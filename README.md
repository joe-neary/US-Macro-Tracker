================================================================================
  US ECONOMIC INDICATORS TRACKER  v8
  Automated macro dashboard with regime detection, leading indicators,
  probabilistic HMM classification, and forward-looking signals
================================================================================

OVERVIEW
--------
A single Python script that:
  1. Extracts ISM Manufacturing, ISM Services, and Chicago PMI values
     directly from the official PDF reports (no manual data entry)
  2. Fetches 17 additional indicators live from the FRED API, including
     credit spreads, yield curve, JOLTS, building permits, breakeven
     inflation, CFNAI, and Economic Policy Uncertainty
  3. Estimates a 3-factor Dynamic Factor Model (Growth, Inflation, Credit)
  4. Runs a 4-state Hidden Markov Model for probabilistic regime detection
  5. Computes a weighted macro score, directional forecast, and conviction signal
  6. Backtests regime classification against 219 months of equity returns
  7. Generates a polished Excel dashboard with 7 sheets

Indicators covered (20 total):
  Job Market           -- Unemployment Rate, Initial Jobless Claims,
                          Nonfarm Payrolls
  Inflation            -- CPI (YoY), Core CPI (YoY), PPI (MoM),
                          CRB Commodity Index (YoY)
  Economic Activities  -- ISM Manufacturing PMI, ISM Services PMI,
                          Chicago PMI, Consumer Confidence
  Leading Indicators   -- HY Credit Spread (OAS), IG Credit Spread (OAS),
                          10Y-3M Yield Curve, Building Permits (YoY),
                          JOLTS Job Openings (YoY), JOLTS Quits Rate,
                          10Y Breakeven Inflation, CFNAI,
                          Economic Policy Uncertainty (EPU)


PROJECT STRUCTURE
-----------------
  US Economic Tracker v2/
  ├── economic_tracker.py                  <- main script (run this)
  ├── US_Economic_Tracker.xlsx             <- generated output (7 sheets)
  ├── US_Economic_Tracker_TEMPLATE.xlsx    <- auto-generated template
  ├── .tracker_state.json                  <- state file for change alerts
  ├── READ-ME.txt                          <- this file
  └── Reports/
      ├── ISM Manufacturing.pdf            <- replace each month
      ├── ISM Services.pdf                 <- replace each month
      ├── <MNI Chicago ...>.pdf            <- replace each month
      ├── ISM_Manual_Input.xlsx            <- auto-maintained data log
      └── *.csv                            <- PMI history (10 years)


EXCEL OUTPUT (7 SHEETS)
-----------------------
  Executive Summary    -- Headline regime/outlook/score at row 4 (no scrolling),
                          HMM probability distribution, sub-indicators,
                          weighted score breakdown, 20-indicator table,
                          DFM factors, sector playbook, backtest, conviction
  Job Market           -- Detailed breakdowns, history tables, charts
  Inflation            -- Detailed breakdowns, history tables, charts
  Economic Activities  -- Detailed breakdowns, history tables, charts
  Leading Indicators   -- Credit spreads, yield curve, JOLTS, permits,
                          CFNAI, EPU with charts
  Market Positioning   -- VIX, NFCI, STLFSI, bank lending, yield curve,
                          manual reference sources
  Guide & Sources      -- Signal interpretation, FRED series codes,
                          methodology documentation


================================================================================
  QUICK START
================================================================================

  1. Install Python 3.8+ from https://python.org/downloads/
     (tick "Add Python to PATH" during installation)

  2. Install required packages:
       pip install -r requirements.txt

     Optional (for toast notifications):
       pip install plyer

  3. Set up your FRED API key:
       Copy .env.example to .env and paste your key
       Get a free key at https://fred.stlouisfed.org/docs/api/api_key.html

  4. Drop ISM/MNI PDF reports into the data/pdfs/ folder

  5. Run the Excel tracker:
       py -3 economic_tracker.py

  6. Or run the web dashboard:
       Double-click run_dashboard.bat
       Or: py -3 web_dashboard.py (then open http://localhost:5000)

  7. Output saved to output/US_Economic_Tracker.xlsx


================================================================================
  AUTO-SCHEDULING
================================================================================

Run the setup script to configure daily background updates:

  py -3 automation/setup_daemon.py

This auto-detects your OS and sets it up:
  - Windows: VBS launcher in Startup folder (starts on login, runs hidden)
  - macOS: launchd service (starts on login, auto-restarts on failure)
  - Linux: systemd user service (starts on login, auto-restarts on failure)

Options:

  py -3 automation/setup_daemon.py --time 08:30    Custom time (default 07:00)
  py -3 automation/setup_daemon.py --remove        Stop and remove the daemon

The daemon sleeps until the scheduled time, runs the full pipeline, then
sleeps until tomorrow. If your machine is off at the scheduled time, it
retries every 3 hours (up to 3 attempts) once you are back online.

Manual alternative (runs in a visible terminal):

  py -3 economic_tracker.py --daemon 07:00

Manual setup (Windows Task Scheduler):
  1. Open Task Scheduler (search "Task Scheduler" in Start menu)
  2. Create Basic Task > Name: US Economic Tracker
  3. Trigger: Daily at your preferred time
  4. Action: Start a Program
     Program/script: py.exe
     Arguments: -3 "C:\path\to\economic_tracker.py"
     Start in: C:\path\to\US Economic Tracker v2

macOS: use launchd or cron (see bottom of this file for examples)


================================================================================
  REGIME CHANGE ALERTS
================================================================================

The script tracks state between runs via .tracker_state.json.
When a significant change is detected, it fires an alert:

  - Regime change (e.g. Reflation -> Stagflation)
  - Weighted score shift > 0.15
  - Directional forecast change (e.g. STABILISING -> DETERIORATING)

Alerts appear in the console output. If the 'plyer' package is installed,
a Windows toast notification will also pop up.

Testing notifications:

  1. Install plyer:  pip install plyer
  2. Run:  py -3 automation/test_alert.py
     Or double-click automation/test_alert.bat
  3. A toast notification should appear in the bottom-right of your screen


================================================================================
  TEMPLATE SYSTEM
================================================================================

The template system preserves your custom Excel formatting across runs.

  1. Run the script once to generate the initial output + template
  2. Open US_Economic_Tracker_TEMPLATE.xlsx in Excel
  3. Customise formatting: fonts, colours, column widths, chart styles
  4. Save and close
  5. Future runs will load the template and update data only

The Executive Summary, Market Positioning, Chart Data, and Guide sheets
always rebuild from scratch (complex layout). Category sheets (Job Market,
Inflation, Economic Activities, Leading Indicators) preserve formatting.


================================================================================
  KEY FEATURES (v8)
================================================================================

WEIGHTED MACRO SCORE
  Each indicator is weighted by predictive timing:
    Leading indicators (credit spreads, yield curve, JOLTS, permits, CFNAI, EPU): 2-3x
    Coincident indicators (ISM, claims, payrolls): 1.5-2x
    Lagging indicators (unemployment, CPI, consumer sentiment): 1x
  Output: normalised score from -1.0 to +1.0 with tier breakdown.

HMM REGIME DETECTION
  4-state Gaussian Hidden Markov Model trained on 220 months of DFM factor
  history (back to 2008, includes GFC, COVID, inflation spike).
  Outputs probability distribution: "Stagflation 94% | Deflation 6%"
  Built-in implementation using numpy/scipy -- no external dependency.

3-FACTOR DFM
  Growth factor: 10 series including CFNAI, building permits, JOLTS
  Inflation factor: 6 series including breakeven inflation, M2
  Credit factor: HY OAS, IG OAS, NFCI, 10Y-3M yield curve

DIRECTIONAL FORECAST
  Compares each leading indicator's current value vs 3-month-ago value.
  Outputs: "WEAKENING: 2 of 9 leading indicators deteriorating"

REGIME BACKTEST
  219 months of historical equity returns (NASDAQ) classified by regime.
  Validates that the regime framework predicts differential equity outcomes.

CONVICTION SIGNAL
  Combines HMM confidence with forecast direction:
  HIGH / MEDIUM / LOW / TRANSITION


================================================================================
  PDF NAMING CONVENTIONS
================================================================================
  Report                     Filename must contain
  -----------------------------------------------------------
  ISM Manufacturing PMI      manuf
  ISM Services PMI           serv   or   non-manuf
  MNI Chicago PMI            mni    or   barometer


================================================================================
  FRED API KEY
================================================================================
The script uses a free FRED API key (embedded in the script).
If it stops working: get a free key at fred.stlouisfed.org/docs/api/api_key.html
and replace FRED_API_KEY in the script.


================================================================================
  TROUBLESHOOTING
================================================================================

"py is not recognised"
  Python is not in PATH. Reinstall and tick "Add Python to PATH".

"ModuleNotFoundError: No module named 'pdfplumber'"
  Run: pip install pdfplumber openpyxl requests pandas numpy statsmodels

"ERROR: Cannot save -- close US_Economic_Tracker.xlsx in Excel and re-run."
  The output file is open in Excel. Close it first.

"[WARN] FRED XXXX: 500 Server Error"
  FRED API is temporarily down. Re-run later. The script handles missing
  data gracefully -- it will still generate the workbook with available data.

Values show "ADD PDF"
  No matching PDF was found in the Reports/ folder. Add the PDF and re-run.


================================================================================
  WEB DASHBOARD
================================================================================

The tracker also includes a live web dashboard that displays the same
analysis in a browser — no Excel needed.

LOCAL (for testing):

  1. Install Flask:
       pip install flask flask-limiter

  2. Run the dashboard:
       py -3 web_dashboard.py

  3. Open http://localhost:5000 in your browser
     (first load takes ~30 seconds while FRED data is fetched)

DEPLOY ON RAILWAY (public URL):

  1. Push this folder to a GitHub repository

  2. Go to https://railway.app and sign in with GitHub

  3. Click "New Project" > "Deploy from GitHub repo" > select your repo

  4. Go to Variables tab and add:
       FRED_API_KEY = e0ba612289f5c62c1276b774652a483b
     (or use your own key from fred.stlouisfed.org)

  5. Railway auto-detects Python, installs dependencies, and deploys.
     You'll get a public URL like:
       https://us-economic-tracker-production.up.railway.app

  The dashboard caches FRED data for 1 hour. No manual refresh needed.
  Auto-redeploys whenever you push code changes to GitHub.
  Railway free tier: $5/month credit (more than enough for this app).

FILES:
  web_dashboard.py           Flask server (imports from economic_tracker.py)
  templates/dashboard.html   Dashboard page (dark theme, Chart.js charts)
  requirements.txt           Python dependencies (for Railway/pip)
  Procfile                   Railway start command (gunicorn)


================================================================================
  DATA SOURCES
================================================================================
  FRED API           https://fred.stlouisfed.org
  ISM Reports        https://ismworld.org
  MNI / Chicago PMI  https://www.mni-indicators.com
  NASDAQ Composite   FRED: NASDAQCOM (backtest)
  Baker-Bloom-Davis  Economic Policy Uncertainty Index

================================================================================
