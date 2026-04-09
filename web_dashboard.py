#!/usr/bin/env python3
"""
US Economic Tracker - Web Dashboard
────────────────────────────────────
Flask server that serves the same analysis as the Excel tracker
via a live web dashboard. Imports all logic from economic_tracker.py.

Run locally:  py -3 web_dashboard.py
Deploy:       Railway / Render / any WSGI host (see Procfile)
"""
import os, sys, time, json, threading, traceback
sys.dont_write_bytecode = True

# ── Patch FRED_API_KEY from env before importing the tracker ──────────────────
# In production the key lives in an environment variable, not in source code.
_env_key = os.environ.get("FRED_API_KEY")

import economic_tracker as et

if _env_key:
    et.FRED_API_KEY = _env_key

import numpy as np
from flask import Flask, render_template, jsonify
from flask.json.provider import DefaultJSONProvider
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address


class _NumpyJSONProvider(DefaultJSONProvider):
    """Handle numpy types that the default encoder chokes on."""
    @staticmethod
    def default(obj):
        if isinstance(obj, np.integer):   return int(obj)
        if isinstance(obj, np.floating):  return float(obj)
        if isinstance(obj, np.bool_):     return bool(obj)
        if isinstance(obj, np.ndarray):   return obj.tolist()
        raise TypeError(f"Not JSON serializable: {type(obj)}")

# ── App setup ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.json_provider_class = _NumpyJSONProvider
app.json = _NumpyJSONProvider(app)
app.config["JSON_SORT_KEYS"] = False

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["30 per minute"],
    storage_uri="memory://",
)

# ── Data cache ────────────────────────────────────────────────────────────────
# Fetching FRED data takes ~15-30 seconds. Cache results for 1 hour so page
# loads are instant and we don't hammer the API.

_cache = {"data": None, "timestamp": 0, "lock": threading.Lock(), "error": None}
CACHE_TTL = 86400  # 24 hours - refreshes once per day on first visit

# ── Per-indicator freshness thresholds ──────────────────────────────────────
# Expected lag (in months) between "now" and the latest observation date.
# An indicator is FRESH (green) if its latest month >= (current month - lag).
# Default is 2 months for any indicator not listed here.
EXPECTED_LAG_MONTHS = {
    # Daily→monthly aggregated: should have at least previous month
    "hy_spread": 1, "ig_spread": 1, "t10y3m": 1, "breakeven_10y": 1, "epu": 1, "dxy": 1,
    "initial_claims": 1,
    # Monthly early/mid-month releases: expect current - 2
    "unemployment_rate": 2, "nonfarm_payrolls": 2, "adp": 2, "ahe_yoy": 2,
    "cpi_yoy": 2, "core_cpi_yoy": 2, "core_pce_yoy": 2, "ppi_mom": 2, "crb_yoy": 2,
    "building_permits": 2, "consumer_conf": 2, "cfnai": 2,
    "indpro_mom": 2, "retail_mom": 2, "import_yoy": 2, "housing_starts": 2,
    # JOLTS ~2-month publication lag: expect current - 3
    "jolts_openings": 3, "jolts_quits": 3,
    # Manual CSV (ISM/Chicago): expect current - 2
    "ism_mfg": 2, "ism_svc": 2, "spg_svc": 2, "chicago_pmi": 2,
}

# ── Maximum acceptable lag before flagging as stale (red) ───────────────────
# Between EXPECTED_LAG and MAX_LAG = amber (old but latest available).
# Beyond MAX_LAG = red (fetch error or data not updating).
MAX_LAG_MONTHS = {
    "hy_spread": 2, "ig_spread": 2, "t10y3m": 2, "breakeven_10y": 2, "epu": 2, "dxy": 2,
    "initial_claims": 2,
    "unemployment_rate": 3, "nonfarm_payrolls": 3, "adp": 3, "ahe_yoy": 3,
    "cpi_yoy": 3, "core_cpi_yoy": 3, "core_pce_yoy": 3, "ppi_mom": 3, "crb_yoy": 3,
    "building_permits": 4, "consumer_conf": 3, "cfnai": 3,
    "indpro_mom": 3, "retail_mom": 3, "import_yoy": 3, "housing_starts": 4,
    "jolts_openings": 5, "jolts_quits": 5,
    "ism_mfg": 3, "ism_svc": 3, "spg_svc": 3, "chicago_pmi": 3,
}

# ── Data source URLs for chart subtitle links ───────────────────────────────
SOURCE_URLS = {
    "unemployment_rate": "https://fred.stlouisfed.org/series/UNRATE",
    "initial_claims":    "https://fred.stlouisfed.org/series/IC4WSA",
    "nonfarm_payrolls":  "https://fred.stlouisfed.org/series/PAYEMS",
    "adp":               "https://fred.stlouisfed.org/series/ADPMNUSNERSA",
    "ahe_yoy":           "https://fred.stlouisfed.org/series/CES0500000003",
    "cpi_yoy":           "https://fred.stlouisfed.org/series/CPIAUCSL",
    "core_cpi_yoy":      "https://fred.stlouisfed.org/series/CPILFESL",
    "core_pce_yoy":      "https://fred.stlouisfed.org/series/PCEPILFE",
    "import_yoy":        "https://fred.stlouisfed.org/series/IR",
    "ppi_mom":           "https://fred.stlouisfed.org/series/PPIACO",
    "crb_yoy":           "https://fred.stlouisfed.org/series/PALLFNFINDEXM",
    "indpro_mom":        "https://fred.stlouisfed.org/series/INDPRO",
    "retail_mom":        "https://fred.stlouisfed.org/series/RSAFS",
    "consumer_conf":     "https://fred.stlouisfed.org/series/UMCSENT",
    "hy_spread":         "https://fred.stlouisfed.org/series/BAMLH0A0HYM2",
    "ig_spread":         "https://fred.stlouisfed.org/series/BAMLC0A0CM",
    "t10y3m":            "https://fred.stlouisfed.org/series/T10Y3M",
    "building_permits":  "https://fred.stlouisfed.org/series/PERMIT",
    "jolts_openings":    "https://fred.stlouisfed.org/series/JTSJOL",
    "jolts_quits":       "https://fred.stlouisfed.org/series/JTSQUR",
    "breakeven_10y":     "https://fred.stlouisfed.org/series/T10YIE",
    "cfnai":             "https://fred.stlouisfed.org/series/CFNAI",
    "epu":               "https://fred.stlouisfed.org/series/USEPUINDXD",
    "dxy":               "https://fred.stlouisfed.org/series/DTWEXBGS",
    "housing_starts":    "https://fred.stlouisfed.org/series/HOUST",
    "ism_mfg":     "https://www.ismworld.org/supply-management-news-and-reports/reports/ism-report-on-business/pmi/pmi-composite-index/",
    "ism_svc":     "https://www.ismworld.org/supply-management-news-and-reports/reports/ism-report-on-business/services/services-pmi/",
    "chicago_pmi": "https://www.mni-indicators.com",
    "spg_svc":     "https://www.pmi.spglobal.com/Public/Home/PressRelease",
}


def _format_transition(hmm_result):
    """Format HMM transition matrix for template display."""
    A = hmm_result.get("transition")
    if A is None:
        return None
    names = et.REGIME_NAMES_HMM
    # Relabel rows/cols to match regime names (same order as HMM state labeling)
    # A is raw state indices - need to map through the same labeling as estimate_hmm_regime
    # For simplicity, return raw matrix with regime-order rows
    return {
        "regimes": names,
        "rows": [[round(float(A[i][j]) * 100, 1) for j in range(4)] for i in range(4)],
    }

def _optimal_allocation(regime_probs):
    """Probability-weighted sector allocation from regime probabilities."""
    if not regime_probs:
        return None
    try:
        sector_scores = {}
        for regime, prob_pct in regime_probs.items():
            prob = prob_pct / 100.0
            for sector, ret in et.REGIME_SECTORS.get(regime, []):
                sector_scores[sector] = sector_scores.get(sector, 0) + prob * ret
        # Normalise positive scores to weights summing to 100%
        total_pos = sum(max(0, v) for v in sector_scores.values())
        if total_pos <= 0:
            return None
        weights = [(s, round(max(0, v) / total_pos * 100, 1)) for s, v in sector_scores.items()]
        weights.sort(key=lambda x: -x[1])
        # Also include the raw expected return
        return [{"sector": s, "weight": w, "expected": round(sector_scores[s], 1)} for s, w in weights if w > 0]
    except Exception:
        return None

_SHORT_NAMES = {
    "unemployment_rate": "Unemployment", "initial_claims": "Claims",
    "nonfarm_payrolls": "NFP", "adp": "ADP", "ahe_yoy": "Hourly Earnings",
    "cpi_yoy": "CPI", "core_cpi_yoy": "Core CPI", "core_pce_yoy": "Core PCE",
    "import_yoy": "Import Prices", "ppi_mom": "PPI", "crb_yoy": "CRB",
    "ism_mfg": "ISM Mfg", "ism_svc": "ISM Svc", "spg_svc": "S&P Svc",
    "chicago_pmi": "Chicago PMI", "indpro_mom": "Ind. Prod", "retail_mom": "Retail",
    "consumer_conf": "Sentiment", "hy_spread": "HY Spread", "ig_spread": "IG Spread",
    "t10y3m": "Yield Curve", "building_permits": "Permits", "jolts_openings": "JOLTS Open",
    "jolts_quits": "Quits Rate", "breakeven_10y": "Breakeven", "cfnai": "CFNAI",
    "epu": "EPU", "dxy": "Dollar", "housing_starts": "Housing",
}

def _compute_correlation(indicators):
    """Compute pairwise correlation matrix from indicator histories."""
    try:
        labels, series = [], []
        for ind in indicators:
            h = ind.get("history", [])
            if len(h) >= 6:
                labels.append(_SHORT_NAMES.get(ind["key"], ind["key"]))
                series.append(h[-12:] if len(h) >= 12 else h)
        if len(labels) < 3:
            return None
        # Pad to equal length (min of all series lengths)
        min_len = min(len(s) for s in series)
        matrix = np.array([s[-min_len:] for s in series])
        corr = np.corrcoef(matrix)
        return {
            "labels": labels,
            "matrix": [[round(float(corr[i][j]), 2) for j in range(len(labels))] for i in range(len(labels))],
        }
    except Exception:
        return None

def _get_recessions():
    """Fetch NBER recession periods from FRED USREC series."""
    try:
        usrec = et._fred("USREC", 500)
        if not usrec:
            return []
        recessions = []
        in_rec = False
        start = None
        for d, v in reversed(usrec):
            if v == 1 and not in_rec:
                start = d[:7]
                in_rec = True
            elif v == 0 and in_rec:
                recessions.append({"start": start, "end": d[:7]})
                in_rec = False
        if in_rec:
            recessions.append({"start": start, "end": usrec[0][0][:7]})
        return recessions
    except Exception:
        return []

def _read_recent_alerts(limit=5):
    """Read recent alerts from the JSON log file."""
    if not os.path.isfile(et.ALERT_LOG):
        return []
    try:
        with open(et.ALERT_LOG, encoding="utf-8") as f:
            log = json.load(f)
        return log[:limit]
    except Exception:
        return []

def _gather_data():
    """Run the full analysis pipeline - same steps as main() in economic_tracker.py."""
    result = {}

    # Step 0: Parse PDFs and update CSVs so PMI data is current
    et.import_pmi_csvs()
    try:
        mfg_d, mfg_v, svc_d, svc_v, chi_d, chi_v, spg_d, spg_v = et.parse_reports_folder()
        et.update_pmi_csvs(mfg_d, mfg_v, svc_d, svc_v, chi_d, chi_v, spg_d, spg_v)
    except Exception as e:
        print(f"  [WARN] PDF parsing failed: {e}")

    # Step 1: Indicator data
    data = et.fetch_all()
    # Override PMI values from CSV (more current than Excel workbook)
    for pmi_key, csv_path in et.CSV_PMI_MAP.items():
        csv_rows = et._read_pmi_csv(csv_path)
        if csv_rows and pmi_key in data:
            data[pmi_key] = (csv_rows[0][1], csv_rows[:24])
    ff_hist = et.fetch_fedfunds(48)
    pos_data = et.fetch_positioning()

    # Step 2: DFM
    dfm_results = {"growth": None, "inflation": None, "credit": None}
    try:
        dfm_hist = et.fetch_dfm_history()
        g_series = [
            ("unrate",      dfm_hist["unrate"][1]),
            ("ic4wsa",      dfm_hist["ic4wsa_m"][1]),
            ("payems",      dfm_hist["payems_m"][1]),
            ("umcsent",     dfm_hist["umcsent"][1]),
            ("ism_mfg",     dfm_hist["ism_mfg_csv"][1]),
            ("ism_svc",     dfm_hist["ism_svc_csv"][1]),
            ("chicago_pmi", dfm_hist["chicago_pmi_csv"][1]),
        ]
        if dfm_hist.get("cfnai") and len(dfm_hist["cfnai"][1]) >= 36:
            g_series.append(("cfnai", dfm_hist["cfnai"][1]))
        if dfm_hist.get("permit_yoy") and len(dfm_hist["permit_yoy"][1]) >= 36:
            g_series.append(("permits", dfm_hist["permit_yoy"][1]))
        if dfm_hist.get("jolts") and len(dfm_hist["jolts"][1]) >= 36:
            g_series.append(("jolts", dfm_hist["jolts"][1]))
        if dfm_hist.get("adp_m") and len(dfm_hist["adp_m"][1]) >= 36:
            g_series.append(("adp", dfm_hist["adp_m"][1]))
        if dfm_hist.get("indpro_mom") and len(dfm_hist["indpro_mom"][1]) >= 36:
            g_series.append(("indpro", dfm_hist["indpro_mom"][1]))
        if dfm_hist.get("housing_yoy") and len(dfm_hist["housing_yoy"][1]) >= 36:
            g_series.append(("housing", dfm_hist["housing_yoy"][1]))
        g_panel = et._build_aligned_panel(g_series)

        i_series = [
            ("cpi",  dfm_hist["cpi_yoy"][1]),
            ("core", dfm_hist["core_yoy"][1]),
            ("ppi",  dfm_hist["ppi_mom"][1]),
            ("crb",  dfm_hist["crb_yoy"][1]),
        ]
        if dfm_hist.get("breakeven") and len(dfm_hist["breakeven"][1]) >= 36:
            i_series.append(("breakeven", dfm_hist["breakeven"][1]))
        if dfm_hist.get("m2_yoy") and len(dfm_hist["m2_yoy"][1]) >= 36:
            i_series.append(("m2", dfm_hist["m2_yoy"][1]))
        if dfm_hist.get("pce_yoy") and len(dfm_hist["pce_yoy"][1]) >= 36:
            i_series.append(("pce", dfm_hist["pce_yoy"][1]))
        if dfm_hist.get("import_yoy") and len(dfm_hist["import_yoy"][1]) >= 36:
            i_series.append(("imports", dfm_hist["import_yoy"][1]))
        i_panel = et._build_aligned_panel(i_series)

        c_series = []
        if dfm_hist.get("hy_oas") and len(dfm_hist["hy_oas"][1]) >= 36:
            c_series.append(("hy_oas", dfm_hist["hy_oas"][1]))
        if dfm_hist.get("ig_oas") and len(dfm_hist["ig_oas"][1]) >= 36:
            c_series.append(("ig_oas", dfm_hist["ig_oas"][1]))
        if dfm_hist.get("nfci_m") and len(dfm_hist["nfci_m"][1]) >= 36:
            c_series.append(("nfci", dfm_hist["nfci_m"][1]))
        if dfm_hist.get("t10y3m") and len(dfm_hist["t10y3m"][1]) >= 36:
            c_series.append(("t10y3m", dfm_hist["t10y3m"][1]))
        c_panel = et._build_aligned_panel(c_series) if len(c_series) >= 2 else None

        dfm_results = {
            "growth":    et._estimate_dfm(g_panel,  anchor_col="payems", factor_name="Growth"),
            "inflation": et._estimate_dfm(i_panel,  anchor_col="cpi",    factor_name="Inflation"),
            "credit":    et._estimate_dfm(c_panel,  anchor_col="nfci",   factor_name="Credit") if c_panel is not None else None,
        }
    except Exception as e:
        print(f"  [WARN] DFM estimation failed: {e}")

    # Step 3: HMM
    hmm_result = None
    try:
        hmm_hist = et.fetch_hmm_history()
        hmm_g_series = []
        for name, key in [("unrate","unrate"),("ic4wsa","ic4wsa_m"),("payems","payems_m"),
                          ("umcsent","umcsent"),("cfnai","cfnai"),("permits","permits"),
                          ("indpro","indpro_m"),("housing","housing_yoy")]:
            hist = hmm_hist.get(key, (None, []))[1]
            if hist and len(hist) >= 36:
                hmm_g_series.append((name, hist))
        hmm_g_panel = et._build_aligned_panel(hmm_g_series, n=700) if len(hmm_g_series) >= 3 else None
        hmm_i_series = []
        for name, key in [("cpi","cpi_yoy"),("core","core_yoy"),("ppi","ppi_mom")]:
            hist = hmm_hist.get(key, (None, []))[1]
            if hist and len(hist) >= 36:
                hmm_i_series.append((name, hist))
        hmm_i_panel = et._build_aligned_panel(hmm_i_series, n=700) if len(hmm_i_series) >= 2 else None
        hmm_g = et._estimate_dfm(hmm_g_panel, anchor_col="payems", factor_name="HMM-Growth", rolling_zscore=120) if hmm_g_panel is not None else None
        hmm_i = et._estimate_dfm(hmm_i_panel, anchor_col="cpi", factor_name="HMM-Inflation", rolling_zscore=120)
        if hmm_g and hmm_i:
            hmm_result = et.estimate_hmm_regime(hmm_g["series"], hmm_i["series"])
    except Exception as e:
        print(f"  [WARN] HMM estimation failed: {e}")

    # Step 4: Derived signals
    vote_regime = et.classify_regime(data)
    dfm_factors = None
    if dfm_results.get("growth") and dfm_results.get("inflation"):
        dfm_factors = {
            "growth": dfm_results["growth"]["latest"],
            "inflation": dfm_results["inflation"]["latest"],
        }
    dfm_regime = et.classify_regime(data, dfm_factors) if dfm_factors else vote_regime

    current_regime = hmm_result["regime"] if hmm_result else dfm_regime
    w_label, w_score, w_lead, w_coin, w_lag = et.weighted_signal(data)
    forecast = et.directional_forecast(data, dfm_results)
    conviction = et.compute_conviction(hmm_result, forecast)
    stance = et.monetary_policy_stance(ff_hist)

    # Sahm rule
    unrate_val, unrate_hist = data.get("unemployment_rate", (None, []))
    sahm = et.compute_sahm_rule(unrate_hist) if unrate_hist else None

    # Step 5: Build indicator table
    # Pre-load PMI CSV history so ISM/Chicago charts always have data
    _pmi_csv_hist = {}
    for pmi_key, csv_path in et.CSV_PMI_MAP.items():
        csv_rows = et._read_pmi_csv(csv_path)  # [(YYYY-MM-01, float), ...] newest-first
        if csv_rows:
            _pmi_csv_hist[pmi_key] = csv_rows

    indicators = []
    for key, name, category, fmt, thresholds, color, y_label in et.METRICS:
        val, hist = data.get(key, (None, []))
        pressure, impact, fed = et.classify(key, val)
        weight = et.INDICATOR_WEIGHTS.get(key, 1.0)
        itype = et.INDICATOR_TYPE.get(key, "Lagging")
        unit = et.UNIT_MAP.get(key, "")

        # For PMI keys, prefer CSV history (24 months) over manual input
        if key in _pmi_csv_hist:
            csv_h = _pmi_csv_hist[key]
            use_hist = csv_h  # newest-first, up to 24 months
            n_hist = 24
        else:
            use_hist = hist  # newest-first from FRED
            n_hist = 12

        # Trend: compare current vs 3 months ago
        trend = "flat"
        if val is not None and len(use_hist) >= 4:
            ago = use_hist[3][1] if len(use_hist) > 3 else use_hist[-1][1]
            if ago != 0:
                diff = val - ago
                inverted = key in ("hy_spread", "ig_spread", "epu", "unemployment_rate", "initial_claims")
                if inverted:
                    diff = -diff
                if diff > abs(ago) * 0.03:
                    trend = "up"
                elif diff < -abs(ago) * 0.03:
                    trend = "down"

        # Anomaly detection: flag if value is >4 sigma from recent history
        anomaly = False
        if val is not None and len(use_hist) >= 6:
            hist_vals = [h[1] for h in use_hist[:12]]
            h_mean = sum(hist_vals) / len(hist_vals)
            h_std = (sum((v - h_mean)**2 for v in hist_vals) / len(hist_vals)) ** 0.5
            if h_std > 0 and abs(val - h_mean) > 4 * h_std:
                anomaly = True

        # History for sparkline and charts (oldest first)
        trimmed = list(reversed(use_hist[:n_hist])) if use_hist else []
        spark_data = [h[1] for h in trimmed]
        # Date labels as MM-YY (e.g. '03-26')
        date_labels = [f"{d[0][5:7]}-{d[0][2:4]}" for d in trimmed]

        # Last observation date and freshness check (traffic light)
        last_date_str = ""
        last_date_sort = 0
        data_status = "stale"  # red by default
        if trimmed:
            newest_iso = trimmed[-1][0]  # e.g. "2026-03-01"
            last_date_str = f"{newest_iso[5:7]}-{newest_iso[2:4]}"  # "03-26"
            last_date_sort = int(newest_iso[2:4]) * 100 + int(newest_iso[5:7])  # YYMM for sorting

            # Traffic light: green (fresh) / amber (old but expected) / red (stale/error)
            obs_months = int(newest_iso[:4]) * 12 + int(newest_iso[5:7])
            now = time.gmtime()
            cur_months = now.tm_year * 12 + now.tm_mon
            months_behind = cur_months - obs_months
            lag = EXPECTED_LAG_MONTHS.get(key, 2)
            max_lag = MAX_LAG_MONTHS.get(key, lag + 1)
            if months_behind <= lag:
                data_status = "fresh"   # green - up to date
            elif months_behind <= max_lag:
                data_status = "old"     # amber - latest available, slow publisher
            else:
                data_status = "stale"   # red - something wrong

        indicators.append({
            "key": key,
            "name": name,
            "category": category,
            "value": round(val, 2) if val is not None else None,
            "unit": unit,
            "signal": pressure,
            "impact": impact,
            "weight": weight,
            "type": itype,
            "trend": trend,
            "history": spark_data,
            "dates": date_labels,
            "last_date": last_date_str,
            "last_date_sort": last_date_sort,
            "data_status": data_status,
            "source_url": SOURCE_URLS.get(key, ""),
            "cont_score": round(et.continuous_score(key, val), 2) if val is not None else 0.0,
            "anomaly": anomaly,
        })

    # Step 6: Positioning
    pos_indicators = []
    for pkey in et.POS_KEYS:
        meta = et.POS_META[pkey]
        pname = meta[0]
        val, raw = pos_data.get(pkey, (None, []))
        signal, impact, action = et.classify_pos(pkey, val)
        pos_indicators.append({
            "key": pkey,
            "name": pname,
            "value": round(val, 2) if val is not None else None,
            "signal": signal,
            "label": et.pos_lbl(signal),
            "impact": impact,
            "action": action,
        })

    pos_overall, pos_color = et.overall_positioning(
        [et.classify_pos(k, pos_data.get(k, (None, []))[0])[0] for k in et.POS_KEYS]
    )

    # Step 7: Sector playbook
    sectors = et.REGIME_SECTORS.get(current_regime, [])

    # Step 8: DFM factor scores for display
    signal_fns = {"growth": et._dfm_signal, "inflation": et._inf_signal, "credit": et._credit_signal}
    factor_scores = {}
    for fname in ("growth", "inflation", "credit"):
        dfm_r = dfm_results.get(fname)
        if dfm_r:
            sig_label, _ = signal_fns[fname](dfm_r["latest"])
            factor_scores[fname] = {
                "score": round(dfm_r["latest"], 2),
                "momentum": round(dfm_r.get("momentum", 0), 2),
                "signal": sig_label,
            }

    # HMM probabilities
    hmm_probs = {}
    hmm_confidence = 0
    transition_risk = False
    if hmm_result:
        hmm_probs = {k: round(v * 100, 1) for k, v in hmm_result["probabilities"].items()}
        hmm_confidence = round(hmm_result["confidence"] * 100, 1)
        transition_risk = hmm_result.get("transition_risk", False)

    # Style factors
    style = et.STYLE_FACTORS.get(stance, [])

    # Calendar entries
    calendar = []
    try:
        for days, key, name, source, priority, rd in et.build_calendar_entries()[:12]:
            calendar.append({
                "key": key, "name": name, "source": source,
                "priority": priority, "date": rd.strftime("%d %b %Y"),
                "days": days,
            })
    except Exception:
        pass

    # Methodology ranking (dynamic readings)
    hmm_reading = f"{current_regime} {hmm_confidence:.0f}%" if hmm_result else "Unavailable"
    score_reading = f"{w_score:+.2f} {w_label.upper()}"
    dfm_reading = (f"Growth {dfm_factors['growth']:+.2f}, Inflation {dfm_factors['inflation']:+.2f}"
                   if dfm_factors else "Unavailable")
    methodology = [
        {"rank": 1, "name": "HMM Regime Detection", "reading": hmm_reading,
         "reliability": "HIGH",
         "desc": "4-state Hidden Markov Model on growth/inflation DFM factors (220 months).",
         "fail": "Overfits to GFC/COVID; may misclassify novel regimes. Trust when >80%.",
         "interp": "Primary signal. >80% = act on regime. <60% = treat as TRANSITION."},
        {"rank": 2, "name": "Weighted Macro Score", "reading": score_reading,
         "reliability": "HIGH",
         "desc": "29 indicators weighted by timing: leading 2-3x, coincident 1.5-2x, lagging 1x.",
         "fail": "Leading indicators dominate; can front-run false signals.",
         "interp": "Best single number. >+0.3 = bullish. <-0.3 = bearish."},
        {"rank": 3, "name": "DFM Factor Model", "reading": dfm_reading,
         "reliability": "MEDIUM",
         "desc": "3 latent factors (Growth, Inflation, Credit) via Kalman filter on 15+ series.",
         "fail": "Kalman unstable with <36 months; PCA fallback less precise.",
         "interp": "Cross-ref with HMM. Momentum arrows more useful than levels."},
        {"rank": 4, "name": "Directional Forecast", "reading": forecast[0],
         "reliability": "MEDIUM",
         "desc": "3-6 month outlook from 9 leading indicators' 3-month momentum.",
         "fail": "Pure extrapolation - cannot predict shocks or policy pivots.",
         "interp": "Trend confirmation only. Most useful when it agrees with HMM."},
        {"rank": 5, "name": "Conviction Signal", "reading": conviction[0],
         "reliability": "SUPPORTING",
         "desc": "Combines HMM confidence with forecast direction for allocation sizing.",
         "fail": "Derivative - inherits all upstream weaknesses.",
         "interp": "TRANSITION is the most actionable signal (raise cash)."},
    ]

    # Check for walk-forward cache (produced by --walk-forward CLI)
    wf_cache_path = os.path.join(et.DATA_DIR, "walk_forward_cache.json")
    walk_forward = None
    wf_regime_history = None
    wf_current_regime = None
    if os.path.isfile(wf_cache_path):
        try:
            with open(wf_cache_path, encoding="utf-8") as f:
                wf_data = json.load(f)
            if wf_data:
                walk_forward = {
                    "dates": [r["date"] for r in wf_data if r.get("sp500_return") is not None],
                    "returns": [r["sp500_return"] for r in wf_data if r.get("sp500_return") is not None],
                    "regimes": [r["regime"] for r in wf_data if r.get("sp500_return") is not None],
                }
                # Build stacked area history from walk-forward probs
                wf_regime_history = {"dates": [], "Early Recovery": [], "Reflation": [], "Stagflation": [], "Deflation": []}
                for r in wf_data:
                    if r.get("probs"):
                        wf_regime_history["dates"].append(r["date"])
                        for regime in et.REGIME_NAMES_HMM:
                            wf_regime_history[regime].append(r["probs"].get(regime, 0))
                wf_current_regime = wf_data[-1].get("regime") if wf_data else None
        except Exception:
            pass

    # Backtest: prefer walk-forward cache (out-of-sample) over in-sample filtered
    backtest = None
    if walk_forward and walk_forward.get("dates") and len(walk_forward["dates"]) > 24:
        # Use walk-forward (out-of-sample) as primary backtest
        import math as _math
        wf_regime_names = et.REGIME_NAMES_HMM
        wf_buckets = {r: [] for r in wf_regime_names}
        for i, ret in enumerate(walk_forward["returns"]):
            wf_buckets[walk_forward["regimes"][i]].append(ret)
        wf_stats = {}
        for r in wf_regime_names:
            rets = wf_buckets[r]
            n = len(rets)
            if n < 2:
                wf_stats[r] = {"n": n, "avg": 0, "sharpe": 0, "hit": 0}
                continue
            avg = sum(rets) / n
            std = (sum((x - avg)**2 for x in rets) / (n - 1)) ** 0.5
            sharpe = (avg / std * _math.sqrt(12)) if std > 0 else 0
            hit = sum(1 for x in rets if x > 0) / n * 100
            wf_stats[r] = {"n": n, "avg": round(avg, 2), "sharpe": round(sharpe, 2), "hit": round(hit, 1)}
        cum = [100.0]
        for ret in walk_forward["returns"]:
            cum.append(cum[-1] * (1 + ret / 100))
        backtest = {
            "dates": walk_forward["dates"],
            "returns": walk_forward["returns"],
            "regimes": walk_forward["regimes"],
            "stats": wf_stats,
            "cumulative": [round(c, 1) for c in cum[1:]],
            "is_walk_forward": True,
        }

    rh_filtered = hmm_result.get("regime_history_filtered") if hmm_result else None
    if backtest is None and rh_filtered:
        try:
            sp_monthly = et._fred("SPASTT01USM661N", 400)  # OECD S&P 500 monthly index
            if sp_monthly and len(sp_monthly) > 24:
                # Already monthly (newest-first). Compute returns.
                sp_rets = {}
                for i in range(len(sp_monthly) - 1):
                    ym = sp_monthly[i][0][:7]
                    sp_rets[ym] = round(
                        (sp_monthly[i][1] - sp_monthly[i+1][1])
                        / sp_monthly[i+1][1] * 100, 2)
                # Match regime to returns
                regime_names = et.REGIME_NAMES_HMM
                bt_dates, bt_returns, bt_regimes = [], [], []
                regime_buckets = {r: [] for r in regime_names}
                for idx, ym in enumerate(rh_filtered["dates"]):
                    if ym not in sp_rets:
                        continue
                    # Dominant regime (forward-only, no look-ahead)
                    probs = {r: rh_filtered[r][idx] for r in regime_names}
                    dom = max(probs, key=probs.get)
                    ret = sp_rets[ym]
                    bt_dates.append(ym)
                    bt_returns.append(ret)
                    bt_regimes.append(dom)
                    regime_buckets[dom].append(ret)
                # Per-regime stats
                import math
                bt_stats = {}
                for r in regime_names:
                    rets = regime_buckets[r]
                    n = len(rets)
                    if n < 2:
                        bt_stats[r] = {"n": n, "avg": 0, "sharpe": 0, "hit": 0}
                        continue
                    avg = sum(rets) / n
                    std = (sum((x - avg)**2 for x in rets) / (n - 1)) ** 0.5
                    sharpe = (avg / std * math.sqrt(12)) if std > 0 else 0
                    hit = sum(1 for x in rets if x > 0) / n * 100
                    bt_stats[r] = {"n": n, "avg": round(avg, 2), "sharpe": round(sharpe, 2), "hit": round(hit, 1)}
                # Cumulative return (indexed to 100)
                cum = [100.0]
                for ret in bt_returns:
                    cum.append(cum[-1] * (1 + ret / 100))
                backtest = {
                    "dates": bt_dates,
                    "returns": bt_returns,
                    "regimes": bt_regimes,
                    "stats": bt_stats,
                    "cumulative": [round(c, 1) for c in cum[1:]],
                    "is_walk_forward": False,
                }
        except Exception as e:
            print(f"  [WARN] Backtest failed: {e}")

    # Tooltips
    tooltips = {}
    if hasattr(et, 'METRIC_TOOLTIPS'):
        tooltips = {k: v for k, v in et.METRIC_TOOLTIPS.items()}

    result = {
        "regime": current_regime,
        "regime_desc": et.REGIME_DESC.get(current_regime, ""),
        "regime_color": et.REGIME_COLORS.get(current_regime, "666666"),
        "regime_equity": et.REGIME_EQUITY.get(current_regime, ("N/A", "666666")),
        "hmm_probs": hmm_probs,
        "hmm_confidence": hmm_confidence,
        "stability": round(hmm_result.get("stability", 1.0) * 100, 1) if hmm_result else 100.0,
        "wf_regime_history": wf_regime_history,
        "wf_regime": wf_current_regime,
        "is_transition": wf_current_regime is not None and wf_current_regime != current_regime,
        "transition_from": wf_current_regime if (wf_current_regime and wf_current_regime != current_regime) else None,
        "transition_to": current_regime if (wf_current_regime and wf_current_regime != current_regime) else None,
        "transition_risk": transition_risk,
        "regime_history": hmm_result.get("regime_history") if hmm_result else None,
        "regime_history_filtered": hmm_result.get("regime_history_filtered") if hmm_result else None,
        "transition_matrix": _format_transition(hmm_result) if hmm_result else None,
        "backtest": backtest,
        "walk_forward": walk_forward,
        "macro_score": round(w_score, 2),
        "macro_label": w_label,
        "tier_scores": {
            "leading": round(w_lead, 2),
            "coincident": round(w_coin, 2),
            "lagging": round(w_lag, 2),
        },
        "forecast_label": forecast[0],
        "forecast_color": forecast[1],
        "forecast_summary": forecast[2],
        "conviction_label": conviction[0],
        "conviction_color": conviction[1],
        "conviction_desc": conviction[2],
        "monetary_stance": stance,
        "sahm_rule": sahm,
        "indicators": indicators,
        "positioning": pos_indicators,
        "pos_overall": pos_overall,
        "pos_color": pos_color,
        "sectors": sectors,
        "factor_scores": factor_scores,
        "style_factors": style,
        "vote_regime": vote_regime,
        "calendar": calendar,
        "methodology": methodology,
        "tooltips": tooltips,
        "correlation": _compute_correlation(indicators),
        "portfolio": _optimal_allocation(hmm_probs),
        "updated_at": time.strftime("%Y-%m-%d %H:%M UTC", time.gmtime()),
        "alerts": _read_recent_alerts(5),
        "recessions": _get_recessions(),
    }

    # Run alert check (so alerts fire even when only the web dashboard runs)
    try:
        old_state = et.load_previous_state()
        new_alerts = et.check_alerts(old_state, current_regime, hmm_confidence / 100,
                                     w_score, forecast[0] if forecast else "")
        et.save_current_state(current_regime, hmm_confidence / 100, w_score,
                              forecast[0] if forecast else "", w_lead)
        if new_alerts:
            result["alerts"] = _read_recent_alerts(5)  # re-read after new alerts logged
    except Exception:
        pass

    return result


def _get_data():
    """Return cached data, refreshing if stale."""
    now = time.time()
    with _cache["lock"]:
        if _cache["data"] and (now - _cache["timestamp"]) < CACHE_TTL:
            return _cache["data"], _cache["error"]

    # Fetch outside the lock to avoid blocking other requests
    try:
        fresh = _gather_data()
        with _cache["lock"]:
            _cache["data"] = fresh
            _cache["timestamp"] = time.time()
            _cache["error"] = None
        return fresh, None
    except Exception as e:
        traceback.print_exc()
        with _cache["lock"]:
            _cache["error"] = str(e)
            # Return stale data if available
            return _cache["data"], str(e)


# ── Security headers ──────────────────────────────────────────────────────────

@app.after_request
def _security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com; "
        "img-src 'self' data:; "
        "connect-src 'self'"
    )
    return response


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    data, error = _get_data()
    if data is None:
        return render_template("dashboard.html", data=None, error=error or "Loading data..."), 503
    return render_template("dashboard.html", data=data, error=error)


@app.route("/methodology")
def methodology_page():
    data, _ = _get_data()
    return render_template("methodology.html", data=data)

@app.route("/api/v1/summary")
def api_summary():
    """JSON endpoint: regime, score, forecast, conviction."""
    data, error = _get_data()
    if data is None:
        return jsonify({"error": error or "Loading"}), 503
    return jsonify({
        "regime": data.get("regime"),
        "confidence": data.get("hmm_confidence"),
        "probabilities": data.get("hmm_probs"),
        "macro_score": data.get("macro_score"),
        "macro_label": data.get("macro_label"),
        "forecast": data.get("forecast_label"),
        "conviction": data.get("conviction_label"),
        "monetary_stance": data.get("monetary_stance"),
        "sahm_rule": data.get("sahm_rule"),
        "updated_at": data.get("updated_at"),
    })

@app.route("/api/v1/indicators")
def api_indicators():
    """JSON endpoint: all indicator data."""
    data, error = _get_data()
    if data is None:
        return jsonify({"error": error or "Loading"}), 503
    return jsonify(data.get("indicators", []))

@app.route("/refresh")
def refresh():
    """Force a data refresh, then redirect back to the dashboard."""
    with _cache["lock"]:
        _cache["timestamp"] = 0  # expire the cache
    _get_data()  # re-fetch now
    from flask import redirect
    return redirect("/")


@app.route("/api/data")
def api_data():
    data, error = _get_data()
    if data is None:
        return jsonify({"error": error or "Data not yet available"}), 503
    return jsonify(data)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    print(f"\n  Starting US Economic Tracker Dashboard on http://localhost:{port}")
    print(f"  First load will take ~30 seconds while FRED data is fetched.\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
