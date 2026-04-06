#!/usr/bin/env python3
"""
US Economic Tracker — Web Dashboard
────────────────────────────────────
Flask server that serves the same analysis as the Excel tracker
via a live web dashboard. Imports all logic from economic_tracker.py.

Run locally:  py -3 web_dashboard.py
Deploy:       Railway / Render / any WSGI host (see Procfile)
"""
import os, sys, time, threading, traceback
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
CACHE_TTL = 86400  # 24 hours — refreshes once per day on first visit


def _gather_data():
    """Run the full analysis pipeline — same steps as main() in economic_tracker.py."""
    result = {}

    # Step 1: Indicator data
    data = et.fetch_all()
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
                          ("umcsent","umcsent"),("cfnai","cfnai"),("permits","permits")]:
            hist = hmm_hist.get(key, (None, []))[1]
            if hist and len(hist) >= 36:
                hmm_g_series.append((name, hist))
        hmm_g_panel = et._build_aligned_panel(hmm_g_series, n=220) if len(hmm_g_series) >= 3 else None
        hmm_i_series = []
        for name, key in [("cpi","cpi_yoy"),("core","core_yoy"),("ppi","ppi_mom"),("crb","crb_yoy")]:
            hist = hmm_hist.get(key, (None, []))[1]
            if hist and len(hist) >= 36:
                hmm_i_series.append((name, hist))
        hmm_i_panel = et._build_aligned_panel(hmm_i_series, n=220) if len(hmm_i_series) >= 2 else None
        hmm_g = et._estimate_dfm(hmm_g_panel, anchor_col="payems", factor_name="HMM-Growth") if hmm_g_panel is not None else None
        hmm_i = et._estimate_dfm(hmm_i_panel, anchor_col="cpi", factor_name="HMM-Inflation")
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
    indicators = []
    for key, name, category, fmt, thresholds, color, y_label in et.METRICS:
        val, hist = data.get(key, (None, []))
        pressure, impact, fed = et.classify(key, val)
        weight = et.INDICATOR_WEIGHTS.get(key, 1.0)
        itype = et.INDICATOR_TYPE.get(key, "Lagging")
        unit = et.UNIT_MAP.get(key, "")

        # Trend: compare current vs 3 months ago
        trend = "flat"
        if val is not None and len(hist) >= 4:
            ago = hist[3][1] if len(hist) > 3 else hist[-1][1]
            if ago != 0:
                diff = val - ago
                inverted = key in ("hy_spread", "ig_spread", "epu", "unemployment_rate", "initial_claims")
                if inverted:
                    diff = -diff
                if diff > abs(ago) * 0.03:
                    trend = "up"
                elif diff < -abs(ago) * 0.03:
                    trend = "down"

        # History for sparkline (last 12 months, oldest first)
        spark_data = [h[1] for h in reversed(hist[:12])] if hist else []

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
         "desc": "20 indicators weighted by timing: leading 2-3x, coincident 1.5-2x, lagging 1x.",
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
        "transition_risk": transition_risk,
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
        "updated_at": time.strftime("%Y-%m-%d %H:%M UTC", time.gmtime()),
    }
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
