# -*- coding: utf-8 -*-
"""
==============================================================================
NAZEV SKRIPTU: comm_decision_RELIANCE_PROD_WIN7_V40.py
VERZE: PRODUCTION V40 (WIN 7 / PY 3.4.4 + COOLING + WIP TRACKING)
==============================================================================

Poznamka:
- Skript bezi v rezimu kompatibility s Python 3.4 (Win7), proto je kod psany
  konzervativne (napr. bez f-stringu).
- Architektura: SQL read -> rozhodnuti start/stop -> zapis prikazu do queue.
"""
import os
import time
import sys
import logging
import json
import ctypes
from math import cos, pi, sin
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta

# Nutné: pip install numpy pywin32 pytz
import numpy as np
import win32com.client
import pytz

# Centralni runtime konfigurace je v slovniku CONFIG nize.
# Nastaveni zde ma primy dopad na rozhodovani o povelu SET_ZATOP.

# === GLOBÁLNÍ KONFIGURACE ===
TZ = pytz.timezone("Europe/Prague")


def time_features_from_datetime(ts):
    minute_of_day = ts.hour * 60.0 + ts.minute + ts.second / 60.0
    angle = 2.0 * pi * minute_of_day / 1440.0
    return sin(angle), cos(angle)
DAY_KEYS = ["po", "ut", "st", "ct", "pa", "so", "ne"]

CONFIG = {
    # Perioda hlavni smycky (sekundy mezi rozhodovacimi cykly).
    "poll_seconds": 60,
    # Vypis per-second countdownu mezi cykly.
    "countdown_enabled": False,
    # Fallback buffer pred startem smeny, pokud DB nedoda vlastni hodnotu.
    "default_buffer": 30,
    # Maximalni mezera pro slouceni navazujicich smen do jednoho okna.
    "shift_merge_gap_minutes": 2,
    # Zapina/vypina DB detekci "aktivnich" van.
    "active_bath_check": {"enabled": True},
    # Hrube limity pro teploty (priprava pro sanity check).
    "temp_sanity_limits": {"min": 5.0, "max": 110.0},
    # SQL timeout (sekundy) pro DB operace.
    "sql_timeout": 40,
    # Max cas (s) pro SQL chybu teplot pred failover do manualu.
    "sql_error_grace_seconds": 240,
    # Vyradit nedokoncene perf snapshoty po teto dobe (uvolneni RAM).
    "perf_snapshot_ttl_hours": 120,
    # Self-restart ochrana pro dlouhy autonomni beh.
    "self_restart": {
        "enabled": True,
        # Pri poll=60s restart po 2 chybach typicky stihne watchdog 3 min.
        "max_consecutive_loop_errors": 2
    },
    "cooling": {
        # Cooling guard muze pred koncem smeny zastavit topeni, pokud je linka
        # potvrzene "ticha" a bez aktivniho WIP.
        "enabled": True,
        "trigger_min_before_end": 40,  # Buffer 40 minut
        "quiet_time_minutes": 5,  # 5 minut ticha
        "flow_ttl_minutes": 7,
        "stop_confirm_cycles": 2,
        "vstup_day_offset": -2,
        "inventory_limit_pos": 16,  # Promazávání RAM
        "position_aliases": {6: 5, 8: 7},
        "safety_temp_margin": 1.0,
        "routes_cfg_path": r"C:\POL1_PREDICTOR\routes_config.json",
        "xgb_dir": r"C:\POL1_PREDICTOR\NN\models_XGBoost\cooling",
        "positions": ["pozice2", "pozice3", "pozice5", "pozice7", "pozice10", "pozice14"],
        "done_positions": [15, 16, 17, 18, 19]
    },
    # Pocet potvrzovacich hitu pro vyhodnoceni "arrival" na target.
    "arrival_confirmation_hits": 2,
    "positions": {
        # T_target = cilova teplota; T_min_active = hranice "pozice je aktivni".
        "pozice2": {"T_target": 50.0, "T_min_active": 30.0},
        "pozice3": {"T_target": 50.0, "T_min_active": 30.0},
        "pozice5": {"T_target": 65.0, "T_min_active": 40.0},
        "pozice7": {"T_target": 65.0, "T_min_active": 40.0},
        "pozice10": {"T_target": 36.0, "T_min_active": 25.0},
        "pozice14": {"T_target": 38.0, "T_min_active": 25.0},
    },
    "models": {
        # Cesty k JSON exportum modelu pro manualni inference.
        "mlp_dir": r"C:\POL1_PREDICTOR\NN\models_MLP\manual",
        "mlp_suffix": "_time_to_target_mlp.json",
        "xgb_dir": r"C:\POL1_PREDICTOR\NN\models_XGBoost\heating",
        "xgb_suffix": "_time_to_target_xgb.json",
        "bin_enabled": False,
    },
    "predictive": {
        # Po AI startu drzet TOP=1 do konce cilove smeny nebo do zacatku
        # cooling okna (pokud je cooling aktivni).
        "hold_after_predictive_start": True,
        # Ochrana pred nerealnymi ETA outliery.
        "eta_cap_minutes": 720.0,
    },
    "sql": {
        # Primarni SQL zdroj teplot a technologickych dat.
        "server": r"XXXX",
        "database": "XXXX",
        "temp_table": "XXXX",
        "time_column": "XXXX",
        "vymenik_column": "XXXX",
        "median_window": 5,
        "driver": "{SQL Server}",  # Změněno na standardní SQL Server pro Win7
        "username": "XXXXX", "password": "XXXXX",
        "temp_columns": {
            "pozice2": "Pozice2", "pozice3": "Pozice3", "pozice5": "Pozice5",
            "pozice7": "Pozice7", "pozice10": "Pozice10", "pozice14": "Pozice14",
        },
    },
    "sql_bus": {
        # "Bus" tabulky: stav systemu, command queue, performance report.
        "state_table": "dbo.POL1_State",
        "cmd_table": "dbo.POL1_CommandQueue",
        "report_table": "dbo.POL1_PerformanceLog",
        "issued_by": "PY_PROD",
        # Stary INPROGRESS prikaz po TTL uz neblokuje novy enqueue.
        "pending_inprogress_ttl_minutes": 10,
        # Po ERROR stejného SET_ZATOP pockej X sekund pred dalsim enqueue.
        "error_retry_cooldown_seconds": 120,
    },
    "output": {
        "log_file": r"C:\POL1_PREDICTOR\logs\pol1_production.log",
        # Rozhodovaci trace v kazdem cyklu (servisni ladeni).
        "decision_trace_every_cycle": True,
    },
}


# --- RUČNÍ INFERENCE (Kompatibilní s Py 3.4) ---

class ManualXGBInference(object):
    """
    Minimalni XGBoost inference nad JSON exportem stromu.
    Ucel: beh na Win7/Py3.4 bez runtime zavislosti na xgboost.
    """
    def __init__(self, json_path):
        with open(json_path, 'r') as f:
            data = json.load(f)
        self.trees = data['learner']['gradient_booster']['model']['trees']
        try:
            raw_base = data['learner']['learner_model_param']['base_score']
            if isinstance(raw_base, list):
                self.base_score = float(raw_base[0])
            else:
                raw_txt = str(raw_base).strip()
                if raw_txt.startswith('[') and raw_txt.endswith(']'):
                    raw_txt = raw_txt[1:-1].strip()
                self.base_score = float(raw_txt)
        except Exception:
            self.base_score = 0.5

    def predict(self, x):
        val = self.base_score
        for tree in self.trees:
            # Legacy format: tree['nodes'][...]
            if 'nodes' in tree:
                nodes, curr = tree['nodes'], 0
                while 'leaf' not in nodes[curr]:
                    node = nodes[curr]
                    if x[node['split_index']] < node['split_condition']:
                        curr = node['left_child']
                    else:
                        curr = node['right_child']
                val += nodes[curr]['leaf']
                continue

            # New XGBoost JSON format (arrays)
            curr = 0
            left_children = tree['left_children']
            right_children = tree['right_children']
            split_indices = tree['split_indices']
            split_conditions = tree['split_conditions']
            base_weights = tree['base_weights']

            while True:
                left = left_children[curr]
                right = right_children[curr]
                if left == -1 and right == -1:
                    val += base_weights[curr]
                    break
                split_idx = split_indices[curr]
                split_cond = split_conditions[curr]
                if x[split_idx] < split_cond:
                    curr = left
                else:
                    curr = right
        return float(val)


class ManualMLPInference(object):
    """
    Minimalni MLP inference nad JSON exportem vah/scaleru.
    Ucel: konzistentni predikce bez sklearn runtime.
    """
    def __init__(self, json_path):
        with open(json_path, 'r') as f: data = json.load(f)
        self.mean, self.scale = np.array(data['scaler']['mean']), np.array(data['scaler']['scale'])
        self.weights = [np.array(w) for w in data['mlp']['weights']]
        self.biases = [np.array(b) for b in data['mlp']['biases']]

    def predict(self, x_raw):
        x = (np.array(x_raw) - self.mean) / self.scale
        for i in range(len(self.weights) - 1):
            x = np.maximum(np.dot(x, self.weights[i]) + self.biases[i], 0)
        y = np.dot(x, self.weights[-1]) + self.biases[-1]
        return float(np.asarray(y).reshape(-1)[0])


# --- POMOCNÉ FUNKCE ---

def setup_logger(log_path):
    # Kombinace file + stdout loggeru kvuli provozni dohledatelnosti.
    logger = logging.getLogger("POL1Prod")
    logger.setLevel(logging.INFO)
    if not os.path.exists(os.path.dirname(log_path)): os.makedirs(os.path.dirname(log_path))
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = RotatingFileHandler(log_path, maxBytes=5000000, backupCount=5)
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    return logger


def _to_int_safe(v, default):
    try:
        if v is None:
            return default
        txt = str(v).strip().replace(",", ".")
        if txt == "" or txt.lower() == "none":
            return default
        return int(float(txt))
    except Exception:
        return default


def _to_bit_safe(v, default=0):
    if v is None:
        return 1 if default else 0
    if isinstance(v, bool):
        return 1 if v else 0
    txt = str(v).strip().lower()
    if txt in ("1", "true", "-1", "yes", "y", "on"):
        return 1
    if txt in ("0", "false", "no", "n", "off", ""):
        return 0
    try:
        return 0 if int(float(txt)) == 0 else 1
    except Exception:
        return 1 if default else 0


def _build_shift_window_for_date(hours, d):
    # Sestavi casove okno smeny pro konkretni den.
    # Pokud je zapnuti pozdeji nez vypnuti, bere se to jako smena pres pulnoc.
    day_key = DAY_KEYS[d.weekday()]
    zh = _to_int_safe(hours.get(day_key + "_zap", -1), -1)
    zm = _to_int_safe(hours.get(day_key + "_zap_min", 0), 0)
    vh = _to_int_safe(hours.get(day_key + "_vyp", -1), -1)
    vm = _to_int_safe(hours.get(day_key + "_vyp_min", 0), 0)
    if zh < 0 or vh < 0:
        return None
    s = TZ.localize(datetime(d.year, d.month, d.day, zh, zm))
    d2 = d + timedelta(days=1) if zh > vh or (zh == vh and zm >= vm) else d
    e = TZ.localize(datetime(d2.year, d2.month, d2.day, vh, vm))
    return (s, e)


def _build_merged_shift_windows(hours, now, days_back, days_fwd):
    # Slozi okna smen v rozsahu kolem "now" a slouci navazujici intervaly.
    windows = []
    start_d = (now - timedelta(days=days_back)).date()
    total_days = days_back + days_fwd + 1
    for i in range(total_days):
        d = start_d + timedelta(days=i)
        win = _build_shift_window_for_date(hours, d)
        if win:
            windows.append(win)
    if not windows:
        return []
    windows.sort(key=lambda w: w[0])
    merged = [[windows[0][0], windows[0][1]]]
    gap = timedelta(minutes=CONFIG.get("shift_merge_gap_minutes", 1))
    for s, e in windows[1:]:
        if s <= merged[-1][1] + gap:
            if e > merged[-1][1]:
                merged[-1][1] = e
        else:
            merged.append([s, e])
    return [(x[0], x[1]) for x in merged]


def get_shift_window(hours, now):
    # Vrati aktualni smenu, jinak None.
    for s, e in _build_merged_shift_windows(hours, now, 1, 8):
        if s <= now < e:
            return (s, e)
    return None


# --- HLAVNÍ MANAŽERY ---

class CoolingManager(object):
    """
    Ochranna logika pred koncem smeny:
    - sleduje WIP(work in progress) v SQL,
    - meri "ticho" na vstupu,
    - povoli stop az po vice potvrzenych cyklech.
    """
    def __init__(self, logger, bus):
        self.logger, self.bus, self.known_routes = logger, bus, {}
        # Inicializace stavových proměnných pro tracking
        self.active_wip = {}
        self.last_max_id = 0
        self.last_input_time = datetime.now(TZ)
        self.current_shift_start = None
        self.stop_confirm_hits = 0
        self.cycle_counter = 0

    def load_resources(self):
        # Volitelne nacteni routingu (cache pro dalsi rozsireni logiky).
        try:
            path = CONFIG["cooling"]["routes_cfg_path"]
            if os.path.exists(path):
                with open(path, 'r') as f:
                    self.known_routes = json.load(f)
        except Exception as e:
            self.logger.error("Cooling Resource Error: " + str(e))

    def analyze(self, now, sched, temps):
        """
        Vrací True = Pokračovat v topení, False = Povolit stop (vše hotovo).
        """
        # Gatekeeper pred koncem smeny: rozhoduje, jestli je bezpecne zastavit topeni.
        if not CONFIG["cooling"]["enabled"]:
            return True

        self.cycle_counter += 1
        win = get_shift_window(sched["hours"], now)
        if not win:
            return True

        # 1. RESET PŘI ZMĚNĚ SMĚNY
        if self.current_shift_start != win[0]:
            self.current_shift_start = win[0]
            self.active_wip = {}
            self.last_max_id = 0
            self.last_input_time = now
            self.stop_confirm_hits = 0
            self.logger.info("=== NOVA SMENA: {} | Reset cooling trackingu ===".format(win[0]))

        # 2. KONTROLA OKNA PRO CHLAZENÍ (jen před koncem směny)
        rem = (win[1] - now).total_seconds() / 60.0
        trigger_limit = CONFIG["cooling"]["trigger_min_before_end"]

        if not (0 < rem <= trigger_limit):
            # Mimo okno chlazení - resetujeme pomocné proměnné a topíme dál
            self.active_wip = {}
            self.stop_confirm_hits = 0
            return True

        # 3. SQL SCAN ROZPRACOVANÉ VÝROBY (WIP)
        if self.bus.cn:
            try:
                with self.bus.cn.cursor() as cur:
                    # Inkrementalni scan: od minuleho last_max_id, jinak od zacatku smeny.
                    # SQL dotaz pro Python 3.4 (použití .format místo f-strings)
                    shift_start_sql = win[0].strftime("%Y-%m-%d %H:%M:%S")
                    done_positions = set([int(x) for x in CONFIG["cooling"].get("done_positions", [15, 16, 17, 18, 19])])
                    if self.last_max_id > 0:
                        sql_filter = "v.IDP > {}".format(self.last_max_id)
                    else:
                        day_offset = int(CONFIG["cooling"].get("vstup_day_offset", -2))
                        sql_filter = "v.vstup IS NOT NULL AND DATEADD(day, {0}, CAST(v.vstup AS datetime)) >= '{1}'".format(
                            day_offset, shift_start_sql
                        )

                    # IDP a czbozi se získávají z tabulky vany/zbozi
                    query = ("SELECT v.pozice, v.IDP, CAST(v.czbozi AS VARCHAR(64)) "
                             "FROM dbo.vany v WITH (NOLOCK) "
                             "WHERE " + sql_filter)

                    cur.execute(query)
                    rows = cur.fetchall()

                    for row in rows:
                        try:
                            curr_pos = int(row[0]) if row[0] is not None else None
                        except Exception:
                            curr_pos = None
                        try:
                            if row[1] is None:
                                continue
                            idp = int(row[1])
                        except Exception:
                            continue
                        czbozi = str(row[2]).strip() if row[2] is not None else ""

                        if idp > self.last_max_id:
                            self.last_max_id = idp

                        if not czbozi:
                            continue
                        if curr_pos in done_positions:
                            done_keys = [k for k in self.active_wip if k[0] == czbozi]
                            for k in done_keys:
                                self.active_wip.pop(k, None)
                            continue

                        # Aktualizace času posledního vstupu
                        self.last_input_time = now

                        # Tracking konkrétního kusu zboží (včetně jeho pozice a IDP)
                        if (czbozi, idp) in self.active_wip:
                            self.active_wip[(czbozi, idp)].update({
                                "last_seen_ts": now,
                                "last_pos": curr_pos,
                                "last_idp": idp
                            })
                        else:
                            self.active_wip[(czbozi, idp)] = {
                                "last_seen_ts": now,
                                "last_pos": curr_pos,
                                "last_idp": idp
                            }
            except Exception as e:
                self.logger.warning("[COOLING] WIP scan failed: " + str(e))
                return True

        # 4. ČIŠTĚNÍ STALÉHO ZBOŽÍ (Flow TTL)
        # Kusy dlouho nevidene v toku odmazeme, aby netvorily falesny WIP.
        ttl_seconds = CONFIG["cooling"].get("flow_ttl_minutes", 10) * 60
        to_remove = []
        for item_key, data in self.active_wip.items():
            if (now - data["last_seen_ts"]).total_seconds() > ttl_seconds:
                to_remove.append(item_key)

        for item_key in to_remove:
            self.active_wip.pop(item_key, None)

        # 5. VYHODNOCENÍ STOP PODMÍNKY (Ticho + Prázdno)
        quiet_limit = CONFIG["cooling"]["quiet_time_minutes"] * 60
        is_quiet = (now - self.last_input_time).total_seconds() >= quiet_limit

        # Pokud je ticho a v lince není žádné aktivní zboží
        if is_quiet and len(self.active_wip) == 0:
            self.stop_confirm_hits += 1
        else:
            self.stop_confirm_hits = 0

        # Potvrzení stopu po více cyklech (prevence náhodných výpadků SQL)
        # Multi-hit potvrzeni chrani pred nahodnym "tichym" cyklem.
        confirm_cycles = max(1, int(CONFIG["cooling"].get("stop_confirm_cycles", 2)))
        if self.stop_confirm_hits > confirm_cycles:
            self.stop_confirm_hits = confirm_cycles
        allow_stop = self.stop_confirm_hits >= confirm_cycles

        # Logování každé 3 cykly pro přehled
        if self.cycle_counter % 3 == 0:
            self.logger.info(
                "[COOLING] WIP: {} ks | Ticho: {:.1f}/{} min | StopConfirm: {}/{} | Status: {}"
                .format(len(self.active_wip),
                        (now - self.last_input_time).total_seconds() / 60.0,
                        CONFIG["cooling"]["quiet_time_minutes"],
                        self.stop_confirm_hits, confirm_cycles,
                        "STOP_ALLOWED" if allow_stop else "RUNNING")
            )

        return not allow_stop

class PerformanceTracker(object):
    """
    Sledovani kvality predikce:
    - pri prechodu prikazu 0->1 ulozi snapshot (start_temp + ETA),
    - po dosazeni targetu ulozi realny cas a odchylku do SQL reportu.
    """
    def __init__(self, logger, bus):
        self.logger, self.bus = logger, bus
        self.tracking_data = {}

    def trigger_snapshot(self, now, temps, get_eta_fn, active_positions, label):
        # Snapshot reprezentuje predikci v okamziku startu topeni.
        tracked = 0
        detail_lines = []
        for pos in active_positions:
            t_now = temps.get(pos)
            if t_now is None:
                continue
            target = CONFIG["positions"][pos]["T_target"]
            if pos not in self.tracking_data:
                self.tracking_data[pos] = []
            preds = []
            for m in ["XGB", "MLP"]:
                eta = get_eta_fn(m, pos, t_now, target)
                if eta is None:
                    continue
                preds.append({"model": m, "eta": max(0.0, float(eta))})
            if preds:
                self.tracking_data[pos].append({
                    "start_time": now,
                    "start_temp": float(t_now),
                    "predictions": preds,
                    "hits": 0
                })
                tracked += 1
                eta_xgb = None
                eta_mlp = None
                eta_max = None
                for pred in preds:
                    if pred["model"] == "XGB":
                        eta_xgb = float(pred["eta"])
                    elif pred["model"] == "MLP":
                        eta_mlp = float(pred["eta"])
                    if eta_max is None or float(pred["eta"]) > eta_max:
                        eta_max = float(pred["eta"])
                dx = float(target) - float(t_now)
                xgb_txt = "{:.1f}".format(eta_xgb) if eta_xgb is not None else "-"
                mlp_txt = "{:.1f}".format(eta_mlp) if eta_mlp is not None else "-"
                max_txt = "{:.1f}".format(eta_max) if eta_max is not None else "-"
                detail_lines.append(
                    "[PERF] DETAIL {} | T={:.1f}C | T_target={:.1f}C | dT={:.1f}C | ETA_XGB={}m | ETA_MLP={}m | ETA_MAX={}m".format(
                        pos, float(t_now), float(target), dx, xgb_txt, mlp_txt, max_txt
                    )
                )
        self.logger.info("[PERF] START {} | tracked_positions={}".format(label, tracked))
        for line in detail_lines:
            self.logger.info(line)

    def check_arrivals(self, now, temps):
        # Arrival potvrzujeme az po vice hitech, aby vysledek nebyl citlivy na sum.
        confirm_hits = int(CONFIG.get("arrival_confirmation_hits", 2))
        ttl_hours = float(CONFIG.get("perf_snapshot_ttl_hours", 120))
        ttl_seconds = max(1.0, ttl_hours * 3600.0)
        for pos in list(self.tracking_data.keys()):
            curr_t = temps.get(pos)
            target_t = CONFIG["positions"][pos]["T_target"]
            keep = []
            dropped_ttl = 0
            for trk in self.tracking_data[pos]:
                age_sec = (now - trk["start_time"]).total_seconds()
                if age_sec > ttl_seconds:
                    dropped_ttl += 1
                    continue
                if curr_t is None:
                    keep.append(trk)
                    continue
                if curr_t >= target_t:
                    trk["hits"] += 1
                else:
                    trk["hits"] = 0
                if trk["hits"] >= confirm_hits:
                    dur = (now - trk["start_time"]).total_seconds() / 60.0
                    for p in trk["predictions"]:
                        self.bus.write_result(now, pos, trk["start_temp"], dur, p["model"], p["eta"], dur - p["eta"])
                else:
                    keep.append(trk)
            self.tracking_data[pos] = keep
            if dropped_ttl > 0:
                self.logger.warning("[PERF] TTL drop: pos={} dropped={} (> {}h)".format(pos, dropped_ttl, ttl_hours))


class Predictor(object):
    """
    Hlavni orchestrator:
    - nacita modely,
    - vyhodnocuje smeny + AI predzatop,
    - zapisuje SET_ZATOP do command queue.
    """
    def __init__(self, logger):
        self.logger = logger
        self.bus = SqlBus(logger)
        self.models = {}
        self.cool = CoolingManager(logger, self.bus)
        self.tracker = PerformanceTracker(logger, self.bus)
        self.last_zatop, self.is_latched = None, False
        self.latch_until = None
        self.latch_shift_start = None
        self.latch_shift_end = None
        self.sql_error_since = None
        self.last_temp_state = None
        self.last_temp_state_log_ts = None
        self.consecutive_loop_errors = 0
        self.last_schedule_signature = None
        self.last_shift_start = None
        self.was_in_shift = False
        self.last_active_pos_set = None
        self.last_decision_signature = None

    def load_models(self):
        # Chybejici model jedne pozice nesmi zastavit celou aplikaci.
        c = CONFIG["models"]
        for p in CONFIG["positions"]:
            m_p, x_p = os.path.join(c["mlp_dir"], p + c["mlp_suffix"]), os.path.join(c["xgb_dir"], p + c["xgb_suffix"])
            if os.path.exists(m_p):
                self.models["MLP_" + p] = ManualMLPInference(m_p)
            else:
                self.logger.warning("[MODEL] Missing MLP model: " + m_p)
            if os.path.exists(x_p):
                self.models["XGB_" + p] = ManualXGBInference(x_p)
            else:
                self.logger.warning("[MODEL] Missing XGB model: " + x_p)
        self.cool.load_resources()

    def get_eta(self, m, p, t_now, t_target):
        # Feature vektor je [T_now, T_target, delta].
        try:
            if t_now is None or t_target is None:
                return 0.0
            if float(t_now) >= float(t_target):
                return 0.0
        except Exception:
            return 0.0
        key = m + "_" + p
        if key in self.models:
            sin_t, cos_t = time_features_from_datetime(datetime.now(TZ))
            return self.models[key].predict([t_now, t_target, t_target - t_now, sin_t, cos_t])
        return 0

    def get_active_positions(self, now, temps):
        # Preferuje DB detekci; fallback je aktualni live teplota.
        if not CONFIG.get("active_bath_check", {}).get("enabled", True):
            return list(CONFIG["positions"].keys())
        if self.bus.cn:
            try:
                c = CONFIG["sql"]
                lookback = now - timedelta(hours=80 if now.weekday() in [0, 5, 6] else 30)
                sel = ", ".join(["MAX({})".format(v) for v in c["temp_columns"].values()])
                q = "SELECT {} FROM {} WHERE {} > ?".format(sel, c["temp_table"], c["time_column"])
                with self.bus.cn.cursor() as cur:
                    # Win7/older ODBC compatibility: prefer naive datetime binding.
                    lookback_sql = lookback.replace(tzinfo=None)
                    try:
                        row = cur.execute(q, (lookback_sql,)).fetchone()
                    except Exception:
                        # Fallback path for strict drivers: bind as string + CAST.
                        q2 = "SELECT {} FROM {} WHERE {} > CAST(? AS DATETIME)".format(
                            sel, c["temp_table"], c["time_column"])
                        row = cur.execute(q2, (lookback_sql.strftime("%Y-%m-%d %H:%M:%S"),)).fetchone()
                    active = []
                    keys = list(c["temp_columns"].keys())
                    for i, k in enumerate(keys):
                        if not row:
                            continue
                        raw_val = row[i]
                        if raw_val is None:
                            continue
                        try:
                            v = float(raw_val)
                        except Exception:
                            continue
                        if v >= CONFIG["positions"][k]["T_min_active"]:
                            active.append(k)
                    if active:
                        return active
            except Exception as e:
                self.logger.warning("[ACTIVE] Detector fallback: " + str(e))
        active_live = []
        for p in CONFIG["positions"]:
            t = temps.get(p)
            if t is not None and t >= CONFIG["positions"][p]["T_min_active"]:
                active_live.append(p)
        return active_live if active_live else list(CONFIG["positions"])

    def _clear_latch(self):
        self.is_latched = False
        self.latch_until = None
        self.latch_shift_start = None
        self.latch_shift_end = None

    def _build_schedule_signature(self, sched_raw):
        if not sched_raw:
            return None
        h = sched_raw.get("hours", {})
        parts = []
        for k in DAY_KEYS:
            parts.append("{}={:d}:{:d}->{:d}:{:d}".format(
                k,
                _to_int_safe(h.get(k + "_zap", -1), -1),
                _to_int_safe(h.get(k + "_zap_min", 0), 0),
                _to_int_safe(h.get(k + "_vyp", -1), -1),
                _to_int_safe(h.get(k + "_vyp_min", 0), 0),
            ))
        parts.append("auto={}".format(int(bool(sched_raw.get("auto")))))
        parts.append("buf={}".format(int(sched_raw.get("buf", CONFIG.get("default_buffer", 30)))))
        parts.append("cool={}".format(int(bool(sched_raw.get("cooling_on", True)))))
        return "|".join(parts)

    def _get_next_shift_window(self, hours, now):
        for s, e in _build_merged_shift_windows(hours, now, 1, 8):
            if now < s:
                return (s, e)
        return None

    def _find_window_for_timestamp(self, hours, ts, now):
        if ts is None:
            return None
        for s, e in _build_merged_shift_windows(hours, now, 2, 8):
            if s <= ts < e:
                return (s, e)
        return None

    def _calc_latch_until(self, now, shift_start, shift_end, cooling_on):
        # Drzime preheat do zacatku cooling okna; bez cooling do konce smeny.
        if not CONFIG.get("predictive", {}).get("hold_after_predictive_start", True):
            return shift_start
        if cooling_on and CONFIG.get("cooling", {}).get("enabled", False):
            trig = int(CONFIG["cooling"].get("trigger_min_before_end", 40))
            cool_start = shift_end - timedelta(minutes=trig)
            if cool_start > now:
                return cool_start
        return shift_end

    def run_forever(self):
        # Jedina smycka, ktera muze vytvaret nove povely v command queue.
        self.logger.info("=== POL1 PROD WIN7 V40 START ===")
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000001)
        except Exception as e:
            self.logger.warning("[WINAPI] SetThreadExecutionState failed: " + str(e))
        self.load_models()
        self.bus.connect()

        while True:
            try:
                if not self.bus.is_ready():
                    self.bus.connect()
                now = datetime.now(TZ)
                sched_raw = self.bus.read_schedule()
                temps, temp_state, temp_detail = self.read_temps()
                if sched_raw and self.last_zatop is None and sched_raw.get("zatop_db") is not None:
                    self.last_zatop = bool(sched_raw.get("zatop_db"))
                state_changed = (temp_state != self.last_temp_state)
                should_emit_state_log = state_changed
                if (not should_emit_state_log) and (self.last_temp_state_log_ts is not None):
                    should_emit_state_log = (now - self.last_temp_state_log_ts).total_seconds() >= 300

                if temp_state == "ok":
                    self.sql_error_since = None
                    self.bus.update_heartbeat()
                    if state_changed and self.last_temp_state is not None:
                        self.logger.info("[SAFETY] Temperature source recovered (state=OK).")
                    self.last_temp_state_log_ts = None
                elif temp_state == "sql_error":
                    grace = int(CONFIG.get("sql_error_grace_seconds", 240))
                    if self.sql_error_since is None:
                        self.sql_error_since = now
                    elapsed = (now - self.sql_error_since).total_seconds()
                    if elapsed <= grace:
                        # Kratky grace rezim: zmrazime rozhodovani, ale drzim heartbeat.
                        if should_emit_state_log:
                            self.logger.warning(
                                "[SAFETY] Temps SQL_ERROR ({}). Freeze rozhodovani {}/{}s."
                                .format(temp_detail, int(elapsed), grace))
                        if self.bus.is_ready():
                            self.bus.update_heartbeat()
                    else:
                        if should_emit_state_log:
                            self.logger.error(
                                "[SAFETY] Temps SQL_ERROR > {}s ({}); heartbeat suspended for failover."
                                .format(grace, temp_detail))
                        self._clear_latch()
                        self.last_zatop = None
                elif temp_state == "no_rows":
                    # NoRows = tvrdy fail-safe: heartbeat neaktualizujeme.
                    self.sql_error_since = None
                    if should_emit_state_log:
                        self.logger.error("[SAFETY] Empty temperature source (NO_ROWS); heartbeat suspended for failover.")
                    self._clear_latch()
                    self.last_zatop = None
                elif temp_state == "all_invalid":
                    # Vsechny vzorky jsou null/mimo sanity rozsah -> fail-safe.
                    self.sql_error_since = None
                    if should_emit_state_log:
                        self.logger.error("[SAFETY] Temperature source ALL_INVALID; heartbeat suspended for failover.")
                    self._clear_latch()
                    self.last_zatop = None

                if temp_state != "ok" and should_emit_state_log:
                    self.last_temp_state_log_ts = now
                self.last_temp_state = temp_state

                if sched_raw and temps and temp_state == "ok":
                    sched_sig = self._build_schedule_signature(sched_raw)
                    schedule_changed = False
                    if self.last_schedule_signature is None:
                        self.last_schedule_signature = sched_sig
                    elif sched_sig != self.last_schedule_signature:
                        self.logger.info("[SCHEDULE] ZMENA HARMONOGRAMU detekovana.")
                        self.logger.info("[SCHEDULE] OLD: {}".format(self.last_schedule_signature))
                        self.logger.info("[SCHEDULE] NEW: {}".format(sched_sig))
                        self.last_schedule_signature = sched_sig
                        schedule_changed = True

                    win = get_shift_window(sched_raw["hours"], now)
                    in_shift = True if win else False
                    if in_shift:
                        if self.last_shift_start != win[0]:
                            self.logger.info("[SHIFT] NOVA SMENA: {} -> {}".format(win[0], win[1]))
                        self.last_shift_start = win[0]
                        self.was_in_shift = True
                    else:
                        if self.was_in_shift:
                            self.logger.info("[SHIFT] MIMO SMENU")
                        self.was_in_shift = False

                    active_pos = self.get_active_positions(now, temps)
                    active_set = set(active_pos)
                    if self.last_active_pos_set is None:
                        self.last_active_pos_set = set(active_set)
                    elif active_set != self.last_active_pos_set:
                        became_inactive = sorted(list(self.last_active_pos_set - active_set))
                        became_active = sorted(list(active_set - self.last_active_pos_set))
                        if became_inactive:
                            self.logger.info("[ACTIVE] Neaktivni vany: {}".format(", ".join(became_inactive)))
                        if became_active:
                            self.logger.info("[ACTIVE] Nove aktivni vany: {}".format(", ".join(became_active)))
                        self.last_active_pos_set = set(active_set)

                    self.tracker.check_arrivals(now, temps)
                    should_start, reason = False, ""
                    next_win = self._get_next_shift_window(sched_raw["hours"], now)
                    next_s = next_win[0] if next_win else None
                    next_e = next_win[1] if next_win else None
                    max_eta = 0.0

                    if schedule_changed and self.is_latched:
                        old_until = self.latch_until
                        target_win = self._find_window_for_timestamp(
                            sched_raw["hours"], self.latch_shift_start, now)
                        if target_win is None and win is not None:
                            target_win = win
                        if target_win is None and next_win is not None and self.latch_shift_start is not None:
                            if next_win[0].date() == self.latch_shift_start.date():
                                target_win = next_win

                        if target_win is None:
                            self.logger.info("[LATCH] CLEAR: target shift missing in NEW schedule.")
                            self._clear_latch()
                        else:
                            self.latch_shift_start = target_win[0]
                            self.latch_shift_end = target_win[1]
                            self.latch_until = self._calc_latch_until(
                                now,
                                self.latch_shift_start,
                                self.latch_shift_end,
                                sched_raw.get("cooling_on", True))
                            old_until_txt = old_until.strftime("%Y-%m-%d %H:%M:%S") if old_until else "-"
                            new_until_txt = self.latch_until.strftime("%Y-%m-%d %H:%M:%S") if self.latch_until else "-"
                            self.logger.info(
                                "[LATCH] RECALC from NEW schedule: shift={}..{} old_until={} new_until={}".format(
                                    self.latch_shift_start, self.latch_shift_end,
                                    old_until_txt, new_until_txt))

                    if self.is_latched and self.latch_until is not None and now >= self.latch_until:
                        self.logger.info("[LATCH] RELEASE: now={} >= until={}".format(now, self.latch_until))
                        self._clear_latch()

                    if not sched_raw["auto"]:
                        should_start, reason = False, "MANUAL MODE"
                        self._clear_latch()
                        self.last_zatop = None
                    elif in_shift:
                        should_start, reason = True, "Active Shift"
                        if sched_raw["cooling_on"] and not self.cool.analyze(now, sched_raw, temps):
                            should_start, reason = False, "COOLING GUARD: VYPÍNÁM"
                    else:
                        latch_active = self.is_latched and (self.latch_until is None or now < self.latch_until)
                        if latch_active:
                            should_start = True
                            if self.latch_until is None:
                                reason = "LATCH HOLD"
                            else:
                                reason = "LATCH HOLD until {}".format(self.latch_until.strftime("%Y-%m-%d %H:%M:%S"))
                        elif not next_win:
                            should_start, reason = False, "No future shift"
                            self._clear_latch()
                        else:
                            eta_cap = float(CONFIG.get("predictive", {}).get("eta_cap_minutes", 720.0))
                            max_eta = 0.0
                            for pos in active_pos:
                                t_now = temps.get(pos)
                                if t_now is None:
                                    continue
                                t_tgt = CONFIG["positions"][pos]["T_target"]
                                eta = max(self.get_eta("XGB", pos, t_now, t_tgt),
                                          self.get_eta("MLP", pos, t_now, t_tgt))
                                try:
                                    eta = float(eta)
                                except Exception:
                                    eta = 0.0
                                if eta < 0:
                                    eta = 0.0
                                if eta > eta_cap:
                                    eta = eta_cap
                                if eta > max_eta:
                                    max_eta = eta

                            if max_eta > 0 and now >= next_s - timedelta(minutes=(max_eta + sched_raw["buf"])):
                                should_start = True
                                reason = "AI Start: {:.0f}m".format(max_eta)
                                self.is_latched = True
                                self.latch_shift_start = next_s
                                self.latch_shift_end = next_e
                                self.latch_until = self._calc_latch_until(
                                    now, next_s, next_e, sched_raw.get("cooling_on", True))
                                self.logger.info("[LATCH] SET: shift={}..{} hold_until={}".format(
                                    self.latch_shift_start, self.latch_shift_end, self.latch_until))
                            else:
                                reason = "Waiting"

                    if CONFIG.get("output", {}).get("decision_trace_every_cycle", True):
                        latch_until_txt = self.latch_until.strftime("%Y-%m-%d %H:%M:%S") if self.latch_until else "-"
                        next_shift_txt = next_s.strftime("%Y-%m-%d %H:%M:%S") if next_s else "-"
                        decision_sig = (
                            bool(sched_raw["auto"]),
                            bool(in_shift),
                            bool(self.is_latched),
                            latch_until_txt,
                            next_shift_txt,
                            round(float(max_eta), 1),
                            int(len(active_pos)),
                            bool(should_start),
                            str(reason),
                        )
                        if decision_sig != self.last_decision_signature:
                            self.last_decision_signature = decision_sig
                            self.logger.info(
                                "[DECISION] auto={} in_shift={} latched={} latch_until={} next_shift={} "
                                "max_eta={:.1f} active={} should_start={} reason={}".format(
                                    sched_raw["auto"], in_shift, self.is_latched, latch_until_txt,
                                    next_shift_txt, float(max_eta), len(active_pos), should_start, reason)
                            )

                    current_zatop = sched_raw.get("zatop_db")
                    if current_zatop is None:
                        current_zatop = self.last_zatop
                    else:
                        current_zatop = bool(current_zatop)

                    need_change = (current_zatop != should_start)
                    pending_same = False
                    if sched_raw["auto"] and need_change:
                        pending_same = self.bus.has_pending_zatop(should_start)

                    recent_error_same = False
                    if sched_raw["auto"] and need_change and not pending_same:
                        recent_error_same = self.bus.has_recent_error_zatop(should_start)

                    if sched_raw["auto"] and need_change and not pending_same and not recent_error_same:
                        prev_zatop = current_zatop
                        if self.bus.enqueue_zatop(should_start):
                            self.last_zatop = should_start
                            if should_start and prev_zatop is not True:
                                self.tracker.trigger_snapshot(now, temps, self.get_eta, active_pos, "CMD_0_TO_1")
                                inactive_now = [p for p in CONFIG["positions"] if p not in active_set]
                                self.logger.info(
                                    "[ACTIVE] TOP 0->1 | Aktivni vany: {} | Neaktivni vany: {}".format(
                                        ", ".join(active_pos) if active_pos else "-",
                                        ", ".join(inactive_now) if inactive_now else "-")
                                )
                            self.logger.info("ZMĚNA: {} ({})".format(should_start, reason))
                        else:
                            self.logger.warning("QUEUE FAIL: {} ({}) - retry next cycle".format(should_start, reason))
                    elif sched_raw["auto"] and need_change and not pending_same and recent_error_same:
                        self.logger.warning(
                            "QUEUE COOLDOWN: skip retry SET_ZATOP={} for {}s".format(
                                should_start, int(CONFIG["sql_bus"].get("error_retry_cooldown_seconds", 120))
                            )
                        )
                self.consecutive_loop_errors = 0

            except Exception as e:
                self.consecutive_loop_errors += 1
                self.logger.error("Loop Error: " + str(e))
                sr = CONFIG.get("self_restart", {})
                if sr.get("enabled", True):
                    max_err = int(sr.get("max_consecutive_loop_errors", 10))
                    if self.consecutive_loop_errors >= max_err:
                        self.logger.error("[SELF-RESTART] Too many loop errors ({}/{}), exiting for scheduler restart."
                                          .format(self.consecutive_loop_errors, max_err))
                        raise SystemExit(2)

            if CONFIG.get("countdown_enabled", False):
                for i in range(CONFIG["poll_seconds"], 0, -1):
                    sys.stdout.write("\r[ {} ] Dalsi kontrola za {:d}s... ".format(
                        datetime.now().strftime("%H:%M:%S"), i))
                    sys.stdout.flush()
                    time.sleep(1)
            else:
                time.sleep(CONFIG["poll_seconds"])

    def find_next_shift(self, hours, now):
        # Vrati nejblizsi budouci start smeny.
        for s, _ in _build_merged_shift_windows(hours, now, 1, 8):
            if now < s:
                return s
        return None

    def read_temps(self):
        if not self.bus.cn:
            return {}, "sql_error", "DB not connected"
        # Median poslednich N vzorku snizuje vliv nahodnych ústřelů teplot.
        c = CONFIG["sql"]
        lim = CONFIG.get("temp_sanity_limits", {})
        t_min = float(lim.get("min", -9999.0))
        t_max = float(lim.get("max", 9999.0))
        q = "SELECT TOP {} {}, {} FROM {} ORDER BY {} DESC".format(
            c["median_window"], c["time_column"], ", ".join(c["temp_columns"].values()), c["temp_table"],
            c["time_column"])
        try:
            with self.bus.cn.cursor() as cur:
                rows = cur.execute(q).fetchall()
                if not rows:
                    return {}, "no_rows", "No rows in source table"
                temps = {}
                for i, k in enumerate(c["temp_columns"].keys()):
                    v = [r[i + 1] for r in rows if r[i + 1] is not None and t_min <= float(r[i + 1]) <= t_max]
                    if v: temps[k] = float(np.median(v))
                if not temps:
                    return {}, "all_invalid", "All samples null/out-of-range"
                return temps, "ok", ""
        except Exception as e:
            self.logger.error("[SQL] read_temps failed: " + str(e))
            self.bus.drop_connection()
            return {}, "sql_error", str(e)

class AdoCursor(object):
    def __init__(self, ado_conn, command_timeout=None):
        self._ado_conn = ado_conn
        self._command_timeout = command_timeout
        self._rs = None
        self._rows = None
        self.description = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    def _sql_literal(self, value):
        if value is None:
            return "NULL"
        if isinstance(value, bool):
            return "1" if value else "0"
        if isinstance(value, int):
            return str(value)
        if isinstance(value, float):
            return repr(value)
        if isinstance(value, datetime):
            return "'" + value.strftime("%Y-%m-%d %H:%M:%S") + "'"
        text = str(value).replace("'", "''")
        return "'" + text + "'"

    def _bind_params(self, sql, params):
        if not params:
            return sql
        out = sql
        for p in params:
            out = out.replace("?", self._sql_literal(p), 1)
        return out

    def execute(self, sql, params=None):
        q = self._bind_params(sql, params)
        if self._command_timeout is not None:
            try:
                self._ado_conn.CommandTimeout = int(self._command_timeout)
            except Exception:
                pass
        result = self._ado_conn.Execute(q)
        if isinstance(result, (tuple, list)):
            self._rs = result[0]
        else:
            self._rs = result
        self._rows = None
        self.description = []
        try:
            if self._rs and hasattr(self._rs, "Fields"):
                cnt = self._rs.Fields.Count
                for i in range(cnt):
                    self.description.append((self._rs.Fields(i).Name, None, None, None, None, None, None))
        except Exception:
            self.description = []
        return self

    def _materialize(self):
        if self._rows is not None:
            return self._rows
        rows = []
        if not self._rs:
            self._rows = rows
            return rows
        try:
            while not self._rs.EOF:
                cnt = self._rs.Fields.Count
                rows.append(tuple(self._rs.Fields(i).Value for i in range(cnt)))
                self._rs.MoveNext()
        except Exception:
            pass
        self._rows = rows
        return rows

    def fetchone(self):
        rows = self._materialize()
        if not rows:
            return None
        return rows[0]

    def fetchall(self):
        return self._materialize()

    def close(self):
        try:
            if self._rs:
                self._rs.Close()
        except Exception:
            pass
        self._rs = None


class AdoConnection(object):
    def __init__(self, ado_conn, command_timeout=None):
        self._ado_conn = ado_conn
        self._command_timeout = command_timeout
        self.autocommit = True

    def cursor(self):
        return AdoCursor(self._ado_conn, self._command_timeout)

    def Execute(self, sql):
        return self._ado_conn.Execute(sql)

    def close(self):
        try:
            self._ado_conn.Close()
        except Exception:
            pass


class SqlBus(object):
    """
    SQL vrstva aplikace:
    - connect/reconnect,
    - heartbeat,
    - cteni scheduleru,
    - zapis commandu a performance metrik.
    """
    def __init__(self, logger):
        self.logger, self.cn = logger, None

    def drop_connection(self):
        if self.cn is not None:
            try:
                self.cn.close()
            except Exception:
                pass
        self.cn = None

    def connect(self):
        # Pri chybe se cn nastavi na None; reconnect probiha v hlavni smycce.
        try:
            self.drop_connection()
            c = CONFIG['sql']
            timeout = int(CONFIG.get("sql_timeout", 40))
            cs = (
                "Provider=SQLOLEDB.1;Password={};Persist Security Info=True;"
                "User ID={};Initial Catalog={};Data Source={};"
            ).format(c['password'], c['username'], c['database'], c['server'])
            ado_conn = win32com.client.Dispatch("ADODB.Connection")
            try:
                ado_conn.ConnectionTimeout = timeout
            except Exception:
                pass
            try:
                ado_conn.CommandTimeout = timeout
            except Exception:
                pass
            ado_conn.Open(cs)
            self.cn = AdoConnection(ado_conn, timeout)
            self.logger.info("[SQL] Connected")
        except Exception as e:
            self.logger.error("[SQL] Connect Error: " + str(e))
            self.drop_connection()

    def is_ready(self):
        return self.cn is not None

    def update_heartbeat(self):
        # Keep-alive pro nadrazeny system (ID je 1 ve state table).
        if not self.cn: return
        try:
            with self.cn.cursor() as cur:
                cur.execute(
                    "UPDATE {} SET PythonHeartbeat=GETDATE() WHERE Id=1".format(CONFIG['sql_bus']['state_table']))
        except Exception as e:
            self.logger.error("[SQL] update_heartbeat failed: " + str(e))
            self.drop_connection()

    def read_schedule(self):
        # Nacte auto/manual, buffer a casy smen ze state table.
        if not self.cn: return None
        try:
            with self.cn.cursor() as cur:
                cur.execute("SELECT TOP 1 * FROM {} WHERE Id=1".format(CONFIG['sql_bus']['state_table']))
                row = cur.fetchone()
                if not row: return None
                cols = [c[0] for c in cur.description]
                default_buf = int(CONFIG.get("default_buffer", 30))
                if "BufferMinutes" in cols:
                    raw_buf = row[cols.index("BufferMinutes")]
                else:
                    raw_buf = default_buf
                try:
                    if raw_buf is None:
                        buf = default_buf
                    else:
                        txt = str(raw_buf).strip().replace(",", ".")
                        if txt == "":
                            buf = default_buf
                        else:
                            buf = int(float(txt))
                except Exception as e:
                    self.logger.warning("[SQL] BufferMinutes cast failed, using default: " + str(e))
                    buf = default_buf
                if buf < 0:
                    buf = default_buf
                vals = {}
                for k in DAY_KEYS:
                    vals[k + "_zap"] = row[cols.index(k + "_zap")]
                    vals[k + "_vyp"] = row[cols.index(k + "_vyp")]
                    vals[k + "_zap_min"] = row[cols.index(k + "_zap_min")] if k + "_zap_min" in cols else 0
                    vals[k + "_vyp_min"] = row[cols.index(k + "_vyp_min")] if k + "_vyp_min" in cols else 0
                raw_zatop = row[cols.index("zatop")] if "zatop" in cols else None
                raw_auto = row[cols.index("auto_man")] if "auto_man" in cols else 0
                raw_cooling = row[cols.index("CoolingEnabled")] if "CoolingEnabled" in cols else 1
                auto_mode = bool(_to_bit_safe(raw_auto, 0))
                cooling_on = bool(_to_bit_safe(raw_cooling, 1))
                if raw_zatop is None:
                    zatop_db = None
                else:
                    zatop_db = bool(_to_bit_safe(raw_zatop, 0))
                return {"hours": vals, "auto": auto_mode,
                        "buf": buf,
                        "cooling_on": cooling_on,
                        "zatop_db": zatop_db}
        except Exception as e:
            self.logger.error("[SQL] read_schedule failed: " + str(e))
            self.drop_connection()
            return None

    def write_result(self, now, pos, start_temp, actual, model, pred, diff):
        # Ulozi vyhodnoceni kvality modelu po potvrzenem arrival.
        if not self.cn:
            return
        try:
            with self.cn.cursor() as cur:
                q = "INSERT INTO {} (ReportDate, Position, StartTemp, TargetTemp, PredictedMinutes, ActualMinutes, DiffMinutes, AiModelUsed) VALUES (?,?,?,?,?,?,?,?)".format(
                    CONFIG['sql_bus']['report_table'])
                data = (now.date(), pos, float(start_temp), float(CONFIG["positions"][pos]["T_target"]),
                        int(round(float(pred))), int(round(float(actual))), int(round(float(diff))), model)
                cur.execute(q, data)
        except Exception as e:
            self.logger.error("[SQL] write_result failed: " + str(e))
            self.drop_connection()

    def enqueue_zatop(self, val):
        # Vlozi prikaz SET_ZATOP do command queue.
        if not self.cn: return False
        try:
            with self.cn.cursor() as cur:
                q = "INSERT INTO {} (TsCreated, CmdType, ValueBit, Status, IssuedBy) VALUES (GETDATE(), 'SET_ZATOP', ?, 'NEW', ?)".format(
                    CONFIG['sql_bus']['cmd_table'])
                cur.execute(q, (1 if val else 0, CONFIG['sql_bus']['issued_by']))
            return True
        except Exception as e:
            self.logger.error("[SQL] enqueue_zatop failed: " + str(e))
            self.drop_connection()
            return False

    def has_pending_zatop(self, val):
        # Vrati True, pokud uz v queue existuje cekajici stejny prikaz.
        # Stary INPROGRESS po TTL ignorujeme (neblokuje novy enqueue).
        if not self.cn:
            return False
        try:
            ttl_min = int(CONFIG.get("sql_bus", {}).get("pending_inprogress_ttl_minutes", 10))
            if ttl_min < 1:
                ttl_min = 1
            with self.cn.cursor() as cur:
                q = ("SELECT TOP 1 1 FROM {} "
                     "WHERE CmdType='SET_ZATOP' AND ValueBit=? AND ("
                     "Status='NEW' OR (Status='INPROGRESS' AND "
                     "ISNULL(TsStarted, TsCreated) >= DATEADD(minute, -?, GETDATE()))) "
                     "ORDER BY TsCreated DESC").format(CONFIG['sql_bus']['cmd_table'])
                row = cur.execute(q, (1 if val else 0, ttl_min)).fetchone()
                return row is not None
        except Exception as e:
            self.logger.error("[SQL] has_pending_zatop failed: " + str(e))
            self.drop_connection()
            return False

    def has_recent_error_zatop(self, val):
        # Brani opakovanemu enqueue stejneho prikazu hned po ERROR.
        if not self.cn:
            return False
        try:
            cooldown = int(CONFIG.get("sql_bus", {}).get("error_retry_cooldown_seconds", 120))
            if cooldown <= 0:
                return False
            with self.cn.cursor() as cur:
                q = ("SELECT TOP 1 1 FROM {} "
                     "WHERE CmdType='SET_ZATOP' AND ValueBit=? AND Status='ERROR' "
                     "AND ISNULL(TsFinished, TsCreated) >= DATEADD(second, -?, GETDATE()) "
                     "ORDER BY ISNULL(TsFinished, TsCreated) DESC").format(CONFIG['sql_bus']['cmd_table'])
                row = cur.execute(q, (1 if val else 0, cooldown)).fetchone()
                return row is not None
        except Exception as e:
            self.logger.error("[SQL] has_recent_error_zatop failed: " + str(e))
            self.drop_connection()
            return False


if __name__ == "__main__":
    Predictor(setup_logger(CONFIG["output"]["log_file"])).run_forever()
