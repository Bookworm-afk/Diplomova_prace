"""
generate_baseline_snapshot.py

Creates a baseline snapshot under NN/reports with:
1) segment counts per position/mode
2) segment lengths (min/median/max) per position/mode
3) current eval metrics (MAE, p90, p95) for existing models
   + signed error metrics (bias/median + under/over-predict ratio)
4) model inventory with hash and timestamps

This script is read-only for datasets/models. It only writes report files.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple
import configparser
import hashlib
import json
import statistics

import joblib
import numpy as np
import xgboost as xgb
from segment_split_manifest import load_manifest, get_split_file_paths


PROJECT_ROOT = Path(__file__).parent
CONFIG_PATH = PROJECT_ROOT / "CONFIG.cfg"

DATA_MODE = "augmented"  # "augmented" -> train_allclean_DeltaT, "raw" -> train_allclean
if DATA_MODE == "augmented":
    TRAIN_DATA = {
        "heating": PROJECT_ROOT / "NN" / "data" / "train_allclean_DeltaT" / "heating",
        "cooling": PROJECT_ROOT / "NN" / "data" / "train_allclean_DeltaT" / "cooling",
    }
elif DATA_MODE == "raw":
    TRAIN_DATA = {
        "heating": PROJECT_ROOT / "NN" / "data" / "train_allclean" / "heating",
        "cooling": PROJECT_ROOT / "NN" / "data" / "train_allclean" / "cooling",
    }
else:
    raise ValueError("Unsupported DATA_MODE={!r}. Use 'augmented' or 'raw'.".format(DATA_MODE))

MODEL_DIRS = {
    ("MLP", "heating"): PROJECT_ROOT / "NN" / "models_MLP" / "heating",
    ("MLP", "cooling"): PROJECT_ROOT / "NN" / "models_MLP" / "cooling",
    ("XGB", "heating"): PROJECT_ROOT / "NN" / "models_XGBoost" / "heating",
    ("XGB", "cooling"): PROJECT_ROOT / "NN" / "models_XGBoost" / "cooling",
}

MODEL_SUFFIX = {
    ("MLP", "heating"): "time_to_target_mlp.joblib",
    ("MLP", "cooling"): "time_to_tmin_mlp.joblib",
    ("XGB", "heating"): "time_to_target_xgb.json",
    ("XGB", "cooling"): "time_to_tmin_xgb.json",
}

FRACTIONS = {
    "heating": [1.0, 0.9, 0.7, 0.4, 0.2],
    "cooling": [0.9, 0.7, 0.4, 0.2, 0.1],
}
USE_SPLIT_MANIFEST_FOR_EVAL = True
SPLIT_MANIFEST_DATA_MODE = DATA_MODE
EVAL_SPLIT = "test"
SPLIT_MANIFEST_PATH = PROJECT_ROOT / "NN" / "splits" / f"segment_split_manifest_{SPLIT_MANIFEST_DATA_MODE}.json"


@dataclass
class SegmentData:
    temps: np.ndarray
    dt_min: float


def load_temperature_map() -> Dict[str, float]:
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding="utf-8")
    section = "TEMPERATURE"
    if section not in cfg:
        raise RuntimeError(f"Missing [{section}] in {CONFIG_PATH}")
    result: Dict[str, float] = {}
    for key, value in cfg[section].items():
        result[key.strip().lower()] = float(value)
    return result


def parse_position(file_path: Path) -> str:
    name = file_path.name.lower()
    return name.split("_", 1)[0] if "_" in name else ""


def load_segment(path: Path) -> SegmentData:
    times: List[datetime] = []
    temps: List[float] = []

    with path.open("r", encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split()
            # Supports both raw and DeltaT-augmented rows; temperature is always column index 2.
            if len(parts) < 3:
                continue
            ts_str = f"{parts[0]} {parts[1]}"
            try:
                ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
                t = float(parts[2].replace(",", "."))
            except ValueError:
                continue
            times.append(ts)
            temps.append(t)

    if len(temps) < 2:
        return SegmentData(temps=np.array([], dtype=float), dt_min=0.0)

    dt_sec = (times[1] - times[0]).total_seconds()
    if dt_sec <= 0:
        dt_sec = 60.0
    return SegmentData(temps=np.array(temps, dtype=float), dt_min=dt_sec / 60.0)


def compute_segment_stats() -> Dict[str, Dict[str, dict]]:
    stats: Dict[str, Dict[str, dict]] = {"heating": {}, "cooling": {}}

    for mode, folder in TRAIN_DATA.items():
        files = sorted(folder.glob("*.txt")) if folder.exists() else []
        by_pos: Dict[str, List[int]] = {}

        for file_path in files:
            pos = parse_position(file_path)
            if not pos.startswith("pozice"):
                continue
            seg = load_segment(file_path)
            if seg.temps.size == 0:
                continue
            by_pos.setdefault(pos, []).append(int(seg.temps.size))

        for pos, lengths in sorted(by_pos.items()):
            stats[mode][pos] = {
                "count_segments": int(len(lengths)),
                "length_samples_min": int(min(lengths)),
                "length_samples_median": float(statistics.median(lengths)),
                "length_samples_max": int(max(lengths)),
            }

    return stats


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def collect_model_inventory() -> Dict[str, List[dict]]:
    inventory: Dict[str, List[dict]] = {}
    for (model_type, mode), folder in MODEL_DIRS.items():
        key = f"{model_type}_{mode}"
        inventory[key] = []
        if not folder.exists():
            continue
        for model_path in sorted(folder.iterdir()):
            if not model_path.is_file():
                continue
            st = model_path.stat()
            inventory[key].append(
                {
                    "file": model_path.name,
                    "size_bytes": int(st.st_size),
                    "modified_local": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                    "sha256": sha256_file(model_path),
                }
            )
    return inventory


class XGBBoosterAdapter:
    def __init__(self, booster: xgb.Booster):
        self.booster = booster

    def predict(self, x: np.ndarray) -> np.ndarray:
        arr = np.asarray(x, dtype=float)
        if arr.ndim == 1:
            arr = arr.reshape(1, -1)
        dm = xgb.DMatrix(arr)
        pred = self.booster.predict(dm)
        return np.asarray(pred, dtype=float)


def load_model(model_type: str, mode: str, position: str):
    model_dir = MODEL_DIRS[(model_type, mode)]
    suffix = MODEL_SUFFIX[(model_type, mode)]
    path = model_dir / f"{position}_{suffix}"
    if not path.exists():
        return None

    if model_type == "MLP":
        return joblib.load(path)

    # Prefer sklearn wrapper; fallback to Booster for xgboost/sklearn incompatibilities.
    try:
        model = xgb.XGBRegressor()
        model.load_model(path)
        return model
    except TypeError as exc:
        if "_estimator_type" not in str(exc):
            raise
        booster = xgb.Booster()
        booster.load_model(str(path))
        return XGBBoosterAdapter(booster)


def build_feature(mode: str, t_target: float, t_now: float) -> np.ndarray:
    if mode == "heating":
        delta_t = t_target - t_now
    else:
        delta_t = t_now - t_target
    return np.array([[float(t_now), float(t_target), float(delta_t)]], dtype=float)


def resolve_eval_files(mode: str, position: str, manifest: dict | None) -> List[Path]:
    folder = TRAIN_DATA[mode]
    if manifest is None:
        return sorted(folder.glob(f"{position}_*.txt")) if folder.exists() else []
    return get_split_file_paths(manifest, mode, position, EVAL_SPLIT, folder)


def evaluate_position(
    mode: str,
    model_type: str,
    position: str,
    t_target: float,
    manifest: dict | None,
) -> Tuple[List[float], List[float], int]:
    model = load_model(model_type, mode, position)
    if model is None:
        return [], [], 0

    files = resolve_eval_files(mode, position, manifest)
    abs_errors: List[float] = []
    signed_errors: List[float] = []

    for file_path in files:
        seg = load_segment(file_path)
        temps = seg.temps
        dt_min = seg.dt_min
        if temps.size < 3 or dt_min <= 0:
            continue

        target_idx = int(temps.size - 1)
        if target_idx < 3:
            continue
        total_steps = target_idx

        for frac in FRACTIONS[mode]:
            steps_remaining = int(round(total_steps * frac))
            if steps_remaining <= 0:
                continue
            i = target_idx - steps_remaining
            if i < 0:
                i = 0
                steps_remaining = target_idx
            if i >= target_idx:
                continue

            t_now = float(temps[i])
            y_true = float(steps_remaining * dt_min)
            x = build_feature(mode, t_target, t_now)
            y_pred = float(model.predict(x)[0])
            if y_pred < 0:
                y_pred = 0.0
            signed_err = float(y_pred - y_true)
            signed_errors.append(signed_err)
            abs_errors.append(abs(signed_err))

    return abs_errors, signed_errors, len(files)


def summarize_abs_errors(abs_errors: List[float]) -> dict:
    if not abs_errors:
        return {"n_samples": 0, "mae_min": None, "p90_min": None, "p95_min": None}

    arr = np.array(abs_errors, dtype=float)
    return {
        "n_samples": int(arr.size),
        "mae_min": float(arr.mean()),
        "p90_min": float(np.percentile(arr, 90)),
        "p95_min": float(np.percentile(arr, 95)),
    }


def summarize_signed_errors(signed_errors: List[float]) -> dict:
    if not signed_errors:
        return {
            "bias_min": None,
            "median_signed_min": None,
            "p10_signed_min": None,
            "p90_signed_min": None,
            "underpredict_pct": None,
            "overpredict_pct": None,
        }

    arr = np.array(signed_errors, dtype=float)
    under = float(np.mean(arr < 0.0) * 100.0)
    over = float(np.mean(arr > 0.0) * 100.0)
    return {
        "bias_min": float(arr.mean()),
        "median_signed_min": float(np.percentile(arr, 50)),
        "p10_signed_min": float(np.percentile(arr, 10)),
        "p90_signed_min": float(np.percentile(arr, 90)),
        "underpredict_pct": under,
        "overpredict_pct": over,
    }


def compute_eval_metrics(temp_map: Dict[str, float], manifest: dict | None) -> Dict[str, dict]:
    results: Dict[str, dict] = {}
    for model_type in ("MLP", "XGB"):
        for mode in ("heating", "cooling"):
            key = f"{model_type}_{mode}"
            per_position: Dict[str, dict] = {}
            all_abs_errors: List[float] = []
            all_signed_errors: List[float] = []

            for position, t_target in sorted(temp_map.items()):
                abs_errors, signed_errors, n_segments = evaluate_position(mode, model_type, position, t_target, manifest)
                position_metrics = summarize_abs_errors(abs_errors)
                position_metrics.update(summarize_signed_errors(signed_errors))
                position_metrics["n_segments_scanned"] = int(n_segments)
                per_position[position] = position_metrics
                all_abs_errors.extend(abs_errors)
                all_signed_errors.extend(signed_errors)

            global_metrics = summarize_abs_errors(all_abs_errors)
            global_metrics.update(summarize_signed_errors(all_signed_errors))
            results[key] = {
                "global": global_metrics,
                "per_position": per_position,
            }

    return results


def build_summary_text(
    snapshot_dir: Path,
    segment_stats: Dict[str, Dict[str, dict]],
    eval_metrics: Dict[str, dict],
    model_inventory: Dict[str, List[dict]],
    eval_scope: str,
) -> str:
    lines: List[str] = []
    lines.append(f"Baseline snapshot: {snapshot_dir.name}")
    lines.append(f"Created: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"Data mode: {DATA_MODE}")
    lines.append(f"Train heating dir: {TRAIN_DATA['heating']}")
    lines.append(f"Train cooling dir: {TRAIN_DATA['cooling']}")
    lines.append("")
    lines.append(f"[Eval scope] {eval_scope}")
    lines.append("")

    lines.append("[Segment counts and lengths]")
    for mode in ("heating", "cooling"):
        lines.append(f"- {mode}:")
        if not segment_stats[mode]:
            lines.append("  no segments")
            continue
        total_segments = sum(v["count_segments"] for v in segment_stats[mode].values())
        lines.append(f"  positions={len(segment_stats[mode])}, total_segments={total_segments}")
    lines.append("")

    lines.append("[Current eval metrics (global)]")
    for key in ("MLP_heating", "MLP_cooling", "XGB_heating", "XGB_cooling"):
        g = eval_metrics.get(key, {}).get("global", {})
        lines.append(
            f"- {key}: N={g.get('n_samples', 0)}, "
            f"MAE={g.get('mae_min')}, p90={g.get('p90_min')}, p95={g.get('p95_min')}, "
            f"bias={g.get('bias_min')}, under%={g.get('underpredict_pct')}, over%={g.get('overpredict_pct')}"
        )
    lines.append("")

    lines.append("[Model inventory]")
    for key in ("MLP_heating", "MLP_cooling", "XGB_heating", "XGB_cooling"):
        lines.append(f"- {key}: {len(model_inventory.get(key, []))} files")

    return "\n".join(lines) + "\n"


def create_snapshot_dir() -> Path:
    reports_dir = PROJECT_ROOT / "NN" / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    base_name = f"baseline_{datetime.now().date().isoformat()}"
    target = reports_dir / base_name
    if target.exists():
        target = reports_dir / f"{base_name}_{datetime.now().strftime('%H%M%S')}"
    target.mkdir(parents=True, exist_ok=False)
    return target


def main() -> None:
    temp_map = load_temperature_map()
    snapshot_dir = create_snapshot_dir()

    manifest = None
    eval_scope = "all_segments"
    if USE_SPLIT_MANIFEST_FOR_EVAL:
        if SPLIT_MANIFEST_PATH.exists():
            manifest = load_manifest(SPLIT_MANIFEST_PATH)
            eval_scope = f"{EVAL_SPLIT}_split_from_manifest({SPLIT_MANIFEST_PATH.name})"
            print(f"[INFO] Eval metrics scope: {eval_scope}")
        else:
            print(f"[WARN] Split manifest not found: {SPLIT_MANIFEST_PATH}. Falling back to all segments.")

    segment_stats = compute_segment_stats()
    model_inventory = collect_model_inventory()
    eval_metrics = compute_eval_metrics(temp_map, manifest)

    (snapshot_dir / "segment_stats.json").write_text(
        json.dumps(segment_stats, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (snapshot_dir / "eval_metrics.json").write_text(
        json.dumps(eval_metrics, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (snapshot_dir / "model_inventory.json").write_text(
        json.dumps(model_inventory, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (snapshot_dir / "summary.txt").write_text(
        build_summary_text(snapshot_dir, segment_stats, eval_metrics, model_inventory, eval_scope),
        encoding="utf-8",
    )

    print(f"[INFO] Baseline snapshot created: {snapshot_dir}")
    print(f"[INFO] Files: segment_stats.json, eval_metrics.json, model_inventory.json, summary.txt")


if __name__ == "__main__":
    main()
