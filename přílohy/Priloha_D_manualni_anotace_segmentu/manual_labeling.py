from __future__ import annotations

from pathlib import Path
from typing import List, Set
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import pyodbc
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.widgets import SpanSelector
import matplotlib.dates as mdates
import configparser
import re


# ==========================
# KONFIGURACE
# ==========================

SQL_SERVER = r"XXX"
SQL_DATABASE = "XXXX"
SQL_TABLE = "dbo.teploty_pol1"

USE_TRUSTED_CONNECTION = False
SQL_USERNAME = "XXXX"
SQL_PASSWORD = "XXXX"

SQL_DRIVER = "{ODBC Driver 17 for SQL Server}"

TIME_COLUMN = "Cas"


# ==========================
# Kterou pozici chceš labelovat (sloupec v DB)
# vyber pozice2, 3, 5, 7, 10, 14
TANK_COLUMN = "pozice14"
# Původní FROM_DATETIME použijeme jen jako default,
# pokud v cílové složce ještě nic není:
FROM_DATETIME = "2025-10-25 00:00:00"
# TO_DATETIME už NEBUDEME používat jako konstantu,
# vždy se dopočítá na dnešní den 23:59:59.
# Začátek „dne“ (shift) – 14:00 do 14:00 další den
SHIFT_HOUR = 12  # 14:00
# ==========================

PROJECT_ROOT = Path(__file__).parent

# Kam ukládat ručně označené úseky
OUTPUT_DIR = PROJECT_ROOT / "NN" / "data" / "train_manual"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# CONFIG.cfg s Tset/Tmin (sekce [TEMPATURE], klíče pozice2 = 50, …
CONFIG_PATH = PROJECT_ROOT / "CONFIG.cfg"

# Zakážeme defaultní klávesové zkratky Matplotlibu, aby nám „s“ neotvíral save dialog
plt.rcParams["keymap.save"] = []
plt.rcParams["keymap.quit"] = []
plt.rcParams["keymap.fullscreen"] = []
plt.rcParams["keymap.home"] = []
plt.rcParams["keymap.back"] = []
plt.rcParams["keymap.forward"] = []


# ==========================
# Pomocné funkce: Tmin z CONFIG.cfg
# ==========================

def load_tmin_from_config() -> float | None:
    """
    Z CONFIG.cfg / [TEMPERATURE] načte hodnotu pro danou pozici.
    Použijeme ji jako 'Tmin' do titulku grafu.
    """
    if not CONFIG_PATH.exists():
        print(f"[WARN] CONFIG.cfg ({CONFIG_PATH}) neexistuje – Tmin nebude v titulku.")
        return None

    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding="utf-8")

    section = "TEMPERATURE"
    if section not in cfg:
        print(f"[WARN] V CONFIG.cfg chybí sekce [{section}] – Tmin nebude v titulku.")
        return None

    key = TANK_COLUMN.lower()
    if key not in cfg[section]:
        print(f"[WARN] V CONFIG.cfg není položka '{key}' – Tmin nebude v titulku.")
        return None

    try:
        return float(cfg[section][key])
    except ValueError:
        print(f"[WARN] Hodnota '{key}' v CONFIG.cfg nejde převést na float.")
        return None


# ==========================
# DB PŘIPOJENÍ
# ==========================

def get_connection() -> pyodbc.Connection:
    if USE_TRUSTED_CONNECTION:
        conn_str = (
            f"DRIVER={SQL_DRIVER};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={SQL_DATABASE};"
            "Trusted_Connection=yes;"
        )
    else:
        conn_str = (
            f"DRIVER={SQL_DRIVER};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={SQL_DATABASE};"
            f"UID={SQL_USERNAME};"
            f"PWD={SQL_PASSWORD};"
        )

    print("[INFO] Connection string:")
    print(conn_str.replace(SQL_PASSWORD, "***"))
    return pyodbc.connect(conn_str)


# ==========================
# NAČTENÍ DAT Z DB
# ==========================

def load_data_from_db(
    tank_col: str,
    from_dt: str,
    to_dt: str,
) -> pd.DataFrame:
    """
    Načte z DB data pro jeden tank (sloupec) v daném časovém rozsahu.
    Vrací DF se sloupci: Tank, Cas (datetime), Temp (float)
    """
    query = f"""
        SELECT {TIME_COLUMN}, {tank_col}
        FROM {SQL_TABLE}
        WHERE {TIME_COLUMN} >= ? AND {TIME_COLUMN} <= ?
        ORDER BY {TIME_COLUMN} ASC
    """

    print("[INFO] Spouštím dotaz na DB:")
    print(query)
    print(f"[INFO] Parametry: FROM={from_dt}, TO={to_dt}")

    with get_connection() as conn:
        df = pd.read_sql(query, conn, params=[from_dt, to_dt])

    if df.empty:
        print("[WARN] Z DB se nenačetla žádná data v daném rozsahu.")
        return pd.DataFrame(columns=["Tank", "Cas", "Temp"])

    # převod času
    df[TIME_COLUMN] = pd.to_datetime(df[TIME_COLUMN], errors="coerce")
    df = df.dropna(subset=[TIME_COLUMN])

    # převod teploty na float (i s čárkou)
    df[tank_col] = (
        df[tank_col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .replace("nan", np.nan)
        .astype(float)
    )

    df = df.sort_values(TIME_COLUMN).reset_index(drop=True)

    df["Tank"] = tank_col
    df.rename(columns={TIME_COLUMN: "Cas", tank_col: "Temp"}, inplace=True)
    return df[["Tank", "Cas", "Temp"]]


# ==========================
# GROUP 14:00 → 14:00
# ==========================

def split_into_shift_days(df: pd.DataFrame, shift_hour: int = 12):
    """
    Rozdělí DF na bloky 12:00 → 12:00 následujícího dne.

    Logika:
      - Pro každý řádek určím "den směny" takto:
          base_day = floor(Cas na půlnoc)
          pokud Cas.hour < SHIFT_HOUR, patří ještě do předchozího dne směny.
      - Všechna data s day_key = D tvoří okno [D 12:00, D+1 12:00).
    """
    cas = df["Cas"]
    base_day = cas.dt.floor("D")
    mask_before_shift = cas.dt.hour < shift_hour
    day_key = base_day.copy()
    day_key[mask_before_shift] = day_key[mask_before_shift] - pd.Timedelta(days=1)

    windows: List[tuple[pd.Timestamp, pd.DataFrame]] = []
    for day, df_day in df.groupby(day_key):
        df_day = df_day.sort_values("Cas").reset_index(drop=True)
        windows.append((day, df_day))
    windows.sort(key=lambda x: x[0])
    return windows


# ==========================
# ZJIŠTĚNÍ DNŮ, KTERÉ UŽ JSOU NALABELOVANÉ
# ==========================

def get_already_labeled_shift_days_for_pozice(base_name: str) -> Set[date]:
    """
    Projde OUTPUT_DIR a najde soubory typu
      pozice2_02_11_2025_05.txt
    ale místo parsování data z názvu se podívá na PRVNÍ řádek souboru,
    přečte timestamp, a z něj spočítá "shift day" stejnou logikou
    jako při dělení 12:00 → 12:00:

        shift_day = (timestamp - SHIFT_HOUR).date()

    Tak se správně trefíme i na případy, kdy náběh začal po půlnoci.
    """
    labeled_shift_days: Set[date] = set()

    for path in OUTPUT_DIR.glob(f"{base_name}_*.txt"):
        try:
            with path.open("r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    parts = line.split()
                    if len(parts) < 3:
                        continue
                    ts_str = parts[0] + " " + parts[1]  # "YYYY-MM-DD HH:MM:SS"
                    try:
                        ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        continue

                    # stejná logika jako ve split_into_shift_days:
                    shift_day = (ts - timedelta(hours=SHIFT_HOUR)).date()
                    labeled_shift_days.add(shift_day)
                    break  # první validní řádek stačí
        except OSError:
            continue

    return labeled_shift_days


# ==========================
# VÝPOČET DYNAMICKÉHO FROM_DATETIME
# ==========================

def compute_dynamic_from_datetime(
    labeled_shift_days: Set[date],
    default_from: str,
    shift_hour: int,
) -> str:
    """
    Určí, od kdy tahat data z DB.

    - pokud ještě nemáš žádné nalabelované dny -> vrátí default_from
    - jinak vezme max(labeled_shift_days) a vrátí datetime
      pro TENTO shift den v čase shift_hour:00

      např. last_day = 2025-12-07, shift_hour = 14
      → FROM_DATETIME = "2025-12-07 14:00:00"
    """
    if not labeled_shift_days:
        print(f"[INFO] V cílové složce zatím nejsou žádné segmenty – používám FROM_DATETIME={default_from}")
        return default_from

    last_day = max(labeled_shift_days)  # date posledního shift dne

    from_dt = datetime(
        year=last_day.year,
        month=last_day.month,
        day=last_day.day,
        hour=shift_hour,
        minute=0,
        second=0,
    )
    from_str = from_dt.strftime("%Y-%m-%d %H:%M:%S")

    print(
        f"[INFO] Nejpozdější již nalabelovaný shift den: {last_day}. "
        f"Nová FROM_DATETIME = {from_str}"
    )
    return from_str


# ==========================
# INTERAKTIVNÍ LABELER
# ==========================

class WindowLabeler:
    def __init__(
        self,
        tank_name: str,
        output_dir: Path,
        tmin_value: float | None,
        start_counter: int = 1,
    ):
        self.tank_name = tank_name          # např. "Pozice2"
        self.base_name = tank_name.lower()  # "pozice2"
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.global_index = start_counter   # číslování segmentů (1,2,3,…)
        self.selected_bounds = None         # (i_min, i_max) v indexech
        self.quit_all = False
        self.saved_files: List[Path] = []

        self.tmin = tmin_value

    def label_window(self, day_key: pd.Timestamp, df_window: pd.DataFrame) -> bool:
        """
        Zobrazí jedno okno (14:00–14:00), nechá uživatele vybrat úsek
        a případně uloží. Vrací True pokud pokračovat, False pokud ukončit všechno.
        """
        if df_window.empty:
            return True

        x_dt = df_window["Cas"].to_numpy()
        y = df_window["Temp"].to_numpy()

        x_num = mdates.date2num(x_dt)

        fig, ax = plt.subplots(figsize=(10, 5))

        start_time = day_key + pd.Timedelta(hours=SHIFT_HOUR)
        end_time = day_key + pd.Timedelta(days=1, hours=SHIFT_HOUR)

        # Tset (cílová teplota pro heating) z CONFIG.cfg
        if self.tmin is not None:
            tset_text = f"Tset={self.tmin:.1f} °C"
        else:
            tset_text = "Tset=? °C"

        fig.suptitle(
            f"{self.tank_name} ({tset_text})\n"
            f"okno {start_time} → {end_time}"
        )

        ax.plot(x_dt, y)
        ax.set_xlabel("Čas")
        ax.set_ylabel("Teplota (°C)")
        ax.grid(True)

        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d\n%H:%M"))
        fig.autofmt_xdate()

        text_help = (
            "Myší označ úsek (táhnutím) → pak:\n"
            "  's' = uložit vybraný úsek\n"
            "  'n' = další okno (bez uložení)\n"
            "  'u' = undo posledně uložený segment (smazat soubor)\n"
            "  'q' = ukončit skript"
        )
        fig.text(0.01, 0.01, text_help, fontsize=9, va="bottom")

        self.selected_bounds = None

        def on_select(xmin, xmax):
            lo = min(xmin, xmax)
            hi = max(xmin, xmax)

            mask = (x_num >= lo) & (x_num <= hi)
            idxs = np.where(mask)[0]
            if len(idxs) == 0:
                print("[WARN] Ve vybraném úseku nejsou žádná data.")
                self.selected_bounds = None
                return

            i_min = int(idxs[0])
            i_max = int(idxs[-1])
            self.selected_bounds = (i_min, i_max)

            print(
                f"[INFO] Vybrán úsek indexů {i_min}–{i_max} "
                f"({df_window['Cas'].iloc[i_min]} → {df_window['Cas'].iloc[i_max]})"
            )

        span = SpanSelector(
            ax,
            on_select,
            "horizontal",
            useblit=True,
            interactive=True,
            drag_from_anywhere=True,
        )

        def on_key(event):
            if event.key == "s":
                if self.selected_bounds is None:
                    print("[WARN] Nejprve označ úsek myší.")
                    return
                self.save_selection(df_window)
            elif event.key == "n":
                print("[INFO] Okno přeskočeno, další 14–14.")
                plt.close(fig)
            elif event.key == "u":
                self.undo_last_save()
            elif event.key == "q":
                print("[INFO] Ukončuji dle požadavku uživatele (q).")
                self.quit_all = True
                plt.close(fig)

        fig.canvas.mpl_connect("key_press_event", on_key)

        plt.tight_layout()
        plt.show()

        if self.quit_all:
            return False
        return True

    def save_selection(self, df_window: pd.DataFrame):
        i_min, i_max = self.selected_bounds
        segment = df_window.iloc[i_min: i_max + 1].copy()

        first_ts = segment["Cas"].iloc[0]
        day_label = first_ts.strftime("%d_%m_%Y")  # dd_mm_yyyy

        out_name = f"{self.base_name}_{day_label}_{self.global_index:02d}.txt"
        out_path = self.output_dir / out_name

        with out_path.open("w", encoding="utf-8") as f:
            for _, row in segment.iterrows():
                ts = row["Cas"]
                ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
                temp = float(row["Temp"])
                f.write(f"{ts_str} {temp:.2f}\n")

        print(
            f"[INFO] Uložen segment #{self.global_index} → {out_path} "
            f"({len(segment)} vzorků; "
            f"{segment['Cas'].iloc[0]} → {segment['Cas'].iloc[-1]})"
        )

        self.saved_files.append(out_path)
        self.global_index += 1

    def undo_last_save(self):
        if not self.saved_files:
            print("[INFO] Není co undo – ještě nic nebylo uloženo v tomto běhu.")
            return

        last_path = self.saved_files.pop()
        if last_path.exists():
            try:
                last_path.unlink()
                print(f"[INFO] Undo: smazán soubor {last_path}")
            except Exception as e:
                print(f"[WARN] Nepodařilo se smazat {last_path}: {e}")
        else:
            print(f"[WARN] Undo: soubor {last_path} už neexistuje.")

        if self.global_index > 1:
            self.global_index -= 1
            print(f"[INFO] Čítač segmentů vrácen na {self.global_index}")


# ==========================
# MAIN
# ==========================

def main():
    base_name = TANK_COLUMN.lower()

    # 1) zjistit, jaké shift dny už jsou nalabelované
    labeled_shift_days = get_already_labeled_shift_days_for_pozice(base_name)
    today = date.today()

    print("[INFO] Již nalabelované shift dny (podle timestampů v souborech):")
    for d in sorted(labeled_shift_days):
        print("  ", d)

    # 2) spočítat dynamické FROM_DATETIME podle již existujících segmentů
    dynamic_from = compute_dynamic_from_datetime(
        labeled_shift_days=labeled_shift_days,
        default_from=FROM_DATETIME,
        shift_hour=SHIFT_HOUR,
    )

    # 2b) automatické TO_DATETIME = dnešní den 23:59:59
    today_end = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
    dynamic_to = today_end.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[INFO] Dynamické TO_DATETIME = {dynamic_to}")

    # 3) načíst z DB jen nové období
    print(f"[INFO] Načítám data pro {TANK_COLUMN} z DB…")
    df = load_data_from_db(TANK_COLUMN, dynamic_from, dynamic_to)

    if df.empty:
        print("[ERROR] Z DB nepřišla žádná nová data, končím.")
        return

    print(f"[INFO] Načteno {len(df)} řádků.")

    print("[INFO] Rozděluji na okna 14:00 → 14:00…")
    windows = split_into_shift_days(df, SHIFT_HOUR)
    print(f"[INFO] Počet oken: {len(windows)}")

    # Tmin z CONFIG.cfg (pro titulek)
    tmin_value = load_tmin_from_config()
    if tmin_value is not None:
        print(f"[INFO] Tmin pro {TANK_COLUMN} z CONFIG.cfg: {tmin_value} °C")

    labeler = WindowLabeler(TANK_COLUMN, OUTPUT_DIR, tmin_value, start_counter=1)

    for idx, (day_key, df_window) in enumerate(windows, start=1):
        # datum směny = datum start_time (14:00)
        start_time = day_key + pd.Timedelta(hours=SHIFT_HOUR)
        end_time = day_key + pd.Timedelta(days=1, hours=SHIFT_HOUR)
        shift_date = start_time.date()

        # přeskoč budoucí dny (po dnešku)
        if shift_date > today:
            print(
                f"[INFO] Přeskakuji budoucí okno {idx}/{len(windows)} "
                f"({start_time} → {end_time}), den {shift_date} > dnešek {today}."
            )
            continue

        print(
            f"\n[INFO] Okno {idx}/{len(windows)}: "
            f"{start_time} → {end_time} "
            f"({len(df_window)} vzorků, den směny {shift_date})"
        )

        cont = labeler.label_window(day_key, df_window)
        if not cont:
            break

    print("\n[INFO] Hotovo.")


if __name__ == "__main__":
    main()
