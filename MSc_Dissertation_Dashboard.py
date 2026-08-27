from __future__ import annotations

import json
import re
from pathlib import Path

import numpy as np
import pandas as pd
from plotly.offline import get_plotlyjs


ROOT = Path(__file__).resolve().parent
INPUT_DIR = ROOT / "output"
OUTPUT_HTML = ROOT / "VFD_Dissertation_Dashboard.html"
SIZE_MCR_FILE = ROOT / "Size_&_MCR.xlsx"
SHEET = "Combined_Report"
EQUIPMENT_CANDIDATES = ("SW1", "SW2", "SW3", "FAN1", "FAN2", "FAN3", "FAN4")


def normalise(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def find_general(columns: list[str], *aliases: str) -> str | None:
    indexed = {normalise(c): c for c in columns}
    for alias in aliases:
        key = normalise(alias)
        if key in indexed:
            return indexed[key]
    for alias in aliases:
        key = normalise(alias)
        for norm, original in indexed.items():
            if key and key in norm:
                return original
    return None


def find_equipment_field(columns: list[str], equipment: str, field: str) -> str | None:
    eq = normalise(equipment)
    endings = {
        "rh": ("rh", "runninghours", "runhours"),
        "auto": ("auto", "autohours"),
        "bypass": ("bypass", "bypasshours"),
        "load": ("hz", "hzpercent", "load", "averageload"),
        "actual": ("kwh", "actualenergy"),
        "baseline": ("max", "baseline", "fixedspeedbaseline"),
    }[field]
    for column in columns:
        key = normalise(column)
        if eq in key and any(key.endswith(suffix) for suffix in endings):
            return column
    return None


def numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def null_sum(series: pd.Series) -> float:
    return series.sum(min_count=1)


def weighted_average(values: pd.Series, weights: pd.Series) -> float:
    mask = values.notna() & weights.notna() & (weights > 0)
    if not mask.any():
        return np.nan
    return float(np.average(values[mask], weights=weights[mask]))


def report_weighted_average(group: pd.DataFrame, value_column: str) -> float:
    values = numeric(group[value_column])
    hours = numeric(group["report_hours"])
    weighted_mask = values.notna() & hours.notna() & (hours > 0)
    if weighted_mask.any():
        return float(np.average(values[weighted_mask], weights=hours[weighted_mask]))
    valid = values.dropna()
    return float(valid.mean()) if not valid.empty else np.nan


def safe_ratio_series(numerator: pd.Series, denominator: pd.Series, multiplier: float = 100.0) -> pd.Series:
    return np.where(denominator.notna() & denominator.ne(0), numerator / denominator * multiplier, np.nan)


def spearman_pair(frame: pd.DataFrame, x_column: str, y_column: str, minimum: int = 3) -> tuple[float, int]:
    pairs = frame[[x_column, y_column]].apply(numeric).dropna()
    if len(pairs) < minimum or pairs[x_column].nunique() < 2 or pairs[y_column].nunique() < 2:
        return np.nan, len(pairs)
    rho = pairs[x_column].rank(method="average").corr(pairs[y_column].rank(method="average"))
    return float(rho), len(pairs)


def parse_fan_midpoint(value: object) -> float:
    if pd.isna(value):
        return np.nan
    text = str(value).replace("–", "-").replace("—", "-")
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*(?:-|to)\s*(-?\d+(?:\.\d+)?)", text, re.I)
    return (float(match.group(1)) + float(match.group(2))) / 2 if match else np.nan


def classify_segment(value: object) -> str:
    text = str(value).strip().lower()
    if "port" in text or "harbour" in text or "harbor" in text:
        return "Port Stay"
    if "sea" in text or "passage" in text or "sailing" in text:
        return "Sea Passage"
    return "Other / Unknown"


def records(frame: pd.DataFrame) -> list[dict]:
    return json.loads(frame.to_json(orient="records", date_format="iso", date_unit="ms"))


def aggregate_equipment(frame: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    output: list[dict] = []
    for group_key, group in frame.groupby(keys, dropna=False, sort=True):
        values = group_key if isinstance(group_key, tuple) else (group_key,)
        row = dict(zip(keys, values))
        for col in ("running_hours", "auto_hours", "manual_hours", "bypass_hours",
                    "operating_hours", "actual_energy", "baseline_energy", "saving"):
            row[col] = null_sum(group[col])
        row["avg_load"] = weighted_average(group["avg_load"], group["running_hours"])
        for col in ("sea_temp", "me_load", "pid_setpoint", "fan_temp_mid"):
            row[col] = group[col].mean() if group[col].notna().any() else np.nan
        row["interval_count"] = int(group["interval_id"].nunique())
        output.append(row)
    return pd.DataFrame(output)


def build_multivariable_sw_daily(interval: pd.DataFrame) -> pd.DataFrame:
    sw = interval[interval["equipment_category"].eq("SW Pumps")]
    keys = ["vessel", "vessel_category", "vessel_size", "report_date", "month", "segment_type"]
    output: list[dict] = []
    for group_key, group in sw.groupby(keys, dropna=False, sort=True):
        reports = group.drop_duplicates("interval_id")
        sea_temp = report_weighted_average(reports, "sea_temp")
        pid_setpoint = report_weighted_average(reports, "pid_setpoint")
        me_load = report_weighted_average(reports, "me_load")
        mcr_values = reports["mcr_kw"].dropna()
        mcr_kw = float(mcr_values.iloc[0]) if not mcr_values.empty else np.nan
        me_load_pct_mcr = me_load / mcr_kw * 100 if pd.notna(me_load) and pd.notna(mcr_kw) and mcr_kw > 0 else np.nan
        running_hours = null_sum(group["running_hours"])
        operating_hours = null_sum(group["operating_hours"])
        actual_energy = null_sum(group["actual_energy"])
        baseline_energy = null_sum(group["baseline_energy"])
        saving = null_sum(group["saving"])
        cooling_margin = pid_setpoint - sea_temp if pd.notna(pid_setpoint) and pd.notna(sea_temp) else np.nan
        row = dict(zip(keys, group_key))
        row.update({
            "sea_temp": sea_temp,
            "pid_setpoint": pid_setpoint,
            "cooling_temperature_margin": cooling_margin,
            "sea_temp_pid_interaction": sea_temp * pid_setpoint if pd.notna(sea_temp) and pd.notna(pid_setpoint) else np.nan,
            "me_load_pct_mcr": me_load_pct_mcr,
            "cooling_margin_me_load_interaction": cooling_margin * me_load_pct_mcr if pd.notna(cooling_margin) and pd.notna(me_load_pct_mcr) else np.nan,
            "sea_temp_me_load_interaction": sea_temp * me_load_pct_mcr if pd.notna(sea_temp) and pd.notna(me_load_pct_mcr) else np.nan,
            "pid_me_load_interaction": pid_setpoint * me_load_pct_mcr if pd.notna(pid_setpoint) and pd.notna(me_load_pct_mcr) else np.nan,
            "avg_load": weighted_average(group["avg_load"], group["running_hours"]),
            "running_hours": running_hours,
            "manual_pct": null_sum(group["manual_hours"]) / operating_hours * 100 if pd.notna(operating_hours) and operating_hours != 0 else np.nan,
            "bypass_pct": null_sum(group["bypass_hours"]) / operating_hours * 100 if pd.notna(operating_hours) and operating_hours != 0 else np.nan,
            "saving_pct": saving / baseline_energy * 100 if pd.notna(baseline_energy) and baseline_energy != 0 else np.nan,
            "actual_energy": actual_energy,
            "baseline_energy": baseline_energy,
            "energy_per_running_hour": actual_energy / running_hours if pd.notna(running_hours) and running_hours != 0 else np.nan,
            "saving_per_running_hour": saving / running_hours if pd.notna(running_hours) and running_hours != 0 else np.nan,
        })
        output.append(row)
    return pd.DataFrame(output)


def build_parallel_imbalance(interval: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    meaningful = (
        interval["running_hours"].gt(0)
        | interval["avg_load"].notna()
        | interval["actual_energy"].notna()
        | interval["baseline_energy"].notna()
    )
    available_rows = interval[meaningful]
    available_units = available_rows[["vessel", "equipment"]].drop_duplicates()
    vessels = sorted(available_rows["vessel"].dropna().unique())
    output: list[dict] = []
    for (vessel, category), group in available_rows.groupby(["vessel", "equipment_category"], sort=True):
        unit_rows: list[dict] = []
        for equipment, equipment_group in group.groupby("equipment", sort=True):
            unit_rows.append({"equipment": equipment, "running_hours": null_sum(equipment_group["running_hours"])})
        run_values = [row for row in unit_rows if pd.notna(row["running_hours"])]
        imbalance = np.nan
        most_used = None
        least_used = None
        if len(unit_rows) >= 2 and len(run_values) == len(unit_rows):
            mean_hours = float(np.mean([row["running_hours"] for row in run_values]))
            if mean_hours > 0:
                ordered = sorted(run_values, key=lambda row: row["running_hours"])
                imbalance = (ordered[-1]["running_hours"] - ordered[0]["running_hours"]) / mean_hours
                most_used = ordered[-1]["equipment"]
                least_used = ordered[0]["equipment"]
        size_values = group["vessel_size"].dropna()
        output.append({
            "vessel": vessel,
            "vessel_size": size_values.iloc[0] if not size_values.empty else "Not specified",
            "equipment_category": category,
            "available_parallel_units": len(unit_rows),
            "normalised_running_hour_imbalance": imbalance,
            "most_used_unit": most_used,
            "least_used_unit": least_used,
        })
    imbalance_frame = pd.DataFrame(output)
    unavailable_cells = len(vessels) * len(EQUIPMENT_CANDIDATES) - len(available_units)
    return imbalance_frame, len(available_units), int(unavailable_cells)


def load_size_mcr() -> dict[str, dict[str, object]]:
    if not SIZE_MCR_FILE.is_file():
        raise FileNotFoundError(f"Size/MCR workbook not found: {SIZE_MCR_FILE}")
    try:
        source = pd.read_excel(SIZE_MCR_FILE, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Could not read {SIZE_MCR_FILE.name}: {exc}") from exc
    source.columns = [str(c).strip() for c in source.columns]
    columns = list(source.columns)
    indexed = {normalise(c): c for c in columns}
    vessel_col = next((indexed[normalise(a)] for a in ("Vessel", "Vessel Name", "Name") if normalise(a) in indexed), None)
    number_col = next((indexed[normalise(a)] for a in ("Vessel Number", "Vessel No", "Vessel ID") if normalise(a) in indexed), None)
    size_col = find_general(columns, "Size", "Vessel Size", "Vessel Category")
    mcr_col = find_general(columns, "MCR", "MCR kW", "Main Engine MCR")
    if not mcr_col or (not vessel_col and not number_col):
        raise ValueError("Size/MCR workbook needs an explicit vessel identifier and MCR column")
    lookup: dict[str, dict[str, object]] = {}
    for _, row in source.iterrows():
        identity = row[vessel_col] if vessel_col else row[number_col]
        if pd.isna(identity):
            continue
        if not vessel_col:
            number = pd.to_numeric(identity, errors="coerce")
            if pd.isna(number):
                continue
            identity = f"Vessel {int(number)}"
        lookup[normalise(identity)] = {
            "vessel_size": None if not size_col or pd.isna(row[size_col]) else str(row[size_col]).strip(),
            "mcr_kw": pd.to_numeric(row[mcr_col], errors="coerce"),
        }
    return lookup


def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    paths = [INPUT_DIR / f"VFD Savings Report - Vessel {i}.xlsx" for i in range(1, 15)]
    missing = [str(path) for path in paths if not path.is_file()]
    if missing:
        raise FileNotFoundError("Missing final workbooks: " + ", ".join(missing))

    size_mcr = load_size_mcr()
    interval_rows: list[dict] = []
    report_rows: list[dict] = []
    matched_vessels: list[str] = []
    unmatched_vessels: list[str] = []
    read_ok = 0
    for vessel_number, path in enumerate(paths, 1):
        vessel_name = f"Vessel {vessel_number}"
        vessel_info = size_mcr.get(normalise(vessel_name))
        if vessel_info and pd.notna(vessel_info.get("mcr_kw")) and float(vessel_info["mcr_kw"]) > 0:
            matched_vessels.append(vessel_name)
        else:
            unmatched_vessels.append(vessel_name)
        vessel_size = vessel_info.get("vessel_size") if vessel_info else None
        mcr_kw = vessel_info.get("mcr_kw") if vessel_info else np.nan
        try:
            source = pd.read_excel(path, sheet_name=SHEET, engine="openpyxl")
        except Exception as exc:
            raise RuntimeError(f"Could not read {SHEET!r} from {path.name}: {exc}") from exc
        read_ok += 1
        source.columns = [str(c).strip() for c in source.columns]
        columns = list(source.columns)
        general = {
            "segment": find_general(columns, "Segment Description", "Segment"),
            "date": find_general(columns, "Date Time", "Datetime", "Date"),
            "hours": find_general(columns, "Hours", "Report Hours"),
            "sea_temp": find_general(columns, "Sea Water Temperature", "Seawater Temperature"),
            "me_load": find_general(columns, "ME Load", "Main Engine Load"),
            "total_ae": find_general(columns, "Total AE kWh", "Total Auxiliary Engine kWh"),
            "pid": find_general(columns, "PID Setpoint", "PID_set_point", "SW - PID_set_point"),
            "fan_range": find_general(columns, "Fan ER Temperature Range", "ER Range Temp", "FAN - ER Range Temp."),
            "vessel_category": find_general(columns, "Vessel Category", "Vessel Size", "Size Category"),
        }
        if general["date"] is None:
            raise ValueError(f"No date column found in {path.name}")
        dates = pd.to_datetime(source[general["date"]], errors="coerce")
        if dates.notna().sum() == 0:
            raise ValueError(f"Dates could not be parsed in {path.name}")

        def series(name: str) -> pd.Series:
            column = general[name]
            return numeric(source[column]) if column else pd.Series(np.nan, index=source.index)

        segment_raw = source[general["segment"]] if general["segment"] else pd.Series(None, index=source.index)
        category_raw = source[general["vessel_category"]] if general["vessel_category"] else pd.Series(None, index=source.index)
        categories = category_raw.apply(
            lambda value: str(value).strip() if pd.notna(value) and str(value).strip().lower() not in {"", "not specified", "unknown", "nan"}
            else (vessel_size or "Not specified")
        )
        fan_raw = source[general["fan_range"]] if general["fan_range"] else pd.Series(None, index=source.index)
        raw_me_load = series("me_load")
        me_load_pct = raw_me_load / float(mcr_kw) * 100 if pd.notna(mcr_kw) and float(mcr_kw) > 0 else pd.Series(np.nan, index=source.index)
        common = pd.DataFrame({
            "date_time": dates,
            "report_date": dates.dt.strftime("%Y-%m-%d"),
            "month": dates.dt.strftime("%Y-%m"),
            "segment_type": segment_raw.map(classify_segment),
            "report_hours": series("hours"),
            "sea_temp": series("sea_temp"),
            "me_load": raw_me_load,
            "mcr_kw": mcr_kw,
            "me_load_pct_mcr": me_load_pct,
            "total_ae_energy": series("total_ae"),
            "pid_setpoint": series("pid"),
            "fan_temp_mid": fan_raw.map(parse_fan_midpoint),
            "vessel_category": categories,
            "vessel_size": vessel_size or "Not specified",
        })
        for source_index in source.index:
            if pd.isna(common.at[source_index, "date_time"]):
                continue
            interval_id = f"V{vessel_number}-{source_index + 2}"
            report_rows.append({
                "interval_id": interval_id, "vessel": vessel_name,
                **common.loc[source_index].to_dict(),
            })

        for equipment in EQUIPMENT_CANDIDATES:
            fields = {name: find_equipment_field(columns, equipment, name)
                      for name in ("rh", "auto", "bypass", "load", "actual", "baseline")}
            if not any(fields.values()):
                continue
            values = {name: numeric(source[column]) if column else pd.Series(np.nan, index=source.index)
                      for name, column in fields.items()}
            meaningful = ((values["rh"] > 0) | (values["load"] > 0) |
                          (values["actual"] > 0) | (values["baseline"] > 0)) & dates.notna()
            for source_index in source.index[meaningful]:
                rh, auto, bypass = values["rh"].at[source_index], values["auto"].at[source_index], values["bypass"].at[source_index]
                actual, baseline = values["actual"].at[source_index], values["baseline"].at[source_index]
                manual = max(rh - auto, 0) if pd.notna(rh) and pd.notna(auto) else np.nan
                operating = rh + bypass if pd.notna(rh) and pd.notna(bypass) else np.nan
                saving = baseline - actual if pd.notna(actual) and pd.notna(baseline) else np.nan
                interval_rows.append({
                    "interval_id": f"V{vessel_number}-{source_index + 2}",
                    "vessel": vessel_name,
                    "vessel_category": common.at[source_index, "vessel_category"],
                    "vessel_size": common.at[source_index, "vessel_size"],
                    "date_time": common.at[source_index, "date_time"],
                    "report_hours": common.at[source_index, "report_hours"],
                    "report_date": common.at[source_index, "report_date"],
                    "month": common.at[source_index, "month"],
                    "segment_type": common.at[source_index, "segment_type"],
                    "equipment_category": "SW Pumps" if equipment.startswith("SW") else "Engine-Room Fans",
                    "equipment": equipment,
                    "running_hours": rh, "auto_hours": auto, "manual_hours": manual,
                    "bypass_hours": bypass, "operating_hours": operating,
                    "avg_load": values["load"].at[source_index],
                    "actual_energy": actual, "baseline_energy": baseline, "saving": saving,
                    "sea_temp": common.at[source_index, "sea_temp"],
                    "me_load": common.at[source_index, "me_load"],
                    "mcr_kw": common.at[source_index, "mcr_kw"],
                    "me_load_pct_mcr": common.at[source_index, "me_load_pct_mcr"],
                    "pid_setpoint": common.at[source_index, "pid_setpoint"],
                    "fan_temp_mid": common.at[source_index, "fan_temp_mid"],
                    "total_ae_energy": common.at[source_index, "total_ae_energy"],
                })

    interval = pd.DataFrame(interval_rows)
    reports = pd.DataFrame(report_rows)
    if interval.empty:
        raise ValueError("No meaningful equipment observations were detected")
    daily_equipment = aggregate_equipment(interval, ["vessel", "vessel_category", "vessel_size", "report_date", "month", "equipment_category", "equipment"])
    daily_category = aggregate_equipment(interval, ["vessel", "vessel_category", "vessel_size", "report_date", "month", "equipment_category"])

    daily_vfd = interval.groupby(["vessel", "vessel_category", "vessel_size", "report_date", "month"], as_index=False).agg(
        total_vfd_energy=("actual_energy", null_sum), total_vfd_baseline=("baseline_energy", null_sum),
        total_vfd_saving=("saving", null_sum), running_hours=("running_hours", null_sum),
        auto_hours=("auto_hours", null_sum), manual_hours=("manual_hours", null_sum), bypass_hours=("bypass_hours", null_sum),
    )
    # Report-level context is aggregated from unique report intervals, never duplicated equipment rows.
    operational_keys = ["vessel", "vessel_category", "vessel_size", "report_date", "month"]
    daily_operational_rows: list[dict] = []
    for group_key, group in reports.groupby(operational_keys, dropna=False, sort=True):
        row = dict(zip(operational_keys, group_key))
        segment_types = sorted({str(value) for value in group["segment_type"].dropna() if str(value).strip()})
        row.update({
            "total_ae_energy": null_sum(group["total_ae_energy"]),
            "sea_temp": report_weighted_average(group, "sea_temp"),
            "pid_setpoint": report_weighted_average(group, "pid_setpoint"),
            "me_load": group["me_load"].mean() if group["me_load"].notna().any() else np.nan,
            "mcr_kw": group["mcr_kw"].dropna().iloc[0] if group["mcr_kw"].notna().any() else np.nan,
            "segment_type": " / ".join(segment_types) if segment_types else "Other / Unknown",
            "interval_count": int(group["interval_id"].nunique()),
        })
        daily_operational_rows.append(row)
    daily_operational = pd.DataFrame(daily_operational_rows)
    daily_operational["me_load_pct_mcr"] = np.where(
        daily_operational["mcr_kw"] > 0,
        daily_operational["me_load"] / daily_operational["mcr_kw"] * 100,
        np.nan,
    )
    daily_operational["cooling_temperature_margin"] = daily_operational["pid_setpoint"] - daily_operational["sea_temp"]
    daily_operational["sea_temp_pid_interaction"] = daily_operational["sea_temp"] * daily_operational["pid_setpoint"]
    daily_operational["cooling_margin_me_load_interaction"] = daily_operational["cooling_temperature_margin"] * daily_operational["me_load_pct_mcr"]
    daily_operational["sea_temp_me_load_interaction"] = daily_operational["sea_temp"] * daily_operational["me_load_pct_mcr"]
    daily_operational["pid_me_load_interaction"] = daily_operational["pid_setpoint"] * daily_operational["me_load_pct_mcr"]
    context_columns = ["vessel", "vessel_category", "vessel_size", "report_date", "month", "segment_type", "sea_temp", "me_load", "mcr_kw", "me_load_pct_mcr", "pid_setpoint", "cooling_temperature_margin", "sea_temp_pid_interaction", "cooling_margin_me_load_interaction", "sea_temp_me_load_interaction", "pid_me_load_interaction", "total_ae_energy", "interval_count"]
    daily_equipment = daily_equipment.drop(columns=["sea_temp", "me_load", "pid_setpoint", "interval_count"], errors="ignore").merge(
        daily_operational[context_columns], how="left",
        on=["vessel", "vessel_category", "vessel_size", "report_date", "month"],
    )
    daily_category = daily_category.drop(columns=["sea_temp", "me_load", "pid_setpoint", "interval_count"], errors="ignore").merge(
        daily_operational[context_columns], how="left",
        on=["vessel", "vessel_category", "vessel_size", "report_date", "month"],
    )
    daily_category["manual_pct"] = safe_ratio_series(daily_category["manual_hours"], daily_category["operating_hours"])
    daily_category["bypass_pct"] = safe_ratio_series(daily_category["bypass_hours"], daily_category["operating_hours"])
    daily_category["saving_pct"] = safe_ratio_series(daily_category["saving"], daily_category["baseline_energy"])
    daily_category["energy_per_running_hour"] = safe_ratio_series(daily_category["actual_energy"], daily_category["running_hours"], 1.0)
    daily_category["saving_per_running_hour"] = safe_ratio_series(daily_category["saving"], daily_category["running_hours"], 1.0)
    vessel_day = daily_operational.merge(
        daily_vfd, how="left", on=["vessel", "vessel_category", "vessel_size", "report_date", "month"]
    )

    box = interval[["vessel", "vessel_category", "vessel_size", "segment_type", "equipment_category", "equipment", "avg_load"]].dropna(subset=["avg_load"])
    if len(box) > 12000:
        box = box.iloc[np.linspace(0, len(box) - 1, 12000, dtype=int)]
    metadata = {
        "files_processed": read_ok,
        "interval_rows": len(interval),
        "daily_rows": len(daily_equipment),
        "date_min": interval["date_time"].min().strftime("%Y-%m-%d"),
        "date_max": interval["date_time"].max().strftime("%Y-%m-%d"),
        "report_interval_count": int(reports["interval_id"].nunique()),
        "report_ae_sum": null_sum(reports["total_ae_energy"]),
        "vessel_day_ae_sum": null_sum(vessel_day["total_ae_energy"]),
        "vessels_matched": len(matched_vessels),
        "unmatched_vessels": unmatched_vessels,
    }
    return interval, daily_equipment, daily_category, daily_operational, vessel_day, box, metadata


HTML = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>VFD Dissertation Dashboard</title><script>__PLOTLY__</script>
<style>
:root{--blue:#17365d;--mid:#2f75b5;--cyan:#5bc0de;--green:#49a078;--orange:#f28e2b;--red:#d9534f;--grey:#eef1f5;--ink:#263238}
*{box-sizing:border-box}body{margin:0;background:var(--grey);color:var(--ink);font:14px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}
header{background:linear-gradient(120deg,#102a43,#245b8f);color:#fff;padding:22px 4vw}h1{margin:0;font-size:clamp(22px,3vw,34px)}header p{margin:6px 0 0;opacity:.9}
.filters{position:sticky;top:0;z-index:20;background:#fff;padding:10px 3vw;display:grid;grid-template-columns:repeat(7,minmax(130px,1fr));gap:9px;box-shadow:0 2px 10px #0002}
label{font-size:11px;font-weight:700;color:var(--blue)}select,button{width:100%;padding:8px;border:1px solid #cbd5df;border-radius:6px;background:#fff;color:#17365d}.reset{align-self:end;background:var(--blue);color:#fff;cursor:pointer}
.nav{display:flex;gap:8px;padding:14px 4vw 4px;flex-wrap:wrap}.nav button{width:auto;padding:9px 15px;cursor:pointer}.nav button.active{background:var(--blue);color:#fff}
main{padding:10px 4vw 35px}.section{display:none}.section.active{display:block}.section h2{color:var(--blue);margin:8px 0}.note{color:#5c6773;margin:0 0 12px}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:12px 0}.card,.chart-card{background:#fff;border-radius:10px;box-shadow:0 2px 8px #17365d14}.kpi{padding:15px}.kpi b{display:block;color:#64748b;font-size:11px;text-transform:uppercase}.kpi span{display:block;color:var(--blue);font-size:clamp(18px,2vw,27px);font-weight:750;margin-top:5px}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.chart-card{padding:12px;min-height:390px;min-width:0;overflow:hidden;position:relative}.chart-card.wide{grid-column:1/-1}.chart-title{font-weight:700;color:var(--blue);padding:2px 8px}.plot{height:335px;width:100%;min-width:0;max-width:100%;position:relative;overflow:hidden}.plot .js-plotly-plot,.plot .plot-container,.plot .svg-container{max-width:100%!important}.plot .modebar-container{left:0!important;right:0!important;width:100%!important;max-width:100%!important;padding:3px 5px 0!important;box-sizing:border-box!important;pointer-events:none}.plot .modebar{display:flex!important;flex-wrap:wrap!important;justify-content:flex-end!important;max-width:100%!important;margin-left:auto!important;pointer-events:auto}.plot .modebar-group{display:flex!important;flex:0 0 auto!important}.plot .modebar-btn{padding:2px 3px!important}.plot .modebar-btn svg{width:17px!important;height:17px!important}.table-wrap{height:330px;overflow:auto;margin-top:7px}table{width:100%;border-collapse:collapse;font-size:12px}th{position:sticky;top:0;background:var(--blue);color:#fff;text-align:left}th,td{padding:7px 9px;border-bottom:1px solid #e4e8ed;white-space:nowrap}tr:nth-child(even){background:#f7f9fb}.hidden{display:none!important}
@media(max-width:1000px){.filters{grid-template-columns:repeat(3,1fr)}.kpis{grid-template-columns:repeat(3,1fr)}}@media(max-width:700px){.filters{position:static;grid-template-columns:1fr 1fr}.grid{grid-template-columns:1fr}.chart-card.wide{grid-column:auto}.kpis{grid-template-columns:1fr 1fr}.plot{height:310px}}
</style></head><body>
<header><h1>Monitored Equipment Energy &amp; Operational Performance</h1><p>Dissertation dashboard · 14-vessel consolidated analysis</p></header>
<div class="filters">
 <label>Vessel Category / Size<select id="fCat"></select></label><label>Vessel<select id="fVessel"></select></label>
 <label>Equipment Category<select id="fEqCat"></select></label><label>Equipment<select id="fEquip"></select></label>
 <label>Segment Type<select id="fSegment"></select></label><label>Group Results By<select id="fGroup"><option>Vessel</option><option>Vessel Category</option><option>Equipment Category</option><option>Equipment</option></select></label>
 <button class="reset" id="reset">Reset Filters</button>
</div>
<nav class="nav"><button data-section="overview" class="active">1 · Overview</button><button data-section="utilisation">2 · VFD Utilisation</button><button data-section="energy">3 · Energy Performance</button><button data-section="drivers">4 · Operational Drivers</button></nav>
<main>
<section id="overview" class="section active"><h2>Overview</h2><p class="note">Estimated energy saving relative to an equivalent fixed-speed operating baseline.</p><div id="kpis" class="kpis"></div><div class="grid">
<div class="chart-card"><div class="chart-title">Data coverage by vessel</div><div id="coverage" class="plot"></div></div><div class="chart-card"><div class="chart-title">Actual versus baseline energy by equipment category</div><div id="overviewEnergy" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Estimated saving percentage by vessel</div><div id="savingVessel" class="plot"></div></div><div class="chart-card"><div class="chart-title">Monthly actual, baseline and saving trend</div><div id="overviewMonthly" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Equipment availability</div><div id="availability" class="table-wrap"></div></div></div></section>
<section id="utilisation" class="section"><h2>VFD Utilisation</h2><div class="grid">
<div class="chart-card"><div class="chart-title">Auto / Manual / Bypass Operating Share</div><div id="modeShare" class="plot"></div></div><div class="chart-card"><div class="chart-title">Running-Hour-Weighted Average VFD Load</div><div id="weightedLoad" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Share of Equipment Running Hours in Intervals Above 80% and 90% Load</div><div id="highLoad" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Parallel-Unit Running-Hour Distribution by Vessel</div><div id="parallelHeatmap" class="plot"></div><p class="note">Running-hour shares are calculated separately within each vessel and equipment category. Unavailable units are excluded rather than treated as zero-hour equipment.</p></div>
<div class="chart-card"><div class="chart-title">Distribution of Report-Interval Average Load</div><div id="loadBox" class="plot"></div><p class="note">Each observation represents the average VFD load recorded for one report interval.</p></div><div class="chart-card"><div class="chart-title">Daily Running-Hour-Weighted Average Load</div><div id="dailyLoad" class="plot"></div><p class="note">The number and composition of active vessels vary across the monitoring period. The rolling average is provided to clarify the underlying utilisation trend.</p></div>
<div class="chart-card"><div class="chart-title">Port Stay versus Sea Passage Load</div><div id="segmentLoad" class="plot"></div></div><div class="chart-card"><div class="chart-title">Utilisation Summary</div><div id="utilTable" class="table-wrap"></div></div></div></section>
<section id="energy" class="section"><h2>Energy Performance</h2><p class="note">Estimated energy saving relative to an equivalent fixed-speed operating baseline.</p><div class="grid">
<div class="chart-card"><div class="chart-title">Actual energy versus fixed-speed baseline</div><div id="energyActualBase" class="plot"></div></div><div class="chart-card"><div class="chart-title">Estimated saving percentage</div><div id="energySavingPct" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Energy per VFD running hour</div><div id="energyPerHour" class="plot"></div></div><div class="chart-card"><div class="chart-title">Saving per VFD running hour</div><div id="savingPerHour" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Cumulative estimated savings</div><div id="cumulative" class="plot"></div></div><div class="chart-card"><div class="chart-title">Monthly energy trend</div><div id="energyMonthly" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">SW Pumps versus Engine-Room Fans</div><div id="categoryEnergy" class="plot"></div></div><div class="chart-card"><div class="chart-title">Port Stay versus Sea Passage saving percentage</div><div id="segmentSaving" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Port Stay versus Sea Passage saving per running hour</div><div id="segmentSavingHour" class="plot"></div></div><div class="chart-card"><div class="chart-title">Total Auxiliary Energy vs Energy of Monitored VFD-Controlled Equipment</div><div id="aeMonitored" class="plot"></div><p id="contributionNote" class="note"></p></div>
<div class="chart-card"><div class="chart-title">Monitored VFD-Controlled Equipment Share of Total Auxiliary Energy</div><div id="contribution" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Energy summary</div><div id="energyTable" class="table-wrap"></div></div></div></section>
<section id="drivers" class="section"><h2>Operational Drivers</h2><p class="note">The multivariable charts visualise the simultaneous variation of SW-pump load with sea-water temperature, PID temperature setting, main-engine load and operational segment. These visual relationships are exploratory and do not establish causal effects.<br>Cooling Temperature Margin is calculated as PID Temperature Setpoint minus Sea-Water Temperature. It represents a proxy for the available external cooling-temperature difference and not the actual PID control error, because the measured controlled-water temperature is not available in the present dataset.</p><div class="grid">
<div class="chart-card"><div class="chart-title">SW-Pump Load by Sea-Water Temperature and Operational Segment</div><div id="seaTempLoad" class="plot"></div></div><div class="chart-card"><div class="chart-title">Main-Engine Load vs Daily SW-Pump Load</div><div id="meLoadSW" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Multivariable SW-Pump Operational Relationship</div><div id="multivariableScatter" class="plot"></div><p class="note">Marker colour represents PID Temperature Setpoint, marker size represents Main-Engine Load % MCR, and marker symbol represents operational segment.</p></div><div class="chart-card"><div class="chart-title">Sea-Water Temperature and PID Setpoint Interaction</div><div id="pidInteractionHeatmap" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Daily SW-Pump Load versus Cooling-Temperature Margin</div><div id="coolingMarginLoad" class="plot"></div></div>
<div class="chart-card wide"><div class="chart-title">Spearman Correlation Matrix</div><div id="corrMatrix" class="plot"></div></div>
<div class="chart-card"><div class="chart-title">Operational Driver Correlation Summary</div><div id="driverCorrTable" class="table-wrap"></div></div><div class="chart-card"><div class="chart-title">Correlation Summary Table</div><div id="corrTable" class="table-wrap"></div></div></div></section>
</main><script>
const DATA=__DATA__; const C={actual:'#2f75b5',base:'#9aa6b2',save:'#49a078',orange:'#f28e2b',red:'#d9534f',cyan:'#5bc0de',blue:'#17365d'};
function num(x){if(x===null||x===undefined||x===""){return null;}const value=Number(x);return Number.isFinite(value)?value:null;}
function between(x,low,high){const value=num(x);return value!==null&&value>=low&&value<=high;}
const sum=(rows,key)=>{const a=rows.map(r=>num(r[key])).filter(v=>v!==null);return a.length?a.reduce((x,y)=>x+y,0):null};
const ratio=(a,b,m=100)=>a!==null&&b!==null&&b!==0?a/b*m:null; const fmt=(x,d=1)=>x===null||!Number.isFinite(x)?'—':x.toLocaleString('en-GB',{maximumFractionDigits:d});
function uniq(a){return [...new Set(a.filter(x=>x!==null&&x!==undefined&&x!==''))].sort((a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true}));}
function fillSelect(id,vals){const e=document.getElementById(id);e.innerHTML='<option value="All">All</option>'+uniq(vals).map(v=>`<option>${v}</option>`).join('');}
fillSelect('fCat',DATA.interval.map(r=>r.vessel_size));fillSelect('fVessel',DATA.interval.map(r=>r.vessel));fillSelect('fEqCat',DATA.interval.map(r=>r.equipment_category));fillSelect('fEquip',DATA.interval.map(r=>r.equipment));fillSelect('fSegment',DATA.interval.map(r=>r.segment_type));
const filterIds=['fCat','fVessel','fEqCat','fEquip','fSegment'];
let groupSelectionTouched=false;
function selected(){return {vessel_size:fCat.value,vessel:fVessel.value,equipment_category:fEqCat.value,equipment:fEquip.value,segment_type:fSegment.value}}
function filtered(rows,ignoreSegment=false){const s=selected();return rows.filter(r=>Object.entries(s).every(([k,v])=>v==='All'||(ignoreSegment&&k==='segment_type')||r[k]===v));}
function groupKey(){return {'Vessel':'vessel','Vessel Category':'vessel_category','Equipment Category':'equipment_category','Equipment':'equipment'}[fGroup.value]}
function aggregate(rows,key=groupKey()){const m=new Map();rows.forEach(r=>{const k=r[key]??'Not specified';if(!m.has(k))m.set(k,[]);m.get(k).push(r)});return [...m].map(([name,x])=>({name,rows:x,run:sum(x,'running_hours'),auto:sum(x,'auto_hours'),manual:sum(x,'manual_hours'),bypass:sum(x,'bypass_hours'),op:sum(x,'operating_hours'),actual:sum(x,'actual_energy'),base:sum(x,'baseline_energy'),saving:sum(x,'saving'),load:weighted(x)}));}
function weighted(rows){let n=0,d=0;rows.forEach(r=>{const v=num(r.avg_load),w=num(r.running_hours);if(v!==null&&w!==null&&w>0){n+=v*w;d+=w}});return d?n/d:null}
const layout=(y='',extra={})=>Object.assign({margin:{l:60,r:20,t:18,b:70},paper_bgcolor:'#fff',plot_bgcolor:'#fff',font:{family:'system-ui',color:'#263238'},legend:{orientation:'h',x:0,y:-.22},yaxis:{title:y,gridcolor:'#e9edf2',zeroline:false},xaxis:{gridcolor:'#eef1f5'},hovermode:'closest'},extra);
const AXIS_TITLES={
 coverage:{x:'Vessel',y:'Count'},overviewEnergy:{x:'Equipment Category',y:'Energy (kWh)'},savingVessel:{x:'Vessel',y:'Estimated Saving (%)'},overviewMonthly:{x:'Calendar Month',y:'Energy (kWh)'},
 modeShare:{x:'Selected Group',y:'Operating Time (%)'},weightedLoad:{x:'Selected Group',y:'Running-Hour-Weighted Average Load (%)'},highLoad:{x:'Selected Group',y:'Equipment Running Hours Share (%)'},parallelHeatmap:{x:'Equipment',y:'Vessel'},loadBox:{x:'Equipment Unit / Selected Group',y:'Report-Interval Average Load (%)'},dailyLoad:{x:'Calendar Date',y:'Running-Hour-Weighted Average Load (%)'},segmentLoad:{x:'Segment Type',y:'Running-Hour-Weighted Average Load (%)'},
 energyActualBase:{x:'Selected Group',y:'Energy (kWh)'},energySavingPct:{x:'Selected Group',y:'Estimated Saving (%)'},energyPerHour:{x:'Selected Group',y:'Energy per Running Hour (kWh/h)'},savingPerHour:{x:'Selected Group',y:'Saving per Running Hour (kWh/h)'},cumulative:{x:'Calendar Date',y:'Cumulative Estimated Saving (kWh)'},energyMonthly:{x:'Calendar Month',y:'Energy (kWh)'},categoryEnergy:{x:'Equipment Category',y:'Energy (kWh)'},segmentSaving:{x:'Segment Type',y:'Estimated Saving (%)'},segmentSavingHour:{x:'Segment Type',y:'Saving per Running Hour (kWh/h)'},aeMonitored:{x:'Total Auxiliary Energy (kWh)',y:'Monitored VFD-Controlled Equipment Energy (kWh)'},contribution:{x:'Vessel',y:'Share of Total Auxiliary Energy (%)'},
 seaTempLoad:{x:'Sea-Water Temperature Bin (°C)',y:'Running-Hour-Weighted Daily SW-Pump Load (%)'},meLoadSW:{x:'Main-Engine Load (% MCR)',y:'Daily Weighted SW-Pump Load (%)'},multivariableScatter:{x:'Sea-Water Temperature (°C)',y:'Daily Weighted SW-Pump Load (%)'},pidInteractionHeatmap:{x:'Sea-Water Temperature Bin (°C)',y:'PID Temperature Setpoint Bin (°C)'},coolingMarginLoad:{x:'Cooling Temperature Margin (°C)',y:'Daily Weighted SW-Pump Load (%)'},corrMatrix:{x:'Variable',y:'Variable'}
};
function axisTitleText(title,fallback){if(typeof title==='string'&&title.trim())return title;if(title&&typeof title.text==='string'&&title.text.trim())return title.text;return fallback}
const cfg={responsive:true,displayModeBar:true,displaylogo:false,modeBarButtonsToRemove:['lasso2d','select2d']}; function plot(id,traces,lay={}){const finalLayout=Object.assign({},lay),titles=AXIS_TITLES[id];if(titles){finalLayout.xaxis=Object.assign({},finalLayout.xaxis||{});finalLayout.yaxis=Object.assign({},finalLayout.yaxis||{});finalLayout.xaxis.title={text:axisTitleText(finalLayout.xaxis.title,titles.x),standoff:14,font:{size:12,color:C.blue}};finalLayout.yaxis.title={text:axisTitleText(finalLayout.yaxis.title,titles.y),standoff:14,font:{size:12,color:C.blue}};finalLayout.xaxis.automargin=true;finalLayout.yaxis.automargin=true;finalLayout.margin=Object.assign({},finalLayout.margin||{});finalLayout.margin.l=Math.max(Number(finalLayout.margin.l)||0,85);finalLayout.margin.b=Math.max(Number(finalLayout.margin.b)||0,95)}Plotly.react(id,traces,finalLayout,cfg)}
function bar(id,names,series,y='',mode='group'){plot(id,series.map((s,i)=>({type:'bar',name:s.name,x:names,y:s.values,marker:{color:s.color}})),layout(y,{barmode:mode}))}
function line(id,x,series,y=''){plot(id,series.map(s=>({type:'scatter',mode:'lines+markers',name:s.name,x,y:s.values,line:{color:s.color}})),layout(y))}
function table(id,headers,rows){document.getElementById(id).innerHTML=`<table><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(v=>`<td>${v}</td>`).join('')}</tr>`).join('')}</tbody></table>`}
function by(rows,key){const out={};rows.forEach(r=>(out[r[key]]??=[]).push(r));return out}
function monthly(rows){return Object.entries(by(rows,'month')).sort().map(([name,x])=>({name,actual:sum(x,'actual_energy'),base:sum(x,'baseline_energy'),saving:sum(x,'saving')}))}
function renderOverview(){const r=filtered(DATA.interval),a=aggregate(r);const vessels=uniq(r.map(x=>x.vessel)).length,units=uniq(r.map(x=>x.vessel+'|'+x.equipment)).length,run=sum(r,'running_hours'),actual=sum(r,'actual_energy'),saving=sum(r,'saving'),base=sum(r,'baseline_energy');
 const cards=[['Selected Vessels',fmt(vessels,0)],['Available Equipment Units',fmt(units,0)],['VFD Running Hours',fmt(run)],['Actual Energy of Monitored VFD-Controlled Equipment',fmt(actual)+' kWh'],['Estimated Saving',fmt(saving)+' kWh'],['Estimated Saving %',fmt(ratio(saving,base))+'%']];kpis.innerHTML=cards.map(x=>`<div class="card kpi"><b>${x[0]}</b><span>${x[1]}</span></div>`).join('');
 const cov=Object.entries(by(r,'vessel')).map(([name,x])=>({name,intervals:uniq(x.map(z=>z.interval_id)).length,days:uniq(x.map(z=>z.report_date)).length}));bar('coverage',cov.map(x=>x.name),[{name:'Report intervals',values:cov.map(x=>x.intervals),color:C.blue},{name:'Calendar days',values:cov.map(x=>x.days),color:C.cyan}],'Count');
 const cats=aggregate(r,'equipment_category');bar('overviewEnergy',cats.map(x=>x.name),[{name:'Actual energy',values:cats.map(x=>x.actual),color:C.actual},{name:'Fixed-speed baseline',values:cats.map(x=>x.base),color:C.base}],'Energy (kWh)');
 const vs=aggregate(r,'vessel');bar('savingVessel',vs.map(x=>x.name),[{name:'Estimated saving %',values:vs.map(x=>ratio(x.saving,x.base)),color:C.save}],'Estimated saving (%)');
 const mo=monthly(r);line('overviewMonthly',mo.map(x=>x.name),[{name:'Actual',values:mo.map(x=>x.actual),color:C.actual},{name:'Baseline',values:mo.map(x=>x.base),color:C.base},{name:'Saving',values:mo.map(x=>x.saving),color:C.save}],'Energy (kWh)');
 const av=aggregate(r,'equipment').sort((x,y)=>x.name.localeCompare(y.name));table('availability',['Equipment','Vessel-equipment units','Intervals','Running hours','Weighted load','Actual kWh','Baseline kWh'],av.map(x=>[x.name,fmt(uniq(x.rows.map(z=>z.vessel+'|'+z.equipment)).length,0),fmt(uniq(x.rows.map(z=>z.interval_id)).length,0),fmt(x.run),fmt(x.load)+'%',fmt(x.actual),fmt(x.base)]));}
function pctText(value){return value===null||!Number.isFinite(value)?'—':fmt(value,1)+'%'}
function meaningfulEquipmentRow(r){return num(r.running_hours)>0||num(r.avg_load)!==null||num(r.actual_energy)!==null||num(r.baseline_energy)!==null}
function availableEquipmentUnits(rows){const units=new Map();rows.forEach(r=>{const key=r.vessel+'|'+r.equipment;if(!units.has(key))units.set(key,[]);units.get(key).push(r)});return [...units.values()].filter(x=>x.some(meaningfulEquipmentRow)).length}
function thresholdShares(rows){const valid=rows.filter(r=>num(r.running_hours)>0&&num(r.avg_load)!==null),den=sum(valid,'running_hours');return {p80:ratio(sum(valid.filter(r=>num(r.avg_load)>80),'running_hours'),den),p90:ratio(sum(valid.filter(r=>num(r.avg_load)>90),'running_hours'),den)}}
function labelledBar(id,names,series,y='',mode='group'){plot(id,series.map(s=>({type:'bar',name:s.name,x:names,y:s.values,text:s.values.map(pctText),textposition:'outside',cliponaxis:false,marker:{color:s.color}})),layout(y,{barmode:mode,margin:{l:65,r:20,t:35,b:75}}))}
function parallelFilteredRows(){const s=selected();return DATA.interval.filter(r=>Object.entries(s).every(([k,v])=>k==='equipment'||v==='All'||r[k]===v))}
function renderParallelHeatmap(){const base=parallelFilteredRows(),swUnits=['SW1','SW2','SW3'],fanUnits=['FAN1','FAN2','FAN3','FAN4'];let columns=fEqCat.value==='SW Pumps'?swUnits:fEqCat.value==='Engine-Room Fans'?fanUnits:[...swUnits,...fanUnits];if(fEquip.value!=='All')columns=columns.filter(e=>e===fEquip.value);const vessels=uniq(base.filter(meaningfulEquipmentRow).map(r=>r.vessel)),annotations=[],z=[],custom=[];vessels.forEach(v=>{const vr=base.filter(r=>r.vessel===v),size=vr[0]?.vessel_size??'Not specified',groupStats={};['SW Pumps','Engine-Room Fans'].forEach(category=>{const candidates=category==='SW Pumps'?swUnits:fanUnits,units=candidates.map(e=>{const rows=vr.filter(r=>r.equipment===e),available=rows.some(meaningfulEquipmentRow),run=available?sum(rows,'running_hours'):null;return {equipment:e,available,run}}).filter(u=>u.available),runUnits=units.filter(u=>u.run!==null),completeRuns=units.length>0&&runUnits.length===units.length,total=completeRuns?runUnits.reduce((a,u)=>a+u.run,0):null,mean=completeRuns?total/runUnits.length:null,ordered=[...runUnits].sort((a,b)=>b.run-a.run),imbalance=completeRuns&&runUnits.length>=2&&mean>0?(ordered[0].run-ordered[ordered.length-1].run)/mean:null;groupStats[category]={units,total,count:units.length,equal:units.length?100/units.length:null,imbalance,most:completeRuns&&runUnits.length>=2?ordered[0].equipment:'—',least:completeRuns&&runUnits.length>=2?ordered[ordered.length-1].equipment:'—'}});const zr=[],cr=[];columns.forEach(e=>{const category=e.startsWith('SW')?'SW Pumps':'Engine-Room Fans',g=groupStats[category],u=g.units.find(x=>x.equipment===e);if(!u){zr.push(null);cr.push([v,size,e,category,null,g.total,null,g.count,g.equal,null,g.imbalance,g.most,g.least]);annotations.push({text:'N/A',x:e,y:v,showarrow:false,font:{size:10,color:'#4b5563'}});return}const share=u.run!==null&&g.total!==null&&g.total>0?u.run/g.total*100:null,deviation=share!==null&&g.equal!==null?share-g.equal:null;if(share===null)annotations.push({text:'N/A',x:e,y:v,showarrow:false,font:{size:10,color:'#4b5563'}});zr.push(share);cr.push([v,size,e,category,u.run,g.total,share,g.count,g.equal,deviation,g.imbalance,g.most,g.least])});z.push(zr);custom.push(cr)});plot('parallelHeatmap',[{type:'heatmap',x:columns,y:vessels,z,customdata:custom,zmin:0,zmax:100,colorscale:[[0,'#f7fbff'],[.35,'#6baed6'],[.7,'#fdae61'],[1,'#b2182b']],colorbar:{title:{text:'Running-Hour<br>Share (%)'}},hoverongaps:false,hovertemplate:'Vessel: %{customdata[0]}<br>Vessel Size: %{customdata[1]}<br>Equipment: %{customdata[2]}<br>Equipment Category: %{customdata[3]}<br>Equipment Running Hours: %{customdata[4]:,.2f}<br>Parallel-Group Total Running Hours: %{customdata[5]:,.2f}<br>Running-Hour Share: %{customdata[6]:.2f}%<br>Number of Available Parallel Units: %{customdata[7]}<br>Equal-Use Reference Share: %{customdata[8]:.2f}%<br>Deviation from Equal-Use Reference: %{customdata[9]:.2f} percentage points<br>Normalised Group Imbalance: %{customdata[10]:.3f}<br>Most-Used Unit: %{customdata[11]}<br>Least-Used Unit: %{customdata[12]}<extra></extra>'}],layout('',{xaxis:{title:'Equipment'},yaxis:{title:'Vessel',autorange:'reversed'},annotations,plot_bgcolor:'#c7cdd4',margin:{l:95,r:105,t:20,b:70}}))}
function renderDailyUtilisationTrend(rows){const valid=rows.filter(r=>num(r.avg_load)!==null&&num(r.running_hours)>0),days=Object.entries(by(valid,'report_date')).sort().map(([date,x])=>({date,load:weighted(x),hours:sum(x,'running_hours'),vessels:uniq(x.map(r=>r.vessel)).length,units:uniq(x.map(r=>r.vessel+'|'+r.equipment)).length}));days.forEach((d,i)=>{const cutoff=Date.parse(d.date)-29*86400000,window=days.slice(0,i+1).filter(x=>Date.parse(x.date)>=cutoff).map(x=>x.load).filter(x=>x!==null);d.rolling=window.length>=7?window.reduce((a,b)=>a+b,0)/window.length:null});const custom=days.map(d=>[d.load,d.rolling,d.hours,d.vessels,d.units,fSegment.value]),hover='Date: %{x}<br>Daily Weighted Average Load: %{customdata[0]:.2f}%<br>30-Day Rolling Average: %{customdata[1]:.2f}%<br>Equipment Running Hours: %{customdata[2]:,.2f}<br>Active Vessels: %{customdata[3]}<br>Available Equipment Units: %{customdata[4]}<br>Segment Type selection: %{customdata[5]}<extra></extra>';plot('dailyLoad',[{type:'scatter',mode:'lines+markers',name:'Daily Weighted Average Load',x:days.map(d=>d.date),y:days.map(d=>d.load),customdata:custom,line:{color:'#8fc4e8',width:1.2},marker:{size:3},hovertemplate:hover},{type:'scatter',mode:'lines',name:'30-Day Rolling Average',x:days.map(d=>d.date),y:days.map(d=>d.rolling),customdata:custom,line:{color:C.blue,width:3},hovertemplate:hover}],layout('Load (%)',{xaxis:{title:'Calendar Date',gridcolor:'#eef1f5'},margin:{l:65,r:20,t:20,b:75}}))}
function renderLoadBox(){const rows=filtered(DATA.box).filter(r=>num(r.avg_load)!==null),key=groupSelectionTouched?groupKey():'equipment',defaultOrder=['FAN1','FAN2','FAN3','FAN4','SW1','SW2','SW3'],names=key==='equipment'?defaultOrder.filter(name=>rows.some(r=>r.equipment===name)):uniq(rows.map(r=>r[key]));plot('loadBox',names.map(name=>({type:'box',name,y:rows.filter(r=>(r[key]??'Not specified')===name).map(r=>r.avg_load),boxpoints:false,marker:{color:C.blue}})),layout('Report-Interval Average Load (%)'))}
function renderSegmentLoad(rows){const order=['Port Stay','Sea Passage','Other / Unknown'],segments=order.filter(segment=>rows.some(r=>r.segment_type===segment)),stats=segments.map(segment=>{const x=rows.filter(r=>r.segment_type===segment&&num(r.avg_load)!==null&&num(r.running_hours)>0);return {segment,load:weighted(x),hours:sum(x,'running_hours'),vessels:uniq(x.map(r=>r.vessel)).length,units:uniq(x.map(r=>r.vessel+'|'+r.equipment)).length}});plot('segmentLoad',[{type:'bar',name:'Weighted Average Load',x:stats.map(x=>x.segment),y:stats.map(x=>x.load),text:stats.map(x=>pctText(x.load)),textposition:'outside',cliponaxis:false,customdata:stats.map(x=>[x.hours,x.vessels,x.units]),marker:{color:C.cyan},hovertemplate:'Segment Type: %{x}<br>Weighted Average Load: %{y:.2f}%<br>Equipment Running Hours: %{customdata[0]:,.2f}<br>Selected Vessels: %{customdata[1]}<br>Available Equipment Units: %{customdata[2]}<extra></extra>'}],layout('Load (%)',{margin:{l:65,r:20,t:35,b:75}}))}
function renderUtilisation(){const r=filtered(DATA.interval),a=aggregate(r),names=a.map(x=>x.name),modeSeries=[{name:'Auto',values:a.map(x=>ratio(x.auto,x.op)),color:C.actual},{name:'Manual',values:a.map(x=>ratio(x.manual,x.op)),color:C.orange},{name:'Bypass',values:a.map(x=>ratio(x.bypass,x.op)),color:C.red}];bar('modeShare',names,modeSeries,'Operating time (%)','stack');labelledBar('weightedLoad',names,[{name:'Weighted Average Load',values:a.map(x=>x.load),color:C.blue}],'Load (%)');const hl=a.map(g=>thresholdShares(g.rows));labelledBar('highLoad',names,[{name:'Above 80%',values:hl.map(x=>x.p80),color:C.orange},{name:'Above 90%',values:hl.map(x=>x.p90),color:C.red}],'Equipment Running Hours share (%)');renderParallelHeatmap();renderLoadBox();renderDailyUtilisationTrend(r);renderSegmentLoad(r);table('utilTable',['Group','Equipment Running Hours','Auto %','Manual %','Bypass %','Weighted Load %','Available Equipment Units','Share Above 80%','Share Above 90%'],a.map((x,i)=>[x.name,fmt(x.run),pctText(ratio(x.auto,x.op)),pctText(ratio(x.manual,x.op)),pctText(ratio(x.bypass,x.op)),pctText(x.load),fmt(availableEquipmentUnits(x.rows),0),pctText(hl[i].p80),pctText(hl[i].p90)]));}
function selectedVesselDays(rows){const groups=by(rows.map(r=>Object.assign({},r,{day_key:r.vessel+'|'+r.report_date})),'day_key');return Object.values(groups).map(x=>{const intervals=new Map();x.forEach(r=>{if(!intervals.has(r.interval_id))intervals.set(r.interval_id,num(r.total_ae_energy))});const ae=[...intervals.values()].filter(v=>v!==null);return {vessel:x[0].vessel,vessel_size:x[0].vessel_size,report_date:x[0].report_date,total_ae_energy:ae.length?ae.reduce((a,b)=>a+b,0):null,monitored_energy:sum(x,'actual_energy')}})}
function renderEnergy(){const r=filtered(DATA.interval),a=aggregate(r),names=a.map(x=>x.name);bar('energyActualBase',names,[{name:'Actual',values:a.map(x=>x.actual),color:C.actual},{name:'Baseline',values:a.map(x=>x.base),color:C.base}],'Energy (kWh)');bar('energySavingPct',names,[{name:'Estimated saving %',values:a.map(x=>ratio(x.saving,x.base)),color:C.save}],'Estimated saving (%)');bar('energyPerHour',names,[{name:'Energy / running h',values:a.map(x=>ratio(x.actual,x.run,1)),color:C.actual}],'kWh / running hour');bar('savingPerHour',names,[{name:'Saving / running h',values:a.map(x=>ratio(x.saving,x.run,1)),color:C.save}],'kWh / running hour');
 const dates=Object.entries(by(r,'report_date')).sort();let cum=0;line('cumulative',dates.map(x=>x[0]),[{name:'Cumulative saving',values:dates.map(x=>{const v=sum(x[1],'saving');if(v!==null)cum+=v;return cum}),color:C.save}],'Estimated saving (kWh)');const mo=monthly(r);line('energyMonthly',mo.map(x=>x.name),[{name:'Actual',values:mo.map(x=>x.actual),color:C.actual},{name:'Baseline',values:mo.map(x=>x.base),color:C.base},{name:'Saving',values:mo.map(x=>x.saving),color:C.save}],'Energy (kWh)');
 const cat=aggregate(r,'equipment_category');bar('categoryEnergy',cat.map(x=>x.name),[{name:'Actual',values:cat.map(x=>x.actual),color:C.actual},{name:'Saving',values:cat.map(x=>x.saving),color:C.save}],'Energy (kWh)');const seg=aggregate(r,'segment_type');bar('segmentSaving',seg.map(x=>x.name),[{name:'Estimated saving %',values:seg.map(x=>ratio(x.saving,x.base)),color:C.save}],'Estimated saving (%)');bar('segmentSavingHour',seg.map(x=>x.name),[{name:'Saving / running h',values:seg.map(x=>ratio(x.saving,x.run,1)),color:C.orange}],'kWh / running hour');
 const days=selectedVesselDays(r),valid=days.filter(x=>num(x.total_ae_energy)>0&&num(x.monitored_energy)!==null&&num(x.monitored_energy)>=0&&num(x.monitored_energy)<=num(x.total_ae_energy)),invalid=days.filter(x=>num(x.total_ae_energy)!==null&&num(x.monitored_energy)!==null&&num(x.monitored_energy)>num(x.total_ae_energy));
 plot('aeMonitored',uniq(valid.map(x=>x.vessel)).map(v=>{const z=valid.filter(x=>x.vessel===v);return {type:'scattergl',mode:'markers',name:v,x:z.map(x=>x.total_ae_energy),y:z.map(x=>x.monitored_energy),customdata:z.map(x=>[x.vessel_size,x.report_date]),hovertemplate:'Vessel: '+v+'<br>Vessel Size: %{customdata[0]}<br>Date: %{customdata[1]}<br>Total Auxiliary Energy: %{x:.1f} kWh<br>Actual Energy of Monitored VFD-Controlled Equipment: %{y:.1f} kWh<extra></extra>'}}),layout('Actual Energy of Monitored VFD-Controlled Equipment (kWh)',{xaxis:{title:'Total Auxiliary Energy (kWh)',gridcolor:'#eef1f5'}}));contributionNote.textContent=`${invalid.length.toLocaleString()} invalid contribution observation(s) above 100% excluded from this scatter plot and contribution calculation only.`;
 const vc=Object.entries(by(valid,'vessel')).map(([name,x])=>({name,n:sum(x,'monitored_energy'),d:sum(x,'total_ae_energy')}));bar('contribution',vc.map(x=>x.name),[{name:'Monitored equipment share',values:vc.map(x=>ratio(x.n,x.d)),color:C.cyan}],'Share of total auxiliary energy (%)');table('energyTable',['Group','Actual kWh','Baseline kWh','Saving kWh','Saving %','Energy/run h','Saving/run h'],a.map(x=>[x.name,fmt(x.actual),fmt(x.base),fmt(x.saving),fmt(ratio(x.saving,x.base))+'%',fmt(ratio(x.actual,x.run,1)),fmt(ratio(x.saving,x.run,1))]));}
function average(rows,key){const x=rows.map(r=>num(r[key])).filter(v=>v!==null);return x.length?x.reduce((a,b)=>a+b,0)/x.length:null}
function reportWeightedAverage(rows,key){const weightedRows=rows.map(r=>[num(r[key]),num(r.report_hours)]).filter(x=>x[0]!==null&&x[1]!==null&&x[1]>0);if(weightedRows.length){const d=weightedRows.reduce((a,x)=>a+x[1],0);return weightedRows.reduce((a,x)=>a+x[0]*x[1],0)/d}return average(rows,key)}
function driverDaily(){const groups={};filtered(DATA.interval).forEach(r=>{const k=r.vessel+'|'+r.report_date+'|'+r.equipment_category;(groups[k]??=[]).push(r)});return Object.values(groups).map(x=>{const reports=[...new Map(x.map(r=>[r.interval_id,r])).values()],run=sum(x,'running_hours'),op=sum(x,'operating_hours'),actual=sum(x,'actual_energy'),base=sum(x,'baseline_energy'),saving=sum(x,'saving'),mcr=average(reports,'mcr_kw'),me=average(reports,'me_load'),mePct=mcr!==null&&mcr>0&&me!==null?me/mcr*100:null,sea=reportWeightedAverage(reports,'sea_temp'),pid=reportWeightedAverage(reports,'pid_setpoint'),margin=sea!==null&&pid!==null?pid-sea:null,interaction=sea!==null&&pid!==null?sea*pid:null,segment=uniq(reports.map(r=>r.segment_type)).join(' / ')||'Other / Unknown',energyHour=ratio(actual,run,1),savingHour=ratio(saving,run,1);return {vessel:x[0].vessel,vessel_category:x[0].vessel_category,vessel_size:x[0].vessel_size,report_date:x[0].report_date,segment_type:segment,equipment_category:x[0].equipment_category,sea_temp:sea,me_load:me,mcr_kw:mcr,me_load_pct_mcr:mePct,pid_setpoint:pid,cooling_temperature_margin:margin,sea_temp_pid_interaction:interaction,cooling_margin_me_load_interaction:margin!==null&&mePct!==null?margin*mePct:null,sea_temp_me_load_interaction:sea!==null&&mePct!==null?sea*mePct:null,pid_me_load_interaction:pid!==null&&mePct!==null?pid*mePct:null,avg_load:weighted(x),running_hours:run,auto_hours:sum(x,'auto_hours'),manual_hours:sum(x,'manual_hours'),bypass_hours:sum(x,'bypass_hours'),operating_hours:op,actual_energy:actual,baseline_energy:base,saving:saving,manual_pct:ratio(sum(x,'manual_hours'),op),bypass_pct:ratio(sum(x,'bypass_hours'),op),saving_pct:ratio(saving,base),energy_per_running_hour:energyHour,saving_per_running_hour:savingHour,energy_per_hour:energyHour,saving_per_hour:savingHour}})}
function multivariableDaily(){const groups={};filtered(DATA.interval).forEach(r=>{if(r.equipment_category!=='SW Pumps')return;const k=r.vessel+'|'+r.report_date+'|'+r.segment_type;(groups[k]??=[]).push(r)});return Object.values(groups).map(x=>{const reports=[...new Map(x.map(r=>[r.interval_id,r])).values()],run=sum(x,'running_hours'),op=sum(x,'operating_hours'),actual=sum(x,'actual_energy'),base=sum(x,'baseline_energy'),saving=sum(x,'saving'),mcr=average(reports,'mcr_kw'),me=reportWeightedAverage(reports,'me_load'),mePct=mcr!==null&&mcr>0&&me!==null?me/mcr*100:null,sea=reportWeightedAverage(reports,'sea_temp'),pid=reportWeightedAverage(reports,'pid_setpoint'),margin=sea!==null&&pid!==null?pid-sea:null,energyHour=ratio(actual,run,1),savingHour=ratio(saving,run,1);return {vessel:x[0].vessel,vessel_category:x[0].vessel_category,vessel_size:x[0].vessel_size,report_date:x[0].report_date,segment_type:x[0].segment_type,equipment_category:x[0].equipment_category,sea_temp:sea,pid_setpoint:pid,cooling_temperature_margin:margin,sea_temp_pid_interaction:sea!==null&&pid!==null?sea*pid:null,me_load_pct_mcr:mePct,cooling_margin_me_load_interaction:margin!==null&&mePct!==null?margin*mePct:null,sea_temp_me_load_interaction:sea!==null&&mePct!==null?sea*mePct:null,pid_me_load_interaction:pid!==null&&mePct!==null?pid*mePct:null,avg_load:weighted(x),running_hours:run,manual_pct:ratio(sum(x,'manual_hours'),op),bypass_pct:ratio(sum(x,'bypass_hours'),op),saving_pct:ratio(saving,base),actual_energy:actual,baseline_energy:base,energy_per_running_hour:energyHour,saving_per_running_hour:savingHour}})}
function scatter(id,rows,xk,yk,xlab,ylab,color=C.actual){const valid=rows.filter(r=>num(r[xk])!==null&&num(r[yk])!==null);plot(id,[{type:'scattergl',mode:'markers',x:valid.map(r=>r[xk]),y:valid.map(r=>r[yk]),text:valid.map(r=>r.vessel+' · '+r.report_date),marker:{color,opacity:.6,size:7}}],layout(ylab,{xaxis:{title:xlab,gridcolor:'#eef1f5'}}));return valid}
function rank(a){const s=a.map((v,i)=>[v,i]).sort((x,y)=>x[0]-y[0]),r=Array(a.length);for(let i=0;i<s.length;){let j=i;while(j+1<s.length&&s[j+1][0]===s[i][0])j++;const q=(i+j+2)/2;for(let k=i;k<=j;k++)r[s[k][1]]=q;i=j+1}return r}
function pearson(a,b){const n=a.length;if(n<3)return null;const am=a.reduce((x,y)=>x+y)/n,bm=b.reduce((x,y)=>x+y)/n;let p=0,da=0,db=0;for(let i=0;i<n;i++){const x=a[i]-am,y=b[i]-bm;p+=x*y;da+=x*x;db+=y*y}return da&&db?p/Math.sqrt(da*db):null}
function spearman(rows,a,b,minN=20){const p=rows.map(r=>[num(r[a]),num(r[b])]).filter(x=>x[0]!==null&&x[1]!==null);return {rho:p.length>=minN?pearson(rank(p.map(x=>x[0])),rank(p.map(x=>x[1]))):null,n:p.length}}
function lowessTrend(rows,xk,yk,frac=.25,pointCount=100){const pairs=rows.map(r=>[num(r[xk]),num(r[yk])]).filter(p=>p[0]!==null&&p[1]!==null).sort((a,b)=>a[0]-b[0]),n=pairs.length;if(n<3)return {x:[],y:[]};const span=Math.max(3,Math.ceil(frac*n)),indices=uniq(Array.from({length:Math.min(pointCount,n)},(_,i)=>Math.round(i*(n-1)/(Math.min(pointCount,n)-1)))),xs=[],ys=[];indices.forEach(index=>{const x0=pairs[index][0],distances=pairs.map(p=>Math.abs(p[0]-x0)).sort((a,b)=>a-b),h=distances[Math.min(span-1,n-1)],weighted=pairs.map(p=>{const u=h>0?Math.min(1,Math.abs(p[0]-x0)/h):(p[0]===x0?0:1),w=u<1?Math.pow(1-Math.pow(u,3),3):0;return [p[0],p[1],w]}).filter(p=>p[2]>0);let sw=0,sx=0,sy=0,sxx=0,sxy=0;weighted.forEach(p=>{sw+=p[2];sx+=p[2]*p[0];sy+=p[2]*p[1];sxx+=p[2]*p[0]*p[0];sxy+=p[2]*p[0]*p[1]});const denominator=sw*sxx-sx*sx,b=denominator?((sw*sxy-sx*sy)/denominator):0,a=sw?((sy-b*sx)/sw):null;if(a!==null){xs.push(x0);ys.push(a+b*x0)}});return {x:xs,y:ys}}
function centralTrendRows(rows,xk,lower=.01,upper=.99){const values=rows.map(r=>num(r[xk])).filter(v=>v!==null).sort((a,b)=>a-b);if(!values.length)return [];const lo=values[Math.floor((values.length-1)*lower)],hi=values[Math.ceil((values.length-1)*upper)];return rows.filter(r=>num(r[xk])!==null&&num(r[xk])>=lo&&num(r[xk])<=hi)}
function driverScatter(id,rows,kind,renderChart=true){const xk=kind==='sea'?'sea_temp':'me_load_pct_mcr',xlab=kind==='sea'?'Sea-Water Temperature (°C)':'Main-Engine Load (% MCR)',valid=rows.filter(r=>num(r.running_hours)>0&&between(r.avg_load,0,100)&&(kind==='sea'?between(r.sea_temp,Number.MIN_VALUE,40):between(r.me_load_pct_mcr,Number.MIN_VALUE,110)));const overall=spearman(valid,xk,'avg_load',3),traces=uniq(valid.map(r=>r.vessel)).map(v=>{const z=valid.filter(r=>r.vessel===v),custom=kind==='sea'?z.map(r=>[r.vessel_size,r.report_date,r.running_hours]):z.map(r=>[r.vessel_size,r.report_date,r.me_load,r.mcr_kw,r.running_hours]);return {type:'scattergl',mode:'markers',name:v,x:z.map(r=>r[xk]),y:z.map(r=>r.avg_load),customdata:custom,showlegend:kind==='sea',marker:{size:5,color:'#6f9db5',opacity:kind==='sea'?0.65:0.16},hovertemplate:kind==='sea'?`Vessel: ${v}<br>Vessel Size: %{customdata[0]}<br>Date: %{customdata[1]}<br>Sea-Water Temperature: %{x:.2f} °C<br>SW-Pump Load: %{y:.2f}%<br>SW Running Hours: %{customdata[2]:.2f}<extra></extra>`:`Vessel: ${v}<br>Vessel Size: %{customdata[0]}<br>Date: %{customdata[1]}<br>Raw ME Load: %{customdata[2]:.2f}<br>MCR: %{customdata[3]:.2f} kW<br>ME Load % MCR: %{x:.2f}%<br>SW-Pump Load: %{y:.2f}%<br>SW Running Hours: %{customdata[4]:.2f}<extra></extra>`}});if(kind==='me'){const trend=lowessTrend(valid,xk,'avg_load');traces.push({type:'scatter',mode:'lines',name:'LOWESS trend',x:trend.x,y:trend.y,line:{color:C.blue,width:3.5},hovertemplate:'LOWESS trend<br>Main-Engine Load: %{x:.2f}% MCR<br>Smoothed Daily SW-Pump Load: %{y:.2f}%<extra></extra>'})}const note=`Spearman ρ = ${overall.rho===null?'—':fmt(overall.rho,3)}; n = ${overall.n.toLocaleString()}`;if(renderChart)plot(id,traces,layout('Daily weighted SW-pump load (%)',{xaxis:{title:xlab,gridcolor:'#eef1f5'},yaxis:{range:[50,100]},annotations:[{text:note,xref:'paper',yref:'paper',x:.5,y:.98,showarrow:false,font:{size:12,color:C.blue}}],margin:{l:75,r:25,t:35,b:75}}));const results=[{driver:kind==='sea'?'Sea-Water Temperature':'Main-Engine Load (% MCR)',vessel:'Overall',size:'All selected',rho:overall.rho,n:overall.n}];uniq(valid.map(r=>r.vessel)).forEach(v=>{const z=valid.filter(r=>r.vessel===v),s=spearman(z,xk,'avg_load',20);if(s.rho!==null)results.push({driver:kind==='sea'?'Sea-Water Temperature':'Main-Engine Load (% MCR)',vessel:v,size:z[0].vessel_size,rho:s.rho,n:s.n})});return results}
function renderSeaTemperatureBins(){const segments=['Port Stay','Sea Passage'],binStarts=Array.from({length:20},(_,i)=>i*2),binCentres=binStarts.map(x=>x+1),binLabels=binStarts.map(x=>`${x}–${x+2} °C`),valid=filtered(DATA.dailyCategory).filter(r=>r.equipment_category==='SW Pumps'&&segments.includes(r.segment_type)&&num(r.running_hours)>0&&between(r.avg_load,0,100)&&num(r.sea_temp)>0&&num(r.sea_temp)<=40),selectedSegments=segments.filter(segment=>fSegment.value==='All'||fSegment.value===segment),traces=selectedSegments.map(segment=>{const cells=new Map(binStarts.map(start=>[start,[]]));valid.filter(r=>r.segment_type===segment).forEach(r=>{const start=Math.min(38,Math.floor(num(r.sea_temp)/2)*2);cells.get(start).push(r)});const stats=binStarts.map((start,i)=>{const rows=cells.get(start),hours=sum(rows,'running_hours'),load=rows.length>=20?weighted(rows):null;return {label:binLabels[i],load,n:rows.length,hours}});return {type:'scatter',mode:'lines+markers',name:segment,x:binCentres,y:stats.map(x=>x.load),customdata:stats.map(x=>[segment,x.label,x.load,x.n,x.hours]),connectgaps:false,line:{color:segment==='Port Stay'?C.actual:C.orange,width:2.8},marker:{color:segment==='Port Stay'?C.actual:C.orange,size:7},hovertemplate:'Operational Segment: %{customdata[0]}<br>Temperature Bin: %{customdata[1]}<br>Running-Hour-Weighted SW-Pump Load: %{customdata[2]:.2f}%<br>Valid Daily Observations: %{customdata[3]}<br>Total SW-Pump Running Hours: %{customdata[4]:,.2f}<extra></extra>'}});plot('seaTempLoad',traces,layout('Running-Hour-Weighted Daily SW-Pump Load (%)',{xaxis:{title:'Sea-Water Temperature Bin (°C)',range:[0,40],tickmode:'array',tickvals:binCentres,ticktext:binLabels,tickangle:-45,gridcolor:'#eef1f5'},yaxis:{title:'Running-Hour-Weighted Daily SW-Pump Load (%)',range:[0,100],gridcolor:'#e9edf2',zeroline:false}}))}
function validPidRows(rows){return rows.filter(r=>num(r.sea_temp)>0&&num(r.sea_temp)<=40&&num(r.pid_setpoint)>0&&between(r.avg_load,0,100)&&num(r.running_hours)>0)}
function validMultivariableRows(rows){return rows.filter(r=>num(r.sea_temp)>0&&num(r.sea_temp)<=40&&num(r.pid_setpoint)>0&&between(r.avg_load,0,100)&&num(r.running_hours)>0&&num(r.me_load_pct_mcr)>0&&num(r.me_load_pct_mcr)<=110)}
function systematicRenderSample(rows,limit=8000){if(rows.length<=limit)return rows;return Array.from({length:limit},(_,i)=>rows[Math.floor(i*(rows.length-1)/(limit-1))])}
function renderMultivariableScatter(allRows){const rows=systematicRenderSample(allRows),symbols={'Port Stay':'circle','Sea Passage':'diamond','Other / Unknown':'square'};plot('multivariableScatter',[{type:'scattergl',mode:'markers',x:rows.map(r=>r.sea_temp),y:rows.map(r=>r.avg_load),customdata:rows.map(r=>[r.vessel,r.vessel_size,r.report_date,r.segment_type,r.pid_setpoint,r.cooling_temperature_margin,r.me_load_pct_mcr,r.running_hours]),marker:{size:rows.map(r=>6+Math.min(110,Math.max(0,r.me_load_pct_mcr))/110*18),symbol:rows.map(r=>symbols[r.segment_type]||'square'),opacity:.62,color:rows.map(r=>r.pid_setpoint),colorscale:'Viridis',showscale:true,colorbar:{title:{text:'PID Temperature<br>Setpoint (°C)'}}},hovertemplate:'Vessel: %{customdata[0]}<br>Vessel Size: %{customdata[1]}<br>Date: %{customdata[2]}<br>Segment Type: %{customdata[3]}<br>Sea-Water Temperature: %{x:.2f} °C<br>PID Temperature Setpoint: %{customdata[4]:.2f} °C<br>Cooling Temperature Margin: %{customdata[5]:.2f} °C<br>Main-Engine Load: %{customdata[6]:.2f}% MCR<br>Daily Weighted SW-Pump Load: %{y:.2f}%<br>SW-Pump Running Hours: %{customdata[7]:.2f}<extra></extra>'}],layout('Daily Weighted SW-Pump Load (%)',{xaxis:{title:'Sea-Water Temperature (°C)',gridcolor:'#eef1f5'},annotations:[{text:`Valid observations: ${allRows.length.toLocaleString()} · Rendered observations: ${rows.length.toLocaleString()}`,xref:'paper',yref:'paper',x:0,y:1.08,showarrow:false,font:{size:12,color:C.blue}}],margin:{l:65,r:115,t:45,b:75}}))}
function renderPidInteractionHeatmap(rows){const cells=new Map();rows.forEach(r=>{const seaBin=Math.floor(r.sea_temp/2)*2,pidBin=Math.floor(r.pid_setpoint),key=seaBin+'|'+pidBin;if(!cells.has(key))cells.set(key,{seaBin,pidBin,n:0,hours:0,loadHours:0,me:0,margin:0});const c=cells.get(key);c.n++;c.hours+=r.running_hours;c.loadHours+=r.avg_load*r.running_hours;c.me+=r.me_load_pct_mcr;c.margin+=r.cooling_temperature_margin});const seaBins=uniq([...cells.values()].map(c=>c.seaBin)).sort((a,b)=>a-b),pidBins=uniq([...cells.values()].map(c=>c.pidBin)).sort((a,b)=>a-b),x=seaBins.map(v=>`${v}–${v+2} °C`),y=pidBins.map(v=>`${v}–${v+1} °C`),annotations=[];let sufficient=0,na=0;const z=pidBins.map((p,yi)=>seaBins.map((s,xi)=>{const c=cells.get(s+'|'+p);if(!c||c.n<20){na++;annotations.push({text:'N/A',x:x[xi],y:y[yi],showarrow:false,font:{size:9,color:'#4b5563'}});return null}sufficient++;return c.loadHours/c.hours})),custom=pidBins.map(p=>seaBins.map(s=>{const c=cells.get(s+'|'+p);return c?[`${s}–${s+2} °C`,`${p}–${p+1} °C`,c.loadHours/c.hours,c.me/c.n,c.margin/c.n,c.n,c.hours]:[`${s}–${s+2} °C`,`${p}–${p+1} °C`,null,null,null,0,0]}));plot('pidInteractionHeatmap',[{type:'heatmap',x,y,z,customdata:custom,zmin:0,zmax:100,colorscale:'YlOrRd',colorbar:{title:{text:'Weighted SW-Pump<br>Load (%)'}},hoverongaps:false,hovertemplate:'Sea-Water Temperature Bin: %{customdata[0]}<br>PID Temperature Setpoint Bin: %{customdata[1]}<br>Weighted Average SW-Pump Load: %{customdata[2]:.2f}%<br>Average Main-Engine Load: %{customdata[3]:.2f}% MCR<br>Average Cooling Temperature Margin: %{customdata[4]:.2f} °C<br>Valid Observations: %{customdata[5]}<br>Total SW-Pump Running Hours: %{customdata[6]:.2f}<extra></extra>'}],layout('',{xaxis:{title:'Sea-Water Temperature bins',gridcolor:'#eef1f5'},yaxis:{title:'PID Temperature Setpoint bins',gridcolor:'#eef1f5'},annotations:[...annotations,{text:`Valid observations: ${rows.length.toLocaleString()} · Sufficient cells: ${sufficient} · N/A cells: ${na}`,xref:'paper',yref:'paper',x:0,y:1.1,showarrow:false,font:{size:12,color:C.blue}}],plot_bgcolor:'#c7cdd4',margin:{l:105,r:105,t:50,b:90}}))}
function renderCoolingMarginChart(rows){const overall=spearman(rows,'cooling_temperature_margin','avg_load',3),trend=lowessTrend(centralTrendRows(rows,'cooling_temperature_margin'),'cooling_temperature_margin','avg_load'),traces=[{type:'scattergl',mode:'markers',name:'Daily observations',x:rows.map(r=>r.cooling_temperature_margin),y:rows.map(r=>r.avg_load),customdata:rows.map(r=>[r.vessel,r.vessel_size,r.report_date,r.segment_type,r.sea_temp,r.pid_setpoint,r.running_hours]),marker:{size:3,color:'#6f9db5',opacity:.12},hovertemplate:'Vessel: %{customdata[0]}<br>Vessel Size: %{customdata[1]}<br>Date: %{customdata[2]}<br>Segment Type: %{customdata[3]}<br>Sea-Water Temperature: %{customdata[4]:.2f} °C<br>PID Temperature Setpoint: %{customdata[5]:.2f} °C<br>Cooling Temperature Margin: %{x:.2f} °C<br>Daily Weighted SW-Pump Load: %{y:.2f}%<br>SW-Pump Running Hours: %{customdata[6]:.2f}<extra></extra>'},{type:'scatter',mode:'lines',name:'LOWESS trend',x:trend.x,y:trend.y,line:{color:C.blue,width:3.5},hovertemplate:'LOWESS trend<br>Cooling Temperature Margin: %{x:.2f} °C<br>Smoothed Daily SW-Pump Load: %{y:.2f}%<extra></extra>'}],note=`Spearman ρ = ${overall.rho===null?'—':fmt(overall.rho,3)}; n = ${overall.n.toLocaleString()}`;plot('coolingMarginLoad',traces,layout('Daily Weighted SW-Pump Load (%)',{xaxis:{title:'Cooling Temperature Margin (°C)',gridcolor:'#eef1f5'},yaxis:{range:[50,100]},annotations:[{text:note,xref:'paper',yref:'paper',x:.5,y:.98,showarrow:false,font:{size:12,color:C.blue}}],margin:{l:75,r:25,t:35,b:75}}));const results=[{driver:'Cooling Temperature Margin',vessel:'Overall',size:'All selected',rho:overall.rho,n:overall.n}];uniq(rows.map(r=>r.vessel)).forEach(v=>{const z=rows.filter(r=>r.vessel===v),s=spearman(z,'cooling_temperature_margin','avg_load',20);if(s.rho!==null)results.push({driver:'Cooling Temperature Margin',vessel:v,size:z[0].vessel_size,rho:s.rho,n:s.n})});return results}
function renderDrivers(){const dc=driverDaily(),sw=dc.filter(x=>x.equipment_category==='SW Pumps'),pid=validPidRows(sw),multiDaily=multivariableDaily(),multi=validMultivariableRows(multiDaily),driverResults=[...driverScatter('seaTempLoad',sw,'sea',false),...driverScatter('meLoadSW',sw,'me')];renderSeaTemperatureBins();renderMultivariableScatter(multi);renderPidInteractionHeatmap(multi);driverResults.push(...renderCoolingMarginChart(pid));table('driverCorrTable',['Driver','Vessel','Vessel Size','Spearman ρ','Valid Observations'],driverResults.map(x=>[x.driver,x.vessel,x.size,fmt(x.rho,3),fmt(x.n,0)]));
 dc.forEach(r=>{r.sea_corr=num(r.sea_temp)>0&&num(r.sea_temp)<=40?num(r.sea_temp):null;r.me_corr=num(r.me_load_pct_mcr)>0&&num(r.me_load_pct_mcr)<=110?num(r.me_load_pct_mcr):null;r.load_corr=num(r.running_hours)>0&&num(r.avg_load)>=0&&num(r.avg_load)<=100?num(r.avg_load):null;r.manual_corr=num(r.manual_pct)>=0&&num(r.manual_pct)<=100?num(r.manual_pct):null;r.bypass_corr=num(r.bypass_pct)>=0&&num(r.bypass_pct)<=100?num(r.bypass_pct):null;r.saving_corr=num(r.baseline_energy)>0&&num(r.saving_pct)>=-100&&num(r.saving_pct)<=100?num(r.saving_pct):null;r.actual_corr=num(r.actual_energy)>=0?num(r.actual_energy):null;r.energy_hour_corr=num(r.running_hours)>0&&num(r.actual_energy)>=0?num(r.energy_per_hour):null;r.saving_hour_corr=num(r.running_hours)>0?num(r.saving_per_hour):null;const validPid=r.equipment_category==='SW Pumps'&&num(r.sea_temp)>0&&num(r.sea_temp)<=40&&num(r.pid_setpoint)>0&&between(r.avg_load,0,100)&&num(r.running_hours)>0;r.pid_corr=validPid?r.pid_setpoint:null;r.cooling_margin_corr=validPid?r.cooling_temperature_margin:null});
 const matrixRows=multiDaily.map(r=>{const validSea=num(r.sea_temp)>0&&num(r.sea_temp)<=40,validPid=num(r.pid_setpoint)>0,validHours=num(r.running_hours)>0;return {...r,avg_load:validHours&&between(r.avg_load,0,100)?num(r.avg_load):null,sea_temp:validSea?num(r.sea_temp):null,pid_setpoint:validPid?num(r.pid_setpoint):null,cooling_temperature_margin:validSea&&validPid?num(r.cooling_temperature_margin):null,me_load_pct_mcr:num(r.me_load_pct_mcr)>0&&num(r.me_load_pct_mcr)<=110?num(r.me_load_pct_mcr):null,running_hours:validHours?num(r.running_hours):null}}),matrixMetrics=[['avg_load','Weighted Daily SW-Pump Load'],['sea_temp','Sea-Water Temperature'],['pid_setpoint','PID Temperature Setpoint'],['cooling_temperature_margin','Cooling Temperature Margin'],['me_load_pct_mcr','ME Load % MCR'],['running_hours','SW-Pump Running Hours']],matrixStats=matrixMetrics.map(a=>matrixMetrics.map(b=>spearman(matrixRows,a[0],b[0],3))),matrix=matrixStats.map(row=>row.map(x=>x.rho)),sampleSizes=matrixStats.map(row=>row.map(x=>x.n)),missingAnnotations=[];matrix.forEach((row,i)=>row.forEach((value,j)=>{if(value===null)missingAnnotations.push({text:'N/A',x:matrixMetrics[j][1],y:matrixMetrics[i][1],showarrow:false,font:{size:10,color:'#4b5563'}})}));plot('corrMatrix',[{type:'heatmap',z:matrix,x:matrixMetrics.map(x=>x[1]),y:matrixMetrics.map(x=>x[1]),customdata:sampleSizes,zmin:-1,zmax:1,colorscale:[[0,'#b2182b'],[.5,'#f7f7f7'],[1,'#2166ac']],colorbar:{title:'ρ'},hoverongaps:false,hovertemplate:'%{y} ↔ %{x}<br>Spearman ρ: %{z:.3f}<br>Valid paired observations: %{customdata}<extra></extra>'}],layout('',{margin:{l:205,r:35,t:15,b:155},plot_bgcolor:'#c7cdd4',annotations:missingAnnotations}));const metrics=[['sea_corr','Sea-Water Temperature'],['pid_corr','PID Temperature Setpoint'],['cooling_margin_corr','Cooling Temperature Margin'],['me_corr','ME Load % MCR'],['load_corr','Weighted VFD Load %'],['manual_corr','Manual Operation %'],['bypass_corr','Bypass Operation %'],['saving_corr','Estimated Saving %'],['actual_corr','Actual Category Energy (kWh)'],['energy_hour_corr','Energy per VFD Running Hour'],['saving_hour_corr','Saving per VFD Running Hour']],pairs=[];for(let i=0;i<metrics.length;i++)for(let j=i+1;j<metrics.length;j++){const z=spearman(dc,metrics[i][0],metrics[j][0],20);if(z.rho!==null)pairs.push([metrics[i][1]+' ↔ '+metrics[j][1],z.rho,z.n])}pairs.sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));table('corrTable',['Exploratory Association','Spearman ρ','Valid Paired Observations'],pairs.map(x=>[x[0],fmt(x[1],3),fmt(x[2],0)]));}
const renderers={overview:renderOverview,utilisation:renderUtilisation,energy:renderEnergy,drivers:renderDrivers};let active='overview',dirty=new Set(Object.keys(renderers));function render(){renderers[active]();dirty.delete(active);setTimeout(()=>window.dispatchEvent(new Event('resize')),0)}
document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('.nav button').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.section').forEach(x=>x.classList.toggle('active',x.id===b.dataset.section));active=b.dataset.section;if(dirty.has(active))render()});
filterIds.forEach(id=>document.getElementById(id).onchange=()=>{dirty=new Set(Object.keys(renderers));render()});fGroup.onchange=()=>{groupSelectionTouched=true;dirty=new Set(Object.keys(renderers));render()};reset.onclick=()=>{filterIds.forEach(id=>document.getElementById(id).value='All');fGroup.value='Vessel';groupSelectionTouched=false;dirty=new Set(Object.keys(renderers));render()};render();
</script></body></html>'''


def build_dashboard() -> dict:
    interval, daily_equipment, daily_category, daily_operational, vessel_day, box, metadata = load_data()
    parallel_imbalance, available_vessel_equipment_units, unavailable_heatmap_cells = build_parallel_imbalance(interval)
    daily_sw_multivariable = build_multivariable_sw_daily(interval)
    sw_daily = daily_category[daily_category["equipment_category"] == "SW Pumps"]
    sea_valid = sw_daily[
        sw_daily["sea_temp"].gt(0) & sw_daily["sea_temp"].le(40)
        & sw_daily["avg_load"].ge(0) & sw_daily["avg_load"].le(100)
        & sw_daily["running_hours"].gt(0)
    ]
    me_valid = sw_daily[
        sw_daily["me_load_pct_mcr"].gt(0) & sw_daily["me_load_pct_mcr"].le(110)
        & sw_daily["avg_load"].ge(0) & sw_daily["avg_load"].le(100)
        & sw_daily["running_hours"].gt(0)
    ]
    pid_valid_mask = (
        sw_daily["sea_temp"].gt(0) & sw_daily["sea_temp"].le(40)
        & sw_daily["pid_setpoint"].gt(0)
        & sw_daily["avg_load"].ge(0) & sw_daily["avg_load"].le(100)
        & sw_daily["running_hours"].gt(0)
    )
    pid_valid = sw_daily[pid_valid_mask]
    cooling_rho, cooling_count = spearman_pair(
        pid_valid, "cooling_temperature_margin", "avg_load"
    )
    vessel_cooling_correlations = 0
    for _, vessel_rows in pid_valid.groupby("vessel"):
        vessel_rho, _ = spearman_pair(
            vessel_rows, "cooling_temperature_margin", "avg_load", minimum=20
        )
        if pd.notna(vessel_rho):
            vessel_cooling_correlations += 1
    multivariable_valid_mask = (
        daily_sw_multivariable["sea_temp"].gt(0) & daily_sw_multivariable["sea_temp"].le(40)
        & daily_sw_multivariable["pid_setpoint"].gt(0)
        & daily_sw_multivariable["avg_load"].ge(0) & daily_sw_multivariable["avg_load"].le(100)
        & daily_sw_multivariable["running_hours"].gt(0)
        & daily_sw_multivariable["me_load_pct_mcr"].gt(0) & daily_sw_multivariable["me_load_pct_mcr"].le(110)
    )
    multivariable_valid = daily_sw_multivariable[multivariable_valid_mask].copy()
    multivariable_valid["sea_temp_bin"] = np.floor(multivariable_valid["sea_temp"] / 2) * 2
    multivariable_valid["pid_setpoint_bin"] = np.floor(multivariable_valid["pid_setpoint"])
    heatmap_counts = multivariable_valid.groupby(
        ["sea_temp_bin", "pid_setpoint_bin"], dropna=False
    ).size()
    sufficient_heatmap_cells = int(heatmap_counts.ge(20).sum())
    heatmap_grid_cells = int(
        multivariable_valid["sea_temp_bin"].nunique()
        * multivariable_valid["pid_setpoint_bin"].nunique()
    )
    valid_daily_trend_rows = interval[
        interval["avg_load"].notna() & interval["running_hours"].gt(0)
    ]
    daily_trend_dates = pd.to_datetime(
        valid_daily_trend_rows["report_date"].dropna().unique(), errors="coerce"
    )
    invalid_contribution = vessel_day[
        vessel_day["total_ae_energy"].notna() & vessel_day["total_vfd_energy"].notna()
        & vessel_day["total_vfd_energy"].gt(vessel_day["total_ae_energy"])
    ]
    metadata.update({
        "valid_sea_driver_observations": len(sea_valid),
        "valid_me_driver_observations": len(me_valid),
        "valid_daily_pid_observations": len(pid_valid),
        "excluded_invalid_pid_observations": int((~pid_valid_mask).sum()),
        "cooling_margin_observations": cooling_count,
        "overall_cooling_margin_spearman_rho": cooling_rho,
        "vessels_with_valid_cooling_margin_correlations": vessel_cooling_correlations,
        "valid_multivariable_observations": len(multivariable_valid),
        "rendered_multivariable_scatter_observations": min(len(multivariable_valid), 8000),
        "heatmap_sufficient_cells": sufficient_heatmap_cells,
        "heatmap_na_cells": heatmap_grid_cells - sufficient_heatmap_cells,
        "multivariable_sea_temp_min": multivariable_valid["sea_temp"].min(),
        "multivariable_sea_temp_max": multivariable_valid["sea_temp"].max(),
        "multivariable_pid_min": multivariable_valid["pid_setpoint"].min(),
        "multivariable_pid_max": multivariable_valid["pid_setpoint"].max(),
        "multivariable_me_load_min": multivariable_valid["me_load_pct_mcr"].min(),
        "multivariable_me_load_max": multivariable_valid["me_load_pct_mcr"].max(),
        "available_vessel_equipment_units": available_vessel_equipment_units,
        "parallel_heatmap_vessels": int(interval["vessel"].nunique()),
        "sw_valid_imbalance_groups": int(parallel_imbalance.loc[
            parallel_imbalance["equipment_category"].eq("SW Pumps"),
            "normalised_running_hour_imbalance",
        ].notna().sum()),
        "fan_valid_imbalance_groups": int(parallel_imbalance.loc[
            parallel_imbalance["equipment_category"].eq("Engine-Room Fans"),
            "normalised_running_hour_imbalance",
        ].notna().sum()),
        "unavailable_parallel_heatmap_cells": unavailable_heatmap_cells,
        "daily_trend_observations": int(len(daily_trend_dates)),
        "daily_trend_date_min": pd.Series(daily_trend_dates).min().strftime("%Y-%m-%d"),
        "daily_trend_date_max": pd.Series(daily_trend_dates).max().strftime("%Y-%m-%d"),
        "invalid_contribution_observations": len(invalid_contribution),
        "correlation_matrix_observations": int(daily_category["running_hours"].gt(0).sum()),
    })
    payload = {
        "interval": records(interval), "dailyEquipment": records(daily_equipment),
        "dailyCategory": records(daily_category), "dailyOperational": records(daily_operational),
        "dailySWMultivariable": records(daily_sw_multivariable),
        "parallelImbalance": records(parallel_imbalance),
        "vesselDay": records(vessel_day),
        "box": records(box), "meta": metadata,
    }
    html = HTML.replace("__PLOTLY__", get_plotlyjs()).replace(
        "__DATA__", json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    )
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    external_scripts = re.findall(r"<script\b[^>]*\bsrc\s*=\s*['\"]([^'\"]+)", html, re.I)
    if external_scripts:
        raise RuntimeError("Offline check failed: external script reference(s) are present")
    if "Plotly.newPlot" not in html and "plotly.js" not in html.lower():
        raise RuntimeError("Offline check failed: Plotly JavaScript is not embedded")
    if not np.isclose(metadata["report_ae_sum"], metadata["vessel_day_ae_sum"], equal_nan=True):
        raise RuntimeError("AE-energy validation failed: vessel-day total differs from unique report intervals")
    if any(label in HTML for label in (
        "Manual operation versus estimated saving percentage",
        "Bypass operation versus estimated saving percentage",
        "Fan ER temperature midpoint versus fan load",
    )):
        raise RuntimeError("Removed Operational Drivers scatter card is still present")
    if 'id="pidLoad"' in HTML or "PID Setpoint vs Daily SW-Pump Load" in HTML:
        raise RuntimeError("Old standalone PID chart is still present")
    if any(chart_id not in HTML for chart_id in (
        'id="meLoadSW"', 'id="multivariableScatter"', 'id="pidInteractionHeatmap"',
        'id="coolingMarginLoad"', 'id="corrMatrix"'
    )):
        raise RuntimeError("One or more multivariable Operational Drivers charts are missing")
    if "PID Temperature Setpoint" not in HTML or "Cooling Temperature Margin" not in HTML:
        raise RuntimeError("PID variables are missing from the Operational Drivers analysis")
    if "N/A" not in HTML:
        raise RuntimeError("Missing matrix-cell N/A display is not configured")
    if "scatter3d" in HTML.lower():
        raise RuntimeError("A prohibited 3D chart is present")
    if 'id="corrMatrix"' not in HTML:
        raise RuntimeError("Spearman correlation matrix is missing")
    required_interactions = {
        "sea_temp_pid_interaction", "cooling_margin_me_load_interaction",
        "sea_temp_me_load_interaction", "pid_me_load_interaction",
    }
    if not required_interactions.issubset(daily_sw_multivariable.columns):
        raise RuntimeError("One or more machine-learning-ready interaction fields are missing")
    if 'id="hourBalance"' in HTML or "Running-hour balance between parallel units" in HTML:
        raise RuntimeError("Old fleet-total parallel-unit chart is still present")
    if 'id="parallelHeatmap"' not in HTML:
        raise RuntimeError("Parallel-unit running-hour heatmap is missing")
    if "30-Day Rolling Average" not in HTML:
        raise RuntimeError("Daily utilisation rolling average is missing")
    if "filtered(DATA.dailyEquipment,true)" in HTML:
        raise RuntimeError("Daily utilisation trend still ignores the Segment Type filter")
    for required_label in (
        "Available Equipment Units", "Share Above 80%", "Share Above 90%",
        "Normalised Group Imbalance", "Equal-Use Reference Share",
    ):
        if required_label not in HTML:
            raise RuntimeError(f"Required utilisation output is missing: {required_label}")
    if any(term in HTML for term in ("pid_lower", "pid_upper", "pid_range_width")):
        raise RuntimeError("Unexpected PID range parsing logic remains")
    calculated_margin = daily_category["pid_setpoint"] - daily_category["sea_temp"]
    if not np.allclose(
        daily_category["cooling_temperature_margin"], calculated_margin, equal_nan=True
    ):
        raise RuntimeError("Cooling Temperature Margin validation failed")
    if re.search(r"\bVFD Energy\b", HTML, re.I):
        raise RuntimeError("Ambiguous visible label 'VFD Energy' remains")
    metadata["html_size"] = OUTPUT_HTML.stat().st_size
    metadata["output_path"] = str(OUTPUT_HTML)
    return metadata


if __name__ == "__main__":
    result = build_dashboard()
    print(f"Available vessel-equipment units: {result['available_vessel_equipment_units']}")
    print(f"Vessels included in parallel-unit heatmap: {result['parallel_heatmap_vessels']}")
    print(f"SW groups with valid imbalance calculation: {result['sw_valid_imbalance_groups']}")
    print(f"Fan groups with valid imbalance calculation: {result['fan_valid_imbalance_groups']}")
    print(f"Unavailable heatmap cells displayed as N/A: {result['unavailable_parallel_heatmap_cells']}")
    print(f"Daily trend observations: {result['daily_trend_observations']}")
    print(f"Daily trend date range: {result['daily_trend_date_min']} to {result['daily_trend_date_max']}")
    print(f"Output HTML path: {result['output_path']}")
    print(f"HTML file size: {result['html_size']} bytes")
