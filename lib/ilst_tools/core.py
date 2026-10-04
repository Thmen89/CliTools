from __future__ import annotations

import csv
import getpass
import json
import os
import re
import tempfile
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

FIELDS = {
    "ENTITY_GROUP": ["ENTITYGROUP", "ENTITYGRP", "GROUP", "GROUPNAME"],
    "LEGAL_ENTITY": ["LEGALENTITY", "LEGALENTITYCODE", "LECODE", "LE", "ENTITYCODE"],
    "CONTINGENCY": ["CONTINGENCY", "CONTINGENCYNAME"],
    "SCENARIO": ["SCENARIO", "SCENARIONAME", "STRESSSCENARIO"],
    "TIME_HORIZON": ["TIMEHORIZON", "HORIZON", "DAY", "DAYS"],
    "BUCKET_CFP": ["BUCKETCFP", "CFPBUCKET", "BUCKET", "TIMEBUCKET"],
    "DATA_TYPE": ["DATATYPE", "SUBTYPE"],
    "CURRENCY": ["CURRENCY", "CCY", "CURRENCYCODE"],
    "AMOUNT": ["AMOUNT", "AMOUNTUSD", "USDAMOUNT", "AMT", "VALUE"],
}
REQUIRED = ("CONTINGENCY", "SCENARIO", "BUCKET_CFP", "AMOUNT")
LABELS = {"ENTITY_GROUP": "Entity group", "LEGAL_ENTITY": "Legal entity", "CONTINGENCY": "Contingency",
          "SCENARIO": "Scenario", "BUCKET_CFP": "Bucket", "DATA_TYPE": "Data type", "CURRENCY": "Currency",
          "TIME_HORIZON": "Day"}
BLANK = "(blank)"
HORIZON_END = "M12"
FULL_UNTIL_DAYS = 31
HORIZON_END_DAYS = 366


class DataError(Exception):
    pass


def _norm(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def parse_date(text: str) -> str:
    t = str(text).strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%d/%m/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(t, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    raise ValueError(f"'{text}' is not a valid date - use YYYY-MM-DD.")


def fmt_date(d: str) -> str:
    return datetime.strptime(d, "%Y%m%d").strftime("%d-%b-%Y")


def list_reports(data_dir: str, pattern: str) -> dict:
    folder = Path(data_dir)
    if not folder.is_dir():
        raise DataError(f"Data folder not reachable: {folder}")
    rx = re.compile(re.escape(pattern).replace(re.escape("{date}"), r"(\d{8})") + r"$", re.I)
    found = {}
    for p in folder.glob(pattern.replace("{date}", "*")):
        m = rx.match(p.name)
        if m:
            try:
                found[parse_date(m.group(1))] = p
            except ValueError:
                raise DataError(f"{p.name}: '{m.group(1)}' in the file name is not a valid date.") from None
    if not found:
        raise DataError(f"No files matching '{pattern}' in {folder}")
    return dict(sorted(found.items()))


def _find_header(path: Path, extra_aliases: dict | None) -> tuple:
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        lines = [f.readline() for _ in range(30)]
    first_error = None
    for skip, line in enumerate(lines):
        if not line.strip():
            continue
        delim = max(",;|\t", key=line.count)
        header = next(csv.reader([line.rstrip("\r\n")], delimiter=delim))
        try:
            return delim, header, skip, resolve_columns(header, extra_aliases)
        except DataError as e:
            first_error = first_error or e
    raise first_error or DataError(f"{path.name} is empty.")


def resolve_columns(header: list, extra_aliases: dict | None = None) -> dict:
    normed = [_norm(h) for h in header]
    used, out = set(), {}
    for fld, aliases in FIELDS.items():
        extra = (extra_aliases or {}).get(fld, [])
        extra = [extra] if isinstance(extra, str) else extra
        for alias in [_norm(a) for a in extra] + [_norm(fld)] + aliases:
            hits = [i for i, n in enumerate(normed) if n == alias and i not in used]
            if hits:
                out[fld] = hits[0]
                used.add(hits[0])
                break
    missing = [f for f in REQUIRED if f not in out]
    if missing:
        raise DataError(f"Column(s) {', '.join(missing)} not found. Header has: {', '.join(header)}. "
                        'If a column was renamed, add it to column_aliases, e.g. {"AMOUNT": ["NEW_NAME"]}.')
    return out


@dataclass
class Report:
    date: str
    path: Path
    df: pd.DataFrame
    bucket_days: dict

    def require(self, *fields: str, why: str = "") -> None:
        missing = [f for f in fields if f and f not in self.df.columns]
        if missing:
            raise DataError(f"{self.path.name} has no {', '.join(LABELS.get(f, f) for f in missing)} column"
                            + (f" - needed for {why}." if why else "."))

    def flow_buckets(self, upto: str | None = None) -> list:
        if upto is None:
            hi = HORIZON_END_DAYS
        elif upto in self.bucket_days:
            hi = self.bucket_days[upto]
        else:
            raise DataError(f"Bucket {upto} is not in {self.path.name}.")
        return [b for b, d in self.bucket_days.items() if 0 < d <= hi]

    def window_buckets(self) -> list:
        return [b for b, d in self.bucket_days.items() if FULL_UNTIL_DAYS < d <= HORIZON_END_DAYS]


_CACHE: OrderedDict = OrderedDict()
CACHE_SIZE = 6


def fingerprint(path) -> tuple:
    p = Path(path).resolve()
    st = p.stat()
    return str(p), st.st_mtime_ns, st.st_size


def load_report(date: str, path: Path, extra_aliases: dict | None = None) -> Report:
    path = Path(path)
    key = (fingerprint(path), date, json.dumps(extra_aliases or {}, sort_keys=True))
    if key in _CACHE:
        _CACHE.move_to_end(key)
        return _CACHE[key]
    delim, header, skip, cols = _find_header(path, extra_aliases)
    names = {header[i]: fld for fld, i in cols.items()}
    text = [h for h, f in names.items() if f not in ("AMOUNT", "TIME_HORIZON")]
    raw = _read_csv(path, list(names), text, delim, skip)
    by_name = {str(h).strip(): f for h, f in names.items()}
    df = pd.DataFrame(index=raw.index)
    for src in raw.columns:
        fld, col = by_name[str(src).strip()], raw[src]
        if fld == "AMOUNT":
            df[fld] = _to_amount(col, path.name)
        elif fld == "TIME_HORIZON":
            df[fld] = pd.to_numeric(col, errors="coerce")
        else:
            df[fld] = _to_category(col)
    rep = Report(date, path, df, _bucket_days(df, path.name))
    _CACHE[key] = rep
    while len(_CACHE) > CACHE_SIZE:
        _CACHE.popitem(last=False)
    return rep


def _read_csv(path: Path, columns: list, text: list, delim: str, skip: int = 0) -> pd.DataFrame:
    return pd.read_csv(path, usecols=columns, sep=delim, dtype={h: str for h in text}, skiprows=skip,
                       keep_default_na=False, na_values=[""], encoding="utf-8-sig")


def _to_category(col: pd.Series) -> pd.Categorical:
    cat = pd.Categorical(col)
    labels = pd.Index(cat.categories.astype(str)).str.strip()
    labels = labels.where(labels != "", BLANK)
    new_codes, uniq = pd.factorize(labels)
    codes = np.asarray(cat.codes)
    out = np.where(codes >= 0, new_codes[np.maximum(codes, 0)] if len(new_codes) else 0, -1)
    uniq = list(uniq)
    if (out < 0).any():
        if BLANK not in uniq:
            uniq.append(BLANK)
        out[out < 0] = uniq.index(BLANK)
    return pd.Categorical.from_codes(out, categories=uniq)


def _to_amount(col: pd.Series, source: str) -> pd.Series:
    if pd.api.types.is_numeric_dtype(col):
        v = col.astype("float64")
    else:
        s = col.astype("string").str.strip()
        neg = s.str.startswith("(", na=False) & s.str.endswith(")", na=False)
        v = pd.to_numeric(s.str.replace(r"[,\s()]", "", regex=True), errors="coerce").astype("float64")
        v[neg] = -v[neg].abs()
    bad = ~np.isfinite(v.to_numpy())
    if bad.any():
        rows = np.flatnonzero(bad)
        shown = ", ".join(f"row {r + 2}: '{col.iloc[r] if not pd.isna(col.iloc[r]) else ''}'" for r in rows[:3])
        raise DataError(f"{source}: {len(rows)} amount(s) are blank or not numbers ({shown}"
                        + (", …" if len(rows) > 3 else "") + "). Nothing was calculated.")
    return v


_UNIT_DAYS = {"D": 1, "DAY": 1, "DAYS": 1, "W": 7, "WK": 7, "WKS": 7, "WEEK": 7, "WEEKS": 7,
              "M": 30, "MTH": 30, "MTHS": 30, "MONTH": 30, "MONTHS": 30, "Y": 365, "YR": 365, "YEAR": 365}


def label_days(label) -> int | None:
    s = re.sub(r"[\s_\-]", "", str(label).upper())
    m = re.fullmatch(r"([A-Z]+)(\d+)", s) or re.fullmatch(r"(\d+)([A-Z]+)", s)
    if not m:
        return None
    a, b = m.groups()
    unit, n = (a, b) if a.isalpha() else (b, a)
    return int(n) * _UNIT_DAYS[unit] if unit in _UNIT_DAYS else None


def _bucket_days(df: pd.DataFrame, source: str) -> dict:
    days, unknown = {}, []
    for label in df["BUCKET_CFP"].cat.categories:
        d = label_days(label)
        if d is None and "TIME_HORIZON" in df.columns:
            d = df.loc[df["BUCKET_CFP"] == label, "TIME_HORIZON"].min()
        if d is None or pd.isna(d):
            unknown.append(label)
        else:
            days[label] = float(d)
    if unknown:
        raise DataError(f"{source}: cannot place bucket(s) {', '.join(unknown)} in time order "
                        "(label not recognised and no TIME_HORIZON). Nothing was calculated.")
    return dict(sorted(days.items(), key=lambda kv: kv[1]))


@dataclass
class Prepositioning:
    rules: dict = field(default_factory=dict)
    names: dict = field(default_factory=dict)
    path: Path | None = None

    def factor(self, group: str, date: str) -> float:
        f = 1.0
        for start, value in self.rules.get(str(group).upper(), []):
            if start <= date:
                f = value
        return f

    def active(self, groups, date: str) -> list:
        return [(g, self.factor(g, date)) for g in groups if self.factor(g, date) != 1.0]


def pct(f: float) -> str:
    return f"{round(f * 100, 4):g}%"


def parse_factor(value, where: str = "") -> float:
    s = str(value).strip()
    try:
        f = float(s.rstrip("%"))
    except ValueError:
        raise DataError(f"{where}'{value}' is not a percentage - use 75, 75% or 0.75.") from None
    f = f / 100 if (s.endswith("%") or f > 1) else f
    if not 0 <= f <= 1:
        raise DataError(f"{where}{value} must be between 0% and 100%.")
    return f


def _parse_prepos(data: dict, path: Path) -> Prepositioning:
    out = Prepositioning(path=path)
    for name, spec in (data.get("entity_groups") or {}).items():
        if name.startswith("_"):
            continue
        if not isinstance(spec, dict):
            raise DataError(f"{path.name}, {name}: expected {{\"YYYY-MM-DD\": factor}}.")
        steps = []
        for start, value in spec.items():
            try:
                day = parse_date(start)
            except ValueError as e:
                raise DataError(f"{path.name}, {name}: {e}") from None
            steps.append((day, parse_factor(value, f"{path.name}, {name}, {start}: ")))
        out.rules[name.upper()] = sorted(steps)
        out.names[name.upper()] = name
    return out


def load_prepositioning(path) -> Prepositioning:
    path = Path(path)
    if not path.exists():
        raise DataError(f"Prepositioning file not found: {path}. Binding needs it (it may list no factors).")
    return _parse_prepos(_read_json(path), path)


def set_prepositioning(path, group: str, start: str, factor: float | None, expect: list) -> Prepositioning:
    path = Path(path)

    def change(data):
        doc = _parse_prepos(data, path)
        if doc.rules.get(group.upper(), []) != list(expect):
            raise DataError(f"{group} was changed by someone else since you opened this screen. "
                            "Nothing saved - reload and try again.")
        groups = data.setdefault("entity_groups", {})
        key = next((k for k in groups if k.upper() == group.upper()), group)
        steps = {parse_date(k): v for k, v in groups.get(key, {}).items()}
        before = pct(dict(doc.rules.get(group.upper(), [])).get(start, 1.0)) if start in steps else "none"
        steps.pop(start, None)
        if factor is not None:
            steps[start] = pct(factor)
        if steps:
            groups[key] = {f"{d[:4]}-{d[4:6]}-{d[6:]}": v for d, v in sorted(steps.items())}
        else:
            groups.pop(key, None)
        data["_log"] = (data.get("_log", []) + [{
            "when": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "who": getpass.getuser(),
            "group": key, "from": f"{start[:4]}-{start[4:6]}-{start[6:]}",
            "before": before, "after": pct(factor) if factor is not None else "removed"}])[-500:]
        _parse_prepos(data, path)
        return data

    _locked_update(path, change)
    return load_prepositioning(path)


def factor_series(rep: Report, prepos: Prepositioning) -> pd.Series:
    df = rep.df
    ones = pd.Series(1.0, index=df.index)
    if not prepos.rules:
        return ones
    rep.require("ENTITY_GROUP", why="applying prepositioning")
    factors = {g: prepos.factor(g, rep.date) for g in df["ENTITY_GROUP"].cat.categories}
    if all(f == 1.0 for f in factors.values()):
        return ones
    f = df["ENTITY_GROUP"].map(factors).astype("float64")
    return f.where(df["BUCKET_CFP"].isin(rep.window_buckets()), 1.0)


def adjusted_amount(rep: Report, prepos: Prepositioning) -> pd.Series:
    return rep.df["AMOUNT"] * factor_series(rep, prepos)


def group_of(rep: Report, level: str) -> dict:
    if level == "ENTITY_GROUP" or "ENTITY_GROUP" not in rep.df.columns:
        return {}
    pairs = rep.df[["LEGAL_ENTITY", "ENTITY_GROUP"]].drop_duplicates().astype(str)
    dup = pairs[pairs.duplicated("LEGAL_ENTITY", keep=False)]
    if len(dup):
        le = dup["LEGAL_ENTITY"].iloc[0]
        raise DataError(f"{rep.path.name}: legal entity {le} appears under "
                        f"{', '.join(sorted(dup[dup.LEGAL_ENTITY == le].ENTITY_GROUP))} - fix the extract.")
    return dict(zip(pairs["LEGAL_ENTITY"], pairs["ENTITY_GROUP"]))


def requirement_grid(rep: Report, level: str, prepos: Prepositioning) -> pd.DataFrame:
    rep.require(level, why="binding by " + LABELS[level].lower())
    flows = rep.flow_buckets()
    if not flows:
        raise DataError(f"{rep.path.name} has no buckets between D1 and {HORIZON_END}.")
    amt = adjusted_amount(rep, prepos)
    keep = rep.df["BUCKET_CFP"].isin(flows)
    keys = [rep.df[level][keep], rep.df["SCENARIO"][keep], rep.df["BUCKET_CFP"][keep]]
    g = (amt[keep].groupby(keys, observed=True).sum()
         .unstack(-1, fill_value=0.0).reindex(columns=flows, fill_value=0.0))
    g.columns = [str(c) for c in g.columns]
    return g.cumsum(axis=1)


def binding(rep: Report, level: str, prepos: Prepositioning) -> dict:
    cum = requirement_grid(rep, level, prepos)
    groups = group_of(rep, level)
    out = {}
    for ent, sub in cum.groupby(level=0, observed=True):
        sub = sub.droplevel(0)
        worst = sub.min(axis=1).sort_values(kind="stable")
        where = sub.idxmin(axis=1)
        scen = worst.index[0]
        nxt = None
        if len(worst) > 1:
            s2 = worst.index[1]
            nxt = {"scen": str(s2), "bucket": str(where[s2]), "req": float(worst.iloc[1])}
        eg = groups.get(str(ent), str(ent))
        out[str(ent)] = {"scen": str(scen), "bucket": str(where[scen]), "req": float(worst.iloc[0]),
                         "group": eg, "factor": prepos.factor(eg, rep.date), "next": nxt}
    return out


def _read_json(path: Path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise DataError(f"{Path(path).name} is not valid JSON (line {e.lineno}: {e.msg}).") from None


def _locked_update(path: Path, change) -> None:
    path = Path(path)
    lock = path.with_name(path.name + ".lock")
    deadline = time.monotonic() + 10
    while True:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{getpass.getuser()} {datetime.now():%Y-%m-%d %H:%M:%S}".encode())
            os.close(fd)
            break
        except FileExistsError:
            try:
                if time.time() - lock.stat().st_mtime > 120:
                    lock.unlink()
                    continue
            except FileNotFoundError:
                continue
            if time.monotonic() > deadline:
                raise DataError(f"{path.name} is being saved by someone else - try again in a moment.") from None
            time.sleep(0.2)
    try:
        data = _read_json(path) if path.exists() else {}
        data = change(data)
        fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=f".{path.name}.", suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
                f.write("\n")
            os.replace(tmp, path)
        except BaseException:
            Path(tmp).unlink(missing_ok=True)
            raise
    finally:
        lock.unlink(missing_ok=True)


def load_settings(path) -> dict:
    path = Path(path)
    if not path.exists():
        return {}
    return {k: v for k, v in _read_json(path).items() if not k.startswith("_")}


def save_setting(path, key: str, value) -> None:
    def change(data):
        if value is None:
            data.pop(key, None)
        else:
            data[key] = value
        return data
    _locked_update(Path(path), change)


def money(v: float, units: str = "mn", decimals: int = 1, sign: bool = False) -> str:
    div = {"1": 1, "k": 1e3, "mn": 1e6, "bn": 1e9}[units]
    x = v / div
    if round(x, decimals) == 0:
        return "0"
    return f"{x:{'+' if sign else ''},.{decimals}f}"
