from __future__ import annotations

from dataclasses import dataclass, field, replace

import pandas as pd

from . import core

BREAKDOWNS = ["CONTINGENCY", "DATA_TYPE", "LEGAL_ENTITY", "CURRENCY", "BUCKET_CFP", "TIME_HORIZON"]
FILTER_FIELDS = ["CONTINGENCY", "DATA_TYPE", "CURRENCY", "LEGAL_ENTITY", "ENTITY_GROUP"]
SCOPES = {"cumulative": "Cumulative from D1", "bucket": "One bucket only", "days": "Range of days"}
DIMS = ["ENTITY_GROUP", "LEGAL_ENTITY", "SCENARIO", "CONTINGENCY", "DATA_TYPE", "CURRENCY", "BUCKET_CFP", "TIME_HORIZON"]
MAX_EXPORT_ROWS = 1_000_000


@dataclass(frozen=True)
class Side:
    date: str
    scenario: str
    horizon: str = "M2"
    level: str = "ENTITY_GROUP"
    entity: str = ""
    scope: str = "cumulative"
    days: tuple = (1, 30)
    filters: tuple = ()

    def span(self) -> str:
        if self.scope == "bucket":
            return f"{self.horizon} only"
        if self.scope == "days":
            return f"days {self.days[0]:g}–{self.days[1]:g}"
        return f"D1–{self.horizon} cumulative"

    def who(self) -> str:
        return self.entity or ("All entity groups" if self.level == "ENTITY_GROUP" else "All legal entities")

    def describe(self) -> str:
        parts = [core.fmt_date(self.date), self.who(), self.scenario, self.span()]
        return " · ".join(parts + [f"{core.LABELS[f]} = {v}" for f, v in self.filters])


@dataclass(frozen=True)
class View:
    sides: tuple
    breakdown: str = "CONTINGENCY"
    drill: tuple = ()
    prepositioned: bool = True

    @property
    def compare(self) -> bool:
        return len(self.sides) == 2

    def drilled(self, key) -> "View":
        drill = self.drill + ((self.breakdown, key),)
        used = {f for f, _ in drill} | {f for s in self.sides for f, _ in s.filters}
        nxt = next((b for b in BREAKDOWNS if b not in used), None)
        if nxt is None:
            raise ValueError("Nothing left to break down.")
        return replace(self, drill=drill, breakdown=nxt)


@dataclass
class Result:
    view: View
    columns: list
    rows: list
    totals: list
    context: list
    sources: list = field(default_factory=list)


def _match(col: pd.Series, fld: str, value) -> pd.Series:
    return (col == float(value)) if fld == "TIME_HORIZON" else (col.astype(str) == str(value))


def select(side: Side, drill: tuple, rep: core.Report) -> pd.Series:
    df = rep.df
    rep.require("SCENARIO", "BUCKET_CFP", side.level if side.entity else "",
                "TIME_HORIZON" if side.scope == "days" else "",
                *[f for f, _ in side.filters + drill], why="this selection")
    if side.scenario not in df["SCENARIO"].cat.categories:
        raise core.DataError(f"Scenario '{side.scenario}' is not in {rep.path.name}.")
    mask = df["SCENARIO"] == side.scenario
    if side.scope == "cumulative":
        mask &= df["BUCKET_CFP"].isin(rep.flow_buckets(side.horizon))
    elif side.scope == "bucket":
        if side.horizon not in rep.bucket_days:
            raise core.DataError(f"Bucket {side.horizon} is not in {rep.path.name}.")
        mask &= df["BUCKET_CFP"] == side.horizon
    else:
        lo, hi = side.days
        mask &= df["TIME_HORIZON"].between(lo, hi)
    if side.entity:
        if side.entity not in df[side.level].cat.categories:
            raise core.DataError(f"{core.LABELS[side.level]} '{side.entity}' is not in {rep.path.name}.")
        mask &= df[side.level] == side.entity
    for fld, value in side.filters + drill:
        mask &= _match(df[fld], fld, value)
    return mask


def _factors_used(rep, mask, factors) -> list:
    if "ENTITY_GROUP" not in rep.df.columns:
        return []
    window = mask & rep.df["BUCKET_CFP"].isin(rep.window_buckets()) & (factors != 1.0)
    if not window.any():
        return []
    pairs = pd.DataFrame({"g": rep.df["ENTITY_GROUP"][window].astype(str), "f": factors[window]}).drop_duplicates()
    return sorted(zip(pairs["g"], pairs["f"]))


def compute(view: View, report, prepos: core.Prepositioning) -> Result:
    series, sources, notes, relevant = [], [], [], False
    for side in view.sides:
        rep = report(side.date)
        rep.require(view.breakdown, why="this breakdown")
        mask = select(side, view.drill, rep)
        if view.prepositioned or "ENTITY_GROUP" in rep.df.columns:
            factors = core.factor_series(rep, prepos)
        else:
            factors = pd.Series(1.0, index=rep.df.index)
        used = _factors_used(rep, mask, factors)
        relevant = relevant or bool(used)
        amount = rep.df["AMOUNT"] * factors if view.prepositioned else rep.df["AMOUNT"]
        key = rep.df[view.breakdown][mask]
        key = key.map(lambda d: "?" if pd.isna(d) else f"{d:g}") if view.breakdown == "TIME_HORIZON" else key.astype(str)
        series.append(amount[mask].groupby(key.to_numpy()).sum())
        sources.append(rep.path.name)
        notes.append(", ".join(f"{g} {core.pct(f)}" for g, f in used) or "none")

    frame = pd.concat(series, axis=1).fillna(0.0)
    if view.compare:
        frame["change"] = frame.iloc[:, 1] - frame.iloc[:, 0]
        columns, order = ["A", "B", "Change"], frame["change"].abs()
    else:
        columns, order = ["Amount"], frame.iloc[:, 0].abs()
    if view.breakdown == "BUCKET_CFP":
        days = report(view.sides[-1].date).bucket_days
        frame = frame.loc[sorted(frame.index, key=lambda k: days.get(k, 1e9))]
    elif view.breakdown == "TIME_HORIZON":
        frame = frame.loc[sorted(frame.index, key=lambda k: float(k) if k != "?" else 1e9)]
    else:
        frame = frame.loc[order.sort_values(ascending=False, kind="stable").index]

    if view.compare:
        lines = [f"A  {view.sides[0].describe()}", f"B  {view.sides[1].describe()}"]
    else:
        lines = [view.sides[0].describe()]
    path = " › ".join(f"{core.LABELS[f]} = {v}" for f, v in view.drill)
    lines.append(f"By {core.LABELS[view.breakdown].lower()}" + (f" · within {path}" if path else ""))
    if not view.prepositioned:
        lines.append("Prepositioning not applied (raw flows)")
    elif relevant:
        lines.append("Prepositioning applied to M2–M12 flows: "
                     + (f"A {notes[0]} · B {notes[1]}" if view.compare else notes[0]))
    rows = [(k, [float(v) for v in r]) for k, r in zip(frame.index, frame.to_numpy())]
    return Result(view, columns, rows, [float(v) for v in frame.sum().to_numpy()], lines, sources)


def underlying(view: View, report, prepos: core.Prepositioning) -> pd.DataFrame:
    parts = []
    for label, side in zip("AB", view.sides):
        rep = report(side.date)
        mask = select(side, view.drill, rep)
        factors = core.factor_series(rep, prepos) if view.prepositioned else pd.Series(1.0, index=rep.df.index)
        dims = [d for d in DIMS if d in rep.df.columns]
        rows = rep.df.loc[mask, dims].copy()
        for d in dims:
            if d != "TIME_HORIZON":
                rows[d] = rows[d].astype(str)
        rows["Factor"] = factors[mask]
        rows["Raw USD"] = rep.df["AMOUNT"][mask]
        rows["USD"] = rows["Raw USD"] * rows["Factor"]
        agg = rows.groupby(dims + ["Factor"], dropna=False, sort=False)[["Raw USD", "USD"]].sum().reset_index()
        agg.insert(0, "Report", pd.Timestamp(side.date).date())
        if view.compare:
            agg.insert(0, "Side", label)
        parts.append(agg)
    out = pd.concat(parts, ignore_index=True).rename(columns=core.LABELS)
    if len(out) > MAX_EXPORT_ROWS:
        raise core.DataError(f"{len(out):,} rows is more than Excel can hold - narrow the selection first.")
    return out
