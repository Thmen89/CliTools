from __future__ import annotations

import json
import textwrap
from collections import OrderedDict
from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pandas as pd

from . import analysis as an
from . import core
from . import excel as xl
from .ui import actions, ask, choose, confirm, label_value, numbered, pair, say, show_result, title, width

DEFAULTS = {"units": "mn", "decimals": 1, "level": "ENTITY_GROUP", "history_reports": 10,
            "rows_shown": 15, "default_horizon": "M2", "export_dir": "", "column_aliases": {},
            "file_pattern": "GetILSTAggT2_{date}.csv"}
SETTING_KEYS = ("units", "decimals", "level", "history_reports", "rows_shown", "default_horizon", "export_dir")
BACK = ("b", "back", "q", "quit", "..")


def run(ctx, screen: str = "binding") -> None:
    app = App(ctx)
    {"binding": app.binding, "analysis": app.analysis, "settings": app.settings}[screen]()


class App:
    def __init__(self, ctx):
        self.ui, self.base = ctx.ui, dict(ctx.config)
        self.app_dir = Path(getattr(ctx, "app_dir", None) or Path.cwd())
        self.settings_file = self.path(self.base.get("settings_file", "config/ilst.json"))
        self.prepos_file = self.path(self.base.get("prepositioning_file", "config/ilst_prepositioning.json"))
        self.banner = self.base.get("banner", "")
        self.msg = None
        self.dates: list = []
        self.paths: dict = {}
        self.summaries: OrderedDict = OrderedDict()
        self.reload()

    def path(self, p) -> Path:
        p = Path(str(p)).expanduser()
        return p if p.is_absolute() else self.app_dir / p

    def reload(self):
        saved = {k: v for k, v in core.load_settings(self.settings_file).items() if k in SETTING_KEYS}
        self.cfg = {**DEFAULTS, "rows_shown": self.base.get("top_n", DEFAULTS["rows_shown"]), **self.base, **saved}

    def title(self, text, *context, warn=""):
        title(self.ui, text, *context, warn=warn, banner=self.banner)

    def bind(self, date, level, pp):
        key = (core.fingerprint(self.paths[date]), date, level, json.dumps(self.cfg.get("column_aliases") or {}),
               tuple(sorted((k, tuple(v)) for k, v in pp.rules.items())))
        if key not in self.summaries:
            self.summaries[key] = core.binding(self.report(date), level, pp)
            while len(self.summaries) > 64:
                self.summaries.popitem(last=False)
        self.summaries.move_to_end(key)
        return self.summaries[key]

    def m(self, v, sign=False):
        return core.money(v, self.cfg["units"], int(self.cfg["decimals"]), sign)

    def open_reports(self) -> bool:
        try:
            if "data_dir" not in self.cfg:
                raise core.DataError("No data folder configured (data_dir in config/app.json).")
            self.paths = core.list_reports(self.path(self.cfg["data_dir"]), self.cfg["file_pattern"])
            self.dates = list(self.paths)
            return True
        except core.DataError as e:
            self.title("ILST", warn=str(e))
            return False

    def report(self, date):
        with self.ui.status(f"Reading {self.paths[date].name}…"):
            return core.load_report(date, self.paths[date], self.cfg.get("column_aliases"))

    def prepos(self):
        return core.load_prepositioning(self.prepos_file)

    def previous(self, date):
        i = self.dates.index(date)
        return self.dates[i - 1] if i else None

    def flash(self):
        if self.msg:
            say(self.ui)
            say(self.ui, *self.msg)
            self.msg = None

    @staticmethod
    def describe_error(e) -> str:
        if isinstance(e, OSError):
            return f"{e.strerror or e}: {e.filename}" if getattr(e, "filename", None) else str(e)
        return str(e)

    def loop(self, render, handle, enter=None) -> None:
        while True:
            ok = True
            try:
                render()
            except (core.DataError, ValueError, OSError) as e:
                ok = False
                self.title("ILST", warn=self.describe_error(e))
                actions(self.ui, "s: settings", "b: back")
            self.flash()
            cmd = ask()
            if cmd is None or cmd.lower() in BACK:
                return
            if not cmd:
                if ok and enter:
                    try:
                        enter()
                    except (core.DataError, ValueError, OSError) as e:
                        self.msg = (self.describe_error(e), "yellow")
                continue
            if not ok:
                if cmd.lower() == "s":
                    self.settings()
                else:
                    self.msg = ("Fix the problem above first - s: settings, b: back.", "yellow")
                continue
            try:
                if handle(cmd.lower().strip()) == "back":
                    return
            except (core.DataError, ValueError, OSError) as e:
                self.msg = (self.describe_error(e), "yellow")

    def export(self, sheets, name) -> None:
        if not xl.available():
            raise core.DataError("Excel export needs openpyxl:  pip install openpyxl")
        folder = xl.folder(self.path(self.cfg["export_dir"]) if self.cfg["export_dir"] else None)
        rows = sum(len(s.frame) for s in sheets)
        with self.ui.status(f"Writing {rows:,} rows to Excel…"):
            path = xl.write(folder / name, sheets, self.cfg["units"], int(self.cfg["decimals"]))
        self.msg = (f"Saved {path}", "green")

    def binding(self):
        if not self.open_reports():
            return
        st = {"date": self.dates[-1], "level": self.cfg["level"], "names": []}

        def render():
            date, level = st["date"], st["level"]
            prev = self.previous(date)
            pp = self.prepos()
            cur = self.bind(date, level, pp)
            old = self.bind(prev, level, pp) if prev else {}
            st["cur"], st["old"], st["prev"], st["names"] = cur, old, prev, sorted(cur)
            groups = sorted({b["group"] for b in cur.values()})
            active = pp.active(groups, date)
            self.title(f"ILST Binding · {core.fmt_date(date)}",
                  f"{core.LABELS[level]} · USD {self.cfg['units']}"
                  + (f" · change vs {core.fmt_date(prev)}" if prev else ""),
                  f"Requirement: cumulative D1 to the binding bucket (up to {core.HORIZON_END}), Day 0 excluded",
                  warn=("Prepositioning on M2–M12 flows: " + ", ".join(f"{g} {core.pct(f)}" for g, f in active))
                  if active else "")
            for line in pair("     Scenario · bucket", f"{'Requirement':>12}  {'Change':>10}"):
                say(self.ui, line, "dim")
            for n, e in enumerate(st["names"], 1):
                b, p = cur[e], old.get(e)
                for line in numbered(n, e):
                    say(self.ui, line, "bold")
                change = self.m(b["req"] - p["req"], True) if p else "new"
                for line in pair(f"{b['scen']} · {b['bucket']}", f"{self.m(b['req']):>12}  {change:>10}", indent=5):
                    say(self.ui, line)
                if p and (p["scen"], p["bucket"]) != (b["scen"], b["bucket"]):
                    say(self.ui, f"     moved - was {p['scen']} · {p['bucket']}", "yellow")
            other = "legal entity" if level == "ENTITY_GROUP" else "entity group"
            actions(self.ui, "Number: detail", "d: date", f"l: by {other}", "x: export", "s: settings", "q: quit")

        def handle(cmd):
            if cmd.isdigit():
                n = int(cmd)
                if not 1 <= n <= len(st["names"]):
                    raise ValueError("Pick a number from the list.")
                self.entity(st["names"][n - 1], st["date"], st["level"])
            elif cmd == "d":
                st["date"] = self.ask_date("Report date", st["date"])
            elif cmd == "l":
                st["level"] = "LEGAL_ENTITY" if st["level"] == "ENTITY_GROUP" else "ENTITY_GROUP"
            elif cmd == "s":
                self.settings()
                st["level"] = self.cfg["level"]
            elif cmd == "x":
                tag = "eg" if st["level"] == "ENTITY_GROUP" else "le"
                self.export([self.binding_sheet(st)], f"binding_{st['date']}_{tag}_{datetime.now():%H%M%S}.xlsx")
            else:
                raise ValueError("Type a number, d, l, x, s or q.")

        self.loop(render, handle)

    def binding_sheet(self, st) -> xl.Sheet:
        rows = []
        for e in st["names"]:
            b, p, nxt = st["cur"][e], st["old"].get(e), st["cur"][e]["next"] or {}
            rows.append({core.LABELS[st["level"]]: e, "Scenario": b["scen"], "Bucket": b["bucket"],
                         "Requirement": b["req"], "Previous scenario": p["scen"] if p else None,
                         "Previous bucket": p["bucket"] if p else None, "Previous requirement": p["req"] if p else None,
                         "Change": b["req"] - p["req"] if p else None,
                         "Next tightest": f"{nxt['scen']} · {nxt['bucket']}" if nxt else None,
                         "Gap to next": nxt["req"] - b["req"] if nxt else None, "Prepositioning": b["factor"]})
        prev = st["prev"]
        return xl.Sheet(f"Binding {core.fmt_date(st['date'])}", pd.DataFrame(rows),
                        {c: "money" for c in ("Requirement", "Previous requirement", "Change", "Gap to next")}
                        | {"Prepositioning": "pct"},
                        [f"{core.LABELS[st['level']]} · cumulative D1 to binding bucket (≤ {core.HORIZON_END}), "
                         "Day 0 excluded · prepositioning applied to M2–M12 flows"
                         + (f" · previous = {core.fmt_date(prev)}" if prev else "")],
                        self.paths[st["date"]].name, tab="Binding")

    def ask_date(self, heading, current):
        recent = list(reversed(self.dates[-8:]))
        say(self.ui)
        say(self.ui, heading, "bold")
        for line in textwrap.wrap("   ".join(f"{i}:\u00a0{core.fmt_date(d)[:6]}" for i, d in enumerate(recent, 1)), width()):
            say(self.ui, line, "dim")
        raw = ask("Number, date or -1/+1", core.fmt_date(current))
        if raw is None or raw == core.fmt_date(current):
            return current
        if raw.isdigit() and 1 <= int(raw) <= len(recent):
            return recent[int(raw) - 1]
        if raw[:1] in "+-" and raw[1:].isdigit():
            i = self.dates.index(current) + int(raw)
            if not 0 <= i < len(self.dates):
                raise ValueError(f"No report {raw} from {core.fmt_date(current)}.")
            return self.dates[i]
        date = core.parse_date(raw)
        if date not in self.paths:
            earlier = [d for d in self.dates if d < date]
            raise ValueError(f"No report for {core.fmt_date(date)}."
                             + (f" Nearest earlier: {core.fmt_date(earlier[-1])}." if earlier else ""))
        return date

    def pick(self, heading, options, current=None):
        options = [str(o) for o in options]
        say(self.ui)
        say(self.ui, heading, "bold")
        if len(options) <= 12:
            for i, o in enumerate(options, 1):
                for line in numbered(i, o):
                    say(self.ui, line)
        elif max(len(o) for o in options) <= 8:
            for line in textwrap.wrap("  ".join(options), width()):
                say(self.ui, line, "dim")
        else:
            say(self.ui, f"{len(options)} to choose from - type part of the name.", "dim")
        raw = ask("Number or name", current or "")
        if raw is None or (current and raw == current):
            return current
        if raw.isdigit() and 1 <= int(raw) <= len(options) and len(options) <= 12:
            return options[int(raw) - 1]
        low = raw.lower()
        hits = ([o for o in options if o.lower() == low]
                or [o for o in options if o.lower().startswith(low)]
                or [o for o in options if any(w.startswith(low) for w in o.lower().split())]
                or [o for o in options if low in o.lower()])
        if len(hits) == 1:
            return hits[0]
        if not hits:
            raise ValueError(f"Nothing matches '{raw}'.")
        i = choose(self.ui, f"{len(hits)} match '{raw}'", hits[:30])
        return current if i is None else hits[i]

    def binding_view(self, ent, level, a_date, a, b_date, b):
        return an.View((an.Side(a_date, a["scen"], a["bucket"], level, ent),
                        an.Side(b_date, b["scen"], b["bucket"], level, ent)))

    def entity(self, ent, date, level):
        prev = self.previous(date)

        def render():
            pp = self.prepos()
            b = self.bind(date, level, pp).get(ent)
            if b is None:
                raise core.DataError(f"{ent} is not in the {core.fmt_date(date)} report.")
            p = self.bind(prev, level, pp).get(ent) if prev else None
            self.cur_b, self.prev_b = b, p
            self.title(f"{ent} · {core.fmt_date(date)}", f"{core.LABELS[level]} · USD {self.cfg['units']}")
            rows = [("Binding", f"{b['scen']} · {b['bucket']}"), ("Requirement", self.m(b["req"]))]
            if p:
                rows += [(f"On {core.fmt_date(prev)}", f"{p['scen']} · {p['bucket']}  {self.m(p['req'])}"),
                         ("Change", self.m(b["req"] - p["req"], True))]
            nxt = b["next"]
            if nxt:
                rows.append(("Next tightest", f"{nxt['scen']} · {nxt['bucket']}  "
                                              f"({self.m(nxt['req'] - b['req'])} less)"))
            rows.append(("Prepositioning", f"{core.pct(b['factor'])} on M2–M12 flows ({b['group']})"))
            for label, value in rows:
                for line in label_value(label, value):
                    say(self.ui, line)
            say(self.ui)
            if p:
                say(self.ui, f"  1  What moved it since {core.fmt_date(prev)}")
            say(self.ui, "  2  Compare with another date")
            say(self.ui, "  3  Binding history")
            actions(self.ui, "1, 2 or 3", "b: back")

        def handle(cmd):
            if cmd == "1" and self.prev_b:
                self.show_view(self.binding_view(ent, level, prev, self.prev_b, date, self.cur_b),
                               f"What moved {ent}")
            elif cmd == "2":
                other = self.ask_date("Compare with which report?", prev or date)
                o = self.bind(other, level, self.prepos()).get(ent)
                if o is None:
                    raise ValueError(f"{ent} is not in the {core.fmt_date(other)} report.")
                a_date, a, b_date, b = (other, o, date, self.cur_b) if other <= date else (date, self.cur_b, other, o)
                self.show_view(self.binding_view(ent, level, a_date, a, b_date, b),
                               f"{ent} · {core.fmt_date(a_date)} vs {core.fmt_date(b_date)}")
            elif cmd == "3":
                self.history(ent, level, date)
            else:
                raise ValueError("Type 1, 2, 3 or b.")

        self.loop(render, handle)

    def show_view(self, view, heading):
        note = "Binding to binding: the change includes any switch of scenario, bucket or prepositioning factor."
        if self.result(view, heading, note=note) == "edit":
            self.editor(view)

    def history(self, ent, level, end):
        st = {"start": None, "end": end, "points": []}

        def render():
            dates = [d for d in self.dates if (st["start"] or "") <= d <= st["end"]]
            if st["start"] is None:
                dates = dates[-int(self.cfg["history_reports"]):]
            if not dates:
                raise ValueError("No reports in that range.")
            pp = self.prepos()
            read = ([self.previous(dates[0])] if self.previous(dates[0]) else []) + dates
            hist = {d: self.bind(d, level, pp).get(ent) for d in read}
            st["points"] = [(d, hist[d], hist.get(self.previous(d))) for d in reversed(dates) if hist[d]]
            self.title(f"{ent} · binding history",
                       f"{core.fmt_date(dates[0])} to {core.fmt_date(dates[-1])} · USD {self.cfg['units']} · "
                       "* = binding scenario or bucket moved")
            if not st["points"]:
                say(self.ui, f"  {ent} is not in any report in this range.", "dim")
            else:
                for line in pair("     Report · scenario · bucket", f"{'Requirement':>12}  {'Change':>10}"):
                    say(self.ui, line, "dim")
            for n, (d, b, p) in enumerate(st["points"], 1):
                moved = bool(p) and (p["scen"], p["bucket"]) != (b["scen"], b["bucket"])
                left = f"{n:>3}  {core.fmt_date(d)[:6]}  {b['scen']} · {b['bucket']}{' *' if moved else ''}"
                change = self.m(b["req"] - p["req"], True) if p else ""
                for line in pair(left, f"{self.m(b['req']):>12}  {change:>10}"):
                    say(self.ui, line, "yellow" if moved else "")
            actions(self.ui, "Number: what moved it that day", "r: date range", "x: export with chart", "b: back")

        def handle(cmd):
            if cmd.isdigit():
                n = int(cmd)
                if not 1 <= n <= len(st["points"]):
                    raise ValueError("Pick a number from the list.")
                d, b, p = st["points"][n - 1]
                if not p:
                    raise ValueError("No earlier report to compare with.")
                self.show_view(self.binding_view(ent, level, self.previous(d), p, d, b),
                               f"What moved {ent} on {core.fmt_date(d)}")
            elif cmd == "r":
                first = st["points"][-1][0] if st["points"] else st["end"]
                a = ask("From (YYYY-MM-DD)", f"{first[:4]}-{first[4:6]}-{first[6:]}")
                z = ask("To (YYYY-MM-DD)", f"{st['end'][:4]}-{st['end'][4:6]}-{st['end'][6:]}")
                if a is None or z is None:
                    return
                a, z = sorted([core.parse_date(a), core.parse_date(z)])
                st["start"], st["end"] = a, z
            elif cmd == "x":
                if not st["points"]:
                    raise ValueError("Nothing to export in this range.")
                pts = list(reversed(st["points"]))
                frame = pd.DataFrame([{"Report": pd.Timestamp(d).date(), "Scenario": b["scen"], "Bucket": b["bucket"],
                                       "Requirement": b["req"], "Change": b["req"] - p["req"] if p else None,
                                       "Prepositioning": b["factor"]} for d, b, p in pts])
                flag = [i for i, (_, b, p) in enumerate(pts) if p and (p["scen"], p["bucket"]) != (b["scen"], b["bucket"])]
                sheet = xl.Sheet(f"{ent} binding history", frame,
                                 {"Report": "date", "Requirement": "money", "Change": "money", "Prepositioning": "pct"},
                                 [f"{core.LABELS[level]} {ent} · cumulative D1 to binding bucket (≤ {core.HORIZON_END}), "
                                  "Day 0 excluded · line colour = binding scenario · highlighted = binding moved",
                                  "Prepositioning column = factor in force on each report date"],
                                 f"{len(pts)} reports to {self.paths[pts[-1][0]].name}", flag=flag,
                                 chart={"x": "Report", "y": "Requirement", "group": "Scenario",
                                        "title": f"{ent} · binding requirement"}, tab="History")
                safe = "".join(c if c.isalnum() else "_" for c in ent)
                self.export([sheet], f"binding_history_{safe}_{pts[0][0]}_{pts[-1][0]}.xlsx")
            else:
                raise ValueError("Type a number, r, x or b.")

        self.loop(render, handle)

    def result(self, view, heading, back="back", note=""):
        pp = self.prepos()
        st = {"stack": [an.compute(view, self.report, pp)], "all": False, "edit": False}

        def render():
            res = st["stack"][-1]
            units = f"USD {self.cfg['units']} · negative = requirement" + (" · Change = B − A" if res.view.compare else "")
            self.title(heading, *res.context, units, *([note] if note else []))
            show_result(self.ui, res, self.cfg["units"], int(self.cfg["decimals"]),
                        None if st["all"] else int(self.cfg["rows_shown"]))
            acts = ["Number: break that row down", "c: change breakdown"]
            if len(st["stack"]) > 1:
                acts.append("u: up a level")
            actions(self.ui, *acts, "a: all rows", "e: edit selection", "x: export", "d: export with data", f"b: {back}")

        def handle(cmd):
            res = st["stack"][-1]
            if cmd.isdigit():
                n = int(cmd)
                if not 1 <= n <= len(res.rows):
                    raise ValueError("Pick a row number.")
                st["stack"].append(an.compute(res.view.drilled(res.rows[n - 1][0]), self.report, pp))
                st["all"] = False
            elif cmd == "c":
                cols = self.report(res.view.sides[-1].date).df.columns
                used = {f for f, _ in res.view.drill} | {f for s in res.view.sides for f, _ in s.filters}
                options = [b for b in an.BREAKDOWNS if b in cols and b not in used]
                i = choose(self.ui, "Break down by", [core.LABELS[b] for b in options],
                           options.index(res.view.breakdown) if res.view.breakdown in options else 0)
                if i is not None:
                    st["stack"][-1] = an.compute(replace(res.view, breakdown=options[i]), self.report, pp)
            elif cmd == "u" and len(st["stack"]) > 1:
                st["stack"].pop()
            elif cmd == "a":
                st["all"] = not st["all"]
            elif cmd == "e":
                st["edit"] = True
                return "back"
            elif cmd in ("x", "d"):
                sheets = [xl.result_sheet(res, heading, self.cfg["units"])]
                if cmd == "d":
                    with self.ui.status("Collecting the underlying rows…"):
                        sheets.append(xl.data_sheet(an.underlying(res.view, self.report, pp), res))
                stamp = res.view.sides[-1].date
                self.export(sheets, f"ilst_{stamp}_{datetime.now():%H%M%S}.xlsx")
            else:
                raise ValueError("Type a row number, c, u, a, e, x, d or b.")

        self.loop(render, handle)
        return "edit" if st["edit"] else None

    def analysis(self):
        if not self.open_reports():
            return
        try:
            view = self.default_view()
        except (core.DataError, ValueError, OSError) as e:
            self.title("ILST Analysis", warn=self.describe_error(e))
            ask("Enter to go back")
            return
        self.editor(view)

    def default_view(self):
        date, level = self.dates[-1], self.cfg["level"]
        rep = self.report(date)
        rep.require(level, "SCENARIO", why="analysis")
        scen = sorted(rep.df["SCENARIO"].cat.categories)[0]
        flows = rep.flow_buckets()
        horizon = self.cfg["default_horizon"] if self.cfg["default_horizon"] in rep.bucket_days else flows[-1]
        return an.View((an.Side(date, scen, horizon, level),))

    FIELDS = ["Date", "Level", "Entity", "Scenario", "Scope", "Filters"]

    def editor(self, view):
        st = {"view": replace(view, drill=(), breakdown="CONTINGENCY")}

        def render():
            v = st["view"]
            self.title("ILST Analysis", "Choose what to look at: change a field by its number, Enter shows the numbers.",
                       f"USD {self.cfg['units']} · negative = requirement")
            for k, side in enumerate(v.sides):
                if v.compare:
                    say(self.ui, "A" if k == 0 else "B", "bold")
                values = [core.fmt_date(side.date), core.LABELS[side.level], side.who(), side.scenario, side.span(),
                          ", ".join(f"{core.LABELS[f]} = {x}" for f, x in side.filters) or "none"]
                for i, (label, value) in enumerate(zip(self.FIELDS, values), k * len(self.FIELDS) + 1):
                    for line in label_value(f"{i:>3}  {label}", value, 17):
                        say(self.ui, line)
            say(self.ui)
            basis = "applied (binding basis)" if v.prepositioned else "not applied (raw flows)"
            for line in label_value("  p  Prepositioning", basis, 22):
                say(self.ui, line)
            compare = "c: remove comparison" if v.compare else "c: compare with…"
            actions(self.ui, "Enter: show", "Number: change", compare, "p: prepositioning", "b: back")

        def show():
            if self.result(st["view"], "ILST Analysis", back="back to selection") == "edit":
                return

        def handle(cmd):
            v = st["view"]
            if cmd == "p":
                st["view"] = replace(v, prepositioned=not v.prepositioned)
            elif cmd == "c":
                if v.compare:
                    st["view"] = replace(v, sides=v.sides[:1])
                else:
                    a = v.sides[0]
                    prev = self.previous(a.date)
                    options = ["A with the previous report" + (f" ({core.fmt_date(prev)})" if prev else ""),
                               "A copy of A - then change what differs"]
                    i = choose(self.ui, "Compare A with", options, 0 if prev else 1)
                    if i is None:
                        return None
                    if i == 0 and not prev:
                        raise ValueError("There is no earlier report.")
                    b = a
                    if i == 0:
                        a, b = replace(a, date=prev), a
                    st["view"] = replace(v, sides=(a, b))
            elif cmd.isdigit() and 1 <= int(cmd) <= len(v.sides) * len(self.FIELDS):
                k, f = divmod(int(cmd) - 1, len(self.FIELDS))
                sides = list(v.sides)
                sides[k] = self.edit_side(sides[k], self.FIELDS[f])
                st["view"] = replace(v, sides=tuple(sides))
            else:
                raise ValueError("Type a field number, Enter, c, p or b.")
            return None

        self.loop(render, handle, enter=show)

    def edit_side(self, side, fld):
        rep = self.report(side.date)
        if fld == "Date":
            date = self.ask_date("Report date", side.date)
            return replace(side, date=date)
        if fld == "Level":
            level = "LEGAL_ENTITY" if side.level == "ENTITY_GROUP" else "ENTITY_GROUP"
            rep.require(level, why="this level")
            return replace(side, level=level, entity="")
        if fld == "Entity":
            names = ["All"] + sorted(rep.df[side.level].cat.categories)
            name = self.pick(f"{core.LABELS[side.level]} - All, or one by number or name", names, side.entity or "All")
            return replace(side, entity="" if name in (None, "All") else name)
        if fld == "Scenario":
            return replace(side, scenario=self.pick("Scenario", sorted(rep.df["SCENARIO"].cat.categories), side.scenario))
        if fld == "Scope":
            i = choose(self.ui, "Scope", list(an.SCOPES.values()), list(an.SCOPES).index(side.scope))
            if i is None:
                return side
            scope = list(an.SCOPES)[i]
            if scope == "days":
                rep.require("TIME_HORIZON", why="a range of days")
                lo = ask("From day", f"{side.days[0]:g}")
                hi = ask("To day", f"{side.days[1]:g}")
                if lo is None or hi is None:
                    return side
                try:
                    lo, hi = sorted([float(lo), float(hi)])
                except ValueError:
                    raise ValueError("Days are numbers, e.g. 29 and 31.") from None
                return replace(side, scope=scope, days=(lo, hi))
            buckets = rep.flow_buckets(list(rep.bucket_days)[-1]) if scope == "cumulative" else list(rep.bucket_days)
            heading = "Cumulative from D1 up to" if scope == "cumulative" else "Which bucket"
            bucket = self.pick(heading, buckets, side.horizon if side.horizon in buckets else None)
            return replace(side, scope=scope, horizon=bucket or side.horizon)
        filters = list(side.filters)
        options = [f"Remove {core.LABELS[f]} = {x}" for f, x in filters] + ["Add a filter"]
        i = choose(self.ui, "Filters", options, len(options) - 1)
        if i is None:
            return side
        if i < len(filters):
            filters.pop(i)
            return replace(side, filters=tuple(filters))
        used = {f for f, _ in filters} | ({side.level} if side.entity else set())
        fields = [f for f in an.FILTER_FIELDS if f in rep.df.columns and f not in used]
        j = choose(self.ui, "Filter on", [core.LABELS[f] for f in fields])
        if j is None:
            return side
        fld = fields[j]
        value = self.pick(core.LABELS[fld], sorted(rep.df[fld].cat.categories))
        if value is None:
            return side
        return replace(side, filters=tuple(filters + [(fld, value)]))

    SETTINGS = [
        ("units", "Units", "choice", ["mn", "bn", "k", "1"], "How amounts are shown"),
        ("decimals", "Decimals", "int", (0, 3), "Decimal places"),
        ("level", "Default level", "choice", ["ENTITY_GROUP", "LEGAL_ENTITY"], "Entity group or legal entity"),
        ("history_reports", "History length", "int", (2, 250), "Reports shown in binding history"),
        ("rows_shown", "Rows shown", "int", (5, 500), "Rows in a breakdown before 'more'"),
        ("default_horizon", "Analysis horizon", "text", None, "Suggested horizon in Analysis, e.g. M2"),
        ("export_dir", "Export folder", "path", None, "Where exports are saved (blank = Downloads)"),
    ]

    def settings(self):
        def render():
            self.reload()
            self.title("ILST Settings", f"Saved to {self.settings_file}",
                  f"Data folder: {self.cfg.get('data_dir', '(not set)')}")
            for n, (key, label, _k, _x, help_) in enumerate(self.SETTINGS, 1):
                value = self.cfg.get(key)
                shown = {"ENTITY_GROUP": "Entity group", "LEGAL_ENTITY": "Legal entity"}.get(value, value)
                for line in pair(f"{n:>3}  {label}", str(shown or "Downloads" if key == "export_dir" else shown)):
                    say(self.ui, line)
            say(self.ui)
            say(self.ui, "  p  Prepositioning factors")
            actions(self.ui, "Number: change", "r <n>: reset to default", "p: prepositioning", "b: back")

        def handle(cmd):
            if cmd == "p":
                return self.prepositioning()
            if cmd.startswith("r ") and cmd[2:].strip().isdigit():
                key, label = self.setting(int(cmd[2:]))[:2]
                core.save_setting(self.settings_file, key, None)
                self.msg = (f"{label} reset to default.", "green")
                return None
            if not cmd.isdigit():
                raise ValueError("Type a number, r <n>, p or b.")
            key, label, kind, extra, help_ = self.setting(int(cmd))
            if kind == "choice":
                i = choose(self.ui, label, extra, extra.index(self.cfg[key]) if self.cfg[key] in extra else 0)
                if i is None:
                    return None
                value = extra[i]
            else:
                say(self.ui, help_, "dim")
                raw = ask(label, str(self.cfg.get(key) or ""))
                if raw is None:
                    return None
                value = self.parse_setting(kind, raw, extra)
            core.save_setting(self.settings_file, key, value)
            self.reload()
            self.msg = (f"{label} saved.", "green")
            return None

        self.loop(render, handle)
        self.reload()

    def setting(self, n):
        if not 1 <= n <= len(self.SETTINGS):
            raise ValueError("Pick a setting number.")
        return self.SETTINGS[n - 1]

    @staticmethod
    def parse_setting(kind, raw, extra):
        if kind == "int":
            if not raw.strip().isdigit() or not extra[0] <= int(raw) <= extra[1]:
                raise ValueError(f"Enter a whole number from {extra[0]} to {extra[1]}.")
            return int(raw)
        if kind == "text":
            if core.label_days(raw) is None:
                raise ValueError(f"'{raw}' is not a bucket like D1, Wk2 or M3.")
            return raw.strip()
        if kind == "path":
            if raw and not Path(raw).expanduser().is_dir():
                raise ValueError(f"Folder not found: {raw}")
            return raw
        return raw

    def prepositioning(self):
        groups_in_data: list = []
        as_of = datetime.now().strftime("%Y%m%d")
        if self.paths or ("data_dir" in self.cfg and self._try_reports()):
            as_of = self.dates[-1]
            try:
                rep = self.report(as_of)
                if "ENTITY_GROUP" in rep.df.columns:
                    groups_in_data = list(rep.df["ENTITY_GROUP"].cat.categories)
            except core.DataError:
                pass
        st = {"rows": [], "pp": None}

        def render():
            pp = self.prepos()
            st["pp"] = pp
            names = {g.upper(): g for g in groups_in_data}
            for k, v in pp.names.items():
                names.setdefault(k, v)
            st["rows"] = sorted(names.values(), key=str.upper)
            self.title("Prepositioning",
                  "Scales an entity group's flows in M2–M12 when finding binding. D1–M1 always count in full. "
                  "Groups not listed are at 100%.", f"Factors as at {core.fmt_date(as_of)} · {self.prepos_file}")
            last = self._last_change()
            if last:
                say(self.ui, f"Last change: {last}", "dim")
                say(self.ui)
            for n, g in enumerate(st["rows"], 1):
                steps = pp.rules.get(g.upper(), [])
                since = max((s for s, _ in steps if s <= as_of), default=None)
                now = f"{core.pct(pp.factor(g, as_of))}" + (f" since {core.fmt_date(since)}" if since else "")
                for line in pair(f"{n:>3}  {g}", now):
                    say(self.ui, line, "yellow" if pp.factor(g, as_of) != 1 else "")
                for s, v in steps:
                    if s > as_of:
                        say(self.ui, f"     then {core.pct(v)} from {core.fmt_date(s)}", "dim")
                if g.upper() not in {x.upper() for x in groups_in_data} and groups_in_data:
                    say(self.ui, "     not in the latest report", "dim")
            actions(self.ui, "Number: change", "a: add entity group", "h <n>: history / remove", "b: back")

        def handle(cmd):
            if cmd.isdigit():
                return self._edit_factor(self._pick(st["rows"], int(cmd)), as_of)
            if cmd == "a":
                name = ask("Entity group name (exactly as in the report)")
                if name:
                    self._edit_factor(name.strip(), as_of)
                return None
            if cmd.startswith("h ") and cmd[2:].strip().isdigit():
                return self._steps(self._pick(st["rows"], int(cmd[2:])))
            raise ValueError("Type a number, a, h <n> or b.")

        self.loop(render, handle)

    def _try_reports(self) -> bool:
        try:
            self.paths = core.list_reports(self.path(self.cfg["data_dir"]), self.cfg["file_pattern"])
            self.dates = list(self.paths)
            return True
        except core.DataError:
            return False

    @staticmethod
    def _pick(rows, n):
        if not 1 <= n <= len(rows):
            raise ValueError("Pick a number from the list.")
        return rows[n - 1]

    def _edit_factor(self, group, as_of):
        pp = self.prepos()
        seen = list(pp.rules.get(group.upper(), []))
        say(self.ui)
        say(self.ui, f"{group}: {core.pct(pp.factor(group, as_of))} on {core.fmt_date(as_of)}", "bold")
        raw = ask("New factor (e.g. 75%)", core.pct(pp.factor(group, as_of)))
        if raw is None:
            return
        factor = core.parse_factor(raw)
        raw = ask("From report date", f"{as_of[:4]}-{as_of[4:6]}-{as_of[6:]}")
        if raw is None:
            return
        start = core.parse_date(raw)
        if self.dates and start <= self.dates[-1]:
            say(self.ui, f"This changes binding for reports already published from {core.fmt_date(start)} "
                         "onwards (history is recalculated with the new factor).", "yellow")
        if not confirm(f"Set {group} to {core.pct(factor)} from {core.fmt_date(start)}?"):
            self.msg = ("Nothing changed.", "dim")
            return
        core.set_prepositioning(self.prepos_file, group, start, factor, expect=seen)
        self.msg = (f"Saved: {group} {core.pct(factor)} from {core.fmt_date(start)}.", "green")

    def _steps(self, group):
        pp = self.prepos()
        steps = list(pp.rules.get(group.upper(), []))
        if not steps:
            self.msg = (f"{group} has no factors set - always 100%.", "dim")
            return
        i = choose(self.ui, f"{group} - choose a step to remove (b to keep all)",
                   [f"{core.pct(v)} from {core.fmt_date(s)}" for s, v in steps])
        if i is None:
            return
        s, v = steps[i]
        if confirm(f"Remove {group} {core.pct(v)} from {core.fmt_date(s)}?"):
            core.set_prepositioning(self.prepos_file, group, s, None, expect=steps)
            self.msg = ("Removed.", "green")

    def _last_change(self):
        try:
            log = core._read_json(self.prepos_file).get("_log") or []
        except (core.DataError, OSError):
            return None
        if not log:
            return None
        e = log[-1]
        return f"{e.get('group')} {e.get('before')} → {e.get('after')} from {e.get('from')} ({e.get('who')}, {e.get('when')})"
