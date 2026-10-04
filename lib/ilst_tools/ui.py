from __future__ import annotations

import shutil
import textwrap

from . import core


def width() -> int:
    return max(40, min(shutil.get_terminal_size((100, 24)).columns, 100))


def say(ui, text: str = "", style: str = "") -> None:
    if len(text) > width():
        indent = " " * (len(text) - len(text.lstrip()))
        for part in textwrap.wrap(text, width(), subsequent_indent=indent):
            ui.print(ui.colour(part, style) if ui.rich else part)
        return
    ui.print(ui.colour(text, style) if ui.rich else text)


def title(ui, text: str, *context: str, warn: str = "", banner: str = "") -> None:
    ui.clear()
    say(ui)
    if banner:
        say(ui, banner, "bold yellow")
    for part in textwrap.wrap(text, width()):
        say(ui, part, "bold")
    for line in context:
        for part in textwrap.wrap(line, width()) or [""]:
            say(ui, part, "dim")
    if warn:
        for part in textwrap.wrap(warn, width()):
            say(ui, part, "yellow")
    say(ui)


def pair(left: str, right: str, indent: int = 0) -> list:
    w = width()
    lead = " " * indent
    if len(lead) + len(left) + 2 + len(right) <= w:
        return [lead + left + " " * (w - len(lead) - len(left) - len(right)) + right]
    lines = textwrap.wrap(left, w - indent, initial_indent=lead, subsequent_indent=lead) or [lead]
    return lines + [" " * max(0, w - len(right)) + right]


def numbered(n: int, name: str) -> list:
    return textwrap.wrap(name, width() - 5, initial_indent=f"{n:>3}  ", subsequent_indent="     ") or [f"{n:>3}"]


def label_value(label: str, value: str, label_w: int = 16) -> list:
    return textwrap.wrap(value, width(), initial_indent=label.ljust(label_w),
                         subsequent_indent=" " * label_w) or [label]


def actions(ui, *items: str) -> None:
    say(ui)
    for line in textwrap.wrap("  ·  ".join(items), width()):
        say(ui, line, "dim")


def ask(label: str = "", default: str = "") -> str | None:
    try:
        raw = input(f"{label}{f' [{default}]' if default else ''} › ").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return None
    return raw or default


def confirm(label: str) -> bool:
    ans = ask(f"{label} (y/n)", "n")
    return bool(ans) and ans.lower().startswith("y")


def choose(ui, heading: str, options: list, default: int | None = None) -> int | None:
    say(ui)
    say(ui, heading, "bold")
    for i, opt in enumerate(options, 1):
        for line in numbered(i, str(opt)):
            say(ui, line)
    while True:
        raw = ask("Number", str(default + 1) if default is not None else "")
        if raw is None or raw.lower() in ("q", "b", "back"):
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw) - 1
        say(ui, f"Type a number from 1 to {len(options)}.", "yellow")


def show_result(ui, res, units: str, decimals: int, limit: int | None) -> int:
    def m(v, sign=False):
        return core.money(v, units, decimals, sign)

    compare = len(res.columns) == 3
    if compare:
        for label, v in zip(res.columns, res.totals):
            for line in pair(label, m(v, label == "Change"), indent=2):
                say(ui, line, "bold" if label == "Change" else "")
    else:
        for line in pair("Total", m(res.totals[0]), indent=2):
            say(ui, line, "bold")
    say(ui)
    if not res.rows:
        say(ui, "  Nothing in this selection.", "dim")
        return 0
    shown = res.rows if limit is None else res.rows[:limit]
    w = width()
    for n, (key, vals) in enumerate(shown, 1):
        if compare:
            right = m(vals[2], True)
            detail = f"{m(vals[0])} → {m(vals[1])}"
            head = numbered(n, str(key))
            if len(head) == 1 and len(head[0]) + len(right) + 2 <= w:
                say(ui, head[0] + " " * (w - len(head[0]) - len(right)) + right)
            else:
                for line in head:
                    say(ui, line)
                say(ui, " " * max(0, w - len(right)) + right)
            say(ui, "     " + detail, "dim")
        else:
            head = numbered(n, str(key))
            right = m(vals[0])
            if len(head) == 1 and len(head[0]) + len(right) + 2 <= w:
                say(ui, head[0] + " " * (w - len(head[0]) - len(right)) + right)
            else:
                for line in head:
                    say(ui, line)
                say(ui, " " * max(0, w - len(right)) + right)
    if len(shown) < len(res.rows):
        rest = [sum(v[i] for _, v in res.rows[len(shown):]) for i in range(len(res.columns))]
        say(ui)
        for line in pair(f"     {len(res.rows) - len(shown)} more (a: show all)", m(rest[-1], compare)):
            say(ui, line, "dim")
    return len(shown)
