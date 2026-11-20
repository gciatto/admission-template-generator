"""Build the blank applicant workbook for one programme.

Python port of the original JavaScript layout builder and its XML post-processor.
It creates a new workbook and never reads application data. tools/generate.py
calls it, then applies the programme's texts, deadline and colours.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any, Literal

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.views import Pane, Selection

FILE_COUNTRIES = Path(__file__).with_name("countries.json")

NAMES = [
    "Instructions",
    "Personal & contacts",
    "Degrees",
    "Exams",
    "Work experience",
    "Language skills",
    "Internationalization",
    "Training",
    "Other activities",
    "Degree class legend",
    "Exam sector legend",
    "Lists",
]
TABS = "334155 2563EB 7C3AED 9333EA 0F766E C2410C BE185D 4D7C0F A16207 64748B 475569 94A3B8"
INPUT, GUIDE, EXAMPLE, FORMULA, HEADER = "FFF2A6", "DBEAFE", "FCE7F3", "E8EDF2", "23374D"
TEXT, BORDER, ERROR_FILL, ERROR_TEXT = "172033", "AAB6C4", "FCE3E3", "9C1C23"
SIDES = ("left", "right", "top", "bottom")
OTHER_GUIDE = "No listed cluster fits. Explain the subject in Other details."
_REVIEW = "Applicant classifications are subject to committee review."
SOURCE_NOTES = {
    "Degree class legend": "Clusters group official Italian degree classes for "
    "self-classification. " + _REVIEW,
    "Exam sector legend": "Clusters group official Italian subject codes (SSD) for "
    "self-classification. " + _REVIEW,
}
DEFAULT_PROMPT = "Use the specified format or choose a listed value."
EMPTY_PROMPT = "Choose a listed value or use the specified format."
DATE_RULE = {
    "type": "date",
    "operator": "between",
    "f1": "DATE(1900,1,1)",
    "f2": "DATE(2100,12,31)",
}
YES = ["Yes", "No"]
LEVEL = ["A1", "A2", "B1", "B2", "C1", "C2", "Not assessed", "Other"]
REFERENCE = "'Personal & contacts'!$B$34"
FREEZE_ROWS_ONLY = {"Personal & contacts", "Degree class legend", "Exam sector legend"}
# Later rows copy these inputs from the previous row (editable yellow defaults).
REPEATED_FIELDS = {
    "Degrees": [
        "Institution country",
        "Credit system",
        "Hours per credit",
        "Final scale minimum",
        "Final pass threshold",
        "Final scale maximum",
        "Final scale direction",
        "Exam scale minimum",
        "Exam pass threshold",
        "Exam scale maximum",
        "Exam scale direction",
    ],
    "Exams": [
        "Degree ID",
        "Grading type",
        "Scale minimum",
        "Pass threshold",
        "Scale maximum",
        "Scale direction",
    ],
    "Work experience": ["Country"],
    "Internationalization": ["Host country"],
    "Training": ["Country"],
    "Other activities": ["Country"],
}


def argb(rgb: str) -> str:
    return "FF" + rgb


def bounds(ref: str) -> tuple[int, int, int, int]:
    c1, r1, c2, r2 = range_boundaries(ref)
    assert c1 and r1 and c2 and r2, ref
    return c1, r1, c2, r2


def col(n: int) -> str:
    """Zero-based column index to letters."""
    return get_column_letter(n + 1)


def cells(ref: str) -> list[tuple[int, int]]:
    c1, r1, c2, r2 = bounds(ref)
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


class Sheet:
    """Accumulates cell styles the way the original authoring API merged them.

    Every touched cell starts from Carlito 11; each format call merges only the
    properties it names. Styles are materialised once, in Builder.materialise().
    """

    def __init__(self, name: str, tab: str) -> None:
        self.name, self.tab = name, tab
        self.values: dict[tuple[int, int], Any] = {}
        self.styles: dict[tuple[int, int], dict[str, Any]] = {}
        self.merges: list[str] = []
        self.widths: dict[str, float] = {}
        self.heights: dict[int, float] = {}
        self.tables: list[tuple[str, str, str]] = []
        self.validations: list[tuple[str, dict[str, Any], str]] = []
        self.conditional: list[tuple[str, str, bool]] = []
        self.unlocked: list[str] = []
        self.freeze = 0

    def style(self, rc: tuple[int, int]) -> dict[str, Any]:
        return self.styles.setdefault(
            rc,
            {
                "font": {"name": "Carlito", "size": 11, "bold": False, "italic": False},
                "fill": None,
                "numfmt": "General",
                "wrap": False,
                "h": None,
                "v": None,
                "border": {},
            },
        )

    def fmt(
        self,
        ref: str,
        *,
        fill: str | None = None,
        font: dict[str, Any] | None = None,
        wrap: bool | None = None,
        h: str | None = None,
        v: str | None = None,
        numfmt: str | None = None,
        height: float | None = None,
    ) -> None:
        for rc in cells(ref):
            s = self.style(rc)
            if fill:
                s["fill"] = fill
            if font:
                s["font"].update(font)
            if wrap is not None:
                s["wrap"] = wrap
            if h:
                s["h"] = h
            if v:
                s["v"] = v
            if numfmt:
                s["numfmt"] = numfmt
        if height:
            _, r1, _, r2 = bounds(ref)
            for r in range(r1, r2 + 1):
                self.heights[r] = height

    def borders(self, ref: str, outside: bool = False) -> None:
        c1, r1, c2, r2 = bounds(ref)
        anchors = {}
        for m in self.merges:
            mc1, mr1, mc2, mr2 = bounds(m)
            anchors[(mr1, mc1)] = (mc1, mr1, mc2, mr2)
        covered = {rc for m in self.merges for rc in cells(m)[1:]}
        for r, c in cells(ref):
            if (r, c) in covered:
                continue
            # A merged block takes the borders of its whole extent.
            bc1, br1, bc2, br2 = anchors.get((r, c), (c, r, c, r))
            sides = {
                "left": not outside or bc1 == c1,
                "right": not outside or bc2 == c2,
                "top": not outside or br1 == r1,
                "bottom": not outside or br2 == r2,
            }
            self.style((r, c))["border"].update({k: BORDER for k, on in sides.items() if on})

    def put(self, ref: str, rows: list[list[Any]]) -> None:
        c1, r1, _, _ = bounds(ref)
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                rc = (r1 + i, c1 + j)
                self.style(rc)
                if value is None or value == "":
                    self.values.pop(rc, None)
                else:
                    self.values[rc] = value

    def merge(self, ref: str) -> None:
        self.merges.append(ref)

    def validate(self, ref: str, rule: dict[str, Any], prompt: str = DEFAULT_PROMPT) -> None:
        self.validations.append((ref, rule, prompt or EMPTY_PROMPT))


def heading(s: Sheet, title: str, note: str, end: str = "H") -> None:
    s.fmt(f"A1:{end}4", font={"name": "Arial", "size": 11, "color": TEXT})
    s.merge(f"A2:{end}2")
    s.put("A2", [[title]])
    s.fmt("A2", font={"name": "Arial", "size": 16, "bold": True, "color": TEXT})
    s.merge(f"A3:{end}4")
    s.put("A3", [[note]])
    s.fmt(f"A3:{end}4", wrap=True, fill=GUIDE, height=16)


# Column definitions: plain dicts mirroring the original builder.
Column = dict[str, Any]
RowFn = Callable[[int, dict[str, str], int, int], str]


def text(h: str, e: Any = "", hint: str = "", req: bool = False) -> Column:
    return {"h": h, "e": e, "hint": hint, "req": req, "type": "text"}


def num(h: str, e: Any = 0, hint: str = "", req: bool = False) -> Column:
    return {"h": h, "e": e, "hint": hint, "req": req, "type": "decimal"}


def date(h: str, e: str = "2025-06-30", hint: str = "", req: bool = False) -> Column:
    return {"h": h, "e": e, "hint": hint, "req": req, "type": "date"}


def choice(h: str, values: list[str], e: Any = None, hint: str = "", req: bool = False) -> Column:
    first = values[0] if values else None
    return {
        "h": h,
        "values": values,
        "e": first if e is None else e,
        "hint": hint,
        "req": req,
        "type": "list",
    }


def calc(h: str, fn: RowFn, hint: str = "Automatic; do not edit") -> Column:
    return {"h": h, "fn": fn, "hint": hint, "type": "formula", "e": "Calculated"}


class Builder:
    def __init__(self, config: dict[str, Any]) -> None:
        """config: a programme validated by tools.generate.load_config."""
        self.sheets = {n: Sheet(n, t) for n, t in zip(NAMES, TABS.split())}
        self.lists: dict[tuple[str, ...], tuple[str, str]] = {}
        self.tables: list[dict[str, Any]] = []
        self.countries = json.loads(FILE_COUNTRIES.read_text(encoding="utf-8")) + ["Other"]
        self.config = config

    def clusters(self, field: str) -> tuple[list[str], str]:
        """Dropdown values and example for a classification: its clusters, then Other."""
        names = [g["name"] for g in self.config[field]] + ["Other"]
        return names, names[0]

    def add_list(self, values: list[str]) -> str:
        key = tuple(values)
        if key not in self.lists:
            c = col(len(self.lists))
            self.sheets["Lists"].put(f"{c}1", [[v] for v in values])
            self.lists[key] = (
                f"Options{len(self.lists) + 1}",
                f"'Lists'!${c}$1:${c}${len(values)}",
            )
        return self.lists[key][0]

    def list_ref(self, x: Column) -> str:
        return x.get("ref") or self.add_list(x["values"])

    def country(self, h: str) -> Column:
        return choice(
            h, self.countries, "IT — Italy", "Select country; use Other details if absent.", True
        )

    def table(self, sheet: str, name: str, cols: list[Column], start: int, count: int) -> dict:
        s = self.sheets[sheet]
        end, last, header = start + count - 1, col(len(cols) - 1), start - 1
        s.fmt(
            f"A{header}:{last}{end}", font={"name": "Arial", "size": 11, "color": TEXT}, height=27
        )
        s.fmt(f"A{start}:{last}{end}", fill=INPUT, v="center", wrap=True)
        s.put(f"A{header}", [[x["h"] + (" *" if x.get("req") else "") for x in cols]])
        s.fmt(
            f"A{header}:{last}{header}",
            fill=HEADER,
            font={"name": "Arial", "size": 11, "bold": True, "color": "FFFFFF"},
            wrap=True,
            h="center",
            v="center",
            height=38,
        )
        s.tables.append((name, f"A{header}:{last}{end}", "TableStyleLight1"))
        sample, hints = start - 3, start - 2
        s.fmt(f"A{sample}:{last}{sample}", numfmt="@")
        s.put(f"A{sample}", [[x["e"] for x in cols]])
        s.fmt(
            f"A{sample}:{last}{sample}",
            fill=EXAMPLE,
            font={"name": "Arial", "size": 10, "italic": True, "color": "334155"},
            wrap=True,
            height=28,
        )
        s.put(f"A{hints}", [[x.get("hint") or "" for x in cols]])
        s.fmt(
            f"A{hints}:{last}{hints}",
            fill=GUIDE,
            font={"name": "Arial", "size": 10, "color": "334155"},
            wrap=True,
            height=44,
        )
        m = {x["h"]: col(i) for i, x in enumerate(cols)}
        for i, x in enumerate(cols):
            c, h = col(i), x["h"]
            rng = f"{c}{start}:{c}{end}"
            s.widths[c] = (
                40
                if "relevant" in h or "description" in h.lower()
                else 32
                if "cluster" in h
                else 10
                if h == "ID"
                else 38
                if h == "Check"
                else 14
                if x["type"] == "date"
                else 17
                if x["type"] == "decimal"
                else 22
            )
            if x["type"] == "formula":
                s.fmt(rng, fill=FORMULA)
                for r in range(start, end + 1):
                    s.put(f"{c}{r}", [[x["fn"](r, m, start, end)]])
                continue
            s.unlocked.append(rng)
            if x["type"] == "list":
                s.validate(rng, {"type": "list", "f1": self.list_ref(x)}, x["hint"])
            elif x["type"] == "date":
                s.fmt(rng, numfmt="yyyy-mm-dd")
                s.validate(rng, DATE_RULE, "Enter a real Excel date. Display: yyyy-mm-dd.")
            elif x["type"] == "decimal":
                s.fmt(rng, numfmt="0.##")
                low = re.search(r"mark|minimum|threshold|maximum|average|score", h, re.IGNORECASE)
                s.validate(
                    rng,
                    {
                        "type": "whole" if h == "Remaining exams" else "decimal",
                        "operator": "between",
                        "f1": -1000000 if low else 0,
                        "f2": 1000000,
                    },
                )
            else:
                s.fmt(rng, numfmt="@")
                limit = 300 if "relevant" in h else 250
                s.validate(rng, {"type": "textLength", "operator": "lessThanOrEqual", "f1": limit})
            if h in REPEATED_FIELDS.get(sheet, []):
                # Different degrees may have different grading systems: ask again after a
                # Degree ID change.
                for r in range(start + 1, end + 1):
                    prev = f"{c}{r - 1}"
                    changed = f"$C{r}<>$C{r - 1}," if sheet == "Exams" and h != "Degree ID" else ""
                    s.put(f"{c}{r}", [[f'=IF(OR($B{r}="",{changed}{prev}=""),"",{prev})']])
        s.borders(f"A{header}:{last}{end}")
        if "Check" in m:
            c = m["Check"]
            s.conditional.append((f"{c}{start}:{c}{end}", f'{c}{start}<>""', True))
        meta = {"sheet": sheet, "name": name, "start": start, "end": end, "cols": cols, "map": m}
        self.tables.append(meta)
        s.freeze = 9
        return meta

    def checked_table(
        self,
        sheet: str,
        name: str,
        cols: list[Column],
        start: int,
        count: int,
        extra: Sequence[Callable[[int, dict[str, str], Callable[[str], str]], str]] = (),
    ) -> dict:
        original = list(cols)

        def check(r: int, m: dict[str, str], *_: int) -> str:
            def v(h: str) -> str:
                return f"{m[h]}{r}"

            terms = [
                f'IF({v(x["h"])}="","Missing {x["h"]}; ","")' for x in original if x.get("req")
            ]
            for x in original:
                h = x["h"]
                if x["type"] == "date":
                    terms.append(
                        f'IF(AND({v(h)}<>"",NOT(ISNUMBER({v(h)}))),"Invalid date: {h}; ","")'
                    )
                if x["type"] == "decimal":
                    terms.append(
                        f'IF(AND({v(h)}<>"",NOT(ISNUMBER({v(h)}))),"Invalid number: {h}; ","")'
                    )
                if x["type"] == "list":
                    terms.append(
                        f'IF(AND({v(h)}<>"",COUNTIF({self.list_ref(x)},{v(h)})=0),'
                        f'"Invalid choice: {h}; ","")'
                    )
            if "Other details" in m:
                terms.append(
                    f'IF(AND(COUNTIF(A{r}:{col(len(original) - 1)}{r},"Other")>0,'
                    f'{v("Other details")}=""),"Specify Other; ","")'
                )
            if "Start date" in m:
                end = start + count - 1

                def span(h: str) -> str:
                    return f"${m[h]}${start}:${m[h]}${end}"

                until = f'MIN(IF({v("Ongoing?")}="Yes",{REFERENCE},{v("End date")}),{REFERENCE})'
                terms.append(
                    f'IF(AND(ISNUMBER({v("Start date")}),{REFERENCE}<>"",'
                    f'COUNTIFS({span("Start date")},"<="&{until},{span("End date")},'
                    f'">="&{v("Start date")})+COUNTIFS({span("Start date")},"<="&{until},'
                    f'{span("End date")},"",{span("Ongoing?")},"Yes")>1),'
                    f'"Overlapping dates: committee review; ","")'
                )
            terms += [f(r, m, v) for f in extra]
            return f'=IF(B{r}="","",{"&".join(terms) if terms else chr(34) * 2})'

        cols.append(calc("Check", check))
        return self.table(sheet, name, cols, start, count)

    def declare(self, s: Sheet) -> None:
        s.put("A5", [["Status *"]])
        s.put("B5", [["Not started"]])
        s.fmt("B5", fill=INPUT)
        s.unlocked.append("B5")
        s.validate("B5", {"type": "list", "f1": self.add_list(["Not started", "Provided", "None"])})
        s.merge("C5:H5")
        s.put("C5", [["Choose Provided if entering rows; otherwise choose None."]])
        s.fmt("C5:H5", fill=GUIDE)

    def legends(self) -> None:
        for sheet, groups in [
            ("Degree class legend", self.config["degree_classes"]),
            ("Exam sector legend", self.config["exam_sectors"]),
        ]:
            s = self.sheets[sheet]
            heading(
                s,
                sheet + " — read only",
                "If your document states a listed code, select its cluster. Otherwise select "
                "the closest cluster from these descriptions. Use Other only if no cluster fits.",
                "E",
            )
            grouped = [
                [
                    g["name"],
                    ", ".join(c["code"] for c in g["codes"]),
                    "; ".join(c["en"] for c in g["codes"]),
                    "; ".join(c["it"] for c in g["codes"]),
                    g["description"],
                ]
                for g in groups
            ]
            grouped.append(["Other", "", "", "", OTHER_GUIDE])
            n = len(grouped)
            s.put(
                "A6",
                [
                    [
                        "Cluster to select",
                        "Included Italian codes",
                        "English names",
                        "Italian names",
                        "Plain-language guide",
                    ]
                ],
            )
            s.put("A7", grouped)
            s.fmt(f"A6:E{6 + n}", font={"name": "Arial", "size": 11}, wrap=True, v="center")
            s.fmt("A6:E6", fill=HEADER, font={"bold": True, "color": "FFFFFF"}, height=30)
            s.widths.update(A=31, B=25, C=43, D=43, E=62)
            for i, row in enumerate(grouped):
                # Enough wrapped lines for the longest cell, given each column's width.
                lines = max(
                    math.ceil(len(v) / max(1, s.widths[col(j)] * 1.05)) for j, v in enumerate(row)
                )
                s.fmt(f"A{7 + i}:E{7 + i}", height=max(28, lines * 13 + 6))
            s.fmt(f"E7:E{6 + n}", fill=GUIDE)
            s.borders(f"A6:E{6 + n}")
            s.merge(f"A{9 + n}:E{10 + n}")
            s.put(
                f"A{9 + n}",
                [[SOURCE_NOTES[sheet]]],
            )
            s.fmt(f"A{9 + n}:E{10 + n}", wrap=True)
            s.freeze = 6

    def personal(self) -> None:
        p = self.sheets["Personal & contacts"]
        heading(
            p,
            "Personal information and contacts",
            "Fill every required yellow cell. Write names in full. Expand national "
            "abbreviations, even when your ID uses them (MD → Mohammed). * = required.",
            "D",
        )
        p.put(
            "A6", [["Field", "Your answer", "Expected format / guidance", "Example (do not copy)"]]
        )
        p.put("A7", [[x[0], None, x[2], x[1]] for x in PERSONAL])
        p.fmt(
            "A6:D35",
            font={"name": "Arial", "size": 11, "color": TEXT},
            height=28,
            wrap=True,
            v="center",
        )
        p.fmt("A6:D6", fill=HEADER, font={"bold": True, "color": "FFFFFF"}, height=34)
        p.fmt("B7:B35", fill=INPUT, numfmt="@")
        p.unlocked.append("B7:B35")
        p.fmt("C7:C35", fill=GUIDE)
        p.fmt("D7:D35", fill=EXAMPLE)
        p.fmt("C7:C8", height=36)
        p.widths.update(A=29, B=30, C=64, D=27)
        p.tables.append(("PersonalDetails", "A6:D35", "TableStyleMedium2"))
        p.borders("A6:D35")
        for i, (label, _, hint) in enumerate(PERSONAL):
            r = i + 7
            rule: dict[str, Any] = {"type": "textLength", "operator": "lessThanOrEqual", "f1": 250}
            if re.search(r"Birth date|expiry|Reference date", label):
                p.fmt(f"B{r}", numfmt="yyyy-mm-dd")
                rule = DATE_RULE
            elif label in (
                "Citizenship *",
                "Home country *",
                "Residence country *",
                "ID issuing country *",
            ):
                rule = {"type": "list", "f1": self.add_list(self.countries)}
            elif label == "Applicant category *":
                categories = [
                    "EU candidate",
                    "Non-EU equivalent candidate",
                    "Non-EU candidate residing abroad",
                    "Other",
                    "Unknown",
                ]
                rule = {"type": "list", "f1": self.add_list(categories)}
            elif label == "ID document type *":
                rule = {"type": "list", "f1": self.add_list(["Passport", "Identity card", "Other"])}
            elif re.search("applicable|confirmed", label):
                rule = {"type": "list", "f1": self.add_list(YES)}
            p.validate(f"B{r}", rule, hint)
        p.merge("A38:D40")
        p.fmt("A38:D40", wrap=True)
        missing = "+".join(f'IF(B{i + 7}="",1,0)' for i, x in enumerate(PERSONAL) if "*" in x[0])
        p.put(
            "A38",
            [
                [
                    (
                        f'=IF({missing}>0,"Complete all required personal fields. ","")'
                        '&IF(AND(COUNTIF(B7:B35,"Other")>0,B23=""),'
                        '"Explain Other in Other category details. ","")'
                        '&IF(AND(B30="Yes",OR(B32="",B33="")),'
                        '"Add permit evidence filename and page. ","")'
                        '&IF(B35<>"Yes","Confirm accuracy. ","")'
                        "&IF(AND(ISNUMBER(B27),ISNUMBER(B34),B27<B34),"
                        '"Identity document expires before reference date. ","")'
                    )
                ]
            ],
        )
        p.conditional.append(("A38:D40", 'A38<>""', False))
        p.freeze = 6

    def degrees_and_exams(self) -> None:
        sh = self.sheets
        evidence = EVIDENCE
        heading(
            sh["Degrees"],
            "Degrees",
            "Select a degree cluster using the Degree class legend. A listed official code "
            "determines the cluster; otherwise choose the closest description. * = required.",
        )
        self.declare(sh["Degrees"])
        scale_dir = ["Higher is better", "Lower is better", "Not numeric", "Other"]
        honours = ["Yes", "No", "Not applicable"]
        deg = [
            id_column("D"),
            text(
                "Degree title",
                "BSc Computer Science",
                "Official title; add English translation if useful.",
                True,
            ),
            choice(
                "Qualification level",
                ["Bachelor / first cycle", "Master / second cycle", "Doctorate", "Other"],
            ),
            text("Institution", "Example University", "Official awarding institution.", True),
            self.country("Institution country"),
            text("Programme student ID", "00012345", "Your matriculation / programme number."),
            date("Enrolment date", "2021-09-01", "Start of this programme.", True),
            choice(
                "Status",
                ["Graduated", "Graduating"],
                "Graduated",
                "Graduating = not yet awarded.",
                True,
            ),
            date("Graduation date", "2024-07-15", "Required only for Graduated."),
            date("Expected graduation", "2027-07-15", "Required only for Graduating."),
            num("Formal duration (months)", 36, "Official full-time programme length.", True),
            calc(
                "Actual duration (months)",
                lambda r, *_: (
                    f'=IF(OR(G{r}="",H{r}<>"Graduated",I{r}="",I{r}<G{r}),"",ROUND((I{r}-G{r})/30.4375,1))'
                ),
                "Completed only; elapsed days / 30.4375.",
            ),
            choice(
                "Main admission degree?", YES, "Yes", "Select Yes for exactly one degree.", True
            ),
            choice(
                "Degree cluster",
                *self.clusters("degree_classes"),
                "See Degree class legend: use the cluster matching your official code, or the closest description if unlabelled.",
                True,
            ),
            choice(
                "Credit system",
                ["ECTS", "Local credits", "Other"],
                "ECTS",
                "Do not invent conversions.",
                True,
            ),
            num("Hours per credit", 25, "Documented student workload, not contact hours only."),
            num("Programme credits", 180, "Original credit system.", True),
            num("Credits acquired", 180, "Passed / recognised credits only.", True),
            num("Programme ECTS", 180, "Only if documented; blank if unavailable."),
            num("ECTS acquired", 180, "Only if documented; blank if unavailable."),
            num("Final mark / GPA", 105, "Awarded degree result; blank until awarded."),
            num(
                "Final scale minimum", 0, "Lowest possible final result, including failing values."
            ),
            num("Final pass threshold", 66, "Minimum final result needed to pass."),
            num("Final scale maximum", 110, "Highest possible numeric final result."),
            choice("Final scale direction", scale_dir),
            choice("Final honours", honours, "No"),
            num("Weighted exam average", 27.5, "Use original exam scale; do not convert."),
            num("Exam scale minimum", 0),
            num("Exam pass threshold", 18),
            num("Exam scale maximum", 30),
            choice("Exam scale direction", scale_dir),
            num("Remaining exams", 0, "Whole number; 0 if none.", True),
            *evidence(),
            other_details(),
        ]
        self.checked_table(
            "Degrees",
            "Degrees",
            deg,
            10,
            8,
            [
                lambda r, m, v: (
                    f'IF(AND({v("Status")}="Graduated",{v("Graduation date")}=""),"Add graduation date; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Status")}="Graduating",{v("Expected graduation")}=""),"Add expected graduation; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Graduation date")}<>"",{v("Graduation date")}<{v("Enrolment date")}),"Graduation before enrolment; ","")'
                ),
                lambda r, m, v: (
                    f'IF({v("Credits acquired")}>{v("Programme credits")},"Acquired credits exceed total; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Credit system")}<>"ECTS",{v("Hours per credit")}=""),"Add documented hours per credit or explain unavailable in Other details; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Final mark / GPA")}<>"",OR({v("Final scale minimum")}="",{v("Final pass threshold")}="",{v("Final scale maximum")}="")),"Complete final grading scale; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Final mark / GPA")}<>"",OR({v("Final mark / GPA")}<{v("Final scale minimum")},{v("Final mark / GPA")}>{v("Final scale maximum")})),"Final grade outside range; ","")'
                ),
            ],
        )
        degree_ids = choice(
            "Degree ID", [], "D001", "Select an automatically generated degree ID.", True
        )
        degree_ids["ref"] = "DegreeIDs"
        lookup = "Degrees!$A$10:$AI$17"
        exam = [
            id_column("EX"),
            text("Exam title", "Algorithms", "One row per attempt; copy transcript exactly.", True),
            degree_ids,
            choice(
                "Result",
                ["Passed", "Failed"],
                "Passed",
                "Record every attempt shown on the transcript.",
                True,
            ),
            date(
                "Result date", "2024-06-20", "Exam / attempt date, including failed attempts.", True
            ),
            num(
                "Original credits", 6, "Credits for this course, in the degree credit system.", True
            ),
            calc(
                "Credit system",
                lambda r, *_: (
                    f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},{lookup},15,FALSE),"Unknown degree"))'
                ),
            ),
            calc(
                "Hours per credit",
                lambda r, *_: (
                    f'=IF(C{r}="","",IFERROR(IF(VLOOKUP(C{r},{lookup},16,FALSE)=0,"",VLOOKUP(C{r},{lookup},16,FALSE)),""))'
                ),
            ),
            num("Documented ECTS", 6, "No self-invented conversion."),
            choice(
                "Grading type",
                ["Numeric", "Pass/fail", "Other"],
                "Numeric",
                "Use Pass/fail if no numeric mark.",
                True,
            ),
            num("Mark", 28, "Blank for pass/fail grading."),
            num("Scale minimum", 0),
            num("Pass threshold", 18),
            num("Scale maximum", 30),
            choice("Scale direction", scale_dir),
            choice("Honours", honours, "No"),
            choice(
                "Exam cluster",
                *self.clusters("exam_sectors"),
                "See Exam sector legend: use the cluster matching your official SSD, or the closest description if unlabelled.",
                True,
            ),
            *evidence(),
            other_details(),
        ]
        heading(
            sh["Exams"],
            "Exams",
            "Select a Degree ID from Degrees. Select an exam cluster using the Exam sector "
            "legend: match a stated SSD, otherwise choose the closest description.",
        )
        self.declare(sh["Exams"])
        self.checked_table(
            "Exams",
            "Exams",
            exam,
            10,
            200,
            [
                lambda r, m, v: (
                    f'IF(COUNTIF(Degrees!$A$10:$A$17,{v("Degree ID")})<>1,"Unknown degree ID; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Grading type")}="Numeric",OR({v("Mark")}="",{v("Scale minimum")}="",{v("Pass threshold")}="",{v("Scale maximum")}="")),"Complete numeric mark and scale; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Mark")}<>"",OR({v("Mark")}<{v("Scale minimum")},{v("Mark")}>{v("Scale maximum")})),"Mark outside range; ","")'
                ),
            ],
        )

    def activities(self) -> None:
        sh = self.sheets
        checks = [
            lambda r, m, v: f'IF(AND({v("Ongoing?")}="No",{v("End date")}=""),"Add end date; ","")',
            lambda r, m, v: (
                f'IF(AND({v("Ongoing?")}="Yes",{v("End date")}<>""),"Ongoing: clear end date; ","")'
            ),
            lambda r, m, v: (
                f'IF(AND({v("End date")}<>"",{v("End date")}<{v("Start date")}),"End before start; ","")'
            ),
            lambda r, m, v: (
                f'IF({REFERENCE}="","Set reference date in Personal; ",IF({v("Start date")}>{REFERENCE},"Starts after reference date; ",""))'
            ),
        ]

        def duration(r: int, m: dict[str, str], *_: int) -> str:
            s, e, o = f"{m['Start date']}{r}", f"{m['End date']}{r}", f"{m['Ongoing?']}{r}"
            until = f'MIN(IF({o}="Yes",{REFERENCE},{e}),{REFERENCE})'
            return (
                f'=IF(OR({s}="",{REFERENCE}="",{o}="",AND({o}="No",{e}="")),"",'
                f'IF({until}<{s},"",ROUND(({until}-{s}+1)/30.4375,1)))'
            )

        def period() -> list[Column]:
            return [
                date("Start date", "2024-01-01", "Real date.", True),
                date("End date", "2024-06-30", "Blank only if ongoing."),
                choice("Ongoing?", YES, "No", "Yes: leave End date blank.", True),
                calc(
                    "Duration (months)",
                    duration,
                    "Calendar duration through reference date; inclusive days / 30.4375. "
                    "Not an eligibility score.",
                ),
            ]

        def relevance() -> Column:
            return text(
                "Role / why relevant",
                "Built a reporting tool used to plan production.",
                "Max 300 characters: your contribution and relevance to computing, business or "
                "digital transformation.",
                True,
            )

        def link() -> Column:
            x = choice(
                "Related activity ID",
                [],
                "W001",
                "Optional: select the same activity recorded elsewhere. Do not count twice.",
            )
            x["ref"] = "ActivityIDs"
            return x

        work = [
            id_column("W"),
            text("Position", "Junior data analyst", "Short role title.", True),
            text("Employer", "Example Ltd", "Organisation name.", True),
            self.country("Country"),
            text("City", "Bologna", "Work location.", True),
            *period(),
            choice(
                "Activity type",
                ["Employment", "Internship", "Volunteering", "Self-employed", "Other"],
                "Employment",
                "Select the relationship.",
                True,
            ),
            choice("Paid?", YES, "Yes", "Separate from activity type.", True),
            choice(
                "Schedule",
                ["Full-time", "Part-time", "Variable", "Other"],
                "Full-time",
                "Specify actual hours below.",
                True,
            ),
            num("Hours per week", 40, "Average weekly workload.", True),
            relevance(),
            *EVIDENCE(),
            other_details(),
        ]
        heading(
            sh["Work experience"],
            "Work experience",
            "One row per role. Include internships and volunteering; clearly identify their "
            "type. Pink rows are examples. * = required.",
        )
        self.declare(sh["Work experience"])
        self.checked_table("Work experience", "Work", work, 10, 30, checks)

        languages = [
            id_column("L"),
            choice(
                "Language",
                [
                    "English",
                    "Italian",
                    "French",
                    "German",
                    "Spanish",
                    "Portuguese",
                    "Arabic",
                    "Chinese",
                    "Russian",
                    "Turkish",
                    "Persian",
                    "Other",
                ],
                "English",
                "Other: specify language in Other details.",
                True,
            ),
            choice("Native?", YES, "No", "Native is not a CEFR certificate level.", True),
            choice(
                "Certified CEFR",
                [*LEVEL, "Not stated"],
                "B2",
                "Use Not assessed if no certification.",
                True,
            ),
            *[
                choice(f"{h} self-assessment", LEVEL, "B2", "Your own CEFR assessment.", True)
                for h in ("Speaking", "Reading", "Writing", "Listening")
            ],
            choice(
                "Evidence route",
                [
                    "Language certificate",
                    "English-taught qualification",
                    "Transcript language exam",
                    "Native declaration",
                    "Self-assessment only",
                    "Other",
                ],
                "Language certificate",
                "Select the evidence you actually submit.",
                True,
            ),
            choice(
                "Certificate type",
                [
                    "IELTS",
                    "TOEFL",
                    "Cambridge",
                    "University language centre",
                    "Other",
                    "Not applicable",
                ],
                "Cambridge",
            ),
            text("Issuer", "Cambridge English", "Official certificate / evidence provider."),
            text("Certificate ID", "ABC00123", "If present."),
            date("Certificate date", "2025-06-01", "Award / test date, if applicable."),
            date("Certificate expiry", "2027-06-01", "Blank if none is stated."),
            num("Test score", 175, "Optional original score."),
            text("Score scale", "80–230", "Original test scale; not a conversion."),
            *[{**x, "req": False} for x in EVIDENCE()],
            other_details(),
        ]
        heading(
            sh["Language skills"],
            "Language skills",
            "One row per language / certificate. Self-assessment does not prove a certified "
            "level. Evidence is required for a claimed certificate.",
        )
        self.declare(sh["Language skills"])
        self.checked_table(
            "Language skills",
            "Languages",
            languages,
            10,
            20,
            [
                lambda r, m, v: (
                    f'IF(AND({v("Evidence route")}<>"Self-assessment only",{v("Evidence route")}<>"Native declaration",OR({v("Evidence filename")}="",{v("Page / section")}="")),"Add evidence reference; ","")'
                ),
                lambda r, m, v: (
                    f'IF(AND({v("Evidence route")}="Language certificate",OR({v("Issuer")}="",{v("Certificate date")}="")),"Add certificate issuer and date; ","")'
                ),
            ],
        )

        abroad = [
            id_column("AB"),
            choice(
                "Experience type",
                ["Study exchange", "Internship", "Work", "Research", "Volunteering", "Other"],
                "Study exchange",
                "Physical stay outside your home country.",
                True,
            ),
            text("Host organisation", "Example University", "Employer / university / host.", True),
            self.country("Host country"),
            text("Host city", "Berlin", "Physical location.", True),
            choice(
                "Participation",
                ["In person", "Hybrid", "Remote", "Other"],
                "In person",
                "Remote activity is flagged for committee review.",
                True,
            ),
            *period(),
            link(),
            relevance(),
            *EVIDENCE(),
            other_details(),
        ]
        heading(
            sh["Internationalization"],
            "Internationalization",
            "Home country = country where you spent most of your life (Personal). Exclude stays "
            "in that country. Link an existing work/training activity instead of claiming it "
            "twice.",
        )
        self.declare(sh["Internationalization"])
        home = "'Personal & contacts'!$B$13"
        self.checked_table(
            "Internationalization",
            "Abroad",
            abroad,
            10,
            30,
            [
                *checks,
                lambda r, m, v: (
                    f'IF({home}="","Set home country in Personal; ",IF({v("Host country")}={home},"Exclude home-country experience; ",""))'
                ),
                lambda r, m, v: (
                    f'IF({v("Participation")}="Remote","Remote: not a physical stay abroad; ","")'
                ),
            ],
        )

        training = [
            id_column("T"),
            text("Activity title", "Data analytics summer school", "Short title.", True),
            choice(
                "Activity type",
                [
                    "Internal internship",
                    "External internship",
                    "School project",
                    "Summer school",
                    "Postgraduate course",
                    "Professional course",
                    "Workshop",
                    "Language course",
                    "Other",
                ],
                "Summer school",
                "Select closest category.",
                True,
            ),
            text("Institution", "Example Institute", "Provider / host.", True),
            self.country("Country"),
            *period(),
            num("Total hours", 48, "Actual documented learning / work hours.", True),
            choice(
                "Completion",
                ["Completed", "Ongoing", "Not completed", "Other"],
                "Completed",
                "Status at reference date.",
                True,
            ),
            choice(
                "Part of main degree?",
                YES,
                "No",
                "Yes if already part of the academic curriculum.",
                True,
            ),
            link(),
            relevance(),
            *EVIDENCE(),
            other_details(),
        ]
        heading(
            sh["Training"],
            "Training and additional education",
            "Include training beyond your main academic programme. Identify curricular "
            "activities and link any internship already recorded under Work.",
        )
        self.declare(sh["Training"])
        self.checked_table("Training", "Training", training, 10, 30, checks)

        others = [
            id_column("O"),
            text("Activity title", "Open-source inventory project", "Short name.", True),
            choice(
                "Activity type",
                [
                    "Independent project",
                    "Open-source contribution",
                    "Startup / entrepreneurship",
                    "Competition / award",
                    "Volunteering",
                    "Research activity",
                    "Other",
                ],
                "Independent project",
                "Only activities that do not fit another sheet.",
                True,
            ),
            text("Organisation / context", "Independent", "Organisation or Independent.", True),
            self.country("Country"),
            *period(),
            num("Total hours", 120, "Documented estimate; explain basis in relevance."),
            link(),
            relevance(),
            *EVIDENCE(),
            other_details(),
        ]
        heading(
            sh["Other activities"],
            "Other relevant activities",
            "Use for activities that do not fit the other sheets. Explain your contribution and "
            "its relevance. Do not repeat an existing activity.",
        )
        self.declare(sh["Other activities"])
        self.checked_table("Other activities", "OtherActivities", others, 10, 30, checks)

    def lists_sheet(self) -> None:
        # Hidden, read-only option lists backing the dropdowns.
        s = self.sheets["Lists"]
        for i, values in enumerate(self.lists):
            s.put(f"{col(i)}1", [[v] for v in values])
            s.widths[col(i)] = 34

    def instructions(self) -> None:
        ins = self.sheets["Instructions"]
        heading(
            ins,
            "Applicant evidence workbook",
            "Complete this workbook and submit it together with the documents you reference. "
            "No macros are required.",
        )
        ins.widths.update(A=6, B=32, **{c: 15 for c in "CDEFGH"})
        ins.merge("A6:H6")
        ins.put(
            "A6",
            [["Yellow = input. Blue = guidance. Pink = example. Grey = formula. Red = review."]],
        )
        ins.fmt("A6:H6", fill=INPUT, font={"bold": True}, height=30)
        for i, (number, title, body) in enumerate(STEPS):
            r = 8 + i * 2
            ins.merge(f"A{r}:A{r + 1}")
            ins.merge(f"B{r}:B{r + 1}")
            ins.merge(f"C{r}:H{r + 1}")
            ins.put(f"A{r}", [[number]])
            ins.put(f"B{r}", [[title]])
            ins.put(f"C{r}", [[body]])
            ins.fmt(
                f"A{r}:H{r + 1}",
                font={"name": "Arial", "size": 11, "color": TEXT},
                wrap=True,
                v="center",
                height=23,
            )
            ins.fmt(f"B{r}", font={"bold": True})
        ins.merge("A34:H35")
        ins.put(
            "A34",
            [
                [
                    (
                        "Home country: the country where you have spent most of your life. If tied, "
                        "explain in Personal → Other category details. Durations are calendar months "
                        "(days / 30.4375), not awarded points. Overlapping periods and "
                        "self-classifications require committee review."
                    )
                ]
            ],
        )
        ins.fmt("A34:H35", wrap=True, height=27, font={"size": 10, "color": "475569"})
        ins.put("A37", [["Section", "Declared status", "Rows entered", "Rows with warnings"]])
        ins.fmt("A37:D37", fill=HEADER, font={"color": "FFFFFF", "bold": True}, height=30)
        categories = NAMES[2:9]
        for i, n in enumerate(categories):
            r = 38 + i
            ts = [t for t in self.tables if t["sheet"] == n]
            ins.put(f"A{r}", [[n]])
            ins.put(f"B{r}", [[f"='{n}'!B5"]])
            ins.put(
                f"C{r}",
                [["=" + "+".join(f"COUNTA('{n}'!B{t['start']}:B{t['end']})" for t in ts)]],
            )
            ins.put(
                f"D{r}",
                [
                    [
                        "="
                        + "+".join(
                            f"COUNTIF('{n}'!{t['map']['Check']}{t['start']}:"
                            f'{t["map"]["Check"]}{t["end"]},"?*")'
                            for t in ts
                        )
                    ]
                ],
            )
        ins.fmt("A38:D44", font={"name": "Arial", "size": 11}, height=30)
        ins.widths["A"] = 26
        ins.merge("E37:H37")
        ins.put("E37", [["Section check"]])
        for i in range(len(categories)):
            r = 38 + i
            ins.merge(f"E{r}:H{r}")
            ins.put(
                f"E{r}",
                [
                    [
                        (
                            f'=IF(B{r}="Not started","Choose Provided or None",'
                            f'IF(AND(B{r}="None",C{r}>0),"None conflicts with entered rows",'
                            f'IF(AND(B{r}="Provided",C{r}=0),"Add rows or choose None",'
                            f'IF(D{r}>0,"Review red Check cells","Review evidence before submission"))))'
                        )
                    ]
                ],
            )
        ins.fmt("E38:H44", font={"name": "Arial", "size": 11}, wrap=True, height=30)
        ins.borders("A37:D44")
        ins.borders("E37:H44", outside=True)
        ins.merge("A46:H47")
        ins.put(
            "A46",
            [
                [
                    (
                        '=IF(COUNTIF(Degrees!M10:M17,"Yes")<>1,'
                        '"Education: select exactly one main admission degree.",'
                        '"Main admission degree selected. Check its transcript and grading scales.")'
                    )
                ]
            ],
        )
        ins.fmt("A46:H47", wrap=True)

    def activity_ids(self) -> str:
        s = self.sheets["Lists"]
        c = col(len(self.lists) + 1)
        row = 1
        for n in ("Work experience", "Training", "Other activities"):
            for r in range(10, 40):
                s.put(f"{c}{row}", [[f"=IF('{n}'!A{r}=\"\",\"\",'{n}'!A{r})"]])
                row += 1
        return f"'Lists'!${c}$1:${c}${row - 1}"

    def build(self) -> openpyxl.Workbook:
        self.legends()
        self.personal()
        self.degrees_and_exams()
        self.activities()
        self.lists_sheet()
        self.instructions()
        names = [*self.lists.values(), ("DegreeIDs", "'Degrees'!$A$10:$A$17")]
        names.append(("ActivityIDs", self.activity_ids()))
        return self.materialise(names)

    def materialise(self, names: list[tuple[str, str]]) -> openpyxl.Workbook:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # type: ignore[arg-type]
        wb._fonts = IndexedList([Font(name="Carlito", sz=11)])  # type: ignore[attr-defined]
        for name, ref in names:
            wb.defined_names[name] = DefinedName(name, attr_text=ref)
        wb.security = WorkbookProtection(lockStructure=True)
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        side = Side(style="thin", color=argb(BORDER))
        for s in self.sheets.values():
            ws = wb.create_sheet(s.name)
            ws.sheet_properties.tabColor = argb(s.tab)
            unlocked = {rc for ref in s.unlocked for rc in cells(ref)}
            for (r, c), st in s.styles.items():
                cell = ws.cell(r, c, s.values.get((r, c)))
                f = st["font"]
                cell.font = Font(
                    name=f["name"],
                    sz=f["size"],
                    b=f["bold"] or None,
                    i=f["italic"] or None,
                    color=argb(f["color"]) if f.get("color") else None,
                )
                if st["fill"]:
                    cell.fill = PatternFill("solid", fgColor=argb(st["fill"]))
                cell.number_format = st["numfmt"]
                if st["wrap"] or st["h"] or st["v"]:
                    cell.alignment = Alignment(
                        wrap_text=st["wrap"] or None, horizontal=st["h"], vertical=st["v"]
                    )
                if st["border"]:
                    cell.border = Border(
                        **{k: side if k in st["border"] else Side() for k in SIDES}  # type: ignore[arg-type]
                    )
                cell.protection = Protection(locked=(r, c) not in unlocked, hidden=False)
            for ref in s.merges:
                ws.merge_cells(ref)
            for letter, w in s.widths.items():
                ws.column_dimensions[letter].width = w
            for r, h in s.heights.items():
                ws.row_dimensions[r].height = h
            for name, ref, style in s.tables:
                c1, r1, c2, _ = bounds(ref)
                header = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
                ws.add_table(
                    Table(
                        displayName=name,
                        ref=ref,
                        tableColumns=[TableColumn(id=i, name=n) for i, n in enumerate(header, 1)],
                        tableStyleInfo=TableStyleInfo(
                            name=style, showRowStripes=False, showColumnStripes=False
                        ),
                    )
                )
            for ref, rule, prompt in s.validations:
                ws.add_data_validation(
                    DataValidation(
                        type=rule["type"],
                        operator=rule.get("operator"),
                        formula1=str(rule["f1"]),
                        formula2=str(rule["f2"]) if "f2" in rule else None,
                        sqref=ref,
                        allow_blank=True,
                        showInputMessage=True,
                        showErrorMessage=True,
                        errorStyle="stop",
                        errorTitle="Check your entry",
                        error="Use the listed values or the expected data type.",
                        promptTitle="How to fill this cell",
                        prompt=prompt[:255],
                    )
                )
            for ref, formula, bold in s.conditional:
                ws.conditional_formatting.add(
                    ref,
                    FormulaRule(
                        formula=[formula],
                        fill=PatternFill("solid", bgColor=argb(ERROR_FILL)),
                        font=Font(color=argb(ERROR_TEXT), b=bold or None),
                    ),
                )
            view = ws.sheet_view
            view.showGridLines = False
            view.zoomScale = view.zoomScaleNormal = 85
            if s.freeze:
                x = 0 if s.name in FREEZE_ROWS_ONLY else 2
                top = f"{'C' if x else 'A'}{s.freeze + 1}"
                active: Literal["bottomRight", "bottomLeft"] = "bottomRight" if x else "bottomLeft"
                view.pane = Pane(
                    xSplit=x or None,
                    ySplit=s.freeze,
                    topLeftCell=top,
                    activePane=active,
                    state="frozen",
                )
                view.selection = [Selection(pane=active, activeCell=top, sqref=top)]
            else:
                view.selection = [Selection(activeCell="A1", sqref="A1")]
            ws.protection = SheetProtection(
                sheet=True,
                objects=True,
                scenarios=True,
                formatCells=True,
                formatColumns=True,
                formatRows=False,
                insertColumns=True,
                insertRows=True,
                deleteColumns=True,
                deleteRows=True,
                sort=True,
                autoFilter=False,
                selectLockedCells=False,
                selectUnlockedCells=False,
            )
            if s.name == "Lists":
                ws.sheet_state = "hidden"
        return wb


def id_column(prefix: str) -> Column:
    return calc(
        "ID",
        lambda r, m, start, end: f'=IF(B{r}="","","{prefix}"&TEXT(ROW()-{start - 1},"000"))',
    )


def EVIDENCE() -> list[Column]:
    return [
        text(
            "Evidence filename",
            "transcript.pdf",
            "Exact attachment filename; separate files with ;",
            True,
        ),
        text("Page / section", "2–4", "Locate the supporting passage.", True),
    ]


def other_details() -> Column:
    return text("Other details", "", "Required when any category is Other.")


_ABBREV = (
    "Write in full; expand national abbreviations, e.g. MD → Mohammed, even if used on your ID."
)
_DATE = "Real date; displayed yyyy-mm-dd."
_PERMIT = "Required when permit / visa is applicable."
PERSONAL = [
    ["First name(s) *", "Alex", _ABBREV],
    ["Surname(s) *", "Example", _ABBREV],
    ["Birth date *", "2001-04-15", _DATE],
    ["Birth place *", "Bologna", "City / region / country."],
    ["Citizenship *", "IT — Italy", "Select your main citizenship."],
    ["Additional citizenships", "", "Country names; separate with ;, or leave blank."],
    ["Home country *", "IT — Italy", "Country where you have spent most of your life."],
    ["Residential address *", "12 Example Street", "Street and building number."],
    ["City *", "Bologna", "Current residence."],
    ["Postal code *", "00123", "Text; keep leading zeros."],
    ["Residence country *", "IT — Italy", "Current residence."],
    ["Phone *", "+39 0123456789", "Include international calling code."],
    ["Personal email *", "alex@example.invalid", "Email you check regularly."],
    ["University email", "", "If already assigned."],
    ["Student / application number", "", "If already assigned; text, not a number."],
    ["Applicant category *", "EU candidate", "Select the category used in your application."],
    ["Other category details", "", "Complete if any selected category is Other."],
    ["ID document type *", "Passport", "Select document type."],
    ["ID document number *", "AB0012345", "Keep letters and leading zeros."],
    ["ID issuing country *", "IT — Italy", "Country that issued the document."],
    ["ID expiry date *", "2030-04-15", _DATE],
    ["ID evidence filename *", "identity.pdf", "Exact submitted filename."],
    ["ID page / side *", "1–2", "Include all relevant sides."],
    ["Permit / visa applicable? *", "No", "If applicable, complete the next three fields."],
    ["Permit / visa expiry", "", "Leave blank only if not applicable or no expiry is stated."],
    ["Permit evidence filename", "", _PERMIT],
    ["Permit page / section", "", _PERMIT],
    [
        "Reference date *",
        "2026-11-20",
        "Use the application deadline. All ongoing durations stop on this date.",
    ],
    ["Accuracy confirmed? *", "Yes", "Confirm that your entries match the submitted evidence."],
]
STEPS = [
    [
        "1",
        "Start with Personal",
        "Enter your legal details, home country and the application deadline as Reference date.",
    ],
    [
        "2",
        "Edit yellow cells only",
        "Pink = example; blue = guidance; grey = formula. * = required. Rows are compact: increase a row height to read longer entries.",
    ],
    [
        "3",
        "Choose the section status",
        "On each category sheet choose Provided or None. A blank table is not a declaration of no experience.",
    ],
    [
        "4",
        "Use the empty rows",
        "Examples are locked above each table. Start in row 10 on every input table. Leave unused rows completely empty.",
    ],
    [
        "5",
        "Use dropdowns and real dates",
        "Dates display as yyyy-mm-dd. If ISO typing remains text in your Excel locale, enter a local date and check its ISO display.",
    ],
    [
        "6",
        "Use Other honestly",
        "Choose Other if needed and complete Other details. Unknown means you do not know; Not applicable means it does not apply.",
    ],
    [
        "7",
        "Attach evidence",
        "Type the exact submitted filename and page / section for each claim. Submit the actual files separately; a filename alone proves nothing.",
    ],
    [
        "8",
        "Link your records",
        "IDs appear automatically when you enter a title. Select a Degree ID for each exam and Related activity ID for repeated contexts.",
    ],
    [
        "9",
        "Check carried-forward settings",
        "Starting a row copies selected settings from the previous row. Yellow defaults are editable and stay linked until overwritten. Check each value. New degree: enter exam grading settings again.",
    ],
    [
        "10",
        "Check before submitting",
        "Review every red Check cell, all required fields, section statuses and evidence attachments. No red warning does not certify completeness.",
    ],
    [
        "11",
        "Save and submit",
        "Save as .xlsx with your name or application number. Keep the provided row order; filtering is allowed, sorting and row insertion are blocked.",
    ],
    [
        "12",
        "If you run out of rows",
        "Ask the admissions office for an expanded copy. Capacity: 8 degrees, 200 exams, 30 activities per sheet, 20 language records.",
    ],
]
