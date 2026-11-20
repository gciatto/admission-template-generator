"""Customize the approved blank layout and optionally convert it to OpenDocument."""

import argparse
import re
import shutil
import subprocess
import tempfile
import zipfile
from copy import copy
from datetime import date
from math import ceil
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import yaml
from openpyxl.styles import Alignment
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent
LAYOUT = ROOT / "layout/applicant-template.xlsx"
PALETTE = {
    "input": "FFF2A6",
    "guidance": "DBEAFE",
    "example": "FCE7F3",
    "formula": "E8EDF2",
    "header": "23374D",
    "text": "172033",
    "header_text": "FFFFFF",
    "muted_text": "334155",
    "secondary_text": "475569",
    "error_fill": "FCE3E3",
    "error_text": "9C1C23",
    "border": "AAB6C4",
}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def keys(value, allowed, where):
    require(isinstance(value, dict), f"{where} must be a mapping")
    require(not value.keys() - allowed, f"Unknown keys in {where}: {value.keys() - allowed}")


def string(value, where, limit=1000):
    require(
        isinstance(value, str) and 0 < len(value.strip()) <= limit,
        f"{where} must be a non-empty string of at most {limit} characters",
    )
    require(
        not value.startswith(("=", "+", "-", "@")), f"Formula-like text is forbidden in {where}"
    )
    return value


def load_config(path):
    config = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    keys(config, {"programme", "degree_classes", "exam_sectors", "colors"}, "config")
    programme = config.get("programme")
    keys(
        programme,
        {"slug", "name", "institution", "academic_year", "url", "deadline", "topics"},
        "programme",
    )
    for field in ("slug", "name"):
        string(programme.get(field), f"programme.{field}")
    require(re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", programme["slug"]), "Invalid programme.slug")
    for field in ("institution", "academic_year", "url"):
        if field in programme:
            string(programme[field], f"programme.{field}")
    topics = programme.get("topics")
    require(isinstance(topics, list) and bool(topics), "programme.topics must be a non-empty list")
    for topic in topics:
        string(topic, "programme.topics entry", 120)
    if "deadline" in programme:
        programme["deadline"] = date.fromisoformat(str(programme["deadline"]))
    for field in ("degree_classes", "exam_sectors"):
        groups = config.get(field)
        require(isinstance(groups, list) and bool(groups), f"{field} must be a non-empty list")
        names, codes = set(), set()
        for group in groups:
            keys(group, {"name", "description", "codes"}, field)
            name = string(group.get("name"), f"{field}.name", 120)
            require(
                name.casefold() not in {"other", "unknown", "not applicable"},
                f"Reserved cluster: {name}",
            )
            require(name.casefold() not in names, f"Duplicate cluster: {name}")
            names.add(name.casefold())
            string(group.get("description"), f"{field}.description")
            require(
                isinstance(group.get("codes"), list) and bool(group["codes"]),
                f"{name}: codes must not be empty",
            )
            for code in group["codes"]:
                keys(code, {"code", "en", "it"}, f"{name}.codes")
                for part in ("code", "en", "it"):
                    string(code.get(part), f"{name}.codes.{part}", 250)
                require(
                    code["code"].casefold() not in codes,
                    f"Code belongs to multiple clusters: {code['code']}",
                )
                codes.add(code["code"].casefold())
    colors = config.get("colors", {})
    keys(colors, {"palette", "tabs"}, "colors")
    keys(colors.get("palette", {}), set(PALETTE), "colors.palette")
    layout = openpyxl.load_workbook(LAYOUT, read_only=True)
    keys(colors.get("tabs", {}), set(layout.sheetnames), "colors.tabs")
    layout.close()
    for section in colors.values():
        for name, color in section.items():
            require(
                isinstance(color, str) and re.fullmatch(r"#[0-9a-fA-F]{6}", color),
                f"Invalid color {name}: use '#RRGGBB'",
            )
    return config


def set_clusters(workbook, sheet_name, field, legend_name, groups):
    sheet = workbook[sheet_name]
    column = next(c.column_letter for c in sheet[9] if c.value == f"{field} *")
    validation = next(d for d in sheet.data_validations.dataValidation if f"{column}10" in d)
    name = validation.formula1
    destination, area = next(workbook.defined_names[name].destinations)
    values = workbook[destination]
    old_range = values[area]
    first = old_range[0][0]
    for row in old_range:
        row[0].value = None
    choices = [g["name"] for g in groups] + ["Other"]
    for index, choice in enumerate(choices):
        cell = values.cell(first.row + index, first.column, choice)
        cell.data_type = "s"
    new_range = f"'{destination}'!${first.column_letter}${first.row}:${first.column_letter}${first.row + len(choices) - 1}"
    workbook.defined_names[name] = DefinedName(name, attr_text=new_range)
    sheet[f"{column}7"] = choices[0]
    sheet[f"{column}7"].data_type = "s"

    legend = workbook[legend_name]
    # Keep the approved header/layout, then replace all programme-specific rows and source notes.
    style = [copy(c._style) for c in legend[7][:5]]
    for merged in list(legend.merged_cells.ranges):
        if merged.min_row >= 7:
            legend.unmerge_cells(str(merged))
    legend.delete_rows(7, legend.max_row - 6)
    rows = [
        [
            g["name"],
            ", ".join(c["code"] for c in g["codes"]),
            "; ".join(c["en"] for c in g["codes"]),
            "; ".join(c["it"] for c in g["codes"]),
            g["description"],
        ]
        for g in groups
    ]
    rows.append(
        ["Other", "", "", "", "No listed cluster fits. Explain the subject in Other details."]
    )
    for index, row in enumerate(rows, 7):
        for col, value in enumerate(row, 1):
            cell = legend.cell(index, col, value)
            cell.data_type = "s"
            cell._style = copy(style[col - 1])
        legend.row_dimensions[index].height = max(
            28,
            max(
                ceil(
                    len(str(value))
                    / max(
                        1,
                        legend.column_dimensions[openpyxl.utils.get_column_letter(col)].width
                        * 1.05,
                    )
                )
                * 13
                + 6
                for col, value in enumerate(row, 1)
            ),
        )
    legend.merge_cells(
        start_row=len(rows) + 9, start_column=1, end_row=len(rows) + 10, end_column=5
    )
    legend.cell(
        len(rows) + 9,
        1,
        "Source: this programme's YAML configuration. Applicant classifications are subject to committee review.",
    )
    legend.cell(len(rows) + 9, 1).alignment = Alignment(wrap_text=True)


def recolor(workbook, colors):
    replacements = {
        old: colors.get("palette", {}).get(role, "#" + old)[1:].upper()
        for role, old in PALETTE.items()
    }

    def change(color):
        if color is not None and color.type == "rgb" and color.rgb[-6:] in replacements:
            color.rgb = "FF" + replacements[color.rgb[-6:]]

    for sheet in workbook:
        if sheet.title in colors.get("tabs", {}):
            sheet.sheet_properties.tabColor = colors["tabs"][sheet.title][1:]
        for row in sheet:
            for cell in row:
                fill, font, border = copy(cell.fill), copy(cell.font), copy(cell.border)
                change(fill.fgColor)
                change(fill.bgColor)
                change(font.color)
                for side in ("left", "right", "top", "bottom", "start", "end", "diagonal"):
                    edge = getattr(border, side)
                    if edge:
                        change(edge.color)
                cell.fill, cell.font, cell.border = fill, font, border
        for cf in sheet.conditional_formatting:
            for rule in sheet.conditional_formatting[cf]:
                if rule.dxf:
                    rule.dxf = copy(rule.dxf)
                    if rule.dxf.fill:
                        rule.dxf.fill = copy(rule.dxf.fill)
                        change(rule.dxf.fill.fgColor)
                        change(rule.dxf.fill.bgColor)
                    if rule.dxf.font:
                        rule.dxf.font = copy(rule.dxf.font)
                        change(rule.dxf.font.color)


def build(config_path, output_dir):
    config = load_config(config_path)
    workbook = openpyxl.load_workbook(LAYOUT)
    programme = config["programme"]
    workbook.properties.title = f"{programme['name']} — applicant evidence"
    instructions = workbook["Instructions"]
    instructions["A2"] = programme["name"]
    instructions["A2"].data_type = "s"
    details = [
        programme.get(k) for k in ("institution", "academic_year", "url") if programme.get(k)
    ]
    instructions["A3"] = ". ".join(
        details + ["Complete this workbook and submit the supporting documents separately."]
    )
    instructions["A3"].data_type = "s"
    if "deadline" in programme:
        workbook["Personal & contacts"]["B34"] = programme["deadline"]
        workbook["Personal & contacts"]["B34"].number_format = "yyyy-mm-dd"
    topics = ", ".join(programme["topics"])
    for sheet in workbook:
        for row in sheet:
            for cell in row:
                if cell.data_type != "f" and isinstance(cell.value, str):
                    cell.value = cell.value.replace(
                        "computing, business or digital transformation", topics
                    )
        for dv in sheet.data_validations.dataValidation:
            if dv.prompt:
                dv.prompt = dv.prompt.replace(
                    "computing, business or digital transformation", topics
                )[:255]
        if sheet.title in {
            "Work experience",
            "Training",
            "Other activities",
            "Internationalization",
        }:
            for cell in sheet[8]:
                if isinstance(cell.value, str):
                    width = sheet.column_dimensions[cell.column_letter].width
                    sheet.row_dimensions[8].height = max(
                        sheet.row_dimensions[8].height or 44,
                        ceil(len(cell.value) / (width * 1.1)) * 10 + 6,
                    )
    set_clusters(
        workbook, "Degrees", "Degree cluster", "Degree class legend", config["degree_classes"]
    )
    set_clusters(workbook, "Exams", "Exam cluster", "Exam sector legend", config["exam_sectors"])
    recolor(workbook, config.get("colors", {}))
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"{programme['slug']}-applicant-template.xlsx"
    # Save atomically so a failed generation never replaces a previous complete workbook.
    with tempfile.TemporaryDirectory(dir=output_dir) as staging:
        temp = Path(staging) / target.name
        workbook.save(temp)
        temp.replace(target)
    return target


def convert_ods(xlsx, executable=None):
    executable = executable or shutil.which("libreoffice") or shutil.which("soffice")
    if executable is None and Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists():
        executable = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    require(
        executable is not None,
        "LibreOffice is required for --ods; install it or use --libreoffice PATH",
    )
    xlsx = Path(xlsx).resolve()
    # An isolated profile avoids interfering with the user's open LibreOffice session.
    with (
        tempfile.TemporaryDirectory() as profile,
        tempfile.TemporaryDirectory(dir=xlsx.parent) as staging,
    ):
        subprocess.run(
            [
                str(executable),
                f"-env:UserInstallation={Path(profile).as_uri()}",
                "--headless",
                "--convert-to",
                "ods:calc8",
                "--outdir",
                staging,
                str(xlsx),
            ],
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
        result = Path(staging) / xlsx.with_suffix(".ods").name
        require(result.exists(), "LibreOffice returned without producing an ODS file")
        verify_ods(result)
        target = xlsx.with_suffix(".ods")
        result.replace(target)
        return target


def verify_ods(path):
    ns = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0"}
    with zipfile.ZipFile(path) as archive:
        require(
            archive.read("mimetype").decode() == "application/vnd.oasis.opendocument.spreadsheet",
            "Invalid ODS mimetype",
        )
        root = ET.fromstring(archive.read("content.xml"))
    tables = root.findall(".//table:table", ns)
    require(len(tables) == 12, "ODS lost worksheets")
    protected = f"{{{ns['table']}}}protected"
    require(all(t.get(protected) == "true" for t in tables), "ODS lost sheet protection")
    require(bool(root.findall(".//table:content-validation", ns)), "ODS lost input validation")
    require(
        any(c.get(f"{{{ns['table']}}}formula") for c in root.findall(".//table:table-cell", ns)),
        "ODS lost formulas",
    )
    for cell in root.findall(".//table:table-cell", ns):
        require(
            "error" not in [v for k, v in cell.attrib.items() if k.endswith("}value-type")],
            "ODS contains a formula error",
        )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("configs", nargs="+", type=Path, help="One or more programme YAML files")
    parser.add_argument("--output-dir", type=Path, default=Path("dist"))
    parser.add_argument("--ods", action="store_true", help="Also produce ODS using LibreOffice")
    parser.add_argument("--libreoffice", help="Path to the LibreOffice executable")
    args = parser.parse_args()
    configs = [load_config(path) for path in args.configs]
    slugs = [c["programme"]["slug"] for c in configs]
    require(len(slugs) == len(set(slugs)), "Programme slugs must be unique within a build")
    for config in args.configs:
        xlsx = build(config, args.output_dir)
        print(xlsx)
        if args.ods:
            print(convert_ods(xlsx, args.libreoffice))


if __name__ == "__main__":
    main()
