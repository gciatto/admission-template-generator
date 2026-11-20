# Admission template generator

Generate protected applicant evidence workbooks for different master's admission
procedures. Each programme YAML supplies its name, topics of interest, relevant
degree-class and exam-sector clusters, and optional colours. Builds produce one
`.xlsx` and, with LibreOffice, one `.ods` per programme.

## Build

Install [uv](https://docs.astral.sh/uv/getting-started/installation/) (it provides
Python 3.11+) and LibreOffice (Calc). On Ubuntu:

```sh
sudo apt-get install libreoffice-calc
uv sync
uv run poe generate        # = python generate.py programmes/*.y*ml --ods
```

On macOS the generator also detects
`/Applications/LibreOffice.app/Contents/MacOS/soffice`. Otherwise use
`--libreoffice /path/to/soffice`. Use `uv run poe generate-xlsx` (or omit `--ods`)
to generate only Excel files.
Use `--output-dir PATH` to change the output directory (default: `dist/`).
Existing outputs of the same name are replaced atomically; do not use the output
directory to store completed applications.

```sh
uv run poe test            # pytest (also runs the unittest suites)
uv run poe static-checks   # compileall, ruff format --check, ruff check, mypy
uv run poe static-fixes    # ruff --fix + ruff format
uv run poe install-hooks   # optional local Git hooks (scripts/githooks.py)
```

Manage dependencies with `uv add` / `uv add --dev`; never edit `uv.lock` by hand.
Renovate (`renovate.json`) opens dependency update PRs once enabled on the repository.

GitHub Actions tests and builds **all** programme `.yml` / `.yaml` files on pushes,
pull requests and manual runs. Download `applicant-templates` from the run's
Artifacts section. This local repository has no remote until one is configured.

## Configure a programme

Copy `programmes/digital-transformation-management.yml` and edit it. Use a unique
lowercase, hyphenated `programme.slug`: it determines output filenames.

```yaml
programme:
  slug: example-master
  name: Example Master Programme
  institution: Example University
  academic_year: '2027/28'       # optional
  url: https://example.org      # optional
  deadline: '2027-05-20'        # optional, prefills the editable reference date
  topics: [data analysis, statistics]

degree_classes:
  - name: Statistics
    description: Probability, statistical inference, data analysis and modelling.
    codes:
      - code: L-41
        en: Statistics
        it: Statistica

exam_sectors:
  - name: Computer science
    description: Programming, algorithms, software and computing systems.
    codes:
      - code: INF/01
        en: Computer science
        it: Informatica

colors:
  palette:
    input: '#FFF2A6'
    guidance: '#DBEAFE'
    example: '#FCE7F3'
    header: '#23374D'
  tabs:
    Degrees: '#7C3AED'
    Exams: '#9333EA'
```

Supported programme properties: `slug`, `name`, `institution`, `academic_year`,
`url`, `deadline`, `topics`. The full default YAML lists every available palette
role and tab name. Omitted colour roles preserve the approved layout's colours.
Quote hex colours so YAML does not interpret them as comments. Choose readable
foreground/background combinations; colour contrast is not automatically assessed.

Each classification group has a unique `name`, a plain-language `description`,
and a list of `{code, en, it}` records. A code cannot belong to two clusters within
the same classification. **Other** is added automatically. Codes are displayed in
the legends; dropdowns contain only cluster names. Applicants with a stated code
select its corresponding cluster. Applicants without a code choose the closest
description. The committee reviews both cases.

The default classifications reproduce the original committee rubric, including
its SSD codes; this is not an exhaustive or automatically updated national
catalogue. Review the YAML for each admission procedure. Do not assume a code
or eligibility rule remains appropriate for another year or programme.

Programme topics replace the relevance guidance across activity sheets. Other
form fields and capacities remain shared: 8 degrees, 200 exam attempts, 20 language
records and 30 activities per category. This is configuration of content, not a
general form designer or automatic admission scoring system.

## Layout and provenance

`layout/applicant-template.xlsx` is the blank, approved layout. The supported
generator edits that template using public dependencies and preserves its formulas,
validations and protection. Generated outputs are ignored.

`tools/build_layout.py` builds that layout from code (sheets, columns, formulas,
validations, styles and protection; country names are in `tools/countries.json`).
It is a Python port of the original JavaScript builder. `tests/test_layout.py`
checks that the builder reproduces the committed layout exactly, and, with
LibreOffice, recalculates sample inputs to check IDs, durations, cross-sheet
lookups and carried-forward defaults. To change the layout, edit the builder, run
`uv run poe layout`, inspect the workbook, and commit both files.

Yellow defaults in later rows are editable formulas linked to the previous row.
Changing an earlier value updates later linked defaults until an applicant
overwrites them. Grading defaults reset when an exam's Degree ID changes. Unique
claims (names, dates, marks, experiences and evidence filenames) are never copied.

LibreOffice uses an isolated temporary profile for conversion. The generator
checks that ODS retains sheets, protection, validation and formulas; the integration
test also checks recalculation of a carried-forward grading bound. XLSX table
objects may be represented differently in Calc. Sheet protection prevents accidental
edits, not disclosure, and applicants must still check carried-forward values.

Implementation references:
[openpyxl named ranges](https://openpyxl.readthedocs.io/en/latest/defined_names.html),
[validation](https://openpyxl.readthedocs.io/en/stable/validation.html),
[LibreOffice conversion](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html),
[GitHub artifact uploads](https://github.com/actions/upload-artifact).
