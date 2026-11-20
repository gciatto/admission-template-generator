# Repository instructions

This repository builds blank applicant evidence workbooks for master's admissions.

- Read README.md before changing the generation pipeline.
- The supported generator is `generate.py`. Programme definitions belong in
  `programmes/*.yml` or `programmes/*.yaml`, not in Python or the layout workbook.
- `layout/applicant-template.xlsx` is the approved blank layout, not an output.
  Preserve its input fields, table borders, ISO date formats, validations, formula
  IDs, hidden dropdown sheet, and sheet/workbook protection.
- Never put real applicant data, credentials, or filled workbooks in this repo.
- Yellow inherited defaults must stay editable. IDs, lookups, durations and checks
  must stay locked. Defaults must stay blank until the row title is entered,
  preserve numeric zero, and stop copying exam grading settings when Degree ID changes.
- Degree and exam selectors contain cluster names and Other, never individual
  codes, Unknown or Not applicable. Legends remain visible, read-only, without
  Excel tables or filter buttons.
- Changes must work with public Python dependencies on GitHub-hosted Ubuntu.
- Change the layout through `tools/build_layout.py`, then `uv run poe layout`;
  never edit the layout workbook by hand. `tests/test_layout.py` must stay green.
- Keep fixes small. Add a focused unittest for nontrivial behaviour changes.
- Manage dependencies with uv (`uv add`, `uv add --dev`); never hand-edit `uv.lock`.
- Run `uv run poe static-checks`, `uv run poe test` and `uv run poe generate`
  before delivering changes.
  If LibreOffice is unavailable, report that ODS checks were skipped.
- Inspect affected workbook views when changing layout or colours. Do not treat
  sheet protection as encryption or Excel validation as proof that a claim is true.
- Do not push or create a GitHub repository unless the user asks.
