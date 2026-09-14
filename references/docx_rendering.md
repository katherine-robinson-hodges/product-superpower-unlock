# Rendering the Product Superpower Canvas (docx)

Fill in `assets/product_superpower_canvas_template.docx` (header fields + all
four quadrant boxes) using python-docx (preinstalled — open with
`docx.Document()`, edit paragraph/cell text, save; no need to unzip/edit raw
XML for simple text insertion like this).

**Known template quirk — fix this every time:** earlier versions of the
template had the two quadrant rows in the second table carrying a fixed
minimum row height (`trHeight` ~3120 twips), left over from when the
template was meant to be filled in by hand — this pushed the second row onto
a page 2 even with trimmed content, because the table couldn't shrink below
it. The current template no longer sets this, but apply the fix defensively
anyway (it's a no-op if there's nothing to override, and protects against a
future template revision reintroducing it):

```python
from docx.enum.table import WD_ROW_HEIGHT_RULE
for row in quad_table.rows:
    row.height_rule = WD_ROW_HEIGHT_RULE.AUTO
    row.height = None
```

Also delete any leftover unused blank paragraphs in a cell after filling it
(don't leave empty paragraphs past your last bullet) — they add vertical
space for nothing.

**Page-count check — keep this lightweight, don't chase it:**
The bullet budget in the main skill (max 3 bullets/quadrant, ~12-15 words
each) is *designed* to fit the template on one page. Trust it as the default
guarantee rather than treating rendering as mandatory.

- Only attempt an automated check if a CLI converter is already installed
  and available without launching a GUI app: `pandoc file.docx -o file.pdf`
  or `soffice --headless --convert-to pdf file.docx`. If one of these exists,
  run it, then check with `pdfinfo file.pdf | grep Pages` (or count pages via
  `python3 -c "import fitz; print(fitz.open('file.pdf').page_count)"` if
  PyMuPDF is available).
- If neither is installed, **don't install anything and don't try to render
  it another way** — skip the check and rely on the bullet budget. Do a
  quick manual sanity pass instead: total bullets across all four quadrants
  should be ≤ 10-12, and no single bullet should need wrapping to 3+ lines.
- Never automate Microsoft Word, Pages, or any other GUI application (e.g.
  via AppleScript/`osascript`) to open the file, count pages, or export a
  PDF. It's slow, flaky across Word versions, pops up visible app windows,
  and turns a one-page sanity check into a multi-step debugging session.

If a render *does* show more than one page, don't shrink margins/fonts —
trim content instead: cut each quadrant to at most 2 bullets, or shorten
bullets to a single short clause (~8-10 words), and re-check once. If still
over a page, pick the single highest-leverage bullet per quadrant and drop
the rest rather than retrying indefinitely.

Save the completed file to a location the user can find it with zero extra
steps. Default to the OS's standard `~/Downloads` folder if it already
exists — no `mkdir`, no path decision, no faff. If `~/Downloads` doesn't
exist, fall back to the current working directory or the session's
scratchpad; don't create `~/Downloads` yourself and don't use a
project-specific hardcoded path like `/mnt/user-data/outputs/`, which doesn't
exist on every platform. Present it alongside the headline recommendation and
a brief conversational summary (don't just hand over the file with no
context).
