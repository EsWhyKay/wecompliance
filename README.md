# BC99 – Compliance Advice Around the World (Single-page Accordion)

## What you get
- `index.html`: the generated, searchable, single-page website (accordion list by country)
- `data.xlsx`: the **replaceable** Excel source (keep the same sheet name & columns)
- `generate.py`: generator that reads `data.xlsx`, resolves link titles, and regenerates `index.html`
- `assets/`: placeholders (team photo placeholder SVG)

## Update workflow (recommended)
1. Replace `data.xlsx` with your updated version (keep the same columns).
2. Run:
   - `python generate.py`
3. Open `index.html`.

## Notes
- Link titles are fetched via HTTP from the page `<title>` tag when you run `generate.py`.
- If fetching fails (network/CORS/blocked), it falls back to the domain name.
- Titles are cached in `link_titles.json` to speed up repeated runs.

## Excel format requirements
Sheet: `BC99` (or the first sheet)

Columns (exact):
- `Country`
- `食品`, `f_link`
- `保健品/膳食补充剂`, `s_link`
- `药品`, `d_link`
- `动物食品`, `fe_link`
