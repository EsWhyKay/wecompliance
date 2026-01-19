# WeCompliance 
- A website for personnal use.
## What this package does
- Uses `data.xlsx` as the single source of truth.
- Generates **one HTML page per worksheet** (e.g., `BC99.html`).
- Provides `index.html` as a landing page.
- Countries are grouped under regions: **亚洲 / 美洲 / 欧洲 / 大洋洲** (fallback **其他**).
- In the 4 compliance columns (**食品 / 保健品/膳食补充剂 / 药品 / 动物食品**), if the cell value **starts with `Y` or `N`**, the page will render:
  - `Y...` as a **green check SVG** + the remaining text
  - `N...` as a **red cross SVG** + the remaining text

## Update workflow
1. Replace `data.xlsx` (keep the same column headers).
2. Run:
   ```bash
   python generate.py
   ```
3. Open `index.html`.

## Column requirements (per sheet)
Required headers:
- `Country`, `食品`, `f_link`, `保健品/膳食补充剂`, `s_link`, `药品`, `d_link`, `动物食品`, `fe_link`

Optional:
- `Region` (if missing or blank, generator uses `region_map.csv`, else falls back to `其他`).

