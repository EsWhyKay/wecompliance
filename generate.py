import json
import sys
import re
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

try:
    import requests
    from bs4 import BeautifulSoup
except Exception:  # allow generation even if requests/bs4 not installed
    requests = None
    BeautifulSoup = None

BASE_DIR = Path(__file__).resolve().parent
DATA_XLSX = BASE_DIR / "data.xlsx"
OUT_HTML = BASE_DIR / "index.html"
CACHE_PATH = BASE_DIR / "link_titles.json"

TITLE = "BC99"
SUBTITLE = "Compliance Advice Around the World"

# By default, generation is offline-safe (no outbound requests).
# To fetch real webpage titles, run:  python generate.py --fetch-titles
ENABLE_FETCH_TITLES = "--fetch-titles" in sys.argv

COLUMNS = [
    "Country",
    "食品", "f_link",
    "保健品/膳食补充剂", "s_link",
    "药品", "d_link",
    "动物食品", "fe_link",
]


def clean_text(x) -> str:
    if x is None:
        return "-"
    if isinstance(x, float) and pd.isna(x):
        return "-"
    s = str(x).strip()
    return s if s else "-"


def clean_url(x) -> str:
    s = clean_text(x)
    if s == "-":
        return "-"
    # if multiple links are separated by newlines/spaces, keep first for now
    # (can be extended later)
    s = s.splitlines()[0].strip()
    return s


def fallback_title(url: str) -> str:
    try:
        host = urlparse(url).netloc
        host = re.sub(r"^www\.", "", host)
        return host or url
    except Exception:
        return url


def load_cache() -> dict:
    if CACHE_PATH.exists():
        try:
            return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_cache(cache: dict) -> None:
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_title(url: str, timeout=12) -> str:
    """Fetch <title> from a URL. If unavailable, returns a domain fallback."""
    if not ENABLE_FETCH_TITLES:
        return fallback_title(url)

    if requests is None or BeautifulSoup is None:
        return fallback_title(url)

    headers = {
        "User-Agent": "Mozilla/5.0 (BC99 Generator)",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    try:
        r = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        t = soup.title.string.strip() if soup.title and soup.title.string else ""
        if t:
            # collapse whitespace
            t = re.sub(r"\s+", " ", t)
            return t
    except Exception:
        pass
    return fallback_title(url)


def build_link_html(url: str, title: str) -> str:
    if url == "-":
        return "-"
    safe_title = html_escape(title) if title else html_escape(fallback_title(url))
    safe_url = html_escape(url)
    return f'<a href="{safe_url}" target="_blank" rel="noopener">↗ {safe_title}</a>'


def html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def nl2br(s: str) -> str:
    """Escape HTML and convert newlines to <br>."""
    return html_escape(s).replace("\n", "<br>")


def main():
    if not DATA_XLSX.exists():
        raise FileNotFoundError(f"Missing {DATA_XLSX}")

    xl = pd.ExcelFile(DATA_XLSX)
    sheet = "BC99" if "BC99" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(DATA_XLSX, sheet_name=sheet)

    # Ensure required columns
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Excel columns missing: " + ", ".join(missing) + "\n" +
            "Expected columns: " + ", ".join(COLUMNS)
        )

    df = df[COLUMNS].copy()

    cache = load_cache()

    # Resolve link titles
    link_cols = ["f_link", "s_link", "d_link", "fe_link"]
    titles = {col: [] for col in link_cols}

    for _, row in df.iterrows():
        for col in link_cols:
            url = clean_url(row[col])
            if url == "-":
                titles[col].append("-")
                continue

            if url in cache and isinstance(cache[url], str) and cache[url].strip():
                titles[col].append(cache[url])
                continue

            t = fetch_title(url)
            cache[url] = t
            titles[col].append(t)

    save_cache(cache)

    # Build accordion HTML blocks
    blocks = []
    for i, row in df.iterrows():
        country = html_escape(clean_text(row["Country"]))

        food = clean_text(row["食品"])
        supp = clean_text(row["保健品/膳食补充剂"])
        drug = clean_text(row["药品"])
        feed = clean_text(row["动物食品"])

        f_url = clean_url(row["f_link"])
        s_url = clean_url(row["s_link"])
        d_url = clean_url(row["d_link"])
        fe_url = clean_url(row["fe_link"])

        f_link = build_link_html(f_url, titles["f_link"][i])
        s_link = build_link_html(s_url, titles["s_link"][i])
        d_link = build_link_html(d_url, titles["d_link"][i])
        fe_link = build_link_html(fe_url, titles["fe_link"][i])

        block = f"""
<details>
  <summary>{country}</summary>
  <div class=\"content\">
    <div class=\"grid\">
      <div class=\"cell-title\">食品</div>
      <div class=\"cell-title\">保健品</div>
      <div class=\"cell-title\">药品</div>
      <div class=\"cell-title\">动物食品</div>

      <div class=\"cell\">{nl2br(food)}</div>
      <div class=\"cell\">{nl2br(supp)}</div>
      <div class=\"cell\">{nl2br(drug)}</div>
      <div class=\"cell\">{nl2br(feed)}</div>
    </div>

    <div class=\"links\">
      <div class=\"linkcell\">{f_link}</div>
      <div class=\"linkcell\">{s_link}</div>
      <div class=\"linkcell\">{d_link}</div>
      <div class=\"linkcell\">{fe_link}</div>
    </div>
  </div>
</details>
"""
        blocks.append(block)

    accordion_html = "\n".join(blocks)

    html = f"""<!DOCTYPE html>
<html lang=\"zh-CN\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>{TITLE}</title>
  <style>
    :root {{
      --bg-main: #f4f1ec;
      --bg-card: #fbfaf7;
      --bg-expand: #ffffff;
      --border-soft: #ddd6cc;
      --text-main: #2f2f2f;
      --text-muted: #7a746c;
      --accent: #8b7e6a;
      --link: #5f543f;
      --shadow: 0 6px 20px rgba(40, 30, 20, 0.06);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      background: var(--bg-main);
      color: var(--text-main);
    }}

    header {{
      text-align: center;
      padding: 34px 16px 18px;
    }}

    header h1 {{
      margin: 0;
      font-size: 38px;
      letter-spacing: 1px;
      color: var(--text-main);
    }}

    header h2 {{
      margin: 10px 0 0;
      font-size: 16px;
      font-weight: 500;
      color: var(--text-muted);
    }}

    .controls {{
      max-width: 1180px;
      margin: 0 auto 14px;
      padding: 0 16px;
      display: flex;
      gap: 12px;
      align-items: center;
    }}

    .controls input {{
      flex: 1;
      padding: 11px 12px;
      border: 1px solid var(--border-soft);
      border-radius: 10px;
      background: #fff;
      font-size: 14px;
      outline: none;
    }}

    .controls input:focus {{
      border-color: #cdbfae;
      box-shadow: 0 0 0 3px rgba(205, 191, 174, 0.25);
    }}

    .controls button {{
      padding: 10px 14px;
      border: 1px solid var(--border-soft);
      background: var(--bg-card);
      border-radius: 10px;
      cursor: pointer;
      color: var(--accent);
      font-weight: 600;
    }}

    .controls button:hover {{
      filter: brightness(0.99);
    }}

    .container {{
      max-width: 1180px;
      margin: 0 auto;
      padding: 0 16px 40px;
    }}

    details {{
      background: var(--bg-card);
      border: 1px solid var(--border-soft);
      border-radius: 14px;
      margin-bottom: 10px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }}

    summary {{
      padding: 14px 18px;
      font-size: 16px;
      cursor: pointer;
      list-style: none;
      font-weight: 700;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }}

    summary::-webkit-details-marker {{ display: none; }}

    summary .chev {{
      color: var(--text-muted);
      font-weight: 900;
      transform: rotate(0deg);
      transition: transform 0.15s ease;
    }}

    details[open] summary .chev {{
      transform: rotate(90deg);
    }}

    .content {{
      background: var(--bg-expand);
      border-top: 1px solid var(--border-soft);
      padding: 16px 18px;
    }}

    .grid {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
      font-size: 14px;
    }}

    .cell-title {{
      font-weight: 800;
      color: var(--accent);
    }}

    .cell {{
      line-height: 1.55;
      white-space: normal;
    }}

    .links {{
      margin-top: 12px;
      padding-top: 12px;
      border-top: 1px dashed var(--border-soft);
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
      font-size: 13px;
      color: var(--text-muted);
    }}

    a {{
      color: var(--link);
      text-decoration: none;
    }}

    a:hover {{
      text-decoration: underline;
    }}

    .linkcell {{
      min-height: 18px;
    }}

    .footer {{
      max-width: 1180px;
      margin: 20px auto 50px;
      padding: 0 16px;
      color: var(--text-muted);
    }}

    .team {{
      margin-top: 24px;
      background: var(--bg-card);
      border: 1px solid var(--border-soft);
      border-radius: 14px;
      box-shadow: var(--shadow);
      padding: 18px;
    }}

    .team h3 {{
      margin: 0 0 12px;
      color: var(--text-main);
      font-size: 16px;
    }}

    .team-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
    }}

    .member {{
      background: #fff;
      border: 1px solid var(--border-soft);
      border-radius: 14px;
      padding: 12px;
      display: grid;
      grid-template-columns: 54px 1fr;
      gap: 10px;
      align-items: center;
    }}

    .avatar {{
      width: 54px;
      height: 54px;
      border-radius: 14px;
      overflow: hidden;
      border: 1px solid var(--border-soft);
      background: #fff;
    }}

    .avatar img {{
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
    }}

    .member .name {{
      font-weight: 800;
      color: var(--text-main);
      margin-bottom: 3px;
    }}

    .member .role {{
      font-size: 13px;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 4px;
    }}

    .member .desc {{
      font-size: 13px;
      color: var(--text-muted);
      line-height: 1.4;
    }}

    .member .desc a {{
      color: var(--link);
      word-break: break-all;
    }}

    @media (max-width: 980px) {{
      .grid, .links {{ grid-template-columns: 1fr 1fr; }}
      .team-grid {{ grid-template-columns: 1fr; }}
      .controls {{ flex-wrap: wrap; }}
    }}
  </style>

  <script>
    function toggleAll(open) {{
      document.querySelectorAll('details').forEach(d => d.open = open);
    }}

    function filterList() {{
      const q = (document.getElementById('search').value || '').toLowerCase();
      document.querySelectorAll('details').forEach(d => {{
        const txt = d.innerText.toLowerCase();
        d.style.display = txt.includes(q) ? '' : 'none';
      }});
    }}
  </script>
</head>

<body>
  <header>
    <h1>{TITLE}</h1>
    <h2>{SUBTITLE}</h2>
  </header>

  <div class="controls">
    <input id="search" placeholder="Search country or content…" oninput="filterList()" />
    <button onclick="toggleAll(true)">Expand all</button>
    <button onclick="toggleAll(false)">Collapse all</button>
  </div>

  <div class="container">
    {accordion_html}
  </div>

  <div class="footer">
    <div class="team">
      <h3>Team</h3>
      <div class="team-grid">

        <div class="member">
          <div class="avatar">
            <img src="assets/photos/xiaowen.wang.jpg" onerror="this.onerror=null;this.src='assets/photo-placeholder.svg';" alt="Xiaowen Wang" />
          </div>
          <div>
            <div class="name">xiaowen.wang@wecare-bio.com</div>
            <div class="role">Consultant</div>
            <div class="desc">负责亚欧板块合规信息</div>
          </div>
        </div>

        <div class="member">
          <div class="avatar">
            <img src="assets/photos/yixuan.fan.jpg" onerror="this.onerror=null;this.src='assets/photo-placeholder.svg';" alt="Yixuan Fan" />
          </div>
          <div>
            <div class="name">yixuan.fan@wecare-bio.com</div>
            <div class="role">Consultant</div>
            <div class="desc">负责美洲板块合规信息</div>
          </div>
        </div>

        <div class="member">
          <div class="avatar">
            <img src="assets/photos/kay.sun.jpg" onerror="this.onerror=null;this.src='assets/photo-placeholder.svg';" alt="Yu-kun Sun" />
          </div>
          <div>
            <div class="name">kay.sun@wecare-life.com</div>
            <div class="role">Manager</div>
            <div class="desc">
              <a href="http://www.linkedin.com/in/yu-kun-sun" target="_blank" rel="noopener">LinkedIn</a>
            </div>
          </div>
        </div>

      </div>
      <div style="margin-top:10px;font-size:12px;">Photo placement: replace images under <code>assets/photos/</code> with the same filenames.</div>
    </div>
  </div>
</body>
</html>
"""

    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"Generated: {OUT_HTML}")


if __name__ == "__main__":
    main()
