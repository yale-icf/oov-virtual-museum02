"""
OOV Corpus Full-Text Search
Flask app on port 5051.

Features:
  - FTS5 full-text search with Porter stemming (phrase search, boolean ops)
  - Filter by document type, period, issuing country, language
  - Year range filter
  - KWIC snippets with highlighted matches
  - Hits-per-year bar chart
  - Each result links to the museum viewer

Usage:
    python scripts/corpus_search.py
    # open http://localhost:5051
"""

import io
import json
import os
import sqlite3

from flask import Flask, g, jsonify, render_template_string, request, send_file, send_from_directory
from PIL import Image

SITE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(SITE_DIR, "scripts", "corpus.db")
THUMBS_DIR = os.path.join(SITE_DIR, "thumbnails")
JPEG_DIR = r"C:\Users\ks2479\Documents\my-project\origins-of-value\JPEG Files"

app = Flask(__name__)

# Build a lookup dict {stem_lowercase: full_path} from the JPEG Files directory
def _build_image_index():
    idx = {}
    if not os.path.isdir(JPEG_DIR):
        return idx
    for root, _dirs, files in os.walk(JPEG_DIR):
        for fname in files:
            if fname.lower().endswith(".jpg"):
                stem = os.path.splitext(fname)[0].lower()
                idx[stem] = os.path.join(root, fname)
    return idx

IMAGE_INDEX = _build_image_index()

# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH, check_same_thread=False)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA query_only = ON")
    return g.db


@app.teardown_appcontext
def close_db(exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def get_stats():
    db = get_db()
    rows = db.execute("SELECT key, value FROM corpus_stats").fetchall()
    return {r["key"]: r["value"] for r in rows}


def get_filter_options():
    """Return sorted distinct values for each filter dimension."""
    db = get_db()
    def distinct(col):
        rows = db.execute(
            f"SELECT DISTINCT {col} FROM docs WHERE {col} != '' ORDER BY {col}"
        ).fetchall()
        # Each cell may be comma-joined; collect unique atomic values
        values = set()
        for row in rows:
            for v in row[0].split(","):
                v = v.strip()
                if v:
                    values.add(v)
        return sorted(values)

    return {
        "types":     distinct("type"),
        "periods":   distinct("period"),
        "countries": distinct("issuing_country"),
        "languages": distinct("language"),
    }


def search(query, sel_types, sel_periods, sel_countries, sel_languages,
           min_year, max_year, limit=200):
    """
    Run FTS5 search and apply metadata filters.
    Returns (results, year_counts) where year_counts = {year: count}.
    """
    db = get_db()

    fts_query = query.strip()
    if not fts_query:
        # Browse all (no FTS constraint): join docs directly
        sql = """
            SELECT d.id, d.title, d.type, d.period,
                   d.issuing_country, d.subject_country,
                   d.issue_year, d.language, d.owner,
                   d.word_count,
                   substr(d.ocr_text, 1, 400) AS snippet
            FROM docs d
            WHERE 1=1
        """
        params = []
    else:
        sql = """
            SELECT d.id, d.title, d.type, d.period,
                   d.issuing_country, d.subject_country,
                   d.issue_year, d.language, d.owner,
                   d.word_count,
                   snippet(docs_fts, 0, '<mark>', '</mark>', ' … ', 32) AS snippet
            FROM docs_fts
            JOIN docs d ON docs_fts.rowid = d.rowid
            WHERE docs_fts MATCH ?
        """
        params = [fts_query]

    rows = db.execute(sql + " ORDER BY d.issue_year", params).fetchall()

    # Post-filter (Python-side; fast for 545 rows)
    results = []
    year_counts = {}

    for row in rows:
        r = dict(row)

        # Type filter (any selected type appears in comma-joined cell)
        if sel_types:
            doc_types = {t.strip() for t in r["type"].split(",")}
            if not doc_types.intersection(sel_types):
                continue

        # Period filter
        if sel_periods:
            doc_periods = {p.strip() for p in r["period"].split(",")}
            if not doc_periods.intersection(sel_periods):
                continue

        # Issuing country filter
        if sel_countries:
            doc_countries = {c.strip() for c in r["issuing_country"].split(",")}
            if not doc_countries.intersection(sel_countries):
                continue

        # Language filter
        if sel_languages:
            doc_langs = {l.strip() for l in r["language"].split(",")}
            if not doc_langs.intersection(sel_languages):
                continue

        # Year range
        yr = r.get("issue_year")
        if min_year and yr and yr < min_year:
            continue
        if max_year and yr and yr > max_year:
            continue

        results.append(r)
        if yr:
            year_counts[yr] = year_counts.get(yr, 0) + 1

    return results[:limit], year_counts


# ---------------------------------------------------------------------------
# Year chart (inline SVG)
# ---------------------------------------------------------------------------

def make_year_chart(year_counts, width=700, height=120):
    """Return an SVG string: a bar chart of hits per year."""
    if not year_counts:
        return ""

    years = sorted(year_counts)
    counts = [year_counts[y] for y in years]
    max_count = max(counts) or 1
    n = len(years)

    pad_left, pad_right, pad_top, pad_bot = 5, 5, 8, 20
    chart_w = width - pad_left - pad_right
    chart_h = height - pad_top - pad_bot

    bar_w = max(1, chart_w / n - 1)

    bars = []
    labels = []
    for i, (yr, cnt) in enumerate(zip(years, counts)):
        x = pad_left + i * (chart_w / n)
        bar_h = cnt / max_count * chart_h
        y = pad_top + chart_h - bar_h
        bars.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" '
            f'height="{bar_h:.1f}" fill="#8b5e3c" opacity="0.75">'
            f'<title>{yr}: {cnt} hit{"s" if cnt != 1 else ""}</title></rect>'
        )
        # Label every ~50 years or if few bars
        if n <= 20 or yr % 50 == 0:
            lx = x + bar_w / 2
            ly = height - 2
            labels.append(
                f'<text x="{lx:.1f}" y="{ly}" text-anchor="middle" '
                f'font-size="9" fill="#5a3e2b">{yr}</text>'
            )

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{width}" height="{height}" style="display:block">'
        + "".join(bars)
        + "".join(labels)
        + "</svg>"
    )
    return svg


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>OOV Corpus Search</title>
<style>
  *, *::before, *::after { box-sizing: border-box; }

  body {
    font-family: Georgia, 'Times New Roman', serif;
    background: #f5efe0;
    color: #2c1a0e;
    margin: 0;
    padding: 0;
  }

  header {
    background: #3b2008;
    color: #f5efe0;
    padding: 14px 24px;
  }
  header h1 { margin: 0; font-size: 1.3rem; font-weight: normal; letter-spacing: .04em; }
  header small { opacity: .7; font-size: .8rem; }

  .layout {
    display: flex;
    gap: 0;
    min-height: calc(100vh - 56px);
  }

  /* ---- sidebar ---- */
  aside {
    width: 230px;
    flex-shrink: 0;
    background: #ece3ce;
    border-right: 1px solid #c4a87a;
    padding: 16px 14px;
    font-size: .85rem;
  }
  aside h2 {
    font-size: .9rem;
    text-transform: uppercase;
    letter-spacing: .06em;
    color: #5a3e2b;
    margin: 0 0 10px;
    border-bottom: 1px solid #c4a87a;
    padding-bottom: 4px;
  }
  .filter-group { margin-bottom: 16px; }
  .filter-group h3 {
    font-size: .8rem;
    text-transform: uppercase;
    letter-spacing: .05em;
    color: #7a5533;
    margin: 0 0 6px;
  }
  .filter-group label { display: block; margin: 2px 0; cursor: pointer; }
  .filter-group input[type=checkbox] { margin-right: 4px; }
  .year-row { display: flex; gap: 6px; align-items: center; }
  .year-row input[type=number] {
    width: 64px;
    padding: 3px 5px;
    border: 1px solid #c4a87a;
    background: #fdf8ef;
    color: #2c1a0e;
    border-radius: 3px;
    font-size: .82rem;
  }

  /* ---- main ---- */
  main {
    flex: 1;
    padding: 20px 24px;
    max-width: 900px;
  }

  .search-bar {
    display: flex;
    gap: 8px;
    margin-bottom: 10px;
  }
  .search-bar input[type=text] {
    flex: 1;
    padding: 8px 12px;
    border: 1px solid #c4a87a;
    background: #fdf8ef;
    color: #2c1a0e;
    border-radius: 4px;
    font-size: 1rem;
    font-family: inherit;
  }
  .search-bar button, .btn {
    padding: 8px 18px;
    background: #3b2008;
    color: #f5efe0;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: .9rem;
    font-family: inherit;
  }
  .search-bar button:hover, .btn:hover { background: #5a3e2b; }

  .result-meta { color: #7a5533; font-size: .8rem; margin-bottom: 4px; }
  .result-count { color: #7a5533; font-size: .85rem; margin-bottom: 14px; }

  .chart-box {
    background: #ece3ce;
    border: 1px solid #c4a87a;
    border-radius: 4px;
    padding: 10px 14px 6px;
    margin-bottom: 18px;
  }
  .chart-box h3 { margin: 0 0 6px; font-size: .8rem; color: #5a3e2b; text-transform: uppercase; letter-spacing: .04em; }

  .results { list-style: none; margin: 0; padding: 0; }
  .result-card {
    background: #fdf8ef;
    border: 1px solid #c4a87a;
    border-radius: 4px;
    padding: 12px 16px;
    margin-bottom: 12px;
    display: flex;
    gap: 14px;
    align-items: flex-start;
  }
  .result-thumb {
    flex-shrink: 0;
    width: 110px;
    border: 1px solid #c4a87a;
    border-radius: 3px;
    cursor: zoom-in;
    display: block;
  }
  .result-body { flex: 1; min-width: 0; }
  .result-card h4 {
    margin: 0 0 4px;
    font-size: .95rem;
    color: #3b2008;
  }
  .snippet {
    font-size: .85rem;
    color: #3e2a14;
    margin: 6px 0 0;
    line-height: 1.5;
    border-left: 3px solid #c4a87a;
    padding-left: 10px;
  }
  mark {
    background: #f0c040;
    color: #2c1a0e;
    padding: 0 1px;
    border-radius: 2px;
  }
  .no-ocr { color: #9e8060; font-style: italic; font-size: .82rem; }

  /* Lightbox */
  #lb { display:none; position:fixed; inset:0; background:rgba(0,0,0,.88);
        z-index:999; flex-direction:column; align-items:center; justify-content:center; }
  #lb.open { display:flex; }
  #lb-content { position:relative; display:inline-block; line-height:0; }
  #lb-img { max-width:92vw; max-height:84vh; object-fit:contain;
            border:2px solid #c4a87a; border-radius:4px; display:block; }
  #lb-canvas { position:absolute; top:0; left:0; pointer-events:none;
               border-radius:4px; }
  #lb-close { position:absolute; top:12px; right:18px; color:#f5efe0; font-size:1.6rem;
               cursor:pointer; line-height:1; background:none; border:none; }
  #lb-close:hover { color:#f0c040; }
  #lb-footer { margin-top:8px; text-align:center; }
  #lb-footer a { color:#f0c040; font-size:.85rem; text-decoration:none; }
  #lb-footer a:hover { text-decoration:underline; }

  .hint {
    font-size: .78rem;
    color: #9e8060;
    margin-top: 6px;
  }
</style>
</head>
<body>
<header>
  <h1>Origins of Value — Corpus Search</h1>
  <small>{{ stats.total_docs }} documents · {{ stats.total_words }} words indexed
    · built {{ stats.built_at[:10] if stats.built_at else '?' }}</small>
</header>

<form method="get" action="/search">
<div class="layout">

  <!-- Sidebar filters -->
  <aside>
    <h2>Filters</h2>

    {% if options.periods %}
    <div class="filter-group">
      <h3>Period</h3>
      {% for p in options.periods %}
      <label>
        <input type="checkbox" name="period" value="{{ p }}"
          {{ 'checked' if p in sel_periods }}>{{ p }}
      </label>
      {% endfor %}
    </div>
    {% endif %}

    {% if options.types %}
    <div class="filter-group">
      <h3>Document Type</h3>
      {% for t in options.types %}
      <label>
        <input type="checkbox" name="type" value="{{ t }}"
          {{ 'checked' if t in sel_types }}>{{ t }}
      </label>
      {% endfor %}
    </div>
    {% endif %}

    {% if options.countries %}
    <div class="filter-group">
      <h3>Issuing Country</h3>
      {% for c in options.countries %}
      <label>
        <input type="checkbox" name="country" value="{{ c }}"
          {{ 'checked' if c in sel_countries }}>{{ c }}
      </label>
      {% endfor %}
    </div>
    {% endif %}

    {% if options.languages %}
    <div class="filter-group">
      <h3>Language</h3>
      {% for l in options.languages %}
      <label>
        <input type="checkbox" name="language" value="{{ l }}"
          {{ 'checked' if l in sel_languages }}>{{ l }}
      </label>
      {% endfor %}
    </div>
    {% endif %}

    <div class="filter-group">
      <h3>Issue Year</h3>
      <div class="year-row">
        <input type="number" name="min_year" placeholder="from"
          value="{{ min_year or '' }}" min="1000" max="2100">
        <span>–</span>
        <input type="number" name="max_year" placeholder="to"
          value="{{ max_year or '' }}" min="1000" max="2100">
      </div>
    </div>

    <button type="submit" class="btn" style="width:100%;margin-top:8px">Apply</button>
    <a href="/" style="display:block;text-align:center;margin-top:8px;font-size:.8rem;color:#7a5533">Reset all</a>
  </aside>

  <!-- Main content -->
  <main>
    <div class="search-bar">
      <input type="text" name="q" value="{{ query|e }}"
        placeholder='Search OCR text — try "dividend" or "bearer bond" or "société"'>
      <button type="submit">Search</button>
    </div>
    <p class="hint">
      Tip: use quotes for phrases (<code>"bearer bond"</code>), <code>*</code> for
      prefix (<code>franc*</code>), <code>AND</code>/<code>OR</code>/<code>NOT</code>
      for boolean. Porter stemming is on (search→searches→searching).
    </p>

    {% if searched %}

    {% if error %}
    <p style="color:#a00">Search error: {{ error }}</p>
    {% else %}

    <p class="result-count">
      {{ results|length }} result{{ 's' if results|length != 1 }}
      {% if query %} for <strong>{{ query|e }}</strong>{% endif %}
      {% if results|length == 200 %} (showing first 200){% endif %}
    </p>

    {% if chart_svg %}
    <div class="chart-box">
      <h3>Hits per Year</h3>
      {{ chart_svg|safe }}
    </div>
    {% endif %}

    <ul class="results">
    {% for r in results %}
    <li class="result-card">
      <img class="result-thumb" src="/image/{{ r.id }}"
           alt="{{ r.id }}" title="Click to open zoomable view"
           onclick="lb('{{ r.id }}','{{ query|e }}')" onerror="this.style.display='none'">
      <div class="result-body">
        <h4>{{ r.title or r.id }}</h4>
        <div class="result-meta">
          {{ r.id }}
          {% if r.type %} · {{ r.type }}{% endif %}
          {% if r.period %} · {{ r.period }}{% endif %}
          {% if r.issuing_country %} · {{ r.issuing_country }}{% endif %}
          {% if r.issue_year %} · {{ r.issue_year }}{% endif %}
          {% if r.language %} · {{ r.language }}{% endif %}
          {% if r.owner %} · {{ r.owner }}{% endif %}
        </div>
        {% if r.snippet %}
        <div class="snippet">{{ r.snippet|safe }}</div>
        {% else %}
        <div class="no-ocr">No OCR text available for this document.</div>
        {% endif %}
      </div>
    </li>
    {% endfor %}
    </ul>

    {% if not results %}
    <p style="color:#7a5533;font-style:italic">No results found.</p>
    {% endif %}

    {% endif %}{# end error #}
    {% endif %}{# end searched #}
  </main>

</div>
</form>
<div id="lb" onclick="if(event.target===this)closeLb()">
  <button id="lb-close" onclick="closeLb()">✕</button>
  <div id="lb-content">
    <img id="lb-img" src="" alt="" onload="onImgLoad()">
    <canvas id="lb-canvas"></canvas>
  </div>
  <div id="lb-footer"><a id="lb-link" href="#" target="_blank">Open in museum viewer ↗</a></div>
</div>
<script>
var _lbBoxes = null;
var _lbQuery = '';

function lb(docId, query) {
  _lbBoxes = null;
  _lbQuery = (query || '').trim();
  var img = document.getElementById('lb-img');
  img.src = '/fullres/' + docId;
  img.alt = docId;
  document.getElementById('lb-link').href = '../viewer.html?id=' + docId;
  document.getElementById('lb').classList.add('open');
  if (_lbQuery) {
    fetch('/boxes/' + docId)
      .then(function(r) { return r.json(); })
      .then(function(data) {
        _lbBoxes = data;
        drawHighlights();
      })
      .catch(function() {});
  }
}

function onImgLoad() {
  var canvas = document.getElementById('lb-canvas');
  var img = document.getElementById('lb-img');
  canvas.width = img.offsetWidth;
  canvas.height = img.offsetHeight;
  canvas.style.width = img.offsetWidth + 'px';
  canvas.style.height = img.offsetHeight + 'px';
  drawHighlights();
}

function drawHighlights() {
  var canvas = document.getElementById('lb-canvas');
  var ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (!_lbBoxes || !_lbQuery || !_lbBoxes.words) return;

  var terms = _lbQuery.toLowerCase().split(/\s+/).filter(function(t) { return t.length >= 2; });
  ctx.fillStyle = 'rgba(255, 200, 0, 0.45)';

  _lbBoxes.words.forEach(function(w) {
    var wt = w.text.toLowerCase();
    for (var i = 0; i < terms.length; i++) {
      if (wt.includes(terms[i])) {
        ctx.fillRect(
          w.x0 * canvas.width,
          w.y0 * canvas.height,
          (w.x1 - w.x0) * canvas.width,
          (w.y1 - w.y0) * canvas.height
        );
        break;
      }
    }
  });
}

function closeLb() {
  document.getElementById('lb').classList.remove('open');
  document.getElementById('lb-img').src = '';
  var ctx = document.getElementById('lb-canvas').getContext('2d');
  ctx.clearRect(0, 0, 1, 1);
  _lbBoxes = null;
}
document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeLb(); });
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/image/<doc_id>")
def serve_image(doc_id):
    return send_from_directory(THUMBS_DIR, f"{doc_id}.jpg")


@app.route("/fullres/<doc_id>")
def serve_fullres(doc_id):
    """Serve a ~1500px-wide JPEG from the local originals directory."""
    path = IMAGE_INDEX.get(doc_id.lower())
    if not path or not os.path.exists(path):
        # Fall back to thumbnail
        return send_from_directory(THUMBS_DIR, f"{doc_id}.jpg")
    img = Image.open(path)
    if img.width > 1500 or img.height > 1500:
        img.thumbnail((1500, 1500), Image.LANCZOS)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=82)
    buf.seek(0)
    return send_file(buf, mimetype="image/jpeg")


@app.route("/boxes/<doc_id>")
def serve_boxes(doc_id):
    """Return word bounding-box JSON for doc_id from the corpus DB."""
    db = get_db()
    row = db.execute(
        "SELECT words_json FROM boxes WHERE image_id = ?", (doc_id,)
    ).fetchone()
    if not row:
        return jsonify({"words": []})
    return jsonify({"words": json.loads(row["words_json"])})


@app.route("/")
def index():
    stats = get_stats()
    options = get_filter_options()
    return render_template_string(
        TEMPLATE,
        stats=stats,
        options=options,
        query="",
        sel_types=set(),
        sel_periods=set(),
        sel_countries=set(),
        sel_languages=set(),
        min_year=None,
        max_year=None,
        results=[],
        chart_svg="",
        searched=False,
        error=None,
    )


@app.route("/search")
def search_route():
    stats = get_stats()
    options = get_filter_options()

    query = request.args.get("q", "").strip()
    sel_types = set(request.args.getlist("type"))
    sel_periods = set(request.args.getlist("period"))
    sel_countries = set(request.args.getlist("country"))
    sel_languages = set(request.args.getlist("language"))

    def _int(key):
        v = request.args.get(key, "").strip()
        try:
            return int(v) if v else None
        except ValueError:
            return None

    min_year = _int("min_year")
    max_year = _int("max_year")

    results = []
    year_counts = {}
    error = None

    try:
        results, year_counts = search(
            query, sel_types, sel_periods, sel_countries, sel_languages,
            min_year, max_year
        )
    except Exception as e:
        error = str(e)

    chart_svg = make_year_chart(year_counts) if year_counts else ""

    return render_template_string(
        TEMPLATE,
        stats=stats,
        options=options,
        query=query,
        sel_types=sel_types,
        sel_periods=sel_periods,
        sel_countries=sel_countries,
        sel_languages=sel_languages,
        min_year=min_year,
        max_year=max_year,
        results=results,
        chart_svg=chart_svg,
        searched=True,
        error=error,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if not os.path.exists(DB_PATH):
        print(f"Error: corpus.db not found at {DB_PATH}")
        print("Run build_corpus_index.py first.")
        raise SystemExit(1)
    print(f"OOV Corpus Search → http://localhost:5051")
    app.run(host="127.0.0.1", port=5051, debug=False)
