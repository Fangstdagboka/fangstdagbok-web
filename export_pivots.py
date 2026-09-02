#!/usr/bin/env python3
"""
Export the five named pivot-table views (per municipality) from
Fangst 2026 v1.xlsx into static, self-contained HTML pages with
month -> day -> boat expand/collapse, replicating what the OneDrive
"Embed" iframes used to show -- but with columns sized to their
content and the first two (identifying) columns frozen while the
data columns scroll horizontally, like Excel's freeze panes.

Usage:
    python3 export_pivots.py <path-to-xlsx> <output-dir>
"""
import re
import sys
import html
from datetime import datetime

import openpyxl
import pandas as pd

MINUS = '−'  # matches the glyph the client-side JS uses for "expanded"

KVITFISK_SPECIES = ['Breiflabb', 'Hyse', 'Lange', 'Lyr', 'Sei', 'Torsk']
PELAGISK_SPECIES = ['Kolmule', 'Makrell', 'Sild', 'Øyepål', 'Sølvtorsk', 'Strømsild']

MONTH_ORDER = ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli',
               'August', 'September', 'Oktober', 'November', 'Desember']

MUNICIPALITIES = ['HERØY', 'SANDE', 'VANYLVEN']
PAGE_PREFIX = {'HERØY': 'H', 'SANDE': 'S', 'VANYLVEN': 'V'}


def load_sheet(wb, name):
    ws = wb[name]
    rows = ws.iter_rows(values_only=True)
    headers = next(rows)
    data = [r for r in rows if any(v is not None for v in r)]
    return pd.DataFrame(data, columns=headers)


def clean_month(s):
    return str(s).strip() if s is not None else ''


def fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return f"{int(round(v)):,}".replace(',', ' ')


def fmt_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, datetime):
        return v.strftime('%d.%m.%Y')
    return str(v)


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    try:
        return datetime.strptime(str(v).strip(), '%d.%m.%Y')
    except Exception:
        return None


def px_width(strings, base=22, per_char=7.2, minimum=44, maximum=260):
    """Estimate a sensible column pixel width from the text it will hold."""
    longest = max((len(str(s)) for s in strings if s), default=1)
    return int(max(minimum, min(maximum, base + per_char * longest)))


def slugify(s):
    """Turn a month/date/boat/species name into a stable row id fragment.

    Content-based (not a running counter) so that adding a new day or boat
    elsewhere in the sheet doesn't shift every id after it -- that would
    silently break the saved expand/collapse state in localStorage.
    """
    s = re.sub(r'[^A-Za-z0-9æøåÆØÅ]+', '-', str(s))
    return s.strip('-').lower() or 'x'


# ---------- View builders ----------

def build_species_by_day(df, kommune, species_list):
    """KvitPrDag / PelPrDag shape: Month -> Day -> Boat, columns = species."""
    d = df[df['Kommune'] == kommune].copy()
    art_col = 'Art FAO' if 'Art FAO' in d.columns else 'Art'
    d = d[d[art_col].isin(species_list)]
    d['Landingsmåned'] = d['Landingsmåned'].map(clean_month)
    d['_date'] = d['Landingsdato'].map(parse_date)

    months = {}
    for month, mdf in d.groupby('Landingsmåned'):
        month_total = {sp: 0 for sp in species_list}
        days = {}
        for date, ddf in mdf.groupby('_date'):
            if date is None:
                continue
            day_total = {sp: 0 for sp in species_list}
            boats = {}
            for boat, bdf in ddf.groupby('Fartøynavn'):
                boat_vals = bdf.groupby(art_col)['Produktvekt'].sum().to_dict()
                row = {sp: boat_vals.get(sp, 0) for sp in species_list}
                boats[boat] = row
                for sp in species_list:
                    day_total[sp] += row[sp]
            days[date] = {'total': day_total, 'boats': boats}
            for sp in species_list:
                month_total[sp] += day_total[sp]
        months[month] = {'total': month_total, 'days': days}
    return {'kind': 'species_by_day', 'species': species_list, 'months': months}


def build_species_by_boat(df, kommune, species_list):
    """KvitPrBat shape: Boat -> Day, columns = species."""
    d = df[df['Kommune'] == kommune].copy()
    d = d[d['Art'].isin(species_list)]
    d['_date'] = d['Landingsdato'].map(parse_date)

    boats = {}
    for boat, bdf in d.groupby('Fartøynavn'):
        boat_total = {sp: 0 for sp in species_list}
        days = {}
        for date, ddf in bdf.groupby('_date'):
            if date is None:
                continue
            vals = ddf.groupby('Art')['Produktvekt'].sum().to_dict()
            row = {sp: vals.get(sp, 0) for sp in species_list}
            days[date] = row
            for sp in species_list:
                boat_total[sp] += row[sp]
        boats[boat] = {'total': boat_total, 'days': days}
    return {'kind': 'species_by_boat', 'species': species_list, 'boats': boats}


def build_catch_by_boat(df, kommune):
    """FangstPerBat shape: Boat -> Species, values = Siste fangstdag (max) + Kg (sum)."""
    d = df[df['Kommune'] == kommune].copy()
    d['_sfd'] = d['Siste fangstdato'].map(parse_date)

    boats = {}
    for boat, bdf in d.groupby('Fartøynavn'):
        species = {}
        for art, adf in bdf.groupby('Art'):
            species[art] = {'kg': adf['Produktvekt'].sum(), 'last_date': adf['_sfd'].max()}
        boats[boat] = {'kg': bdf['Produktvekt'].sum(), 'last_date': bdf['_sfd'].max(), 'species': species}
    return {'kind': 'catch_by_boat', 'boats': boats}


def build_art_by_month(df, kommune):
    """ArtPerBat shape: Month -> Species -> Boat, values = Siste fangstdato (max) + Kilo (sum)."""
    d = df[df['Kommune'] == kommune].copy()
    d['Landingsmåned'] = d['Landingsmåned'].map(clean_month)
    d['_sfd'] = d['Siste fangstdato'].map(parse_date)

    months = {}
    for month, mdf in d.groupby('Landingsmåned'):
        species = {}
        for art, adf in mdf.groupby('Art'):
            boats = {}
            for boat, bdf in adf.groupby('Fartøynavn'):
                boats[boat] = {'kg': bdf['Produktvekt'].sum(), 'last_date': bdf['_sfd'].max()}
            species[art] = {'kg': adf['Produktvekt'].sum(), 'last_date': adf['_sfd'].max(), 'boats': boats}
        months[month] = {'kg': mdf['Produktvekt'].sum(), 'last_date': mdf['_sfd'].max(), 'species': species}
    return {'kind': 'art_by_month', 'months': months}


# ---------- HTML rendering ----------

BASE_CSS = """
<style>
:root { color-scheme: light dark; }
body { font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; padding: 0; background: #fff; color: #1a1a1a; }
html, body { height: 100%; overscroll-behavior: contain; }
body { display: flex; flex-direction: column; overflow: hidden; }
.scrollwrap {
  flex: 1 1 auto;
  min-height: 0;
  overflow: auto;
  overscroll-behavior: contain;
  -webkit-overflow-scrolling: touch;
  max-width: 100%;
  box-sizing: border-box;
  border: 1px solid #ccc;
}
.updated { flex: 0 0 auto; }
table.pivot { border-collapse: separate; border-spacing: 0; font-size: 14px; }
table.pivot th, table.pivot td { border: 1px solid #ddd; padding: 6px 12px; text-align: right; white-space: nowrap; }
table.pivot .fz1, table.pivot .fz2, table.pivot .fz3 { text-align: left; }
tr.row-month .fz1 { font-size: 16px; }
tr.row-date .fz2 { font-weight: bold; }
table.pivot thead th { background: #f4f4f4; position: sticky; top: 0; z-index: 3; }
table.pivot .fz1 { position: sticky; left: 0; z-index: 2; }
table.pivot thead .fz1 { z-index: 5; }
table.pivot .fz2, table.pivot .fz3 {
  position: sticky; z-index: 2;
}
table.pivot .fz3 { box-shadow: 2px 0 4px -2px rgba(0,0,0,0.25); }
table.pivot thead .fz2, table.pivot thead .fz3 { z-index: 5; }
tr.level-0 { background: #eef3f8; font-weight: 600; cursor: pointer; }
tr.level-1 { background: #f8fafc; cursor: pointer; }
tr.level-2 { background: #ffffff; }
tr.level-0 .fz1, tr.level-0 .fz2, tr.level-0 .fz3 { background: #eef3f8; }
tr.level-1 .fz1, tr.level-1 .fz2, tr.level-1 .fz3 { background: #f8fafc; }
tr.level-2 .fz1, tr.level-2 .fz2, tr.level-2 .fz3 { background: #ffffff; }
tr.hidden { display: none; }
.toggle { display: inline-block; width: 1em; }
.updated { font-size: 12px; color: #777; padding: 6px 4px 10px; }
</style>
<script>
// Expanding a branch (month/boat/species row) cascades open everything
// below it in one click; collapsing it closes the whole branch again.
function cascadeSet(id, expand) {
  var rows = document.querySelectorAll('[data-parent="' + id + '"]');
  rows.forEach(function(r){
    r.classList.toggle('hidden', !expand);
    if (r.dataset.id) {
      cascadeSet(r.dataset.id, expand);
      var childToggle = document.getElementById('toggle-' + r.dataset.id);
      if (childToggle) childToggle.textContent = expand ? '−' : '+';
    }
  });
  var toggle = document.getElementById('toggle-' + id);
  if (toggle) toggle.textContent = expand ? '−' : '+';
}

function toggleRow(id) {
  var rows = document.querySelectorAll('[data-parent="' + id + '"]');
  var expand = rows.length > 0 && rows[0].classList.contains('hidden');
  cascadeSet(id, expand);
  saveViewState();
}

// Per-page (per URL path), per-browser memory of which rows are open --
// so a returning visitor sees the layout they left, instead of the
// server-rendered default every time. Falls back to that default silently
// wherever storage is unavailable (private browsing, cross-site iframe
// restrictions in some browsers, etc.).
function storageKey() {
  return 'fangst-pivot:' + location.pathname;
}

function saveViewState() {
  var state = {};
  document.querySelectorAll('[id^="row-"]').forEach(function(r){
    var id = r.id.slice(4);
    var kids = document.querySelectorAll('[data-parent="' + id + '"]');
    if (kids.length && !kids[0].classList.contains('hidden')) state[id] = true;
  });
  try { localStorage.setItem(storageKey(), JSON.stringify(state)); } catch (e) {}
}

function restoreViewState() {
  var raw = null;
  try { raw = localStorage.getItem(storageKey()); } catch (e) {}
  if (raw === null) return; // nothing saved yet -- keep the page's built-in default

  var state = {};
  try { state = JSON.parse(raw) || {}; } catch (e) {}

  document.querySelectorAll('[data-parent]').forEach(function(r){ r.classList.add('hidden'); });
  document.querySelectorAll('[id^="toggle-"]').forEach(function(t){ t.textContent = '+'; });

  // Document order is parent-before-child, so unhiding a parent's children
  // here always runs before we reach that child's own row -- no recursion
  // needed to rebuild arbitrarily nested saved states.
  document.querySelectorAll('[id^="row-"]').forEach(function(r){
    var id = r.id.slice(4);
    if (!state[id]) return;
    document.querySelectorAll('[data-parent="' + id + '"]').forEach(function(c){
      c.classList.remove('hidden');
    });
    var t = document.getElementById('toggle-' + id);
    if (t) t.textContent = '−';
  });
}

document.addEventListener('DOMContentLoaded', restoreViewState);
</script>
"""


def page_shell(title, updated, thead_html, tbody_html, frozen_px, extra_cols):
    """frozen_px: list of pixel widths for the frozen (fz1, fz2, fz3, ...) columns,
    left to right. extra_cols: number of scrolling data columns after them."""
    col_rules = []
    left_rules = []
    running_left = 0
    for i, w in enumerate(frozen_px, start=1):
        col_rules.append(f"col.c{i}{{width:{w}px}}")
        left_rules.append(f"table.pivot .fz{i} {{ left:{running_left}px; }}")
        running_left += w
    style = f"<style>{''.join(col_rules)}{''.join(left_rules)}</style>"
    colgroup_cols = ''.join(f"<col class='c{i}'>" for i in range(1, len(frozen_px) + 1))
    colgroup = f"<colgroup>{colgroup_cols}{'<col>' * extra_cols}</colgroup>"
    return (
        f"<!doctype html><html><head><meta charset='utf-8'><title>{html.escape(title)}</title>"
        f"{BASE_CSS}{style}</head><body>"
        f"<div class='scrollwrap'><table class='pivot'>{colgroup}"
        f"<thead>{thead_html}</thead><tbody>{tbody_html}</tbody></table></div>"
        f"<div class='updated'>Sist oppdatert: {html.escape(updated)}</div>"
        f"</body></html>"
    )


def month_key(m):
    try:
        return MONTH_ORDER.index(m)
    except ValueError:
        return 99


def render_species_tree(view, title, updated, current_month=None):
    species = view['species']
    months = view['months']
    rows = []
    col1_texts = list(months.keys())
    col2_texts = []
    col3_texts = []

    for month in sorted(months.keys(), key=month_key):
        mdata = months[month]
        mid = f"m-{slugify(month)}"
        is_cur = (month == current_month)
        m_glyph = MINUS if is_cur else '+'
        cells = ''.join(f"<td>{fmt_num(mdata['total'][sp])}</td>" for sp in species)
        rows.append(
            f'<tr class="level-0 row-month" id="row-{mid}" onclick="toggleRow(\'{mid}\')">'
            f'<td class="fz1">{html.escape(month)}<span class="toggle" id="toggle-{mid}">{m_glyph}</span></td>'
            f'<td class="fz2"></td><td class="fz3"></td>{cells}</tr>'
        )
        for date in sorted(mdata['days'].keys()):
            ddata = mdata['days'][date]
            did = f"d-{slugify(month)}-{slugify(fmt_date(date))}"
            col2_texts.append(fmt_date(date))
            d_hidden = '' if is_cur else ' hidden'
            d_glyph = MINUS if is_cur else '+'
            dcells = ''.join(f"<td>{fmt_num(ddata['total'][sp])}</td>" for sp in species)
            rows.append(
                f'<tr class="level-1 row-date{d_hidden}" data-parent="{mid}" data-id="{did}" id="row-{did}" '
                f'onclick="event.stopPropagation(); toggleRow(\'{did}\')">'
                f'<td class="fz1"></td><td class="fz2">{fmt_date(date)}<span class="toggle" id="toggle-{did}">{d_glyph}</span></td>'
                f'<td class="fz3"></td>{dcells}</tr>'
            )
            for boat, brow in sorted(ddata['boats'].items()):
                col3_texts.append(str(boat))
                b_hidden = '' if is_cur else ' hidden'
                bcells = ''.join(f"<td>{fmt_num(brow[sp])}</td>" for sp in species)
                rows.append(
                    f'<tr class="level-2{b_hidden}" data-parent="{did}">'
                    f'<td class="fz1"></td><td class="fz2"></td><td class="fz3">{html.escape(str(boat))}</td>{bcells}</tr>'
                )
    header_cells = ''.join(f"<th>{html.escape(sp)}</th>" for sp in species)
    thead = f"<tr><th class='fz1'>Mnd</th><th class='fz2'>Dato</th><th class='fz3'>Båt</th>{header_cells}</tr>"
    col1_px = px_width(col1_texts)
    col2_px = px_width(col2_texts or ['Dato'])
    col3_px = px_width(col3_texts or ['Båt'])
    return page_shell(title, updated, thead, ''.join(rows), [col1_px, col2_px, col3_px], len(species))


def render_boat_species_tree(view, title, updated):
    species = view['species']
    boats = view['boats']
    rows = []
    col1_texts = list(boats.keys())
    col2_texts = []
    for boat in sorted(boats.keys()):
        bdata = boats[boat]
        bid = f"b-{slugify(boat)}"
        cells = ''.join(f"<td>{fmt_num(bdata['total'][sp])}</td>" for sp in species)
        rows.append(
            f'<tr class="level-0" id="row-{bid}" onclick="toggleRow(\'{bid}\')">'
            f'<td class="fz1"><span class="toggle" id="toggle-{bid}">+</span>{html.escape(str(boat))}</td><td class="fz2"></td>{cells}</tr>'
        )
        for date in sorted(bdata['days'].keys()):
            drow = bdata['days'][date]
            col2_texts.append(fmt_date(date))
            dcells = ''.join(f"<td>{fmt_num(drow[sp])}</td>" for sp in species)
            rows.append(
                f'<tr class="level-1 row-date hidden" data-parent="{bid}">'
                f'<td class="fz1"></td><td class="fz2">{fmt_date(date)}</td>{dcells}</tr>'
            )
    header_cells = ''.join(f"<th>{html.escape(sp)}</th>" for sp in species)
    thead = f"<tr><th class='fz1'>Båt</th><th class='fz2'>Dato</th>{header_cells}</tr>"
    col1_px = px_width(col1_texts)
    col2_px = px_width(col2_texts or ['Dato'])
    return page_shell(title, updated, thead, ''.join(rows), [col1_px, col2_px], len(species))


def render_catch_by_boat(view, title, updated):
    boats = view['boats']
    rows = []
    col1_texts = list(boats.keys())
    col2_texts = []
    for boat in sorted(boats.keys()):
        bdata = boats[boat]
        bid = f"b-{slugify(boat)}"
        rows.append(
            f'<tr class="level-0" id="row-{bid}" onclick="toggleRow(\'{bid}\')">'
            f'<td class="fz1"><span class="toggle" id="toggle-{bid}">+</span>{html.escape(str(boat))}</td>'
            f'<td class="fz2"></td><td>{fmt_date(bdata["last_date"])}</td><td>{fmt_num(bdata["kg"])}</td></tr>'
        )
        for art, sdata in sorted(bdata['species'].items()):
            col2_texts.append(str(art))
            rows.append(
                f'<tr class="level-1 hidden" data-parent="{bid}">'
                f'<td class="fz1"></td><td class="fz2">{html.escape(str(art))}</td>'
                f'<td>{fmt_date(sdata["last_date"])}</td><td>{fmt_num(sdata["kg"])}</td></tr>'
            )
    thead = "<tr><th class='fz1'>Båt</th><th class='fz2'>Art</th><th>Siste fangstdag</th><th>Kg</th></tr>"
    col1_px = px_width(col1_texts)
    col2_px = px_width(col2_texts or ['Art'])
    return page_shell(title, updated, thead, ''.join(rows), [col1_px, col2_px], 2)


def render_art_by_month(view, title, updated, current_month=None):
    months = view['months']
    rows = []
    col1_texts = list(months.keys())
    col2_texts = []
    col3_texts = []

    for month in sorted(months.keys(), key=month_key):
        mdata = months[month]
        mid = f"m-{slugify(month)}"
        is_cur = (month == current_month)
        m_glyph = MINUS if is_cur else '+'
        rows.append(
            f'<tr class="level-0 row-month" id="row-{mid}" onclick="toggleRow(\'{mid}\')">'
            f'<td class="fz1">{html.escape(month)}<span class="toggle" id="toggle-{mid}">{m_glyph}</span></td>'
            f'<td class="fz2"></td><td class="fz3"></td>'
            f'<td>{fmt_date(mdata["last_date"])}</td><td>{fmt_num(mdata["kg"])}</td></tr>'
        )
        for art, sdata in sorted(mdata['species'].items()):
            sid = f"s-{slugify(month)}-{slugify(art)}"
            col2_texts.append(str(art))
            s_hidden = '' if is_cur else ' hidden'
            s_glyph = MINUS if is_cur else '+'
            rows.append(
                f'<tr class="level-1{s_hidden}" data-parent="{mid}" data-id="{sid}" id="row-{sid}" '
                f'onclick="event.stopPropagation(); toggleRow(\'{sid}\')">'
                f'<td class="fz1"></td><td class="fz2"><span class="toggle" id="toggle-{sid}">{s_glyph}</span>{html.escape(str(art))}</td>'
                f'<td class="fz3"></td>'
                f'<td>{fmt_date(sdata["last_date"])}</td><td>{fmt_num(sdata["kg"])}</td></tr>'
            )
            for boat, bdata in sorted(sdata['boats'].items()):
                col3_texts.append(str(boat))
                b_hidden = '' if is_cur else ' hidden'
                rows.append(
                    f'<tr class="level-2{b_hidden}" data-parent="{sid}">'
                    f'<td class="fz1"></td><td class="fz2"></td><td class="fz3">{html.escape(str(boat))}</td>'
                    f'<td>{fmt_date(bdata["last_date"])}</td><td>{fmt_num(bdata["kg"])}</td></tr>'
                )
    thead = "<tr><th class='fz1'>Mnd</th><th class='fz2'>Art</th><th class='fz3'>Båt</th><th>Siste fangstdato</th><th>Kilo</th></tr>"
    col1_px = px_width(col1_texts)
    col2_px = px_width(col2_texts or ['Art'])
    col3_px = px_width(col3_texts or ['Båt'])
    return page_shell(title, updated, thead, ''.join(rows), [col1_px, col2_px, col3_px], 2)


def main():
    xlsx_path = sys.argv[1]
    out_dir = sys.argv[2]
    import os
    os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    out_df = load_sheet(wb, 'OUTPUT')
    pel_df = load_sheet(wb, 'CatchUnzipPel')
    # "Last updated" stamp = the date this export actually ran (Oslo time),
    # not a manually-maintained cell in the workbook -- avoids the pages
    # silently going stale if that cell isn't kept in sync.
    try:
        from zoneinfo import ZoneInfo
        now = datetime.now(ZoneInfo('Europe/Oslo'))
    except Exception:
        now = datetime.utcnow()
    updated = now.strftime('%d.%m.%Y')
    current_month = MONTH_ORDER[now.month - 1]

    manifest = []
    for kommune in MUNICIPALITIES:
        prefix = PAGE_PREFIX[kommune]

        v = build_species_by_day(out_df, kommune, KVITFISK_SPECIES)
        fn = f"{prefix}-KvitPrDag.html"
        open(f"{out_dir}/{fn}", 'w', encoding='utf-8').write(
            render_species_tree(v, f"{kommune} - Kvitfisk per dag", updated, current_month))
        manifest.append(fn)

        if kommune in ('HERØY', 'VANYLVEN'):
            v = build_species_by_day(pel_df, kommune, PELAGISK_SPECIES)
            fn = f"{prefix}-PelPrDag.html"
            open(f"{out_dir}/{fn}", 'w', encoding='utf-8').write(
                render_species_tree(v, f"{kommune} - Pelagisk per dag", updated, current_month))
            manifest.append(fn)

        v = build_species_by_boat(out_df, kommune, KVITFISK_SPECIES)
        fn = f"{prefix}-KvitPrBat.html"
        open(f"{out_dir}/{fn}", 'w', encoding='utf-8').write(
            render_boat_species_tree(v, f"{kommune} - Kvitfisk per fartøy", updated))
        manifest.append(fn)

        v = build_catch_by_boat(out_df, kommune)
        fn = f"{prefix}-FangstPerBat.html"
        open(f"{out_dir}/{fn}", 'w', encoding='utf-8').write(
            render_catch_by_boat(v, f"{kommune} - Fangst per fartøy", updated))
        manifest.append(fn)

        v = build_art_by_month(out_df, kommune)
        fn = f"{prefix}-ArtPerBat.html"
        open(f"{out_dir}/{fn}", 'w', encoding='utf-8').write(
            render_art_by_month(v, f"{kommune} - Fangst per art", updated, current_month))
        manifest.append(fn)

    print(f"Wrote {len(manifest)} files to {out_dir}:")
    for m in manifest:
        print(" -", m)


if __name__ == '__main__':
    main()
