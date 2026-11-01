"""Exports Excel (plusieurs onglets) et page HTML autonome, triable."""
from __future__ import annotations

import html as _html
import io
from datetime import date

import pandas as pd


def to_excel_bytes(sheets: dict[str, pd.DataFrame], money_cols: set[str],
                   link_cols: set[str], note: str = "",
                   formulas: dict[str, tuple[str, str]] | None = None) -> bytes:
    """
    sheets   : {nom d'onglet: DataFrame déjà renommé pour l'affichage}
    formulas : {colonne résultat: (colonne A, colonne B)} → =A-B
    """
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    formulas = formulas or {}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, index=False, sheet_name=name[:31])
            ws = xw.sheets[name[:31]]
            cols = list(df.columns)
            pos = {c: i + 1 for i, c in enumerate(cols)}
            for c in ws[1]:
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
                c.alignment = Alignment(wrap_text=True, vertical="center")
            for r in range(2, ws.max_row + 1):
                for c in ws[r]:
                    c.font = Font(name="Arial", size=10)
                for res, (a, b) in formulas.items():
                    if res in pos and a in pos and b in pos:
                        ws.cell(r, pos[res]).value = (
                            f"={get_column_letter(pos[a])}{r}-{get_column_letter(pos[b])}{r}")
                for col in money_cols & set(cols):
                    ws.cell(r, pos[col]).number_format = "$#,##0;($#,##0);-"
                for col in link_cols & set(cols):
                    cell = ws.cell(r, pos[col])
                    if cell.value:
                        cell.hyperlink = cell.value
                        cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")
            for i, col in enumerate(cols, start=1):
                width = min(max(len(str(col)), *(len(str(v)) for v in df[col].head(200)), 8), 50)
                ws.column_dimensions[get_column_letter(i)].width = width + 2
            ws.row_dimensions[1].height = 36
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            if note:
                ws.cell(ws.max_row + 2, 1, note).font = Font(name="Arial", italic=True, size=9)
    return buf.getvalue()


def to_html_page(df: pd.DataFrame, title: str, money_cols: set[str],
                 link_cols: set[str], positive_cols: set[str] = frozenset(),
                 subtitle: str = "") -> str:
    def fmt(col, v):
        if pd.isna(v):
            return ""
        if col in link_cols:
            return f'<a href="{_html.escape(str(v))}" target="_blank">fiche</a>'
        if col in money_cols:
            return f"${v:,.0f}".replace(",", " ")
        if isinstance(v, bool):
            return "✔" if v else ""
        if isinstance(v, float) and v.is_integer():
            return f"{int(v)}"
        return _html.escape(str(v))

    def style(col, v):
        if col in positive_cols and isinstance(v, (int, float)) and not pd.isna(v):
            return ' class="pos"' if v > 0 else (' class="neg"' if v < 0 else "")
        return ""

    head = "".join(f"<th>{_html.escape(str(c))}</th>" for c in df.columns)
    body = "\n".join(
        "<tr>" + "".join(f"<td{style(c, v)}>{fmt(c, v)}</td>" for c, v in row.items()) + "</tr>"
        for _, row in df.iterrows())
    return f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_html.escape(title)}</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:24px;color:#1f2328;background:#fff}}
 h1{{font-size:20px;margin:0 0 4px}} p.sub{{color:#57606a;margin:0 0 16px;font-size:13px}}
 .wrap{{overflow-x:auto}}
 table{{border-collapse:collapse;font-size:13px;width:100%}}
 th{{background:#1f4e78;color:#fff;text-align:left;padding:8px;position:sticky;top:0;cursor:pointer;white-space:nowrap}}
 td{{padding:6px 8px;border-bottom:1px solid #d0d7de;vertical-align:top}}
 tr:nth-child(even) td{{background:#f6f8fa}}
 td.pos{{color:#1a7f37;font-weight:600}} td.neg{{color:#cf222e}}
</style></head><body>
<h1>{_html.escape(title)}</h1>
<p class="sub">{len(df)} valeurs · relevé le {date.today():%d/%m/%Y} · {_html.escape(subtitle)}
 · cliquer un en-tête pour trier</p>
<div class="wrap"><table><thead><tr>{head}</tr></thead><tbody>
{body}
</tbody></table></div>
<script>
document.querySelectorAll('th').forEach((th,i)=>th.addEventListener('click',()=>{{
 const tb=th.closest('table').tBodies[0], rows=[...tb.rows], asc=th.dataset.asc!=='1';
 const v=r=>{{const t=r.cells[i].innerText.replace(/[$\\s%]/g,'');const n=parseFloat(t);return isNaN(n)?t:n}};
 rows.sort((a,b)=>{{const x=v(a),y=v(b);return (x>y?1:x<y?-1:0)*(asc?1:-1)}});
 th.dataset.asc=asc?'1':'0'; rows.forEach(r=>tb.appendChild(r));
}}));
</script>
</body></html>"""
