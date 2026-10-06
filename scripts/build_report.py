#!/usr/bin/env python3
"""Build the daily prisoners report as a Word file in the approved design.

Usage: python3 scripts/build_report.py reports/YYYY-MM-DD.json [output.docx]

The design (styles, header, footer, bullets, colours, page setup) comes from
templates/report_template.docx; only the body is generated from the JSON.
See REPORT_FORMAT.md for the JSON structure.
"""
import json
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "templates", "report_template.docx")

NAVY, MAROON, GREY_TEXT = "1B365D", "8E2430", "5F6B7A"
LIGHT, GROUP, GRID = "EEF2F7", "DCE4EE", "C9D1DC"
STATUS = {
    "confirmed": ("مؤكد", "2E7D32"),
    "attributed": ("منسوب", "B26A00"),
    "unverified": ("غير محقق", "7A7A7A"),
    "alert": ("تنبيه", MAROON),
}
FULL_WIDTH = 9638
HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

FONTS = '<w:rFonts w:asciiTheme="minorBidi" w:hAnsiTheme="minorBidi" w:cstheme="minorBidi"/>'
SIZE = '<w:sz w:val="26"/><w:szCs w:val="26"/>'
NUM_BULLET, NUM_SOURCES = 2, 3


class Links:
    """Collects hyperlink relationships for document.xml.rels."""

    def __init__(self):
        self.targets = []

    def rid(self, url):
        self.targets.append(url)
        return f"rIdL{len(self.targets)}"


LINKS = Links()


# ---------------------------------------------------------------- runs

def run(text, bold=False, color=None, shade=None, size=None):
    props = FONTS
    if bold:
        props += "<w:b/><w:bCs/>"
    if color:
        props += f'<w:color w:val="{color}"/>'
    props += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>' if size else SIZE
    if shade:
        props += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    return f'<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def link(text, url):
    return (f'<w:hyperlink r:id="{LINKS.rid(url)}" w:history="1"><w:r><w:rPr>{FONTS}'
            f'<w:rStyle w:val="Hyperlink"/>{SIZE}</w:rPr>'
            f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:hyperlink>')


TOKEN = re.compile(r"\*\*(.+?)\*\*|\[([^\]]+)\]\(([^)\s]+)\)")


def rich(text, bold=False, color=None):
    """Render **bold** and [text](url) markup into runs."""
    out, pos = [], 0
    for m in TOKEN.finditer(text or ""):
        if m.start() > pos:
            out.append(run(text[pos:m.start()], bold, color))
        if m.group(1) is not None:
            out.append(run(m.group(1), True, color))
        else:
            out.append(link(m.group(2), m.group(3)))
        pos = m.end()
    if pos < len(text or ""):
        out.append(run(text[pos:], bold, color))
    return "".join(out)


def badge(status, label=None, dot=True):
    name, fill = STATUS[status]
    text = label or name
    return run(f" ● {text} " if dot else text, True, "FFFFFF", fill)


def links_list(items, color=None):
    parts = []
    for i, item in enumerate(items):
        if i:
            parts.append(run("، ", color=color))
        if isinstance(item, str):
            parts.append(rich(item, color=color))
        elif item.get("url"):
            parts.append(link(item["text"], item["url"]))
        else:
            parts.append(run(item["text"], color=color))
    return "".join(parts)


# ---------------------------------------------------------------- paragraphs

def para(content, jc="both", spacing='w:line="276" w:lineRule="auto"', extra=""):
    return f'<w:p><w:pPr>{extra}<w:bidi/><w:spacing {spacing}/><w:jc w:val="{jc}"/></w:pPr>{content}</w:p>'


def heading(text, level):
    return f'<w:p><w:pPr><w:pStyle w:val="{level}"/></w:pPr>{run(text)}</w:p>'


def spacer():
    return para("", spacing='w:line="200" w:lineRule="auto"')


def bullet(content, level=0, num=NUM_BULLET, after=70):
    return (f'<w:p><w:pPr><w:pStyle w:val="a4"/><w:numPr><w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{num}"/></w:numPr><w:bidi/>'
            f'<w:spacing w:after="{after}" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
            f"{content}</w:p>")


def status_prefix(block):
    if block.get("status"):
        return badge(block["status"], block.get("label")) + run(" ")
    return ""


def sources_line(items, level=1):
    content = run("المصادر: ", True, GREY_TEXT) + links_list(items, GREY_TEXT)
    return bullet(content, level, NUM_SOURCES, after=90)


# ---------------------------------------------------------------- tables

def mar(top, side):
    return (f'<w:tcMar><w:top w:w="{top}" w:type="dxa"/><w:left w:w="{side}" w:type="dxa"/>'
            f'<w:bottom w:w="{top}" w:type="dxa"/><w:right w:w="{side}" w:type="dxa"/></w:tcMar>')


def borders(color=GRID, val="single", sz=4, tag="tblBorders"):
    sides = "".join(f'<w:{s} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
                    for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return f"<w:{tag}>{sides}</w:{tag}>"


NO_BORDERS = borders("FFFFFF", "none", 0)


def table(widths, rows_xml, tbl_borders=None, indent=None, look="04A0"):
    total = sum(widths)
    ind = f'<w:tblInd w:w="{indent}" w:type="dxa"/>' if indent is not None else ""
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return (f'<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="{total}" w:type="dxa"/>{ind}'
            f'{tbl_borders or borders()}<w:tblLayout w:type="fixed"/><w:tblCellMar>'
            f'<w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar>'
            f'<w:tblLook w:val="{look}"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
            f'{"".join(rows_xml)}</w:tbl>')


def cell(width, paragraphs, fill=None, margin=mar(70, 110), span=None, tc_borders=""):
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span:
        props += f'<w:gridSpan w:val="{span}"/>'
    props += tc_borders
    if fill:
        props += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    props += f'{margin}<w:vAlign w:val="center"/>'
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{''.join(paragraphs)}</w:tc>"


def row(cells, header=False, cant_split=True):
    trpr = "<w:tblHeader/>" if header else ("<w:cantSplit/>" if cant_split else "")
    return f"<w:tr>{f'<w:trPr>{trpr}</w:trPr>' if trpr else ''}{''.join(cells)}</w:tr>"


def cpara(content, jc="both", line=270):
    return para(content, jc, f'w:line="{line}" w:lineRule="auto"')


def header_row(widths, titles, margin=mar(70, 110), compact=False):
    return row([cell(w, [para(run(t, True, "FFFFFF"), "center", 'w:line="270" w:lineRule="auto"')
                         if not compact else f'<w:p><w:pPr><w:bidi/><w:jc w:val="center"/></w:pPr>{run(t, True, "FFFFFF")}</w:p>'],
                     NAVY, margin) for w, t in zip(widths, titles)], header=True)


def status_cell_content(value, dot=True):
    """A table cell value: text, {"status": ...}, a list of statuses, or {"links": [...]}."""
    if isinstance(value, dict) and "status" in value:
        value = [value]
    if isinstance(value, list) and value and isinstance(value[0], dict) and "status" in value[0]:
        return [cpara(badge(v["status"], v.get("label"), dot), "center") for v in value], "center"
    if isinstance(value, dict) and "links" in value:
        return [cpara(links_list(value["links"]), "center")], "center"
    return None, None


# ---------------------------------------------------------------- sections

def title_block(r):
    widths = [2301, 7342]
    banner_border = ('<w:tcBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/>'
                     '<w:left w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/>'
                     f'<w:bottom w:val="single" w:sz="36" w:space="0" w:color="{MAROON}"/>'
                     '<w:right w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/></w:tcBorders>')
    banner = cell(9638, [
        para(run("تقرير الرصد اليومي", True, "C9D6E8"), "center", 'w:after="60" w:line="276" w:lineRule="auto"'),
        para(run("أخبار وانتهاكات الأسرى الفلسطينيين", True, "FFFFFF"), "center",
             'w:after="80" w:line="276" w:lineRule="auto"'),
        para(run(r["date"], color="FFFFFF"), "center"),
    ], NAVY, mar(260, 300), span=2, tc_borders=banner_border)
    rows = [row([banner], cant_split=False)]
    ex = f"<w:tblPrEx>{borders()}</w:tblPrEx>"
    for label, value in r["meta"]:
        cells = cell(2300, [cpara(run(label, True, NAVY), "center")], LIGHT) + \
            cell(7338, [cpara(rich(value))])
        rows.append(f"<w:tr>{ex}{cells}</w:tr>")
    return table(widths, rows, NO_BORDERS)


def key_block():
    rows = []
    for status, desc in (("confirmed", "ورد في مصدرين مستقلين على الأقل، أو في بيان رسمي أو حقوقي."),
                         ("attributed", "رواية صادرة عن جهة محددة، ولم يتوفر لها تأكيد مستقل."),
                         ("unverified", "معلومات ناقصة، أو تاريخها غير مؤكد، أو تتضارب فيها المصادر.")):
        rows.append(row([cell(1700, [cpara(badge(status), "center")]),
                         cell(7938, [cpara(run(desc))])], cant_split=False))
    label = para(run("مفتاح التصنيف", True, NAVY), extra="<w:keepNext/>",
                 spacing='w:after="60" w:line="276" w:lineRule="auto"')
    return label + table([1700, 7938], rows)


def stats_block(stats):
    if not stats:
        return ""
    n = len(stats)
    widths = [FULL_WIDTH // n] * n
    widths[-1] += FULL_WIDTH - sum(widths)
    tile_border = (f'<w:tcBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="{MAROON}"/>'
                   '<w:left w:val="single" w:sz="12" w:space="0" w:color="FFFFFF"/>'
                   '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/>'
                   '<w:right w:val="single" w:sz="12" w:space="0" w:color="FFFFFF"/></w:tcBorders>')
    cells = [cell(w, [para(run(s["value"], True, NAVY), "center", 'w:after="40" w:line="276" w:lineRule="auto"'),
                      para(run(s["label"], color=GREY_TEXT), "center", 'w:line="250" w:lineRule="auto"')],
                  LIGHT, mar(160, 100), tc_borders=tile_border) for w, s in zip(widths, stats)]
    return heading("الملخص التنفيذي", "1") + table(widths, [row(cells, cant_split=False)], NO_BORDERS)


def brief_block(brief, not_observed):
    widths = [858, 5650, 1598, 1970]
    m = mar(55, 90)
    compact = lambda content, jc: f'<w:p><w:pPr><w:bidi/><w:jc w:val="{jc}"/></w:pPr>{content}</w:p>'
    rows = [header_row(widths, ["م", "الخبر", "التصنيف", "المصدر"], m, compact=True)]
    number = 0
    for item in brief:
        if "group" in item:
            rows.append(row([cell(sum(widths), [compact(run(item["group"], True, NAVY), "center")],
                                  GROUP, m, span=4)]))
            continue
        number += 1
        statuses = item["status"] if isinstance(item["status"], list) else [item["status"]]
        statuses = [s if isinstance(s, dict) else {"status": s} for s in statuses]
        rows.append(row([
            cell(widths[0], [compact(run(str(number), True, NAVY), "center")], LIGHT, m),
            cell(widths[1], [compact(rich(item["text"]), "both")], None, m),
            cell(widths[2], [compact(badge(s["status"], s.get("label"), dot=False), "center")
                             for s in statuses], None, m),
            cell(widths[3], [compact(links_list(item.get("sources", [])), "center")], None, m),
        ]))
    out = para(run("موجز الأخبار", True, NAVY),
               extra=f'<w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="3" w:color="{MAROON}"/></w:pBdr>',
               spacing='w:after="40" w:line="276" w:lineRule="auto"')
    out += para(run("كل خبر مع تصنيفه ومصدره. التفاصيل الكاملة في الصفحات التالية.", color=GREY_TEXT),
                spacing='w:after="80" w:line="276" w:lineRule="auto"')
    out += table(widths, rows, indent=-438, look="0000")
    if not_observed:
        out += para(run("لم يُرصد خلال الفترة: ", True, MAROON) + rich(not_observed),
                    spacing='w:before="70" w:after="30" w:line="276" w:lineRule="auto"')
    return out


def data_table(block):
    cols = block["columns"]
    widths = [c["width"] for c in cols]
    scale = FULL_WIDTH / sum(widths)
    widths = [round(w * scale) for w in widths]
    widths[-1] += FULL_WIDTH - sum(widths)
    rows = [header_row(widths, [c["title"] for c in cols])]
    zebra = block.get("zebra", False)
    for r_index, values in enumerate(block["rows"]):
        cells = []
        for c_index, (w, value) in enumerate(zip(widths, values)):
            first = c_index == 0 and block.get("first_col_label", True)
            fill = LIGHT if first or (zebra and r_index % 2 == 0) else None
            paras, _ = status_cell_content(value)
            if paras is None:
                jc = cols[c_index].get("align", "both")
                if value in ("", "—", None):
                    paras = [cpara(run("—"), "center")]
                elif first:
                    paras = [cpara(rich(value, True, block.get("label_color", NAVY)), jc)]
                else:
                    paras = [cpara(rich(value, color=cols[c_index].get("color")), jc)]
            cells.append(cell(w, paras, fill))
        rows.append(row(cells))
    return table(widths, rows)


def render_block(block):
    kind = block["type"]
    if kind == "h2":
        return heading(block["text"], "2")
    if kind == "bullet":
        return bullet(status_prefix(block) + rich(block["text"]), block.get("level", 0))
    if kind == "sources":
        return sources_line(block["items"], block.get("level", 1))
    if kind == "para":
        return para(status_prefix(block) + rich(block["text"]),
                    spacing='w:after="100" w:line="276" w:lineRule="auto"')
    if kind == "table":
        return data_table(block)
    if kind == "note":
        return para(rich(block["text"]), spacing='w:before="80" w:after="100" w:line="276" w:lineRule="auto"')
    raise ValueError(f"unknown block type: {kind}")


def details_block(sections):
    out = para("", spacing='w:line="276" w:lineRule="auto"')
    out += para(run("التفاصيل", True, GREY_TEXT, size=36), spacing='w:line="276" w:lineRule="auto"')
    for section in sections:
        out += heading(section["title"], "1")
        out += "".join(render_block(b) for b in section.get("blocks", []))
    return out


def body(r):
    return "".join([
        title_block(r), spacer(), key_block(),
        stats_block(r.get("stats")),
        spacer(), brief_block(r.get("brief", []), r.get("not_observed")),
        details_block(r.get("sections", [])),
    ])


# ---------------------------------------------------------------- packaging

def build(report, out_path):
    global LINKS
    LINKS = Links()
    with zipfile.ZipFile(TEMPLATE) as tpl:
        files = {i.filename: tpl.read(i.filename) for i in tpl.infolist()}
    doc = files["word/document.xml"].decode("utf-8")
    sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S).group(0)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    files["word/document.xml"] = (head + body(report) + sect + "</w:body></w:document>").encode("utf-8")

    rels = files["word/_rels/document.xml.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship [^>]*Type="%s"[^>]*/>' % re.escape(HYPERLINK_TYPE), "", rels)
    new = "".join(f'<Relationship Id="rIdL{i}" Type="{HYPERLINK_TYPE}" Target="{escape(u, {chr(34): "&quot;"})}" '
                  f'TargetMode="External"/>' for i, u in enumerate(LINKS.targets, 1))
    files["word/_rels/document.xml.rels"] = rels.replace("</Relationships>", new + "</Relationships>").encode("utf-8")

    header = files["word/header1.xml"].decode("utf-8")
    header = re.sub(r"(<w:t[^>]*>)\s*\|\s*[^<]*(</w:t>)",
                    lambda m: f"{m.group(1)}   |   {escape(report['header_date'])}{m.group(2)}", header)
    files["word/header1.xml"] = header.encode("utf-8")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in files.items():
            out.writestr(name, data)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".docx"
    with open(src, encoding="utf-8") as fh:
        report = json.load(fh)
    build(report, out)
    print(out)


if __name__ == "__main__":
    main()
