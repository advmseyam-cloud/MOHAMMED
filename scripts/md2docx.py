#!/usr/bin/env python3
"""Convert an Arabic daily report (Markdown) into a right-to-left Word file.

Usage: python3 scripts/md2docx.py reports/2026-10-06.md [output.docx]
Requires pandoc. The output defaults to the input path with a .docx suffix.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

FONT = "Arial"
HEADING_COLOR = "1F4E3D"


def rewrite_zip(path, edits):
    """Rewrite selected members of a zip archive via callables."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in edits:
                data = edits[item.filename](data.decode("utf-8")).encode("utf-8")
            dst.writestr(item, data)
    os.replace(tmp, path)


def patch_styles(xml):
    fonts = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}" />'
    xml = re.sub(r"<w:rFonts [^>]*/>", fonts, xml)
    xml = xml.replace('w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"',
                      'w:val="ar-SA" w:eastAsia="en-US" w:bidi="ar-SA"')
    # Arabic text is sized by szCs, so mirror every sz that lacks one.
    xml = re.sub(r'(<w:sz w:val="(\d+)" />)(?!\s*<w:szCs)', r'\1<w:szCs w:val="\2" />', xml)
    xml = re.sub(r'<w:color w:val="[0-9A-F]{6}" w:themeColor="accent1"[^>]*/>',
                 f'<w:color w:val="{HEADING_COLOR}" />', xml)
    borders = ("<w:tblBorders>" + "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="999999" />'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>")
    xml = xml.replace('<w:tblInd w:w="0" w:type="dxa" />\n      <w:tblCellMar>',
                      '<w:bidiVisual /><w:tblInd w:w="0" w:type="dxa" />' + borders + "<w:tblCellMar>", 1)
    return xml


def patch_document(xml):
    xml = re.sub(r"<w:tblPr>(?!<w:bidiVisual)", "<w:tblPr><w:bidiVisual />", xml)
    return xml


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".docx"
    if not shutil.which("pandoc"):
        sys.exit("pandoc is required")
    with tempfile.TemporaryDirectory() as tmp:
        ref = os.path.join(tmp, "reference.docx")
        with open(ref, "wb") as fh:
            fh.write(subprocess.check_output(
                ["pandoc", "--print-default-data-file", "reference.docx"]))
        rewrite_zip(ref, {"word/styles.xml": patch_styles})
        subprocess.check_call(["pandoc", src, "-f", "gfm", "-o", out,
                               "--reference-doc", ref, "-M", "lang=ar", "-M", "dir=rtl"])
    rewrite_zip(out, {"word/document.xml": patch_document})
    print(out)


if __name__ == "__main__":
    main()
