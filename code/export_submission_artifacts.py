#!/usr/bin/env python3
from __future__ import annotations

import csv
import math
import re
import shutil
import subprocess
import tempfile
import textwrap
import zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parent
ANALYSIS = WORKSPACE.parent / "analysis"
WRITING = WORKSPACE.parent / "writing"
RESULTS = WRITING / "results_package_v6"
PRIMARY_META = ANALYSIS / "results" / "primary_meta_v5_20260324"
AF_META = ANALYSIS / "results" / "af_sensitivity_meta_v1_20260323"
REQUIRED_SOURCE_INPUTS = [
    ANALYSIS / "prisma_flow_current_v12.md",
    ANALYSIS / "side_search_triage.tsv",
    ANALYSIS / "fulltext_review_log.tsv",
    ANALYSIS / "risk_of_bias_quips_working_v5.tsv",
    ANALYSIS / "extraction_master_v5.tsv",
    RESULTS / "study_characteristics_table_v5.tsv",
    RESULTS / "main_results_table_v5.tsv",
    RESULTS / "sensitivity_results_table_v5.tsv",
    RESULTS / "evidence_gap_table_v5.tsv",
    RESULTS / "nonpooled_evidence_table_v5.tsv",
]

RENDERED = ROOT / "rendered"
UPLOAD_TABLES = ROOT / "upload_ready" / "tables"
UPLOAD_FIGURES = ROOT / "upload_ready" / "figures"
UPLOAD_SUPP = ROOT / "upload_ready" / "supplement"
FIG_DPI = 300
FIG_NATIVE_SCALE = 2
FIG_EXPORT_SCALE = 1
FONT_REGULAR_CANDIDATES = [
    "/usr/share/fonts/opentype/urw-base35/NimbusSans-Regular.otf",
    "/usr/share/fonts/type1/urw-base35/NimbusSans-Regular.t1",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
FONT_BOLD_CANDIDATES = [
    "/usr/share/fonts/opentype/urw-base35/NimbusSans-Bold.otf",
    "/usr/share/fonts/type1/urw-base35/NimbusSans-Bold.t1",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)

REF_BY_PMID = {
    "30376054": 7,
    "32298733": 8,
    "37418748": 9,
    "30590586": 11,
    "33433302": 12,
    "37947123": 13,
    "30243978": 14,
    "39120771": 15,
    "34648724": 16,
    "8889364": 18,
    "11734445": 19,
    "21724460": 20,
    "24342759": 21,
    "26612581": 22,
    "29277336": 23,
    "31081538": 24,
    "34677657": 25,
    "35131555": 26,
    "36642535": 27,
    "38773880": 28,
    "39206667": 29,
    "40008168": 30,
    "40056262": 31,
    "40818776": 32,
    "35304425": 33,
    "33394326": 34,
    "40132385": 35,
    "41733912": 36,
    "27464791": 37,
    "20339144": 38,
    "37656346": 39,
    "37734857": 40,
}


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def clear_dir(path: Path) -> None:
    ensure_dir(path)
    for child in path.iterdir():
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()


def validate_source_inputs() -> None:
    missing = [path for path in REQUIRED_SOURCE_INPUTS if not path.exists()]
    if missing:
        missing_list = "\n".join(f"- {path}" for path in missing)
        raise FileNotFoundError(
            "Submission-package export requires the complete workspace source tree.\n"
            f"Missing required inputs:\n{missing_list}\n"
            "Run export_submission_artifacts.py from the full osa_meta_20260323 workspace "
            "rather than from an isolated package copy."
        )


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = FONT_BOLD_CANDIDATES if bold else FONT_REGULAR_CANDIDATES
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def px(value: int | float, scale: float = 1.0) -> int:
    return int(round(value * scale))


def scale_box(box: tuple[int, int, int, int], scale: float = 1.0) -> tuple[int, int, int, int]:
    return tuple(px(v, scale) for v in box)


def save_image_outputs(image: Image.Image, pdf_path: Path, png_path: Path) -> None:
    ensure_dir(pdf_path.parent)
    ensure_dir(png_path.parent)
    export_image = image
    if FIG_EXPORT_SCALE != 1:
        export_image = image.resize(
            (image.width * FIG_EXPORT_SCALE, image.height * FIG_EXPORT_SCALE),
            resample=getattr(Image, "Resampling", Image).LANCZOS,
        )
    export_dpi = FIG_DPI * FIG_EXPORT_SCALE
    export_image.save(png_path, format="PNG", dpi=(export_dpi, export_dpi))
    image_rgb = export_image.convert("RGB")
    image_rgb.save(pdf_path, format="PDF", resolution=export_dpi)


_MEASURE_IMAGE = Image.new("RGB", (16, 16), "white")
_MEASURE_DRAW = ImageDraw.Draw(_MEASURE_IMAGE)


def rgb_triplet(color: str) -> tuple[float, float, float]:
    value = color.lstrip("#")
    if len(value) != 6:
        raise ValueError(f"Unsupported color: {color}")
    return tuple(int(value[i : i + 2], 16) / 255 for i in (0, 2, 4))


def ps_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def text_width_pixels(text: str, font: ImageFont.ImageFont) -> int:
    return _MEASURE_DRAW.textbbox((0, 0), text, font=font)[2]


class PSCanvas:
    def __init__(self, width: int, height: int):
        self.width = width
        self.height = height
        self.commands: list[str] = [
            "%!PS-Adobe-3.0 EPSF-3.0",
            f"%%BoundingBox: 0 0 {width} {height}",
            "%%LanguageLevel: 2",
            "1 setlinejoin 1 setlinecap",
        ]

    def _y(self, y_top: float) -> float:
        return self.height - y_top

    def _emit(self, line: str) -> None:
        self.commands.append(line)

    def set_stroke(self, color: str) -> None:
        r, g, b = rgb_triplet(color)
        self._emit(f"{r:.4f} {g:.4f} {b:.4f} setrgbcolor")

    def set_fill(self, color: str) -> None:
        self.set_stroke(color)

    def line(self, x1: float, y1: float, x2: float, y2: float, color: str, width: float = 1.0, dash: tuple[int, ...] | None = None) -> None:
        self._emit("gsave")
        self.set_stroke(color)
        self._emit(f"{width:.2f} setlinewidth")
        if dash:
            self._emit(f"[{' '.join(str(int(v)) for v in dash)}] 0 setdash")
        self._emit(f"newpath {x1:.2f} {self._y(y1):.2f} moveto {x2:.2f} {self._y(y2):.2f} lineto stroke")
        self._emit("grestore")

    def rect(self, x: float, y: float, w: float, h: float, *, fill: str | None = None, stroke: str | None = None, stroke_width: float = 1.0) -> None:
        yb = self.height - y - h
        if fill:
            self._emit("gsave")
            self.set_fill(fill)
            self._emit(f"newpath {x:.2f} {yb:.2f} moveto {w:.2f} 0 rlineto 0 {h:.2f} rlineto {-w:.2f} 0 rlineto closepath fill")
            self._emit("grestore")
        if stroke:
            self._emit("gsave")
            self.set_stroke(stroke)
            self._emit(f"{stroke_width:.2f} setlinewidth")
            self._emit(f"newpath {x:.2f} {yb:.2f} moveto {w:.2f} 0 rlineto 0 {h:.2f} rlineto {-w:.2f} 0 rlineto closepath stroke")
            self._emit("grestore")

    def polygon(self, points: list[tuple[float, float]], *, fill: str | None = None, stroke: str | None = None, stroke_width: float = 1.0) -> None:
        if not points:
            return
        path = [f"{points[0][0]:.2f} {self._y(points[0][1]):.2f} moveto"]
        for x, y in points[1:]:
            path.append(f"{x:.2f} {self._y(y):.2f} lineto")
        path.append("closepath")
        joined = " ".join(path)
        if fill:
            self._emit("gsave")
            self.set_fill(fill)
            self._emit(f"newpath {joined} fill")
            self._emit("grestore")
        if stroke:
            self._emit("gsave")
            self.set_stroke(stroke)
            self._emit(f"{stroke_width:.2f} setlinewidth")
            self._emit(f"newpath {joined} stroke")
            self._emit("grestore")

    def text(self, x: float, y: float, text: str, *, size: int, bold: bool = False, color: str = "#111827", anchor: str = "left") -> None:
        font = "NimbusSans-Bold" if bold else "NimbusSans-Regular"
        safe = ps_escape(text)
        baseline = self.height - y - size * 0.82
        self._emit("gsave")
        self.set_fill(color)
        self._emit(f"/{font} findfont {size} scalefont setfont")
        if anchor == "center":
            self._emit(f"{x:.2f} {baseline:.2f} moveto ({safe}) dup stringwidth pop 2 div neg 0 rmoveto show")
        elif anchor == "right":
            self._emit(f"{x:.2f} {baseline:.2f} moveto ({safe}) dup stringwidth pop neg 0 rmoveto show")
        else:
            self._emit(f"{x:.2f} {baseline:.2f} moveto ({safe}) show")
        self._emit("grestore")

    def wrapped_text(self, x: float, y: float, text: str, *, font: ImageFont.ImageFont, size: int, max_width: int, color: str, bold: bool = False, line_gap: int = 4, align: str = "left") -> int:
        line_height = _MEASURE_DRAW.textbbox((0, 0), "Ag", font=font)[3]
        for line in wrap_text_pixels(_MEASURE_DRAW, text, font, max_width):
            if align == "center":
                line_width = text_width_pixels(line, font)
                self.text(x + max_width / 2, y, line, size=size, bold=bold, color=color, anchor="center")
            else:
                self.text(x, y, line, size=size, bold=bold, color=color, anchor="left")
            y += line_height + line_gap
        return y

    def save_eps(self, path: Path) -> None:
        ensure_dir(path.parent)
        path.write_text("\n".join(self.commands + ["showpage", "%%EOF"]) + "\n", encoding="utf-8")


def render_eps_to_outputs(eps_path: Path, pdf_path: Path, png_path: Path, png_dpi: int = FIG_DPI) -> None:
    ensure_dir(pdf_path.parent)
    ensure_dir(png_path.parent)
    subprocess.run(
        [
            "gs",
            "-dSAFER",
            "-dBATCH",
            "-dNOPAUSE",
            "-dEPSCrop",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dCompressFonts=true",
            "-sDEVICE=pdfwrite",
            f"-sOutputFile={pdf_path}",
            str(eps_path),
        ],
        check=True,
    )
    if not pdf_has_embedded_fonts(pdf_path):
        raise RuntimeError(f"Embedded-font check failed for {pdf_path}")
    subprocess.run(
        [
            "gs",
            "-dSAFER",
            "-dBATCH",
            "-dNOPAUSE",
            "-dEPSCrop",
            "-dTextAlphaBits=4",
            "-dGraphicsAlphaBits=4",
            "-sDEVICE=pngalpha",
            f"-r{png_dpi}",
            f"-sOutputFile={png_path}",
            str(eps_path),
        ],
        check=True,
    )


def pdf_has_embedded_fonts(pdf_path: Path) -> bool:
    reader = PdfReader(str(pdf_path))
    for page in reader.pages:
        resources = page.get("/Resources")
        if not resources or "/Font" not in resources:
            continue
        for _, font_ref in resources["/Font"].items():
            font = font_ref.get_object()
            descriptor = font.get("/FontDescriptor")
            if descriptor:
                desc = descriptor.get_object()
                if any(key in desc for key in ("/FontFile", "/FontFile2", "/FontFile3")):
                    return True
            if "/DescendantFonts" in font:
                for descendant in font["/DescendantFonts"]:
                    desc_font = descendant.get_object()
                    descriptor = desc_font.get("/FontDescriptor")
                    if descriptor:
                        desc = descriptor.get_object()
                        if any(key in desc for key in ("/FontFile", "/FontFile2", "/FontFile3")):
                            return True
    return False


def read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    ensure_dir(path.parent)
    if not rows:
        raise ValueError(f"No rows for {path}")
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def escape_md(text: str) -> str:
    return (text or "").replace("|", "\\|").replace("\n", " ").strip()


def markdown_table_only(rows: list[dict[str, str]]) -> str:
    headers = list(rows[0].keys())
    lines: list[str] = []
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for row in rows:
        lines.append("| " + " | ".join(escape_md(row.get(h, "")) for h in headers) + " |")
    return "\n".join(lines)


def markdown_table_block(title: str, subtitle: str, rows: list[dict[str, str]], heading: str = "#") -> str:
    heading_line = f"{heading} {title}" if heading else f"**{title}**"
    return "\n".join(
        [
            heading_line,
            "",
            markdown_table_only(rows),
            "",
            subtitle,
            "",
        ]
    )


def write_markdown_table(path: Path, title: str, subtitle: str, rows: list[dict[str, str]]) -> None:
    lines = [markdown_table_block(title, subtitle, rows, heading="")]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_pandoc(src: Path, dst: Path) -> None:
    ensure_dir(dst.parent)
    subprocess.run(["pandoc", str(src), "-s", "-o", str(dst)], check=True)


def set_cell_width(cell: ET.Element, width: int) -> None:
    tc_pr = cell.find(qn(W_NS, "tcPr"))
    if tc_pr is None:
        tc_pr = ET.SubElement(cell, qn(W_NS, "tcPr"))
    tc_w = tc_pr.find(qn(W_NS, "tcW"))
    if tc_w is None:
        tc_w = ET.SubElement(tc_pr, qn(W_NS, "tcW"))
    tc_w.set(qn(W_NS, "type"), "dxa")
    tc_w.set(qn(W_NS, "w"), str(width))


def patch_docx(
    docx_path: Path,
    *,
    add_line_numbers: bool = False,
    add_page_numbers: bool = False,
    landscape: bool = False,
    document_role: str = "generic",
    repeat_table_headers: bool = False,
    narrow_margins: bool = False,
    table_widths: list[list[int]] | None = None,
    line_spacing_twips: int = 480,
    paragraph_before_twips: int = 0,
    paragraph_after_twips: int = 0,
    font_name: str = "Times New Roman",
    font_size_half_points: int = 24,
    heading_color: str = "000000",
    table_line_spacing_twips: int | None = None,
    table_font_size_half_points: int | None = None,
    prevent_row_splits: bool = False,
) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(tmp)

        styles_path = tmp / "word" / "styles.xml"
        settings_path = tmp / "word" / "settings.xml"
        document_path = tmp / "word" / "document.xml"
        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        content_types_path = tmp / "[Content_Types].xml"

        styles_root = ET.parse(styles_path).getroot()
        doc_defaults = styles_root.find(qn(W_NS, "docDefaults"))
        if doc_defaults is None:
            doc_defaults = ET.SubElement(styles_root, qn(W_NS, "docDefaults"))
        ppr_default = doc_defaults.find(qn(W_NS, "pPrDefault"))
        if ppr_default is None:
            ppr_default = ET.SubElement(doc_defaults, qn(W_NS, "pPrDefault"))
        ppr = ppr_default.find(qn(W_NS, "pPr"))
        if ppr is None:
            ppr = ET.SubElement(ppr_default, qn(W_NS, "pPr"))
        spacing = ppr.find(qn(W_NS, "spacing"))
        if spacing is None:
            spacing = ET.SubElement(ppr, qn(W_NS, "spacing"))
        spacing.set(qn(W_NS, "before"), str(paragraph_before_twips))
        spacing.set(qn(W_NS, "after"), str(paragraph_after_twips))
        spacing.set(qn(W_NS, "line"), str(line_spacing_twips))
        spacing.set(qn(W_NS, "lineRule"), "auto")
        rpr_default = doc_defaults.find(qn(W_NS, "rPrDefault"))
        if rpr_default is None:
            rpr_default = ET.SubElement(doc_defaults, qn(W_NS, "rPrDefault"))
        rpr = rpr_default.find(qn(W_NS, "rPr"))
        if rpr is None:
            rpr = ET.SubElement(rpr_default, qn(W_NS, "rPr"))
        rfonts = rpr.find(qn(W_NS, "rFonts"))
        if rfonts is None:
            rfonts = ET.SubElement(rpr, qn(W_NS, "rFonts"))
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(qn(W_NS, attr), font_name)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            rfonts.attrib.pop(qn(W_NS, attr), None)
        sz = rpr.find(qn(W_NS, "sz"))
        if sz is None:
            sz = ET.SubElement(rpr, qn(W_NS, "sz"))
        sz.set(qn(W_NS, "val"), str(font_size_half_points))
        sz_cs = rpr.find(qn(W_NS, "szCs"))
        if sz_cs is None:
            sz_cs = ET.SubElement(rpr, qn(W_NS, "szCs"))
        sz_cs.set(qn(W_NS, "val"), str(font_size_half_points))
        color = rpr.find(qn(W_NS, "color"))
        if color is None:
            color = ET.SubElement(rpr, qn(W_NS, "color"))
        color.set(qn(W_NS, "val"), "000000")
        for attr in ("themeColor", "themeTint", "themeShade"):
            color.attrib.pop(qn(W_NS, attr), None)

        heading_ids = {
            "Title",
            "Subtitle",
            "Heading1",
            "Heading2",
            "Heading3",
            "Heading4",
            "Heading5",
            "Heading6",
            "Heading7",
            "Heading8",
            "Heading9",
        }
        zero_space_ids = {"Normal", "BodyText", "FirstParagraph", "Compact", "Author", "Date"}
        for style in styles_root.findall(qn(W_NS, "style")):
            style_id = style.attrib.get(qn(W_NS, "styleId"), "")
            style_type = style.attrib.get(qn(W_NS, "type"), "")
            if style_type not in {"paragraph", "character"}:
                continue
            style_rpr = style.find(qn(W_NS, "rPr"))
            if style_rpr is None:
                style_rpr = ET.SubElement(style, qn(W_NS, "rPr"))
            style_rfonts = style_rpr.find(qn(W_NS, "rFonts"))
            if style_rfonts is None:
                style_rfonts = ET.SubElement(style_rpr, qn(W_NS, "rFonts"))
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                style_rfonts.set(qn(W_NS, attr), font_name)
            for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                style_rfonts.attrib.pop(qn(W_NS, attr), None)
            if style_id in heading_ids:
                style_color = style_rpr.find(qn(W_NS, "color"))
                if style_color is None:
                    style_color = ET.SubElement(style_rpr, qn(W_NS, "color"))
                style_color.set(qn(W_NS, "val"), heading_color)
                for attr in ("themeColor", "themeTint", "themeShade"):
                    style_color.attrib.pop(qn(W_NS, attr), None)
            if style_type == "paragraph" and style_id in zero_space_ids.union(heading_ids):
                style_ppr = style.find(qn(W_NS, "pPr"))
                if style_ppr is None:
                    style_ppr = ET.SubElement(style, qn(W_NS, "pPr"))
                style_spacing = style_ppr.find(qn(W_NS, "spacing"))
                if style_spacing is None:
                    style_spacing = ET.SubElement(style_ppr, qn(W_NS, "spacing"))
                if style_id in heading_ids:
                    style_spacing.set(qn(W_NS, "before"), "60")
                    style_spacing.set(qn(W_NS, "after"), "60")
                else:
                    style_spacing.set(qn(W_NS, "before"), str(paragraph_before_twips))
                    style_spacing.set(qn(W_NS, "after"), str(paragraph_after_twips))
        ET.ElementTree(styles_root).write(styles_path, encoding="utf-8", xml_declaration=True)

        settings_root = ET.parse(settings_path).getroot()
        if add_line_numbers:
            ln = settings_root.find(qn(W_NS, "lnNumType"))
            if ln is None:
                ln = ET.SubElement(settings_root, qn(W_NS, "lnNumType"))
            ln.set(qn(W_NS, "countBy"), "1")
            ln.set(qn(W_NS, "restart"), "continuous")
            ln.set(qn(W_NS, "distance"), "360")
        ET.ElementTree(settings_root).write(settings_path, encoding="utf-8", xml_declaration=True)

        footer_path = tmp / "word" / "footer1.xml"
        footer_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">
  <w:p>
    <w:pPr><w:jc w:val="right"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
"""
        next_rid = None
        if add_page_numbers:
            rels_root = ET.parse(rels_path).getroot()
            footer_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            footer_rels = [
                rel
                for rel in rels_root.findall(qn(PKG_REL_NS, "Relationship"))
                if rel.attrib.get("Type") == footer_rel_type
            ]
            if footer_rels:
                footer_rel = footer_rels[0]
                next_rid = footer_rel.attrib.get("Id")
                footer_target = footer_rel.attrib.get("Target", "footer1.xml")
                existing_footer_path = tmp / "word" / footer_target
                if not existing_footer_path.exists():
                    existing_footer_path.write_text(footer_xml, encoding="utf-8")
            else:
                footer_path.write_text(footer_xml, encoding="utf-8")
                existing_ids = []
                for rel in rels_root.findall(qn(PKG_REL_NS, "Relationship")):
                    rid = rel.attrib.get("Id", "")
                    if rid.startswith("rId") and rid[3:].isdigit():
                        existing_ids.append(int(rid[3:]))
                next_rid = f"rId{max(existing_ids or [0]) + 1}"
                footer_rel = ET.SubElement(rels_root, qn(PKG_REL_NS, "Relationship"))
                footer_rel.set("Id", next_rid)
                footer_rel.set("Type", footer_rel_type)
                footer_rel.set("Target", "footer1.xml")
                ET.ElementTree(rels_root).write(rels_path, encoding="utf-8", xml_declaration=True)

                ct_root = ET.parse(content_types_path).getroot()
                if not any(node.attrib.get("PartName") == "/word/footer1.xml" for node in ct_root.findall(qn(CT_NS, "Override"))):
                    override = ET.SubElement(ct_root, qn(CT_NS, "Override"))
                    override.set("PartName", "/word/footer1.xml")
                    override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
                ET.ElementTree(ct_root).write(content_types_path, encoding="utf-8", xml_declaration=True)

        doc_root = ET.parse(document_path).getroot()
        body = doc_root.find(qn(W_NS, "body"))
        if body is None:
            raise RuntimeError("No body found in generated docx")

        def replace_paragraph_with_lines(para: ET.Element, lines: list[str]) -> None:
            ppr = para.find(qn(W_NS, "pPr"))
            for child in list(para):
                if child.tag != qn(W_NS, "pPr"):
                    para.remove(child)
            if ppr is None:
                ppr = ET.SubElement(para, qn(W_NS, "pPr"))
            for idx, line in enumerate(lines):
                run = ET.SubElement(para, qn(W_NS, "r"))
                t = ET.SubElement(run, qn(W_NS, "t"))
                if line.startswith(" ") or line.endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = line
                if idx < len(lines) - 1:
                    br_run = ET.SubElement(para, qn(W_NS, "r"))
                    ET.SubElement(br_run, qn(W_NS, "br"))

        sects = doc_root.findall(f".//{qn(W_NS, 'sectPr')}")
        if not sects:
            raise RuntimeError("No sectPr found in generated docx")
        sect = sects[-1]
        if add_page_numbers and next_rid:
            existing = None
            for child in sect.findall(qn(W_NS, "footerReference")):
                if child.attrib.get(qn(W_NS, "type")) == "default":
                    existing = child
                    break
            if existing is None:
                existing = ET.Element(qn(W_NS, "footerReference"))
                existing.set(qn(W_NS, "type"), "default")
                sect.insert(0, existing)
            existing.set(qn(R_NS, "id"), next_rid)
        if add_line_numbers:
            ln = sect.find(qn(W_NS, "lnNumType"))
            if ln is None:
                ln = ET.SubElement(sect, qn(W_NS, "lnNumType"))
            ln.set(qn(W_NS, "countBy"), "1")
            ln.set(qn(W_NS, "restart"), "continuous")
            ln.set(qn(W_NS, "distance"), "360")
        pg_sz = sect.find(qn(W_NS, "pgSz"))
        if pg_sz is None:
            pg_sz = ET.SubElement(sect, qn(W_NS, "pgSz"))
        if landscape:
            pg_sz.set(qn(W_NS, "w"), "16838")
            pg_sz.set(qn(W_NS, "h"), "11906")
            pg_sz.set(qn(W_NS, "orient"), "landscape")
        else:
            pg_sz.set(qn(W_NS, "w"), "12240")
            pg_sz.set(qn(W_NS, "h"), "15840")
            pg_sz.attrib.pop(qn(W_NS, "orient"), None)
        pg_mar = sect.find(qn(W_NS, "pgMar"))
        if pg_mar is None:
            pg_mar = ET.SubElement(sect, qn(W_NS, "pgMar"))
        margin_values = (
            {
                "top": "720",
                "right": "720",
                "bottom": "720",
                "left": "720",
                "header": "360",
                "footer": "360",
                "gutter": "0",
            }
            if narrow_margins
            else {
                "top": "1440",
                "right": "1440",
                "bottom": "1440",
                "left": "1440",
                "header": "720",
                "footer": "720",
                "gutter": "0",
            }
        )
        for side, value in margin_values.items():
            pg_mar.set(qn(W_NS, side), value)

        if document_role == "manuscript":
            children = list(body)

            def para_text(node: ET.Element) -> str:
                return "".join(t.text or "" for t in node.findall(f".//{qn(W_NS, 't')}")).strip()

            def is_page_break_para(node: ET.Element) -> bool:
                return node.tag == qn(W_NS, "p") and node.find(f".//{qn(W_NS, 'br')}[@{qn(W_NS, 'type')}='page']") is not None

            abstract_idx = next(
                (
                    idx
                    for idx, child in enumerate(children)
                    if child.tag == qn(W_NS, "p") and para_text(child) == "Abstract"
                ),
                None,
            )
            if abstract_idx is not None and abstract_idx > 0 and is_page_break_para(children[abstract_idx - 1]):
                body.remove(children[abstract_idx - 1])
                abstract_para = children[abstract_idx]
                abstract_ppr = abstract_para.find(qn(W_NS, "pPr"))
                if abstract_ppr is None:
                    abstract_ppr = ET.SubElement(abstract_para, qn(W_NS, "pPr"))
                if abstract_ppr.find(qn(W_NS, "pageBreakBefore")) is None:
                    ET.SubElement(abstract_ppr, qn(W_NS, "pageBreakBefore"))

            if landscape:
                final_pg_sz = sect.find(qn(W_NS, "pgSz"))
                if final_pg_sz is None:
                    final_pg_sz = ET.SubElement(sect, qn(W_NS, "pgSz"))
                final_pg_sz.set(qn(W_NS, "w"), "12240")
                final_pg_sz.set(qn(W_NS, "h"), "15840")
                final_pg_sz.attrib.pop(qn(W_NS, "orient"), None)
                final_pg_mar = sect.find(qn(W_NS, "pgMar"))
                if final_pg_mar is None:
                    final_pg_mar = ET.SubElement(sect, qn(W_NS, "pgMar"))
                for side, value in {
                    "top": "1440",
                    "right": "1440",
                    "bottom": "1440",
                    "left": "1440",
                    "header": "720",
                    "footer": "720",
                    "gutter": "0",
                }.items():
                    final_pg_mar.set(qn(W_NS, side), value)

                children = list(body)
                first_table_idx = next(
                    (
                        idx
                        for idx, child in enumerate(children)
                        if child.tag == qn(W_NS, "p") and para_text(child).startswith("Table 1.")
                    ),
                    None,
                )
                second_table_idx = next(
                    (
                        idx
                        for idx, child in enumerate(children)
                        if child.tag == qn(W_NS, "p") and para_text(child).startswith("Table 2.")
                    ),
                    None,
                )

                def build_section_break(orientation: str) -> ET.Element:
                    break_para = ET.Element(qn(W_NS, "p"))
                    break_ppr = ET.SubElement(break_para, qn(W_NS, "pPr"))
                    break_sect = ET.Element(qn(W_NS, "sectPr"))
                    break_type = break_sect.find(qn(W_NS, "type"))
                    if break_type is None:
                        break_type = ET.SubElement(break_sect, qn(W_NS, "type"))
                    break_type.set(qn(W_NS, "val"), "nextPage")
                    break_pg_sz = break_sect.find(qn(W_NS, "pgSz"))
                    if break_pg_sz is None:
                        break_pg_sz = ET.SubElement(break_sect, qn(W_NS, "pgSz"))
                    break_pg_mar = break_sect.find(qn(W_NS, "pgMar"))
                    if break_pg_mar is None:
                        break_pg_mar = ET.SubElement(break_sect, qn(W_NS, "pgMar"))
                    if orientation == "landscape":
                        break_pg_sz.set(qn(W_NS, "w"), "16838")
                        break_pg_sz.set(qn(W_NS, "h"), "11906")
                        break_pg_sz.set(qn(W_NS, "orient"), "landscape")
                        margins = {
                            "top": "720",
                            "right": "720",
                            "bottom": "720",
                            "left": "720",
                            "header": "360",
                            "footer": "360",
                            "gutter": "0",
                        }
                    else:
                        break_pg_sz.set(qn(W_NS, "w"), "12240")
                        break_pg_sz.set(qn(W_NS, "h"), "15840")
                        break_pg_sz.attrib.pop(qn(W_NS, "orient"), None)
                        margins = {
                            "top": "1440",
                            "right": "1440",
                            "bottom": "1440",
                            "left": "1440",
                            "header": "720",
                            "footer": "720",
                            "gutter": "0",
                        }
                    for side, value in margins.items():
                        break_pg_mar.set(qn(W_NS, side), value)
                    break_ppr.append(break_sect)
                    return break_para

                if first_table_idx is not None:
                    insert_idx = first_table_idx
                    if insert_idx > 0 and is_page_break_para(children[insert_idx - 1]):
                        body.remove(children[insert_idx - 1])
                        insert_idx -= 1
                    body.insert(insert_idx, build_section_break("portrait"))

                children = list(body)
                second_table_idx = next(
                    (
                        idx
                        for idx, child in enumerate(children)
                        if child.tag == qn(W_NS, "p") and para_text(child).startswith("Table 2.")
                    ),
                    None,
                )
                if second_table_idx is not None:
                    insert_idx = second_table_idx
                    if insert_idx > 0 and is_page_break_para(children[insert_idx - 1]):
                        body.remove(children[insert_idx - 1])
                        insert_idx -= 1
                    body.insert(insert_idx, build_section_break("landscape"))

            body_children = list(body)
            for idx in range(1, len(body_children) - 1):
                child = body_children[idx]
                if child.tag != qn(W_NS, "p"):
                    continue
                if para_text(child):
                    continue
                prev_node = body_children[idx - 1]
                next_node = body_children[idx + 1]
                if (
                    prev_node.tag == qn(W_NS, "p")
                    and para_text(prev_node).startswith("Table ")
                    and next_node.tag == qn(W_NS, "tbl")
                ):
                    body.remove(child)

        tables = doc_root.findall(f".//{qn(W_NS, 'tbl')}")
        for idx, tbl in enumerate(tables):
            if repeat_table_headers:
                first_row = tbl.find(qn(W_NS, "tr"))
                if first_row is not None:
                    tr_pr = first_row.find(qn(W_NS, "trPr"))
                    if tr_pr is None:
                        tr_pr = ET.SubElement(first_row, qn(W_NS, "trPr"))
                    if tr_pr.find(qn(W_NS, "tblHeader")) is None:
                        ET.SubElement(tr_pr, qn(W_NS, "tblHeader"))
            if prevent_row_splits:
                for tr in tbl.findall(qn(W_NS, "tr")):
                    tr_pr = tr.find(qn(W_NS, "trPr"))
                    if tr_pr is None:
                        tr_pr = ET.SubElement(tr, qn(W_NS, "trPr"))
                    if tr_pr.find(qn(W_NS, "cantSplit")) is None:
                        ET.SubElement(tr_pr, qn(W_NS, "cantSplit"))
            if table_widths and idx < len(table_widths):
                widths = table_widths[idx]
                tbl_pr = tbl.find(qn(W_NS, "tblPr"))
                if tbl_pr is None:
                    tbl_pr = ET.SubElement(tbl, qn(W_NS, "tblPr"))
                tbl_w = tbl_pr.find(qn(W_NS, "tblW"))
                if tbl_w is None:
                    tbl_w = ET.SubElement(tbl_pr, qn(W_NS, "tblW"))
                tbl_w.set(qn(W_NS, "type"), "dxa")
                tbl_w.set(qn(W_NS, "w"), str(sum(widths)))
                tbl_layout = tbl_pr.find(qn(W_NS, "tblLayout"))
                if tbl_layout is None:
                    tbl_layout = ET.SubElement(tbl_pr, qn(W_NS, "tblLayout"))
                tbl_layout.set(qn(W_NS, "type"), "fixed")
                tbl_grid = tbl.find(qn(W_NS, "tblGrid"))
                if tbl_grid is not None:
                    for child in list(tbl_grid):
                        tbl_grid.remove(child)
                else:
                    tbl_grid = ET.Element(qn(W_NS, "tblGrid"))
                    tbl.insert(1, tbl_grid)
                for width in widths:
                    col = ET.SubElement(tbl_grid, qn(W_NS, "gridCol"))
                    col.set(qn(W_NS, "w"), str(width))
                for tr in tbl.findall(qn(W_NS, "tr")):
                    for width, cell in zip(widths, tr.findall(qn(W_NS, "tc"))):
                        set_cell_width(cell, width)
            if table_line_spacing_twips is not None or table_font_size_half_points is not None:
                for para in tbl.findall(f".//{qn(W_NS, 'p')}"):
                    ppr = para.find(qn(W_NS, "pPr"))
                    if ppr is None:
                        ppr = ET.SubElement(para, qn(W_NS, "pPr"))
                    spacing = ppr.find(qn(W_NS, "spacing"))
                    if spacing is None:
                        spacing = ET.SubElement(ppr, qn(W_NS, "spacing"))
                    spacing.set(qn(W_NS, "before"), "0")
                    spacing.set(qn(W_NS, "after"), "0")
                    if table_line_spacing_twips is not None:
                        spacing.set(qn(W_NS, "line"), str(table_line_spacing_twips))
                        spacing.set(qn(W_NS, "lineRule"), "auto")
                if table_font_size_half_points is not None:
                    for run in tbl.findall(f".//{qn(W_NS, 'r')}"):
                        rpr = run.find(qn(W_NS, "rPr"))
                        if rpr is None:
                            rpr = ET.SubElement(run, qn(W_NS, "rPr"))
                        sz = rpr.find(qn(W_NS, "sz"))
                        if sz is None:
                            sz = ET.SubElement(rpr, qn(W_NS, "sz"))
                        sz.set(qn(W_NS, "val"), str(table_font_size_half_points))
                        sz_cs = rpr.find(qn(W_NS, "szCs"))
                        if sz_cs is None:
                            sz_cs = ET.SubElement(rpr, qn(W_NS, "szCs"))
                        sz_cs.set(qn(W_NS, "val"), str(table_font_size_half_points))
        table_title_prefixes = (
            "Table 1.",
            "Table 2.",
            "Table 3.",
        )
        table_note_prefixes = (
            "Table 1 summarizes retained articles",
            "Each line pools two cohort-specific estimates",
            "Focused competing-risk and exploratory harmonization checks",
            "Article-level cohort map.",
        )
        no_top_space_headings = {
            "Abstract",
            "Figure legends",
            "Additional files",
        }
        title_page_prefixes = (
            "Title:",
            "Running title:",
            "Article type:",
            "Authors:",
            "Affiliations:",
            "Corresponding author:",
            "Keywords:",
        )
        cover_signature_lines = {
            "Sincerely,",
            "Bing Li, MD, PhD",
            "Professor of Medicine",
            "Director of Respiratory and Critical Care Medicine",
            "Shanghai Pulmonary Hospital",
            "School of Medicine, Tongji University",
            "507 Zhengmin Road, Shanghai 200433, China",
            "Phone: +86-021-65115006",
            "Email: libing044162@163.com",
        }
        for para in doc_root.findall(f".//{qn(W_NS, 'p')}"):
            text = "".join(node.text or "" for node in para.findall(f".//{qn(W_NS, 't')}")).strip()
            if document_role == "manuscript" and text.startswith("Affiliations:"):
                content = text[len("Affiliations:") :].strip()
                parts = [part.strip() for part in re.split(r"(?=\d\.\s*)", content) if part.strip()]
                if parts:
                    replace_paragraph_with_lines(para, ["Affiliations:"] + parts)
                    text = "".join(node.text or "" for node in para.findall(f".//{qn(W_NS, 't')}")).strip()
            if document_role == "manuscript" and text.startswith("Corresponding author:"):
                content = text[len("Corresponding author:") :].strip()
                parts = [part.strip() for part in re.split(r"(?=Department of|Shanghai Pulmonary Hospital|507 Zhengmin Road|Phone:|Email:)", content) if part.strip()]
                if parts:
                    replace_paragraph_with_lines(para, ["Corresponding author:"] + parts)
                    text = "".join(node.text or "" for node in para.findall(f".//{qn(W_NS, 't')}")).strip()
            ppr = para.find(qn(W_NS, "pPr"))
            if ppr is None:
                ppr = ET.SubElement(para, qn(W_NS, "pPr"))
            spacing = ppr.find(qn(W_NS, "spacing"))
            if spacing is None:
                spacing = ET.SubElement(ppr, qn(W_NS, "spacing"))
            if document_role == "manuscript" and text in no_top_space_headings:
                spacing.set(qn(W_NS, "before"), "0")
                spacing.set(qn(W_NS, "after"), "60")
                spacing.set(qn(W_NS, "line"), "240")
                spacing.set(qn(W_NS, "lineRule"), "auto")
                continue
            if text.startswith("Abbreviations:"):
                spacing.set(qn(W_NS, "before"), "40")
                spacing.set(qn(W_NS, "after"), "0")
                spacing.set(qn(W_NS, "line"), "180")
                spacing.set(qn(W_NS, "lineRule"), "auto")
                for run in para.findall(qn(W_NS, "r")):
                    rpr = run.find(qn(W_NS, "rPr"))
                    if rpr is None:
                        rpr = ET.SubElement(run, qn(W_NS, "rPr"))
                    sz = rpr.find(qn(W_NS, "sz"))
                    if sz is None:
                        sz = ET.SubElement(rpr, qn(W_NS, "sz"))
                    sz.set(qn(W_NS, "val"), "18")
                    sz_cs = rpr.find(qn(W_NS, "szCs"))
                    if sz_cs is None:
                        sz_cs = ET.SubElement(rpr, qn(W_NS, "szCs"))
                    sz_cs.set(qn(W_NS, "val"), "18")
                continue
            if document_role == "manuscript" and any(text.startswith(prefix) for prefix in title_page_prefixes):
                spacing.set(qn(W_NS, "before"), "0")
                spacing.set(qn(W_NS, "after"), "120")
                spacing.set(qn(W_NS, "line"), "480")
                spacing.set(qn(W_NS, "lineRule"), "auto")
                continue
            if any(text.startswith(prefix) for prefix in table_title_prefixes):
                spacing.set(qn(W_NS, "before"), "0")
                spacing.set(qn(W_NS, "after"), "0")
                spacing.set(qn(W_NS, "line"), "200")
                spacing.set(qn(W_NS, "lineRule"), "auto")
                for run in para.findall(qn(W_NS, "r")):
                    rpr = run.find(qn(W_NS, "rPr"))
                    if rpr is None:
                        rpr = ET.SubElement(run, qn(W_NS, "rPr"))
                    if rpr.find(qn(W_NS, "b")) is None:
                        ET.SubElement(rpr, qn(W_NS, "b"))
                    sz = rpr.find(qn(W_NS, "sz"))
                    if sz is None:
                        sz = ET.SubElement(rpr, qn(W_NS, "sz"))
                    sz.set(qn(W_NS, "val"), "22")
                    sz_cs = rpr.find(qn(W_NS, "szCs"))
                    if sz_cs is None:
                        sz_cs = ET.SubElement(rpr, qn(W_NS, "szCs"))
                    sz_cs.set(qn(W_NS, "val"), "22")
                continue
            if any(text.startswith(prefix) for prefix in table_note_prefixes):
                spacing.set(qn(W_NS, "before"), "40")
                spacing.set(qn(W_NS, "after"), "0")
                spacing.set(qn(W_NS, "line"), "200")
                spacing.set(qn(W_NS, "lineRule"), "auto")
                for run in para.findall(qn(W_NS, "r")):
                    rpr = run.find(qn(W_NS, "rPr"))
                    if rpr is None:
                        rpr = ET.SubElement(run, qn(W_NS, "rPr"))
                    sz = rpr.find(qn(W_NS, "sz"))
                    if sz is None:
                        sz = ET.SubElement(rpr, qn(W_NS, "sz"))
                    sz.set(qn(W_NS, "val"), "20")
                    sz_cs = rpr.find(qn(W_NS, "szCs"))
                    if sz_cs is None:
                        sz_cs = ET.SubElement(rpr, qn(W_NS, "szCs"))
                    sz_cs.set(qn(W_NS, "val"), "20")
                continue
            if document_role == "cover_letter":
                if text == "Sincerely,":
                    spacing.set(qn(W_NS, "before"), "80")
                    spacing.set(qn(W_NS, "after"), "72")
                    spacing.set(qn(W_NS, "line"), "320")
                    spacing.set(qn(W_NS, "lineRule"), "auto")
                    continue
                if text in cover_signature_lines:
                    spacing.set(qn(W_NS, "before"), "0")
                    spacing.set(qn(W_NS, "after"), "6")
                    spacing.set(qn(W_NS, "line"), "308")
                    spacing.set(qn(W_NS, "lineRule"), "auto")
                    continue
        ET.ElementTree(doc_root).write(document_path, encoding="utf-8", xml_declaration=True)

        tmp_zip = docx_path.with_suffix(".tmp.docx")
        with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(tmp.rglob("*")):
                if file_path.is_file():
                    zf.write(file_path, file_path.relative_to(tmp))
        tmp_zip.replace(docx_path)


def patch_rtf(rtf_path: Path) -> None:
    text = rtf_path.read_text(encoding="utf-8", errors="ignore")
    if "{\\footer" not in text:
        text = text.replace("{\\rtf1", "{\\rtf1{\\footer\\pard\\plain\\qr\\chpgn\\par}", 1)
    if "\\linemod1" not in text:
        if "\\viewkind4" in text:
            text = text.replace("\\viewkind4", "\\viewkind4\\linemod1\\linex360\\linestarts1", 1)
        else:
            text = text.replace("\\deff0", "\\deff0\\linemod1\\linex360\\linestarts1", 1)
    text = text.replace("\\pard", "\\pard\\sl480\\slmult1")
    rtf_path.write_text(text, encoding="utf-8")


def rebuild_docx_from_rtf_via_soffice(rtf_path: Path, docx_path: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return
    with tempfile.TemporaryDirectory() as tmpdir_name:
        tmpdir = Path(tmpdir_name)
        staged_rtf = tmpdir / rtf_path.name
        staged_docx = tmpdir / docx_path.name
        staged_pdf = tmpdir / f"{docx_path.stem}.pdf"
        shutil.copy2(rtf_path, staged_rtf)
        first = subprocess.run(
            [soffice, "--headless", "--convert-to", "docx", "--outdir", str(tmpdir), str(staged_rtf)],
            check=False,
            capture_output=True,
            text=True,
        )
        if first.returncode != 0 or not staged_docx.exists():
            raise RuntimeError(
                "LibreOffice RTF-to-DOCX conversion failed for canonical manuscript rebuild: "
                f"{first.stdout.strip()} {first.stderr.strip()}".strip()
            )
        second = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmpdir), str(staged_docx)],
            check=False,
            capture_output=True,
            text=True,
        )
        if second.returncode != 0 or not staged_pdf.exists():
            raise RuntimeError(
                "LibreOffice DOCX round-trip validation failed for canonical manuscript rebuild: "
                f"{second.stdout.strip()} {second.stderr.strip()}".strip()
            )
        shutil.copy2(staged_docx, docx_path)


def short_text(text: str, limit: int = 140) -> str:
    text = re.sub(r"\s+", " ", (text or "").strip())
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "…"


def round_half_up(value: str | float, digits: int = 2) -> Decimal:
    quant = Decimal("1").scaleb(-digits)
    return Decimal(str(float(value))).quantize(quant, rounding=ROUND_HALF_UP)


def fmt_num(value: str | float, digits: int = 2) -> str:
    return f"{round_half_up(value, digits):.{digits}f}"


def fmt_effect_label(effect: str | float, lower: str | float, upper: str | float, digits: int = 2) -> str:
    return f"HR {fmt_num(effect, digits)} ({fmt_num(lower, digits)}-{fmt_num(upper, digits)})"


def fmt_i2(value: str | float) -> str:
    num = float(value)
    if abs(num - round(num)) < 1e-9:
        return str(int(round(num)))
    return f"{round_half_up(num, 1):.1f}"


def ref_suffix(pmid: str = "") -> str:
    ref = REF_BY_PMID.get(str(pmid or "").strip())
    return f" [{ref}]" if ref else ""


def normalize_followup(text: str) -> str:
    value = re.sub(r"\s+", " ", (text or "").strip())
    if not value:
        return "NR"
    if value == "5+":
        return ">=5"
    if value == "0.08":
        return "0.08 (~30 d)"
    if value == "in-hospital postoperative follow-up":
        return "in-hospital"
    match = re.fullmatch(r"([\d.]+)\s+(mean|median)", value)
    if match:
        return f"{match.group(2)} {match.group(1)}"
    if "; " in value:
        parts = [part.strip() for part in value.split(";") if part.strip()]
        numeric_parts = []
        suffix = None
        for part in parts:
            token_match = re.fullmatch(r"([\d.]+)(?:\s+(mean|median))?", part)
            if token_match:
                numeric_parts.append(token_match.group(1))
                suffix = suffix or token_match.group(2)
            elif part.endswith("mean in overall DM subgroup"):
                numeric_parts.append(part.split()[0])
                suffix = suffix or "mean"
        if numeric_parts:
            if len(set(numeric_parts)) == 1:
                base = numeric_parts[0]
            else:
                ordered = sorted(float(x) for x in numeric_parts)
                base = f"{ordered[0]:.1f}-{ordered[-1]:.1f}".rstrip("0").rstrip(".")
            return f"{suffix} {base}" if suffix else base
    return value


def compact_metric_family(text: str) -> str:
    return (text or "").replace("; ", " / ")


def compact_outcome_family(text: str) -> str:
    value = re.sub(r"\s+", " ", (text or "").strip())
    replacements = {
        "post-CABG atrial fibrillation": "post-CABG AF",
        "incident hospitalized atrial fibrillation": "incident AF hospitalization",
        "composite readmission or mortality": "readmission or mortality",
        "composite mortality or HF readmission": "mortality or HF readmission",
        "incident heart failure": "incident HF",
        "atrial tachyarrhythmia recurrence": "tachyarrhythmia recurrence",
        "incident hard CVD; all-cause mortality; incident CVD; CVD mortality": "hard CVD; CVD/all-cause mortality",
        "CVD mortality; incident CVD": "CVD mortality; incident CVD",
        "CVD mortality; MACE": "CVD mortality; MACE",
        "ODI; rHB": "ODI / rHB",
        "ODI; T90": "ODI / T90",
        "ODI; nadir SpO2": "ODI / nadir SpO2",
        "HB; T90": "HB / T90",
        "T90; nadir SpO2": "T90 / nadir SpO2",
    }
    return replacements.get(value, value)


def compact_analytic_n(text: str) -> str:
    value = re.sub(r"\s+", " ", (text or "").strip())
    if not value:
        return "NR"
    if "not reported for treated-DM subgroup" in value:
        return "453; treated-DM NR"
    nums = re.findall(r"\d+", value.replace(",", ""))
    if not nums:
        return value
    if len(nums) == 1:
        return f"{int(nums[0]):,}"
    ints = sorted(int(n) for n in nums)
    if len(set(ints)) == 1:
        return str(ints[0])
    return f"{ints[0]}-{ints[-1]}"


def cohort_class_label(text: str) -> str:
    lower = (text or "").lower()
    clinical_keys = [
        "diagnosed with osa",
        "diagnosed osa",
        "suspected osa",
        "clinical osa",
        "clinical sleep-study",
        "clinical sleep study",
        "investigated for osa",
        "clinical suspicion of osa",
        "osa-referred",
        "referral",
        "newly diagnosed",
        "patients with osa",
    ]
    community_keys = [
        "community-based",
        "community-dwelling",
        "community cohort",
        "community adults",
        "community older",
    ]
    specialized_keys = [
        "coronary",
        "atrial fibrillation",
        "heart failure",
        "acute coronary syndrome",
        "surgical",
        "surgery",
        "perioperative",
        "cabg",
        "catheter ablation",
        "hemodialysis",
        "diabetes",
    ]
    if any(key in lower for key in clinical_keys):
        return "Clinical OSA/referral"
    if any(key in lower for key in community_keys):
        return "Community physiology"
    if any(key in lower for key in specialized_keys):
        return "Specialized cardiovascular/surgical"
    return "Specialized cardiovascular/surgical"


def cohort_class_short(text: str) -> str:
    mapping = {
        "Clinical OSA/referral": "Clin",
        "Community physiology": "Comm",
        "Specialized cardiovascular/surgical": "Spec",
    }
    return mapping.get(cohort_class_label(text), cohort_class_label(text))


def study_label(citation_key: str, pmid: str = "") -> str:
    parts = citation_key.split("_")
    year = next((p for p in parts if p.isdigit() and len(p) == 4), "")
    surname = parts[0]
    if surname == "BrianconMarjollet" and year == "2021":
        surname = "Blanchard"
    surname = (
        surname.replace("BrianconMarjollet", "Briancon-Marjollet")
        .replace("RiveraLopez", "Rivera-Lopez")
        .replace("HenriquezBeltran", "Henríquez-Beltrán")
    )
    if year:
        return f"{surname} {year}"
    if pmid:
        return f"{surname} PMID {pmid}"
    return surname


def scale_label(scale: str) -> str:
    mapping = {
        "categorical_high_vs_low": "High vs low",
        "categorical_threshold": "Threshold contrast",
        "continuous_log": "Continuous log-scale",
        "continuous_spline_contrast": "Spline-derived contrast",
        "continuous_standardized": "Per 1 SD",
        "continuous_fixed_increment": "Per reported increment",
        "continuous_fixed_increment_hour": "Per hour",
        "continuous_raw_unit": "Per reported unit",
    }
    return mapping.get(scale, scale.replace("_", " "))


def analysis_layer_label(review_role: str, main_rows: str, sens_rows: str) -> str:
    main_n = int(main_rows or "0")
    sens_n = int(sens_rows or "0")
    role = review_role or ""
    if main_n > 0:
        if "primary_main" in role:
            return "Primary pooled"
        if "primary_main_plus" in role:
            return "Primary + overlap-sensitive"
        if "primary_overlap" in role:
            return "Overlap-sensitive"
        if "single_study" in role:
            return "Single-study anchor"
        return "Primary evidence"
    if sens_n > 0:
        return "Sensitivity/comparator"
    return "Narrative / non-pooled only"


def sensitivity_label(sensitivity_type: str, analysis_cell_id: str) -> str:
    if sensitivity_type == "manual_pooled" and "competing_risk" in analysis_cell_id:
        return "SASHB and incident heart failure: competing-risk sensitivity"
    if sensitivity_type == "primary_exploratory_harmonization":
        return "T90 and incident AF: exploratory harmonized analysis"
    if sensitivity_type == "precision_check_sensitivity":
        return "T90 and incident AF: precision-check harmonization"
    if sensitivity_type == "leave_one_out_range":
        if analysis_cell_id == "HB__CVD_mortality__categorical_high_vs_low":
            return "Leave-one-out range: HB and CVD mortality (high vs low)"
        if analysis_cell_id == "HB__CVD_mortality__continuous_log":
            return "Leave-one-out range: HB and CVD mortality (continuous log)"
        if "all_cause_mortality" in analysis_cell_id or "all-cause_mortality" in analysis_cell_id:
            return "Leave-one-out range: HB and all-cause mortality"
        if "incident_heart_failure" in analysis_cell_id:
            return "Leave-one-out range: SASHB and incident heart failure"
    return analysis_cell_id.replace("__", " | ")


def primary_context_label(row: dict[str, str]) -> str:
    cell = row.get("analysis_cell_id", "")
    if cell == "HB__CVD_mortality__categorical_high_vs_low":
        return "Azarbarzin 2019 (MrOS; SHHS)"
    if cell == "HB__CVD_mortality__continuous_log":
        return "Azarbarzin 2019 (MrOS; SHHS)"
    if cell == "HB__all-cause_mortality__continuous_standardized":
        return "Labarca 2023 (MESA; MrOS)"
    if cell == "SASHB__incident_heart_failure__continuous_standardized":
        return "Azarbarzin 2020 (MrOS men; SHHS men)"
    return short_text(row.get("cell_rationale", ""), 72)


def clean_sensitivity_note(row: dict[str, str]) -> str:
    sensitivity_type = row.get("sensitivity_type", "")
    analysis_cell_id = row.get("analysis_cell_id", "")
    if sensitivity_type == "leave_one_out_range" and analysis_cell_id == "HB__CVD_mortality__categorical_high_vs_low":
        return "Single-study estimates remained positive after omitting either HB high-vs-low cohort."
    if sensitivity_type == "leave_one_out_range" and analysis_cell_id == "HB__CVD_mortality__continuous_log":
        return "Single-study estimates remained positive after omitting either continuous-log HB cohort."
    if sensitivity_type == "leave_one_out_range" and analysis_cell_id == "HB__all-cause_mortality__continuous_standardized":
        return "Direction was preserved after omitting either all-cause mortality cohort."
    if sensitivity_type == "leave_one_out_range" and analysis_cell_id == "SASHB__incident_heart_failure__continuous_standardized":
        return "Direction was preserved after omitting either male-subgroup SASHB cohort."
    if sensitivity_type == "manual_pooled" and "incident_heart_failure" in analysis_cell_id:
        return "Competing-risk modeling yielded a near-identical estimate in the same 2 male SASHB cohorts."
    if sensitivity_type == "primary_exploratory_harmonization":
        return "Rounded-CI rescaling of the Blanchard T90 row to a per-10% scale yielded a similar exploratory AF estimate."
    if sensitivity_type == "precision_check_sensitivity":
        return "Alternative SE derivation for the same Blanchard rescaling yielded a near-identical exploratory AF estimate."
    return first_sentence(row.get("notes", ""))


def pooling_reason_label(value: str) -> str:
    mapping = {
        "continuous_spline_contrast_main": "Spline-derived contrast without a shared poolable scale",
        "continuous_standardized_main": "Only one directly comparable cohort family",
        "single_publication_dualcohort_main": "Single publication with two cohort-specific estimates retained outside the pooled core",
        "continuous_log_main": "Single-publication cohort family or overlap-sensitive row",
        "ahi_adjusted_sensitivity": "Alternate model from the same cohort family",
        "sensitivity_specialized_secondary_prevention": "Specialized secondary-prevention cohort",
        "sensitivity_overlap_shhs_or": "Overlap-prone SHHS reuse with OR-based model",
        "sensitivity_specialized_postmi": "Specialized recent-post-MI cohort",
        "narrative_primary_composite_heterogeneity": "Composite endpoint not directly comparable",
        "same_cohort_metric_comparison_sensitivity": "Comparator metric from the same cohort family",
        "sensitivity_specialized_acs_subgroup": "Specialized ACS cohort",
        "continuous_raw_unit_main": "Single adjusted cohort on the native scale",
        "special_population_sensitivity": "Restricted special-population subgroup",
        "sensitivity_specialized_cad": "Specialized CAD cohort",
        "sensitivity_specialized_adhf": "Specialized acute-heart-failure cohort",
        "sensitivity_specialized_elderly_cvd": "Restricted elderly cardiovascular subgroup",
        "sensitivity_general_population_osa_surrogate": "Community physiology cohort without clinic-defined OSA entry criteria",
        "sensitivity_specialized_perioperative": "Specialized perioperative cohort",
        "sensitivity_specialized_postcabg_af": "Specialized post-CABG cohort",
        "categorical_high_vs_low_main": "Not enough independent cohorts on a shared threshold",
        "sensitivity_unadjusted_threshold": "Unadjusted threshold analysis",
        "sensitivity_specialized_hfref_mixedsdb": "Heart-failure cohort with mixed sleep-disordered breathing",
        "sensitivity_specialized_postablation_af": "Post-ablation AF cohort",
        "sensitivity_specialized_hf_nocthypox": "Specialized decompensated-HF cohort",
    }
    return mapping.get(value, value.replace("_", " "))


def first_sentence(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", (text or "").strip()).replace("…", "")
    if not cleaned:
        return ""
    match = re.search(r"(?<=[.!?])\s", cleaned)
    if match:
        return cleaned[: match.start()].strip()
    return cleaned


def format_summary_effect_label(label: str) -> str:
    cleaned = (label or "").strip()
    prefix = "HR"
    for candidate in ("OR", "RR", "HR"):
        if cleaned.startswith(f"{candidate} "):
            prefix = candidate
            break
    nums = re.findall(r"\d+\.\d+", label or "")
    if len(nums) >= 3:
        return f"{prefix} {fmt_num(nums[0])} ({fmt_num(nums[1])}-{fmt_num(nums[2])})"
    return label


def format_range_label(label: str) -> str:
    nums = re.findall(r"\d+\.\d+", label or "")
    if len(nums) >= 2:
        return f"{fmt_num(nums[0])}-{fmt_num(nums[1])}"
    return label


def evidence_status(row: dict[str, str]) -> tuple[str, str, str]:
    metric = row["metric_family"]
    outcome = row["outcome_family"]
    if metric == "T90" and outcome == "incident atrial fibrillation":
        return (
            "Two AF-oriented T90 cohorts show a positive direction, and the updated SHHS HB AF paper is directionally supportive, but one T90 estimate still requires exploratory harmonization.",
            "Direct native-scale replication remains limited.",
            "T90 AF evidence is directionally positive but still not cleanly replicated on a shared native scale.",
        )
    if metric == "T90" and outcome == "cardiovascular mortality":
        return (
            "An adjusted community MrOS row and the new SantOSA per-1% row now support the same outcome family.",
            "Scale and phenotype mismatch still prevent direct pooling.",
            "T90 cardiovascular-mortality evidence is thicker than before but still not directly replicated on a shared scale.",
        )
    if metric == "T90 / TST90" and outcome == "hard composite cardiovascular outcomes":
        return (
            "Log-scale T90 and spline-based TST90 rows now both support hard composite cardiovascular outcomes.",
            "Effect scales and composite definitions remain heterogeneous.",
            "T90/TST90 hard-composite evidence has broadened meaningfully, but it remains non-poolable.",
        )
    if metric == "ODI" and outcome == "all-cause mortality":
        return (
            "Clinical-OSA and specialized rows are available.",
            "Clean OSA-framed replication is limited.",
            "ODI mortality evidence spans clinical and specialized cohorts without a shared scale.",
        )
    if metric == "ODI" and outcome == "cardiovascular mortality":
        return (
            "Evidence is still anchored by one adjusted community row plus comparator-level SHHS data.",
            "Independent adjusted ODI cohorts remain limited.",
            "ODI cardiovascular-mortality evidence is directionally informative but still not a clean replicated cohort set.",
        )
    if metric == "HB / SASHB" and outcome == "hard cardiovascular outcomes":
        return (
            "HB remains strongest for cardiovascular mortality, and newer HB/SASHB hard-outcome signals now extend to AF and perioperative/postoperative settings.",
            "Most newer non-mortality signals come from overlap-prone or specialized settings.",
            "HB/SASHB hard-outcome evidence is thicker than before but still short of independent publication-level replication beyond the core mortality/HF anchors.",
        )
    if metric == "HB / ODI / T90" and outcome == "stroke or cerebrovascular outcomes":
        return (
            "Stroke-oriented evidence now includes older CAD data, a community SHHS male T90 row, and mixed comparator signals.",
            "Clean OSA-focused cohorts and directly comparable stroke definitions are scarce.",
            "Cerebrovascular hypoxemia evidence is broader than before but still too heterogeneous for synthesis-ready inference.",
        )
    return (row["current_state"], "Barrier not further classified", short_text(row["current_state"], 88))


def overall_rob_label(text: str) -> str:
    label = (text or "").capitalize()
    if label == "Moderate":
        return "Moderate"
    if label == "High":
        return "High"
    if label == "Low":
        return "Low"
    return label or "NR"


def rob_short(text: str) -> str:
    mapping = {"Low": "L", "Moderate": "M", "High": "H"}
    return mapping.get(overall_rob_label(text), overall_rob_label(text))


def build_table1() -> list[dict[str, str]]:
    rows = read_tsv(RESULTS / "study_characteristics_table_v5.tsv")
    out = []
    for row in rows:
        cohorts, n_by_cohort, follow_up = article_cohort_summary(row["article_pmid"])
        out.append(
            {
                "Study": f"{study_label(row['citation_key'], row['article_pmid'])}{ref_suffix(row['article_pmid'])}",
                "Cohort(s)": cohorts,
                "Cls": cohort_class_short(row["population_summary"]),
                "Metric": simplify_metric_label(row["metric_families"]),
                "Outcome": simplify_outcome_label(row["outcome_families"]),
                "Analytic N": n_by_cohort,
                "Follow-up": follow_up,
                "Layer / ROB": f"{layer_label_short(row['review_role'], row['main_pool_rows_n'], row['sensitivity_rows_n'])} / {overall_rob_label(row['overall_judgment'])}",
            }
        )
    return out


def build_main_table1() -> list[dict[str, str]]:
    rows = read_tsv(RESULTS / "study_characteristics_table_v5.tsv")
    out = []
    for row in rows:
        _, n_by_cohort, _ = article_cohort_summary(row["article_pmid"])
        out.append(
            {
                "Study": f"{study_label(row['citation_key'], row['article_pmid'])}{ref_suffix(row['article_pmid'])}",
                "Cls": cohort_class_short(row["population_summary"]),
                "Metric": simplify_metric_label(row["metric_families"]),
                "Main outcome": simplify_outcome_label(row["outcome_families"]),
                "Analytic N": n_by_cohort,
                "Layer": layer_label_short(row["review_role"], row["main_pool_rows_n"], row["sensitivity_rows_n"]),
                "ROB": overall_rob_label(row["overall_judgment"]),
            }
        )
    return out


def build_table2() -> list[dict[str, str]]:
    rows = read_tsv(RESULTS / "main_results_table_v5.tsv")
    out = []
    for row in rows:
        out.append(
            {
                "Metric": row["exposure_metric"],
                "Outcome": row["outcome_primary"],
                "Scale": scale_label(row["scale_family"]),
                "Publication / cohorts": primary_context_label(row),
                "Pooled HR (95% CI)": fmt_effect_label(row["pooled_effect_random"], row["pooled_ci_random_lower"], row["pooled_ci_random_upper"]),
                "I2 (%)": fmt_i2(row["i2_percent"]),
            }
        )
    return out


def build_table3() -> list[dict[str, str]]:
    rows = read_tsv(RESULTS / "sensitivity_results_table_v5.tsv")
    out = []
    for row in rows:
        if row["sensitivity_type"] == "leave_one_out_range":
            continue
        out.append(
            {
                "Sensitivity analysis": sensitivity_label(row["sensitivity_type"], row["analysis_cell_id"]),
                "Cell tested": (
                    "SASHB -> incident HF"
                    if "incident_heart_failure" in row["analysis_cell_id"]
                    else "T90 -> incident AF"
                ),
                "Pooled HR (95% CI)": format_summary_effect_label(row["summary_effect_label"]),
                "I2 (%)": fmt_i2(row["i2_percent"]),
                "Interpretation": clean_sensitivity_note(row),
            }
        )
    return out


def build_leave_one_out_table() -> list[dict[str, str]]:
    rows = [r for r in read_tsv(RESULTS / "sensitivity_results_table_v5.tsv") if r["sensitivity_type"] == "leave_one_out_range"]
    out = []
    for row in rows:
        out.append(
            {
                "Primary pooled cell": sensitivity_label(row["sensitivity_type"], row["analysis_cell_id"]).replace("Leave-one-out range: ", ""),
                "Range after omitting one cohort": format_range_label(row["summary_effect_label"]),
                "Main reading": clean_sensitivity_note(row),
            }
        )
    return out


def build_table4() -> list[dict[str, str]]:
    rows = [r for r in read_tsv(RESULTS / "evidence_gap_table_v5.tsv") if r.get("gap_id") != "embase_layer"]
    out = []
    for row in rows:
        status, barrier, interpretation = evidence_status(row)
        out.append(
            {
                "Metric-outcome family": f"{row['metric_family']} -> {row['outcome_family']}",
                "Barrier": barrier,
                "Evidence summary": interpretation.replace(" remains ", " is ").replace(" currently ", " ").replace(" still ", " "),
            }
        )
    return out


def build_nonpooled_table() -> list[dict[str, str]]:
    rows = read_tsv(RESULTS / "nonpooled_evidence_table_v5.tsv")
    out = []
    for row in rows:
        out.append(
            {
                "Study": f"{study_label(row['study_id'], row['article_pmid'])}{ref_suffix(row['article_pmid'])}",
                "Metric family": row["exposure_metric"],
                "Outcome": row["outcome_primary"],
                "Effect estimate": format_summary_effect_label(row["summary_effect_label"]),
                "Reason retained outside primary pooled analyses": pooling_reason_label(row["pooling_eligibility"]),
                "Main note": clean_nonpooled_note(row["notes"]),
            }
        )
    return out


def build_domain_rob_table() -> list[dict[str, str]]:
    rows = read_tsv(ANALYSIS / "risk_of_bias_quips_working_v5.tsv")
    out = []
    for row in rows:
        out.append(
            {
                "Study": f"{study_label(row['citation_key'], row['article_pmid'])}{ref_suffix(row['article_pmid'])}",
                "Anchor": f"{simplify_metric_label(row['main_metric_family'])} -> {simplify_outcome_label(row['main_outcomes'])}",
                "Part.": rob_short(row["study_participation"]),
                "Attr.": rob_short(row["study_attrition"]),
                "Factor": rob_short(row["prognostic_factor_measurement"]),
                "Outcome": rob_short(row["outcome_measurement"]),
                "Confound.": rob_short(row["confounding"]),
                "Analysis": rob_short(row["statistical_analysis_reporting"]),
                "Overall": rob_short(row["overall_judgment"]),
            }
        )
    return out


def build_extraction_worksheet() -> list[dict[str, str]]:
    rows = read_tsv(ANALYSIS / "extraction_master_v5.tsv")
    out = []
    for row in rows:
        population = row["population_setting"] or row["cohort_name"]
        if row["sex_subgroup"]:
            population = f"{population}; {row['sex_subgroup']} subgroup"
        out.append(
            {
                "Study": f"{study_label(row['study_id'], row['article_pmid'])}{ref_suffix(row['article_pmid'])}",
                "Cohort": row["cohort_name"],
                "Cohort class": cohort_class_label(population),
                "Country": row["country"] or "NR",
                "Metric family": row["exposure_metric"],
                "Outcome family": row["outcome_primary"],
                "Exposure scale": row["exposure_scale"],
                "Analytic N": row["analytic_n"] or "NR",
                "Follow-up (years)": normalize_followup(row["follow_up_years"]),
                "Effect measure": row["effect_measure"],
                "Effect estimate": fmt_num(row["effect_estimate"], 3),
                "Lower CI": fmt_num(row["ci_lower"], 3),
                "Upper CI": fmt_num(row["ci_upper"], 3),
                "AHI adjusted": row["ahi_adjusted"] or "NR",
                "Analysis layer": "Main pool" if row["main_pool_flag"] == "yes" else "Sensitivity/comparator",
                "Pooling status": row["pooling_eligibility"].replace("_", " "),
                "Adjustment summary": short_text(row["adjustment_summary"], 120),
                "Source location": short_text(row["source_location"], 72),
            }
        )
    return out


def clean_nonpooled_note(text: str) -> str:
    cleaned = first_sentence(text)
    replacements = {
        "Main-text excerpt": "The article",
        "Published summary": "The article",
        "the extracted snippet": "the article",
        "extracted": "reported",
        "snippet": "summary",
        "accessible full text": "the full text reviewed for this submission",
        "Abstract-level extraction only.": "Abstract-only specialized comparator row without sufficient full-text model detail for pooling.",
        "recoverable from the full text": "available from the full text reviewed for this submission",
        "recoverable from the PDF text layer": "fully recoverable from the published PDF text layer",
        "The article did not report the exact number of CVD deaths in this 2023 article.": "The adjusted CVD-mortality estimate was available, but the article did not report a corresponding event count.",
        "Full text confirms an unadjusted threshold analysis with 8% mortality for T90 <=20% versus 18% for T90 >20%.": "Unadjusted threshold analysis only; no multivariable threshold model was available for pooling.",
        "This is the paper's highlighted incremental-risk result.": "Incremental-risk result from an alternate model within the same cohort family.",
        "This is currently a single clinical-cohort ODI all-cause mortality row with exact multivariable effect size.": "Single clinical-cohort ODI mortality estimate with a directly extractable multivariable effect size.",
        "This is a robust T90 mortality estimate.": "Adjusted T90 mortality estimate from a specialized heart-failure cohort.",
        "Specialized secondary-prevention cohort.": "Specialized secondary-prevention cohort retained outside the primary pooled analyses.",
        "Baseline characteristics were reported as medians: age 76 [72-80], BMI 27.0 [25.0-29.0], male 100%.": "Single community-cohort HB row without a second directly comparable cohort family.",
        "Baseline characteristics were reported as medians: age 67 [61-75], BMI 27.9 [24.7-31.8], male 46.4%.": "Single community-cohort HB row without a second directly comparable cohort family.",
    }
    for old, new in replacements.items():
        cleaned = cleaned.replace(old, new)
    if "highlighted incremental-risk result" in cleaned:
        cleaned = "Incremental-risk result from an alternate model within the same cohort family."
    if "single clinical-cohort ODI all-cause mortality row" in cleaned:
        cleaned = "Single clinical-cohort ODI mortality estimate with a directly extractable multivariable effect size."
    if "robust T90 mortality estimate" in cleaned:
        cleaned = "Adjusted T90 mortality estimate from a specialized heart-failure cohort."
    if "Abstract-only report retained as specialized comparator evidence" in cleaned:
        cleaned = "Abstract-only specialized comparator row without sufficient full-text model detail for pooling."
    if cleaned.startswith("Specialized secondary-prevention cohort"):
        cleaned = "Specialized secondary-prevention cohort retained outside the primary pooled analyses."
    if cleaned.startswith("Full text confirms an unadjusted threshold analysis"):
        cleaned = "Unadjusted threshold analysis only; no multivariable threshold model was available for pooling."
    if cleaned.startswith("The article did not report the exact number of CVD deaths"):
        cleaned = "The adjusted CVD-mortality estimate was available, but the article did not report a corresponding event count."
    cleaned = cleaned.replace("sufficient full-text model detail", "enough published model detail")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if ";" in cleaned:
        cleaned = cleaned.split(";", 1)[0].rstrip() + "."
    if ", but" in cleaned and len(cleaned) > 150:
        cleaned = cleaned.split(", but", 1)[0].rstrip() + "."
    return cleaned


def nonpooled_group(metric: str) -> str:
    if metric in {"HB", "SASHB", "rHB", "SBII", "nocturnal hypoxemia"}:
        return "HB-family and comparator constructs"
    if metric == "ODI":
        return "ODI family"
    if metric == "T90":
        return "T90 family"
    return "Nadir oxygen saturation and other comparator constructs"


def build_nonpooled_sections() -> list[tuple[str, str, list[dict[str, str]]]]:
    rows = build_nonpooled_table()
    grouped: dict[str, list[dict[str, str]]] = {}
    for row in rows:
        grouped.setdefault(nonpooled_group(row["Metric family"]), []).append(
                {
                    "Study": row["Study"],
                    "Outcome": simplify_outcome_label(row["Outcome"]),
                    "Effect (95% CI)": row["Effect estimate"],
                    "Main pooling barrier": row["Reason retained outside primary pooled analyses"],
                    "Main message": clean_nonpooled_note(row["Main note"]),
                }
        )
    order = [
        "HB-family and comparator constructs",
        "ODI family",
        "T90 family",
        "Nadir oxygen saturation and other comparator constructs",
    ]
    sections = []
    for title in order:
        if title in grouped:
            sections.append((title, "Rows retained outside the four-cell primary pooled analyses.", grouped[title]))
    return sections


def build_domain_rob_sections() -> list[tuple[str, str, list[dict[str, str]]]]:
    rows = build_domain_rob_table()
    return [
        (
            "ROB matrix A",
            "Core QUIPS-style domains shown as compact judgments. L = low, M = moderate, H = high.",
            [
                {
                    "Study": row["Study"],
                    "Anchor": row["Anchor"],
                    "Part.": row["Part."],
                    "Attr.": row["Attr."],
                    "Factor": row["Factor"],
                    "Outcome": row["Outcome"],
                }
                for row in rows
            ],
        ),
        (
            "ROB matrix B",
            "Confounding, analysis/reporting, and overall judgments for the same included articles.",
            [
                {
                    "Study": row["Study"],
                    "Confound.": row["Confound."],
                    "Analysis": row["Analysis"],
                    "Overall": row["Overall"],
                }
                for row in rows
            ],
        ),
    ]


def write_multitable_markdown(
    path: Path,
    title: str,
    intro: str,
    sections: list[tuple[str, str, list[dict[str, str]]]],
    closing_note: str | None = None,
) -> None:
    blocks = [f"# {title}", "", intro, ""]
    for section_title, section_note, rows in sections:
        blocks.append(f"## {section_title}")
        blocks.append("")
        blocks.append(markdown_table_only(rows))
        blocks.append("")
        blocks.append(section_note)
        blocks.append("")
    if closing_note:
        blocks.append(closing_note)
        blocks.append("")
    path.write_text("\n".join(blocks), encoding="utf-8")


def parse_prisma_counts() -> dict[str, int]:
    text = (ANALYSIS / "prisma_flow_current_v12.md").read_text(encoding="utf-8")
    patterns = {
        "pubmed_ids": r"PubMed main query identified: `(\d+)`",
        "pubmed_screened": r"PubMed main records title/abstract screened: `(\d+)`",
        "wos_exported": r"Web of Science Core Collection records exported.*: `(\d+)`",
        "wos_unique": r"Web of Science records unique after deduplication.*: `(\d+)`",
        "embase_exported": r"Embase records exported.*: `(\d+)`",
        "embase_unique": r"Embase unique after deduplication.*: `(\d+)`|Embase records unique after deduplication.*: `(\d+)`",
    }
    out: dict[str, int] = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            groups = [g for g in match.groups() if g]
            out[key] = int(groups[0])
    side_rows = read_tsv(ANALYSIS / "side_search_triage.tsv")
    counts = Counter(row["triage_decision"] for row in side_rows)
    fulltext_rows = read_tsv(ANALYSIS / "fulltext_review_log.tsv")
    extraction_rows = read_tsv(ANALYSIS / "extraction_master_v5.tsv")
    historical_rows = [row for row in extraction_rows if row.get("source_round") != "round26_upgrade"]
    upgrade_rows = [row for row in extraction_rows if row.get("source_round") == "round26_upgrade"]
    historical_pmids = {row["article_pmid"] for row in historical_rows}
    upgrade_pmids = {row["article_pmid"] for row in upgrade_rows}
    historical_nonretained_fulltext_rows = [
        row
        for row in fulltext_rows
        if row["pubmed_id"] not in historical_pmids and row["pubmed_id"] not in upgrade_pmids
    ]
    out["side_unique"] = len(side_rows)
    out["side_high"] = counts.get("high_priority", 0)
    out["side_medium"] = counts.get("medium_priority", 0)
    out["side_low"] = counts.get("low_priority", 0)
    out["side_likely_exclude"] = counts.get("likely_exclude", 0)
    out["screened_total"] = 510 + 187 + 162
    out["pubmed_not_parsed"] = out["pubmed_ids"] - out["pubmed_screened"]
    out["fulltext_reviewed"] = len(fulltext_rows)
    out["included_articles"] = len({row["article_pmid"] for row in historical_rows})
    out["historical_rows"] = len(historical_rows)
    out["historical_primary_rows"] = sum(row["main_pool_flag"] == "yes" for row in historical_rows)
    out["historical_sensitivity_rows"] = sum(row["sensitivity_flag"] == "yes" for row in historical_rows)
    out["updated_included_articles"] = len({row["article_pmid"] for row in extraction_rows})
    out["updated_rows"] = len(extraction_rows)
    out["updated_primary_rows"] = sum(row["main_pool_flag"] == "yes" for row in extraction_rows)
    out["updated_sensitivity_rows"] = sum(row["sensitivity_flag"] == "yes" for row in extraction_rows)
    out["updated_nonpooled_rows"] = out["updated_rows"] - 8
    out["upgrade_retained_articles"] = len({row["article_pmid"] for row in upgrade_rows})
    out["upgrade_retained_rows"] = len(upgrade_rows)
    out["later_retained_postfreeze"] = len(upgrade_pmids)
    out["title_abstract_excluded"] = out["screened_total"] - out["fulltext_reviewed"]
    out["fulltext_not_retained"] = out["fulltext_reviewed"] - out["included_articles"]
    decision_counts = Counter(row["provisional_fulltext_decision"] for row in historical_nonretained_fulltext_rows)
    out["narrative_or_nonextractable"] = decision_counts.get("include_narrative_only_nonextractable", 0) + decision_counts.get("narrative_only", 0)
    out["scope_or_protocol_excluded"] = decision_counts.get("exclude_candidate", 0) + decision_counts.get("separate_intervention", 0)
    out["other_nonretained_context"] = len(historical_nonretained_fulltext_rows) - out["narrative_or_nonextractable"] - out["scope_or_protocol_excluded"]
    return out


def fulltext_exclusion_rows() -> list[dict[str, str]]:
    historical_pmids = {
        row["article_pmid"]
        for row in read_tsv(ANALYSIS / "extraction_master_v5.tsv")
        if row.get("source_round") != "round26_upgrade"
    }
    upgrade_pmids = {
        row["article_pmid"]
        for row in read_tsv(ANALYSIS / "extraction_master_v5.tsv")
        if row.get("source_round") == "round26_upgrade"
    }
    included_rows = [row for row in read_tsv(RESULTS / "study_characteristics_table_v5.tsv") if row["article_pmid"] in historical_pmids]
    included_pmids = {row["article_pmid"] for row in included_rows}
    included_author_year = {
        f"{study_label(row['citation_key'], row['article_pmid']).rsplit(' ', 1)[0]} {row['citation_key'].split('_')[1]}"
        for row in included_rows
        if "_" in row["citation_key"]
    }
    rows = read_tsv(ANALYSIS / "fulltext_review_log.tsv")
    retained_rows = [row for row in rows if row["pubmed_id"] not in included_pmids]
    base_labels = [f"{row['first_author']} {row['pub_year']}" for row in retained_rows]
    label_counts = Counter(base_labels)

    def context_suffix(row: dict[str, str]) -> str:
        design = (row.get("design_prefill") or "").lower()
        status = (row.get("fulltext_status") or "").lower()
        cohort = (row.get("cohort_name_prefill") or "").strip()
        role = (row.get("provisional_role") or "").lower()
        pmid = row.get("pubmed_id", "")
        curated = {
            "21220756": "maintenance hemodialysis pulse-oximetry cohort",
            "22705247": "older community-dwelling men cohort",
            "27464791": "recent myocardial infarction cohort",
            "27690206": "suspected OSA obesity-hypoxaemia note",
            "34226030": "obstructive HCM septal-myectomy cohort",
            "40794640": "RICCADSA / ISAACC / SAVE pooled analysis",
        }
        if pmid in curated:
            return curated[pmid]
        if "letter" in design or "letter" in status:
            return "letter note"
        if cohort:
            cleaned = cohort.replace("OSAS", "OSA").replace("study", "cohort")
            cleaned = cleaned.replace("large ", "").replace("single-centre ", "").replace("single-center ", "")
            cleaned = cleaned.replace(" cohort", "")
            cleaned = cleaned.replace("clinical OSA", "clinical-OSA").replace("overlap mortality note", "overlap note")
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            return cleaned
        if "intervention" in role:
            return "intervention analysis"
        if "comparator" in role:
            return "comparator row"
        return f"PMID {row['pubmed_id']}"

    provisional_labels: list[dict[str, str]] = []
    for row in retained_rows:
        decision = row["provisional_fulltext_decision"]
        role = row["provisional_role"]
        base_label = f"{row['first_author']} {row['pub_year']}"
        context = context_suffix(row)
        pmid = row["pubmed_id"]
        if label_counts[base_label] > 1 or base_label in included_author_year:
            report_label = f"{base_label} ({context})"
        else:
            report_label = base_label
        if pmid in upgrade_pmids:
            disposition = "Later retained in post-freeze supplement"
            reason = "Historical executed package remained frozen; the article was later re-adjudicated and retained in the updated submission dataset."
        elif "composite_highrisk_phenotype_comparator" in role:
            disposition = "Specialized/context/noncanonical comparator report"
            reason = "Composite high-risk OSA phenotype combined HB or ΔHR rather than reporting a separable prespecified metric-family estimate."
        elif "specialized_secondary_prevention" in role:
            disposition = "Specialized/context/noncanonical comparator report"
            reason = "Specialized ACS secondary-prevention cohort with inverse or U-shaped TSA90 associations; not retained as a general OSA prognostic anchor."
        elif "wrong_population_frame" in role:
            disposition = "Protocol exclusion"
            reason = "Community-based nocturnal saturation construct lay outside the prespecified OSA-related metric framework."
        elif decision in {"include_sensitivity"}:
            disposition = "Specialized/context/noncanonical comparator report"
            reason = "Did not yield a prespecified retained row in the historical executed quantitative evidence set."
        elif decision in {"include_narrative_only_nonextractable", "narrative_only"}:
            disposition = "Narrative-only or nonextractable report"
            reason = "No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set."
        elif decision in {"separate_intervention"}:
            disposition = "Protocol exclusion"
            reason = "Intervention or effect-modifier analysis rather than prognostic cohort evidence."
        elif "scope" in role or decision == "exclude_candidate":
            disposition = "Protocol exclusion"
            reason = "Outside the prespecified OSA-related prognostic scope."
        else:
            disposition = "Specialized/context/noncanonical comparator report"
            reason = "Did not yield a prespecified retained row in the historical executed quantitative evidence set."
        provisional_labels.append(
            {
                "Report": report_label,
                "PMID": row["pubmed_id"],
                "Context": context,
                "Final disposition": disposition,
                "Main reason for non-retention": reason,
            }
        )
    report_counts = Counter(item["Report"] for item in provisional_labels)
    out = []
    for item in provisional_labels:
        report = item["Report"]
        if report_counts[report] > 1:
            if report.endswith(")"):
                report = report[:-1] + f"; PMID {item['PMID']})"
            else:
                report = f"{report} (PMID {item['PMID']})"
        out.append(
            {
                "Report": report,
                "PMID": item["PMID"],
                "Context": item["Context"],
                "Final disposition": item["Final disposition"],
                "Main reason for non-retention": item["Main reason for non-retention"],
            }
        )
    return out


def wrap_text_pixels(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        trial = f"{current} {word}"
        width = draw.textbbox((0, 0), trial, font=font)[2]
        if width <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def draw_wrapped_text(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    max_width: int,
    fill: str,
    line_gap: int = 6,
    align: str = "left",
) -> int:
    x, y = xy
    bbox = draw.textbbox((0, 0), "Ag", font=font)
    line_height = bbox[3] - bbox[1]
    for line in wrap_text_pixels(draw, text, font, max_width):
        if align == "center":
            line_box = draw.textbbox((0, 0), line, font=font)
            line_w = line_box[2] - line_box[0]
            draw.text((x + (max_width - line_w) / 2, y), line, font=font, fill=fill)
        else:
            draw.text((x, y), line, font=font, fill=fill)
        y += line_height + line_gap
    return y


def wrapped_block_height(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
    max_width: int,
    line_gap: int = 6,
) -> int:
    lines = wrap_text_pixels(draw, text, font, max_width)
    bbox = draw.textbbox((0, 0), "Ag", font=font)
    line_height = bbox[3] - bbox[1]
    return len(lines) * line_height + max(0, len(lines) - 1) * line_gap


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    fill: str,
    width: int = 5,
    scale: float = 1.0,
) -> None:
    x1, y1 = start
    x2, y2 = end
    draw.line((x1, y1, x2, y2), fill=fill, width=width)
    head_w = px(10, scale)
    head_l = px(18, scale)
    if x1 == x2:
        direction = 1 if y2 > y1 else -1
        head = [(x2 - head_w, y2 - head_l * direction), (x2 + head_w, y2 - head_l * direction), (x2, y2)]
    else:
        direction = 1 if x2 > x1 else -1
        head = [(x2 - head_l * direction, y2 - head_w), (x2 - head_l * direction, y2 + head_w), (x2, y2)]
    draw.polygon(head, fill=fill)


def draw_box_with_text(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int, int, int],
    title: str,
    body: list[str],
    fill: str,
    outline: str,
    title_font: ImageFont.ImageFont,
    body_font: ImageFont.ImageFont,
    *,
    scale: float = 1.0,
) -> None:
    x1, y1, x2, y2 = xy
    draw.rectangle((x1, y1, x2, y2), fill=fill, outline=outline, width=px(3, scale))
    max_width = x2 - x1 - px(44, scale)
    title_h = wrapped_block_height(draw, title, title_font, max_width, line_gap=px(2, scale))
    body_h = 0
    for item in body:
        body_h += wrapped_block_height(draw, item, body_font, max_width, line_gap=px(2, scale)) + px(6, scale)
    total_h = title_h + px(10, scale) + body_h
    inner_h = y2 - y1 - px(36, scale)
    y = y1 + px(18, scale) + max(0, (inner_h - total_h) // 2)
    y = draw_wrapped_text(draw, (x1 + px(22, scale), y), title, title_font, max_width, "#10243c", line_gap=px(2, scale), align="center") + px(10, scale)
    for item in body:
        y = draw_wrapped_text(draw, (x1 + px(22, scale), y), item, body_font, max_width, "#1f2937", line_gap=px(2, scale), align="center") + px(6, scale)


def make_prisma_figure(counts: dict[str, int]) -> Image.Image:
    duplicates_removed = counts["pubmed_ids"] + counts["wos_exported"] + counts["embase_exported"] - counts["screened_total"]
    scale = FIG_NATIVE_SCALE
    img = Image.new("RGB", (px(2200, scale), px(1260, scale)), "white")
    draw = ImageDraw.Draw(img)
    box_title_font = load_font(px(28, scale), bold=True)
    body_font = load_font(px(22, scale), bold=False)

    draw_box_with_text(
        draw,
        scale_box((650, 40, 1550, 190), scale),
        "Records identified from databases",
        [
            f"Total: {counts['pubmed_ids'] + counts['wos_exported'] + counts['embase_exported']}",
            f"PubMed: {counts['pubmed_ids']}",
            f"Web of Science Core Collection: {counts['wos_exported']}",
            f"Embase: {counts['embase_exported']}",
        ],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((120, 230, 700, 350), scale),
        "Duplicate records removed before screening",
        [f"{duplicates_removed} records"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((810, 260, 1510, 380), scale),
        "Records screened",
        [f"{counts['screened_total']} records"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((120, 490, 700, 610), scale),
        "Records excluded",
        [f"{counts['title_abstract_excluded']} records"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((810, 490, 1510, 610), scale),
        "Reports sought for retrieval",
        [f"{counts['fulltext_reviewed']} reports"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((1580, 490, 2060, 610), scale),
        "Reports not retrieved",
        ["0 reports"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((810, 710, 1510, 830), scale),
        "Reports assessed for eligibility",
        [f"{counts['fulltext_reviewed']} reports"],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((1520, 660, 2110, 930), scale),
        "Reports excluded, with reasons",
        [
            f"{counts['fulltext_not_retained']} reports",
            f"Narrative-only/nonextractable: {counts['narrative_or_nonextractable']}",
            f"Scope/intervention exclusions: {counts['scope_or_protocol_excluded']}",
            f"Specialized/context/noncanonical: {counts['other_nonretained_context']}",
        ],
        "#ffffff",
        "#3f4d63",
        box_title_font,
        body_font,
        scale=scale,
    )
    draw_box_with_text(
        draw,
        scale_box((810, 1010, 1510, 1150), scale),
        "Studies included in quantitative evidence set",
        [f"{counts['included_articles']} unique articles", f"{counts['historical_rows']} cohort-level rows in the historical extraction master"],
        "#f6fffb",
        "#0f766e",
        box_title_font,
        body_font,
        scale=scale,
    )
    arrow_color = "#4b5563"
    draw.line(scale_box((1100, 190, 1100, 234), scale), fill=arrow_color, width=px(5, scale))
    draw_arrow(draw, scale_box((1100, 234, 700, 290), scale)[:2], scale_box((1100, 234, 700, 290), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1100, 234, 1160, 260), scale)[:2], scale_box((1100, 234, 1160, 260), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1160, 380, 1160, 490), scale)[:2], scale_box((1160, 380, 1160, 490), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw.line(scale_box((810, 320, 810, 550), scale), fill=arrow_color, width=px(5, scale))
    draw_arrow(draw, scale_box((810, 550, 700, 550), scale)[:2], scale_box((810, 550, 700, 550), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1510, 550, 1580, 550), scale)[:2], scale_box((1510, 550, 1580, 550), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1160, 610, 1160, 710), scale)[:2], scale_box((1160, 610, 1160, 710), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1510, 770, 1520, 770), scale)[:2], scale_box((1510, 770, 1520, 770), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    draw_arrow(draw, scale_box((1160, 830, 1160, 1010), scale)[:2], scale_box((1160, 830, 1160, 1010), scale)[2:], arrow_color, width=px(5, scale), scale=scale)
    return img


def write_prisma_svg(path: Path, counts: dict[str, int]) -> None:
    ensure_dir(path.parent)
    lines = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="980" height="620" viewBox="0 0 980 620">',
        '<style>text{font-family:Arial,sans-serif;fill:#1f2937}.title{font-size:22px;font-weight:700}.box-title{font-size:15px;font-weight:700}.body{font-size:13px}.note{font-size:12px}.box{fill:#f8fbff;stroke:#2a5f8a;stroke-width:2}.audit{fill:#fff8e8;stroke:#a7701b;stroke-width:2;stroke-dasharray:6 4}.arrow{stroke:#4b5563;stroke-width:2.2;fill:none;marker-end:url(#arrow)}</style>',
        '<defs><marker id="arrow" markerWidth="12" markerHeight="8" refX="10" refY="4" orient="auto"><polygon points="0 0, 12 4, 0 8" fill="#4b5563"/></marker></defs>',
        '<text class="title" x="40" y="34">Figure 1. PRISMA flow of the executed three-database package</text>',
    ]

    def add_box(x: int, y: int, w: int, h: int, title: str, body: list[str], cls: str = "box") -> None:
        lines.append(f'<rect class="{cls}" x="{x}" y="{y}" rx="10" ry="10" width="{w}" height="{h}"/>')
        lines.append(f'<text class="box-title" x="{x+14}" y="{y+24}">{escape(title)}</text>')
        cy = y + 48
        width = 34 if w < 270 else 40
        for item in body:
            for wrapped in textwrap.wrap(item, width=width):
                lines.append(f'<text class="body" x="{x+14}" y="{cy}">{escape(wrapped)}</text>')
                cy += 18

    add_box(
        320,
        48,
        320,
        92,
        "Database records identified",
        [
            f"PubMed: {counts['pubmed_ids']}",
            f"Web of Science Core Collection: {counts['wos_exported']}",
            f"Embase: {counts['embase_exported']}",
        ],
    )
    add_box(
        320,
        182,
        320,
        92,
        "Records screened after deduplication",
        [
            "PubMed: 510",
            "Web of Science unique: 187",
            "Embase unique: 162",
            f"Total screened: {counts['screened_total']}",
        ],
    )
    add_box(
        60,
        332,
        240,
        82,
        "Excluded at title/abstract",
        [f"{counts['title_abstract_excluded']} records"],
    )
    add_box(
        360,
        332,
        240,
        82,
        "Full-text reports assessed for eligibility",
        [f"{counts['fulltext_reviewed']} reports"],
    )
    add_box(
        660,
        332,
        260,
        118,
        "Reports excluded after eligibility assessment",
        [
            f"{counts['fulltext_not_retained']} records",
            f"Narrative-only/nonextractable: {counts['narrative_or_nonextractable']}",
            f"Scope / intervention exclusions: {counts['scope_or_protocol_excluded']}",
            f"Context-only / noncanonical rows: {counts['other_nonretained_context']}",
        ],
    )
    add_box(
        360,
        488,
        240,
        82,
        "Included in quantitative evidence set",
        [f"{counts['included_articles']} unique articles", f"{counts['historical_rows']} cohort-level rows in the historical extraction master"],
    )
    lines.extend(
        [
            '<path class="arrow" d="M480 140 L480 182"/>',
            '<path class="arrow" d="M480 274 L480 332"/>',
            '<path class="arrow" d="M420 274 L420 373 L300 373"/>',
            '<path class="arrow" d="M540 274 L790 332"/>',
            '<path class="arrow" d="M480 414 L480 488"/>',
            '</svg>',
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def cohort_short_label(cohort: str) -> str:
    mapping = {
        "MrOS": "MrOS",
        "SHHS": "SHHS",
        "MESA": "MESA",
        "SHHS_men": "SHHS men",
        "MrOS_men": "MrOS men",
        "Cleveland_sleep_cohort": "Cleveland",
        "Blanchard_multicenterOSA": "Multicenter OSA",
    }
    return mapping.get(cohort, cohort.replace("_", " "))


def cohort_name_short(cohort_name: str) -> str:
    mapping = {
        "AF catheter ablation cohort": "AF ablation",
        "CIRCS cohort": "CIRCS",
        "Cleveland Clinic sleep-study cohort": "Cleveland",
        "CoroKind elderly community subgroup": "CoroKind",
        "French major noncardiothoracic surgery cohort": "French surgery",
        "Hong Kong sleep clinic cohort": "Hong Kong clinic",
        "MESA Sleep Study": "MESA",
        "MrOS Sleep Study": "MrOS",
        "OSA-ACS RCIR subgroup": "OSA-ACS",
        "POSA surgical cohort": "POSA surgical",
        "Pays de la Loire cohort": "Pays de la Loire",
        "RICCADSA observational cohort": "RICCADSA",
        "SantOSA": "SantOSA",
        "SHHS": "SHHS",
        "SHHS OSA without baseline AF cohort": "SHHS OSA no AF",
        "SHHS diabetes subgroup": "SHHS DM",
        "SHHS treated-diabetes subgroup": "SHHS treated DM",
        "coronary angiography CAD cohort": "CAD cohort",
        "coronary artery bypass surgery cohort": "CABG cohort",
        "decompensated heart failure pulse-oximetry cohort": "ADHF pulse-ox",
        "patients investigated for OSA in a multicenter referral network": "Multicenter referral",
        "predischarge ADHF LV systolic dysfunction cohort": "Predischarge ADHF",
        "recent myocardial infarction cohort": "Recent MI",
        "single-centre moderate-to-severe OSA cohort": "Single-centre OSA",
        "single-centre referral sleep-clinic cohort": "Referral clinic",
        "stable HF-REF cohort": "HF-REF",
        "suspected OSA diagnostic-sleep-study cohort": "Suspected OSA",
    }
    return mapping.get(cohort_name, cohort_name)


def simplify_metric_label(text: str) -> str:
    return compact_metric_family(text).replace(" / ", "; ")


def simplify_outcome_label(text: str) -> str:
    mapping = {
        "post-CABG atrial fibrillation": "post-CABG AF",
        "cerebrovascular events": "cerebrovascular events",
        "all-cause mortality": "all-cause mortality",
        "composite readmission or mortality": "readmission or mortality",
        "incident hospitalized atrial fibrillation": "incident AF hospitalization",
        "incident ischemic stroke": "incident stroke",
        "CVD mortality": "CVD mortality",
        "hard CVD; CVD/all-cause mortality": "hard CVD; CVD/all-cause mortality",
        "incident heart failure": "incident HF",
        "incident atrial fibrillation": "incident AF",
        "MACE": "MACE",
        "tachyarrhythmia recurrence": "AF recurrence",
        "incident CVD": "incident CVD",
        "MACCE": "MACCE",
        "mortality or HF readmission": "mortality or HF readmission",
        "30-day postoperative cardiovascular events": "30-day postoperative CV events",
    }
    return mapping.get(compact_outcome_family(text), compact_outcome_family(text))


def plain_n_token(text: str) -> str:
    nums = re.findall(r"\d+", (text or "").replace(",", ""))
    if not nums:
        return "NR"
    ints = sorted(int(n) for n in nums)
    if len(set(ints)) == 1:
        return str(ints[0])
    return "/".join(str(n) for n in sorted(set(ints)))


def followup_year_token(text: str) -> str:
    value = re.sub(r"\s+", " ", (text or "").strip())
    if not value:
        return "NR"
    if value == "in-hospital postoperative follow-up":
        return "in-hospital"
    if value == "5+":
        return ">=5 y"
    nums = re.findall(r"\d+(?:\.\d+)?", value)
    if not nums:
        return value
    numeric = [float(x) for x in nums]
    if len(set(numeric)) == 1:
        base = nums[0]
    else:
        ordered = []
        for num in sorted(set(numeric)):
            label = f"{num:.1f}".rstrip("0").rstrip(".")
            ordered.append(label)
        base = "; ".join(ordered)
    return f"{base} y"


def layer_label_short(review_role: str, main_rows: str, sens_rows: str) -> str:
    label = analysis_layer_label(review_role, main_rows, sens_rows)
    mapping = {
        "Primary pooled": "Primary pooled",
        "Primary + overlap-sensitive": "Primary + overlap",
        "Overlap-sensitive": "Overlap-sensitive",
        "Single-study anchor": "Single-study",
        "Primary evidence": "Primary evidence",
        "Sensitivity/comparator": "Comparator",
        "Narrative / non-pooled only": "Narrative only",
    }
    return mapping.get(label, label)


def article_cohort_summary(article_pmid: str) -> tuple[str, str, str]:
    special = {
        "32298733": (
            "SHHS men; MrOS men",
            "SHHS men: 2259 subgroup (4535 interaction model); MrOS men: 2653 (2646 adjusted model)",
            "SHHS men: 10.4; MrOS men: 8.8",
        ),
        "37418748": (
            "MESA; MrOS",
            "MESA hard CVD: 1891; MESA all-cause mortality: 1973; MrOS incident CVD: 1518; MrOS mortality: 2627",
            "MESA: 6.9 median; MrOS incident CVD: 9.4 median; MrOS mortality: 12.0 median",
        ),
        "38773880": (
            "SHHS",
            "CVD mortality: 4485; incident CVD: 3872",
            "NR",
        ),
        "40008168": (
            "SHHS free-of-baseline-CVD subset",
            "CVD mortality: 3714; MACE: 3698",
            "CVD mortality: 11.7 median; MACE: 11.4 median",
        ),
    }
    if article_pmid in special:
        return special[article_pmid]
    rows = [row for row in read_tsv(ANALYSIS / "extraction_master_v5.tsv") if row["article_pmid"] == article_pmid]
    if not rows:
        return ("NR", "NR", "NR")
    grouped: dict[str, dict[str, set[str]]] = {}
    for row in rows:
        cohort = cohort_name_short(row["cohort_name"])
        if row["sex_subgroup"] == "men" and "men" not in cohort.lower():
            cohort = f"{cohort} men"
        slot = grouped.setdefault(cohort, {"n": set(), "fu": set()})
        slot["n"].add(plain_n_token(row["analytic_n"]))
        slot["fu"].add(followup_year_token(row["follow_up_years"]))
    cohort_label = "; ".join(grouped)
    n_chunks = []
    fu_chunks = []
    for cohort, detail in grouped.items():
        n_vals = sorted(detail["n"])
        fu_vals = sorted(detail["fu"])
        n_value = n_vals[0] if len(n_vals) == 1 else " / ".join(n_vals)
        fu_value = fu_vals[0] if len(fu_vals) == 1 else " / ".join(fu_vals)
        n_value = re.sub(r"\s*/\s*", " / ", n_value)
        fu_value = re.sub(r"\s*/\s*", " / ", fu_value)
        label_single_group = len(grouped) == 1 and (len(n_vals) > 1 or len(fu_vals) > 1)
        needs_label = len(grouped) > 1 or label_single_group
        n_chunks.append(f"{cohort}: {n_value}" if needs_label else n_value)
        fu_chunks.append(f"{cohort}: {fu_value}" if needs_label else fu_value)
    joiner = "; "
    return (cohort_label, joiner.join(n_chunks), joiner.join(fu_chunks))


def study_forest_label(study_id: str, cohort: str) -> str:
    base = study_label(study_id)
    cohort_label = cohort_short_label(cohort)
    if len(base) >= 22 or len(base) + len(cohort_label) >= 34:
        return base
    if cohort_label and cohort_label.lower() not in base.lower():
        return f"{base} ({cohort_label})"
    return base


def nice_axis_bounds(values: list[float]) -> tuple[float, float]:
    lo = min(values)
    hi = max(values)
    span = hi - lo
    margin = max(span * 0.12, 0.06)
    lo = min(lo - margin, 1.0 - margin * 0.35)
    hi = max(hi + margin, 1.0 + margin * 0.35)
    if lo <= 0:
        lo = min(v for v in values if v > 0) * 0.8
    span = hi - lo
    if span <= 0.35:
        step = 0.05
    elif span <= 0.8:
        step = 0.10
    elif span <= 1.6:
        step = 0.20
    elif span <= 3.2:
        step = 0.50
    else:
        step = 1.00
    lo = math.floor(lo / step) * step
    hi = math.ceil(hi / step) * step
    lo = max(lo, 0.5)
    return (round(lo, 2), round(hi, 2))


def tick_values(lo: float, hi: float) -> list[float]:
    span = hi - lo
    if span <= 0.35:
        step = 0.05
    elif span <= 0.8:
        step = 0.10
    elif span <= 1.6:
        step = 0.20
    elif span <= 3.2:
        step = 0.50
    else:
        step = 1.00
    ticks: list[float] = []
    current = lo
    while current <= hi + 1e-9:
        ticks.append(round(current, 2))
        current += step
    if 1.0 >= lo and 1.0 <= hi and 1.0 not in ticks:
        ticks.append(1.0)
        ticks.sort()
    if len(ticks) > 6:
        keep = [ticks[0], ticks[len(ticks) // 2], ticks[-1]]
        if 1.0 >= lo and 1.0 <= hi and 1.0 not in keep:
            keep.append(1.0)
        ticks = sorted(set(round(v, 2) for v in keep))
    return ticks


def axis_label(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    if abs(value) >= 2:
        return f"{value:.1f}"
    return f"{value:.2f}"


def xmap(value: float, lo: float, hi: float, left: float, right: float) -> float:
    return left + (value - lo) * (right - left) / (hi - lo)


def draw_forest_panel(
    lines: list[str],
    panel_x: int,
    panel_y: int,
    panel_w: int,
    panel_h: int,
    title: str,
    subtitle: str,
    study_rows: list[dict[str, str]],
    pooled_effect: float,
    pooled_lo: float,
    pooled_hi: float,
    i2: str,
) -> None:
    plot_left = panel_x + 210
    plot_right = panel_x + panel_w - 125
    text_right = panel_x + panel_w - 18
    top = panel_y + 74
    row_gap = 48
    row_y = [top + i * row_gap for i in range(len(study_rows))]
    pooled_y = top + len(study_rows) * row_gap + 8
    axis_y = pooled_y + 36
    values = [pooled_lo, pooled_hi, pooled_effect]
    for row in study_rows:
        values.extend([float(row["ci_lower"]), float(row["ci_upper"]), float(row["effect_estimate"])])
    lo, hi = nice_axis_bounds(values)

    lines.append(f'<rect x="{panel_x}" y="{panel_y}" width="{panel_w}" height="{panel_h}" fill="#ffffff" stroke="#f1f5f9" stroke-width="0.4" rx="10" ry="10"/>')
    lines.append(f'<text x="{panel_x+16}" y="{panel_y+24}" font-size="16" font-weight="700" fill="#111827">{escape(title)}</text>')
    lines.append(f'<text x="{panel_x+16}" y="{panel_y+44}" font-size="12" fill="#4b5563">{escape(subtitle)}</text>')
    lines.append(f'<text x="{panel_x+16}" y="{panel_y+64}" font-size="12" font-weight="700" fill="#111827">Study</text>')
    lines.append(f'<text x="{text_right-86}" y="{panel_y+64}" font-size="12" font-weight="700" fill="#111827">HR (95% CI)</text>')

    if lo < 1 < hi:
        null_x = xmap(1.0, lo, hi, plot_left, plot_right)
        lines.append(f'<line x1="{null_x:.1f}" y1="{top-16}" x2="{null_x:.1f}" y2="{axis_y}" stroke="#9ca3af" stroke-width="1.5" stroke-dasharray="5 4"/>')

    for row, y in zip(study_rows, row_y):
        label = study_forest_label(row["study_id"], row.get("cohort_family", ""))
        est = float(row["effect_estimate"])
        lower = float(row["ci_lower"])
        upper = float(row["ci_upper"])
        weight = float(row.get("random_weight_percent", row.get("fixed_weight_percent", "50")))
        x1 = xmap(max(lower, lo), lo, hi, plot_left, plot_right)
        x2 = xmap(min(upper, hi), lo, hi, plot_left, plot_right)
        xc = xmap(est, lo, hi, plot_left, plot_right)
        size = 8 + math.sqrt(max(weight, 1.0))
        lines.append(f'<text x="{panel_x+16}" y="{y+4}" font-size="12" fill="#111827">{escape(label)}</text>')
        lines.append(f'<line x1="{x1:.1f}" y1="{y}" x2="{x2:.1f}" y2="{y}" stroke="#334155" stroke-width="2"/>')
        lines.append(f'<rect x="{xc-size/2:.1f}" y="{y-size/2:.1f}" width="{size:.1f}" height="{size:.1f}" fill="#2563eb" opacity="0.9"/>')
        lines.append(f'<text x="{text_right-100}" y="{y+4}" font-size="12" fill="#111827">{fmt_effect_label(est, lower, upper)}</text>')

    pooled_x = xmap(pooled_effect, lo, hi, plot_left, plot_right)
    pooled_left = xmap(pooled_lo, lo, hi, plot_left, plot_right)
    pooled_right = xmap(pooled_hi, lo, hi, plot_left, plot_right)
    diamond = [
        f"{pooled_left:.1f},{pooled_y}",
        f"{pooled_x:.1f},{pooled_y-8}",
        f"{pooled_right:.1f},{pooled_y}",
        f"{pooled_x:.1f},{pooled_y+8}",
    ]
    lines.append(f'<text x="{panel_x+16}" y="{pooled_y+4}" font-size="12" font-weight="700" fill="#111827">Random-effects pooled</text>')
    lines.append(f'<polygon points="{" ".join(diamond)}" fill="#0f766e" stroke="#0f766e" stroke-width="1.5"/>')
    lines.append(f'<text x="{text_right-100}" y="{pooled_y+4}" font-size="12" font-weight="700" fill="#111827">{fmt_effect_label(pooled_effect, pooled_lo, pooled_hi)}</text>')

    lines.append(f'<line x1="{plot_left}" y1="{axis_y}" x2="{plot_right}" y2="{axis_y}" stroke="#111827" stroke-width="1.4"/>')
    for tick in tick_values(lo, hi):
        tx = xmap(tick, lo, hi, plot_left, plot_right)
        lines.append(f'<line x1="{tx:.1f}" y1="{axis_y}" x2="{tx:.1f}" y2="{axis_y+6}" stroke="#111827" stroke-width="1"/>')
        lines.append(f'<text x="{tx:.1f}" y="{axis_y+20}" font-size="11" text-anchor="middle" fill="#374151">{axis_label(tick)}</text>')
    lines.append(f'<text x="{(plot_left + plot_right) / 2:.1f}" y="{axis_y+38}" font-size="11" text-anchor="middle" fill="#374151">Hazard ratio</text>')
    lines.append(f'<text x="{panel_x+16}" y="{panel_y+panel_h-16}" font-size="11" fill="#4b5563">Random-effects model; I² = {escape(fmt_i2(i2))}%</text>')


def draw_box_with_text_vector(
    canvas: PSCanvas,
    box: tuple[int, int, int, int],
    title: str,
    body: list[str],
    *,
    fill: str = "#ffffff",
    stroke: str = "#3f4d63",
    title_size: int = 14,
    body_size: int = 11,
) -> None:
    x1, y1, x2, y2 = box
    w = x2 - x1
    h = y2 - y1
    canvas.rect(x1, y1, w, h, fill=fill, stroke=stroke, stroke_width=1.4)
    title_font = load_font(title_size, bold=True)
    body_font = load_font(body_size, bold=False)
    max_width = w - 28
    title_h = wrapped_block_height(_MEASURE_DRAW, title, title_font, max_width, line_gap=2)
    body_h = 0
    for item in body:
        body_h += wrapped_block_height(_MEASURE_DRAW, item, body_font, max_width, line_gap=2) + 4
    total_h = title_h + 8 + body_h
    y = y1 + 14 + max(0, (h - 28 - total_h) // 2)
    y = canvas.wrapped_text(x1 + 14, y, title, font=title_font, size=title_size, max_width=max_width, color="#10243c", bold=True, line_gap=2, align="center") + 8
    for item in body:
        y = canvas.wrapped_text(x1 + 14, y, item, font=body_font, size=body_size, max_width=max_width, color="#1f2937", line_gap=2, align="center") + 4


def draw_arrow_vector(canvas: PSCanvas, start: tuple[int, int], end: tuple[int, int], color: str = "#4b5563", width: float = 2.0) -> None:
    x1, y1 = start
    x2, y2 = end
    canvas.line(x1, y1, x2, y2, color=color, width=width)
    if abs(x2 - x1) < abs(y2 - y1):
        direction = 1 if y2 > y1 else -1
        head = [(x2 - 7, y2 - 12 * direction), (x2 + 7, y2 - 12 * direction), (x2, y2)]
    else:
        direction = 1 if x2 > x1 else -1
        head = [(x2 - 12 * direction, y2 - 7), (x2 - 12 * direction, y2 + 7), (x2, y2)]
    canvas.polygon(head, fill=color, stroke=color, stroke_width=width)


def draw_forest_panel_vector(
    canvas: PSCanvas,
    panel_xy: tuple[int, int, int, int],
    title: str,
    subtitle: str,
    study_rows: list[dict[str, str]],
    pooled_effect: float,
    pooled_lo: float,
    pooled_hi: float,
    i2: str,
) -> None:
    x1, y1, x2, y2 = panel_xy
    panel_w = x2 - x1
    panel_h = y2 - y1
    canvas.rect(x1, y1, panel_w, panel_h, fill="#ffffff", stroke="#f8fafc", stroke_width=0.18)

    title_font = load_font(13, bold=True)
    subtitle_font = load_font(10, bold=False)
    header_font = load_font(10, bold=True)
    label_font = load_font(10, bold=False)
    small_font = load_font(8, bold=False)

    canvas.text(x1 + 12, y1 + 12, title, size=13, bold=True, color="#111827")
    canvas.text(x1 + 12, y1 + 30, subtitle, size=10, color="#5b6574")
    canvas.line(x1 + 12, y1 + 48, x2 - 12, y1 + 48, color="#f0f3f7", width=0.8)
    canvas.text(x1 + 12, y1 + 62, "Study", size=10, bold=True, color="#111827")

    row_effect_labels = [
        fmt_effect_label(float(row["effect_estimate"]), float(row["ci_lower"]), float(row["ci_upper"]))
        for row in study_rows
    ]
    pooled_label = fmt_effect_label(pooled_effect, pooled_lo, pooled_hi)
    effect_width = max(
        [text_width_pixels("HR (95% CI)", header_font)]
        + [text_width_pixels(text, label_font) for text in row_effect_labels + [pooled_label]]
    )

    label_x = x1 + 12
    plot_left = x1 + int(panel_w * 0.34)
    effect_x = x2 - 18 - effect_width
    plot_right = min(x2 - int(panel_w * 0.18), effect_x - 22)
    canvas.text(effect_x, y1 + 62, "HR (95% CI)", size=10, bold=True, color="#111827")

    top = y1 + 90
    row_gap = 42
    row_y = [top + i * row_gap for i in range(len(study_rows))]
    pooled_y = top + len(study_rows) * row_gap + 4
    axis_y = pooled_y + 22

    values = [pooled_lo, pooled_hi, pooled_effect]
    for row in study_rows:
        values.extend([float(row["ci_lower"]), float(row["ci_upper"]), float(row["effect_estimate"])])
    lo, hi = nice_axis_bounds(values)
    if lo < 1 < hi:
        null_x = xmap(1.0, lo, hi, plot_left, plot_right)
        canvas.line(null_x, top - 12, null_x, axis_y, color="#b0b8c3", width=1.2)

    for row, effect_label, y in zip(study_rows, row_effect_labels, row_y):
        label = study_forest_label(row["study_id"], row.get("cohort_family", ""))
        est = float(row["effect_estimate"])
        lower = float(row["ci_lower"])
        upper = float(row["ci_upper"])
        weight = float(row.get("random_weight_percent", row.get("fixed_weight_percent", "50")))
        x_low = xmap(max(lower, lo), lo, hi, plot_left, plot_right)
        x_high = xmap(min(upper, hi), lo, hi, plot_left, plot_right)
        x_mid = xmap(est, lo, hi, plot_left, plot_right)
        square = 6 + math.sqrt(max(weight, 1.0)) * 0.9
        wrap_width = plot_left - label_x - 20
        canvas.wrapped_text(label_x, y - 8, label, font=label_font, size=10, max_width=wrap_width, color="#111827", line_gap=1)
        canvas.line(x_low, y, x_high, y, color="#334155", width=1.5)
        canvas.rect(x_mid - square / 2, y - square / 2, square, square, fill="#2f63d8", stroke="#2856be", stroke_width=0.6)
        canvas.text(effect_x, y - 8, effect_label, size=10, color="#111827")

    diamond = [
        (xmap(pooled_lo, lo, hi, plot_left, plot_right), pooled_y),
        (xmap(pooled_effect, lo, hi, plot_left, plot_right), pooled_y - 8),
        (xmap(pooled_hi, lo, hi, plot_left, plot_right), pooled_y),
        (xmap(pooled_effect, lo, hi, plot_left, plot_right), pooled_y + 8),
    ]
    canvas.text(label_x, pooled_y - 8, "Random-effects pooled", size=10, bold=True, color="#111827")
    canvas.polygon(diamond, fill="#0f766e", stroke="#0f766e", stroke_width=0.8)
    canvas.text(effect_x, pooled_y - 8, pooled_label, size=10, color="#111827")

    canvas.line(plot_left, axis_y, plot_right, axis_y, color="#111827", width=1.2)
    for tick in tick_values(lo, hi):
        tx = xmap(tick, lo, hi, plot_left, plot_right)
        canvas.line(tx, axis_y, tx, axis_y + 5, color="#111827", width=0.8)
        canvas.text(tx, axis_y + 10, axis_label(tick), size=8, color="#374151", anchor="center")
    canvas.text((plot_left + plot_right) / 2, axis_y + 22, "Hazard ratio", size=8, color="#374151", anchor="center")
    canvas.text(label_x, y2 - 16, f"Random-effects model; I² = {fmt_i2(i2)}%", size=8, color="#4b5563")


def build_prisma_vector_eps(path: Path) -> None:
    counts = parse_prisma_counts()
    duplicates_removed = counts["pubmed_ids"] + counts["wos_exported"] + counts["embase_exported"] - counts["screened_total"]
    canvas = PSCanvas(1100, 660)
    draw_box_with_text_vector(
        canvas,
        (320, 18, 780, 116),
        "Records identified from databases",
        [
            f"Total: {counts['pubmed_ids'] + counts['wos_exported'] + counts['embase_exported']}",
            f"PubMed: {counts['pubmed_ids']}",
            f"Web of Science Core Collection: {counts['wos_exported']}",
            f"Embase: {counts['embase_exported']}",
        ],
        title_size=13,
        body_size=10,
    )
    draw_box_with_text_vector(canvas, (60, 155, 350, 228), "Duplicate records removed before screening", [f"{duplicates_removed} records"], title_size=13, body_size=10)
    draw_box_with_text_vector(canvas, (400, 168, 760, 241), "Records screened", [f"{counts['screened_total']} records"], title_size=13, body_size=10)
    draw_box_with_text_vector(canvas, (28, 300, 318, 373), "Records excluded", [f"{counts['title_abstract_excluded']} records"], title_size=13, body_size=10)
    draw_box_with_text_vector(canvas, (400, 300, 760, 373), "Reports sought for retrieval", [f"{counts['fulltext_reviewed']} reports"], title_size=13, body_size=10)
    draw_box_with_text_vector(canvas, (800, 300, 1060, 373), "Reports not retrieved", ["0 reports"], title_size=13, body_size=10)
    draw_box_with_text_vector(canvas, (400, 442, 760, 515), "Reports assessed for eligibility", [f"{counts['fulltext_reviewed']} reports"], title_size=13, body_size=10)
    draw_box_with_text_vector(
        canvas,
        (770, 405, 1085, 568),
        "Reports excluded, with reasons",
        [
            f"{counts['fulltext_not_retained']} reports",
            f"Narrative-only/nonextractable: {counts['narrative_or_nonextractable']}",
            f"Scope/intervention exclusions: {counts['scope_or_protocol_excluded']}",
            f"Specialized/context/noncanonical: {counts['other_nonretained_context']}",
        ],
        title_size=13,
        body_size=10,
    )
    draw_box_with_text_vector(
        canvas,
        (400, 574, 760, 650),
        "Studies included in quantitative evidence set",
        [f"{counts['included_articles']} unique articles", f"{counts['historical_rows']} cohort-level rows in the historical extraction master"],
        fill="#f6fffb",
        stroke="#0f766e",
        title_size=13,
        body_size=10,
    )
    canvas.line(550, 116, 550, 137, color="#4b5563", width=2)
    draw_arrow_vector(canvas, (550, 137), (350, 191))
    draw_arrow_vector(canvas, (550, 137), (580, 168))
    canvas.line(440, 241, 440, 336, color="#4b5563", width=2)
    draw_arrow_vector(canvas, (440, 336), (318, 336))
    draw_arrow_vector(canvas, (580, 241), (580, 300))
    draw_arrow_vector(canvas, (760, 336), (800, 336))
    draw_arrow_vector(canvas, (580, 373), (580, 442))
    draw_arrow_vector(canvas, (760, 486), (770, 486))
    draw_arrow_vector(canvas, (580, 515), (580, 574))
    canvas.save_eps(path)


def build_primary_vector_eps(path: Path) -> None:
    meta_rows = {row["analysis_cell_id"]: row for row in read_tsv(PRIMARY_META / "meta_summary.tsv")}
    canvas = PSCanvas(1140, 740)
    panel_specs = [
        ("HB__CVD_mortality__categorical_high_vs_low", "A. HB and CVD mortality", "High vs low; 1 report / 2 cohorts", (18, 14, 558, 352)),
        ("HB__CVD_mortality__continuous_log", "B. HB and CVD mortality", "Continuous log-scale; 1 report / 2 cohorts", (582, 14, 1122, 352)),
        ("HB__all-cause_mortality__continuous_standardized", "C. HB and all-cause mortality", "Per 1 SD; 1 report / 2 cohorts", (18, 386, 558, 724)),
        ("SASHB__incident_heart_failure__continuous_standardized", "D. SASHB and incident HF", "Per 1 SD; men only; 1 report / 2 cohorts", (582, 386, 1122, 724)),
    ]
    for cell_id, title, subtitle, box in panel_specs:
        study_rows = read_tsv(PRIMARY_META / f"study_level_weights_{cell_id.replace('all-cause', 'all_cause')}.tsv")
        meta = meta_rows[cell_id]
        draw_forest_panel_vector(
            canvas,
            box,
            title,
            subtitle,
            study_rows,
            float(meta["pooled_effect_random"]),
            float(meta["pooled_ci_random_lower"]),
            float(meta["pooled_ci_random_upper"]),
            meta["i2_percent"],
        )
    canvas.save_eps(path)


def build_single_vector_eps(path: Path, analysis_id: str, title: str, subtitle: str) -> None:
    meta_row = next(row for row in read_tsv(AF_META / "meta_summary.tsv") if row["analysis_id"] == analysis_id)
    study_rows = read_tsv(AF_META / f"study_level_weights_{analysis_id}.tsv")
    canvas = PSCanvas(860, 360)
    draw_forest_panel_vector(
        canvas,
        (18, 18, 842, 324),
        title,
        subtitle,
        study_rows,
        float(meta_row["pooled_effect_random"]),
        float(meta_row["pooled_ci_random_lower"]),
        float(meta_row["pooled_ci_random_upper"]),
        meta_row["i2_percent"],
    )
    canvas.save_eps(path)


def write_primary_composite_svg(path: Path) -> None:
    meta_rows = {row["analysis_cell_id"]: row for row in read_tsv(PRIMARY_META / "meta_summary.tsv")}
    panel_specs = [
        (
            "HB__CVD_mortality__categorical_high_vs_low",
            "A. HB and CVD mortality",
            "High vs low exposure",
            28,
            40,
        ),
        (
            "HB__CVD_mortality__continuous_log",
            "B. HB and CVD mortality",
            "Continuous log-scale exposure",
            584,
            40,
        ),
        (
            "HB__all-cause_mortality__continuous_standardized",
            "C. HB and all-cause mortality",
            "Per 1-SD HB increase",
            28,
            388,
        ),
        (
            "SASHB__incident_heart_failure__continuous_standardized",
            "D. SASHB and incident HF",
            "Per 1-SD increase; men only",
            584,
            388,
        ),
    ]
    lines = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="1140" height="710" viewBox="0 0 1140 710">',
        '<rect x="0" y="0" width="1140" height="710" fill="#ffffff"/>',
    ]
    for cell_id, title, subtitle, x, y in panel_specs:
        study_rows = read_tsv(PRIMARY_META / f"study_level_weights_{cell_id.replace('all-cause', 'all_cause')}.tsv")
        meta = meta_rows[cell_id]
        draw_forest_panel(
            lines,
            x,
            y,
            528,
            300,
            title,
            subtitle,
            study_rows,
            float(meta["pooled_effect_random"]),
            float(meta["pooled_ci_random_lower"]),
            float(meta["pooled_ci_random_upper"]),
            meta["i2_percent"],
        )
    lines.append("</svg>")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_single_forest_svg(path: Path, analysis_id: str, title: str, subtitle: str) -> None:
    meta_row = next(row for row in read_tsv(AF_META / "meta_summary.tsv") if row["analysis_id"] == analysis_id)
    study_rows = read_tsv(AF_META / f"study_level_weights_{analysis_id}.tsv")
    lines = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="860" height="360" viewBox="0 0 860 360">',
        '<rect x="0" y="0" width="860" height="360" fill="#ffffff"/>',
    ]
    draw_forest_panel(
        lines,
        24,
        24,
        812,
        300,
        title,
        subtitle,
        study_rows,
        float(meta_row["pooled_effect_random"]),
        float(meta_row["pooled_ci_random_lower"]),
        float(meta_row["pooled_ci_random_upper"]),
        meta_row["i2_percent"],
    )
    lines.append("</svg>")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def draw_forest_panel_pil(
    draw: ImageDraw.ImageDraw,
    panel_xy: tuple[int, int, int, int],
    title: str,
    subtitle: str,
    study_rows: list[dict[str, str]],
    pooled_effect: float,
    pooled_lo: float,
    pooled_hi: float,
    i2: str,
    *,
    scale: float = 1.0,
) -> None:
    x1, y1, x2, y2 = panel_xy
    panel_w = x2 - x1
    draw.rectangle(panel_xy, outline="#e8edf3", width=max(1, px(1, scale)), fill="white")
    title_font = load_font(px(22, scale), bold=True)
    subtitle_font = load_font(px(15, scale), bold=False)
    header_font = load_font(px(14, scale), bold=True)
    label_font = load_font(px(17, scale), bold=False)
    small_font = load_font(px(13, scale), bold=False)
    draw.text((x1 + px(20, scale), y1 + px(18, scale)), title, font=title_font, fill="#111827")
    draw.text((x1 + px(20, scale), y1 + px(46, scale)), subtitle, font=subtitle_font, fill="#5b6574")
    draw.line((x1 + px(20, scale), y1 + px(68, scale), x2 - px(20, scale), y1 + px(68, scale)), fill="#f3f5f8", width=max(1, px(1, scale)))
    draw.text((x1 + px(20, scale), y1 + px(88, scale)), "Study", font=header_font, fill="#111827")
    header = "HR (95% CI)"
    header_box = draw.textbbox((0, 0), header, font=header_font)

    row_effect_labels = [
        fmt_effect_label(float(row["effect_estimate"]), float(row["ci_lower"]), float(row["ci_upper"]))
        for row in study_rows
    ]
    pooled_effect_label = fmt_effect_label(pooled_effect, pooled_lo, pooled_hi)
    effect_width = max(
        [header_box[2] - header_box[0]]
        + [draw.textbbox((0, 0), text, font=label_font)[2] for text in row_effect_labels + [pooled_effect_label]]
    )

    label_x = x1 + px(20, scale)
    plot_left = x1 + int(panel_w * 0.34)
    right_padding = px(24, scale)
    effect_x = x2 - right_padding - effect_width
    plot_right = min(x2 - int(panel_w * 0.18), effect_x - px(34, scale))
    draw.text((effect_x, y1 + px(88, scale)), header, font=header_font, fill="#111827")
    top = y1 + px(128, scale)
    row_gap = px(68, scale)
    row_y = [top + i * row_gap for i in range(len(study_rows))]
    pooled_y = top + len(study_rows) * row_gap + px(6, scale)
    axis_y = pooled_y + px(34, scale)

    values = [pooled_lo, pooled_hi, pooled_effect]
    for row in study_rows:
        values.extend([float(row["ci_lower"]), float(row["ci_upper"]), float(row["effect_estimate"])])
    lo, hi = nice_axis_bounds(values)
    if lo < 1 < hi:
        null_x = xmap(1.0, lo, hi, plot_left, plot_right)
        draw.line((null_x, top - px(20, scale), null_x, axis_y), fill="#b0b8c3", width=max(1, px(2, scale)))

    for row, effect_label, y in zip(study_rows, row_effect_labels, row_y):
        label = study_forest_label(row["study_id"], row.get("cohort_family", ""))
        est = float(row["effect_estimate"])
        lower = float(row["ci_lower"])
        upper = float(row["ci_upper"])
        weight = float(row.get("random_weight_percent", row.get("fixed_weight_percent", "50")))
        x_low = xmap(max(lower, lo), lo, hi, plot_left, plot_right)
        x_high = xmap(min(upper, hi), lo, hi, plot_left, plot_right)
        x_mid = xmap(est, lo, hi, plot_left, plot_right)
        sq = px(11, scale) + int(math.sqrt(max(weight, 1.0)) * 1.4 * scale)
        wrap_width = plot_left - label_x - px(30, scale)
        draw_wrapped_text(draw, (label_x, y - px(12, scale)), label, label_font, wrap_width, "#111827", line_gap=max(1, px(1, scale)))
        draw.line((x_low, y, x_high, y), fill="#334155", width=max(1, px(3, scale)))
        draw.rectangle((x_mid - sq / 2, y - sq / 2, x_mid + sq / 2, y + sq / 2), fill="#2f63d8", outline="#2856be", width=max(1, px(1, scale)))
        draw.text((effect_x, y - px(12, scale)), effect_label, font=label_font, fill="#111827")

    diamond = [
        (xmap(pooled_lo, lo, hi, plot_left, plot_right), pooled_y),
        (xmap(pooled_effect, lo, hi, plot_left, plot_right), pooled_y - px(12, scale)),
        (xmap(pooled_hi, lo, hi, plot_left, plot_right), pooled_y),
        (xmap(pooled_effect, lo, hi, plot_left, plot_right), pooled_y + px(12, scale)),
    ]
    draw.text((label_x, pooled_y - px(12, scale)), "Random-effects pooled", font=header_font, fill="#111827")
    draw.polygon(diamond, fill="#0f766e", outline="#0f766e")
    draw.text((effect_x, pooled_y - px(12, scale)), pooled_effect_label, font=label_font, fill="#111827")

    draw.line((plot_left, axis_y, plot_right, axis_y), fill="#111827", width=max(1, px(2, scale)))
    for tick in tick_values(lo, hi):
        tx = xmap(tick, lo, hi, plot_left, plot_right)
        draw.line((tx, axis_y, tx, axis_y + px(8, scale)), fill="#111827", width=max(1, px(2, scale)))
        label = axis_label(tick)
        tw = draw.textbbox((0, 0), label, font=small_font)[2]
        draw.text((tx - tw / 2, axis_y + px(12, scale)), label, font=small_font, fill="#374151")
    axis_caption = "Hazard ratio"
    tw = draw.textbbox((0, 0), axis_caption, font=small_font)[2]
    draw.text(((plot_left + plot_right - tw) / 2, axis_y + px(28, scale)), axis_caption, font=small_font, fill="#374151")
    draw.text((label_x, axis_y + px(48, scale)), f"Random-effects model; I² = {fmt_i2(i2)}%", font=small_font, fill="#4b5563")


def make_primary_composite_figure() -> Image.Image:
    scale = FIG_NATIVE_SCALE
    img = Image.new("RGB", (px(2200, scale), px(1320, scale)), "white")
    draw = ImageDraw.Draw(img)
    meta_rows = {row["analysis_cell_id"]: row for row in read_tsv(PRIMARY_META / "meta_summary.tsv")}
    panel_specs = [
        ("HB__CVD_mortality__categorical_high_vs_low", "A. HB and CVD mortality", "High vs low; 1 report / 2 cohorts", (30, 20, 1070, 620)),
        ("HB__CVD_mortality__continuous_log", "B. HB and CVD mortality", "Continuous log-scale; 1 report / 2 cohorts", (1130, 20, 2170, 620)),
        ("HB__all-cause_mortality__continuous_standardized", "C. HB and all-cause mortality", "Per 1 SD; 1 report / 2 cohorts", (30, 680, 1070, 1280)),
        ("SASHB__incident_heart_failure__continuous_standardized", "D. SASHB and incident HF", "Per 1 SD; men only; 1 report / 2 cohorts", (1130, 680, 2170, 1280)),
    ]
    for cell_id, title, subtitle, box in panel_specs:
        study_rows = read_tsv(PRIMARY_META / f"study_level_weights_{cell_id.replace('all-cause', 'all_cause')}.tsv")
        meta = meta_rows[cell_id]
        draw_forest_panel_pil(
            draw,
            scale_box(box, scale),
            title,
            subtitle,
            study_rows,
            float(meta["pooled_effect_random"]),
            float(meta["pooled_ci_random_lower"]),
            float(meta["pooled_ci_random_upper"]),
            meta["i2_percent"],
            scale=scale,
        )
    return img


def make_single_forest_figure(analysis_id: str, title: str, subtitle: str) -> Image.Image:
    scale = FIG_NATIVE_SCALE
    img = Image.new("RGB", (px(1040, scale), px(820, scale)), "white")
    draw = ImageDraw.Draw(img)
    meta_row = next(row for row in read_tsv(AF_META / "meta_summary.tsv") if row["analysis_id"] == analysis_id)
    study_rows = read_tsv(AF_META / f"study_level_weights_{analysis_id}.tsv")
    draw_forest_panel_pil(
        draw,
        scale_box((24, 24, 1016, 760), scale),
        title,
        subtitle,
        study_rows,
        float(meta_row["pooled_effect_random"]),
        float(meta_row["pooled_ci_random_lower"]),
        float(meta_row["pooled_ci_random_upper"]),
        meta_row["i2_percent"],
        scale=scale,
    )
    return img


def build_protocol_search_appendix(path: Path) -> None:
    counts = parse_prisma_counts()
    exclusion_rows = fulltext_exclusion_rows()
    duplicates_removed = counts["pubmed_ids"] + counts["wos_exported"] + counts["embase_exported"] - counts["screened_total"]
    protocol_outline = textwrap.dedent(
        f"""\
        # Additional file 1. Protocol and search appendix

        This appendix reports the final protocol rules, the exact executed search strings, the final database yields, the eligibility-assessment accounting, and the operational synthesis rules used for the submitted review package. Development-stage logs, pilot counts, local file paths, and intermediate parsing notes are intentionally omitted from this submission version.

        ## Review question and registration status

        - Working review title: `Beyond-AHI hypoxic metrics and hard cardiovascular and mortality outcomes in adult OSA-related cohorts: a systematic review and meta-analysis`
        - Reporting framework: `PRISMA 2020`
        - Additional reporting cross-check: `MOOSE-relevant items for observational meta-analyses`
        - Registration: protocol outline prepared in a `PROSPERO`-compatible format but not formally registered
        - Protocol freeze date: `2026-03-23` before full-text eligibility assessment and before the final pooled synthesis
        - Core review question: in adult OSA-related cohorts, which beyond-AHI hypoxic metrics show the strongest and most defensible associations with hard cardiovascular or mortality outcomes?

        ## Final eligibility framework

        Eligible reports were original adult human studies with at least one prespecified beyond-AHI hypoxic metric (`HB`, `SASHB`, `T90`, `ODI`, minimum or nadir oxygen saturation) and at least one hard clinical outcome (all-cause mortality, cardiovascular mortality, incident or adjudicated cardiovascular events, incident heart failure, incident atrial fibrillation, or incident stroke). For interpretation, cohorts were grouped into three classes:

        - `clinical OSA/referral cohorts`
        - `community cohorts with OSA-related physiology`
        - `specialized cardiovascular/surgical cohorts`

        Pediatric studies, animal studies, diagnostic-only reports, surrogate-only outcome studies, central-sleep-apnea-focused cohorts without separable OSA results, and non-original or non-extractable reports were excluded from the primary quantitative pool.

        ## Final synthesis rules

        - metric families were kept separate
        - categorical contrasts were not pooled with continuous-scale analyses
        - materially different composite outcomes were not pooled by default
        - adjusted hazard ratios were prioritized
        - when multiple rows from the same cohort family addressed the same metric-outcome cell, default row selection prioritized adjusted hazard ratios, prespecified core metric families over comparator constructs, primary published models over alternate incremental models, and the cleanest shared scale before alternate subgroup or threshold rows
        - alternate scales or models from the same cohort family were retained as overlap-sensitive sensitivity evidence rather than merged into the default pooled cell
        - random-effects meta-analysis with restricted maximum likelihood was used only for directly comparable cells
        - post hoc scale harmonization was labeled exploratory
        - because every poolable cell contained fewer than 10 studies, funnel plots and small-study-effects testing were not performed

        ## Default row-selection algorithm

        | Priority step | Operational rule |
        | --- | --- |
        | 1 | Prefer adjusted hazard ratios over unadjusted, logistic, or purely descriptive rows |
        | 2 | Prefer prespecified core metric families over comparator constructs from the same cohort family |
        | 3 | Prefer the primary published model over alternate incremental or subtype-adjusted models |
        | 4 | Prefer the cleanest shared exposure scale before subgroup-only, threshold-only, or overlap-sensitive rows |
        | 5 | Retain alternate rows from the same cohort family as sensitivity/comparator evidence rather than merge them into the primary pooled analyses |

        ## Protocol deviations and late analytic clarifications

        - The interpretive framing was narrowed from `adult OSA` to `adult OSA-related cohorts` after full eligibility assessment because several retained community cohorts were not restricted to clinic-defined OSA at entry; the prespecified metric and outcome families were not changed.
        - The atrial-fibrillation T90 harmonization was introduced after protocol drafting and is retained only as an exploratory sensitivity analysis.
        - Metric-specific PubMed side searches were used as a supplementary recall audit after closure of the main PubMed corpus; they did not create a new primary pooled cell.

        ## Final search yields and screening position

        - database records identified before deduplication: `{counts['pubmed_ids'] + counts['wos_exported'] + counts['embase_exported']}`
        - duplicate records removed before screening: `{duplicates_removed}`
        - PubMed main query identified: `{counts['pubmed_ids']}` records
        - PubMed records parsed into the screening corpus after export cleaning: `{counts['pubmed_screened']}` records
        - PubMed IDs not parsed into the screening corpus after export cleaning: `{counts['pubmed_not_parsed']}` records
        - The 3 PubMed IDs not retained in the parsed screening corpus were removed during export cleaning because their saved export blocks did not normalize into complete screening rows.
        - Web of Science Core Collection exported after document-type filtering: `{counts['wos_exported']}` records
        - Embase exported after source/publication-type filtering: `{counts['embase_exported']}` records
        - deduplicated title/abstract screening corpus across the three planned databases: `{counts['screened_total']}` records
        - reports sought for retrieval after title/abstract screening: `{counts['fulltext_reviewed']}`
        - reports not retrieved: `0`
        - full-text reports assessed for eligibility against the prespecified protocol rules: `{counts['fulltext_reviewed']}`
        - unique articles contributing to the historical executed quantitative evidence set: `{counts['included_articles']}`

        Reports not retained in the historical executed quantitative evidence set after full-text assessment fell mainly into three broad groups, while a later-rescued subset was retained into the post-freeze supplement:

        - later re-adjudicated into the post-freeze supplement: `{counts['later_retained_postfreeze']}`
        - narrative-only or nonextractable reports: `{counts['narrative_or_nonextractable']}`
        - protocol-scope or intervention-effect-modifier exclusions: `{counts['scope_or_protocol_excluded']}`
        - specialized/context/noncanonical comparator reports that did not yield a prespecified retained row: `{counts['other_nonretained_context']}`

        ## Re-review audit summary

        | Audit object | Re-reviewed set | Review structure | Final effect on the submitted core |
        | --- | --- | --- | --- |
        | Included articles | `{counts['updated_included_articles']}` articles in the final submission dataset (`{counts['included_articles']}` historical + `{counts['upgrade_retained_articles']}` post-freeze upgrade articles) | The second author re-reviewed the retained studies before final submission export, with the post-freeze supplement adjudicated under the same extraction rules | The updated submission dataset was expanded without adding a new primary pooled cell |
        | Full-text non-retained reports | `{counts['fulltext_not_retained']}/{counts['fulltext_not_retained']}` reports | The second author re-reviewed all non-retained full-text decisions; disagreements were adjudicated by the corresponding author | Final non-retained classes were retained and summarized as final-state exclusion categories |
        | Primary pooled inputs | `4/4` pooled cells (`8` cohort-specific estimates) | All effect estimates contributing to the four primary pooled cells were re-checked before export | The four-cell primary pooled structure was retained without change |
        | Non-pooled retained rows | `{counts['updated_nonpooled_rows']}/{counts['updated_nonpooled_rows']}` retained non-pooled rows | The second author re-reviewed all retained sensitivity/comparator and narrative-supporting rows against the extraction worksheet and article PDFs or abstracts | Row-level retention outside the primary pooled analyses was preserved while the post-freeze upgrade thickened the T90/TST90 and specialized-evidence layers |

        ## Post-freeze evidence-upgrade supplement

        After the historical executed package was frozen, we performed a targeted post-freeze evidence-upgrade supplement focused on high-value full texts and open-access anchors identified during manuscript finalization. A final strict-review re-adjudication also rescued one previously screened dual-cohort T90 mortality paper into the updated dataset. This supplement did not alter the Figure 1 historical PRISMA accounting and was adjudicated using the same protocol-concordant extraction and retention rules.

        - targeted studies/open-access anchors reviewed: `8`
        - retained into the updated submission dataset: `{counts['upgrade_retained_articles']}` studies contributing `{counts['upgrade_retained_rows']}` cohort-level rows
        - contextual-only specialized paper acknowledged but not rowed: `1` (`Pinilla 2023`, PMID `37734857`)
        - updated final submission dataset: `{counts['updated_included_articles']}` unique articles, `{counts['updated_rows']}` cohort-level rows, `{counts['updated_primary_rows']}` primary retained rows, and `{counts['updated_sensitivity_rows']}` sensitivity/comparator rows
        - historical executed package preserved for Figure 1 and screening accounting: `{counts['included_articles']}` unique articles and `{counts['historical_rows']}` cohort-level rows
        - effect on the four primary pooled cells: `no new pooled cell was added and the four-cell primary pooled structure remained unchanged`

        ## Final anchor-centered citation-chasing completeness pass

        On `2026-03-26`, we completed a targeted anchor-centered citation-chasing completeness pass around the main HB/SASHB mortality or heart-failure anchors, the T90/ODI atrial-fibrillation and mortality anchors, and all studies carried into the post-freeze evidence-upgrade supplement. Candidate follow-on papers were checked against primary-source `PubMed`, `PMC`, or journal DOI records under the same protocol-concordant retention rules.

        - anchor set interrogated: `Azarbarzin 2019/2020`, `Labarca 2023`, `Blanchard 2021`, `Baumert 2020`, `Oldenburg 2016`, `Heinzinger 2023`, `Kendzerska 2018`, `Trzepizur 2022`, `Hui 2024`, `Vichova 2025`, `Mazzotti 2025`, plus the retained post-freeze upgrade studies
        - strongest rescued article already integrated: `Henríquez-Beltrán 2024` (PMID `37656346`)
        - final result of this pass: `no additional protocol-concordant retained study or retained cohort-level row beyond the integrated 31-article / 54-row updated submission dataset`
        - effect on the four primary pooled cells: `no new independent publication-level replication was identified for HB -> cardiovascular mortality, HB -> all-cause mortality, or SASHB -> incident heart failure`
        - highest-value screened but non-retained candidates confirmed during this pass:
          - `Xu 2026` (PMID `41794120`): composite high-CVD-risk OSA phenotype based on high HB or high ΔHR rather than a separable prespecified metric-family estimate
          - `Zheng 2025` (PMID `41478496`): specialized ACS pooled cohort with inverse or U-shaped `TSA90` associations rather than a general OSA prognostic anchor
          - `Yan 2024` (PMID `37772691`): community `SpO2_TOTAL` construct rather than a prespecified event-based hypoxic metric family
          - `Parekh 2023` (PMID `37698405`): ventilatory burden, a novel but noncanonical construct outside the prespecified exposure families
        - implication: `after the three-database package, the post-freeze upgrade supplement, and the final citation-chasing pass, the main remaining ceiling is publication-level replication scarcity rather than unresolved search incompleteness`

        ## Supplementary PubMed side-search audit

        Metric-specific PubMed side searches were executed as a supplementary recall audit after closure of the main PubMed corpus, not as an independent fourth database stream.

        - unique side-search records versus the main PubMed query: `{counts['side_unique']}`
        - high-priority audit records: `{counts['side_high']}`
        - medium-priority audit records: `{counts['side_medium']}`
        - low-priority audit records: `{counts['side_low']}`
        - likely exclusions on first-pass triage: `{counts['side_likely_exclude']}`
        - final impact on retained articles or retained cohort-level rows: `no article or retained row in the final quantitative evidence set depended exclusively on the side-search stream`
        - effect on the primary pooled analyses: `no new primary pooled cell added`
        """
    )
    exclusion_log = markdown_table_block(
        "Full-text reports not retained after eligibility assessment",
        f"Complete report-level log for all {len(exclusion_rows)} full-text reports that were reviewed but not retained in the historical executed quantitative evidence set.",
        exclusion_rows,
        heading="##",
    )
    pubmed_query = textwrap.dedent(
        """\
        ## Executed PubMed main query

        ```text
        ("Sleep Apnea, Obstructive"[Mesh] OR "Sleep Apnea Syndromes"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab] OR OSA[tiab])
        AND
        ("hypoxic burden"[tiab] OR "oxygen desaturation index"[tiab] OR ODI[tiab] OR T90[tiab] OR "sleep hypoxemia"[tiab] OR "sleep hypoxaemia"[tiab] OR "nocturnal hypoxemia"[tiab] OR "nocturnal hypoxaemia"[tiab] OR "minimum oxygen saturation"[tiab] OR "nadir oxygen saturation"[tiab] OR "lowest oxygen saturation"[tiab] OR "nadir SpO2"[tiab])
        AND
        (mortality[tiab] OR "cardiovascular mortality"[tiab] OR "all-cause mortality"[tiab] OR "cardiovascular event*"[tiab] OR MACE[tiab] OR "major adverse cardiovascular"[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR prognosis[tiab] OR incident[tiab])
        NOT
        (child*[tiab] OR pediatric*[tiab] OR paediatric*[tiab])
        ```
        """
    )
    side_search_queries = textwrap.dedent(
        """\
        ## Executed supplementary PubMed side-search queries

        ```text
        HB:
        ("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
        AND ("hypoxic burden"[tiab])
        AND (mortality[tiab] OR "cardiovascular mortality"[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

        T90 / sleep hypoxemia:
        ("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
        AND (T90[tiab] OR "sleep hypoxemia"[tiab] OR "sleep hypoxaemia"[tiab] OR "time below 90"[tiab] OR "minimum oxygen saturation below 90"[tiab])
        AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

        ODI:
        ("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
        AND ("oxygen desaturation index"[tiab] OR ODI[tiab])
        AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

        Nadir / minimum oxygen saturation:
        ("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
        AND ("minimum oxygen saturation"[tiab] OR "nadir oxygen saturation"[tiab] OR "lowest oxygen saturation"[tiab] OR "nadir SpO2"[tiab])
        AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])
        ```
        """
    )
    wos_query = textwrap.dedent(
        """\
        ## Executed Web of Science Core Collection query

        ```text
        TS=((("obstructive sleep apnea" OR "obstructive sleep apnoea" OR "sleep-disordered breathing" OR "sleep disordered breathing" OR OSA OR OSAHS) AND ("hypoxic burden" OR "sleep apnea-specific hypoxic burden" OR "sleep apnoea-specific hypoxic burden" OR "oxygen desaturation index" OR ODI OR T90 OR "sleep hypoxemia" OR "sleep hypoxaemia" OR "nocturnal hypoxemia" OR "nocturnal hypoxaemia" OR "minimum oxygen saturation" OR "nadir oxygen saturation" OR "lowest oxygen saturation" OR "nadir SpO2" OR "mean oxygen saturation") AND (mortality OR "all-cause mortality" OR "cardiovascular mortality" OR "cardiovascular event*" OR MACE OR MACCE OR "major adverse cardiovascular" OR "heart failure" OR "atrial fibrillation" OR stroke OR prognosis OR incident)) NOT (child* OR pediatric* OR paediatric*))
        ```

        Post-retrieval filters:

        - database: `Web of Science Core Collection`
        - document types: `Article`, `Early Access`
        """
    )
    embase_query = textwrap.dedent(
        """\
        ## Executed Embase query logic

        ```text
        #1 ('obstructive sleep apnea'/exp OR 'obstructive sleep apnea':ti,ab,kw OR 'obstructive sleep apnoea':ti,ab,kw OR 'sleep disordered breathing':ti,ab,kw OR 'sleep-disordered breathing':ti,ab,kw OR osa:ti,ab,kw OR osahs:ti,ab,kw)

        #2 ('hypoxic burden'/exp OR 'hypoxic burden':ti,ab,kw OR 'sleep apnea specific hypoxic burden':ti,ab,kw OR 'sleep apnoea specific hypoxic burden':ti,ab,kw OR 'oxygen desaturation index'/exp OR 'oxygen desaturation index':ti,ab,kw OR odi:ti,ab,kw OR t90:ti,ab,kw OR 'sleep hypoxemia':ti,ab,kw OR 'sleep hypoxaemia':ti,ab,kw OR 'nocturnal hypoxemia':ti,ab,kw OR 'nocturnal hypoxaemia':ti,ab,kw OR 'minimum oxygen saturation':ti,ab,kw OR 'nadir oxygen saturation':ti,ab,kw OR 'lowest oxygen saturation':ti,ab,kw OR 'nadir spo2':ti,ab,kw OR 'mean oxygen saturation':ti,ab,kw)

        #3 ('mortality'/exp OR mortality:ti,ab,kw OR 'all cause mortality':ti,ab,kw OR 'cardiovascular mortality':ti,ab,kw OR 'cardiovascular event*':ti,ab,kw OR mace:ti,ab,kw OR macce:ti,ab,kw OR 'major adverse cardiovascular':ti,ab,kw OR 'heart failure'/exp OR 'heart failure':ti,ab,kw OR 'atrial fibrillation'/exp OR 'atrial fibrillation':ti,ab,kw OR stroke/exp OR stroke:ti,ab,kw OR prognosis/exp OR prognosis:ti,ab,kw OR incident:ti,ab,kw)

        #4 (child/exp OR adolescent/exp OR pediatric*:ti,ab,kw OR paediatric*:ti,ab,kw OR child*:ti,ab,kw)

        Final logic: #1 AND #2 AND #3 NOT #4
        ```

        Post-retrieval filters:

        - publication types: `Article`, `Article in Press`
        - source limit: `[embase]/lim`
        """
    )
    dedup_notes = textwrap.dedent(
        """\
        ## Deduplication and evidence-layer notes

        - Deduplication was performed sequentially against the closed corpus using PMID, DOI, and normalized title matching, followed by manual review of residual ambiguous records.
        - The final quantitative evidence set was built at the cohort-row level rather than the article level.
        - Alternate models or scales from the same cohort family were retained as overlap-sensitive sensitivity evidence rather than merged into the primary pooled analyses.
        - Narrative-only and noncanonical comparator records were tracked separately from prespecified metric-family rows.
        """
    )
    text = "\n\n".join([protocol_outline, exclusion_log, pubmed_query, side_search_queries, wos_query, embase_query, dedup_notes])
    path.write_text(text + "\n", encoding="utf-8")


def write_markdown(path: Path, text: str) -> None:
    ensure_dir(path.parent)
    path.write_text(text.rstrip() + "\n", encoding="utf-8")


def build_submission_manuscript_markdown() -> str:
    base = (ROOT / "manuscript.md").read_text(encoding="utf-8").rstrip()
    page_break = "\n".join(
        [
            "```{=openxml}",
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>',
            "```",
        ]
    )
    base = base.replace("\nAbstract\n", "\n\n" + page_break + "\n\nAbstract\n", 1)
    table_blocks = [
        markdown_table_block(
            "Table 1. Cohort map of included articles",
            "Table 1 summarizes retained articles by cohort class, main metric family, outcome anchor, analytic sample size, evidence layer, and overall risk of bias. For multi-cohort articles, cohort-specific sample sizes are shown with cohort-labeled semicolon-separated entries within the same cell; slash-separated values within a cohort label indicate multiple retained rows from the same cohort family.",
            build_main_table1(),
            heading="",
        ),
        markdown_table_block(
            "Table 2. Primary pooled analyses",
            "Each line pools two cohort-specific estimates from one publication reporting two non-overlapping cohorts; these cells represent cohort-level comparability rather than multi-publication replication.",
            build_table2(),
            heading="",
        ),
        markdown_table_block(
            "Table 3. Sensitivity analyses",
            "Focused competing-risk and exploratory harmonization checks for the primary pooled analyses. Leave-one-out ranges are provided in Additional file 7.",
            build_table3(),
            heading="",
        ),
    ]
    output = base
    for block in table_blocks:
        output += "\n\n" + block.strip()
    return output.rstrip() + "\n"


def export_tables() -> None:
    table_specs = [
        (
            "Table_1_study_characteristics",
            "Table 1. Cohort map of included articles",
            "Article-level cohort map. Cls is shown as Clin, Comm, or Spec. For multi-cohort articles, cohort-specific sample sizes and follow-up values are shown as cohort-labeled semicolon-separated entries within the same cell; slash-separated values within a cohort label indicate multiple retained rows from the same cohort family.",
            build_table1(),
        ),
        (
            "Table_2_primary_pooled_results",
            "Table 2. Primary pooled analyses",
            "Each line pools two cohort-specific estimates from one publication reporting two non-overlapping cohorts; these cells therefore represent cohort-level comparability rather than multi-publication replication.",
            build_table2(),
        ),
        (
            "Table_3_sensitivity_results",
            "Table 3. Sensitivity analyses",
            "Focused competing-risk and exploratory harmonization checks for the primary pooled analyses. Leave-one-out ranges are provided in Additional file 7.",
            build_table3(),
        ),
    ]
    for stem, title, subtitle, rows in table_specs:
        csv_path = UPLOAD_TABLES / f"{stem}.csv"
        md_path = csv_path.with_suffix(".md")
        docx_path = csv_path.with_suffix(".docx")
        write_csv(csv_path, rows)
        write_markdown_table(md_path, title, subtitle, rows)
        run_pandoc(md_path, docx_path)
        table_width_map = {
            "Table_1_study_characteristics": [[1850, 1100, 900, 2550, 4300, 1850, 1450]],
            "Table_2_primary_pooled_results": [[1100, 1700, 1350, 2500, 2200, 900]],
            "Table_3_sensitivity_results": [[2500, 1700, 2100, 900, 3000]],
        }
        patch_docx(
            docx_path,
            landscape=(stem == "Table_1_study_characteristics"),
            repeat_table_headers=True,
            narrow_margins=True,
            table_widths=table_width_map.get(stem),
            line_spacing_twips=230 if stem == "Table_1_study_characteristics" else 220,
            table_line_spacing_twips=230 if stem == "Table_1_study_characteristics" else 200,
            table_font_size_half_points=18 if stem == "Table_1_study_characteristics" else 19,
        )

    # Additional file 2
    af2_csv = UPLOAD_SUPP / "Additional_file_2_Domain_level_risk_of_bias_table.csv"
    af2_md = af2_csv.with_suffix(".md")
    af2_docx = af2_csv.with_suffix(".docx")
    write_csv(af2_csv, build_domain_rob_table())
    write_multitable_markdown(
        af2_md,
        "Additional file 2. Domain-level risk-of-bias matrix",
        "Compact QUIPS-style domain judgments. L = low, M = moderate, H = high. Overall low risk required no high-risk domain and no more than one moderate concern; overall high risk required at least one high-risk domain or multiple major concerns in confounding/reporting; all other studies were rated moderate.",
        build_domain_rob_sections(),
        "Abbreviations: Part.=participation; Attr.=attrition; Factor=prognostic factor; Confound.=confounding.",
    )
    run_pandoc(af2_md, af2_docx)
    patch_docx(
        af2_docx,
        landscape=True,
        repeat_table_headers=True,
        narrow_margins=True,
        table_widths=[
            [1900, 3300, 800, 800, 800, 900],
            [2500, 1100, 1100, 1100],
        ],
        line_spacing_twips=210,
        table_line_spacing_twips=180,
        table_font_size_half_points=18,
    )

    # Additional file 3
    af3_csv = UPLOAD_SUPP / "Additional_file_3_Nonpooled_evidence_table.csv"
    af3_md = af3_csv.with_suffix(".md")
    af3_docx = af3_csv.with_suffix(".docx")
    write_csv(af3_csv, build_nonpooled_table())
    write_multitable_markdown(
        af3_md,
        "Additional file 3. Non-pooled evidence table",
        "Rows retained outside the four-cell primary pooled analyses. Tables are grouped by metric family to keep the supplement readable and to separate prespecified metric families from comparator constructs.",
        build_nonpooled_sections(),
    )
    run_pandoc(af3_md, af3_docx)
    patch_docx(
        af3_docx,
        landscape=True,
        repeat_table_headers=True,
        narrow_margins=True,
        table_widths=[
            [1700, 1600, 1800, 2500, 3300],
            [1700, 1600, 1800, 2500, 3300],
            [1700, 1600, 1800, 2500, 3300],
            [1700, 1600, 1800, 2500, 3300],
        ],
        line_spacing_twips=240,
        table_line_spacing_twips=200,
        table_font_size_half_points=20,
        prevent_row_splits=True,
    )

    # Additional file 7
    af7_csv = UPLOAD_SUPP / "Additional_file_7_Leave_one_out_ranges.csv"
    af7_md = af7_csv.with_suffix(".md")
    af7_docx = af7_csv.with_suffix(".docx")
    write_csv(af7_csv, build_leave_one_out_table())
    write_markdown_table(
        af7_md,
        "Additional file 7. Leave-one-out stability ranges",
        "Leave-one-out ranges for the four primary pooled cells, reported outside the main manuscript table to keep the core sensitivity table compact.",
        build_leave_one_out_table(),
    )
    run_pandoc(af7_md, af7_docx)
    patch_docx(af7_docx, repeat_table_headers=True, narrow_margins=True, table_widths=[[2600, 2300, 3800]], line_spacing_twips=240, table_line_spacing_twips=200, table_font_size_half_points=20)

    # Additional file 8
    af8_csv = UPLOAD_SUPP / "Additional_file_8_Nonpooled_barriers_summary.csv"
    af8_md = af8_csv.with_suffix(".md")
    af8_docx = af8_csv.with_suffix(".docx")
    write_csv(af8_csv, build_table4())
    write_markdown_table(
        af8_md,
        "Additional file 8. Non-pooled barriers summary",
        "Compact metric-outcome barrier matrix summarizing the main reasons non-pooled evidence remained outside synthesis-ready quantitative pooling.",
        build_table4(),
    )
    run_pandoc(af8_md, af8_docx)
    patch_docx(af8_docx, landscape=True, repeat_table_headers=True, narrow_margins=True, table_widths=[[2100, 1700, 4800]], line_spacing_twips=180, table_line_spacing_twips=160, table_font_size_half_points=18)

    extraction_csv = UPLOAD_SUPP / "Additional_file_5_Extraction_and_synthesis_worksheet.csv"
    write_csv(extraction_csv, build_extraction_worksheet())

    script_txt = UPLOAD_SUPP / "Additional_file_6_Core_analysis_and_figure_script.txt"
    script_txt.write_text(Path(__file__).read_text(encoding="utf-8"), encoding="utf-8")


def export_prisma_figure() -> None:
    pdf_path = UPLOAD_FIGURES / "Figure_1_PRISMA_flow.pdf"
    png_path = UPLOAD_FIGURES / "Figure_1_PRISMA_flow.png"
    eps_path = UPLOAD_FIGURES / "_Figure_1_PRISMA_flow.eps"
    build_prisma_vector_eps(eps_path)
    render_eps_to_outputs(eps_path, pdf_path, png_path)
    eps_path.unlink(missing_ok=True)


def export_primary_figure_panels() -> None:
    fig2_pdf = UPLOAD_FIGURES / "Figure_2_primary_pooled_panels.pdf"
    fig2_png = UPLOAD_FIGURES / "Figure_2_primary_pooled_panels.png"
    fig2_eps = UPLOAD_FIGURES / "_Figure_2_primary_pooled_panels.eps"
    build_primary_vector_eps(fig2_eps)
    render_eps_to_outputs(fig2_eps, fig2_pdf, fig2_png)
    fig2_eps.unlink(missing_ok=True)

    fig3_pdf = UPLOAD_FIGURES / "Figure_3_T90_AF_harmonized.pdf"
    fig3_png = UPLOAD_FIGURES / "Figure_3_T90_AF_harmonized.png"
    fig3_eps = UPLOAD_FIGURES / "_Figure_3_T90_AF_harmonized.eps"
    build_single_vector_eps(
        fig3_eps,
        "T90__incident_atrial_fibrillation__per10pct_harmonized_ci_based",
        "T90 and incident AF",
        "Exploratory sensitivity only; harmonized to per 10-percentage-point increase",
    )
    render_eps_to_outputs(fig3_eps, fig3_pdf, fig3_png)
    fig3_eps.unlink(missing_ok=True)

    supp_pdf = UPLOAD_SUPP / "Additional_file_4_AF_precision_check_figure.pdf"
    supp_png = UPLOAD_SUPP / "Additional_file_4_AF_precision_check_figure.png"
    supp_eps = UPLOAD_SUPP / "_Additional_file_4_AF_precision_check_figure.eps"
    build_single_vector_eps(
        supp_eps,
        "T90__incident_atrial_fibrillation__per10pct_harmonized_pvalue_check",
        "T90 and incident AF",
        "Exploratory precision-check only; alternative SE derivation",
    )
    render_eps_to_outputs(supp_eps, supp_pdf, supp_png)
    supp_eps.unlink(missing_ok=True)


def export_protocol_search_appendix() -> None:
    md_path = UPLOAD_SUPP / "Additional_file_1_Protocol_and_search_appendix.md"
    docx_path = UPLOAD_SUPP / "Additional_file_1_Protocol_and_search_appendix.docx"
    build_protocol_search_appendix(md_path)
    run_pandoc(md_path, docx_path)
    patch_docx(
        docx_path,
        repeat_table_headers=True,
        narrow_margins=True,
        table_widths=[
            [1800, 7600],
            [1700, 1900, 1900, 3000],
            [1450, 700, 2950, 1700, 3200],
        ],
        line_spacing_twips=220,
        table_line_spacing_twips=200,
        table_font_size_half_points=19,
    )


def export_manuscript_and_letter() -> None:
    man_md = ROOT / "_submission_manuscript_with_tables.md"
    man_docx = RENDERED / "manuscript_rr.docx"
    man_rtf = RENDERED / "manuscript_rr.rtf"
    cover_md = ROOT / "cover_letter.md"
    cover_docx = RENDERED / "cover_letter_rr.docx"
    cover_rtf = RENDERED / "cover_letter_rr.rtf"
    write_markdown(man_md, build_submission_manuscript_markdown())
    run_pandoc(man_md, man_docx)
    patch_docx(
        man_docx,
        document_role="manuscript",
        add_line_numbers=True,
        add_page_numbers=True,
        landscape=True,
        repeat_table_headers=True,
        narrow_margins=True,
        paragraph_before_twips=0,
        paragraph_after_twips=0,
        font_name="Times New Roman",
        font_size_half_points=24,
        table_widths=[
            [1750, 900, 850, 2400, 3950, 1700, 1350],
            [1100, 1500, 1200, 2400, 2200, 900],
            [2500, 1700, 2100, 900, 3000],
        ],
        table_line_spacing_twips=220,
        table_font_size_half_points=18,
    )
    run_pandoc(man_md, man_rtf)
    patch_rtf(man_rtf)
    rebuild_docx_from_rtf_via_soffice(man_rtf, man_docx)
    patch_docx(
        man_docx,
        document_role="manuscript",
        add_line_numbers=True,
        add_page_numbers=True,
        landscape=True,
        repeat_table_headers=True,
        narrow_margins=True,
        paragraph_before_twips=0,
        paragraph_after_twips=0,
        font_name="Times New Roman",
        font_size_half_points=24,
        table_widths=[
            [1750, 900, 850, 2400, 3950, 1700, 1350],
            [1100, 1500, 1200, 2400, 2200, 900],
            [2500, 1700, 2100, 900, 3000],
        ],
        table_line_spacing_twips=220,
        table_font_size_half_points=18,
    )
    run_pandoc(cover_md, cover_docx)
    patch_docx(
        cover_docx,
        document_role="cover_letter",
        line_spacing_twips=340,
        paragraph_before_twips=0,
        paragraph_after_twips=90,
        font_name="Times New Roman",
        font_size_half_points=22,
        narrow_margins=True,
    )
    run_pandoc(cover_md, cover_rtf)
    patch_rtf(cover_rtf)
    man_md.unlink(missing_ok=True)


def main() -> None:
    validate_source_inputs()
    for path in [RENDERED, UPLOAD_TABLES, UPLOAD_FIGURES, UPLOAD_SUPP]:
        clear_dir(path)
    export_manuscript_and_letter()
    export_tables()
    export_protocol_search_appendix()
    export_prisma_figure()
    export_primary_figure_panels()
    print("Export complete:")
    print(f"- rendered manuscript: {RENDERED / 'manuscript_rr.docx'}")
    print(f"- rendered manuscript: {RENDERED / 'manuscript_rr.rtf'}")
    print(f"- upload tables dir: {UPLOAD_TABLES}")
    print(f"- upload figures dir: {UPLOAD_FIGURES}")
    print(f"- upload supplement dir: {UPLOAD_SUPP}")


if __name__ == "__main__":
    main()
