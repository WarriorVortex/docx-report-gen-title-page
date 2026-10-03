"""Pure rendering of a title page from a TitlePageConfig.

The layout algorithm:

1. Collect the text lines that will be rendered, with their font
   sizes and a gap weight attached to each line (the space above it).

2. Estimate the number of visual lines each field will take. Word
   wraps long strings automatically; the estimate uses average
   character width for the configured font and the content width
   of the page.

3. Compute total text height as (estimated lines × line_height_pt).
   Compute the page's usable height from geometry.

4. The remaining vertical space is distributed among the gap weights.
   Every gap gets (weight / total_weight) × remaining points of
   space_before. If content overflows, gaps collapse to zero; the
   page then overflows visually — reduce font sizes in that case.

This gives an exact fit: the sum of text heights and gaps equals the
page's content height, regardless of margins or fonts.
"""
from dataclasses import dataclass
from typing import List, Optional

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from .config import TitlePageConfig


# Fallbacks for page geometry when python-docx reports None. A4 with
# standard academic margins.
_FALLBACK_PAGE_HEIGHT = Cm(29.7)
_FALLBACK_PAGE_WIDTH = Cm(21.0)
_FALLBACK_TOP_MARGIN = Cm(2.0)
_FALLBACK_BOTTOM_MARGIN = Cm(2.0)
_FALLBACK_LEFT_MARGIN = Cm(3.0)
_FALLBACK_RIGHT_MARGIN = Cm(1.0)


@dataclass
class _Line:
    """One text line with its gap weight and formatting."""
    text: str
    size: int
    bold: bool = False
    right: bool = False
    gap_weight: float = 0.0


def render_config(doc: DocxDocument, cfg: TitlePageConfig) -> None:
    """Render a title page into `doc`, fitting exactly one page.

    The layout is computed from page geometry and the config's gap
    weights, then applied with exact point-based spacing. Empty
    fields are skipped without leaving a placeholder.
    """
    lines = _build_lines(cfg)
    if not lines:
        return

    # The first line has no gap before it — content starts at the top
    # margin. Any weight accumulated before the first present line is
    # dropped.
    lines[0].gap_weight = 0.0

    geometry = _read_geometry(doc)
    line_height = _resolve_line_height(cfg)

    text_height_pt = sum(
        _estimate_lines(line.text, line.size, geometry.content_width_pt)
        * line_height
        for line in lines
    )

    total_weight = sum(line.gap_weight for line in lines)
    remaining = geometry.content_height_pt - text_height_pt
    if remaining < 0:
        remaining = 0.0
    gap_per_weight = remaining / total_weight if total_weight else 0.0

    for i, line in enumerate(lines):
        gap_pt = 0.0 if i == 0 else line.gap_weight * gap_per_weight
        _add_line(doc, cfg, line, gap_pt, line_height)


# ---------- geometry ----------

@dataclass
class _Geometry:
    content_width_pt: float
    content_height_pt: float


def _read_geometry(doc: DocxDocument) -> _Geometry:
    """Return the usable content area of the first section, in points."""
    section = doc.sections[0]

    page_h = section.page_height or _FALLBACK_PAGE_HEIGHT
    page_w = section.page_width or _FALLBACK_PAGE_WIDTH
    top = section.top_margin or _FALLBACK_TOP_MARGIN
    bottom = section.bottom_margin or _FALLBACK_BOTTOM_MARGIN
    left = section.left_margin or _FALLBACK_LEFT_MARGIN
    right = section.right_margin or _FALLBACK_RIGHT_MARGIN

    return _Geometry(
        content_width_pt=float(page_w.pt - left.pt - right.pt),
        content_height_pt=float(page_h.pt - top.pt - bottom.pt),
    )


def _resolve_line_height(cfg: TitlePageConfig) -> float:
    """Return the exact line height in points."""
    max_size = max(cfg.base_size, cfg.title_size, cfg.institution_size)
    if cfg.line_height_pt is not None:
        # Never allow below the safe minimum for the largest font.
        return float(max(cfg.line_height_pt, int(max_size * 1.2)))
    return float(int(max_size * 1.4))


def _estimate_lines(text: str, size_pt: float, width_pt: float) -> int:
    """Estimate how many visual lines `text` occupies.

    Uses an average character width of 0.5 × font size — a reasonable
    approximation for Times New Roman and similar serif fonts at
    normal kerning. Word wraps on word boundaries, so words are packed
    greedily into lines of at most `chars_per_line` characters.

    The estimate is intentionally conservative: slight overestimation
    keeps the title page from overflowing; a small amount of unused
    space at the bottom is preferable.
    """
    if not text:
        return 0

    char_width_pt = size_pt * 0.5
    chars_per_line = max(1, int(width_pt / char_width_pt))

    words = text.split()
    if not words:
        return 1

    lines = 1
    current = 0
    for word in words:
        length = len(word)
        if current == 0:
            current = length
        elif current + 1 + length <= chars_per_line:
            current += 1 + length
        else:
            lines += 1
            current = length
    return lines


# ---------- line collection ----------

def _build_lines(cfg: TitlePageConfig) -> List[_Line]:
    """Turn a config into an ordered list of lines with gap weights.

    Gap weights from skipped sections accumulate and are applied to
    the next present line, so no space is wasted on absent blocks.
    """
    lines: List[_Line] = []
    pending_gap = 0.0

    def add(text: str, size: int, *,
            bold: bool = False, right: bool = False) -> None:
        nonlocal pending_gap
        lines.append(_Line(
            text=text, size=size, bold=bold, right=right,
            gap_weight=pending_gap,
        ))
        pending_gap = 0.0

    # Header block
    if cfg.institution:
        add(cfg.institution, cfg.institution_size)
    if cfg.institute:
        add(cfg.institute, cfg.institution_size)
    if cfg.school:
        add(cfg.school, cfg.institution_size, bold=True)

    pending_gap += cfg.gap_after_header

    # Main block
    work_line = _join(cfg.work_type, cfg.work_number, sep=' ')
    if work_line:
        add(work_line, cfg.title_size, bold=True)

    pending_gap += cfg.gap_after_work

    if cfg.work_title:
        add(cfg.work_title, cfg.title_size, bold=True)

    pending_gap += cfg.gap_after_title

    if cfg.discipline:
        add(cfg.discipline_prefix or '', cfg.base_size)
        pending_gap += cfg.gap_after_discipline_prefix
        add(cfg.discipline, cfg.base_size)

    pending_gap += cfg.gap_after_discipline

    if cfg.student:
        add(cfg.student_label or '', cfg.base_size, right=True)
        pending_gap += cfg.gap_after_student_label
        add(cfg.student, cfg.base_size, right=True)

    pending_gap += cfg.gap_after_student

    if cfg.supervisor:
        supervisor = _join(
            cfg.supervisor_label, cfg.supervisor, sep=' ',
        )
        if supervisor:
            add(supervisor, cfg.base_size, right=True)

    pending_gap += cfg.gap_after_supervisor

    if cfg.city:
        add(cfg.city, cfg.base_size)
    if cfg.year is not None:
        add(str(cfg.year), cfg.base_size)

    return lines


# ---------- paragraph output ----------

def _add_line(
    doc: DocxDocument,
    cfg: TitlePageConfig,
    line: _Line,
    gap_pt: float,
    line_height_pt: float,
) -> None:
    """Add one rendered line to `doc`."""
    if not line.text:
        return

    para = doc.add_paragraph()
    para.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT if line.right
        else WD_ALIGN_PARAGRAPH.CENTER
    )
    pf = para.paragraph_format
    pf.space_before = Pt(gap_pt)
    pf.space_after = Pt(0)
    # Exact line height in points — Word uses this verbatim,
    # independent of the font's own metrics.
    pf.line_spacing = Pt(line_height_pt)

    run = para.add_run(line.text)
    run.font.name = cfg.font
    run.font.size = Pt(line.size)
    run.bold = line.bold


def _join(
    left: Optional[str],
    right: Optional[str],
    *,
    sep: str = ' ',
) -> Optional[str]:
    if left and right:
        return f'{left}{sep}{right}'
    return left or right or None