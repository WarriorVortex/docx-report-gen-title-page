"""Render a title page from a TitlePageConfig.

Layout algorithm:

1. Collect the lines that will be rendered, each with a font size
   and a gap weight (space above it).
2. Estimate how many visual lines each text takes (word wrapping by
   average character width).
3. Compute total text height = sum(lines × line_height). When
   line_height_pt is None, use single spacing — approximate 1.15 ×
   font size for Times New Roman.
4. Distribute the remaining vertical space of the content area among
   the gap weights.
5. If content overflows, gaps collapse to zero.

The last line's bottom edge lands exactly on the bottom margin, so
the year ends up on the last line of the page.

Sections: if custom margins are configured or isolate=True, the title
page is rendered into its own section. A new section is created
afterwards with the original margins restored, so the rest of the
document is unaffected.
"""
from dataclasses import dataclass
from typing import Optional

from docx.document import Document as DocxDocument
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.section import Section
from docx.shared import Cm, Length, Pt

from .config import TitlePageConfig


_FALLBACK_PAGE_HEIGHT = Cm(29.7)
_FALLBACK_PAGE_WIDTH = Cm(21.0)
_FALLBACK_TOP_MARGIN = Cm(2.0)
_FALLBACK_BOTTOM_MARGIN = Cm(2.0)
_FALLBACK_LEFT_MARGIN = Cm(3.0)
_FALLBACK_RIGHT_MARGIN = Cm(1.0)

_ALIGN_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
}

# Approximate ratio of line height to font size for single spacing in
# common serif fonts. Used only to estimate total text height for
# distributing the gap weights.
_SINGLE_LINE_RATIO = 1.15


@dataclass
class _Line:
    """One rendered line."""
    text: str
    size: int
    bold: bool = False
    align: str = 'center'
    gap_weight: float = 0.0
    tab_stop: bool = False


def render_config(doc: DocxDocument, cfg: TitlePageConfig) -> None:
    """Render the title page into `doc`.

    The title page is rendered into its own section when custom
    margins are set, or when cfg.isolate is True. Otherwise it is
    rendered inline; the caller is responsible for the page break.
    """
    needs_section = _needs_section(cfg)

    if not needs_section:
        _render_lines(doc, cfg)
        return

    first = doc.sections[0]
    saved_top = first.top_margin
    saved_bottom = first.bottom_margin
    saved_left = first.left_margin
    saved_right = first.right_margin

    # Apply the title page's margins to the first section.
    _apply_margins(first, cfg)

    # Render the title page into the first section.
    _render_lines(doc, cfg)

    if cfg.isolate:
        _clear_header_footer(first)

    # Start a new section for the rest of the document, restoring the
    # original margins so subsequent content is unaffected.
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    _restore_margins(
        new_section,
        saved_top, saved_bottom, saved_left, saved_right,
    )


def append_page_break_inline(doc: DocxDocument) -> None:
    """Append a page break to the last paragraph.

    Unlike Document.add_page_break (which creates a new paragraph
    containing only a break), this inserts the break inside the
    existing last paragraph. No extra blank line appears at the
    bottom of the title page.
    """
    if not doc.paragraphs:
        return
    last = doc.paragraphs[-1]
    run = last.add_run()
    run.add_break(WD_BREAK.PAGE)


# ---------- sections and margins ----------

def _needs_section(cfg: TitlePageConfig) -> bool:
    """Return True if the title page should be its own section."""
    if cfg.isolate:
        return True
    return any(x is not None for x in (
        cfg.margin_top, cfg.margin_bottom,
        cfg.margin_left, cfg.margin_right,
    ))


def _apply_margins(section: Section, cfg: TitlePageConfig) -> None:
    if cfg.margin_top is not None:
        section.top_margin = Cm(cfg.margin_top)
    if cfg.margin_bottom is not None:
        section.bottom_margin = Cm(cfg.margin_bottom)
    if cfg.margin_left is not None:
        section.left_margin = Cm(cfg.margin_left)
    if cfg.margin_right is not None:
        section.right_margin = Cm(cfg.margin_right)


def _restore_margins(
    section: Section,
    top: Optional[Length],
    bottom: Optional[Length],
    left: Optional[Length],
    right: Optional[Length],
) -> None:
    if top is not None:
        section.top_margin = top
    if bottom is not None:
        section.bottom_margin = bottom
    if left is not None:
        section.left_margin = left
    if right is not None:
        section.right_margin = right


def _clear_header_footer(section: Section) -> None:
    for part in (section.header, section.footer):
        part.is_linked_to_previous = False
        for para in part.paragraphs:
            para.text = ''


# ---------- rendering ----------

def _render_lines(doc: DocxDocument, cfg: TitlePageConfig) -> None:
    lines = _build_lines(cfg)
    if not lines:
        return

    lines[0].gap_weight = 0.0

    geometry = _read_geometry(doc)

    text_height_pt = sum(
        _estimate_lines(line.text, line.size, geometry.content_width_pt)
        * _line_height_pt(line.size, cfg)
        for line in lines
    )

    total_weight = sum(line.gap_weight for line in lines)
    remaining = geometry.content_height_pt - text_height_pt
    if remaining < 0:
        remaining = 0.0
    gap_per_weight = remaining / total_weight if total_weight else 0.0

    for i, line in enumerate(lines):
        gap_pt = 0.0 if i == 0 else line.gap_weight * gap_per_weight
        _add_line(doc, cfg, line, gap_pt, geometry.content_width_pt)


def _line_height_pt(size: int, cfg: TitlePageConfig) -> float:
    """Line height in points for estimation purposes."""
    if cfg.line_height_pt is not None:
        return float(cfg.line_height_pt)
    return size * _SINGLE_LINE_RATIO


# ---------- geometry ----------

@dataclass
class _Geometry:
    content_width_pt: float
    content_height_pt: float


def _read_geometry(doc: DocxDocument) -> _Geometry:
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


def _estimate_lines(text: str, size_pt: float, width_pt: float) -> int:
    """Estimate visual line count. Tabbed lines are one line."""
    if not text:
        return 0
    if '\t' in text:
        return 1

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

def _build_lines(cfg: TitlePageConfig) -> list[_Line]:
    lines: list[_Line] = []
    pending_gap = 0.0

    def add(
        text: str,
        size: int,
        *,
        bold: bool = False,
        align: str = 'center',
        tab_stop: bool = False,
    ) -> None:
        nonlocal pending_gap
        if not text:
            return
        lines.append(_Line(
            text=text, size=size, bold=bold, align=align,
            tab_stop=tab_stop, gap_weight=pending_gap,
        ))
        pending_gap = 0.0

    # Header
    add(cfg.institution or '', cfg.institution_size)
    add(cfg.institute or '', cfg.institution_size)
    add(cfg.school or '', cfg.institution_size, bold=True)
    pending_gap += cfg.gap_after_header

    # Main
    work_line = _join(cfg.work_type, cfg.work_number, sep=' ')
    add(work_line or '', cfg.title_size, bold=True)
    pending_gap += cfg.gap_after_work

    add(cfg.work_title or '', cfg.title_size, bold=True)
    pending_gap += cfg.gap_after_title

    if cfg.discipline:
        if cfg.discipline_prefix:
            add(cfg.discipline_prefix, cfg.base_size)
            pending_gap += cfg.gap_after_discipline_prefix
        add(cfg.discipline, cfg.base_size)
    pending_gap += cfg.gap_after_discipline

    # Students — students_label is fully replaceable.
    if cfg.students:
        if cfg.students_label:
            add(cfg.students_label, cfg.base_size, align='left')
            pending_gap += cfg.gap_after_students_label
        for student in cfg.students:
            add(student.render(), cfg.base_size,
                align='left', tab_stop=True)
    pending_gap += cfg.gap_after_students

    # Supervisors — same as students_label.
    if cfg.supervisors:
        if cfg.supervisors_label:
            add(cfg.supervisors_label, cfg.base_size, align='left')
            pending_gap += cfg.gap_after_supervisors_label
        for supervisor in cfg.supervisors:
            add(supervisor.render(), cfg.base_size,
                align='left', tab_stop=True)
    pending_gap += cfg.gap_after_supervisors

    # Bottom
    add(cfg.city or '', cfg.base_size)
    if cfg.year is not None:
        add(str(cfg.year), cfg.base_size)

    return lines


# ---------- paragraph output ----------

def _add_line(
    doc: DocxDocument,
    cfg: TitlePageConfig,
    line: _Line,
    gap_pt: float,
    content_width_pt: float,
) -> None:
    if not line.text:
        return

    para = doc.add_paragraph()
    para.alignment = _ALIGN_MAP[line.align]

    pf = para.paragraph_format
    pf.space_before = Pt(gap_pt)
    pf.space_after = Pt(0)

    # Single spacing unless explicitly overridden.
    if cfg.line_height_pt is not None:
        pf.line_spacing = Pt(cfg.line_height_pt)
    else:
        pf.line_spacing = 1.0

    if line.tab_stop:
        pf.tab_stops.add_tab_stop(
            Pt(content_width_pt), WD_TAB_ALIGNMENT.RIGHT,
        )

    parts = line.text.split('\t')
    for i, part in enumerate(parts):
        if i > 0:
            para.add_run('\t')
        run = para.add_run(part)
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