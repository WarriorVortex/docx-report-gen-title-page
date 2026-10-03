"""Pure rendering of a title page from a TitlePageConfig.

Kept separate from plugin.py so the rendering logic is testable in
isolation and has no dependency on the plugin lifecycle.
"""
from typing import Optional

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from .config import TitlePageConfig


def render_config(doc: DocxDocument, cfg: TitlePageConfig) -> None:
    """Render a title page into `doc` using explicit formatting.

    Skips any line whose value is missing. Labels and section headers
    are only rendered together with their value — an empty label does
    not appear on the page.
    """
    # ---------- header block ----------
    _add_line(doc, cfg, cfg.institution, cfg.institution_size)
    _add_line(doc, cfg, cfg.institute, cfg.institution_size)
    _add_line(doc, cfg, cfg.school, cfg.institution_size, bold=True)

    _blank(doc, cfg, 2)

    # ---------- main block ----------
    work_line = _join(cfg.work_type, cfg.work_number, sep=' ')
    _add_line(doc, cfg, work_line, cfg.title_size, bold=True)

    _blank(doc, cfg, 1)

    _add_line(doc, cfg, cfg.work_title, cfg.title_size, bold=True)

    _blank(doc, cfg, 1)

    if cfg.discipline:
        _add_line(doc, cfg, cfg.discipline_prefix, cfg.base_size)
        _blank(doc, cfg, 1)
        _add_line(doc, cfg, cfg.discipline, cfg.base_size)

    _blank(doc, cfg, 4)

    # ---------- signature block ----------
    if cfg.student:
        _add_line(
            doc, cfg, cfg.student_label, cfg.base_size, right=True,
        )
        _blank(doc, cfg, 1)
        _add_line(doc, cfg, cfg.student, cfg.base_size, right=True)

    if cfg.supervisor:
        supervisor_line = _join(
            cfg.supervisor_label, cfg.supervisor, sep=' ',
        )
        _blank(doc, cfg, 1)
        _add_line(
            doc, cfg, supervisor_line, cfg.base_size, right=True,
        )

    # ---------- bottom block ----------
    if cfg.city or cfg.year is not None:
        _blank(doc, cfg, cfg.bottom_gap_lines)
        _add_line(doc, cfg, cfg.city, cfg.base_size)
        _add_line(
            doc, cfg,
            str(cfg.year) if cfg.year is not None else None,
            cfg.base_size,
        )


# ---------- helpers ----------

def _add_line(
    doc: DocxDocument,
    cfg: TitlePageConfig,
    text: Optional[str],
    size: int,
    *,
    bold: bool = False,
    right: bool = False,
) -> None:
    if not text:
        return

    para = doc.add_paragraph()
    para.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT if right
        else WD_ALIGN_PARAGRAPH.CENTER
    )
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = cfg.line_spacing

    run = para.add_run(text)
    run.font.name = cfg.font
    run.font.size = Pt(size)
    run.bold = bold


def _blank(doc: DocxDocument, cfg: TitlePageConfig, count: int) -> None:
    for _ in range(count):
        para = doc.add_paragraph()
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = cfg.line_spacing
        run = para.add_run('')
        run.font.size = Pt(cfg.base_size)


def _join(
    left: Optional[str],
    right: Optional[str],
    *,
    sep: str = ' ',
) -> Optional[str]:
    if left and right:
        return f'{left}{sep}{right}'
    return left or right or None