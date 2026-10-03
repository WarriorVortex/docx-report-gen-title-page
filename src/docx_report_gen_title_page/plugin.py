"""TitlePagePlugin — inserts a standardized title page into a Report.

The plugin runs in its `setup(report)` hook, which fires during
Report.__init__, before any user content. It writes the title page
using python-docx primitives with explicit formatting, and closes the
page with a page break so subsequent content starts on page 2.

All content comes from a TitlePageConfig. Missing fields are skipped
without rendering an empty line — labels and section headers are only
rendered together with their value, so a minimal config still
produces a coherent page.
"""
from typing import Any, Optional

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from docx_report_gen import Plugin, Report

from .config import TitlePageConfig


class TitlePagePlugin(Plugin):
    """Renders a title page as the first element of a Report.

    Example:
        from docx_report_gen import Report
        from docx_report_gen_title_page import TitlePagePlugin

        r = Report(plugins=[
            TitlePagePlugin(
                institution='Санкт-Петербургский политехнический '
                            'университет Петра Великого',
                institute='Институт компьютерных наук и технологий',
                school='Высшая школа программной инженерии',
                work_type='ЛАБОРАТОРНАЯ РАБОТА',
                work_number='№2',
                work_title='«Разложение дискретизированных сигналов '
                           'в действительный и комплексный ряд Фурье»',
                discipline='«Применение методов искусственного '
                           'интеллекта для цифровой обработки сигналов»',
                student='студент гр. 5130904/30103 Ерохин В.С.',
                supervisor='Тутыгин В.С.',
                city='Санкт-Петербург',
                year=2026,
            ),
        ])
        r.h1('Introduction')
        r.save('report.docx')
    """

    name = 'title-page'

    def __init__(
        self,
        config: Optional[TitlePageConfig] = None,
        **kwargs: Any,
    ) -> None:
        """Initialize the plugin.

        Args:
            config: a fully constructed TitlePageConfig. If given,
                **kwargs must be empty.
            **kwargs: fields for TitlePageConfig, used when `config` is
                None. Passing both is an error.

        Raises:
            ValueError: both `config` and keyword fields were provided.
            TypeError: an unknown keyword field was passed.
        """
        super().__init__()
        if config is not None and kwargs:
            raise ValueError(
                'Pass either config=... or **kwargs, not both'
            )
        self.config = config if config is not None else TitlePageConfig(**kwargs)

    # ---------- Plugin hook ----------

    def setup(self, report: Report) -> None:
        """Insert the title page into a freshly created Report.

        Uses only the public API of Report: python-docx's own
        Document.add_page_break is not typed, so it is reached through
        Report.page_break() instead. This also keeps the plugin free
        of direct python-docx coupling beyond paragraph formatting.
        """
        self._render(report.doc)
        report.page_break()

    # ---------- rendering ----------

    def _render(self, doc: DocxDocument) -> None:
        cfg = self.config

        # ---------- header block ----------
        self._add_line(doc, cfg.institution, cfg.institution_size)
        self._add_line(doc, cfg.institute, cfg.institution_size)
        self._add_line(doc, cfg.school, cfg.institution_size, bold=True)

        self._blank(doc, 2)

        # ---------- main block ----------
        work_line = self._join(cfg.work_type, cfg.work_number, sep=' ')
        self._add_line(doc, work_line, cfg.title_size, bold=True)

        self._blank(doc, 1)

        self._add_line(doc, cfg.work_title, cfg.title_size, bold=True)

        self._blank(doc, 1)

        # Discipline prefix is a label — render only with the discipline.
        if cfg.discipline:
            self._add_line(doc, cfg.discipline_prefix, cfg.base_size)

            self._blank(doc, 1)

            self._add_line(doc, cfg.discipline, cfg.base_size)

        self._blank(doc, 4)

        # ---------- signature block ----------
        if cfg.student:
            self._add_line(
                doc, cfg.student_label, cfg.base_size, right=True,
            )
            self._blank(doc, 1)
            self._add_line(doc, cfg.student, cfg.base_size, right=True)

        if cfg.supervisor:
            supervisor_line = self._join(
                cfg.supervisor_label, cfg.supervisor, sep=' ',
            )
            self._blank(doc, 1)
            self._add_line(
                doc, supervisor_line, cfg.base_size, right=True,
            )

        # ---------- bottom block ----------
        if cfg.city or cfg.year is not None:
            self._blank(doc, cfg.bottom_gap_lines)
            self._add_line(doc, cfg.city, cfg.base_size)
            self._add_line(
                doc,
                str(cfg.year) if cfg.year is not None else None,
                cfg.base_size,
            )

    # ---------- helpers ----------

    def _add_line(
        self,
        doc: DocxDocument,
        text: Optional[str],
        size: int,
        *,
        bold: bool = False,
        right: bool = False,
    ) -> None:
        """Add a centered or right-aligned line.

        Empty and None text is skipped without producing a paragraph.
        """
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
        pf.line_spacing = self.config.line_spacing

        run = para.add_run(text)
        run.font.name = self.config.font
        run.font.size = Pt(size)
        run.bold = bold

    def _blank(self, doc: DocxDocument, count: int) -> None:
        """Add `count` empty paragraphs at base font size.

        Each paragraph carries an empty run with an explicit font size
        so Word computes the paragraph height from the intended size,
        not from the surrounding style.
        """
        for _ in range(count):
            para = doc.add_paragraph()
            pf = para.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = self.config.line_spacing
            run = para.add_run('')
            run.font.size = Pt(self.config.base_size)

    @staticmethod
    def _join(
        left: Optional[str],
        right: Optional[str],
        *,
        sep: str = ' ',
    ) -> Optional[str]:
        """Join two optional strings, skipping missing parts."""
        if left and right:
            return f'{left}{sep}{right}'
        return left or right or None