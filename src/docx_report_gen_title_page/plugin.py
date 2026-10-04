"""TitlePagePlugin — adds a title_page() method to a Report.

The plugin renders a title page from a TitlePageConfig. Loading an
existing .docx as the base document is a feature of Report itself —
use Report(source=...) instead of a plugin argument.

Page break behavior: when the title page uses its own section (custom
margins or isolate=True), the section break itself starts a new page.
Otherwise, an inline break is appended to the last paragraph — no
extra blank line at the bottom of the title page.
"""
from typing import Any, Optional

from docx_report_gen import Plugin, Report

from ._render import (
    _needs_section,
    append_page_break_inline,
    render_config,
)
from .config import TitlePageConfig


class TitlePagePlugin(Plugin):
    """Adds title_page() to Report.

    Example:
        from docx_report_gen import Report
        from docx_report_gen_title_page import (
            TitlePagePlugin, Student, Supervisor,
        )

        r = Report(plugins=[TitlePagePlugin()])
        r.title_page(
            institution='СПбПУ',
            work_type='ЛАБОРАТОРНАЯ РАБОТА',
            work_number='№2',
            work_title='«...»',
            students=[Student('Ерохин В.С.', group='5130904/30103')],
            supervisors=[Supervisor('Тутыгин В.С.', prefix='доцент')],
            city='Санкт-Петербург',
            year=2026,
        )
        r.h1('Введение')

    Replaceable labels and custom margins:

        r.title_page(
            students_label='Работу выполнил:',
            supervisors_label='Научный руководитель:',
            margin_top=2.5, margin_bottom=2.5,
            margin_left=3.0, margin_right=1.5,
            ...
        )

    Combining with a source-based base document:

        r = Report(source='Титульный лист.docx',
                   plugins=[TitlePagePlugin()])
        r.h1('Введение')
    """

    name = 'title-page'

    def __init__(
        self,
        config: Optional[TitlePageConfig] = None,
        page_break: bool = True,
        isolate: bool = False,
        **kwargs: Any,
    ) -> None:
        super().__init__()

        if config is not None and kwargs:
            raise ValueError(
                'Pass either config=... or **kwargs, not both'
            )

        self.page_break = page_break
        self.isolate = isolate
        self._initial_config: Optional[TitlePageConfig] = None

        if config is not None:
            self._initial_config = config
        elif kwargs:
            self._initial_config = TitlePageConfig(**kwargs)

    # ---------- Plugin hook ----------

    def setup(self, report: Report) -> None:
        registry = self.registry
        assert registry is not None, 'registry must be set during setup'

        registry.block('title_page', self._title_page_handler)

        if self._initial_config is not None:
            self._title_page_handler(report, self._initial_config)

    # ---------- handler ----------

    def _title_page_handler(
        self,
        report: Report,
        config: Optional[TitlePageConfig] = None,
        *,
        page_break: Optional[bool] = None,
        isolate: Optional[bool] = None,
        **kwargs: Any,
    ) -> Report:
        if config is not None and kwargs:
            raise ValueError(
                'title_page(): pass either a config object or keyword '
                'fields, not both'
            )
        cfg = config if config is not None else TitlePageConfig(**kwargs)

        if isolate is not None:
            cfg = TitlePageConfig(
                **{**cfg.__dict__, 'isolate': isolate},
            )

        render_config(report.doc, cfg)

        # A section break (if any) already starts a new page. Only add
        # an explicit page break when no section was created.
        if not _needs_section(cfg):
            pb = self.page_break if page_break is None else page_break
            if pb:
                append_page_break_inline(report.doc)
        return report