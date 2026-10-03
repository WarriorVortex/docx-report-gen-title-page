"""TitlePagePlugin — adds title-page methods to a Report.

The plugin registers two block methods on every Report it is attached
to:

    report.title_page(config=None, **kwargs)
        Build a title page from fields. Accepts either a fully
        constructed TitlePageConfig or keyword fields.

    report.load_title_page(source, page_break=None)
        Load an existing .docx file and append its content as-is,
        preserving styles, sizes and spacing. Missing style
        definitions are copied from the source; existing style IDs
        in the target are kept as-is.

Both methods return the report, so they chain with other block
methods.

Auto-render in setup():
    If the plugin is constructed with `config=`, `source=`, or content
    keyword fields, the page is rendered immediately in `setup()` —
    before any user content. Constructing with no arguments registers
    the methods and waits for an explicit call.
"""
from pathlib import Path
from typing import Any, Optional, Union

from docx_report_gen import Plugin, Report

from ._render import render_config
from .config import TitlePageConfig
from .loader import load_docx_content


PathLike = Union[str, Path]


class TitlePagePlugin(Plugin):
    """Adds title_page() and load_title_page() methods to Report.

    Explicit usage:

        r = Report(plugins=[TitlePagePlugin()])
        r.title_page(
            institution='СПбПУ',
            work_title='«...»',
            student='студент гр. ... Ерохин В.С.',
            supervisor='Тутыгин В.С.',
            city='Санкт-Петербург',
            year=2026,
        )
        r.h1('Введение')

    Or:

        r.load_title_page('Титульный лист.docx')

    Auto-render:

        r = Report(plugins=[
            TitlePagePlugin(institution='СПбПУ', year=2026),
        ])
        # Title page is already on the first sheet.
    """

    name = 'title-page'

    def __init__(
            self,
            config: Optional[TitlePageConfig] = None,
            source: Optional[PathLike] = None,
            page_break: bool = True,
            **kwargs: Any,
    ) -> None:
        """Initialize the plugin.

        Args:
            config: a fully constructed TitlePageConfig for auto-render.
            source: path to an existing .docx title page for auto-render.
                Mutually exclusive with `config` and content kwargs.
            page_break: default for the page break after the title page.
                Used by auto-render and by both methods when they are
                called without an explicit override.
            **kwargs: fields for TitlePageConfig, used when `config`
                and `source` are both absent.

        Raises:
            ValueError: config and kwargs were both given; or source
                and (config or kwargs) were both given.
            FileNotFoundError: source path does not exist.
            TypeError: an unknown content keyword was passed.
        """
        super().__init__()

        if config is not None and kwargs:
            raise ValueError(
                'Pass either config=... or **kwargs, not both'
            )

        if source is not None and (config is not None or kwargs):
            raise ValueError(
                'Pass either source=..., or config=.../**kwargs, not both'
            )

        self.page_break = page_break
        self._initial_source: Optional[Path] = None
        self._initial_config: Optional[TitlePageConfig] = None

        if source is not None:
            path = Path(source)
            if not path.exists():
                raise FileNotFoundError(
                    f'TitlePagePlugin(): source not found: {path}'
                )
            self._initial_source = path
        elif config is not None or kwargs:
            self._initial_config = (
                config if config is not None else TitlePageConfig(**kwargs)
            )

    # ---------- Plugin hook ----------

    def setup(self, report: Report) -> None:
        """Register block methods; auto-render if configured."""
        # Plugin.setup() receives a report with a registry attached;
        # registry is Optional on Plugin only to allow a plugin to be
        # constructed without one. Here it is always set.
        registry = self.registry
        assert registry is not None, 'registry must be set during setup'

        registry.block('title_page', self._title_page_handler)
        registry.block('load_title_page', self._load_title_page_handler)

        if self._initial_source is not None:
            self._load_title_page_handler(
                report, self._initial_source, page_break=True,
            )
        elif self._initial_config is not None:
            self._title_page_handler(report, self._initial_config)

    # ---------- block method handlers ----------

    def _title_page_handler(
        self,
        report: Report,
        config: Optional[TitlePageConfig] = None,
        *,
        page_break: Optional[bool] = None,
        **kwargs: Any,
    ) -> Report:
        """Handler for Report.title_page(...).

        Returns the report so calls can chain.
        """
        if config is not None and kwargs:
            raise ValueError(
                'title_page(): pass either a config object or keyword '
                'fields, not both'
            )
        cfg = (
            config if config is not None
            else TitlePageConfig(**kwargs)
        )

        render_config(report.doc, cfg)

        pb = self.page_break if page_break is None else page_break
        if pb:
            report.page_break()
        return report

    def _load_title_page_handler(
        self,
        report: Report,
        source: PathLike,
        page_break: Optional[bool] = None,
    ) -> Report:
        """Handler for Report.load_title_page(...).

        Returns the report so calls can chain.
        """
        load_docx_content(source, report.doc)

        pb = self.page_break if page_break is None else page_break
        if pb:
            report.page_break()
        return report