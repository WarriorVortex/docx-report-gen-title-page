"""Title page plugin for docx-report-gen.

Inserts a standardized title page as the first element of a report.
The plugin is attached to a Report via the standard plugin mechanism:

    from docx_report_gen import Report
    from docx_report_gen_title_page import TitlePagePlugin

    r = Report(plugins=[TitlePagePlugin(
        institution='Санкт-Петербургский политехнический '
                    'университет Петра Великого',
        work_title='«Разложение сигналов в ряд Фурье»',
        student='студент гр. 5130904/30103 Ерохин В.С.',
        supervisor='Тутыгин В.С.',
        city='Санкт-Петербург',
        year=2026,
    )])
    r.h1('Introduction')
    r.save('report.docx')

Or globally, so every Report picks it up:

    from docx_report_gen.writer import register_plugin
    register_plugin(TitlePagePlugin(...))

The plugin writes into the document in its `setup(report)` hook,
which runs during Report.__init__ — before any user content.
"""
from importlib.metadata import version, PackageNotFoundError

from .config import TitlePageConfig
from .plugin import TitlePagePlugin


try:
    __version__ = version("docx-report-gen-title-page")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"


__all__ = ['TitlePagePlugin', 'TitlePageConfig']