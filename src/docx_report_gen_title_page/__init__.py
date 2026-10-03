"""Title page plugin for docx-report-gen.

Inserts a standardized title page as the first element of a report.
The plugin is attached to a Report via the standard plugin mechanism:

    from docx_report_gen import Report
    from docx_report_gen_title_page import TitlePagePlugin

    r = Report(plugins=[TitlePagePlugin(
        institution='...',
        work_title='...',
        student='...',
        group='...',
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


try:
    __version__ = version("docx-report-gen-title-page")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"


# Public API. The plugin class and its configuration dataclass will
# be defined in dedicated modules and re-exported here.
__all__: list[str] = []


# TODO: implementation phase.
#   from .plugin import TitlePagePlugin
#   from .config import TitlePageConfig
#   __all__ = ['TitlePagePlugin', 'TitlePageConfig']