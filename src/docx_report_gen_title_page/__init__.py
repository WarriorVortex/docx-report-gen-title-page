"""Title page plugin for docx-report-gen.

Adds a `title_page()` method to a Report, rendering a formatted
title page from fields. Optionally isolates the page in its own
section (empty headers/footers), or adds a page break after it.

To use an existing .docx as the base document, use the base package:

    from docx_report_gen import Report
    r = Report(source='Титульный лист.docx')
    r.h1('Введение')

Combine with the plugin:

    from docx_report_gen import Report
    from docx_report_gen_title_page import TitlePagePlugin

    r = Report(source='шаблон.docx', plugins=[TitlePagePlugin()])
    r.h1('Введение')
"""
from importlib.metadata import version, PackageNotFoundError

from .config import Student, Supervisor, TitlePageConfig
from .plugin import TitlePagePlugin


try:
    __version__ = version("docx-report-gen-title-page")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"


__all__ = [
    'TitlePagePlugin',
    'TitlePageConfig',
    'Student',
    'Supervisor',
]