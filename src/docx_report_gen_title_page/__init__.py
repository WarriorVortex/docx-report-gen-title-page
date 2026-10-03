"""Title page plugin for docx-report-gen.

Adds two methods to a Report:

    report.title_page(config=None, **kwargs)
        Build a title page from fields.

    report.load_title_page(source, page_break=None)
        Load an existing .docx file as the title page.

Typical usage:

    from docx_report_gen import Report
    from docx_report_gen_title_page import TitlePagePlugin

    r = Report(plugins=[TitlePagePlugin()])

    r.title_page(
        institution='Санкт-Петербургский политехнический '
                    'университет Петра Великого',
        institute='Институт компьютерных наук и технологий',
        school='Высшая школа программной инженерии',
        work_type='ЛАБОРАТОРНАЯ РАБОТА',
        work_number='№2',
        work_title='«Разложение сигналов в ряд Фурье»',
        discipline='«Применение методов ИИ для ЦОС»',
        student='студент гр. 5130904/30103 Ерохин В.С.',
        supervisor='Тутыгин В.С.',
        city='Санкт-Петербург',
        year=2026,
    )
    r.h1('Введение')
    r.save('report.docx')

Or from an existing file:

    r = Report(plugins=[TitlePagePlugin()])
    r.load_title_page('Титульный лист.docx')
    r.h1('Введение')

Auto-render (no explicit call needed):

    r = Report(plugins=[TitlePagePlugin(source='Титульный лист.docx')])
    r.h1('Введение')

Or globally for every Report:

    from docx_report_gen.writer import register_plugin
    register_plugin(TitlePagePlugin())
"""
from importlib.metadata import version, PackageNotFoundError

from .config import TitlePageConfig
from .loader import load_docx_content
from .plugin import TitlePagePlugin


try:
    __version__ = version("docx-report-gen-title-page")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"


__all__ = ['TitlePagePlugin', 'TitlePageConfig', 'load_docx_content']