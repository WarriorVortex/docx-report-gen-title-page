"""Tests for TitlePagePlugin — rendering and integration with Report."""
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_report_gen import Report
from docx_report_gen_title_page import TitlePageConfig, TitlePagePlugin


# ---------- construction ----------

def test_plugin_name():
    assert TitlePagePlugin.name == 'title-page'


def test_plugin_accepts_kwargs():
    p = TitlePagePlugin(institution='X', year=2026)
    assert p.config.institution == 'X'
    assert p.config.year == 2026


def test_plugin_accepts_config_object():
    cfg = TitlePageConfig(institution='Y')
    p = TitlePagePlugin(config=cfg)
    assert p.config is cfg


def test_plugin_rejects_config_and_kwargs_together():
    with pytest.raises(ValueError, match='not both'):
        TitlePagePlugin(config=TitlePageConfig(), institution='X')


def test_plugin_rejects_unknown_field():
    with pytest.raises(TypeError):
        TitlePagePlugin(nonexistent_field='X')


def test_default_config_has_all_none_content():
    p = TitlePagePlugin()
    assert p.config.institution is None
    assert p.config.work_type is None


# ---------- rendering ----------

def _texts(doc):
    return [p.text for p in doc.paragraphs]


def test_full_title_page(tmp_path):
    r = Report(plugins=[TitlePagePlugin(
        institution='СПбПУ',
        institute='ИКНТ',
        school='ВШПИ',
        work_type='ЛАБОРАТОРНАЯ РАБОТА',
        work_number='№2',
        work_title='«Разложение в ряд Фурье»',
        discipline='«Цифровая обработка сигналов»',
        student='студент гр. 101 Ерохин В.С.',
        supervisor='Тутыгин В.С.',
        city='Санкт-Петербург',
        year=2026,
    )])
    r.h1('Introduction')
    path = tmp_path / 'out.docx'
    r.save(str(path))

    doc = Document(str(path))
    texts = _texts(doc)

    assert 'СПбПУ' in texts
    assert 'ИКНТ' in texts
    assert 'ВШПИ' in texts
    assert 'ЛАБОРАТОРНАЯ РАБОТА №2' in texts
    assert '«Разложение в ряд Фурье»' in texts
    assert 'по дисциплине' in texts
    assert '«Цифровая обработка сигналов»' in texts
    assert 'Выполнил' in texts
    assert 'студент гр. 101 Ерохин В.С.' in texts
    assert 'Руководитель Тутыгин В.С.' in texts
    assert 'Санкт-Петербург' in texts
    assert '2026' in texts
    assert 'Introduction' in texts


def test_missing_fields_are_skipped(tmp_path):
    r = Report(plugins=[TitlePagePlugin(
        institution='X',
        city='Y',
    )])
    r.h1('Body')
    path = tmp_path / 'out.docx'
    r.save(str(path))

    doc = Document(str(path))
    texts = [t for t in _texts(doc) if t]
    # Only X, Y and Body should be non-empty
    assert 'X' in texts
    assert 'Y' in texts
    assert 'Body' in texts


def test_work_type_and_number_joined():
    p = TitlePagePlugin(work_type='ЛАБОРАТОРНАЯ', work_number='№3')
    r = Report(plugins=[p])
    doc = r.doc
    texts = _texts(doc)
    assert 'ЛАБОРАТОРНАЯ №3' in texts


def test_work_type_without_number():
    p = TitlePagePlugin(work_type='РЕФЕРАТ')
    r = Report(plugins=[p])
    texts = _texts(r.doc)
    assert 'РЕФЕРАТ' in texts


def test_supervisor_joined_into_single_line():
    p = TitlePagePlugin(supervisor_label='Руководитель',
                        supervisor='Иванов И.И.')
    r = Report(plugins=[p])
    texts = _texts(r.doc)
    assert 'Руководитель Иванов И.И.' in texts


def test_year_rendered_as_string():
    p = TitlePagePlugin(year=2026)
    r = Report(plugins=[p])
    texts = _texts(r.doc)
    assert '2026' in texts


def test_signature_block_is_right_aligned():
    p = TitlePagePlugin(student='студент')
    r = Report(plugins=[p])
    doc = r.doc
    for para in doc.paragraphs:
        if para.text == 'студент':
            assert para.alignment == WD_ALIGN_PARAGRAPH.RIGHT
            break
    else:
        pytest.fail('student line not found')


def test_header_lines_are_centered():
    p = TitlePagePlugin(institution='University')
    r = Report(plugins=[p])
    doc = r.doc
    for para in doc.paragraphs:
        if para.text == 'University':
            assert para.alignment == WD_ALIGN_PARAGRAPH.CENTER
            break
    else:
        pytest.fail('institution line not found')


def test_page_break_present_at_end(tmp_path):
    """After the title page, a page break must separate it from user content."""
    import zipfile
    from lxml import etree

    r = Report(plugins=[TitlePagePlugin(institution='X')])
    r.h1('Body')
    path = tmp_path / 'out.docx'
    r.save(str(path))

    with zipfile.ZipFile(path) as zf:
        xml = etree.fromstring(zf.read('word/document.xml'))
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    page_breaks = [
        br for br in xml.iter(f'{{{W}}}br')
        if br.get(f'{{{W}}}type') == 'page'
    ]
    assert page_breaks


def test_user_content_follows_title_page():
    """Content added after Report creation appears after the title page."""
    p = TitlePagePlugin(institution='University')
    r = Report(plugins=[p])
    r.h1('First heading')

    texts = [t for t in _texts(r.doc) if t]
    assert texts.index('University') < texts.index('First heading')


def test_plugin_works_with_register_plugin(tmp_path):
    from docx_report_gen.plugins import plugins as global_plugins

    plugin = TitlePagePlugin(institution='Global University')
    try:
        global_plugins.register(plugin)
        r = Report()
        texts = [t for t in _texts(r.doc) if t]
        assert 'Global University' in texts
    finally:
        global_plugins.unregister(plugin)


def test_plugin_setup_runs_on_attach():
    """setup() is called automatically when plugin is passed to Report."""
    p = TitlePagePlugin(institution='Direct')
    r = Report(plugins=[p])
    assert p.report is r
    assert p.registry is not None


def test_empty_config_produces_only_page_break():
    """With an empty config, only the page break should be added."""
    r = Report(plugins=[TitlePagePlugin()])
    non_empty = [t for t in _texts(r.doc) if t]
    assert non_empty == []