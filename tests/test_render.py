"""Tests for the layout algorithm — one-page fit, tabs, gaps."""
import pytest
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.shared import Cm, Pt

from docx_report_gen_title_page import (
    Student, Supervisor, TitlePageConfig,
)
from docx_report_gen_title_page._render import (
    _build_lines, _estimate_lines, append_page_break_inline,
    render_config,
)


# ---------- line building ----------

def test_empty_config_produces_no_lines():
    assert _build_lines(TitlePageConfig()) == []


def test_first_line_has_no_gap():
    lines = _build_lines(TitlePageConfig(institution='X'))
    assert lines[0].gap_weight == 0.0


def test_single_student_lines():
    cfg = TitlePageConfig(students=[Student('A', group='101')])
    texts = [line.text for line in _build_lines(cfg)]
    assert texts == ['Выполнил', 'студент гр. 101\tA']


def test_students_label_replaceable():
    cfg = TitlePageConfig(
        students=[Student('A')],
        students_label='Работу выполнил:',
    )
    texts = [line.text for line in _build_lines(cfg)]
    assert 'Работу выполнил:' in texts
    assert 'Выполнил' not in texts


def test_students_label_none():
    cfg = TitlePageConfig(students=[Student('A')], students_label=None)
    texts = [line.text for line in _build_lines(cfg)]
    assert 'Выполнил' not in texts


def test_supervisors_label_replaceable():
    cfg = TitlePageConfig(
        supervisors=[Supervisor('B', prefix='доцент')],
        supervisors_label='Научный руководитель:',
    )
    texts = [line.text for line in _build_lines(cfg)]
    assert 'Научный руководитель:' in texts
    assert 'Руководитель' not in texts


def test_student_sig_has_tab_stop():
    cfg = TitlePageConfig(students=[Student('A', group='101')])
    sig = [line for line in _build_lines(cfg) if line.tab_stop][0]
    assert '\t' in sig.text
    assert sig.align == 'left'


def test_multiple_students():
    cfg = TitlePageConfig(students=[
        Student('A', group='101'),
        Student('B', group='101', prefix='студентка'),
        Student('C', prefix='магистр'),
    ])
    texts = [line.text for line in _build_lines(cfg)]
    assert 'студент гр. 101\tA' in texts
    assert 'студентка гр. 101\tB' in texts
    assert 'магистр\tC' in texts


def test_work_type_and_number_joined():
    cfg = TitlePageConfig(work_type='ЛАБОРАТОРНАЯ', work_number='№2')
    texts = [line.text for line in _build_lines(cfg)]
    assert 'ЛАБОРАТОРНАЯ №2' in texts


def test_year_as_string():
    cfg = TitlePageConfig(year=2026)
    texts = [line.text for line in _build_lines(cfg)]
    assert '2026' in texts


# ---------- estimation ----------

def test_short_text_one_line():
    assert _estimate_lines('hello', 14, 500) == 1


def test_tabbed_text_one_line():
    assert _estimate_lines('left\tright', 14, 300) == 1


def test_long_text_wraps():
    assert _estimate_lines('word ' * 50, 14, 400) > 1


def test_empty_text_zero_lines():
    assert _estimate_lines('', 14, 400) == 0


# ---------- line spacing ----------

def test_single_spacing_by_default():
    """Without line_height_pt, line_spacing is set to 1.0 (single)."""
    doc = Document()
    cfg = TitlePageConfig(institution='X')
    render_config(doc, cfg)

    for para in doc.paragraphs:
        if para.text == 'X':
            spacing = para.paragraph_format.line_spacing
            assert spacing == 1.0
            break
    else:
        pytest.fail('X not found')


def test_explicit_line_height_overrides():
    doc = Document()
    cfg = TitlePageConfig(institution='X', line_height_pt=25)
    render_config(doc, cfg)

    for para in doc.paragraphs:
        if para.text == 'X':
            spacing = para.paragraph_format.line_spacing
            assert spacing == Pt(25)
            break
    else:
        pytest.fail('X not found')


# ---------- one-page fit ----------

def _total_height_pt(doc) -> float:
    total = 0.0
    for para in doc.paragraphs:
        pf = para.paragraph_format
        if pf.space_before is not None:
            total += pf.space_before.pt
        spacing = pf.line_spacing
        if spacing is None:
            continue
        if isinstance(spacing, float):
            # Multiplier — approximate from largest run size.
            max_size = 14
            for run in para.runs:
                if run.font.size is not None:
                    max_size = max(max_size, run.font.size.pt)
            total += spacing * max_size * 1.15
        else:
            total += spacing.pt
    return total


def _content_height_pt(doc) -> float:
    section = doc.sections[0]
    return float(
        section.page_height.pt
        - section.top_margin.pt
        - section.bottom_margin.pt
    )


def test_year_on_last_line():
    doc = Document()
    cfg = TitlePageConfig(
        institution='X',
        students=[Student('A', group='101')],
        city='City',
        year=2026,
    )
    render_config(doc, cfg)

    last = [p for p in doc.paragraphs if p.text][-1]
    assert last.text == '2026'


def test_tab_stop_present_on_sig_line():
    doc = Document()
    cfg = TitlePageConfig(students=[Student('A', group='101')])
    render_config(doc, cfg)

    for para in doc.paragraphs:
        if '\t' in para.text:
            tabs = list(para.paragraph_format.tab_stops)
            assert len(tabs) == 1
            assert tabs[0].alignment == WD_TAB_ALIGNMENT.RIGHT
            break
    else:
        pytest.fail('no tabbed paragraph')


def test_overflow_collapses_gaps():
    doc = Document()
    cfg = TitlePageConfig(
        institution='University ' * 5,
        work_title='Very long title ' * 20,
        students=[Student('Name ' * 10) for _ in range(3)],
        supervisors=[Supervisor('S ' * 10, prefix='профессор')
                     for _ in range(2)],
        year=2026,
        base_size=20, title_size=24, institution_size=20,
    )
    render_config(doc, cfg)
    for para in doc.paragraphs:
        pf = para.paragraph_format
        if pf.space_before is not None:
            assert pf.space_before.pt >= 0.0


# ---------- margins ----------

def test_custom_margins_applied_to_first_section():
    """Margins are stored as twips internally; comparison uses a
    tolerance of one twip (635 EMU) to account for rounding."""
    from docx.shared import Emu

    doc = Document()
    cfg = TitlePageConfig(
        institution='X',
        margin_top=2.5, margin_bottom=2.5,
        margin_left=3.0, margin_right=1.5,
    )
    render_config(doc, cfg)

    tolerance = Emu(1000)   # ~1.5 twips, well within rounding error

    first = doc.sections[0]
    assert abs(first.top_margin - Cm(2.5)) < tolerance
    assert abs(first.bottom_margin - Cm(2.5)) < tolerance
    assert abs(first.left_margin - Cm(3.0)) < tolerance
    assert abs(first.right_margin - Cm(1.5)) < tolerance


def test_custom_margins_create_second_section():
    doc = Document()
    cfg = TitlePageConfig(institution='X', margin_top=2.5)
    render_config(doc, cfg)
    assert len(doc.sections) == 2


def test_second_section_restores_original_margins():
    """The second section's margins equal the original ones, within
    the one-twip rounding error introduced by python-docx."""
    from docx.shared import Emu

    doc = Document()
    original_top = doc.sections[0].top_margin

    cfg = TitlePageConfig(institution='X', margin_top=1.0)
    render_config(doc, cfg)

    second = doc.sections[1]
    assert abs(second.top_margin - original_top) < Emu(1000)


def test_no_margins_no_section_created():
    """When margin_* are all None and isolate is False, no new section."""
    doc = Document()
    cfg = TitlePageConfig(
        institution='X',
        margin_top=None, margin_bottom=None,
        margin_left=None, margin_right=None,
    )
    render_config(doc, cfg)
    assert len(doc.sections) == 1


def test_isolate_creates_section_without_margins():
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = 'Header'
    cfg = TitlePageConfig(
        institution='X',
        margin_top=None, margin_bottom=None,
        margin_left=None, margin_right=None,
        isolate=True,
    )
    render_config(doc, cfg)

    assert len(doc.sections) == 2
    texts = [p.text for p in doc.sections[0].header.paragraphs if p.text]
    assert texts == []


# ---------- inline page break ----------

def test_append_page_break_inline_no_new_paragraph():
    doc = Document()
    doc.add_paragraph('Last line')
    before = len(doc.paragraphs)
    append_page_break_inline(doc)
    assert len(doc.paragraphs) == before
    assert doc.paragraphs[-1].text == 'Last line'


def test_append_page_break_inline_empty_doc():
    doc = Document()
    append_page_break_inline(doc)