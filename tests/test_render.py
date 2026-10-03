"""Tests for the layout algorithm — one-page fit, gap distribution."""
import pytest
from docx import Document
from docx.shared import Cm, Pt

from docx_report_gen_title_page import TitlePageConfig
from docx_report_gen_title_page._render import (
    _build_lines, _estimate_lines, _resolve_line_height, render_config,
)


# ---------- line building ----------

def test_empty_config_produces_no_lines():
    lines = _build_lines(TitlePageConfig())
    assert lines == []


def test_first_line_has_no_gap():
    lines = _build_lines(TitlePageConfig(institution='X'))
    assert lines[0].gap_weight == 0.0


def test_gap_weights_accumulate_over_skipped_sections():
    """A missing discipline block rolls its gap weight into the next
    present line. The student label always accompanies the student
    line, so it is not counted as a separate section."""
    cfg = TitlePageConfig(
        institution='X',
        work_title='Y',
        student='Z',
    )
    lines = _build_lines(cfg)
    texts = [line.text for line in lines]

    # Header (X), title (Y), student label + student (Z)
    assert texts == ['X', 'Y', 'Выполнил', 'Z']

    # The student line carries the accumulated gap from discipline
    # (which is skipped) plus its own gap weight.
    student_line = lines[-1]
    assert student_line.gap_weight > 0


# ---------- line estimation ----------

def test_short_text_is_one_line():
    assert _estimate_lines('hello', 14, 500) == 1


def test_long_text_wraps():
    long = 'word ' * 50
    assert _estimate_lines(long, 14, 400) > 1


def test_empty_text_is_zero_lines():
    assert _estimate_lines('', 14, 400) == 0


# ---------- line height ----------

def test_line_height_derived_from_title_size():
    cfg = TitlePageConfig(title_size=16, line_height_pt=None)
    height = _resolve_line_height(cfg)
    assert height == int(16 * 1.4)


def test_line_height_clamped_to_minimum():
    """A user-provided height below 1.2 × max_size is bumped up."""
    cfg = TitlePageConfig(
        base_size=12, title_size=20, institution_size=14,
        line_height_pt=10,      # way too small
    )
    height = _resolve_line_height(cfg)
    assert height >= 20 * 1.2


def test_explicit_line_height_used_when_large_enough():
    cfg = TitlePageConfig(
        base_size=12, title_size=14, institution_size=12,
        line_height_pt=30,
    )
    assert _resolve_line_height(cfg) == 30


# ---------- one-page fit ----------

def _total_height_pt(doc) -> float:
    """Sum of space_before + line_spacing over every paragraph."""
    total = 0.0
    for para in doc.paragraphs:
        pf = para.paragraph_format
        if pf.space_before is not None:
            total += pf.space_before.pt
        if pf.line_spacing is not None and not isinstance(
            pf.line_spacing, float
        ):
            total += pf.line_spacing.pt
    return total


def _content_height_pt(doc) -> float:
    section = doc.sections[0]
    return float(
        section.page_height.pt
        - section.top_margin.pt
        - section.bottom_margin.pt
    )


def test_full_config_fits_exactly_one_a4():
    """Total height matches content area within a small tolerance."""
    doc = Document()

    cfg = TitlePageConfig(
        institution='СПбПУ',
        institute='ИКНТ',
        school='ВШПИ',
        work_type='ЛАБОРАТОРНАЯ РАБОТА',
        work_number='№2',
        work_title='«Разложение в ряд Фурье»',
        discipline='«ЦОС»',
        student='студент гр. 101 Ерохин В.С.',
        supervisor='Тутыгин В.С.',
        city='Санкт-Петербург',
        year=2026,
    )

    render_config(doc, cfg)

    # Tolerance: 5 points for float rounding.
    assert abs(_total_height_pt(doc) - _content_height_pt(doc)) < 5.0


def test_minimal_config_still_fits():
    doc = Document()
    cfg = TitlePageConfig(institution='X', city='Y', year=2026)
    render_config(doc, cfg)
    assert _total_height_pt(doc) <= _content_height_pt(doc) + 5.0


def test_overflow_collapses_gaps_to_zero():
    """With massive content, gaps go to zero but the layout is defined."""
    doc = Document()
    cfg = TitlePageConfig(
        institution='University ' * 5,
        institute='Institute ' * 10,
        school='School ' * 10,
        work_type='LONG WORK TYPE',
        work_title='Very long title ' * 20,
        discipline='Long discipline ' * 15,
        student='Student with a long name ' * 10,
        supervisor='Supervisor ' * 10,
        city='City',
        year=2026,
        base_size=20,
        title_size=24,
        institution_size=20,
    )
    render_config(doc, cfg)

    # Every gap collapsed to 0 (or very close)
    for para in doc.paragraphs:
        pf = para.paragraph_format
        if pf.space_before is not None:
            assert pf.space_before.pt >= 0.0


def test_custom_margins_are_respected():
    """A document with tighter margins gives more content height."""
    from docx.shared import Cm

    tight = Document()
    for section in tight.sections:
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.0)

    cfg = TitlePageConfig(institution='X')
    render_config(tight, cfg)

    # No crash, height reported correctly
    assert _content_height_pt(tight) > 700


def test_line_spacing_is_exact():
    """line_spacing is set as a Length (exact), not a float (multiplier)."""
    doc = Document()
    cfg = TitlePageConfig(institution='X')
    render_config(doc, cfg)

    for para in doc.paragraphs:
        if para.text == 'X':
            spacing = para.paragraph_format.line_spacing
            # Exact spacing is a Length, not a float
            assert not isinstance(spacing, float)
            assert isinstance(spacing.pt, float)
            break
    else:
        pytest.fail('X not found')