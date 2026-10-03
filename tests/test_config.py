"""Tests for TitlePageConfig — defaults and field acceptance."""
from docx_report_gen_title_page import TitlePageConfig


def test_all_content_fields_default_to_none():
    cfg = TitlePageConfig()
    assert cfg.institution is None
    assert cfg.institute is None
    assert cfg.school is None
    assert cfg.work_type is None
    assert cfg.work_number is None
    assert cfg.work_title is None
    assert cfg.discipline is None
    assert cfg.student is None
    assert cfg.supervisor is None
    assert cfg.city is None
    assert cfg.year is None


def test_localizable_fields_have_defaults():
    cfg = TitlePageConfig()
    assert cfg.discipline_prefix == 'по дисциплине'
    assert cfg.student_label == 'Выполнил'
    assert cfg.supervisor_label == 'Руководитель'


def test_font_defaults():
    cfg = TitlePageConfig()
    assert cfg.font == 'Times New Roman'
    assert cfg.base_size == 14
    assert cfg.title_size == 16
    assert cfg.institution_size == 14
    assert cfg.line_height_pt is None      # derived from font sizes


def test_gap_weight_defaults():
    cfg = TitlePageConfig()
    assert cfg.gap_after_header == 2.0
    assert cfg.gap_after_work == 0.5
    assert cfg.gap_after_title == 0.5
    assert cfg.gap_after_discipline_prefix == 0.3
    assert cfg.gap_after_discipline == 2.0
    assert cfg.gap_after_student_label == 0.3
    assert cfg.gap_after_student == 0.5
    assert cfg.gap_after_supervisor == 3.0


def test_explicit_values_are_kept():
    cfg = TitlePageConfig(
        institution='University',
        work_type='LAB',
        work_number='№1',
        year=2026,
        gap_after_supervisor=5.0,
        line_height_pt=20,
    )
    assert cfg.institution == 'University'
    assert cfg.work_type == 'LAB'
    assert cfg.work_number == '№1'
    assert cfg.year == 2026
    assert cfg.gap_after_supervisor == 5.0
    assert cfg.line_height_pt == 20


def test_gap_weights_can_be_disabled():
    """Zeroing a weight removes the corresponding gap entirely."""
    cfg = TitlePageConfig(
        gap_after_header=0.0,
        gap_after_supervisor=0.0,
    )
    assert cfg.gap_after_header == 0.0
    assert cfg.gap_after_supervisor == 0.0