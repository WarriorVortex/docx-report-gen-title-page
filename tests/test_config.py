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


def test_layout_defaults():
    cfg = TitlePageConfig()
    assert cfg.font == 'Times New Roman'
    assert cfg.base_size == 14
    assert cfg.title_size == 16
    assert cfg.institution_size == 14
    assert cfg.line_spacing == 1.5
    assert cfg.bottom_gap_lines == 6


def test_explicit_values_are_kept():
    cfg = TitlePageConfig(
        institution='University',
        work_type='LAB',
        work_number='№1',
        year=2026,
        bottom_gap_lines=4,
    )
    assert cfg.institution == 'University'
    assert cfg.work_type == 'LAB'
    assert cfg.work_number == '№1'
    assert cfg.year == 2026
    assert cfg.bottom_gap_lines == 4