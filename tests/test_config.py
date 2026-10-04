"""Tests for TitlePageConfig and participant dataclasses."""
from docx_report_gen_title_page import (
    Student, Supervisor, TitlePageConfig,
)


# ---------- Student ----------

def test_student_defaults():
    s = Student('Ерохин В.С.')
    assert s.name == 'Ерохин В.С.'
    assert s.group is None
    assert s.prefix == 'студент'


def test_student_left_with_group():
    s = Student('Ерохин В.С.', group='5130904/30103')
    assert s.left() == 'студент гр. 5130904/30103'


def test_student_left_without_group():
    s = Student('Иванов И.И.')
    assert s.left() == 'студент'


def test_student_left_empty_prefix():
    s = Student('Ерохин В.С.', prefix='')
    assert s.left() == ''


def test_student_render_full():
    s = Student('Ерохин В.С.', group='5130904/30103')
    assert s.render() == 'студент гр. 5130904/30103\tЕрохин В.С.'


def test_student_render_female():
    s = Student('Петрова А.А.', group='101', prefix='студентка')
    assert s.render() == 'студентка гр. 101\tПетрова А.А.'


def test_student_render_master():
    s = Student('Сидоров П.П.', prefix='магистр')
    assert s.render() == 'магистр\tСидоров П.П.'


def test_student_render_no_prefix():
    s = Student('Ерохин В.С.', prefix='')
    assert s.render() == 'Ерохин В.С.'


# ---------- Supervisor ----------

def test_supervisor_default():
    sup = Supervisor('Тутыгин В.С.')
    assert sup.prefix == ''
    assert sup.render() == 'Тутыгин В.С.'


def test_supervisor_render_with_prefix():
    sup = Supervisor('Тутыгин В.С.', prefix='доцент')
    assert sup.render() == 'доцент\tТутыгин В.С.'


def test_supervisor_render_professor():
    sup = Supervisor('Иванов И.И.', prefix='профессор')
    assert sup.render() == 'профессор\tИванов И.И.'


def test_supervisor_render_teacher():
    sup = Supervisor('Петров П.П.', prefix='преподаватель')
    assert sup.render() == 'преподаватель\tПетров П.П.'


# ---------- TitlePageConfig ----------

def test_all_content_fields_default_to_none():
    cfg = TitlePageConfig()
    assert cfg.institution is None
    assert cfg.institute is None
    assert cfg.school is None
    assert cfg.work_type is None
    assert cfg.work_number is None
    assert cfg.work_title is None
    assert cfg.discipline is None
    assert cfg.city is None
    assert cfg.year is None


def test_participant_lists_default_empty():
    cfg = TitlePageConfig()
    assert cfg.students == []
    assert cfg.supervisors == []


def test_participant_lists_are_per_instance():
    a = TitlePageConfig()
    b = TitlePageConfig()
    a.students.append(Student('X'))
    assert b.students == []


def test_localizable_defaults():
    cfg = TitlePageConfig()
    assert cfg.discipline_prefix == 'по дисциплине'
    assert cfg.students_label == 'Выполнил'
    assert cfg.supervisors_label == 'Руководитель'


def test_font_defaults():
    cfg = TitlePageConfig()
    assert cfg.font == 'Times New Roman'
    assert cfg.base_size == 14
    assert cfg.title_size == 16
    assert cfg.institution_size == 14
    assert cfg.line_height_pt is None


def test_gap_weight_defaults():
    cfg = TitlePageConfig()
    assert cfg.gap_after_header == 2.0
    assert cfg.gap_after_work == 0.5
    assert cfg.gap_after_title == 0.5
    assert cfg.gap_after_discipline_prefix == 0.3
    assert cfg.gap_after_discipline == 2.0
    assert cfg.gap_after_students_label == 0.3
    assert cfg.gap_after_students == 0.7
    assert cfg.gap_after_supervisors_label == 0.3
    assert cfg.gap_after_supervisors == 6.0


def test_explicit_values_kept():
    cfg = TitlePageConfig(
        institution='University',
        students=[Student('A', group='101')],
        supervisors=[Supervisor('B', prefix='доцент')],
        year=2026,
        gap_after_supervisors=5.0,
        line_height_pt=20,
    )
    assert cfg.institution == 'University'
    assert cfg.students[0].group == '101'
    assert cfg.supervisors[0].prefix == 'доцент'
    assert cfg.year == 2026
    assert cfg.gap_after_supervisors == 5.0
    assert cfg.line_height_pt == 20


def test_isolate_defaults_to_false():
    cfg = TitlePageConfig()
    assert cfg.isolate is False