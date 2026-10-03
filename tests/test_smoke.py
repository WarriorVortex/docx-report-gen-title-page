"""Smoke tests — the package imports, version reads from metadata."""
import docx_report_gen_title_page as pkg


def test_package_imports():
    assert pkg is not None


def test_version_is_string():
    assert isinstance(pkg.__version__, str)
    assert pkg.__version__


def test_public_api_exports():
    assert hasattr(pkg, 'TitlePagePlugin')
    assert hasattr(pkg, 'TitlePageConfig')
    assert 'TitlePagePlugin' in pkg.__all__
    assert 'TitlePageConfig' in pkg.__all__