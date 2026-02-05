from __future__ import annotations

from pathlib import Path


def test_gdrive_helpers(load_main, tmp_path):
    main = load_main(env={"ASO_GDRIVE_NAME_FILTER": "asos enesa"})

    text = """
    Link1: https://drive.google.com/file/d/ABCdef12345/view
    Link2: https://drive.google.com/open?id=XYZ987654321
    Link3: https://drive.google.com/uc?export=download&id=ID_12345-6789
    """
    ids = main._extract_gdrive_file_ids(text)
    assert set(ids) == {"ABCdef12345", "XYZ987654321", "ID_12345-6789"}

    cd = "attachment; filename*=UTF-8''ASO%20Enesa.pdf"
    assert main._parse_filename_from_cd(cd) == "ASO Enesa.pdf"
    cd2 = 'attachment; filename="arquivo.pdf"'
    assert main._parse_filename_from_cd(cd2) == "arquivo.pdf"

    html = "<title>ASO ENESA.pdf - Google Drive</title>"
    assert main._parse_filename_from_html(html) == "ASO ENESA.pdf"

    html2 = 'confirm=abc123xyz'
    assert main._parse_confirm_token(html2) == "abc123xyz"

    assert main._safe_filename("..\\evil/..\\a.pdf") == "a.pdf"

    p = Path(tmp_path) / "file.pdf"
    p.write_text("x", encoding="utf-8")
    p2 = main._unique_path(str(p))
    assert p2 != str(p)

    assert main._gdrive_name_matches("ASOS ENESA - 1.pdf") is True
    assert main._gdrive_name_matches("OUTRO.pdf") is False
