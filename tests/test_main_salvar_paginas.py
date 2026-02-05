from __future__ import annotations

from pathlib import Path

from PIL import Image


def test_salvar_paginas_individualmente_flow(load_main, tmp_path, monkeypatch):
    main = load_main()

    img = Image.new("RGB", (10, 10), color="white")
    monkeypatch.setattr(main, "convert_from_path", lambda *_args, **_kwargs: [img, img, img])
    monkeypatch.setattr(main, "ocr_with_fallback", lambda *_args, **_kwargs: "dummy")

    results = [
        ("RASCUNHO", "Ignorar", "", "", "RASCUNHO RASCUNHO RASCUNHO RASCUNHO"),
        ("JOAO DA SILVA", "123.456.789-01", "01/02/2025", "SOLDADOR", "ASO OK"),
        ("JOAO DA SILVA", "123.456.789-01", "01/02/2025", "SOLDADOR", "NAO ASO"),
    ]
    it = iter(results)

    def _fake_extrair(*_args, **_kwargs):
        return next(it)

    monkeypatch.setattr(main, "extrair_dados_completos", _fake_extrair)

    def _fake_eh_aso(texto):
        return "ASO OK" in texto

    monkeypatch.setattr(main, "eh_aso", _fake_eh_aso)

    pdf_path = Path(tmp_path) / "input.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")

    stats = {
        "total_detected": 0,
        "total_processed": 0,
        "success": 0,
        "error": 0,
        "skipped_duplicate": 0,
        "skipped_draft": 0,
        "skipped_non_aso": 0,
        "ocr_failures": [],
        "erros": [],
        "skipped_items": [],
    }
    manifest_items = []
    novos = []

    main.salvar_paginas_individualmente(
        str(pdf_path),
        str(tmp_path),
        "1234",
        lista_novos_arquivos=novos,
        stats=stats,
        manifest_items=manifest_items,
    )

    assert stats["total_detected"] == 3
    assert stats["skipped_draft"] == 1
    assert stats["skipped_non_aso"] == 1
    assert len(novos) == 1
    assert any(item["outcome"] == "SKIPPED_DRAFT" for item in manifest_items)
    assert any(item["outcome"] == "SKIPPED_NON_ASO" for item in manifest_items)
