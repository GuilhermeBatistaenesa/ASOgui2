import json
from pathlib import Path

from reporting import ReportGenerator


def test_report_generator_writes_and_masks(tmp_path):
    stats = {
        "execution_id": "exec-1",
        "tempo_total": "00:00:01",
        "total_detected": 1,
        "total_processed": 1,
        "success": 1,
        "error": 0,
        "erros": [],
        "sucessos": ["JOAO - 12345678901.pdf"],
        "skipped_items": [],
    }

    rg = ReportGenerator(str(tmp_path))
    paths = rg.save_report(stats)

    assert paths["json"]
    assert paths["md"]

    json_path = Path(paths["json"])
    md_path = Path(paths["md"])

    assert json_path.exists()
    assert md_path.exists()

    data = json.loads(json_path.read_text(encoding="utf-8"))
    assert "12345678901" not in json.dumps(data)
    assert "12345678901" not in md_path.read_text(encoding="utf-8")
