import json
import os
import tempfile
import uuid
from datetime import datetime

from outcomes import (
    SUCCESS,
    ERROR,
    SKIPPED_DUPLICATE,
)
from utils_masking import mask_cpf, mask_pii_in_obj
from custom_logger import RpaLogger
from reporting import ReportGenerator
from idempotency import should_skip_duplicate


def _read_all_text(paths):
    content = ""
    for p in paths:
        if p and os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                content += f.read()
    return content


def main():
    execution_id = str(uuid.uuid4())
    started_at = datetime.now()
    cpf_raw = "12345678901"
    cpf_masked = mask_cpf(cpf_raw)
    file_ok_real = f"OK - {cpf_raw}.pdf"
    file_err_real = f"FAIL - {cpf_raw}.pdf"
    file_dup_real = f"DUP - {cpf_raw}.pdf"
    file_ok_display = f"OK - {cpf_masked}.pdf"
    file_err_display = f"FAIL - {cpf_masked}.pdf"
    file_dup_display = f"DUP - {cpf_masked}.pdf"

    with tempfile.TemporaryDirectory() as tmpdir:
        logs_dir = os.path.join(tmpdir, "logs")
        reports_dir = os.path.join(tmpdir, "relatorios")
        os.makedirs(logs_dir, exist_ok=True)
        os.makedirs(reports_dir, exist_ok=True)

        logger = RpaLogger(logs_dir, execution_id=execution_id)
        reporter = ReportGenerator(reports_dir)

        rpa_files = []
        duplicate_file = os.path.join(tmpdir, f"JOAO - {cpf_raw}.pdf")
        with open(duplicate_file, "wb") as f:
            f.write(b"dummy")

        # Fake items
        items = [
            {"file_real": file_ok_real, "file_display": file_ok_display, "outcome": SUCCESS, "message": "OK"},
            {"file_real": file_err_real, "file_display": file_err_display, "outcome": ERROR, "message": "Falha simulada"},
            {"file_real": file_dup_real, "file_display": file_dup_display, "outcome": SKIPPED_DUPLICATE, "message": "Duplicado"},
        ]

        # Log 3 events
        logger.info("Evento sucesso", extra={"cpf": cpf_raw, "outcome": SUCCESS}, step="smoke")
        logger.error("Evento erro", extra={"cpf": cpf_raw, "outcome": ERROR}, step="smoke")
        logger.warning("Evento duplicado", extra={"cpf": cpf_raw, "outcome": SKIPPED_DUPLICATE}, step="smoke")

        # Real idempotency check using project function
        duplicate_dest = os.path.join(tmpdir, items[2]["file_real"])
        with open(duplicate_dest, "wb") as f:
            f.write(b"duplicate")
        assert should_skip_duplicate(duplicate_dest) is True, "should_skip_duplicate failed"

        for it in items:
            if it["outcome"] == SKIPPED_DUPLICATE:
                continue
            rpa_files.append(it["file_real"])

        stats = {
            "execution_id": execution_id,
            "started_at": started_at.isoformat(),
            "tempo_total": "00:00:01",
            "total_detected": 3,
            "total_processed": 2,
            "success": 1,
            "error": 1,
            "skipped_duplicate": 1,
            "skipped_draft": 0,
            "skipped_non_aso": 0,
            "sucessos": [items[0]["file_display"]],
            "erros": [{"arquivo": items[1]["file_display"], "erro": "Falha simulada"}],
            "skipped_items": [f"{SKIPPED_DUPLICATE}: {items[2]['file_display']}"],
            "run_status": "INCONSISTENT",
        }

        stats_masked = mask_pii_in_obj(stats)
        report_paths = reporter.save_report(stats_masked)

        manifest = {
            "execution_id": execution_id,
            "started_at": started_at.isoformat(),
            "finished_at": datetime.now().isoformat(),
            "duration_sec": 1,
            "run_status": "INCONSISTENT",
            "email_status": "SKIPPED_NO_RECIPIENT",
            "email_error": None,
            "paths": {
                "report_json": report_paths.get("json"),
                "report_md": report_paths.get("md"),
                "logs": logs_dir,
            },
            "totals": {
                "total_detected": stats["total_detected"],
                "total_processed": stats["total_processed"],
                "success": stats["success"],
                "error": stats["error"],
                "skipped_duplicate": stats["skipped_duplicate"],
                "skipped_draft": stats["skipped_draft"],
                "skipped_non_aso": stats["skipped_non_aso"],
            },
            "items": [
                {"file_display": items[0]["file_display"], "cpf_masked": cpf_masked, "outcome": SUCCESS},
                {"file_display": items[1]["file_display"], "cpf_masked": cpf_masked, "outcome": ERROR},
                {"file_display": items[2]["file_display"], "cpf_masked": cpf_masked, "outcome": SKIPPED_DUPLICATE},
            ],
        }

        manifest = mask_pii_in_obj(manifest)
        manifest_path = os.path.join(reports_dir, f"manifest_{execution_id}.json")
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=4, ensure_ascii=False)

        rpa_log_path = os.path.join(reports_dir, "rpa_log.csv")
        with open(rpa_log_path, "w", encoding="utf-8") as f:
            f.write("timestamp,cpf,file,status,message\n")
            f.write(f"{datetime.now().isoformat()},{cpf_masked},{file_ok_display},sucesso,ok\n")

        # Assertions
        # JSONL lines are valid JSON
        assert os.path.exists(logger.log_file), "JSONL log file not created"
        with open(logger.log_file, "r", encoding="utf-8") as f:
            lines = [line for line in f.read().splitlines() if line.strip()]
        assert len(lines) >= 3, "Expected at least 3 JSONL lines"
        for line in lines[:3]:
            assert line.startswith("{") and line.endswith("}"), "JSONL line not clean JSON object"
            assert "[" not in line[:10], "JSONL line has prefix text"
            obj = json.loads(line)
            assert obj.get("execution_id") == execution_id, "execution_id mismatch in JSONL"
            assert obj.get("level") in {"INFO", "ERROR", "WARNING"}, "level missing in JSONL"
            assert obj.get("step") == "smoke", "step missing in JSONL"
            assert obj.get("message"), "message missing in JSONL"
            details = obj.get("details") or {}
            extra = obj.get("extra") or {}
            cand = details.get("cpf") or extra.get("cpf")
            if cand:
                assert cand != cpf_raw, "CPF raw in JSONL details"

        # No raw CPF in artifacts
        all_text = _read_all_text([logger.log_file, report_paths.get("json"), report_paths.get("md"), manifest_path, rpa_log_path])
        assert cpf_raw not in all_text, "Raw CPF leaked in artifacts"

        # Duplicate not added to RPA list
        assert items[2]["file_real"] not in rpa_files, "Duplicate entered RPA list"
        assert any(SKIPPED_DUPLICATE in s for s in stats_masked["skipped_items"]), "Duplicate outcome missing"

        # Paths exist
        assert report_paths.get("json"), "Report JSON path missing"
        assert report_paths.get("md"), "Report MD path missing"
        assert os.path.exists(report_paths.get("json")), "Report JSON missing"
        assert os.path.exists(report_paths.get("md")), "Report MD missing"
        assert os.path.exists(manifest_path), "Manifest missing"

    print("SMOKE OK")


if __name__ == "__main__":
    main()
