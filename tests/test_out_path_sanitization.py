import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("pandas")
pytest.importorskip("openpyxl")


def test_out_path_trims_extra_suffix(tmp_path):
    repo_root = Path(__file__).resolve().parents[1]
    pdf_dir = repo_root / "reports"
    raw_out = tmp_path / "Tickets.xlsxclear"
    expected_out = tmp_path / "Tickets.xlsx"

    result = subprocess.run(
        [
            sys.executable,
            str(repo_root / "tickets_parser.py"),
            "--pdf_dir",
            str(pdf_dir),
            "--pattern",
            "Ticket-115249.pdf",
            "--out",
            str(raw_out),
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    assert expected_out.exists(), result.stderr
    assert not raw_out.exists()
