from datetime import datetime, timedelta
from pathlib import Path
import sys

import openpyxl
from openpyxl.utils import get_column_letter
import pytest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import tickets_parser


def test_generated_excel_has_expected_columns(tmp_path, monkeypatch):
    output_path = tmp_path / "Tickets.xlsx"
    dummy_pdf = tmp_path / "Ticket-0001.pdf"
    dummy_pdf.write_bytes(b"")

    monkeypatch.setattr(tickets_parser, "glob", lambda pattern: [str(dummy_pdf)])

    def fake_parse_pdf(path, tz, now):
        return {
            "N° Ticket": "123",
            "Título del ticket": "Prueba",
            "Estado BW": "Abierto",
            "Prioridad": "Alta",
            "Departamento": "IT",
            "Fecha de creación": "",
            "Autor": "Alice",
            "Última respuesta por": "Bob",
            "Última respuesta el": "",
        }

    monkeypatch.setattr(tickets_parser, "parse_pdf", fake_parse_pdf)

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "tickets_parser.py",
            "--pdf_dir",
            str(tmp_path),
            "--out",
            str(output_path),
        ],
    )

    tickets_parser.main()

    workbook = openpyxl.load_workbook(output_path)
    worksheet = workbook["Tickets"]
    header_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))

    expected_order = [
        "N Ticket",
        "Título del ticket",
        "Autor",
        "Prioridad",
        "Área",
        "Departamento",
    ]

    assert header_row[: len(expected_order)] == tuple(expected_order)

    priority_column_index = header_row.index("Prioridad") + 1
    priority_cell_value = worksheet.cell(row=2, column=priority_column_index).value
    assert priority_cell_value in (None, "")

    priority_column_letter = get_column_letter(priority_column_index)
    expected_priority_range = f"{priority_column_letter}2:{priority_column_letter}200"
    expected_formula = '"' + ",".join(tickets_parser.PRIORITY_OPTIONS) + '"'

    priority_validation = None
    for validation in worksheet.data_validations.dataValidation:
        if (
            str(validation.sqref) == expected_priority_range
            and validation.type == "list"
        ):
            priority_validation = validation
            break

    assert priority_validation is not None
    assert priority_validation.formula1 == expected_formula
    assert priority_validation.allow_blank


def test_reprocess_preserves_manual_columns_and_appends_new(tmp_path, monkeypatch):
    output_path = tmp_path / "Tickets.xlsx"
    pdf_one = tmp_path / "Ticket-0001.pdf"
    pdf_two = tmp_path / "Ticket-0002.pdf"
    pdf_one.write_bytes(b"")
    pdf_two.write_bytes(b"")

    def parse_first(path, tz, now):
        assert path == str(pdf_one)
        return {
            "N° Ticket": "123",
            "Título del ticket": "Inicial",
            "Estado BW": "Abierto",
            "Prioridad": "Alta",
            "Departamento": "IT",
            "Fecha de creación": "",
            "Autor": "Alice",
            "Última respuesta por": "Bob",
            "Última respuesta el": "",
        }

    monkeypatch.setattr(tickets_parser, "glob", lambda pattern: [str(pdf_one)])
    monkeypatch.setattr(tickets_parser, "parse_pdf", parse_first)
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "tickets_parser.py",
            "--pdf_dir",
            str(tmp_path),
            "--out",
            str(output_path),
        ],
    )

    tickets_parser.main()

    workbook = openpyxl.load_workbook(output_path)
    worksheet = workbook["Tickets"]
    headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]

    ticket_idx = headers.index("N Ticket") + 1
    priority_idx = headers.index("Prioridad") + 1
    area_idx = headers.index("Área") + 1
    dept_idx = headers.index("Departamento") + 1

    worksheet.cell(row=2, column=priority_idx, value="Manual Prioridad")
    worksheet.cell(row=2, column=area_idx, value="Área Manual")
    worksheet.cell(row=2, column=dept_idx, value="Departamento Manual")

    workbook.save(output_path)
    workbook.close()

    def parse_second(path, tz, now):
        assert path in {str(pdf_one), str(pdf_two)}
        if path == str(pdf_one):
            return {
                "N° Ticket": "123",
                "Título del ticket": "Actualizado",
                "Estado BW": "Cerrado",
                "Prioridad": "Alta",
                "Departamento": "IT",
                "Fecha de creación": "",
                "Autor": "Alice",
                "Última respuesta por": "Bob",
                "Última respuesta el": "",
            }
        return {
            "N° Ticket": "456",
            "Título del ticket": "Nuevo",
            "Estado BW": "Abierto",
            "Prioridad": "Media",
            "Departamento": "Operaciones",
            "Fecha de creación": "",
            "Autor": "Carol",
            "Última respuesta por": "Dave",
            "Última respuesta el": "",
        }

    monkeypatch.setattr(
        tickets_parser,
        "glob",
        lambda pattern: [str(pdf_one), str(pdf_two)],
    )
    monkeypatch.setattr(tickets_parser, "parse_pdf", parse_second)

    tickets_parser.main()

    workbook = openpyxl.load_workbook(output_path)
    worksheet = workbook["Tickets"]
    headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]

    ticket_idx = headers.index("N Ticket") + 1
    title_idx = headers.index("Título del ticket") + 1

    rows = list(worksheet.iter_rows(min_row=2, values_only=True))
    row_by_ticket = {}
    for row in rows:
        ticket = row[ticket_idx - 1]
        if ticket:
            row_by_ticket.setdefault(ticket, []).append(row)

    assert "123" in row_by_ticket
    assert len(row_by_ticket["123"]) == 1

    ticket_123 = row_by_ticket["123"][0]
    assert ticket_123[title_idx - 1] == "Actualizado"
    assert ticket_123[priority_idx - 1] == "Manual Prioridad"
    assert ticket_123[area_idx - 1] == "Área Manual"
    assert ticket_123[dept_idx - 1] == "Departamento Manual"

    assert "456" in row_by_ticket
    assert len(row_by_ticket["456"]) == 1

    ticket_456 = row_by_ticket["456"][0]
    assert ticket_456[priority_idx - 1] in (None, "")
    assert ticket_456[area_idx - 1] in (None, "")
    assert ticket_456[dept_idx - 1] in (None, "")

    workbook.close()


def test_dates_and_durations_are_written_with_native_types(tmp_path, monkeypatch):
    output_path = tmp_path / "Tickets.xlsx"
    pdf_path = tmp_path / "Ticket-0001.pdf"
    pdf_path.write_bytes(b"")

    class FixedDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2024, 1, 10, 12, 0, tzinfo=tz)

    def fake_parse_pdf(path, tz, now):
        assert path == str(pdf_path)
        return {
            "N° Ticket": "789",
            "Título del ticket": "Fechas",
            "Estado BW": "Abierto",
            "Prioridad": "Media",
            "Departamento": "IT",
            "Fecha de creación": "09/01/2024 08:00",
            "Autor": "Eve",
            "Última respuesta por": "Frank",
            "Última respuesta el": "10/01/2024 09:30",
        }

    monkeypatch.setattr(tickets_parser, "glob", lambda pattern: [str(pdf_path)])
    monkeypatch.setattr(tickets_parser, "parse_pdf", fake_parse_pdf)
    monkeypatch.setattr(tickets_parser, "datetime", FixedDateTime)
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "tickets_parser.py",
            "--pdf_dir",
            str(tmp_path),
            "--out",
            str(output_path),
        ],
    )

    tickets_parser.main()

    workbook = openpyxl.load_workbook(output_path)
    worksheet = workbook["Tickets"]
    headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]

    creation_idx = headers.index("Fecha de creación") + 1
    last_response_idx = headers.index("Última respuesta el") + 1
    wait_idx = headers.index("Tiempo parado desde la última respuesta") + 1
    open_idx = headers.index("Tiempo abierto (si sigue abierto)") + 1
    wait_seconds_idx = (
        headers.index("Tiempo parado desde la última respuesta (segundos)") + 1
    )
    open_seconds_idx = (
        headers.index("Tiempo abierto (si sigue abierto) (segundos)") + 1
    )

    creation_cell = worksheet.cell(row=2, column=creation_idx)
    last_response_cell = worksheet.cell(row=2, column=last_response_idx)
    wait_cell = worksheet.cell(row=2, column=wait_idx)
    open_cell = worksheet.cell(row=2, column=open_idx)
    wait_seconds_cell = worksheet.cell(row=2, column=wait_seconds_idx)
    open_seconds_cell = worksheet.cell(row=2, column=open_seconds_idx)

    assert isinstance(creation_cell.value, datetime)
    assert isinstance(last_response_cell.value, datetime)
    assert creation_cell.number_format == "yyyy-mm-dd hh:mm:ss"
    assert last_response_cell.number_format == "yyyy-mm-dd hh:mm:ss"

    assert isinstance(wait_cell.value, timedelta)
    assert isinstance(open_cell.value, timedelta)
    assert wait_cell.number_format == "[h]:mm:ss"
    assert open_cell.number_format == "[h]:mm:ss"

    assert isinstance(wait_seconds_cell.value, (int, float))
    assert isinstance(open_seconds_cell.value, (int, float))
    assert wait_seconds_cell.number_format == "0"
    assert open_seconds_cell.number_format == "0"

    from openpyxl.utils.datetime import to_excel

    assert to_excel(wait_cell.value) == pytest.approx((2 * 3600 + 30 * 60) / 86400)
    assert to_excel(open_cell.value) == pytest.approx((28 * 3600) / 86400)
    assert "1899" not in str(wait_cell.value)
    assert "1899" not in str(open_cell.value)
    assert "/" not in wait_cell.number_format
    assert "/" not in open_cell.number_format

    assert wait_seconds_cell.value == pytest.approx(2 * 3600 + 30 * 60)
    assert open_seconds_cell.value == pytest.approx(28 * 3600)

    workbook.close()
