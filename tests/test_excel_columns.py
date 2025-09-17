from datetime import datetime, timedelta
from pathlib import Path
import sys
import zipfile
import xml.etree.ElementTree as ET

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
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

    creation_cell = worksheet.cell(row=2, column=creation_idx)
    last_response_cell = worksheet.cell(row=2, column=last_response_idx)
    wait_cell = worksheet.cell(row=2, column=wait_idx)
    open_cell = worksheet.cell(row=2, column=open_idx)

    assert isinstance(creation_cell.value, datetime)
    assert isinstance(last_response_cell.value, datetime)
    assert creation_cell.number_format == "yyyy-mm-dd hh:mm:ss"
    assert last_response_cell.number_format == "yyyy-mm-dd hh:mm:ss"

    expected_wait_td = timedelta(hours=2, minutes=30)
    expected_open_td = timedelta(days=1, hours=4)
    assert isinstance(wait_cell.value, timedelta)
    assert isinstance(open_cell.value, timedelta)
    assert wait_cell.value == expected_wait_td
    assert open_cell.value == expected_open_td
    assert wait_cell.number_format == "[h]:mm:ss"
    assert open_cell.number_format == "[h]:mm:ss"

    sheet_index = workbook.sheetnames.index("Tickets") + 1
    sheet_path = f"xl/worksheets/sheet{sheet_index}.xml"
    with zipfile.ZipFile(output_path) as archive:
        sheet_xml = ET.fromstring(archive.read(sheet_path))

    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    def get_raw_cell(coord: str):
        for cell in sheet_xml.iter(f"{namespace}c"):
            if cell.attrib.get("r") == coord:
                return cell.attrib.get("t"), cell.find(f"{namespace}v").text
        raise AssertionError(f"Missing cell {coord} in worksheet XML")

    wait_coord = f"{get_column_letter(wait_idx)}2"
    open_coord = f"{get_column_letter(open_idx)}2"
    wait_type, wait_raw = get_raw_cell(wait_coord)
    open_type, open_raw = get_raw_cell(open_coord)

    expected_wait = expected_wait_td.total_seconds() / (24 * 60 * 60)
    expected_open = expected_open_td.total_seconds() / (24 * 60 * 60)

    assert wait_type == "n"
    assert open_type == "n"
    assert float(wait_raw) == pytest.approx(expected_wait)
    assert float(open_raw) == pytest.approx(expected_open)

    workbook.close()
