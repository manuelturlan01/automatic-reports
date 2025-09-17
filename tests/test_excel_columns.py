from pathlib import Path
import sys

import openpyxl
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

    def parse_second(path, tz, now):
        if path == str(pdf_one):
            return {
                "N° Ticket": "123",
                "Título del ticket": "Actualizado",
                "Estado BW": "Cerrado",
                "Prioridad": "Automática",
                "Departamento": "Auto",
                "Fecha de creación": "",
                "Autor": "Alice",
                "Última respuesta por": "Bob",
                "Última respuesta el": "",
            }
        assert path == str(pdf_two)
        return {
            "N° Ticket": "456",
            "Título del ticket": "Nuevo",
            "Estado BW": "Abierto",
            "Prioridad": "Alta",
            "Departamento": "IT",
            "Fecha de creación": "",
            "Autor": "Carol",
            "Última respuesta por": "Dave",
            "Última respuesta el": "",
        }

    monkeypatch.setattr(
        tickets_parser,
        "glob",
        lambda pattern: sorted([str(pdf_one), str(pdf_two)]),
    )
    monkeypatch.setattr(tickets_parser, "parse_pdf", parse_second)
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

    ticket_idx = headers.index("N Ticket")
    title_idx = headers.index("Título del ticket")
    priority_idx = headers.index("Prioridad")
    area_idx = headers.index("Área")
    dept_idx = headers.index("Departamento")

    data_rows = []
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        ticket_value = row[ticket_idx]
        if ticket_value in (None, ""):
            break
        data_rows.append(row)

    ticket_values = [row[ticket_idx] for row in data_rows]
    assert ticket_values.count("123") == 1
    assert ticket_values[-1] == "456"

    ticket_123 = next(row for row in data_rows if row[ticket_idx] == "123")
    assert ticket_123[title_idx] == "Actualizado"
    assert ticket_123[priority_idx] == "Manual Prioridad"
    assert ticket_123[area_idx] == "Área Manual"
    assert ticket_123[dept_idx] == "Departamento Manual"

    ticket_456 = next(row for row in data_rows if row[ticket_idx] == "456")
    assert ticket_456[priority_idx] in (None, "")

    workbook.close()
