#!/usr/bin/env python3
"""Заполняет `report-template.docx` данными из CSV — скрипт эталонного скилла.

Данные вымышленные, компания «Меридиан» не существует.

Использование:
    python fill_report.py <шаблон.docx> <данные.csv> <выход.docx>

Шаблон и правила заполнения — `../SKILL.md`. Скрипт написан честно под нейтральный
шаблон (простая таблица без слияний); на шаблоне со слитыми ячейками или колонтитулами
с полями компании результат не гарантирован — это `demo-break/`.
"""
import copy
import csv
import sys
from pathlib import Path

from docx import Document

PERIOD = "август 2026"


def load_rows(csv_path: Path) -> list[dict]:
    with csv_path.open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def fill(template_path: Path, csv_path: Path, output_path: Path) -> None:
    rows = load_rows(csv_path)
    total = sum(float(row["amount_usd"]) for row in rows)

    doc = Document(str(template_path))
    table = doc.tables[0]

    template_tr = table.rows[-1]._tr
    tbl = table._tbl
    for _ in rows:
        tbl.append(copy.deepcopy(template_tr))
    tbl.remove(template_tr)

    data_rows = table.rows[1:]
    for row, csv_row in zip(data_rows, rows):
        cells = row.cells
        cells[0].text = csv_row["date"]
        cells[1].text = csv_row["client"]
        cells[2].text = csv_row["product"]
        cells[3].text = csv_row["qty"]
        cells[4].text = f'{float(csv_row["amount_usd"]):,.0f}'.replace(",", " ")

    total_str = f"{total:,.0f}".replace(",", " ")
    for p in doc.paragraphs:
        for run in p.runs:
            if "{{PERIOD}}" in run.text:
                run.text = run.text.replace("{{PERIOD}}", PERIOD)
            if "{{ROWS_COUNT}}" in run.text:
                run.text = run.text.replace("{{ROWS_COUNT}}", str(len(rows)))
            if "{{TOTAL}}" in run.text:
                run.text = run.text.replace("{{TOTAL}}", total_str)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def main() -> None:
    if len(sys.argv) != 4:
        print("Использование: fill_report.py <шаблон.docx> <данные.csv> <выход.docx>")
        raise SystemExit(1)
    fill(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))


if __name__ == "__main__":
    main()
