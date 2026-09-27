#!/usr/bin/env python3
"""Генератор синтетических шаблонов отчёта об отгрузках.

Данные вымышленные, компания «Меридиан» не существует.

Пишет детерминированно:
- `skill-skeleton/templates/report-template.docx` (+ `.md`-близнец) — нейтральный шаблон;
- `demo-naive/report-template.docx` (+ `.md`) — байт-в-байт копия нейтрального шаблона;
- `demo-break/report-template-break.docx` (+ `.md`) — ломающийся шаблон: колонтитулы с полями
  компании, объединённые ячейки, двухуровневая шапка таблицы (провал воспроизводит история 03).

Запуск: python 03-skills/gen/make_templates.py
"""
import shutil
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent

COMPANY = "Меридиан"
TITLE = "ОТЧЁТ ОБ ОТГРУЗКАХ"

PERIOD_PLACEHOLDER = "{{PERIOD}}"
ROWS_PLACEHOLDER = "{{ROWS_COUNT}}"
TOTAL_PLACEHOLDER = "{{TOTAL}}"

HEADER_CELLS = ["Дата", "Клиент", "Товар", "Кол-во", "Сумма, $"]
ROW_PLACEHOLDERS = ["{{DATE}}", "{{CLIENT}}", "{{PRODUCT}}", "{{QTY}}", "{{AMOUNT}}"]


def build_neutral_document() -> Document:
    """Шаблон, который честно заполняет `fill_report.py`: простая таблица без слияний."""
    doc = Document()
    doc.add_paragraph(TITLE)
    doc.add_paragraph(COMPANY)
    doc.add_paragraph(f"Период: {PERIOD_PLACEHOLDER}")

    table = doc.add_table(rows=2, cols=5)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, HEADER_CELLS):
        cell.text = text
    for cell, text in zip(table.rows[1].cells, ROW_PLACEHOLDERS):
        cell.text = text

    doc.add_paragraph(f"Итого по периоду: {ROWS_PLACEHOLDER} позиций, {TOTAL_PLACEHOLDER} $")
    return doc


def build_break_document() -> Document:
    """Шаблон, структура которого ломает наивную подстановку по индексу колонки:
    колонтитулы с полями компании (скрипт их не трогает), двухуровневая шапка таблицы
    (верхняя строка — одна объединённая ячейка) и объединённые ячейки в строке данных
    (Клиент+Товар в одной ячейке).
    """
    doc = Document()

    section = doc.sections[0]
    section.header.paragraphs[0].text = f"{COMPANY} — {TITLE} — Период: {PERIOD_PLACEHOLDER}"
    section.footer.paragraphs[0].text = f"{COMPANY} · внутренний документ"

    doc.add_paragraph(TITLE)
    doc.add_paragraph(COMPANY)
    doc.add_paragraph(f"Период: {PERIOD_PLACEHOLDER}")

    table = doc.add_table(rows=3, cols=5)
    table.style = "Table Grid"

    top = table.rows[0].cells[0].merge(table.rows[0].cells[4])
    top.text = TITLE

    for cell, text in zip(table.rows[1].cells, HEADER_CELLS):
        cell.text = text

    merged_client_product = table.rows[2].cells[1].merge(table.rows[2].cells[2])
    merged_client_product.text = "{{CLIENT}} / {{PRODUCT}}"
    table.rows[2].cells[0].text = "{{DATE}}"
    table.rows[2].cells[3].text = "{{QTY}}"
    table.rows[2].cells[4].text = "{{AMOUNT}}"

    doc.add_paragraph(f"Итого по периоду: {ROWS_PLACEHOLDER} позиций, {TOTAL_PLACEHOLDER} $")
    return doc


def doc_to_markdown(doc: Document, note: str) -> str:
    """Читаемый глазами близнец .docx — тот же текст абзацев и таблиц."""
    lines = [f"<!-- {note} -->", ""]
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text)
    for table in doc.tables:
        lines.append("")
        for row in table.rows:
            lines.append("| " + " | ".join(cell.text for cell in row.cells) + " |")
    return "\n".join(lines) + "\n"


def main() -> None:
    skeleton_dir = ROOT / "skill-skeleton" / "templates"
    naive_dir = ROOT / "demo-naive"
    break_dir = ROOT / "demo-break"
    for d in (skeleton_dir, naive_dir, break_dir):
        d.mkdir(parents=True, exist_ok=True)

    neutral = build_neutral_document()
    skeleton_docx = skeleton_dir / "report-template.docx"
    neutral.save(skeleton_docx)
    (skeleton_dir / "report-template.md").write_text(
        doc_to_markdown(
            neutral,
            "Данные вымышленные, компания «Меридиан» не существует. "
            "Плейлдеры не заполнены — так шаблон выглядит до fill_report.py.",
        ),
        encoding="utf-8",
    )

    # demo-naive использует тот же нейтральный шаблон, байт-в-байт, без скилла и скрипта.
    shutil.copyfile(skeleton_docx, naive_dir / "report-template.docx")
    shutil.copyfile(skeleton_dir / "report-template.md", naive_dir / "report-template.md")

    broken = build_break_document()
    break_docx = break_dir / "report-template-break.docx"
    broken.save(break_docx)
    (break_dir / "report-template-break.md").write_text(
        doc_to_markdown(
            broken,
            "Данные вымышленные, компания «Меридиан» не существует. "
            "Ломающийся шаблон: колонтитулы с полями компании, объединённые ячейки, "
            "двухуровневая шапка таблицы.",
        ),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
