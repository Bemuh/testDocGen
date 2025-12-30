import logging
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from case_resolver import ResolvedCase, Step


def resource_path(relpath: str | Path) -> Path:
    """Devuelve la ruta absoluta correcta dentro o fuera de PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path.cwd()))
    return (base / relpath).resolve()


class FontConfig:
    """ConfiguraciÇün tipogrÇ­fica."""
    def __init__(
        self,
        table_font_name,
        table_font_size,
        pasos_font_name,
        pasos_font_size,
        step_font_name,
        step_font_size,
        expected_font_name,
        expected_font_size,
        exito_font_name,
        exito_font_size,
    ):
        self.table_font_name = table_font_name
        self.table_font_size = table_font_size
        self.pasos_font_name = pasos_font_name
        self.pasos_font_size = pasos_font_size
        self.step_font_name = step_font_name
        self.step_font_size = step_font_size
        self.expected_font_name = expected_font_name
        self.expected_font_size = expected_font_size
        self.exito_font_name = exito_font_name
        self.exito_font_size = exito_font_size


class TestCaseDocumentConfig:
    """ConfiguraciÇün del documento de caso de prueba."""
    def __init__(
        self,
        header_image,
        footer_image,
        project_name,
        analyst_name,
        date,
        success_message,
        font_config,
        template_path=None,
        privacy_classification="DOCUMENTO PRIVADO",
        resultado="Exito",
        evidencia="",
        version="001",
        case_overrides=None,
    ):
        self.header_image = header_image
        self.footer_image = footer_image
        self.project_name = project_name
        self.analyst_name = analyst_name
        self.date = date
        self.success_message = success_message
        self.font_config = font_config
        self.template_path = template_path
        self.privacy_classification = privacy_classification
        self.resultado = resultado
        self.evidencia = evidencia
        self.version = version
        self.case_overrides = case_overrides or {}
        self.user_story = ""  # se setea dinÇ­micamente


def _apply_font(run, name: str, size: int, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold


def _add_centered_image(container, image_ref, width_inches: float):
    img_path = resource_path(image_ref)

    if not img_path.exists():
        logging.warning("La imagen '%s' no se encontrÇü; se omitirÇ­.", img_path)
        return

    para = container.paragraphs[0]
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = para.add_run()
    run.add_picture(str(img_path), width=Inches(width_inches))


def _load_template(cfg: TestCaseDocumentConfig) -> Document:
    if cfg.template_path:
        template_path = resource_path(cfg.template_path)
        if template_path.exists():
            return Document(str(template_path))
        logging.warning("No se encontrÇü el template '%s'; se crearÇ­ un documento nuevo.", template_path)

    default_template = resource_path(
        "17731CP14–ValidaciónMensajesDeCamposObligatorios_HU15_AUT_HU035_REQ003_V001.docx"
    )
    if default_template.exists():
        return Document(str(default_template))
    return Document()


def _set_paragraph_text(paragraph, text: str) -> None:
    paragraph.text = text


def _set_cell_lines(cell, lines: list[str]) -> None:
    if not lines:
        cell.text = ""
        return
    cell.text = lines[0]
    for line in lines[1:]:
        cell.add_paragraph(line)


def _clear_cell(cell) -> None:
    for paragraph in cell.paragraphs:
        paragraph.text = ""


def _set_cell_label(cell, text: str) -> None:
    _clear_cell(cell)
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.bold = True


def _set_cell_value(cell, text: str) -> None:
    _clear_cell(cell)
    cell.paragraphs[0].add_run(text)


def _populate_header(doc: Document, resolved: ResolvedCase) -> None:
    header = doc.sections[0].header
    for paragraph in header.paragraphs:
        paragraph.text = ""
    if header.tables:
        left_cell = header.tables[0].cell(0, 0)
        for paragraph in left_cell.paragraphs:
            paragraph.text = ""
        lines = [
            "Casos De Prueba:",
            f"Proyecto {resolved.project}",
            resolved.header_line,
            resolved.privacidad,
        ]
        for idx, line in enumerate(lines):
            if idx < len(left_cell.paragraphs):
                _set_paragraph_text(left_cell.paragraphs[idx], line)
            else:
                left_cell.add_paragraph(line)
        return

    if header.paragraphs:
        header.paragraphs[0].text = "\n".join(
            [
                "Casos De Prueba:",
                f"Proyecto {resolved.project}",
                resolved.header_line,
                resolved.privacidad,
            ]
        )


def _find_main_table(doc: Document):
    for table in doc.tables:
        content = " ".join(cell.text for row in table.rows for cell in row.cells)
        has_project = "Proyecto" in content or "Nombre del Proyecto" in content
        has_case = "Nombre del Caso de Prueba" in content or "CP-" in content
        if has_project and has_case and len(table.rows) >= 5:
            return table
    return doc.tables[0] if doc.tables else None


def _ensure_row_count(table, count: int) -> None:
    while len(table.rows) < count:
        table.add_row()


def _remove_row(table, row_idx: int) -> None:
    table._tbl.remove(table.rows[row_idx]._tr)


def _clone_row(row):
    return deepcopy(row._tr)


def _populate_steps(table, pasos_row_idx: int, result_row_idx: int, steps: list[Step]):
    if pasos_row_idx is None or result_row_idx is None or pasos_row_idx >= result_row_idx:
        return

    template_row = table.rows[pasos_row_idx + 1] if pasos_row_idx + 1 < len(table.rows) else None
    for idx in range(result_row_idx - 1, pasos_row_idx, -1):
        _remove_row(table, idx)

    insert_before = table.rows[pasos_row_idx + 1] if pasos_row_idx + 1 < len(table.rows) else table.rows[-1]
    for step_index, step in enumerate(steps, start=1):
        if template_row is None:
            new_row = table.add_row()
        else:
            new_tr = _clone_row(template_row)
            insert_before._tr.addprevious(new_tr)
            new_row = table.rows[pasos_row_idx + step_index]

        cells = new_row.cells
        if cells:
            cells[0].text = ""
            run = cells[0].paragraphs[0].add_run(str(step_index))
            run.bold = True
        if len(cells) > 1:
            action_text = step.action.strip()
            expected_text = step.expected.strip()
            cells[1].text = ""
            action_paragraph = cells[1].paragraphs[0]
            action_run = action_paragraph.add_run(action_text)
            action_run.bold = True
            if expected_text:
                cells[1].add_paragraph(expected_text)


def _pick_target_table(doc: Document):
    if not doc.tables:
        return None
    for table in doc.tables:
        if not table.rows or not table.rows[0].cells:
            continue
        head = table.rows[0].cells[0].text
        if "Nombre del Proyecto" in head or "Proyecto" in head:
            return table
    return doc.tables[0]


def _populate_main_table(doc: Document, resolved: ResolvedCase) -> None:
    table = _pick_target_table(doc)
    if table is None:
        return

    pasos_row_idx = None
    result_row_idx = None
    evidencia_row_idx = None

    _ensure_row_count(table, 5)
    row0 = table.rows[0]
    _set_cell_label(row0.cells[0], "Proyecto")
    _set_cell_value(row0.cells[1], resolved.project)

    row1 = table.rows[1]
    _set_cell_label(row1.cells[0], "ID caso de prueba")
    _set_cell_value(row1.cells[1], resolved.case_id)
    if len(row1.cells) >= 4:
        _set_cell_label(row1.cells[2], "Historia de usuario")
        _set_cell_value(row1.cells[3], resolved.historia_usuario)

    row2 = table.rows[2]
    _set_cell_label(row2.cells[0], "Nombre del Caso de Prueba")
    _set_cell_value(row2.cells[1], resolved.title)

    row3 = table.rows[3]
    _set_cell_label(row3.cells[0], "Analista QA")
    _set_cell_value(row3.cells[1], resolved.analyst)

    row4 = table.rows[4]
    _set_cell_label(row4.cells[0], "Fecha")
    _set_cell_value(row4.cells[1], resolved.date)

    for idx, row in enumerate(table.rows):
        cells = [cell.text.strip() for cell in row.cells]
        if not cells:
            continue
        label = cells[0]
        label_norm = " ".join(label.replace("\xa0", " ").split()).lower()
        if label_norm == "pasos":
            pasos_row_idx = idx
        elif label_norm.startswith("resultados del caso de prueba"):
            result_row_idx = idx
            if len(row.cells) > 1:
                _set_cell_label(row.cells[0], "Resultados del caso de prueba")
                _set_cell_value(row.cells[1], resolved.resultado)
        elif label_norm == "evidencia":
            evidencia_row_idx = idx
            if len(row.cells) > 1:
                _set_cell_label(row.cells[0], "Evidencia")
                _set_cell_value(row.cells[1], resolved.evidencia)

    _populate_steps(table, pasos_row_idx, result_row_idx, resolved.steps)

    if result_row_idx is None:
        for idx, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) >= 2 and cells[0].strip().lower() == "exito":
                result_row_idx = idx
                break
    if result_row_idx is None:
        _ensure_row_count(table, len(table.rows) + 1)
        result_row_idx = len(table.rows) - 1
    if evidencia_row_idx is None:
        _ensure_row_count(table, result_row_idx + 2)
        evidencia_row_idx = result_row_idx + 1

    _ensure_row_count(table, result_row_idx + 1)
    row = table.rows[result_row_idx]
    if len(row.cells) > 1:
        _set_cell_label(row.cells[0], "Resultados del caso de prueba")
        _set_cell_value(row.cells[1], resolved.resultado)

    _ensure_row_count(table, evidencia_row_idx + 1)
    row = table.rows[evidencia_row_idx]
    if len(row.cells) > 1:
        _set_cell_label(row.cells[0], "Evidencia")
        _set_cell_value(row.cells[1], "")


def create_test_case_document(resolved: ResolvedCase, cfg: TestCaseDocumentConfig) -> Document:
    doc = _load_template(cfg)
    _populate_header(doc, resolved)
    _populate_main_table(doc, resolved)
    return doc
