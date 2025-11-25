import logging
from pathlib import Path
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_COLOR_INDEX, WD_PARAGRAPH_ALIGNMENT
import sys
from pathlib import Path

def resource_path(relpath: str | Path) -> Path:
    """Devuelve la ruta absoluta correcta dentro o fuera de PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path.cwd()))
    return (base / relpath).resolve()

class FontConfig:
    """Configuración tipográfica."""
    def __init__(self, table_font_name, table_font_size, pasos_font_name, pasos_font_size,
                 step_font_name, step_font_size, expected_font_name, expected_font_size,
                 exito_font_name, exito_font_size):
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
    """Configuración del documento de caso de prueba."""
    def __init__(self, header_image, footer_image, project_name, analyst_name, date, success_message, font_config):
        self.header_image = header_image
        self.footer_image = footer_image
        self.project_name = project_name
        self.analyst_name = analyst_name
        self.date = date
        self.success_message = success_message
        self.font_config = font_config
        self.user_story = ""  # se setea dinámicamente


def _apply_font(run, name: str, size: int, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold


# def _add_centered_image(container, image_path: Path, width_inches: float):
#     if not image_path.exists():
#         logging.warning("La imagen '%s' no se encontró; se omitirá.", image_path)
#         return
#     para = container.paragraphs[0]
#     para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
#     run = para.add_run()
#     run.add_picture(str(image_path), width=Inches(width_inches))

def _add_centered_image(container, image_ref, width_inches: float):
    img_path = resource_path(image_ref)

    if not img_path.exists():
        logging.warning("La imagen '%s' no se encontró; se omitirá.", img_path)
        return

    para = container.paragraphs[0]
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = para.add_run()
    run.add_picture(str(img_path), width=Inches(width_inches))


def create_test_case_document(case_num, title, steps, cfg):
    title = str(title) if pd.notna(title) else "Sin título"
    case_num = str(int(float(case_num)))  # 1.0 → '1'

    doc = Document()

    _add_centered_image(doc.sections[0].header, Path(cfg.header_image), 6)
    _add_centered_image(doc.sections[0].footer, Path(cfg.footer_image), 6)

    # ----- tabla de información -----
    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    etiquetas = ["Proyecto", "Nombre del Caso de Prueba", "Analista QA", "Fecha"]
    valores = [cfg.project_name, title, cfg.analyst_name, cfg.date]

    for fila, (etq, val) in enumerate(zip(etiquetas, valores)):
        table.cell(fila, 0).text = etq
        table.cell(fila, 1).text = val

    for col in range(2):
        for cell in table.column_cells(col):
            run = cell.paragraphs[0].runs[0]
            _apply_font(run, cfg.font_config.table_font_name, cfg.font_config.table_font_size, bold=(col == 0))

    # ajustar anchos
    for cell in table.column_cells(0):
        cell.width = Inches(1.57)
    for cell in table.column_cells(1):
        cell.width = Inches(4.43)

    doc.add_paragraph()  # espacio

    # ----- pasos -----
    pasos_p = doc.add_paragraph("Pasos:")
    _apply_font(pasos_p.runs[0], cfg.font_config.pasos_font_name, cfg.font_config.pasos_font_size, bold=True)

    for action, expected in steps:
        p = doc.add_paragraph(style="List Number")
        _apply_font(p.add_run(action.strip()), cfg.font_config.step_font_name, cfg.font_config.step_font_size, bold=True)
        if expected:
            exp = doc.add_paragraph()
            _apply_font(exp.add_run(expected.strip()),
                        cfg.font_config.expected_font_name, cfg.font_config.expected_font_size)

    # ----- mensaje de éxito -----
    ok_p = doc.add_paragraph()
    ok_run = ok_p.add_run(f"\n{cfg.success_message}\n")
    _apply_font(ok_run, cfg.font_config.exito_font_name, cfg.font_config.exito_font_size, bold=True)
    ok_run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN

    return doc