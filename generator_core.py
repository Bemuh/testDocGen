from pathlib import Path
from processor import process_excel_files
from document_generator import TestCaseDocumentConfig


# def generate_docs(
#     excel_folder: Path,
#     config: TestCaseDocumentConfig,
#     dest_root: Path | None = None        # <── NUEVO parámetro opcional
# ) -> None:
#     """API que genera los documentos Word.

#     excel_folder : carpeta donde están los .xlsx
#     config       : configuración de encabezados/pie, etc.
#     dest_root    : carpeta base donde se crearán los .docx (por defecto: cwd)
#     """
#     process_excel_files(excel_folder, config, dest_root=dest_root)

def generate_docs(
    excel_folder: Path,
    config: TestCaseDocumentConfig,
    dest_root: Path | None = None,
    overwrite: bool = False,
    generate_pdf: bool = True,
    progress_callback=None,
):
    """Genera los documentos. Devuelve un GenerationReport (colisiones + contadores)."""
    return process_excel_files(
        excel_folder,
        config,
        dest_root=dest_root,
        overwrite=overwrite,
        generate_pdf=generate_pdf,
        progress_callback=progress_callback,
    )
