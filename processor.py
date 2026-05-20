import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Final, Optional

# Tipo del callback de progreso: (mensaje, actual, total)
ProgressCallback = Callable[[str, int, int], None]

from data_loader import load_test_cases_from_excel
from document_generator import create_test_case_document, TestCaseDocumentConfig
from case_resolver import build_output_filename, resolve_case
from pdf_converter import convert_to_pdf, init_pdf_session, close_pdf_session


@dataclass
class GenerationReport:
    collisions: list[Path] = field(default_factory=list)
    docx_count: int = 0
    pdf_count: int = 0
    pdf_failed: int = 0

    # Retrocompat: el código antiguo trataba el retorno como list[Path].
    def __iter__(self):
        return iter(self.collisions)

    def __len__(self):
        return len(self.collisions)

    def __bool__(self):
        return bool(self.collisions)

_RE_PLAN: Final = re.compile(r".* - Sprint \d+_(\d+)\s*_?\s*(.+)$")
INVALID_CHARS: Final = re.compile(r'[\\/*?:"<>|]')


def _slugify(text: str, limit: int) -> str:
    text = INVALID_CHARS.sub("_", text.strip())
    return (text if len(text) <= limit else text[:limit].rsplit(" ", 1)[0] or text[:limit]).strip()


def _extract_testplan_parts(filename: str, max_len: int = 100) -> str:
    match = _RE_PLAN.match(filename)
    if not match:
        logging.warning("El nombre de archivo '%s' no sigue la convención; se usará completo", filename)
        return _slugify(filename, max_len)
    number, name = match.groups()
    return f"{number}_{_slugify(name, max_len)}"


# def process_excel_files(
#     folder: Path,
#     cfg: TestCaseDocumentConfig,
#     dest_root: Path | None = None,
#     max_folder_len: int = 40,
#     max_file_len: int = 100
# ) -> None:
#     """Genera los documentos en dest_root / <plan> (por defecto: carpeta actual)."""
#     dest_root = dest_root or Path.cwd()

#     for excel_path in folder.glob("*.xlsx"):
#         if excel_path.name.startswith("~$"):
#             continue

#         base_name = excel_path.stem
#         plan_dir_name = _extract_testplan_parts(base_name, max_folder_len)
#         plan_dir = dest_root / plan_dir_name            # <── ahora bajo dest_root
#         plan_dir.mkdir(parents=True, exist_ok=True)

#         cfg.user_story = plan_dir_name
#         logging.info("▶ Procesando plan '%s' (%s)", cfg.user_story, excel_path.name)

#         try:
#             cases = load_test_cases_from_excel(excel_path)
#         except Exception as exc:                        # noqa: BLE001
#             logging.error("No se pudo leer '%s' (%s). Se omite.", excel_path.name, exc)
#             continue

#         for case_id, data in cases.items():
#             title_slug = _slugify(data["title"], max_file_len)
#             out_file = plan_dir / f"{case_id} {title_slug}.docx"

#             if out_file.exists():
#                 logging.debug("Ya existe %s; se omite.", out_file.name)
#                 continue

#             doc = create_test_case_document(case_id, data["title"], data["steps"], cfg)
#             doc.save(out_file)
#             logging.info("   ✔ %s", out_file.relative_to(dest_root))

def process_excel_files(
    folder: Path,
    cfg: TestCaseDocumentConfig,
    dest_root: Path | None = None,
    max_folder_len: int = 40,
    max_file_len: int = 100,
    overwrite: bool = False,
    generate_pdf: bool = True,
    progress_callback: Optional[ProgressCallback] = None,
) -> GenerationReport:
    def _progress(msg: str, current: int = 0, total: int = 0) -> None:
        if progress_callback is not None:
            try:
                progress_callback(msg, current, total)
            except Exception:  # noqa: BLE001
                logging.debug("progress_callback levantó excepción; se ignora.")

    dest_root = dest_root or Path.cwd()
    report = GenerationReport()

    folder_abs = folder.resolve() if folder else Path.cwd()
    _progress(f"Buscando .xlsx en: {folder_abs}", 0, 1)
    logging.info("Buscando .xlsx en: %s", folder_abs)
    if not folder_abs.exists():
        logging.error("La carpeta origen NO EXISTE: %s", folder_abs)
        _progress("La carpeta origen no existe.", 0, 0)
        return report

    xlsx_files = [p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")]
    logging.info("Encontrados %d archivos .xlsx", len(xlsx_files))
    if not xlsx_files:
        logging.warning(
            "Ningún .xlsx encontrado. Verifica la 'Carpeta origen' en la GUI; "
            "no se admiten subcarpetas."
        )
        _progress("No se encontró ningún .xlsx.", 0, 0)
        return report

    # Pre-cargar casos para conocer el total y poder mostrar % preciso
    _progress(f"Leyendo {len(xlsx_files)} Excel(s)...", 0, len(xlsx_files))
    all_cases: list[tuple[Path, dict]] = []
    for i, excel_path in enumerate(xlsx_files, start=1):
        _progress(f"Leyendo Excel {i}/{len(xlsx_files)}: {excel_path.name}", i, len(xlsx_files))
        try:
            cases = load_test_cases_from_excel(excel_path)
            all_cases.append((excel_path, cases))
        except Exception as exc:  # noqa: BLE001
            logging.error("No se pudo leer '%s' (%s). Se omite.", excel_path.name, exc)

    total_cases = sum(len(cases) for _, cases in all_cases)
    if total_cases == 0:
        _progress("No hay casos de prueba en los Excel encontrados.", 0, 0)
        return report

    _progress(f"Iniciando: {total_cases} casos en {len(all_cases)} Excel(s)...", 0, total_cases)

    if generate_pdf:
        _progress("Abriendo Microsoft Word...", 0, total_cases)
        init_pdf_session()

    try:
        _process_xlsx_loop(
            all_cases, cfg, dest_root, max_folder_len, overwrite,
            generate_pdf, report, _progress, total_cases,
        )
    finally:
        if generate_pdf:
            _progress("Cerrando Microsoft Word...", total_cases, total_cases)
            close_pdf_session()

    summary = f"Listo: {report.docx_count} DOCX"
    if generate_pdf:
        summary += f", {report.pdf_count} PDF"
        if report.pdf_failed:
            summary += f" ({report.pdf_failed} fallidos)"
    _progress(summary, total_cases, total_cases)

    return report


def _process_xlsx_loop(
    all_cases, cfg, dest_root, max_folder_len, overwrite,
    generate_pdf, report, progress, total_cases,
):
    processed = 0
    for excel_path, cases in all_cases:
        logging.info("Procesando Excel: %s", excel_path.name)
        plan_dir = dest_root / _extract_testplan_parts(excel_path.stem, max_folder_len)
        plan_dir.mkdir(parents=True, exist_ok=True)
        cfg.user_story = plan_dir.name

        for case_id, data in cases.items():
            processed += 1
            resolved = resolve_case(case_id, data["title"], data["steps"], cfg, excel_path.stem)
            filename = build_output_filename(resolved)
            out_file = plan_dir / f"{filename}.docx"

            progress(f"DOCX {processed}/{total_cases}: caso {case_id}", processed, total_cases)

            if out_file.exists() and not overwrite:
                report.collisions.append(out_file)
                continue

            doc = create_test_case_document(resolved, cfg)
            doc.save(out_file)
            report.docx_count += 1
            logging.info("   ✔ %s", out_file.relative_to(dest_root))

            if generate_pdf:
                progress(f"PDF  {processed}/{total_cases}: caso {case_id}", processed, total_cases)
                pdf_out = convert_to_pdf(out_file)
                if pdf_out:
                    report.pdf_count += 1
                    logging.info("   📄 %s", pdf_out.relative_to(dest_root))
                else:
                    report.pdf_failed += 1
                    logging.warning("   ⚠ No se generó PDF para %s", out_file.name)

    return report

def regenerate_single(path: Path, cfg: TestCaseDocumentConfig, generate_pdf: bool = True) -> None:
    """
    Regenera un único .docx (usado cuando el usuario acepta sobrescribir).
    Recibe la ruta completa al archivo existente.
    """
    # path = dest_root / <plan_dir> / "<ID> <slug>.docx"
    plan_dir = path.parent
    excel_id = plan_dir.name.split("_")[0]               # 18395, 20683, …
    excel_file = next(plan_dir.parent.glob(f"*_{excel_id}*.xlsx"), None)
    if not excel_file:
        logging.warning("No se encontró el Excel original para %s", path.name)
        return

    cases = load_test_cases_from_excel(excel_file)
    match = re.match(r"(\\d+)", path.name)
    case_id = match.group(1) if match else ""
    if case_id not in cases:
        logging.warning("El caso %s no existe en %s", case_id, excel_file.name)
        return

    data = cases[case_id]
    resolved = resolve_case(case_id, data["title"], data["steps"], cfg, excel_file.stem)
    doc = create_test_case_document(resolved, cfg)
    doc.save(path)
    logging.info("   ↻ %s (sobrescrito)", path.relative_to(plan_dir.parent))

    if generate_pdf:
        pdf_out = convert_to_pdf(path)
        if pdf_out:
            logging.info("   📄 %s", pdf_out.relative_to(plan_dir.parent))
        else:
            logging.warning("   ⚠ No se generó PDF para %s", path.name)
