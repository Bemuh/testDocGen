import logging
import re
from pathlib import Path
from typing import Final

from data_loader import load_test_cases_from_excel
from document_generator import create_test_case_document, TestCaseDocumentConfig

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
    overwrite: bool = False          # <-- NUEVO
) -> list[Path]:                     # <-- devolvemos las colisiones
    dest_root = dest_root or Path.cwd()
    collisions: list[Path] = []

    for excel_path in folder.glob("*.xlsx"):
        if excel_path.name.startswith("~$"):
            continue
        plan_dir = dest_root / _extract_testplan_parts(excel_path.stem, max_folder_len)
        plan_dir.mkdir(parents=True, exist_ok=True)
        cfg.user_story = plan_dir.name

        try:
            cases = load_test_cases_from_excel(excel_path)
        except Exception as exc:      # noqa: BLE001
            logging.error("No se pudo leer '%s' (%s). Se omite.", excel_path.name, exc)
            continue

        for case_id, data in cases.items():
            title_slug = _slugify(data["title"], max_file_len)
            out_file = plan_dir / f"{case_id} {title_slug}.docx"

            if out_file.exists() and not overwrite:
                collisions.append(out_file)
                continue

            doc = create_test_case_document(case_id, data["title"], data["steps"], cfg)
            doc.save(out_file)
            logging.info("   ✔ %s", out_file.relative_to(dest_root))

    return collisions

def regenerate_single(path: Path, cfg: TestCaseDocumentConfig) -> None:
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
    case_id = path.name.split(" ")[0]                    # "22933"
    if case_id not in cases:
        logging.warning("El caso %s no existe en %s", case_id, excel_file.name)
        return

    data = cases[case_id]
    doc = create_test_case_document(case_id, data["title"], data["steps"], cfg)
    doc.save(path)
    logging.info("   ↻ %s (sobrescrito)", path.relative_to(plan_dir.parent))