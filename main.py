import argparse
import datetime as _dt
import logging
from pathlib import Path

from processor import process_excel_files
from document_generator import TestCaseDocumentConfig, FontConfig

# ---------- configuración por defecto ----------
DEFAULT_FONT = FontConfig(
    table_font_name="Segoe UI", table_font_size=12,
    pasos_font_name="Segoe UI",  pasos_font_size=11,
    step_font_name="Segoe UI",   step_font_size=11,
    expected_font_name="Segoe UI", expected_font_size=11,
    exito_font_name="Segoe UI", exito_font_size=12,
)


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="generar-casos-prueba",
        description="Genera documentos Word a partir de suites de Excel exportadas de Azure DevOps."
    )
    p.add_argument(
        "excel_folder", nargs="?", default=".",
        help="Carpeta que contiene los archivos .xlsx (por defecto: carpeta actual)"
    )
    p.add_argument("--project", default="Nombre del Proyecto",
                   help="Nombre del proyecto que aparecerá en la tabla de encabezado")
    p.add_argument("--analyst", default="Nombre del Analista de Calidad",
                   help="Nombre de la persona analista")
    p.add_argument("--date", default=_dt.date.today().strftime("%d/%m/%Y"),
                   help="Fecha que se mostrará en cada documento (por defecto: hoy)")
    p.add_argument("--header-img", default="header.png",
                   help="Imagen de cabecera (opcional)")
    p.add_argument("--footer-img", default="footer.png",
                   help="Imagen de pie de página (opcional)")
    p.add_argument(
        "--verbose", "-v", action="count", default=1,
        help="Nivel de verbosidad log: 0 = solo errores, 1 = INFO (por defecto), 2 = DEBUG"
    )
    return p


def _configure_logging(verbosity: int) -> None:
    level = logging.ERROR if verbosity == 0 else \
            logging.INFO  if verbosity == 1 else \
            logging.DEBUG
    logging.basicConfig(format="%(levelname)s: %(message)s", level=level)


def main() -> None:
    args = _build_parser().parse_args()
    _configure_logging(args.verbose)

    cfg = TestCaseDocumentConfig(
        header_image=args.header_img,
        footer_image=args.footer_img,
        project_name=args.project,
        analyst_name=args.analyst,
        date=args.date,
        success_message="Resultado del caso de prueba: Éxito",
        font_config=DEFAULT_FONT,
    )

    logging.info("Iniciando generación en la carpeta: %s", args.excel_folder)
    process_excel_files(Path(args.excel_folder), cfg)
    logging.info("Generación finalizada con éxito 🎉")


if __name__ == "__main__":
    main()