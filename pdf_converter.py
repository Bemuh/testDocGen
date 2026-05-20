"""Conversión de .docx a .pdf con auto-detección de motor.

Estrategia: probar los motores disponibles en orden y cachear el elegido.
Permite forzar el motor con la variable de entorno TESTDOCGEN_PDF_ENGINE
(valores: auto | word | libreoffice | none).

En Windows convertimos directamente vía Microsoft Word COM (sin docx2pdf)
para evitar el bug de tqdm cuando el ejecutable está empaquetado con
console=False (sys.stdout/stderr = None hacen que tqdm reviente con
"'NoneType' object has no attribute 'write'").
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from typing import Callable, Optional

_engine_cache: Optional[str] = None  # "word" | "libreoffice" | "none"
_soffice_path: Optional[str] = None

# Sesión Word reutilizada durante un lote de conversiones (mucho más rápido
# que abrir/cerrar Word para cada archivo).
_word_session = None
_word_session_lock = threading.Lock()
_word_session_thread_id: Optional[int] = None


def _find_soffice() -> Optional[str]:
    global _soffice_path
    if _soffice_path is not None:
        return _soffice_path or None

    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if not found:
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
        for c in candidates:
            if Path(c).exists():
                found = c
                break

    _soffice_path = found or ""
    return found


def _try_word() -> bool:
    """Verifica si tenemos lo necesario para Word/COM en Windows o
    docx2pdf en macOS."""
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # noqa: F401
            import pythoncom  # noqa: F401
        except ImportError:
            return False
        return True
    if sys.platform == "darwin":
        try:
            import docx2pdf  # noqa: F401
        except ImportError:
            return False
        return True
    return False


def init_pdf_session() -> None:
    """Abre una sesión Word reutilizable. Llamar al inicio de un lote.
    Idempotente: si ya está abierta no hace nada."""
    global _word_session, _word_session_thread_id
    if not sys.platform.startswith("win"):
        return
    if _select_engine() != "word":
        return
    with _word_session_lock:
        if _word_session is not None:
            return
        try:
            import pythoncom
            from win32com.client import DispatchEx
            pythoncom.CoInitialize()
            app = DispatchEx("Word.Application")
            app.Visible = False
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            _word_session = app
            _word_session_thread_id = threading.get_ident()
            logging.info("PDF: sesión Word inicializada (reutilizada para todo el lote).")
        except Exception as exc:  # noqa: BLE001
            logging.warning("PDF: no se pudo abrir Word (%s).", exc)
            _word_session = None
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass


def close_pdf_session() -> None:
    """Cierra la sesión Word. Llamar al final del lote."""
    global _word_session, _word_session_thread_id
    with _word_session_lock:
        if _word_session is None:
            return
        try:
            _word_session.Quit()
        except Exception as exc:  # noqa: BLE001
            logging.debug("PDF: error al cerrar Word (%s).", exc)
        _word_session = None
        _word_session_thread_id = None
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        logging.info("PDF: sesión Word cerrada.")


def _convert_word(docx_path: Path) -> Path:
    """Convierte vía Microsoft Word COM directamente.

    Si no hay sesión abierta crea una efímera (y la cierra al final).
    Si hay sesión abierta la reutiliza (mucho más rápido)."""
    pdf_path = docx_path.with_suffix(".pdf")

    if not sys.platform.startswith("win"):
        # macOS: usar docx2pdf (AppleScript)
        from docx2pdf import convert as _docx2pdf_convert
        _docx2pdf_convert(str(docx_path), str(pdf_path))
        if not pdf_path.exists():
            raise RuntimeError(f"docx2pdf no generó {pdf_path.name}")
        return pdf_path

    own_session = False
    if _word_session is None:
        init_pdf_session()
        own_session = True

    if _word_session is None:
        raise RuntimeError("No se pudo iniciar Microsoft Word vía COM")

    try:
        with _word_session_lock:
            doc = _word_session.Documents.Open(
                str(docx_path.resolve()), ReadOnly=True, Visible=False
            )
            try:
                # wdFormatPDF = 17
                doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
            finally:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
    finally:
        if own_session:
            close_pdf_session()

    if not pdf_path.exists():
        raise RuntimeError(f"Word no generó {pdf_path.name}")
    return pdf_path


def _convert_libreoffice(docx_path: Path) -> Path:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("soffice no encontrado")
    pdf_path = docx_path.with_suffix(".pdf")
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(docx_path.parent), str(docx_path)],
        check=True, timeout=120,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    if not pdf_path.exists():
        raise RuntimeError(f"LibreOffice no generó {pdf_path.name}")
    return pdf_path


def _select_engine() -> str:
    """Decide qué motor usar (con caché). Respeta TESTDOCGEN_PDF_ENGINE."""
    global _engine_cache
    if _engine_cache is not None:
        return _engine_cache

    forced = (os.environ.get("TESTDOCGEN_PDF_ENGINE", "auto") or "auto").strip().lower()
    if forced == "none":
        _engine_cache = "none"
        logging.info("PDF: motor forzado a 'none' por TESTDOCGEN_PDF_ENGINE; se omite generación de PDF.")
        return _engine_cache
    if forced == "word":
        if _try_word():
            _engine_cache = "word"
        else:
            logging.warning("PDF: 'word' forzado pero Word/COM no está disponible; se omite.")
            _engine_cache = "none"
        return _engine_cache
    if forced == "libreoffice":
        if _find_soffice():
            _engine_cache = "libreoffice"
        else:
            logging.warning("PDF: 'libreoffice' forzado pero soffice no se encontró; se omite.")
            _engine_cache = "none"
        return _engine_cache

    # auto: probar Word primero (mayor fidelidad), luego LibreOffice
    if _try_word():
        _engine_cache = "word"
    elif _find_soffice():
        _engine_cache = "libreoffice"
    else:
        _engine_cache = "none"
        logging.warning(
            "PDF: no se encontró Microsoft Word ni LibreOffice; se omite la generación de PDF."
        )

    if _engine_cache != "none":
        logging.info("PDF: motor seleccionado = %s", _engine_cache)
    return _engine_cache


_CONVERTERS: dict[str, Callable[[Path], Path]] = {
    "word": _convert_word,
    "libreoffice": _convert_libreoffice,
}


def convert_to_pdf(docx_path: Path) -> Optional[Path]:
    """Convierte docx_path a PDF con el primer motor disponible.

    Devuelve la ruta del PDF o None si ningún motor pudo o falló."""
    engine = _select_engine()
    if engine == "none":
        return None

    converter = _CONVERTERS.get(engine)
    if not converter:
        return None

    try:
        return converter(docx_path)
    except Exception as exc:  # noqa: BLE001
        logging.warning("PDF: motor '%s' falló para %s (%s).", engine, docx_path.name, exc)
        # Fallback: si Word falló, probar LibreOffice una sola vez
        if engine == "word" and _find_soffice():
            try:
                logging.info("PDF: reintentando con LibreOffice...")
                result = _convert_libreoffice(docx_path)
                global _engine_cache
                _engine_cache = "libreoffice"
                return result
            except Exception as exc2:  # noqa: BLE001
                logging.warning("PDF: LibreOffice también falló (%s).", exc2)
        return None
