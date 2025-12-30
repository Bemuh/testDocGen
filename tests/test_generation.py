from __future__ import annotations

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from case_resolver import build_output_filename, resolve_case
from document_generator import TestCaseDocumentConfig, FontConfig, create_test_case_document


def _docx_text(docx_path: Path, xml_name: str) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read(xml_name)
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    return "".join(t.text or "" for t in root.findall(".//w:t", ns))


def _find_template() -> Path:
    candidates = [
        p for p in Path(".").iterdir()
        if p.suffix.lower() == ".docx" and p.name.startswith("17731CP14")
    ]
    if not candidates:
        raise FileNotFoundError("No se encontrÇü el template DOCX de referencia en el repo.")
    return candidates[0]


def _make_cfg(template_path: Path, overrides=None) -> TestCaseDocumentConfig:
    return TestCaseDocumentConfig(
        header_image="header.png",
        footer_image="footer.png",
        project_name="GarantÇ­as Globales",
        analyst_name="Analista QA",
        date="23/12/2025",
        success_message="Resultado del caso de prueba: Ç%xito",
        font_config=FontConfig(
            table_font_name="Segoe UI", table_font_size=12,
            pasos_font_name="Segoe UI", pasos_font_size=11,
            step_font_name="Segoe UI", step_font_size=11,
            expected_font_name="Segoe UI", expected_font_size=11,
            exito_font_name="Segoe UI", exito_font_size=12,
        ),
        template_path=str(template_path),
        privacy_classification="DOCUMENTO PRIVADO",
        resultado="Exito",
        evidencia="CP14.mp4",
        version="001",
        case_overrides=overrides or {},
    )


def test_filename_builder_sample():
    cfg = _make_cfg(_find_template())
    resolved = resolve_case(
        case_id="17731",
        title="CP-14 – Validación mensajes de campos obligatorios",
        steps=[("A", "B")],
        cfg=cfg,
        source_text="Finagro garantias globales_11922 _ HU15 - AUT_HU035_REQ003",
    )
    expected = "17731CP14–ValidaciónMensajesDeCamposObligatorios_HU15_AUT_HU035_REQ003_V001"
    assert build_output_filename(resolved) == expected


def test_docx_contains_sections_and_overrides(tmp_path: Path):
    template_path = _find_template()
    overrides = {
        "17731": {
            "analista": "QA Editado",
            "fecha": "01/01/2026",
            "titulo": "CP-14 – Caso Editado",
            "evidencia": "video.mp4",
            "version": "002",
        }
    }
    cfg = _make_cfg(template_path, overrides=overrides)
    resolved = resolve_case(
        case_id="17731",
        title="CP-14 – Validación mensajes de campos obligatorios",
        steps=[("Paso 1", "Resultado 1"), ("Paso 2", "Resultado 2")],
        cfg=cfg,
        source_text="Finagro garantias globales_11922 _ HU15 - AUT_HU035_REQ003",
    )
    out_path = tmp_path / "out.docx"
    doc = create_test_case_document(resolved, cfg)
    doc.save(out_path)

    header_text = _docx_text(out_path, "word/header1.xml")
    body_text = _docx_text(out_path, "word/document.xml")

    assert "Casos De Prueba" in header_text
    assert "Proyecto Garant" in header_text
    assert "17731 HU15-AUT HU035 REQ003" in header_text
    assert "DOCUMENTO PRIVADO" in header_text

    assert "Formato Caso de Prueba" in body_text
    assert "Proyecto" in body_text
    assert "Nombre del Caso de Prueba" in body_text
    assert "Pasos" in body_text
    assert "Resultados del caso de prueba" in body_text
    assert "Evidencia" in body_text
    assert "Control De Acceso" in body_text
    assert "Control De Versiones" in body_text

    assert "QA Editado" in body_text
    assert "01/01/2026" in body_text
    assert "Paso 1" in body_text
    assert "Resultado 2" in body_text

    expected_name = "17731CP14–CasoEditado_HU15_AUT_HU035_REQ003_V002"
    assert build_output_filename(resolved) == expected_name
