from __future__ import annotations

from dataclasses import dataclass
import datetime as _dt
import logging
import re
from typing import Iterable, Mapping


@dataclass(frozen=True)
class Step:
    action: str
    expected: str


@dataclass(frozen=True)
class ResolvedCase:
    case_id: str
    cp_num: str
    title: str
    title_slug: str
    user_story_name: str
    project: str
    analyst: str
    date: str
    date_compact: str
    date_slash: str
    historia_usuario: str
    hu_primary: str
    hu_secondary: str
    req: str
    version: str
    steps: list[Step]
    resultado: str
    evidencia: str
    privacidad: str

    @property
    def header_line(self) -> str:
        return (
            f"{self.user_story_name}"
        )


_RE_CP_PREFIX = re.compile(r"^\s*CP\s*[- ]?\s*(\d+)\s*[–\-:]*\s*(.*)$", re.IGNORECASE)
_RE_HU = re.compile(r"HU\s*([0-9]+)", re.IGNORECASE)
_RE_REQ = re.compile(r"REQ\s*([0-9]+)", re.IGNORECASE)
_INVALID_CHARS = re.compile(r'[\\/*?:"<>|]')


def _normalize_case_id(case_id: str | int | float) -> str:
    if isinstance(case_id, (int, float)):
        return str(int(case_id))
    text = str(case_id).strip()
    if text.endswith(".0") and text.replace(".", "", 1).isdigit():
        return text[:-2]
    return text


def _normalize_cp_number(title: str, override: str | None = None) -> str:
    if override:
        return re.sub(r"\D", "", str(override)) or str(override).strip()
    match = _RE_CP_PREFIX.match(title or "")
    if match:
        return match.group(1)
    return ""


def _strip_cp_prefix(title: str) -> str:
    match = _RE_CP_PREFIX.match(title or "")
    if match and match.group(2):
        return match.group(2).strip()
    return title.strip() if title else ""


def _collapse_title_for_filename(title: str) -> str:
    cleaned = []
    for ch in _strip_cp_prefix(title):
        if ch.isalnum():
            cleaned.append(ch)
    return "".join(cleaned)


def _extract_tokens(source_text: str) -> tuple[str, str, str]:
    hu_tokens = _RE_HU.findall(source_text or "")
    req_match = _RE_REQ.search(source_text or "")
    hu_primary = hu_tokens[0] if len(hu_tokens) > 0 else ""
    hu_secondary = hu_tokens[1] if len(hu_tokens) > 1 else ""
    req = req_match.group(1) if req_match else ""
    return hu_primary, hu_secondary, req


def _sanitize_project(text: str) -> str:
    return _INVALID_CHARS.sub("", text or "").strip()


def _parse_date_compact(date_str: str) -> str:
    """Convierte la fecha de la UI a YYYYMMDD.

    Acepta DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD, YYYYMMDD. Si nada cuadra,
    intenta extraer 8 dígitos consecutivos; en último caso usa hoy."""
    text = (date_str or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y%m%d"):
        try:
            return _dt.datetime.strptime(text, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    digits = re.sub(r"\D", "", text)
    if len(digits) == 8:
        return digits
    fallback = _dt.date.today().strftime("%Y%m%d")
    logging.warning("Fecha '%s' no se pudo parsear; se usa la de hoy (%s).", date_str, fallback)
    return fallback


def resolve_case(
    case_id: str | int | float,
    title: str,
    steps: Iterable[tuple[str, str]] | Iterable[Step],
    cfg,
    source_text: str,
) -> ResolvedCase:
    case_id_norm = _normalize_case_id(case_id)
    overrides: Mapping[str, object] = {}
    if getattr(cfg, "case_overrides", None):
        overrides = cfg.case_overrides.get(case_id_norm, {})

    project = overrides.get("proyecto", cfg.project_name)
    analyst = overrides.get("analista", cfg.analyst_name)
    date = overrides.get("fecha", cfg.date)
    privacy = overrides.get("privacidad", cfg.privacy_classification)
    resultado = overrides.get("resultado", cfg.resultado)
    evidencia = overrides.get("evidencia", cfg.evidencia)
    version = overrides.get("version", cfg.version or "001")

    title_final = overrides.get("titulo", title)
    cp_num = _normalize_cp_number(title_final, overrides.get("cp_num"))
    title_slug = _collapse_title_for_filename(title_final)

    source_for_tokens = overrides.get("historia_origen", source_text) or source_text or getattr(cfg, "user_story", "")
    hu_primary, hu_secondary, req = _extract_tokens(str(source_for_tokens))
    hu_primary = overrides.get("hu", hu_primary) or ""
    hu_secondary = overrides.get("hu_secundaria", hu_secondary) or ""
    req = overrides.get("req", req) or ""

    # UserStoryName = stem del Excel tal cual, sólo se quitan chars inválidos de Windows
    raw_excel_name = str(source_for_tokens or "").strip()
    user_story_name = _INVALID_CHARS.sub("", raw_excel_name).strip()

    historia_usuario = overrides.get("historia_usuario", user_story_name)
    if not historia_usuario:
        historia_usuario = getattr(cfg, "user_story", "")

    resolved_steps: list[Step] = []
    override_steps = overrides.get("pasos")
    if override_steps is not None:
        for step in override_steps:
            resolved_steps.append(Step(str(step["accion"]), str(step.get("esperado", ""))))
    else:
        for action, expected in steps:
            resolved_steps.append(Step(str(action or ""), str(expected or "")))

    date_str = str(date).strip()
    date_compact = _parse_date_compact(date_str)
    date_slash = f"{date_compact[:4]}/{date_compact[4:6]}/{date_compact[6:8]}"
    return ResolvedCase(
        case_id=case_id_norm,
        cp_num=cp_num,
        title=str(title_final).strip(),
        title_slug=title_slug,
        user_story_name=user_story_name,
        project=_sanitize_project(str(project)),
        analyst=str(analyst).strip(),
        date=date_str,
        date_compact=date_compact,
        date_slash=date_slash,
        historia_usuario=str(historia_usuario).strip(),
        hu_primary=str(hu_primary),
        hu_secondary=str(hu_secondary),
        req=str(req),
        version=str(version).strip() or "001",
        steps=resolved_steps,
        resultado=str(resultado).strip(),
        evidencia=str(evidencia).strip(),
        privacidad=str(privacy).strip() or "DOCUMENTO PRIVADO",
    )


def build_output_filename(resolved_case: ResolvedCase) -> str:
    cp_num = resolved_case.cp_num or ""
    user_story_name = resolved_case.user_story_name
    version = resolved_case.version or "001"
    return (
        f"{resolved_case.date_compact}_"
        f"{resolved_case.case_id}CP{cp_num}"
        f"_{user_story_name}_V{version}"
    )
