from __future__ import annotations

import logging
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import datetime
from enum import Enum, auto
from functools import lru_cache
from pathlib import Path
from typing import Any, Optional
from xml.sax.saxutils import escape

logger = logging.getLogger(__name__)

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
except ImportError:
    Document = None
    logger.warning(
        "Moduł 'python-docx' nie jest zainstalowany - eksport protokołów do Worda jest niedostępny."
    )

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFError, TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
except ImportError:
    SimpleDocTemplate = None
    logger.warning(
        "Moduł 'reportlab' nie jest zainstalowany - eksport protokołów do PDF jest niedostępny."
    )

# ---------------------------------------------------------------------------
# Stałe
# ---------------------------------------------------------------------------

PROTOCOL_HANDOVER = "przekazanie"
PROTOCOL_RETURN = "zwrot"

DATE_FORMAT = "%d.%m.%Y"

_IDX_MODEL = 0
_IDX_NR_TEL = 1
_IDX_NR_SIM = 2
_IDX_IMEI = 3
_IDX_SERIAL = 4
_IDX_USER = 6
_IDX_EQUIPMENT = 9
_IDX_BATTERY = 10

_CHARGER_KEYWORDS = ("ładowarka", "ladowarka")
_CABLE_KEYWORDS = ("kabel", "usb")
_CASE_KEYWORDS = ("etui",)
_GLASS_KEYWORDS = ("szkło", "szklo")
_BATTERY_GOOD_KEYWORDS = ("dobry",)
_BATTERY_FAIR_KEYWORDS = ("zadowalający", "zadowalajacy")
_BATTERY_WEAK_KEYWORDS = ("słaby", "slaby")


@dataclass(frozen=True)
class BlankWidths:

    company_name: int = 58
    company_address: int = 71
    company_nip: int = 75
    employee: int = 75
    model: int = 54
    imei: int = 72
    serial_number: int = 68
    phone: int = 56
    condition: int = 400  
    signature: int = 60  

DOCX_BLANKS = BlankWidths(condition=396)
PDF_BLANKS = BlankWidths(condition=426)

_HANDOVER_DECLARATIONS: tuple[str, ...] = (
    "Otrzymałem(-am) telefon w stanie sprawnym i bez zastrzeżeń (lub zgodnie z opisem stanu "
    "technicznego wskazanym powyżej).",
    "Zobowiązuję się do użytkowania telefonu zgodnie z jego przeznaczeniem.",
    "Zobowiązuję się do należytego zabezpieczenia telefonu przed utratą, zniszczeniem lub "
    "uszkodzeniem.",
    "Zobowiązuję się do niezwłocznego zgłaszania utraty, kradzieży, uszkodzenia lub awarii "
    "urządzenia.",
    "Zobowiązuję się do zwrotu telefonu wraz z całym przekazanym wyposażeniem niezwłocznie po "
    "ustaniu potrzeby jego używania lub najpóźniej w dniu zakończenia zatrudnienia, chyba że "
    "Pracodawca wskaże inny termin.",
    "Przyjmuję do wiadomości, że telefon został mi powierzony z obowiązkiem zwrotu i ponoszę "
    "odpowiedzialność materialną za powierzone mienie na zasadach określonych w kodeksie pracy.",
    "W przypadku uszkodzenia lub zniszczenia telefonu z mojej winy zobowiązuję się do pokrycia "
    "kosztów jego naprawy lub, jeśli naprawa nie będzie możliwa, kosztów zakupu telefonu.",
)

_RETURN_CONFIRMATION = (
    "Pracodawca potwierdza odbiór telefonu służbowego wraz z wyposażeniem wskazanym w niniejszym "
    "protokole. Strony potwierdzają, że stan zwracanego mienia oraz ewentualne braki, uszkodzenia "
    "lub inne uwagi zostały opisane powyżej."
)

_PDF_FONT_CANDIDATES: tuple[tuple[str, str, str], ...] = (
    ("Arial", "C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("Calibri", "C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf"),
    (
        "DejaVuSans",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ),
)
_PDF_FALLBACK_FONTS = ("Helvetica", "Helvetica-Bold")

_DOCX_FONT = "Palatino Linotype"


# ---------------------------------------------------------------------------
# Model danych i treść protokołu
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ProtocolData:
    is_handover: bool
    date: str
    company_name: str
    company_address: str
    company_nip: str
    employee: str
    model: str
    imei: str
    serial_number: str
    phone_and_sim: str
    condition_line: str
    has_charger: bool
    has_cable: bool
    has_case: bool
    has_glass: bool
    battery_good: bool
    battery_fair: bool
    battery_weak: bool


class _Kind(Enum):
    TITLE = auto()
    SUBTITLE = auto()
    HEADING = auto()
    TEXT = auto()
    TEXT_BOLD = auto()
    BULLET = auto()


@dataclass(frozen=True)
class _Block:
    kind: _Kind
    text: str


def _blank(width: int) -> str:
    return "." * width


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _contains_any(text: str, keywords: Sequence[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def _prepare_data(
    phone_data: Sequence[Any],
    employer_data: Optional[Sequence[Any]],
    protocol_type: str,
    blanks: BlankWidths,
) -> ProtocolData:
    model = _text(phone_data[_IDX_MODEL])
    nr_tel = _text(phone_data[_IDX_NR_TEL])
    nr_sim = _text(phone_data[_IDX_NR_SIM])
    equipment = _text(phone_data[_IDX_EQUIPMENT]).lower()
    battery = _text(phone_data[_IDX_BATTERY]).lower()

    company_name, company_address, company_nip = (
        (_text(value) for value in employer_data) if employer_data else ("", "", "")
    )

    if nr_sim:
        phone_and_sim = f"{nr_tel} (SIM: {nr_sim})"
    else:
        phone_and_sim = nr_tel or _blank(blanks.phone)

    return ProtocolData(
        is_handover=protocol_type == PROTOCOL_HANDOVER,
        date=datetime.now().strftime(DATE_FORMAT),  # noqa: DTZ005
        company_name=company_name or _blank(blanks.company_name),
        company_address=company_address or _blank(blanks.company_address),
        company_nip=company_nip or _blank(blanks.company_nip),
        employee=_text(phone_data[_IDX_USER]) or _blank(blanks.employee),
        model=model or _blank(blanks.model),
        imei=_text(phone_data[_IDX_IMEI]) or _blank(blanks.imei),
        serial_number=_text(phone_data[_IDX_SERIAL]) or _blank(blanks.serial_number),
        condition_line=_blank(blanks.condition),
        phone_and_sim=phone_and_sim,
        has_charger=_contains_any(equipment, _CHARGER_KEYWORDS),
        has_cable=_contains_any(equipment, _CABLE_KEYWORDS),
        has_case=_contains_any(equipment, _CASE_KEYWORDS),
        has_glass=_contains_any(equipment, _GLASS_KEYWORDS),
        battery_good=_contains_any(battery, _BATTERY_GOOD_KEYWORDS) and "brak" not in battery,
        battery_fair=_contains_any(battery, _BATTERY_FAIR_KEYWORDS),
        battery_weak=_contains_any(battery, _BATTERY_WEAK_KEYWORDS),
    )


def _build_content(
    data: ProtocolData, checkbox: Callable[[bool], str]
) -> tuple[list[_Block], tuple[str, str]]:
    handover = data.is_handover
    action = "przekazania" if handover else "zwrotu"

    blocks = [
        _Block(_Kind.TITLE, f"PROTOKÓŁ {action.upper()} TELEFONU SŁUŻBOWEGO"),
        _Block(_Kind.SUBTITLE, f"sporządzony w dniu {data.date}"),
        _Block(_Kind.HEADING, "1. Dane pracodawcy"),
        _Block(_Kind.TEXT, f"Nazwa firmy: {data.company_name}"),
        _Block(_Kind.TEXT, f"Adres: {data.company_address}"),
        _Block(_Kind.TEXT, f"NIP: {data.company_nip}"),
        _Block(_Kind.HEADING, "2. Dane pracownika"),
        _Block(_Kind.TEXT, f"Imię i nazwisko: {data.employee}"),
        _Block(_Kind.HEADING, f"3. Przedmiot {action}"),
        _Block(
            _Kind.TEXT,
            "Pracodawca przekazuje pracownikowi do użytkowania następujące mienie służbowe:"
            if handover
            else "Pracownik zwraca pracodawcy następujące mienie służbowe:",
        ),
        _Block(_Kind.TEXT, f"Marka i model telefonu: {data.model}"),
        _Block(_Kind.TEXT, f"Numer IMEI: {data.imei}"),
        _Block(_Kind.TEXT, f"Numer seryjny: {data.serial_number}"),
        _Block(_Kind.TEXT, f"Numer telefonu (SIM): {data.phone_and_sim}"),
        _Block(_Kind.HEADING, "4. Wyposażenie dodatkowe"),
        _Block(_Kind.TEXT, f"Ładowarka {checkbox(data.has_charger)}"),
        _Block(_Kind.TEXT, f"Kabel USB {checkbox(data.has_cable)}"),
        _Block(_Kind.TEXT, f"Etui {checkbox(data.has_case)}"),
        _Block(_Kind.TEXT, f"Szkło ochronne {checkbox(data.has_glass)}"),
        _Block(
            _Kind.TEXT_BOLD,
            f"Stan techniczny urządzenia w chwili {'przekazania' if handover else 'zwrotu'}:",
        ),
        _Block(_Kind.TEXT, data.condition_line),
        _Block(
            _Kind.TEXT_BOLD,
            f"Stan baterii: Dobry {checkbox(data.battery_good)} / "
            f"Zadowalający {checkbox(data.battery_fair)} / Słaby {checkbox(data.battery_weak)}",
        ),
    ]

    if handover:
        blocks.append(_Block(_Kind.HEADING, "5. Oświadczenie pracownika"))
        blocks.append(
            _Block(
                _Kind.TEXT,
                "Ja, niżej podpisany(-a), potwierdzam odbiór wyżej opisanego telefonu służbowego "
                "wraz z wyposażeniem i oświadczam, że:",
            )
        )
        blocks.extend(
            _Block(_Kind.BULLET, f"{number}. {declaration}")
            for number, declaration in enumerate(_HANDOVER_DECLARATIONS, start=1)
        )
        signatures = ("Przekazujący (Pracodawca):", "Odbierający (Pracownik):")
    else:
        blocks.append(_Block(_Kind.HEADING, "5. Potwierdzenie odbioru przez pracodawcę"))
        blocks.append(_Block(_Kind.TEXT, _RETURN_CONFIRMATION))
        signatures = ("Odbierający / Pracodawca", "Zwracający / Pracownik")

    return blocks, signatures


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _DocxStyle:
    bold: bool = False
    centered: bool = False
    font_size: float = 10
    space_before: float = 0
    space_after: float = 3
    line_spacing: float = 0.9
    indent_inches: float = 0


_DOCX_STYLES: dict[_Kind, _DocxStyle] = {
    _Kind.TITLE: _DocxStyle(bold=True, centered=True, space_after=2),
    _Kind.SUBTITLE: _DocxStyle(bold=True, centered=True, space_after=8),
    _Kind.HEADING: _DocxStyle(bold=True, space_before=6, space_after=2),
    _Kind.TEXT: _DocxStyle(),
    _Kind.TEXT_BOLD: _DocxStyle(bold=True, space_before=2, space_after=4),
    _Kind.BULLET: _DocxStyle(font_size=9, space_after=0.5, line_spacing=1.0, indent_inches=0.15),
}
_DOCX_MARGINS_INCHES = {"top": 0.6, "bottom": 0.6, "left": 0.8, "right": 0.8}
_DOCX_SIGNATURE_SPACING_PT = 16


def _docx_checkbox(checked: bool) -> str:
    return "☑" if checked else "☐"


def _add_docx_paragraph(doc: Any, text: str, style: _DocxStyle) -> Any:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if style.centered else WD_ALIGN_PARAGRAPH.LEFT
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(style.space_before)
    fmt.space_after = Pt(style.space_after)
    fmt.line_spacing = style.line_spacing
    if style.indent_inches:
        fmt.left_indent = Inches(style.indent_inches)

    if text:
        run = paragraph.add_run(text)
        run.bold = style.bold
        run.font.name = _DOCX_FONT
        run.font.size = Pt(style.font_size)
    return paragraph


def generate_docx(
    file_path: str | Path,
    phone_data: Sequence[Any],
    employer_data: Optional[Sequence[Any]],
    protocol_type: str,
) -> None:
    if Document is None:
        raise ImportError("Moduł 'python-docx' nie jest zainstalowany.")

    data = _prepare_data(phone_data, employer_data, protocol_type, DOCX_BLANKS)
    blocks, (left_signature, right_signature) = _build_content(data, _docx_checkbox)
    signature_line = _blank(DOCX_BLANKS.signature)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(_DOCX_MARGINS_INCHES["top"])
        section.bottom_margin = Inches(_DOCX_MARGINS_INCHES["bottom"])
        section.left_margin = Inches(_DOCX_MARGINS_INCHES["left"])
        section.right_margin = Inches(_DOCX_MARGINS_INCHES["right"])

    for block in blocks:
        _add_docx_paragraph(doc, block.text, _DOCX_STYLES[block.kind])

    _add_docx_paragraph(doc, "", _DocxStyle(space_after=_DOCX_SIGNATURE_SPACING_PT))

    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = signature_line
    table.rows[0].cells[1].text = signature_line
    table.rows[1].cells[0].text = left_signature
    table.rows[1].cells[1].text = right_signature

    doc.save(file_path)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

_PDF_MARGINS = {"leftMargin": 40, "rightMargin": 40, "topMargin": 35, "bottomMargin": 35}
_PDF_SIGNATURE_COLUMN_WIDTHS = (260, 250)
_PDF_SIGNATURE_SPACING = 20


def _pdf_checkbox(checked: bool) -> str:
    return "[X]" if checked else "[ ]"


@lru_cache(maxsize=1)
def _get_pdf_fonts() -> tuple[str, str]:
    for name, regular_path, bold_path in _PDF_FONT_CANDIDATES:
        if not Path(regular_path).exists():
            continue
        bold_name = f"{name}-Bold"
        try:
            pdfmetrics.registerFont(TTFont(name, regular_path))
            bold_source = bold_path if Path(bold_path).exists() else regular_path
            pdfmetrics.registerFont(TTFont(bold_name, bold_source))
        except (OSError, TTFError):
            logger.exception("Nie udało się zarejestrować czcionki %s", name)
            continue
        return name, bold_name

    logger.warning("Brak czcionki TTF z polskimi znakami - użyto czcionki Helvetica.")
    return _PDF_FALLBACK_FONTS


def _pdf_styles(regular: str, bold: str) -> dict[_Kind, ParagraphStyle]:
    return {
        _Kind.TITLE: ParagraphStyle("Title", fontName=bold, fontSize=12, alignment=1, spaceAfter=2),
        _Kind.SUBTITLE: ParagraphStyle(
            "Subtitle", fontName=bold, fontSize=9, alignment=1, spaceAfter=8
        ),
        _Kind.HEADING: ParagraphStyle(
            "Heading", fontName=bold, fontSize=9.5, spaceBefore=12, spaceAfter=2
        ),
        _Kind.TEXT: ParagraphStyle("Text", fontName=regular, fontSize=8.5, leading=11, spaceAfter=2),
        _Kind.TEXT_BOLD: ParagraphStyle(
            "TextBold", fontName=bold, fontSize=8.5, leading=11, spaceAfter=2
        ),
        _Kind.BULLET: ParagraphStyle(
            "Bullet", fontName=regular, fontSize=8, leading=10, leftIndent=10, spaceAfter=1.5
        ),
    }


def generate_pdf(
    file_path: str | Path,
    phone_data: Sequence[Any],
    employer_data: Optional[Sequence[Any]],
    protocol_type: str,
) -> None:
    if SimpleDocTemplate is None:
        raise ImportError("Moduł 'reportlab' nie jest zainstalowany.")

    data = _prepare_data(phone_data, employer_data, protocol_type, PDF_BLANKS)
    blocks, (left_signature, right_signature) = _build_content(data, _pdf_checkbox)
    signature_line = _blank(PDF_BLANKS.signature)
    styles = _pdf_styles(*_get_pdf_fonts())
    text_style = styles[_Kind.TEXT]

    def paragraph(text: str, style: ParagraphStyle) -> Paragraph:
        return Paragraph(escape(text), style)

    story: list[Any] = [paragraph(block.text, styles[block.kind]) for block in blocks]
    story.append(Spacer(1, _PDF_SIGNATURE_SPACING))

    signature_table = Table(
        [
            [paragraph(signature_line, text_style), paragraph(signature_line, text_style)],
            [paragraph(left_signature, text_style), paragraph(right_signature, text_style)],
        ],
        colWidths=list(_PDF_SIGNATURE_COLUMN_WIDTHS),
    )
    signature_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(signature_table)

    SimpleDocTemplate(str(file_path), pagesize=A4, **_PDF_MARGINS).build(story)
