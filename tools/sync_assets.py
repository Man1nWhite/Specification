"""Synchronize Module1.bas, the single .dotx template, and the Excel preview sheet.

For the template this step owns the bookmarks, the reviewed content blocks
(assumptions, the list of exclusions, the handover documents, the mounting scope,
the Gantt chart, the terms page), the number in the running head and the 12 pt
text policy. The reviewed edition of the template is the source of truth: where
the review cut a block, the step removes it instead of putting it back, and a
second run leaves the file byte for byte identical.

Run from the repository root with Python 3.10+ after installing requirements-dev.txt.
Excel and Word are not required for this packaging step.
"""

from __future__ import annotations

import os
import posixpath
from copy import deepcopy
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE_DOTX = ROOT / "ТКП_шаблон.dotx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
X_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DC_NS = "http://purl.org/dc/elements/1.1/"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
W = f"{{{W_NS}}}"
X = f"{{{X_NS}}}"
NS = {"w": W_NS}

PREVIEW_SHEET = "Спец_ТКП"

# Word shows this property as the document name when a document is created from
# the template. The visible title on the first page stays untouched.
DOCUMENT_TITLE = "ТКП"

# ---------------------------------------------------------------------------
# Typography policy
# ---------------------------------------------------------------------------
# OOXML stores font sizes in half-points. Every piece of body text is 12 pt
# ("24"); the 10 pt ("20") and 11 pt ("22") sizes the template used to mix are
# retired, both in direct run formatting and in the styles/numbering the
# document actually uses. Running heads and footers keep their own small sizes.
HOUSE_SIZE = "24"
RETIRED_SIZES = frozenset({"20", "22"})

# One table is deliberately set at 11 pt: the stages of the calendar plan in
# Appendix 8. Four columns of names and terms wrap badly at 12 pt, and the table
# is a schedule rather than prose, so the smaller size stays. It is the only
# exception to the house size, and `apply_font_policy` skips exactly this table.
CALENDAR_PLAN_ANCHOR = "Плановая дата выполнения"
CALENDAR_PLAN_SIZE = "22"

# ---------------------------------------------------------------------------
# Reviewed content
# ---------------------------------------------------------------------------
# The running header repeats the number from the «Номер» row of the first page
# through a REF field, so both placeholders have to read the same.
NUMBER_PLACEHOLDER = "№ __/1 от __.__.2026"

# Page two opens with what the offer rests on. The general note that the price
# «may be refined» gave the customer nothing to hold on to, so it names the
# source document. The wording is deliberately one sentence on one line: the
# review asked for short blocks where every sentence carries weight, so the
# four-bullet list of the previous revision is gone, and the document it rests on
# stays inside the same sentence instead of starting a line of its own. The
# constants below are the source of truth, and the sync step rewrites this
# paragraph to them; the empty item list also folds a template that still keeps
# the document on a separate line back into one paragraph.
BASIS_HEADING = "Исходные данные и допущения:"
BASIS_INTRO = (
    "Предложение подготовлено на основании следующих исходных данных: "
    "документ-основание — ____________ (наименование), № ____ от __.__.____, "
    "ревизия ____;"
)
BASIS_ITEMS: tuple[str, ...] = ()
LEGACY_NOTE_HEADING = "Примечание:"

# Then the scope boundary: what the offer does not cover. The introduction that
# used to open the list («В объём и стоимость настоящего предложения не
# входят:») repeated the heading, so the bullets follow it directly.
EXCLUSIONS_HEADING = "Не входит в объём предложения:"
EXCLUSIONS_INTRO = ""
EXCLUSIONS_ITEMS = (
    "поставка, монтаж и подключение датчиков, измерительных преобразователей и "
    "исполнительных механизмов;",
    "внешние кабельные связи между датчиками, механизмами и шкафами управления: "
    "поставка, прокладка и расключение;",
    "строительная часть, подготовка помещений и устройство оснований под шкафы;",
    "поставка и установка АРМ и серверов верхнего уровня;",
    "электропитание шкафов, включая ИБП и контур гарантированного питания;",
    "проектная документация стадий «П» и «Р» по объекту в целом: Исполнитель "
    "разрабатывает РКД в части ПТК АСУ ТП;",
    "метрологическая аттестация измерительных каналов и поверка средств измерений;",
    "экспертиза промышленной безопасности документации и технических устройств;",
    "мероприятия по категорированию объекта КИИ и выполнению требований ФСТЭК "
    "России, включая аттестацию по требованиям информационной безопасности;",
    "согласование документации и получение разрешений в надзорных органах.",
)
# The first revision of this template called the block «Допущения».
LEGACY_EXCLUSIONS_HEADING = "Допущения:"

# A «Обязательства Заказчика» section stood between the exclusions and the
# warranty. It repeated what Appendix 5 already requires from the customer in
# detail, so it was cut: site readiness is stated once, next to the mounting
# scope. A template of the previous revision still carries the block, and the
# sync step removes it rather than leaving a duplicate behind.
LEGACY_DUTIES_HEADING = "Обязательства Заказчика:"

# The parameters table lists the stages but never said how long the whole
# project takes; the critical path of Appendix 8 adds up to 35 weeks.
DURATION_ANCHOR = "ИТОГО без НДС"
DURATION_LABEL = "Общий срок реализации проекта"
DURATION_VALUE = "35 недель"

# Payment terms: the moment of the second payment for the equipment, seven days
# instead of three for the stage advance, and the invoice that follows the signed
# act. The review cut the sentence about the moment title passes to the customer:
# the conditions of payment name the signed УПД once, and the guarantee runs from
# the same document, so the paragraph stops there.
PAYMENT_EQUIPMENT_ANCHOR = "Оплата аванса в размере 50% от стоимости оборудования"
PAYMENT_EQUIPMENT = (
    "Оплата аванса в размере 50% от стоимости оборудования согласно пунктам 1–4 в "
    "течение 7 (семи) календарных дней с даты подписания Сторонами Договора, 50% "
    "от стоимости в течение 7 (семи) календарных дней с даты подписания Сторонами "
    "универсального передаточного документа (УПД), оформляемого при передаче "
    "оборудования на складе Заказчика."
)
PAYMENT_WORKS_ANCHOR = "Оплата работ по пунктам 5–8"
PAYMENT_WORKS = (
    "Оплата работ по пунктам 5–8 осуществляется поэтапно: аванс в размере 50% от "
    "стоимости работ по этапу в течение 7 (семи) календарных дней с даты начала "
    "этапа по календарному плану, 50% от стоимости работ по этапу в течение 7 "
    "(семи) календарных дней с даты подписания акта выполненных работ. Возможен "
    "зачёт аванса пропорционально сумме ежемесячного выполнения."
)
PAYMENT_TRAINING_ANCHOR = "Оплата по пункту 9"
PAYMENT_TRAINING = (
    "Оплата по пункту 9 в размере 100% осуществляется в течение 10 (десяти) "
    "календарных дней с даты направления счёта; счёт выставляется после подписания "
    "Сторонами акта оказанных услуг."
)
PAYMENTS = (
    (PAYMENT_EQUIPMENT_ANCHOR, PAYMENT_EQUIPMENT),
    (PAYMENT_WORKS_ANCHOR, PAYMENT_WORKS),
    (PAYMENT_TRAINING_ANCHOR, PAYMENT_TRAINING),
)

# Openings that are re-worded in place, so the rest of the paragraph (and its
# runs, which carry non-breaking spaces around the certificate numbers) survives.
REWORDINGS = (
    (
        "Поставка оборудования и выполнение работ по созданию АСУ ТП осуществляется",
        "Поставка оборудования для создания АСУ ТП осуществляется",
    ),
    (
        "Система «МИРТС» является основой создаваемой АСУ",
        "ПТК АРКС400 входит в состав системы «МИРТС», которая является основой "
        "создаваемой АСУ",
    ),
)

# The reference pages at the end are appendices in their own right: each gets a
# «Приложение N» label with the same automatic number the other appendices use,
# and a numbered line in the list of appendices. The pair is (the line in the
# list, the opening words of the page the label goes on).
REFERENCE_APPENDICES = (
    ("Описание ПТК АРКС400", "АРКС400 — современные универсальные контроллеры"),
    ("О компании «ЭКСАРА»", "О компании «ЭКСАРА»"),
)
# An earlier revision listed the same pages as unnumbered reference material.
LEGACY_INFORMATION_GROUP = "Информационные материалы"
APPENDIX_LIST_ANCHOR = "Наименование приложения"

MOUNTING_ANCHOR = "Монтаж шкафов на подготовленные основания"
MOUNTING_FIRST_WORK = "Разгрузка оборудования и перемещение к месту установки"
MOUNTING_EXTRA_WORKS = (
    "Монтаж межшкафных связей",
    "Монтаж ЛВС и АРМ",
    "Сборка кроссов",
    "Маркировка кабелей",
)
# Appendix 5 states the site readiness requirements once, without a heading of
# their own: the sentence above the list introduces them, the way the paragraphs
# around it introduce their own lists. The list is what the customer has to
# prepare before mounting, in the order the works need it.
SITE_READINESS_INTRO = "К началу монтажных работ Заказчик обеспечивает:"
SITE_READINESS = (
    "основания и фальшпол: выполненные основания (цоколи) под шкафы по чертежам "
    "общего вида, фальшпол с необходимой несущей способностью и кабельными каналами, "
    "разметка мест установки.",
    "контур заземления: смонтированный контур защитного и рабочего (информационного) "
    "заземления с сопротивлением не более 4 Ом, главная заземляющая шина в помещении.",
    "электропитание: подведённые сети ~220 В (1ф) и ~380 В (3ф) частотой 50 Гц с "
    "запасом по мощности не менее 5 кВт на шкаф, отдельные аппараты защиты.",
    "освещение: рабочее и аварийное освещение не менее 300 лк в зоне монтажа, "
    "розетки ~220 В для подключения инструмента.",
    "пропускной режим: оформленные пропуска для персонала и транспорта, допуск к "
    "работам, согласование огневых и высотных работ, вводный инструктаж по охране труда.",
    "бытовки: помещения для хранения инструмента и спецодежды, места для отдыха и "
    "приёма пищи, доступ к санитарно-бытовым помещениям.",
    "места складирования: крытая запираемая площадка для разгрузки и временного "
    "хранения оборудования с подъездными путями и возможностью применения "
    "грузоподъёмной техники.",
)
# A heading and a nine-item list opened this block in the previous revision. Both
# are gone — the sentence above the list does that job — so a template of that
# revision is brought forward by removing the heading and its duplicates.
LEGACY_SITE_READINESS_HEADING = "Требования к готовности площадки:"
MOUNTING_NOTE_ANCHOR = "Примечание: Прокладка и расключение внешних кабельных связей"
MOUNTING_NOTE = (
    "Примечание: Прокладка и расключение внешних кабельных связей от "
    "датчиков/механизмов до шкафов управления выполняются силами Заказчика. В объём "
    "монтажных работ также не входят: поставка и монтаж кабеленесущих систем (лотки, "
    "короба, эстакады), устройство кабельных и противопожарных проходок через стены и "
    "перекрытия, огнезащита кабельных линий и конструкций, демонтаж существующего "
    "оборудования и кабельных линий, общестроительные и отделочные работы на объекте."
)

DOCUMENTS_ANCHOR = "Кабельный журнал"
DOCUMENTS_STAGE_NOTE = "Оформляется на этапе выполнения монтажных работ"
# Appendix 4 lists the documents themselves, 1..22. A shaded group row that
# titled the set «РКД и ЭД по ГОСТ 34» stood above them in the previous
# revision; the review dropped it, because a title line in the table says the
# same as the heading of the appendix and pushes the list down a row.
LEGACY_GROUP_ROW = "РКД и ЭД по ГОСТ 34"
RKD_DOCUMENTS = (
    "Схема структурная комплекса технических средств автоматизации",
    "Схемы соединений внешних проводок и таблицы соединений внешних проводок",
    "Схемы принципиальные и схемы подключений",
    "Чертежи общего вида шкафов",
    "Ведомость покупных изделий",
    "Программа и методика испытаний (ПМИ)",
    "Протоколы заводских приёмо‑сдаточных испытаний (ЗПИ)",
    "Протоколы приёмо‑сдаточных испытаний (ПСИ)",
    "Описание прикладного программного обеспечения (ППО)",
    "Руководство оператора",
    "Руководство администратора",
    "Инструкция по техническому обслуживанию",
    "Формуляр (паспорт) на оборудование",
    "Ведомость ЗИП",
)
# Handed over with the equipment and after the commissioning, outside the
# ГОСТ 34 group.
HANDOVER_DOCUMENTS = (
    "Паспорта и сертификаты соответствия на покупное оборудование",
    "Акт выполненных пусконаладочных работ",
)
# General rows that repeated the specific items above.
DUPLICATE_DOCUMENTS = frozenset(
    {
        "Документация на ППО",
        "Эксплуатационная документация (руководство оператора, паспорта)",
    }
)

TERMS_TITLE = "Термины и сокращения"
TERMS_INTRO = "В настоящем документе используются следующие термины и сокращения:"
TERMS_COLUMNS = ("Сокращение", "Расшифровка")
TERMS = (
    ("ПТК", "программно-технический комплекс"),
    ("ППО", "прикладное программное обеспечение"),
    ("РКД", "рабочая конструкторская документация"),
    ("УСО", "устройства связи с объектом"),
    ("ЗПИ", "заводские приёмо‑сдаточные испытания"),
    ("ПСИ", "приёмо‑сдаточные испытания на объекте"),
    ("АРМ", "автоматизированное рабочее место"),
    ("КИИ", "критическая информационная инфраструктура"),
    ("АСУ ТП", "автоматизированная система управления технологическим процессом"),
    ("ЗИП", "запасные части, инструменты и принадлежности"),
    ("ИБП", "источник бесперебойного питания"),
    ("ИМ", "исполнительный механизм"),
    ("ЛВС", "локальная вычислительная сеть"),
    ("НДС", "налог на добавленную стоимость"),
    ("ПД", "проектная документация"),
    ("ПМИ", "программа и методика испытаний"),
    ("ПНР", "пусконаладочные работы"),
    ("СИ", "средство измерений"),
    ("ТО", "техническое обслуживание"),
    ("ФСТЭК России", "Федеральная служба по техническому и экспортному контролю"),
    ("ЭД", "эксплуатационная документация"),
)

# Appendix 8 gives the stages in words; the chart shows the same plan in weeks.
# The bars follow the critical path: 20 weeks of equipment, a week of factory
# tests, a week of delivery, three weeks of mounting and ten weeks of
# commissioning — 35 weeks in total.
#
# The chart goes straight under the table of stages, without a heading and a
# caption: the whole appendix is the calendar plan, so «Диаграмма Ганта:» above
# the bars told the reader nothing new. A template of the previous revision
# carries both paragraphs, and the sync step removes them.
LEGACY_GANTT_HEADING = "Диаграмма Ганта:"
LEGACY_GANTT_CAPTION = (
    "Полосы показывают недели выполнения этапов; отсчёт недель — с даты "
    "заключения Договора."
)
GANTT_ANCHOR = "20 (двадцать) недель с даты поступления авансового платежа"
GANTT_STAGES = (
    ("Разработка РКД", 1, 10),
    ("Разработка ППО", 1, 19),
    ("Заказ и сборка оборудования", 1, 20),
    ("Заводские приёмо‑сдаточные испытания", 21, 21),
    ("Доставка оборудования", 22, 22),
    ("Монтажные работы", 23, 25),
    ("Пусконаладочные работы", 26, 35),
    ("Обучение персонала", 35, 35),
)
GANTT_WEEKS = 35

# CT_* child order, so that generated properties stay schema-valid for Word.
RPR_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
)
PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
)
TCPR_ORDER = (
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
    "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
    "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
)

W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
GROUP_ROW_FILL = "D9D9D9"
TERMS_COLUMN_TWIPS = ("2200", "7700")
TERMS_COLUMN_PCT = ("1111", "3889")

# Geometry and colours of the Gantt chart. The tables of the document keep their
# colours: a dark red head, light grey group rows, blue accents and grey grid
# lines.
#
# Widths are the whole page, not the text area: the 35 week columns are only
# legible at 8 pt, so the grid takes all the room a sheet gives it. On the
# portrait page that would be 14574 twips against a 9780-twip text area, which is
# why the chart stands on a page of its own with the sheet turned — landscape
# keeps the same margins and leaves 14712 twips. The chart used to carry a
# 12-twip spare column at the right edge, which was a layout accident rather than
# a decision, so it is gone.
BAR_FILL = "0640A5"
HEAD_FILL = "C00000"
GRID_LINE = "BFBFBF"
GANTT_LABEL_TWIPS = 4949
GANTT_WEEK_TWIPS = 275
GANTT_TABLE_TWIPS = GANTT_LABEL_TWIPS + GANTT_WEEK_TWIPS * GANTT_WEEKS
# Two digits have to fit into one week column, so the scale is set smaller than
# the body text; the stage names stay at the 12 pt house size.
GANTT_SCALE_SIZE = "16"
# The section break that puts the chart on a landscape page of its own. The chart
# is the only content of that section, and the appendix continues on a portrait
# page after it, so the document keeps two breaks and three sections.
GANTT_SECTION_TYPE = "nextPage"


def replace_zip_part(
    path: Path,
    parts: dict[str, bytes],
    *,
    remove_members: set[str] | None = None,
) -> None:
    """Replace just the requested members; preserve the other ZIP entries."""
    remove_members = remove_members or set()
    mode = path.stat().st_mode & 0o777
    with NamedTemporaryFile(dir=path.parent, suffix=path.suffix, delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        with ZipFile(path) as old, ZipFile(temp_path, "w", ZIP_DEFLATED) as new:
            for info in old.infolist():
                if info.filename in remove_members:
                    continue
                new.writestr(info, parts.get(info.filename, old.read(info.filename)))
        os.replace(temp_path, path)
        os.chmod(path, mode)
    finally:
        temp_path.unlink(missing_ok=True)


def write_xml(element: etree._Element) -> bytes:
    return etree.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def word_text(element: etree._Element) -> str:
    return "".join(t.text or "" for t in element.findall(".//w:t", NS))


def bookmark(name: str, bookmark_id: int) -> tuple[etree._Element, etree._Element]:
    start = etree.Element(W + "bookmarkStart", {W + "id": str(bookmark_id), W + "name": name})
    end = etree.Element(W + "bookmarkEnd", {W + "id": str(bookmark_id)})
    return start, end


def run(text: str, example: etree._Element | None = None) -> etree._Element:
    result = etree.Element(W + "r")
    if example is not None:
        props = example.find(W + "rPr")
        if props is not None:
            result.append(deepcopy(props))
    value = etree.SubElement(result, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        value.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    value.text = text
    return result


def remove_bookmark(root: etree._Element, name: str) -> None:
    """Remove the bookmark `name` together with the end that closes it.

    An earlier revision of the template reused bookmark ids: the «Spec_Table»
    bookmark shared its id with a `_Toc` entry, so removing every end with that
    id tore the ends off the other bookmarks as well. The start is paired with
    its own end by their order in the part, which keeps the rest untouched.
    """
    starts = root.xpath(".//w:bookmarkStart[@w:name=$name]", namespaces=NS, name=name)
    if len(starts) != 1:
        raise ValueError(f"Expected one bookmark {name!r}, found {len(starts)}")
    start = starts[0]
    bm_id = start.get(W + "id")
    same_id_starts = root.xpath(
        ".//w:bookmarkStart[@w:id=$id]", namespaces=NS, id=bm_id
    )
    ends = root.xpath(".//w:bookmarkEnd[@w:id=$id]", namespaces=NS, id=bm_id)
    position = same_id_starts.index(start)
    if ends:
        end = ends[position] if position < len(ends) else ends[-1]
        if end.getparent() is not None:
            end.getparent().remove(end)
    if start.getparent() is not None:
        start.getparent().remove(start)


def xml_root(element: etree._Element) -> etree._Element:
    """The root of the part `element` belongs to."""
    node = element
    while node.getparent() is not None:
        node = node.getparent()
    return node


def free_bookmark_id(root: etree._Element) -> int:
    """An id no bookmark of `root` uses yet.

    The ids are hard-coded numbers in older revisions of this step, and the
    document already uses some of them for its own `_Hlk` and `_Toc` bookmarks:
    the same id twice makes Word repair the file, so every created bookmark asks
    for a free one first.
    """
    used = set()
    for tag in ("bookmarkStart", "bookmarkEnd"):
        for element in root.iter(W + tag):
            try:
                used.add(int(element.get(W + "id")))
            except (TypeError, ValueError):
                continue
    return max(used, default=0) + 1


def drop_unpaired_bookmarks(root: etree._Element) -> int:
    """Remove bookmark starts that no end closes.

    The template of the previous revision reused bookmark ids, so removing one
    bookmark could take the end of another with it. A start without its end is
    an invalid part, which Word offers to repair: the start goes instead.
    """
    ends = {end.get(W + "id") for end in root.iter(W + "bookmarkEnd")}
    dropped = 0
    for start in list(root.iter(W + "bookmarkStart")):
        if start.get(W + "id") in ends:
            continue
        if start.getparent() is not None:
            start.getparent().remove(start)
        dropped += 1
    return dropped


def add_empty_bookmark(paragraph: etree._Element, name: str, number: int) -> None:
    start, end = bookmark(name, number)
    paragraph.extend((start, end))


def insert_in_order(
    parent: etree._Element, child: etree._Element, order: tuple[str, ...]
) -> etree._Element:
    """Insert `child` into `parent` where the OOXML content model expects it."""
    name = etree.QName(child).localname
    position = len(parent)
    for index, existing in enumerate(parent):
        existing_name = etree.QName(existing).localname
        if existing_name in order and order.index(existing_name) > order.index(name):
            position = index
            break
    parent.insert(position, child)
    return child


def drop_para_ids(element: etree._Element) -> etree._Element:
    """Drop w14:paraId/w14:textId from a clone so paragraph ids stay unique."""
    for node in element.iter():
        for attribute in (f"{{{W14_NS}}}paraId", f"{{{W14_NS}}}textId"):
            node.attrib.pop(attribute, None)
    return element


def arial_rpr(bold: bool = False) -> etree._Element:
    """Run properties in the house face: Arial, optionally bold."""
    rpr = etree.Element(W + "rPr")
    fonts = etree.SubElement(rpr, W + "rFonts")
    fonts.set(W + "ascii", "Arial")
    fonts.set(W + "hAnsi", "Arial")
    fonts.set(W + "cs", "Arial")
    if bold:
        etree.SubElement(rpr, W + "b")
    return rpr


def styled_run(text: str, *, bold: bool = False) -> etree._Element:
    result = etree.Element(W + "r")
    result.append(arial_rpr(bold))
    value = etree.SubElement(result, W + "t")
    if text != text.strip():
        value.set(XML_SPACE, "preserve")
    value.text = text
    return result


def paragraph_text_of(example: etree._Element, text: str, *, bold: bool = False) -> etree._Element:
    """Copy the paragraph formatting of `example` and give it one run of `text`."""
    paragraph = drop_para_ids(deepcopy(example))
    model = paragraph.find(W + "r")
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    properties = paragraph.find(W + "pPr")
    if properties is not None:
        mark = properties.find(W + "rPr")
        if mark is None and bold:
            mark = insert_in_order(properties, etree.Element(W + "rPr"), PPR_ORDER)
        if mark is not None and bold and mark.find(W + "b") is None:
            insert_in_order(mark, etree.Element(W + "b"), RPR_ORDER)
    new_run = run(text, model)
    if bold:
        rpr = new_run.find(W + "rPr")
        if rpr is None:
            rpr = etree.Element(W + "rPr")
            new_run.insert(0, rpr)
        if rpr.find(W + "b") is None:
            insert_in_order(rpr, etree.Element(W + "b"), RPR_ORDER)
    paragraph.append(new_run)
    return paragraph


def set_paragraph_text(paragraph: etree._Element, text: str) -> None:
    """Rewrite the runs of `paragraph`, keeping its formatting."""
    model = paragraph.find(W + "r")
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    paragraph.append(run(text, model))


def bullet_paragraph(text: str) -> etree._Element:
    """A bulleted list item in the house style (Arial, blue square bullet)."""
    paragraph = etree.Element(W + "p")
    properties = etree.SubElement(paragraph, W + "pPr")
    style = etree.SubElement(properties, W + "pStyle")
    style.set(W + "val", "afb")
    numbering = etree.SubElement(properties, W + "numPr")
    level = etree.SubElement(numbering, W + "ilvl")
    level.set(W + "val", "0")
    number = etree.SubElement(numbering, W + "numId")
    number.set(W + "val", "7")
    indent = etree.SubElement(properties, W + "ind")
    indent.set(W + "left", "0")
    alignment = etree.SubElement(properties, W + "jc")
    alignment.set(W + "val", "both")
    properties.append(arial_rpr())
    paragraph.append(styled_run(text))
    return paragraph


def is_plain_paragraph(paragraph: etree._Element) -> bool:
    """An ordinary paragraph: no bullet, so it cannot be a section heading.

    Every section heading of the page with the offer parameters is a bullet
    itself, which is what tells it from the introduction of its own block.
    """
    return paragraph.find(f"{W}pPr/{W}numPr") is None


def is_item_bullet(paragraph: etree._Element) -> bool:
    """A list item of the page-two blocks: bulleted, and not a bold heading."""
    if paragraph.find(f"{W}pPr/{W}numPr") is None:
        return False
    return paragraph.find(f"{W}pPr/{W}rPr/{W}b") is None


def block_intro(heading: etree._Element) -> etree._Element | None:
    """The introduction paragraph of the block `heading` opens, if there is one."""
    intro = heading.getnext()
    if intro is None or intro.tag != W + "p" or not is_plain_paragraph(intro):
        return None
    return intro


def block_items(heading: etree._Element) -> list[etree._Element]:
    """The bullets of a block: directly under its heading, or under its intro."""
    first = heading.getnext()
    if first is None or first.tag != W + "p":
        return []
    if is_item_bullet(first):
        start = first
    elif is_plain_paragraph(first):
        start = first.getnext()
    else:
        return []
    items = []
    sibling = start
    while sibling is not None and sibling.tag == W + "p" and is_item_bullet(sibling):
        items.append(sibling)
        sibling = sibling.getnext()
    return items


def last_of_block(heading: etree._Element) -> etree._Element:
    """Where the next block has to be attached: after the last item of this one."""
    items = block_items(heading)
    if items:
        return items[-1]
    return block_intro(heading) or heading


def sync_list_block(
    heading: etree._Element,
    intro_text: str,
    items: tuple[str, ...],
    intro_example: etree._Element,
) -> None:
    """Bring the introduction and the bullets of `heading`'s block up to date.

    Unlike the steps that only add what is missing, this one owns its block: the
    wording in the constants wins, extra bullets go and missing ones arrive. An
    empty `intro_text` means the bullets open the block on their own, which is
    how «Не входит в объём предложения» reads.
    """
    intro = block_intro(heading)
    if not intro_text:
        if intro is not None:
            intro.getparent().remove(intro)
            intro = None
    elif intro is None:
        intro = paragraph_text_of(intro_example, intro_text)
        heading.addnext(intro)
    elif word_text(intro).strip() != intro_text:
        set_paragraph_text(intro, intro_text)
    current = block_items(heading)
    for paragraph, text in zip(current, items, strict=False):
        if word_text(paragraph).strip() != text:
            set_paragraph_text(paragraph, text)
    for extra in current[len(items) :]:
        extra.getparent().remove(extra)
    kept = current[: len(items)]
    anchor = kept[-1] if kept else (intro if intro is not None else heading)
    for text in items[len(kept) :]:
        bullet = bullet_paragraph(text)
        anchor.addnext(bullet)
        anchor = bullet


def section_heading(
    body: etree._Element, heading: str, legacy: str | None = None
) -> etree._Element | None:
    """Find a page-two section heading, renaming an earlier wording of it."""
    for paragraph in body.findall(W + "p"):
        text = word_text(paragraph).strip()
        if text == heading:
            return paragraph
        if legacy is not None and text == legacy:
            set_paragraph_text(paragraph, heading)
            return paragraph
    return None


def find_paragraph(body: etree._Element, needle: str, *, exact: bool = False) -> etree._Element:
    """Return the first body-level paragraph holding `needle`."""
    for paragraph in body.findall(W + "p"):
        text = word_text(paragraph).strip()
        matches = (text == needle) if exact else text.startswith(needle)
        if matches:
            return paragraph
    raise ValueError(f"Could not locate the paragraph {needle!r}")


def has_paragraph(body: etree._Element, needle: str) -> bool:
    return any(needle in word_text(p) for p in body.findall(W + "p"))


def find_table(body: etree._Element, needle: str) -> etree._Element:
    for table in body.findall(W + "tbl"):
        if needle in word_text(table):
            return table
    raise ValueError(f"Could not locate the table containing {needle!r}")


def row_label(row: etree._Element) -> str:
    """The text of the last cell of a table row (its description column)."""
    cells = row.findall(W + "tc")
    return word_text(cells[-1]).strip() if cells else ""


def set_cell_text(
    cell: etree._Element,
    text: str,
    *,
    bold: bool = False,
    numbered: bool = True,
    align: str | None = None,
) -> None:
    """Replace a table cell's content with one paragraph holding `text`."""
    paragraphs = cell.findall(W + "p")
    for extra in paragraphs[1:]:
        cell.remove(extra)
    paragraph = paragraphs[0]
    drop_para_ids(paragraph)
    model = paragraph.find(W + "r")
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    properties = paragraph.find(W + "pPr")
    if properties is None:
        properties = etree.Element(W + "pPr")
        paragraph.insert(0, properties)
    if not numbered:
        numbering = properties.find(W + "numPr")
        if numbering is not None:
            properties.remove(numbering)
    if align is not None:
        justification = properties.find(W + "jc")
        if justification is None:
            justification = insert_in_order(properties, etree.Element(W + "jc"), PPR_ORDER)
        justification.set(W + "val", align)
    mark = properties.find(W + "rPr")
    if mark is None and bold:
        mark = insert_in_order(properties, etree.Element(W + "rPr"), PPR_ORDER)
    if mark is not None and bold and mark.find(W + "b") is None:
        insert_in_order(mark, etree.Element(W + "b"), RPR_ORDER)
    if not text:
        # An empty cell keeps its paragraph mark and nothing else: a run with an
        # empty «w:t» would be serialised differently on the next round trip.
        return
    if model is None:
        # The numbered column of these tables holds no run to copy the face
        # from, so the text falls back to the house formatting.
        paragraph.append(styled_run(text, bold=bold))
        return
    new_run = run(text, model)
    if bold:
        rpr = new_run.find(W + "rPr")
        if rpr is None:
            rpr = etree.Element(W + "rPr")
            new_run.insert(0, rpr)
        if rpr.find(W + "b") is None:
            insert_in_order(rpr, etree.Element(W + "b"), RPR_ORDER)
    paragraph.append(new_run)


def item_row(example: etree._Element, text: str) -> etree._Element:
    """A numbered data row cloned from `example`, describing `text`."""
    row = drop_para_ids(deepcopy(example))
    set_cell_text(row.findall(W + "tc")[-1], text)
    return row


def group_row(example: etree._Element, text: str) -> etree._Element:
    """A full-width, bold, shaded row that titles a block of table rows."""
    row = drop_para_ids(deepcopy(example))
    cells = row.findall(W + "tc")
    label, description = cells[0], cells[-1]
    row.remove(description)
    properties = label.find(W + "tcPr")
    for old in properties.findall(W + "tcW"):
        properties.remove(old)
    width = etree.Element(W + "tcW")
    width.set(W + "w", "5000")
    width.set(W + "type", "pct")
    insert_in_order(properties, width, TCPR_ORDER)
    span = etree.Element(W + "gridSpan")
    span.set(W + "val", str(len(cells)))
    insert_in_order(properties, span, TCPR_ORDER)
    shading = etree.Element(W + "shd")
    shading.set(W + "val", "clear")
    shading.set(W + "color", "auto")
    shading.set(W + "fill", GROUP_ROW_FILL)
    insert_in_order(properties, shading, TCPR_ORDER)
    set_cell_text(label, text, bold=True, numbered=False, align="left")
    return row


def field_runs(instruction: str, cached: str, example: etree._Element | None) -> list[etree._Element]:
    """The runs of a Word field with `instruction` and the `cached` result."""
    def holder(child: etree._Element) -> etree._Element:
        result = etree.Element(W + "r")
        if example is not None:
            properties = example.find(W + "rPr")
            if properties is not None:
                result.append(deepcopy(properties))
        result.append(child)
        return result

    begin = etree.Element(W + "fldChar")
    begin.set(W + "fldCharType", "begin")
    # Word refreshes a dirty field when it repaginates, which is what keeps the
    # running header in step with the «Номер» row of the first page.
    begin.set(W + "dirty", "true")
    instruction_text = etree.Element(W + "instrText")
    instruction_text.set(XML_SPACE, "preserve")
    instruction_text.text = instruction
    separate = etree.Element(W + "fldChar")
    separate.set(W + "fldCharType", "separate")
    end = etree.Element(W + "fldChar")
    end.set(W + "fldCharType", "end")
    return [
        holder(begin),
        holder(instruction_text),
        holder(separate),
        run(cached, example),
        holder(end),
    ]



def _apply_title_and_header_bookmarks(
    body: etree._Element, header: etree._Element
) -> None:
    number_cell = customer_cell = object_cell = subject_cell = None
    for table in body.findall(".//w:tbl", NS):
        rows = table.findall("w:tr", NS)
        if len(rows) < 4:
            continue
        cells = [row.findall("w:tc", NS) for row in rows[:4]]
        if all(len(row) >= 2 for row in cells) and [word_text(row[0]) for row in cells[:3]] == [
            "Номер",
            "Заказчик",
            "Объект",
        ]:
            number_cell, customer_cell, object_cell, subject_cell = [row[1] for row in cells]
            break
    if number_cell is None:
        raise ValueError("Could not locate the customer details table")

    number_p = number_cell.find("w:p", NS)
    assert number_p is not None
    original_run = number_p.find("w:r", NS)
    for child in list(number_p):
        if child.tag != W + "pPr":
            number_p.remove(child)
    number_start, number_end = bookmark("Num", free_bookmark_id(xml_root(number_p)))
    number_p.extend(
        (number_start, run(NUMBER_PLACEHOLDER, original_run), number_end)
    )

    for cell, name in ((customer_cell, "Customer"), (object_cell, "Object")):
        p = cell.find("w:p", NS)
        assert p is not None
        for child in list(p):
            if child.tag != W + "pPr":
                p.remove(child)
        add_empty_bookmark(p, name, free_bookmark_id(xml_root(p)))

    if subject_cell is not None:
        subject_p = subject_cell.find("w:p", NS)
        assert subject_p is not None
        for child in list(subject_p):
            if child.tag != W + "pPr":
                subject_p.remove(child)
        subject_start, subject_end = bookmark("Subject", free_bookmark_id(xml_root(subject_p)))
        subject_p.extend(
            (
                subject_start,
                run(
                    "Поставка оборудования и выполнение работ по созданию АСУ ТП "
                    "на базе программно-технического комплекса АРКС400",
                    original_run,
                ),
                subject_end,
            )
        )

    running_header = next(
        (
            p
            for p in header.findall(".//w:p", NS)
            if "Технико-коммерческое предложение" in word_text(p)
        ),
        None,
    )
    if running_header is None:
        raise ValueError("Could not locate the running header")
    example = running_header.find("w:r", NS)
    for child in list(running_header):
        if child.tag != W + "pPr":
            running_header.remove(child)
    header_start, header_end = bookmark("HeaderNum", free_bookmark_id(header))
    # The running head repeats the number of the first page: a REF field on the
    # «Num» bookmark keeps it in step when the template is filled in by hand,
    # and the generators replace the whole bookmark with the sheet value.
    running_header.extend(
        (
            run("Технико-коммерческое предложение ", example),
            header_start,
            *field_runs(" REF Num \\* MERGEFORMAT ", NUMBER_PLACEHOLDER, example),
            header_end,
        )
    )


def _fix_notes_paragraph(notes: etree._Element) -> None:
    for t_el in notes.findall(".//w:t", NS):
        if t_el.text and t_el.text.startswith("римечания к Приложению"):
            t_el.text = "П" + t_el.text


# ---------------------------------------------------------------------------
# Content blocks of the proposal. Every step is idempotent: it looks for its
# block first, so running the tool twice leaves the template identical.
#
# The steps that own a block rewrite it from the constants above — the blocks of
# the shortened page two, the site readiness list, the mounting note, the Gantt
# chart, the appendix numbering and the running head. Those are the parts the
# review rewrote, and the constants are where their wording lives. Steps that
# only add what is missing leave the rest of the document alone, so paragraphs
# edited by hand elsewhere survive: start from a pristine copy of the asset when
# re-wording one of those constants.
# ---------------------------------------------------------------------------
def rework_offer_terms(body: etree._Element) -> None:
    """Page two: what the offer rests on and what it excludes.

    The general note («may be refined») becomes «Исходные данные и допущения»
    naming the source document, and the list of assumptions becomes «Не входит в
    объём предложения». A «Обязательства Заказчика» block of an earlier revision
    is removed: Appendix 5 states the customer's duties in detail, and the
    shortened page two does not repeat them.
    """
    basis = section_heading(body, BASIS_HEADING, LEGACY_NOTE_HEADING)
    if basis is None:
        raise ValueError(f"Could not locate the heading {LEGACY_NOTE_HEADING!r}")
    intro = block_intro(basis)
    if intro is None:
        raise ValueError("Could not locate the body of the note on page two")
    sync_list_block(basis, BASIS_INTRO, BASIS_ITEMS, intro)

    exclusions = section_heading(body, EXCLUSIONS_HEADING, LEGACY_EXCLUSIONS_HEADING)
    if exclusions is None:
        exclusions = paragraph_text_of(basis, EXCLUSIONS_HEADING, bold=True)
        last_of_block(basis).addnext(exclusions)
    sync_list_block(exclusions, EXCLUSIONS_INTRO, EXCLUSIONS_ITEMS, intro)

    duties = section_heading(body, LEGACY_DUTIES_HEADING)
    if duties is not None:
        removed = [duties, *block_items(duties)]
        intro_paragraph = block_intro(duties)
        if intro_paragraph is not None:
            removed.insert(1, intro_paragraph)
        for paragraph in removed:
            if paragraph.getparent() is not None:
                paragraph.getparent().remove(paragraph)


def add_project_duration(body: etree._Element) -> None:
    """The parameters table: how long the whole project takes."""
    table = find_table(body, DURATION_ANCHOR)
    rows = table.findall(W + "tr")
    if any(DURATION_LABEL in word_text(row) for row in rows):
        return
    totals = next(
        row for row in rows if word_text(row).strip().startswith(DURATION_ANCHOR)
    )
    row = drop_para_ids(deepcopy(totals.getprevious()))
    cells = row.findall(W + "tc")
    set_cell_text(cells[0], "")
    set_cell_text(cells[1], DURATION_LABEL, bold=True)
    set_cell_text(cells[2], DURATION_VALUE, bold=True)
    set_cell_text(cells[3], "", bold=True)
    totals.addprevious(row)


def reword_opening(body: etree._Element, old: str, new: str) -> None:
    """Replace the opening words of a paragraph, keeping its remaining runs."""
    if has_paragraph(body, new):
        return
    paragraph = find_paragraph(body, old)
    for value in paragraph.iter(W + "t"):
        if value.text and value.text.startswith(old):
            value.text = new + value.text[len(old) :]
            return
    raise ValueError(f"Could not reword the paragraph starting with {old!r}")


def update_terms_wording(body: etree._Element) -> None:
    """Payment, delivery and certification wordings from the review."""
    for anchor, text in PAYMENTS:
        paragraph = find_paragraph(body, anchor)
        if word_text(paragraph).strip() != text:
            set_paragraph_text(paragraph, text)
    for old, new in REWORDINGS:
        reword_opening(body, old, new)


def extend_appendix_list(body: etree._Element) -> None:
    """The list of appendices ends with the two reference appendices."""
    table = find_table(body, APPENDIX_LIST_ANCHOR)
    titles = {title for title, _ in REFERENCE_APPENDICES}
    for row in table.findall(W + "tr"):
        label = row_label(row)
        numbered = row.find(f".//{W}numPr") is not None
        # An earlier revision listed the pages as unnumbered reference material.
        if label == LEGACY_INFORMATION_GROUP or (label in titles and not numbered):
            table.remove(row)
    labels = {row_label(row) for row in table.findall(W + "tr")}
    example = table.findall(W + "tr")[-1]
    for title, _ in REFERENCE_APPENDICES:
        if title in labels:
            continue
        table.append(item_row(example, title))
        labels.add(title)


def is_appendix_label(paragraph: etree._Element) -> bool:
    """A «Приложение N» label: its number is a SEQ field, not typed in."""
    return any(
        "SEQ App" in (instruction.text or "")
        for instruction in paragraph.iter(W + "instrText")
    )


def appendix_number(body: etree._Element, target: etree._Element) -> str:
    """The number an appendix label at `target`'s place in the document shows."""
    position = body.index(target)
    before = sum(
        1
        for paragraph in body.findall(W + "p")
        if is_appendix_label(paragraph) and body.index(paragraph) < position
    )
    return str(before + 1)


def cached_appendix_number(label: etree._Element) -> etree._Element:
    """The «w:t» of the label that holds the number until Word recalculates it."""
    return [value for value in label.iter(W + "t")][-1]


def appendix_label(example: etree._Element, number: str) -> etree._Element:
    """A copy of the label `example`, showing `number` as the cached result."""
    label = drop_para_ids(deepcopy(example))
    for marker in label.iter(W + "lastRenderedPageBreak"):
        marker.getparent().remove(marker)
    cached_appendix_number(label).text = number
    return label


def add_reference_appendix_labels(body: etree._Element) -> None:
    """Label the reference pages as Приложение 9 and Приложение 10."""
    example = next(
        paragraph for paragraph in body.findall(W + "p") if is_appendix_label(paragraph)
    )
    for _, anchor in REFERENCE_APPENDICES:
        title = find_paragraph(body, anchor)
        label = title.getprevious()
        if label is not None and is_appendix_label(label):
            # Count the labels before the label itself, so its own number is not
            # counted twice on the next run.
            wanted = appendix_number(body, label)
            cached = cached_appendix_number(label)
            if cached.text != wanted:
                cached.text = wanted
            continue
        title.addprevious(appendix_label(example, appendix_number(body, title)))


def add_gost34_documents(body: etree._Element) -> None:
    """Appendix 4: the document set, the handover documents, and no duplicates.

    The step owns its rows: anything the constants list is put back in order, the
    two general rows that repeated the specific items are dropped, and the shaded
    group row of an earlier revision («РКД и ЭД по ГОСТ 34») is removed so the
    list starts at 1.
    """
    table = find_table(body, DOCUMENTS_ANCHOR)
    for cell in table.iter(W + "tc"):
        paragraphs = cell.findall(W + "p")
        if len(paragraphs) < 2:
            continue
        for paragraph in paragraphs[1:]:
            if word_text(paragraph).strip() == DOCUMENTS_STAGE_NOTE:
                cell.remove(paragraph)

    for row in table.findall(W + "tr"):
        if word_text(row).strip() == LEGACY_GROUP_ROW:
            table.remove(row)

    rows = table.findall(W + "tr")
    example = next(row for row in rows[1:] if len(row.findall(W + "tc")) > 1)
    anchor = rows[0]
    for text in RKD_DOCUMENTS:
        match = next(
            (row for row in table.findall(W + "tr") if row_label(row) == text), None
        )
        if match is not None:
            anchor = match
            continue
        row = item_row(example, text)
        anchor.addnext(row)
        anchor = row

    labels = {row_label(row) for row in table.findall(W + "tr")}
    for text in HANDOVER_DOCUMENTS:
        if text not in labels:
            table.append(item_row(example, text))
            labels.add(text)
    for row in table.findall(W + "tr"):
        if row_label(row) in DUPLICATE_DOCUMENTS:
            table.remove(row)


def add_mounting_scope(body: etree._Element) -> None:
    """Appendix 5: extra works, site readiness requirements, a stronger note.

    The readiness block has no heading: an introductory sentence and its bullets
    repeat on every run, so a hand edit cannot leave two copies behind.
    """
    table = find_table(body, MOUNTING_ANCHOR)
    rows = table.findall(W + "tr")
    example = rows[1]
    labels = [row_label(row) for row in rows]

    wanted = [(MOUNTING_FIRST_WORK, MOUNTING_ANCHOR)]
    wanted += [(text, None) for text in MOUNTING_EXTRA_WORKS]
    for text, anchor in wanted:
        if text in labels:
            continue
        row = item_row(example, text)
        if anchor is not None and anchor in labels:
            table.insert(list(table).index(rows[labels.index(anchor)]), row)
        else:
            table.append(row)
        labels.append(text)

    note = find_paragraph(body, MOUNTING_NOTE_ANCHOR)
    if word_text(note).strip() != MOUNTING_NOTE:
        set_paragraph_text(note, MOUNTING_NOTE)

    for heading in [
        paragraph
        for paragraph in body.findall(W + "p")
        if word_text(paragraph).strip() == LEGACY_SITE_READINESS_HEADING
    ]:
        body.remove(heading)

    intros = [
        paragraph
        for paragraph in body.findall(W + "p")
        if word_text(paragraph).strip() == SITE_READINESS_INTRO
    ]
    if not intros:
        intro = paragraph_text_of(note, SITE_READINESS_INTRO)
        note.addprevious(intro)
        intros = [intro]
    # A previous revision kept the nine-item list, and a run that added the list
    # without finding its heading left a second copy behind: keep one block.
    for extra in intros[1:]:
        following = extra.getnext()
        while following is not None and following.tag == W + "p" and is_item_bullet(
            following
        ):
            following.getparent().remove(following)
            following = extra.getnext()
        body.remove(extra)
    intro = intros[0]
    current = []
    following = intro.getnext()
    while following is not None and following.tag == W + "p" and is_item_bullet(
        following
    ):
        current.append(following)
        following = following.getnext()
    for paragraph, text in zip(current, SITE_READINESS, strict=False):
        if word_text(paragraph).strip() != text:
            set_paragraph_text(paragraph, text)
    for extra in current[len(SITE_READINESS) :]:
        extra.getparent().remove(extra)
    kept = current[: len(SITE_READINESS)]
    anchor = kept[-1] if kept else intro
    for text in SITE_READINESS[len(kept) :]:
        bullet = bullet_paragraph(text)
        anchor.addnext(bullet)
        anchor = bullet


def add_terms_page(body: etree._Element) -> None:
    """The final page: terms and abbreviations used in the proposal."""
    if has_paragraph(body, TERMS_TITLE):
        return
    heading = paragraph_text_of(
        find_paragraph(body, "Приложения", exact=True), TERMS_TITLE
    )
    insert_in_order(heading.find(W + "pPr"), etree.Element(W + "pageBreakBefore"), PPR_ORDER)
    intro = paragraph_text_of(find_paragraph(body, BASIS_INTRO), TERMS_INTRO)
    section = body.find(W + "sectPr")
    if section is None:
        raise ValueError("Could not locate the final section properties")
    position = body.index(section)
    for offset, element in enumerate((heading, intro, terms_table(body))):
        body.insert(position + offset, element)


def terms_table(body: etree._Element) -> etree._Element:
    """The abbreviation table, built on the appendix table look."""
    table = drop_para_ids(deepcopy(find_table(body, DOCUMENTS_ANCHOR)))
    rows = table.findall(W + "tr")
    header = rows[0]
    example = next(row for row in rows[1:] if len(row.findall(W + "tc")) > 1)
    for row in rows[1:]:
        table.remove(row)
    header_cells = header.findall(W + "tc")
    for cell, title in zip(header_cells, TERMS_COLUMNS, strict=True):
        set_cell_text(cell, title)
    for abbreviation, meaning in TERMS:
        row = drop_para_ids(deepcopy(example))
        cells = row.findall(W + "tc")
        set_cell_text(cells[0], abbreviation, bold=True, numbered=False, align="left")
        set_cell_text(cells[1], meaning, numbered=False, align="left")
        table.append(row)

    columns = table.find(W + "tblGrid").findall(W + "gridCol")
    for column, width in zip(columns, TERMS_COLUMN_TWIPS, strict=True):
        column.set(W + "w", width)
    for row in table.findall(W + "tr"):
        for cell, share in zip(row.findall(W + "tc"), TERMS_COLUMN_PCT, strict=True):
            properties = cell.find(W + "tcPr")
            for old in properties.findall(W + "tcW"):
                properties.remove(old)
            width = etree.Element(W + "tcW")
            width.set(W + "w", share)
            width.set(W + "type", "pct")
            insert_in_order(properties, width, TCPR_ORDER)
    return table


def gantt_cell(
    width: int,
    *,
    text: str = "",
    fill: str | None = None,
    bold: bool = False,
    color: str | None = None,
    size: str | None = None,
    span: int | None = None,
    align: str = "center",
) -> etree._Element:
    """One cell of the chart: a stage name, a scale mark or a piece of a bar."""
    cell = etree.Element(W + "tc")
    properties = etree.SubElement(cell, W + "tcPr")
    box = etree.SubElement(properties, W + "tcW")
    box.set(W + "w", str(width))
    box.set(W + "type", "dxa")
    if span is not None:
        merged = etree.SubElement(properties, W + "gridSpan")
        merged.set(W + "val", str(span))
    borders = etree.SubElement(properties, W + "tcBorders")
    for side in ("top", "left", "bottom", "right"):
        line = etree.SubElement(borders, W + side)
        line.set(W + "val", "single")
        line.set(W + "sz", "4")
        line.set(W + "space", "0")
        line.set(W + "color", GRID_LINE)
    if fill is not None:
        shading = etree.SubElement(properties, W + "shd")
        shading.set(W + "val", "clear")
        shading.set(W + "color", "auto")
        shading.set(W + "fill", fill)
    vertical = etree.SubElement(properties, W + "vAlign")
    vertical.set(W + "val", "center")

    paragraph = etree.SubElement(cell, W + "p")
    layout = etree.SubElement(paragraph, W + "pPr")
    spacing = etree.SubElement(layout, W + "spacing")
    spacing.set(W + "line", "254")
    spacing.set(W + "lineRule", "auto")
    if align == "left":
        # The week columns have no cell margins at all, so the stage names get
        # their own millimetre of air instead of touching the border.
        indent = etree.SubElement(layout, W + "ind")
        indent.set(W + "left", "57")
    etree.SubElement(layout, W + "contextualSpacing")
    justification = etree.SubElement(layout, W + "jc")
    justification.set(W + "val", align)
    mark = arial_rpr(bold)
    if color is not None:
        tone = etree.SubElement(mark, W + "color")
        tone.set(W + "val", color)
    if size is not None:
        for tag in ("sz", "szCs"):
            value = etree.SubElement(mark, W + tag)
            value.set(W + "val", size)
    layout.append(mark)
    if text:
        letter = etree.SubElement(paragraph, W + "r")
        letter.append(deepcopy(mark))
        etree.SubElement(letter, W + "t").text = text
    return cell


def gantt_row(cells: list[etree._Element]) -> etree._Element:
    row = etree.Element(W + "tr")
    properties = etree.SubElement(row, W + "trPr")
    height = etree.SubElement(properties, W + "trHeight")
    height.set(W + "val", "20")
    row.extend(cells)
    return row


def gantt_table() -> etree._Element:
    """The calendar plan of Appendix 8 drawn as bars, one column per week."""
    table = etree.Element(W + "tbl")
    properties = etree.SubElement(table, W + "tblPr")
    width = etree.SubElement(properties, W + "tblW")
    width.set(W + "w", str(GANTT_TABLE_TWIPS))
    width.set(W + "type", "dxa")
    indent = etree.SubElement(properties, W + "tblInd")
    indent.set(W + "w", "-5")
    indent.set(W + "type", "dxa")
    borders = etree.SubElement(properties, W + "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        line = etree.SubElement(borders, W + side)
        line.set(W + "val", "single")
        line.set(W + "sz", "4")
        line.set(W + "space", "0")
        line.set(W + "color", BAR_FILL)
    # A fixed layout keeps the 190-twip week columns exactly as wide as the grid
    # says, whatever the stage names in the first column wrap to.
    layout = etree.SubElement(properties, W + "tblLayout")
    layout.set(W + "type", "fixed")
    margins = etree.SubElement(properties, W + "tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        margin = etree.SubElement(margins, W + side)
        margin.set(W + "w", "0")
        margin.set(W + "type", "dxa")
    look = etree.SubElement(properties, W + "tblLook")
    look.set(W + "val", "04A0")
    look.set(W + "firstRow", "1")
    look.set(W + "lastRow", "0")
    look.set(W + "firstColumn", "0")
    look.set(W + "lastColumn", "0")
    look.set(W + "noHBand", "0")
    look.set(W + "noVBand", "1")

    grid = etree.SubElement(table, W + "tblGrid")
    for column in (GANTT_LABEL_TWIPS, *(GANTT_WEEK_TWIPS,) * GANTT_WEEKS):
        grid_column = etree.SubElement(grid, W + "gridCol")
        grid_column.set(W + "w", str(column))

    weeks = range(1, GANTT_WEEKS + 1)
    head = [
        gantt_cell(
            GANTT_LABEL_TWIPS,
            text="Этап",
            fill=HEAD_FILL,
            bold=True,
            color="FFFFFF",
            align="left",
        )
    ]
    head += [
        gantt_cell(
            GANTT_WEEK_TWIPS,
            text=str(week),
            fill=HEAD_FILL,
            bold=True,
            color="FFFFFF",
            size=GANTT_SCALE_SIZE,
        )
        for week in weeks
    ]
    table.append(gantt_row(head))

    for label, start, finish in GANTT_STAGES:
        cells = [gantt_cell(GANTT_LABEL_TWIPS, text=label, align="left")]
        cells += [
            gantt_cell(
                GANTT_WEEK_TWIPS,
                fill=BAR_FILL if start <= week <= finish else None,
            )
            for week in weeks
        ]
        table.append(gantt_row(cells))

    total = [
        gantt_cell(
            GANTT_LABEL_TWIPS,
            text=DURATION_LABEL,
            fill=GROUP_ROW_FILL,
            bold=True,
            align="left",
        )
    ]
    total.append(
        gantt_cell(
            GANTT_WEEK_TWIPS * GANTT_WEEKS,
            text=DURATION_VALUE,
            fill=GROUP_ROW_FILL,
            bold=True,
            span=GANTT_WEEKS,
        )
    )
    table.append(gantt_row(total))
    return table


def is_gantt_table(table: etree._Element) -> bool:
    """The chart, as opposed to the tables of the document: it starts with «Этап»."""
    rows = table.findall(W + "tr")
    if not rows:
        return False
    cells = rows[0].findall(W + "tc")
    return bool(cells) and word_text(cells[0]).strip() == "Этап"


def is_blank_paragraph(element: etree._Element) -> bool:
    """An empty spacer paragraph: a paragraph mark and its properties, nothing else."""
    if element.tag != W + "p" or word_text(element):
        return False
    return all(child.tag == W + "pPr" for child in element)


def blank_paragraph_example(body: etree._Element) -> etree._Element | None:
    """A spacer paragraph to clone, so the new one carries the same spacing.

    A paragraph that carries section properties is not a spacer: cloning it would
    copy a section break into the middle of the text.
    """
    for paragraph in body.findall(W + "p"):
        if is_blank_paragraph(paragraph) and paragraph.find(W + "pPr/" + W + "sectPr") is None:
            return paragraph
    return None


SECTPR_ORDER = (
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
    "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
    "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
)


def is_landscape(sectPr: etree._Element) -> bool:
    """Whether the section is printed with the sheet turned."""
    page = sectPr.find(W + "pgSz")
    if page is None:
        return False
    if page.get(W + "orient") == "landscape":
        return True
    return int(page.get(W + "w")) > int(page.get(W + "h"))


def set_page_orientation(sectPr: etree._Element, *, landscape: bool) -> None:
    """Turn the section, keeping its paper size, margins and everything else."""
    page = sectPr.find(W + "pgSz")
    if page is None:
        page = etree.Element(W + "pgSz")
        _insert_in_order(sectPr, page, SECTPR_ORDER)
    width = int(page.get(W + "w"))
    height = int(page.get(W + "h"))
    if landscape != (width > height):
        page.set(W + "w", str(height))
        page.set(W + "h", str(width))
    if landscape:
        page.set(W + "orient", "landscape")
    else:
        page.attrib.pop(W + "orient", None)


def set_section_type(sectPr: etree._Element, value: str) -> None:
    """Say how the section starts: a new page, a new odd page, or no break."""
    for old in sectPr.findall(W + "type"):
        sectPr.remove(old)
    node = etree.Element(W + "type")
    node.set(W + "val", value)
    _insert_in_order(sectPr, node, SECTPR_ORDER)


def _insert_in_order(parent: etree._Element, child: etree._Element, order: tuple[str, ...]) -> None:
    """Put `child` before the first sibling the schema wants after it."""
    rank = order.index(etree.QName(child).localname)
    for index, sibling in enumerate(parent):
        name = etree.QName(sibling).localname
        if name in order and order.index(name) > rank:
            parent.insert(index, child)
            return
    parent.append(child)


def section_break_paragraph(body: etree._Element, *, landscape: bool) -> etree._Element:
    """An empty paragraph that closes a section of the requested orientation.

    Word stores the properties of a section on the last paragraph of that
    section (and the last section of the document in `w:body`), so a break is a
    paragraph mark carrying `w:sectPr`, not an element of its own.
    """
    source = body.find(W + "sectPr")
    if source is None:
        raise ValueError("The template has no section properties")
    example = blank_paragraph_example(body)
    paragraph = etree.Element(W + "p") if example is None else drop_para_ids(deepcopy(example))
    properties = paragraph.find(W + "pPr")
    if properties is None:
        properties = etree.Element(W + "pPr")
        paragraph.insert(0, properties)
    for old in properties.findall(W + "sectPr"):
        properties.remove(old)
    sectPr = deepcopy(source)
    set_page_orientation(sectPr, landscape=landscape)
    set_section_type(sectPr, GANTT_SECTION_TYPE)
    properties.append(sectPr)
    return paragraph


def add_gantt(body: etree._Element) -> None:
    """Appendix 8: the same plan as a bar chart, on a landscape page of its own.

    The chart is generated data — one cell per week, no prose — so this step owns
    both the chart and the section breaks around it: the chart in place is
    rebuilt from the constants, which also repairs the layout of an earlier
    revision (a spare column at the right edge, a total row merged one column too
    wide). A template of an earlier revision may still carry the heading and the
    caption of that revision; both go.

    The grid is 14574 twips wide against a 9780-twip portrait text area, so the
    chart is fenced off by a section break before it and another one after it:
    the appendix keeps its portrait page up to the table of stages, the chart
    gets a landscape page, and Appendix 9 starts on a portrait page again.
    """
    plan = find_table(body, GANTT_ANCHOR)

    # Everything between the table of stages and the next block of the document
    # belongs to the chart: the chart itself, the breaks around it, the spacers.
    # All of it is rebuilt below.
    position = body.index(plan) + 1
    while position < len(body):
        element = body[position]
        kind = etree.QName(element).localname
        if kind == "tbl" and is_gantt_table(element):
            body.remove(element)
            continue
        if kind == "p" and (
            element.find(W + "pPr/" + W + "sectPr") is not None
            or is_blank_paragraph(element)
        ):
            body.remove(element)
            continue
        break

    body.insert(position, section_break_paragraph(body, landscape=False))
    position += 1
    body.insert(position, gantt_table())
    position += 1
    body.insert(position, section_break_paragraph(body, landscape=True))

    stale = {LEGACY_GANTT_HEADING, LEGACY_GANTT_CAPTION}
    for paragraph in body.findall(W + "p"):
        if word_text(paragraph).strip() in stale:
            body.remove(paragraph)


def is_inside(element: etree._Element, ancestor: etree._Element) -> bool:
    """Whether `element` sits anywhere below `ancestor`."""
    parent = element.getparent()
    while parent is not None:
        if parent is ancestor:
            return True
        parent = parent.getparent()
    return False


def raise_text_sizes(
    element: etree._Element, *, keep_inside: tuple[etree._Element, ...] = ()
) -> int:
    """Bring every 10 pt / 11 pt run up to the 12 pt house size.

    Everything below `keep_inside` keeps its own size: the calendar-plan table
    of Appendix 8 is deliberately set at 11 pt.
    """
    changed = 0
    for tag in ("sz", "szCs"):
        for size in element.iter(W + tag):
            if any(is_inside(size, kept) for kept in keep_inside):
                continue
            if size.get(W + "val") in RETIRED_SIZES:
                size.set(W + "val", HOUSE_SIZE)
                changed += 1
    return changed


def calendar_plan_tables(document: etree._Element) -> tuple[etree._Element, ...]:
    """The table of stages in Appendix 8: the one with the «Плановая дата» column."""
    return tuple(
        table
        for table in document.findall(f".//{W}tbl")
        if CALENDAR_PLAN_ANCHOR in word_text(table)
    )


def apply_font_policy(
    document: etree._Element, styles: etree._Element, numbering: etree._Element
) -> None:
    """Apply the 12 pt rule to the text, its numbering and the styles it uses.

    Headers and footers are left alone: they carry their own small sizes, and
    the running head is deliberately quieter than the body text. So is the
    calendar-plan table of Appendix 8, which is set at 11 pt on purpose.
    """
    raise_text_sizes(document, keep_inside=calendar_plan_tables(document))

    used_numbers = {number.get(W + "val") for number in document.iter(W + "numId")}
    used_abstract = {
        number.find(W + "abstractNumId").get(W + "val")
        for number in numbering.findall(W + "num")
        if number.get(W + "numId") in used_numbers
    }
    for abstract in numbering.findall(W + "abstractNum"):
        if abstract.get(W + "abstractNumId") in used_abstract:
            raise_text_sizes(abstract)

    used_styles = set()
    for tag in ("pStyle", "rStyle", "tblStyle"):
        used_styles.update(style.get(W + "val") for style in document.iter(W + tag))
    for style in styles.findall(W + "style"):
        if style.get(W + "styleId") in used_styles:
            link = style.find(W + "link")
            if link is not None:
                used_styles.add(link.get(W + "val"))
    for style in styles.findall(W + "style"):
        if style.get(W + "styleId") in used_styles:
            raise_text_sizes(style)
    defaults = styles.find(W + "docDefaults")
    if defaults is not None:
        raise_text_sizes(defaults)


def _set_document_title(parts: dict[str, bytes]) -> None:
    """Give the template the «ТКП» title property.

    Word uses this property as the document name when a document is created
    from the template, so the new document is called «ТКП» while the visible
    title on the first page keeps its own wording.
    """
    core = etree.fromstring(parts["docProps/core.xml"])
    title = core.find(f"{{{DC_NS}}}title")
    if title is None:
        title = etree.SubElement(core, f"{{{DC_NS}}}title")
    title.text = DOCUMENT_TITLE
    parts["docProps/core.xml"] = write_xml(core)

    app = etree.fromstring(parts["docProps/app.xml"])
    vector = app.find(f"{{{EP_NS}}}TitlesOfParts/{{{VT_NS}}}vector")
    if vector is not None:
        first = vector.find(f"{{{VT_NS}}}lpstr")
        if first is not None:
            first.text = DOCUMENT_TITLE
            parts["docProps/app.xml"] = write_xml(app)


def patch_template() -> None:
    """Update the single ТКП_шаблон.dotx file in place."""
    with ZipFile(TEMPLATE_DOTX) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        header = etree.fromstring(archive.read("word/header1.xml"))
        settings = etree.fromstring(archive.read("word/settings.xml"))
        styles = etree.fromstring(archive.read("word/styles.xml"))
        numbering = etree.fromstring(archive.read("word/numbering.xml"))
        core = archive.read("docProps/core.xml")
        app = archive.read("docProps/app.xml")

    for attached in settings.findall(W + "attachedTemplate"):
        settings.remove(attached)

    body = document.find(".//w:body", NS)
    assert body is not None
    note_paragraphs = [
        p for p in body.findall("w:p", NS) if "римечания к Приложению" in word_text(p)
    ]
    if len(note_paragraphs) != 1:
        raise ValueError("Could not locate the notes after Appendix 3")
    notes = note_paragraphs[0]
    _fix_notes_paragraph(notes)

    spec_start = document.xpath(".//w:bookmarkStart[@w:name='Spec_Table']", namespaces=NS)[0]
    old_spec_paragraph = spec_start.getparent()
    for name in ("Num", "Customer", "Object", "Spec_Table"):
        remove_bookmark(document, name)
    if document.xpath(".//w:bookmarkStart[@w:name='Subject']", namespaces=NS):
        remove_bookmark(document, "Subject")
    if (
        old_spec_paragraph is not notes
        and not word_text(old_spec_paragraph)
        and old_spec_paragraph.getparent() is body
    ):
        body.remove(old_spec_paragraph)

    _apply_title_and_header_bookmarks(body, header)

    spec_p = etree.Element(W + "p")
    add_empty_bookmark(spec_p, "Spec_Table", free_bookmark_id(document))
    body.insert(body.index(notes), spec_p)

    rework_offer_terms(body)
    add_project_duration(body)
    update_terms_wording(body)
    extend_appendix_list(body)
    add_reference_appendix_labels(body)
    add_gost34_documents(body)
    add_mounting_scope(body)
    add_gantt(body)
    add_terms_page(body)
    # Last, so the paragraphs the content steps cloned are normalised too.
    apply_font_policy(document, styles, numbering)
    drop_unpaired_bookmarks(document)

    parts = {
        "word/document.xml": write_xml(document),
        "word/header1.xml": write_xml(header),
        "word/settings.xml": write_xml(settings),
        "word/styles.xml": write_xml(styles),
        "word/numbering.xml": write_xml(numbering),
        "docProps/core.xml": core,
        "docProps/app.xml": app,
    }
    _set_document_title(parts)

    replace_zip_part(
        TEMPLATE_DOTX,
        parts,
        remove_members={"word/_rels/settings.xml.rels"},
    )


def inline_cell(address: str, text: str) -> etree._Element:
    cell = etree.Element(X + "c", r=address, t="inlineStr")
    etree.SubElement(etree.SubElement(cell, X + "is"), X + "t").text = text
    return cell


def preview_sheet_part(sheet_name: str = PREVIEW_SHEET) -> str:
    """Return the ZIP part holding `sheet_name`, resolved through the workbook rels.

    Addressing the part by index (for example sheet26.xml) would silently edit an
    unrelated worksheet as soon as a sheet is inserted, removed or reordered.
    """
    with ZipFile(WORKBOOK) as archive:
        book = etree.fromstring(archive.read("xl/workbook.xml"))
        rels = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    sheets = book.xpath(
        ".//x:sheet[@name=$name]", namespaces={"x": X_NS}, name=sheet_name
    )
    if len(sheets) != 1:
        raise ValueError(f"Expected one worksheet named {sheet_name!r}, found {len(sheets)}")
    rel_id = sheets[0].get(f"{{{R_NS}}}id")
    targets = rels.xpath(
        ".//r:Relationship[@Id=$rel_id]", namespaces={"r": PKG_REL_NS}, rel_id=rel_id
    )
    if len(targets) != 1:
        raise ValueError(f"No relationship {rel_id!r} for worksheet {sheet_name!r}")
    target = targets[0].get("Target").lstrip("/")
    return posixpath.normpath(
        target if target.startswith("xl/") else posixpath.join("xl", target)
    )


def patch_preview_sheet() -> None:
    """Initialize the preview and keep its helper marker column hidden."""
    part = preview_sheet_part()
    with ZipFile(WORKBOOK) as archive:
        root = etree.fromstring(archive.read(part))
    data = root.find(X + "sheetData")
    dimension = root.find(X + "dimension")
    assert data is not None and dimension is not None

    if dimension.get("ref") == "A1:G3":
        header = root.xpath(".//x:c[@r='A1']", namespaces={"x": X_NS})
        if header:
            header[0].clear()
            header[0].set("r", "A1")
            header[0].set("t", "inlineStr")
            header[0].append(etree.Element(X + "is"))
            etree.SubElement(header[0][0], X + "t").text = "№ п/п"
        columns = root.find(X + "cols")
        if columns is not None:
            for col in columns.findall(X + "col"):
                if col.get("min") == "6" and col.get("max") == "6":
                    col.set("hidden", "1")
        replace_zip_part(WORKBOOK, {part: write_xml(root)})
        return

    if len(data) != 1 or data[0].get("r") != "1":
        raise ValueError("The preview sheet is not empty; refusing to erase it")
    data.remove(data[0])
    hints = (
        "Alt+F8: СформироватьТКП (Word) или ОбновитьСпецификациюТКП (предпросмотр)",
        "Сначала заполните позиции и количество на листах 1-17, Spare parts и SW(w.o.VAT).",
        "Берутся C/D/F/G и только количество > 0; цены, итоги и работы не переносятся.",
    )
    first = etree.SubElement(data, X + "row", r="1", spans="1:7")
    for letter, value in zip(
        "ABCDE",
        ("№ п/п", "Наименование и техническая характеристика", "Артикул", "Ед. изм.", "Кол-во"),
        strict=True,
    ):
        first.append(inline_cell(letter + "1", value))
    first.append(inline_cell("G1", hints[0]))
    for row_num in (2, 3):
        row = etree.SubElement(data, X + "row", r=str(row_num), spans="7:7")
        row.append(inline_cell(f"G{row_num}", hints[row_num - 1]))
    dimension.set("ref", "A1:G3")
    sheet_format = root.find(X + "sheetFormatPr")
    assert sheet_format is not None
    columns = etree.Element(X + "cols")
    for col_no, width in ((1, 7), (2, 65), (3, 27), (4, 13), (5, 12), (6, 3), (7, 78)):
        attributes = {
            "min": str(col_no),
            "max": str(col_no),
            "width": str(width),
            "customWidth": "1",
        }
        if col_no == 6:
            attributes["hidden"] = "1"
        etree.SubElement(columns, X + "col", **attributes)
    root.insert(root.index(sheet_format) + 1, columns)
    replace_zip_part(WORKBOOK, {part: write_xml(root)})


def embed_macro() -> None:
    source = (ROOT / "Module1.bas").read_text(encoding="utf-8").replace("\r\n", "\n")
    source = source.replace("\n", "\r\n")
    with ExcelFile(WORKBOOK) as book:
        if book.get_module("Module1").replace("\r\n", "\n").rstrip() != source.replace("\r\n", "\n").rstrip():
            book.set_module("Module1", source)
            book.save(allow_protected=True)
        if book.get_module("Module1").replace("\r\n", "\n").rstrip() != source.replace("\r\n", "\n").rstrip():
            raise ValueError("Embedded VBA source does not match Module1.bas")


if __name__ == "__main__":
    patch_template()
    patch_preview_sheet()
    embed_macro()
    print("Updated the .dotx template, Excel preview sheet, and embedded VBA macro.")
