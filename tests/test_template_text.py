"""Static text checks for ТКП_шаблон.dotx.

The template is the single source of customer-facing wording: everything in the
generated proposal is copied from it verbatim. These checks keep the silent
corruptions out — text pasted from a web page or another language keyboard
arrives with Latin look-alike letters and non-breaking hyphens inside Russian
words, which look identical on screen but break search, spell-check and
copy-paste, and are then mailed to a customer.
"""

import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "ТКП_шаблон.dotx"

# Genuine mixed-script tokens: certificate numbers where a Cyrillic letter is
# part of the official identifier. Anything else that mixes scripts is a typo.
ALLOWED_MIXED_SCRIPT = frozenset({"RU.AЖ49.H04589"})

# Characters that must not appear: non-breaking and soft hyphens, zero-width
# space and BOM. A normal U+00A0 is allowed — it is used deliberately after «№».
FORBIDDEN_CHARS = {
    "\u2010": "HYPHEN (U+2010)",
    "\u2011": "NON-BREAKING HYPHEN (U+2011)",
    "\u00ad": "SOFT HYPHEN (U+00AD)",
    "\u200b": "ZERO WIDTH SPACE (U+200B)",
    "\ufeff": "ZERO WIDTH NO-BREAK SPACE (U+FEFF)",
}

# Wording that must never come back. Each entry was an actual defect found in the
# shipped template (see TEMPLATE_REVIEW.md).
FORBIDDEN_PHRASES = (
    "Объeм",
    "выдаeтся",
    "приeмо",
    "надeжность",
    "счeт",
    "трeхуровневой",
    "Комплексного опробование",
    "Акт о завершении ПНР",
    "третьих сторон и лиц",
    "собственности Заказчику",
    "Оплата работы по пунктам",
    "стоимости работы по этапу",
    "Приемо-сдаточные",
    "данных, Спецификация",
    # A third-party office address used to be hard-coded as the delivery point.
    "Олимпийский проспект",
    "16с1",
)


def document_parts() -> dict[str, etree._Element]:
    """Return every WordprocessingML part that carries text."""
    with ZipFile(TEMPLATE) as archive:
        names = [
            name
            for name in archive.namelist()
            if name.startswith("word/") and name.endswith(".xml") and name.split("/")[-1].split(".")[0] in
            {"document", "header1", "header2", "footer1"}
        ]
        return {name: etree.fromstring(archive.read(name)) for name in names}


def part_text(root: etree._Element) -> str:
    return "\n".join(
        "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
        for p in root.iter(f"{{{W}}}p")
    )


class TemplateTextTest(unittest.TestCase):
    def setUp(self):
        self.assertTrue(TEMPLATE.exists(), f"{TEMPLATE.name} not found")
        self.parts = document_parts()
        self.text = "\n".join(part_text(root) for root in self.parts.values())

    def test_no_look_alike_latin_letters_inside_russian_words(self):
        """A Latin "e" pasted into a Russian word is invisible in print but breaks search."""
        offenders = []
        for word in self.text.split():
            if word in ALLOWED_MIXED_SCRIPT:
                continue
            has_cyrillic = any("а" <= ch.lower() <= "я" or ch in "ёЁ" for ch in word)
            has_latin = any("a" <= ch.lower() <= "z" for ch in word)
            if has_cyrillic and has_latin:
                offenders.append(word)
        self.assertEqual(offenders, [], f"Latin look-alike letters inside Russian words: {offenders}")

    def test_no_forbidden_typographic_characters(self):
        found = {
            label: self.text.count(char)
            for char, label in FORBIDDEN_CHARS.items()
            if char in self.text
        }
        self.assertEqual(found, {}, f"Forbidden characters in the template: {found}")

    def test_known_wording_defects_are_gone(self):
        present = [phrase for phrase in FORBIDDEN_PHRASES if phrase in self.text]
        self.assertEqual(present, [], f"Regression in template wording: {present}")

    def test_delivery_address_stays_fillable(self):
        """The delivery point must be a placeholder, not somebody else's office."""
        self.assertIn("по адресу: ____________", self.text)

    def test_appendix_headings_are_all_present(self):
        headings = [f"Приложение {number}" for number in range(1, 9)]
        for heading in headings:
            self.assertGreaterEqual(
                self.text.count(heading), 1, f"{heading!r} missing from the template"
            )


if __name__ == "__main__":
    unittest.main()
