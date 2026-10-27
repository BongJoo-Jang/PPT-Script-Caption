import unittest
from translations import STRINGS, tr


class TranslationsTests(unittest.TestCase):
    def test_complete_translations(self):
        self.assertEqual(set(STRINGS["ko"]), set(STRINGS["en"]))

    def test_placeholders_match(self):
        from string import Formatter
        fmt = Formatter()
        for key, ko in STRINGS["ko"].items():
            en = STRINGS["en"][key]
            ko_args = {field for _, field, _, _ in fmt.parse(ko) if field is not None}
            en_args = {field for _, field, _, _ in fmt.parse(en) if field is not None}
            self.assertEqual(ko_args, en_args, key)

    def test_t(self):
        self.assertEqual(tr("en", "language_label"), "Language / 언어")
        self.assertIn("35px", tr("ko", "font_label", size=35))


if __name__ == "__main__":
    unittest.main()