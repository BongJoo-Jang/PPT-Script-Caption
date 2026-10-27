import unittest
from caption_core import split_captions


class CaptionSplitterTests(unittest.TestCase):
    def test_explicit_segments(self):
        self.assertEqual(split_captions("첫 문장. // 둘째 문장. // 끝."),
                         ["첫 문장.", "둘째 문장.", "끝."])

    def test_line_breaks(self):
        self.assertEqual(split_captions("Hello.\nWelcome to the talk."),
                         ["Hello.", "Welcome to the talk."])

    def test_empty(self):
        self.assertEqual(split_captions(" \n "), [])

    def test_sentence_boundaries(self):
        self.assertEqual(split_captions("Hello! Now the results? Great."),
                         ["Hello!", "Now the results?", "Great."])


if __name__ == "__main__":
    unittest.main()