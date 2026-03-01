import unittest
from file_management import interpret_txt_data

class TestInterpretingTXT(unittest.TestCase):

    def test_basic_example(self):
        lines = [
            "#hello",
            "epic=",
            "[special]",
            "text  #comment",
            "# another",
            "[wow]",
            "ok"
        ]
        expected = [
            {
                "key": None,
                "content": [
                    {"text": "", "comment": "hello", "extra": ""}
                ]
            },
            {
                "key": "epic",
                "content": [
                    {"text": "", "comment": "", "extra": ""},
                    {"text": "", "comment": "", "extra": "[special]"},
                    {"text": "text", "comment": "comment", "extra": ""},
                    {"text": "", "comment": "another", "extra": ""},
                    {"text": "", "comment": "", "extra": "[wow]"},
                    {"text": "ok", "comment": "", "extra": ""}
                ]
            }
        ]
        result = interpret_txt_data(lines)
        self.assertEqual(expected, result)

    def test_empty_and_none_key(self):
        lines = [
            "", 
            "# just a comment",
            "[extra]",
            "text"
        ]
        expected = [
            {
                "key": None,
                "content": [
                    {"text": "", "comment": "", "extra": ""},
                    {"text": "", "comment": "just a comment", "extra": ""},
                    {"text": "", "comment": "", "extra": "[extra]"},
                    {"text": "text", "comment": "", "extra": ""}
                ]
            }
        ]
        result = interpret_txt_data(lines)
        self.assertEqual(expected, result)

    def test_key_with_value_and_comment(self):
        lines = ["mykey= value # this is a comment", "[meta]"]
        expected = [
            {
                "key": "mykey",
                "content": [
                    {"text": "value", "comment": "this is a comment", "extra": ""},
                    {"text": "", "comment": "", "extra": "[meta]"}
                ]
            }
        ]
        result = interpret_txt_data(lines)
        self.assertEqual(expected, result)

    def test_multiple_keys(self):
        lines = ["k1=one", "two #c", "k2=", "[special]"]
        expected = [
            {
                "key": "k1",
                "content": [
                    {"text": "one", "comment": "", "extra": ""},
                    {"text": "two", "comment": "c", "extra": ""}
                ]
            },
            {
                "key": "k2",
                "content": [
                    {"text": "", "comment": "", "extra": ""},
                    {"text": "", "comment": "", "extra": "[special]"}
                ]
            }
        ]
        result = interpret_txt_data(lines)
        self.assertEqual(expected, result)

    def test_key_with_no_value(self):
        lines = ["emptykey="]
        expected = [
            {
                "key": "emptykey",
                "content": [
                    {"text": "", "comment": "", "extra": ""}
                ]
            }
        ]
        result = interpret_txt_data(lines)
        self.assertEqual(expected, result)

if __name__ == '__main__':
    unittest.main()
