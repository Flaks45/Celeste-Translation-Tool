import os
import tempfile
import unittest
from openpyxl import Workbook
from file_management import *

class TestCheckExtension(unittest.TestCase):
    def test_single_extension_match(self):
        self.assertTrue(check_extension("file.txt", ".txt"))

    def test_single_extension_no_match(self):
        self.assertFalse(check_extension("file.txt", ".pdf"))

    def test_multiple_extensions_list(self):
        self.assertTrue(check_extension("image.png", [".jpg", ".png"]))

    def test_multiple_extensions_tuple(self):
        self.assertTrue(check_extension("document.pdf", (".txt", ".pdf")))

    def test_multiple_extensions_set(self):
        self.assertTrue(check_extension("archive.zip", {".rar", ".zip"}))

    def test_case_insensitive_filepath(self):
        self.assertTrue(check_extension("FILE.TXT", ".txt"))

    def test_no_extension(self):
        self.assertFalse(check_extension("README", ".txt"))

    def test_extension_without_dot(self):
        self.assertTrue(check_extension("file.txt", "txt"))
        self.assertFalse(check_extension("file.txt", "pdf"))

    def test_empty_extensions_iterable(self):
        self.assertFalse(check_extension("file.txt", []))

class TestGetFilename(unittest.TestCase):
    def test_unix_path(self):
        self.assertEqual(get_filename("/home/user/file.txt"), "file")

    def test_windows_path(self):
        self.assertEqual(get_filename("C:\\Users\\User\\document.pdf"), "document")

    def test_no_extension(self):
        self.assertEqual(get_filename("/path/to/README"), "README")

    def test_multiple_dots(self):
        self.assertEqual(get_filename("/path/archive.tar.gz"), "archive.tar")

    def test_filename_only(self):
        self.assertEqual(get_filename("image.png"), "image")

    def test_hidden_file(self):
        self.assertEqual(get_filename("/home/user/.env"), "")

    def test_trailing_slash(self):
        self.assertEqual(get_filename("/path/to/folder/"), "")

    def test_empty_string(self):
        self.assertEqual(get_filename(""), "")

class TestReadTxt(unittest.TestCase):
    def test_reads_valid_txt_file(self):
        with tempfile.NamedTemporaryFile(mode="w+", suffix=".txt", delete=False, encoding="utf-8") as tmp:
            tmp.write("line1\nline2\n")
            tmp_path = tmp.name

        try:
            result = read_txt(tmp_path)
            self.assertEqual(result, ["line1\n", "line2\n"])
        finally:
            os.remove(tmp_path)

    def test_empty_txt_file(self):
        with tempfile.NamedTemporaryFile(mode="w+", suffix=".txt", delete=False, encoding="utf-8") as tmp:
            tmp_path = tmp.name

        try:
            result = read_txt(tmp_path)
            self.assertEqual(result, [])
        finally:
            os.remove(tmp_path)

    def test_raises_value_error_for_wrong_extension(self):
        with self.assertRaises(ValueError):
            read_txt("file.pdf")

    def test_case_insensitive_extension(self):
        with tempfile.NamedTemporaryFile(mode="w+", suffix=".TXT", delete=False, encoding="utf-8") as tmp:
            tmp.write("hello\n")
            tmp_path = tmp.name

        try:
            result = read_txt(tmp_path)
            self.assertEqual(result, ["hello\n"])
        finally:
            os.remove(tmp_path)

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

class TestInterpretedToDataFrame(unittest.TestCase):
    def test_basic_conversion(self):
        elements = [
            {
                "key": "epic",
                "content": [
                    {"text": "hello", "comment": "greet", "extra": "[meta]"},
                    {"text": "world", "comment": "", "extra": ""}
                ]
            }
        ]
        df, merges = interpreted_to_dataframe(elements)

        expected_data = [
            ["epic", "[meta]", "greet", "hello"],
            [None, "", "", "world"]
        ]
        expected_df = pd.DataFrame(expected_data, columns=["Key", "Extra", "Comments", "Contents"])
        expected_merges = [(0, 1)]

        pd.testing.assert_frame_equal(df, expected_df)
        self.assertEqual(merges, expected_merges)

    def test_empty_content_creates_row(self):
        elements = [{"key": "empty", "content": []}]
        df, merges = interpreted_to_dataframe(elements)

        expected_df = pd.DataFrame(
            [["empty", "", "", ""]],
            columns=["Key", "Extra", "Comments", "Contents"]
        )
        expected_merges = [(0, 0)]

        pd.testing.assert_frame_equal(df, expected_df)
        self.assertEqual(merges, expected_merges)

    def test_multiple_keys(self):
        elements = [
            {"key": "k1", "content": [{"text": "a", "comment": "", "extra": ""}]},
            {"key": "k2", "content": [{"text": "b", "comment": "c", "extra": "[x]"}]}
        ]
        df, merges = interpreted_to_dataframe(elements)

        expected_data = [
            ["k1", "", "", "a"],
            ["k2", "[x]", "c", "b"]
        ]
        expected_df = pd.DataFrame(expected_data, columns=["Key", "Extra", "Comments", "Contents"])
        expected_merges = [(0, 0), (1, 1)]

        pd.testing.assert_frame_equal(df, expected_df)
        self.assertEqual(merges, expected_merges)

    def test_key_none(self):
        elements = [{"key": None, "content": [{"text": "data", "comment": "info", "extra": ""}]}]
        df, merges = interpreted_to_dataframe(elements)

        expected_df = pd.DataFrame([[None, "", "info", "data"]], columns=["Key", "Extra", "Comments", "Contents"])
        expected_merges = [(0, 0)]

        pd.testing.assert_frame_equal(df, expected_df)
        self.assertEqual(merges, expected_merges)

class TestUnifyKeys(unittest.TestCase):
    def test_all_languages_have_all_keys(self):
        elements_by_language = {
            "English": [{"key": "k1", "content": [{"text": "a", "comment": "", "extra": ""}]}],
            "Spanish": [{"key": "k1", "content": [{"text": "b", "comment": "", "extra": ""}]}]
        }
        unified = unify_keys(elements_by_language)
        self.assertEqual(len(unified["English"]), 1)
        self.assertEqual(len(unified["Spanish"]), 1)

    def test_adds_missing_keys(self):
        elements_by_language = {
            "English": [{"key": "k1", "content": [{"text": "a", "comment": "", "extra": ""}]}],
            "Spanish": [{"key": "k2", "content": [{"text": "b", "comment": "", "extra": ""}]}]
        }
        unified = unify_keys(elements_by_language)

        english_keys = {item["key"] for item in unified["English"]}
        spanish_keys = {item["key"] for item in unified["Spanish"]}

        self.assertEqual(english_keys, {"k1", "k2"})
        self.assertEqual(spanish_keys, {"k1", "k2"})

    def test_existing_none_key_is_not_added(self):
        elements_by_language = {
            "English": [{"key": None, "content": [{"text": "a", "comment": "", "extra": ""}]}],
            "Spanish": [{"key": "k1", "content": [{"text": "b", "comment": "", "extra": ""}]}]
        }
        unified = unify_keys(elements_by_language)

        english_keys = {item["key"] for item in unified["English"] if item["key"] is not None}
        spanish_keys = {item["key"] for item in unified["Spanish"] if item["key"] is not None}

        self.assertEqual(english_keys, {"k1"})
        self.assertEqual(spanish_keys, {"k1"})

    def test_empty_language_elements(self):
        elements_by_language = {
            "English": [],
            "Spanish": [{"key": "k1", "content": [{"text": "b", "comment": "", "extra": ""}]}]
        }
        unified = unify_keys(elements_by_language)
        english_keys = {item["key"] for item in unified["English"] if item["key"] is not None}
        spanish_keys = {item["key"] for item in unified["Spanish"] if item["key"] is not None}

        self.assertEqual(english_keys, {"k1"})
        self.assertEqual(spanish_keys, {"k1"})

class TestExportFolderToXlsxUnified(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.folder_path = self.temp_dir.name

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_txt_file(self, filename: str, content: str):
        path = os.path.join(self.folder_path, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return path

    def test_basic_export_creates_excel_file(self):
        self.create_txt_file("English.txt", "k1=hello\nk2=world")
        self.create_txt_file("Spanish.txt", "k1=hola\nk2=mundo")

        output_file = os.path.join(self.folder_path, "output.xlsx")
        export_folder_to_xlsx_unified(self.folder_path, output_file)

        self.assertTrue(os.path.exists(output_file))

        wb = load_workbook(output_file)
        self.assertIn("English", wb.sheetnames)
        self.assertIn("Spanish", wb.sheetnames)

        ws_en = wb["English"]
        ws_es = wb["Spanish"]

        self.assertEqual(ws_en["A2"].value, "k1")
        self.assertEqual(ws_en["A3"].value, "k2")
        self.assertIsNone(ws_es["B2"].value)

    def test_unified_keys_added(self):
        self.create_txt_file("English.txt", "k1=hello")
        self.create_txt_file("Spanish.txt", "k2=mundo")

        output_file = os.path.join(self.folder_path, "output.xlsx")
        export_folder_to_xlsx_unified(self.folder_path, output_file)

        wb = load_workbook(output_file)
        ws_en = wb["English"]
        ws_es = wb["Spanish"]

        keys_en = [ws_en[f"A{i}"].value for i in range(2, 4)]
        keys_es = [ws_es[f"A{i}"].value for i in range(2, 4)]

        self.assertEqual(set(keys_en), {"k1", "k2"})
        self.assertEqual(set(keys_es), {"k1", "k2"})

    def test_ignores_non_txt_files(self):
        self.create_txt_file("English.txt", "k1=hello")
        self.create_txt_file("ignore.me", "should not read")

        output_file = os.path.join(self.folder_path, "output.xlsx")
        export_folder_to_xlsx_unified(self.folder_path, output_file)

        wb = load_workbook(output_file)
        self.assertIn("English", wb.sheetnames)
        self.assertNotIn("ignore.me", wb.sheetnames)

class TestAddLanguages(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.xlsx_path = os.path.join(self.temp_dir.name, "workbook.xlsx")

        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "English"

        ws.append(["Key", "Extra", "Comments", "Contents"])

        ws.append(["k1", "", "hello comment", "hello"])
        ws.append(["k2", "", "world comment", "world"])
        wb.save(self.xlsx_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_add_single_language(self):
        add_languages(self.xlsx_path, ["Spanish"])

        wb = load_workbook(self.xlsx_path)
        self.assertIn("Spanish", wb.sheetnames)
        ws = wb["Spanish"]

        self.assertEqual(ws["A1"].value, "Key")
        self.assertEqual(ws["D1"].value, "Contents")
        self.assertEqual(ws["I1"].value, "Original Content")

        self.assertEqual(ws["A2"].value, "k1")
        self.assertEqual(ws["I2"].value, "hello")
        self.assertIsNone(ws["G2"].value)
        self.assertEqual(ws["H2"].value, "hello comment")

    def test_add_multiple_languages(self):
        add_languages(self.xlsx_path, ["Spanish", "French"])
        wb = load_workbook(self.xlsx_path)
        self.assertIn("Spanish", wb.sheetnames)
        self.assertIn("French", wb.sheetnames)

    def test_existing_language_skipped(self):
        add_languages(self.xlsx_path, ["English"])
        wb = load_workbook(self.xlsx_path)
        self.assertEqual(wb.sheetnames.count("English"), 1)

    def test_base_language_specified(self):
        add_languages(self.xlsx_path, ["French"], base_language="English")
        wb = load_workbook(self.xlsx_path)
        ws = wb["French"]
        self.assertEqual(ws["A2"].value, "k1")

    def test_invalid_base_language_raises(self):
        with self.assertRaises(ValueError):
            add_languages(self.xlsx_path, ["Spanish"], base_language="Nonexistent")

class TestAutoFormatExcel(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.xlsx_path = os.path.join(self.temp_dir.name, "test.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        ws.append(["Key", "Extra", "Comments", "Contents"])
        
        ws.append(["k1", "", "short", "Some long content that should wrap"])
        ws.append(["k2", "extra info", "another comment", "Line1\nLine2\nLine3"])
        wb.save(self.xlsx_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_auto_format_applies(self):
        auto_format_excel(self.xlsx_path)
        wb = load_workbook(self.xlsx_path)
        ws = wb.active

        self.assertEqual(ws.freeze_panes, "A2")

        for cell in ws[1]:
            self.assertTrue(cell.font.bold)

        for row in ws.iter_rows():
            for cell in row:
                self.assertTrue(cell.alignment.wrap_text)
                self.assertEqual(cell.alignment.vertical, "top")

        for col_cells in ws.columns:
            column_letter = col_cells[0].column_letter
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            self.assertGreaterEqual(ws.column_dimensions[column_letter].width, max_len)

class TestCheckXlsxColumns(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.xlsx_path = os.path.join(self.temp_dir.name, "test.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Key", "Extra", "Comments", "Contents"])
        wb.save(self.xlsx_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_columns_exist(self):
        result = check_xlsx_columns(self.xlsx_path)
        self.assertTrue(result["Sheet1"])

    def test_missing_columns(self):
        wb = load_workbook(self.xlsx_path)
        ws = wb.create_sheet("Sheet2")
        ws.append(["Key", "Extra", "Contents"])
        wb.save(self.xlsx_path)

        result = check_xlsx_columns(self.xlsx_path)
        self.assertFalse(result["Sheet2"])
        self.assertTrue(result["Sheet1"])

    def test_custom_required_columns(self):
        result = check_xlsx_columns(self.xlsx_path, required_columns=["Key", "Contents"])
        self.assertTrue(result["Sheet1"])

        result = check_xlsx_columns(self.xlsx_path, required_columns=["Comments", "Extra", "MissingCol"])
        self.assertFalse(result["Sheet1"])

class TestReadXlsxRaw(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.xlsx_path = os.path.join(self.temp_dir.name, "test.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        ws.append(["Key", "Extra", "Comments", "Contents"])

        ws.append(["k1", "extra1", "comment1", "content1"])
        ws.append(["k2", None, "comment2", "content2"])
        ws.append([None, None, None, None])
        wb.save(self.xlsx_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_read_xlsx_raw(self):
        result = read_xlsx_raw(self.xlsx_path)
        self.assertIn("Sheet1", result)
        rows = result["Sheet1"]
        self.assertEqual(len(rows), 2)

        self.assertEqual(rows[0], {
            "Key": "k1",
            "Extra": "extra1",
            "Comments": "comment1",
            "Contents": "content1"
        })

        self.assertEqual(rows[1], {
            "Key": "k2",
            "Extra": "",
            "Comments": "comment2",
            "Contents": "content2"
        })

class TestExportXlsxToTxt(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.output_folder = self.temp_dir.name

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_basic_export(self):
        xlsx_data = {
            "English": [
                {"Key": "k1", "Extra": "", "Comments": "comment1", "Contents": "hello"},
                {"Key": "k2", "Extra": "", "Comments": "", "Contents": "world"},
                {"Key": "", "Extra": "", "Comments": "", "Contents": ""}
            ]
        }

        export_xlsx_to_txt(xlsx_data, self.output_folder)

        output_file = os.path.join(self.output_folder, "English.txt")
        self.assertTrue(os.path.exists(output_file))

        with open(output_file, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()

        self.assertEqual(lines[0], "k1= hello # comment1")
        self.assertEqual(lines[1], "k2= world")
        self.assertEqual(lines[2], "")

    def test_extra_and_content_raises_error(self):
        xlsx_data = {
            "English": [
                {"Key": "k1", "Extra": "extra1", "Comments": "", "Contents": "content1"}
            ]
        }

        with self.assertRaises(ValueError):
            export_xlsx_to_txt(xlsx_data, self.output_folder)

    def test_only_extra_or_content(self):
        xlsx_data = {
            "English": [
                {"Key": "k1", "Extra": "extra1", "Comments": "note", "Contents": ""},
                {"Key": "k2", "Extra": "", "Comments": "note2", "Contents": "content2"}
            ]
        }

        export_xlsx_to_txt(xlsx_data, self.output_folder)
        output_file = os.path.join(self.output_folder, "English.txt")
        with open(output_file, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()

        self.assertEqual(lines[0], "k1= extra1 # note")
        self.assertEqual(lines[1], "k2= content2 # note2")

    def test_template_sheet_skipped(self):
        xlsx_data = {
            "Template": [
                {"Key": "k1", "Extra": "", "Comments": "", "Contents": "hello"}
            ],
            "English": [
                {"Key": "k2", "Extra": "", "Comments": "", "Contents": "world"}
            ]
        }

        export_xlsx_to_txt(xlsx_data, self.output_folder)
        self.assertFalse(os.path.exists(os.path.join(self.output_folder, "Template.txt")))
        self.assertTrue(os.path.exists(os.path.join(self.output_folder, "English.txt")))

if __name__ == '__main__':
    unittest.main()
