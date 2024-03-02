import unittest

from xlsql import cli


class TestCli(unittest.TestCase):
    def test_normalize(self):
        def test(input, expected):
            self.assertEqual(expected, cli.normalize(input))

        test("Test Name", "test_name")
        test("  no leading spaces", "no_leading_spaces")
        test("no trailing spaces  ", "no_trailing_spaces")
        test("Filters punctuation!", "filters_punctuation")
        test("Replaces-hyphens", "replaces_hyphens")
        test("ID (SECRET)", "id_secret")

    def test_get_column_names(self):
        def test(sheet_name, headings, expected):
            self.assertEqual(
                expected, cli.get_column_names(sheet_name, headings, print)
            )

        test("Sheet1", ["Name", "ID", "Address"], ["name", "id", "address"])
        test("Sheet2", ["Name", "ID", "Address"], ["name", "id", "address"])
        test("Sheet1", ["", ""], ["EMPTY", "EMPTY_2"])


if __name__ == "__main__":
    unittest.main()
