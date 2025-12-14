import os
import io
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout

# Import the function under test
from convert_invoice import remplacer_symbole_direct


class TestRemplacerSymboleDirect(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def _write_file(self, path, content, encoding="utf-8"):
        with open(path, "w", encoding=encoding) as f:
            f.write(content)

    def _read_file(self, path, encoding="utf-8"):
        with open(path, "r", encoding=encoding) as f:
            return f.read()

    def test_replaces_all_occurrences(self):
        # Arrange
        src = os.path.join(self.tmpdir, "in.txt")
        dst = os.path.join(self.tmpdir, "out.txt")
        self._write_file(src, "a%20b%20c%20d")

        # Act
        f = io.StringIO()
        with redirect_stdout(f):
            remplacer_symbole_direct(src, dst, "%20", " ")
        output = f.getvalue()

        # Assert
        self.assertIn("Le remplacement a été effectué", output)
        self.assertTrue(os.path.exists(dst))
        self.assertEqual(self._read_file(dst), "a b c d")

    def test_no_occurrences_keeps_content(self):
        # Arrange
        src = os.path.join(self.tmpdir, "in.txt")
        dst = os.path.join(self.tmpdir, "out.txt")
        original = "no-change-here"
        self._write_file(src, original)

        # Act
        f = io.StringIO()
        with redirect_stdout(f):
            remplacer_symbole_direct(src, dst, "%20", "-")

        # Assert
        self.assertTrue(os.path.exists(dst))
        self.assertEqual(self._read_file(dst), original)

    def test_empty_file_results_empty(self):
        # Arrange
        src = os.path.join(self.tmpdir, "in.txt")
        dst = os.path.join(self.tmpdir, "out.txt")
        self._write_file(src, "")

        # Act
        remplacer_symbole_direct(src, dst, "%20", "-")

        # Assert
        self.assertTrue(os.path.exists(dst))
        self.assertEqual(self._read_file(dst), "")

    def test_overwrites_existing_output_file(self):
        # Arrange
        src = os.path.join(self.tmpdir, "in.txt")
        dst = os.path.join(self.tmpdir, "out.txt")
        self._write_file(src, "x%20y%20z")
        self._write_file(dst, "PREVIOUS_CONTENT")  # should be overwritten

        # Act
        remplacer_symbole_direct(src, dst, "%20", "-")

        # Assert
        self.assertEqual(self._read_file(dst), "x-y-z")

    def test_handles_non_ascii_characters(self):
        # Arrange
        src = os.path.join(self.tmpdir, "in.txt")
        dst = os.path.join(self.tmpdir, "out.txt")
        # includes accented characters and non-Latin script
        content = "café%20naïve%20東京"
        self._write_file(src, content)

        # Act
        remplacer_symbole_direct(src, dst, "%20", " ")

        # Assert
        self.assertEqual(self._read_file(dst), "café naïve 東京")

    def test_missing_input_file_prints_error_and_no_output(self):
        # Arrange
        src = os.path.join(self.tmpdir, "does_not_exist.txt")
        dst = os.path.join(self.tmpdir, "out.txt")

        # Act
        f = io.StringIO()
        with redirect_stdout(f):
            remplacer_symbole_direct(src, dst, "%20", " ")
        output = f.getvalue()

        # Assert
        self.assertIn("Une erreur s'est produite", output)
        self.assertFalse(os.path.exists(dst))


if __name__ == "__main__":
    unittest.main()
