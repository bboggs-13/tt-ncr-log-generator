import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INDEX = ROOT / "components" / "photo_grid" / "index.html"


class PhotoGridComponentTests(unittest.TestCase):
    def test_version_matches_app(self):
        app = (ROOT / "app.py").read_text(encoding="utf-8")
        wanted = re.search(r'_PHOTO_GRID_VERSION = "([^"]+)"', app).group(1)
        self.assertIn(f"PHOTO_GRID_VERSION {wanted}", INDEX.read_text(encoding="utf-8"))

    def test_uses_the_streamlit_component_messages(self):
        html = INDEX.read_text(encoding="utf-8")
        for message in ("streamlit:componentReady", "streamlit:render",
                        "streamlit:setComponentValue", "streamlit:setFrameHeight"):
            self.assertIn(message, html)

    @unittest.skipUnless(shutil.which("node"), "node not installed")
    def test_script_and_move_logic(self):
        js = re.search(r"<script>(.*)</script>", INDEX.read_text(encoding="utf-8"), re.S).group(1)
        with tempfile.TemporaryDirectory() as tmp:
            grid = Path(tmp) / "grid.js"
            grid.write_text(js, encoding="utf-8")
            subprocess.run(["node", "--check", str(grid)], check=True)
            test = Path(tmp) / "t.js"
            test.write_text(
                f'const {{ moveItem }} = require({str(grid)!r});\n'
                'const a = require("assert");\n'
                'a.deepStrictEqual(moveItem([1,2,3,4,5,6], 5, 0), [6,1,2,3,4,5]);\n'
                'a.deepStrictEqual(moveItem([1,2,3,4,5,6], 0, 3), [2,3,4,1,5,6]);\n'
                'a.deepStrictEqual(moveItem([1,2,3], 2, 2), [1,2,3]);\n'
                'a.deepStrictEqual(moveItem([1,2,3], 0, 7), [1,2,3]);\n', encoding="utf-8")
            subprocess.run(["node", str(test)], check=True)


if __name__ == "__main__":
    unittest.main()
