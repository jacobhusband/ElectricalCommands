import importlib.util
import tempfile
import unittest
from pathlib import Path

import fitz


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / "scripts" / "strip_pdf_layers.py"

spec = importlib.util.spec_from_file_location("strip_pdf_layers", SCRIPT_PATH)
strip_pdf_layers = importlib.util.module_from_spec(spec)
spec.loader.exec_module(strip_pdf_layers)


def _ink_count(pdf_path):
    with fitz.open(pdf_path) as doc:
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
    samples = pix.samples
    return sum(
        1
        for index in range(0, len(samples), 3)
        if samples[index:index + 3] != b"\xff\xff\xff"
    )


def _ocg_count(pdf_path):
    with fitz.open(pdf_path) as doc:
        return len(doc.get_ocgs() or {})


def _catalog_ocproperties_type(pdf_path):
    with fitz.open(pdf_path) as doc:
        value_type, _ = doc.xref_get_key(doc.pdf_catalog(), "OCProperties")
        return value_type


class StripPdfLayersTests(unittest.TestCase):
    def test_layered_pdf_becomes_layer_free_without_deleting_content(self):
        with tempfile.TemporaryDirectory(prefix="acies-strip-layers-") as temp_dir:
            pdf_path = Path(temp_dir) / "layered.pdf"
            doc = fitz.open()
            page = doc.new_page(width=320, height=220)
            visible_layer = doc.add_ocg("visible layer", on=True)
            hidden_layer = doc.add_ocg("hidden layer", on=False)
            page.insert_text(
                (35, 70),
                "VISIBLE CONTENT",
                fontsize=28,
                fill=(0, 0, 0),
                oc=visible_layer,
            )
            page.insert_text(
                (35, 140),
                "HIDDEN CONTENT",
                fontsize=28,
                fill=(0, 0, 0),
                oc=hidden_layer,
            )
            doc.save(pdf_path)
            doc.close()

            before_ink = _ink_count(pdf_path)
            self.assertEqual(2, _ocg_count(pdf_path))
            self.assertGreater(before_ink, 0)

            removed = strip_pdf_layers.strip_pdf_layers(str(pdf_path))

            self.assertEqual(2, removed)
            self.assertEqual(0, _ocg_count(pdf_path))
            self.assertEqual("null", _catalog_ocproperties_type(pdf_path))
            self.assertGreater(
                _ink_count(pdf_path),
                before_ink,
                "Hidden layer content should become visible, not be deleted.",
            )

    def test_non_layered_pdf_is_successful_noop(self):
        with tempfile.TemporaryDirectory(prefix="acies-strip-layers-") as temp_dir:
            pdf_path = Path(temp_dir) / "plain.pdf"
            doc = fitz.open()
            page = doc.new_page(width=200, height=120)
            page.insert_text((30, 60), "PLAIN PDF", fontsize=20)
            doc.save(pdf_path)
            doc.close()

            before_ink = _ink_count(pdf_path)
            before_bytes = pdf_path.read_bytes()

            removed = strip_pdf_layers.strip_pdf_layers(str(pdf_path))

            self.assertEqual(0, removed)
            self.assertEqual(0, _ocg_count(pdf_path))
            self.assertEqual("null", _catalog_ocproperties_type(pdf_path))
            self.assertEqual(before_ink, _ink_count(pdf_path))
            self.assertEqual(before_bytes, pdf_path.read_bytes())


if __name__ == "__main__":
    unittest.main()
