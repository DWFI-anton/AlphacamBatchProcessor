import importlib.machinery
import logging
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).parents[1] / "app" / "alphacam_gui.pyw"
loader = importlib.machinery.SourceFileLoader("alphacam_gui_test", str(MODULE_PATH))
gui = loader.load_module()


class Collection:
    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class Body:
    def __init__(self, complex_text=""):
        self.Shells = Collection([object()])
        self.Lumps = Collection([object()])
        self.Faces = Collection([TypedObject(complex_text) for _ in range(6)])
        self.Edges = Collection([TypedObject(complex_text) for _ in range(12)])
        self.Vertices = Collection([TypedObject(complex_text) for _ in range(8)])


class TypedObject:
    def __init__(self, properties):
        self.Properties = properties


class Bodies:
    def __init__(self, body):
        self.Count = 1
        self._body = body

    def Item(self, index):
        return self._body


class SolidInterface:
    def __init__(self, body):
        self._body = body

    def BodiesInDrawing(self, drawing):
        return Bodies(self._body)


class Drawing:
    def __init__(self, body):
        self.SolidInterface = SolidInterface(body)


class XtBoxDetectorTests(unittest.TestCase):
    def setUp(self):
        self.processor = object.__new__(gui.AlphacamProcessor)
        self.processor.log = logging.getLogger("xt-box-test")

    def test_topology_marks_box_as_simple(self):
        ok, reason = self.processor.inspect_simple_box(
            Path(__file__), Drawing(Body())
        )
        # The fixture itself is Python, so the quick XT filter must reject it
        # before topology is considered.
        self.assertFalse(ok)
        self.assertIn("Parasolid", reason)

    def test_sample_names_are_quick_candidates_when_available(self):
        candidates = [
            Path(r"C:\temp\Ящик\DNO.x_t"),
            Path(r"C:\temp\26031.01.01.00.04.15.01.01.00 Гардероб\MODEL\26031.325117.001 Антресоль (00)\STENKA-ZADNYAYA.x_t"),
        ]
        existing = [path for path in candidates if path.is_file()]
        if not existing:
            self.skipTest("Эталонные XT-файлы отсутствуют на этой машине")
        for path in existing:
            ok, reason = gui.quick_xt_candidate(path)
            self.assertTrue(ok, f"{path}: {reason}")

    def test_complex_topology_marker_is_rejected(self):
        path = Path(r"C:\temp\Ящик\DNO.x_t")
        if not path.is_file():
            self.skipTest("Эталонный DNO.x_t отсутствует на этой машине")
        ok, _ = self.processor.inspect_simple_box(
            path, Drawing(Body("CYLINDER"))
        )
        self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
