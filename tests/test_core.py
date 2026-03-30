"""Basic regression tests for stable helper functions."""

from pathlib import Path, PurePosixPath
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from compresspptx.core import apply_profile, part_from_rel_path, relative_target, resize_fit


class ResizeFitTests(unittest.TestCase):
    def test_keeps_size_if_below_limit(self) -> None:
        self.assertEqual(resize_fit((800, 600), 1200), (800, 600))

    def test_scales_long_edge(self) -> None:
        self.assertEqual(resize_fit((4000, 2000), 1000), (1000, 500))


class PathHelperTests(unittest.TestCase):
    def test_relative_target_uses_posix_separators(self) -> None:
        source = PurePosixPath("ppt/slides/slide1.xml")
        target = PurePosixPath("ppt/media/image1.jpg")
        self.assertEqual(relative_target(source, target), "../media/image1.jpg")

    def test_part_from_rel_path_for_slide_relationship(self) -> None:
        root = Path("workspace")
        rel_path = root / "ppt" / "slides" / "_rels" / "slide1.xml.rels"
        self.assertEqual(part_from_rel_path(rel_path, root), PurePosixPath("ppt/slides/slide1.xml"))

    def test_part_from_rel_path_for_root_relationship(self) -> None:
        root = Path("workspace")
        rel_path = root / "_rels" / ".rels"
        self.assertIsNone(part_from_rel_path(rel_path, root))


class ProfileTests(unittest.TestCase):
    def test_apply_profile_returns_expected_limits(self) -> None:
        self.assertEqual(apply_profile("balanced"), (1600, 900, 960))

    def test_apply_profile_rejects_unknown_value(self) -> None:
        with self.assertRaises(ValueError):
            apply_profile("unknown")


if __name__ == "__main__":
    unittest.main()
