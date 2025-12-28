import os
import unittest

os.environ.setdefault("BOT_TOKEN", "test-token")

from main import ParsedRow, ParsedSpec, _distribute_items_per_section, _recalculate_corpus  # noqa: E402


class RecalculateShelvesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.spec = ParsedSpec(
            source_filename="test.xlsx",
            width_total_mm=2000,
            depth_mm=600,
            height_mm=2400,
            sections_count=2,
            section_width_mm=1000,
            corpus_rows=[
                ParsedRow(
                    name="Полка",
                    thickness_mm=16,
                    length_mm=1000,
                    width_mm=400,
                    qty=5,
                    material="ЛДСП",
                )
            ],
            furniture_items=[],
            total_weight_kg=120.0,
        )

    def test_preserves_shelves_ratio_on_resize(self):
        """Полки пересчитываются по секциям, а не по среднему числу пролётов."""
        new_width = 3000  # станет 3 секции
        corpus_parts, _, _, _ = _recalculate_corpus(self.spec, new_width)

        shelf_part = next(p for p in corpus_parts if "полк" in p["name"].lower())
        self.assertEqual(8, shelf_part["qty"])

    def test_distribute_items_remainder(self):
        """Распределение оставшихся элементов идёт равномерно по секциям."""
        distribution = _distribute_items_per_section(total_qty=5, sections_count=3)
        self.assertEqual([2.0, 2.0, 1.0], distribution)


if __name__ == "__main__":
    unittest.main()
