import os
import unittest

os.environ.setdefault("BOT_TOKEN", "test-token")

from main import (  # noqa: E402
    ParsedRow,
    ParsedSpec,
    _calc_spans_for_section,
    _distribute_items_per_section,
    _distribute_width_evenly,
    _recalculate_corpus,
    _split_sections,
)


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


class RecalculateFacadeTest(unittest.TestCase):
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
                    name="Фасад",
                    thickness_mm=18,
                    length_mm=2300,
                    width_mm=500,
                    qty=8,
                    material="МДФ",
                )
            ],
            furniture_items=[],
            total_weight_kg=120.0,
        )

    def test_facade_qty_uses_spans_ratio(self):
        new_width = 2500  # -> 3 секции, 6 пролётов
        corpus_parts, _, _, _ = _recalculate_corpus(self.spec, new_width)

        facade_part = next(p for p in corpus_parts if "фасад" in p["name"].lower())
        self.assertEqual(12, facade_part["qty"])

    def test_facade_width_description_keeps_proportion(self):
        new_width = 2500
        corpus_parts, _, _, _ = _recalculate_corpus(self.spec, new_width)

        sections = _split_sections(new_width)
        span_widths = []
        for width in sections:
            spans = _calc_spans_for_section(width)
            span_widths.extend(_distribute_width_evenly(width, spans))

        expected_facade_widths = []
        for span_width in span_widths:
            expected_facade_widths.extend(_distribute_width_evenly(span_width, 2))

        facade_part = next(p for p in corpus_parts if "фасад" in p["name"].lower())
        self.assertEqual(expected_facade_widths, facade_part["widths_mm"])


if __name__ == "__main__":
    unittest.main()
