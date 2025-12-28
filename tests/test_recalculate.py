import os
import unittest

os.environ.setdefault("BOT_TOKEN", "test-token")

from main import (  # noqa: E402
    FurnitureItem,
    ParsedRow,
    ParsedSpec,
    _calc_spans_for_section,
    _distribute_items_per_section,
    _distribute_width_evenly,
    _recalculate_corpus,
    _recalculate_furniture,
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
    def test_returns_original_data_when_width_is_same(self):
        spec = ParsedSpec(
            source_filename="test.xlsx",
            width_total_mm=2000,
            depth_mm=600,
            height_mm=2400,
            sections_count=2,
            section_width_mm=1000,
            corpus_rows=[
                ParsedRow(
                    name="Боковина",
                    thickness_mm=16,
                    length_mm=2400,
                    width_mm=600,
                    qty=2,
                    material="ЛДСП",
                )
            ],
            furniture_items=[
                FurnitureItem(
                    name="Петля",
                    qty=8,
                    unit="шт",
                )
            ],
            total_weight_kg=85.5,
        )

        corpus_parts, weight, warnings, furniture_items = _recalculate_corpus(spec, new_width=2000)

        self.assertEqual(1, len(corpus_parts))
        self.assertEqual("Боковина", corpus_parts[0]["name"])
        self.assertEqual(85.5, weight)
        self.assertEqual([], warnings)
        self.assertEqual(1, len(furniture_items))
        self.assertEqual(8, furniture_items[0]["qty"])

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


class RecalculateFurnitureHandlesTest(unittest.TestCase):
    def test_handles_include_scaled_drawers(self):
        spec = ParsedSpec(
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
                    qty=4,
                    material="МДФ",
                ),
                ParsedRow(
                    name="Ящик внутренний",
                    thickness_mm=16,
                    length_mm=500,
                    width_mm=400,
                    qty=4,
                    material="ЛДСП",
                ),
            ],
            furniture_items=[
                FurnitureItem(
                    name="Ручка",
                    qty=8,
                    unit="шт",
                )
            ],
            total_weight_kg=120.0,
        )

        furn_items, warnings, _ = _recalculate_furniture(spec, new_width=3000)

        handles = next(f for f in furn_items if "ручк" in f["name"].lower())
        self.assertEqual(12, handles["qty"])
        self.assertTrue(any("учтены ящики" in w for w in warnings))


if __name__ == "__main__":
    unittest.main()
