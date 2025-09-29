import unittest

from section3_parser import extract_section3_rows


class Section3ParserTests(unittest.TestCase):
    def test_extract_section3_rows_with_prefixed_headers(self):
        source = [
            {
                "AA§Условный номер помещения": "НП-1",
                "AB§Назначение помещения": "Магазин",
                "AC§Этаж расположения": "1",
                "AD§Номер подъезда": "2",
                "AE§Площадь, м2": "35,7",
                "AF§Наименование помещения": "Помещение 101",
                "AG§Площадь части, м2": "35,7",
                "AH§Высота потолков, м": "3,45",
            }
        ]

        rows = extract_section3_rows(source)

        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["unit_number"], "НП-1")
        self.assertEqual(row["purpose"], "Магазин")
        self.assertEqual(row["floor"], "1")
        self.assertEqual(row["entrance_number"], "2")
        self.assertAlmostEqual(row["area"], 35.7)
        self.assertEqual(row["room_name"], "Помещение 101")
        self.assertAlmostEqual(row["part_area"], 35.7)
        self.assertAlmostEqual(row["ceiling_height"], 3.45)

    def test_skip_rows_without_meaningful_values(self):
        source = [
            {
                "AA§Условный номер помещения": "",
                "AB§Назначение помещения": "",
                "AC§Этаж расположения": "",
                "AE§Площадь, м2": "",
            },
            {
                "AA§Условный номер помещения": "НП-2",
                "AB§Назначение помещения": "Офис",
                "AC§Этаж расположения": "2",
                "AE§Площадь, м2": "18,25 м²",
                "AG§Площадь части, м2": "",
            },
        ]

        rows = extract_section3_rows(source)

        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["unit_number"], "НП-2")
        self.assertEqual(row["purpose"], "Офис")
        self.assertEqual(row["floor"], "2")
        self.assertAlmostEqual(row["area"], 18.25)
        # Площадь части должна подставляться из общей площади, если она отсутствует.
        self.assertAlmostEqual(row["part_area"], 18.25)


if __name__ == "__main__":
    unittest.main()

