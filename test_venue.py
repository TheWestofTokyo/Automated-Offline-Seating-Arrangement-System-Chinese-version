import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
import seating_core as core
from test_seating import people


class VenueTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / '新礼堂.xlsx'

    def tearDown(self):
        self.temp.cleanup()

    def write(self, rows):
        wb = Workbook()
        for row in rows:
            wb.active.append(row)
        wb.save(self.path)
        wb.close()

    def test_builtin_name_and_capacity(self):
        self.assertEqual(core.DEFAULT_VENUE.name, '大学生活动中心礼堂')
        self.assertEqual(len(core.DEFAULT_VENUE.seats), 532)
        self.assertEqual(core.DEFAULT_VENUE.counts['available'], 498)

    def test_single_seat_import_and_nonconsecutive_numbering(self):
        self.write([['场地名称', '排号', '区域', '座位号', '用途'],
                    ['新礼堂', 2, '中', 3, '普通'], ['新礼堂', 2, '中', 5, '禁排'],
                    ['新礼堂', 4, '中', 9, '普通']])
        venue = core.import_venue(self.path)
        self.assertEqual(venue.name, '新礼堂')
        self.assertEqual(venue.rows, [2, 4])
        self.assertEqual(venue.sections, ('中',))
        a, s, _ = core.execute_assignment(people(2), venue=venue)
        self.assertEqual(set(a), {(2, '中', 3), (4, '中', 9)})
        self.assertEqual(s[(2, '中', 5)], 'prohibited')
        self.assertEqual(len(s), 3)

    def test_range_import_and_custom_role_locations(self):
        self.write([['排号', '区域', '起始座位号', '结束座位号', '用途'],
                    [7, '贵宾', 101, 102, '主席团'], [9, '东侧', 4, 5, '老师'],
                    [8, '东侧', 10, 16, '普通'], [8, '东侧', 17, 17, '桌子']])
        venue = core.import_venue(self.path)
        self.assertEqual(venue.name, '新礼堂')
        p = people(7) + people(2, 'P', '主席团') + people(2, 'T', '老师')
        a, s, _ = core.execute_assignment(p, venue=venue)
        core.validate_assignments(p, a, venue)
        self.assertEqual(a[(7, '贵宾', 101)]['role'], '主席团')
        self.assertEqual(a[(9, '东侧', 4)]['role'], '老师')
        with self.assertRaises(ValueError):
            core.execute_assignment(p + people(1, 'extra'), venue=venue)

    def test_duplicate_or_overlapping_ranges_are_rejected(self):
        self.write([['排号', '区域', '起始座位号', '结束座位号'], [1, 'A', 1, 5], [1, 'A', 5, 8]])
        with self.assertRaisesRegex(ValueError, '重复或重叠'):
            core.import_venue(self.path)

    def test_invalid_fields_rejected(self):
        for data in ([0, 'A', 1, '普通'], [1.5, 'A', 1, '普通'], [1, '', 1, '普通'],
                     [1, 'A', -1, '普通'], [1, 'A', 1, '不明用途'], ['=1+1', 'A', 1, '普通']):
            with self.subTest(data=data):
                self.write([['排号', '区域', '座位号', '用途'], data])
                with self.assertRaises(ValueError):
                    core.import_venue(self.path)

    def test_reversed_ranges_rejected(self):
        self.write([['排号', '区域', '起始座位号', '结束座位号'], [1, 'A', 8, 2]])
        with self.assertRaisesRegex(ValueError, '结束座位号'):
            core.import_venue(self.path)

    def test_inconsistent_names_rejected(self):
        self.write([['场地名称', '排号', '区域', '座位号'], ['甲', 1, 'A', 1], ['乙', 1, 'A', 2]])
        with self.assertRaisesRegex(ValueError, '一个场地'):
            core.import_venue(self.path)

    def test_both_number_formats_are_rejected(self):
        self.write([['排号', '区域', '座位号', '起始座位号', '结束座位号'], [1, 'A', 1, 1, 2]])
        with self.assertRaisesRegex(ValueError, '一种格式'):
            core.import_venue(self.path)

    def test_template_roundtrip_preserves_builtin_exactly(self):
        core.generate_seat_template(self.path)
        venue = core.import_venue(self.path)
        self.assertEqual(venue.name, core.DEFAULT_VENUE.name)
        self.assertEqual(venue.seats, core.DEFAULT_VENUE.seats)

    def test_custom_exports_have_only_new_geometry_and_name(self):
        self.write([['场地名称', '排号', '区域', '座位号'], ['新礼堂', 20, '西', 13], ['新礼堂', 20, '西', 15]])
        venue = core.import_venue(self.path)
        a, s, _ = core.execute_assignment(people(2), venue=venue)
        result = Path(core.export_results(a, s, self.temp.name, venue))
        wb = load_workbook(result / '01_可视化座位图.xlsx')
        self.assertIn('新礼堂', wb.active['B1'].value)
        self.assertEqual(wb.active['A5'].value, '第20排')
        wb.close()
        wb = load_workbook(result / '02_个人座位明细.xlsx')
        self.assertEqual(wb.active.max_row, 3)
        self.assertEqual(list(wb.active.values)[1][:3], (20, '西', 13))
        wb.close()
        wb = load_workbook(result / '03_班级座位分配表.xlsx')
        self.assertIn('新礼堂', wb.active['A1'].value)
        self.assertIn((20, '西', '13、15', 'A', 2, '普通'), list(wb.active.values))
        wb.close()

    def test_library_roundtrip_and_bad_file_isolation(self):
        directory = Path(self.temp.name) / '场地库'
        core.save_venue(core.DEFAULT_VENUE, directory)
        (directory / 'bad.json').write_text('{invalid', encoding='utf-8')
        venues, warnings = core.load_venues(directory)
        self.assertEqual(len(venues), 1)
        self.assertEqual(venues[0].seats, core.DEFAULT_VENUE.seats)
        self.assertEqual(len(warnings), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
