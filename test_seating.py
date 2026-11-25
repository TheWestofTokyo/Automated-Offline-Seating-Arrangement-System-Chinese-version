import collections
import random
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
import seating_core as core


def people(n, cls='A', role=''):
    return [{'name': f'{cls}-{i}', 'class_name': cls, 'role': role} for i in range(n)]


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / 'input.xlsx'

    def tearDown(self):
        self.temp.cleanup()

    def parse(self, rows):
        wb = Workbook()
        for row in rows:
            wb.active.append(row)
        wb.save(self.path)
        wb.close()
        return core.parse_personnel(self.path)

    def test_parent_name_does_not_override_student(self):
        p, w = self.parse([['姓名', '班级', '家长姓名'], ['甲', '一班', '乙']])
        self.assertEqual(p[0]['name'], '甲')
        self.assertFalse(w)

    def test_ambiguous_headers_rejected(self):
        with self.assertRaisesRegex(ValueError, '多个'):
            self.parse([['姓名', 'Name', '班级'], ['甲', '乙', '一班']])

    def test_physical_excel_row_numbers(self):
        p, w = self.parse([['姓名', '班级'], [None, None], ['甲', None], ['乙', '一班']])
        self.assertIn('第3行', w[0])
        self.assertEqual(p[0]['source_row'], 4)

    def test_duplicates_retained_and_flagged(self):
        p, w = self.parse([['姓名', '班级'], ['甲', '一班'], ['甲', '一班']])
        self.assertEqual(len(p), 2)
        self.assertIn('完全重复', w[0])

    def test_unknown_role_retained_and_flagged(self):
        p, w = self.parse([['姓名', '班级', '身份'], ['甲', '一班', '嘉宾']])
        self.assertEqual((p[0]['role'], p[0]['source_role']), ('', '嘉宾'))
        self.assertIn('普通人员', w[0])

    def test_teacher_without_class(self):
        p, w = self.parse([['姓名', '班级', '身份'], ['甲', None, '教师']])
        self.assertEqual((p[0]['role'], p[0]['class_name']), ('老师', '老师'))
        self.assertFalse(w)

    def test_ordinary_student_role_needs_no_warning(self):
        p, w = self.parse([['姓名', '班级', '身份'], ['甲', '一班', '普通学生']])
        self.assertFalse(w)
        self.assertEqual(p[0]['source_role'], '普通学生')

    def test_uncached_formula_is_explained(self):
        p, w = self.parse([['姓名', '班级'], ['="甲"', '一班'], ['乙', '一班']])
        self.assertEqual(len(p), 1)
        self.assertIn('公式', w[0])

    def test_sheet_can_be_selected(self):
        wb = Workbook()
        wb.active.title = '说明'
        wb.active.append(['说明文字'])
        ws = wb.create_sheet('名单')
        ws.append(['姓名', '单位'])
        ws.append(['甲', '一班'])
        wb.save(self.path)
        wb.close()
        self.assertEqual(core.workbook_sheets(self.path), ['说明', '名单'])
        p, _ = core.parse_personnel(self.path, '名单')
        self.assertEqual(p[0]['name'], '甲')


class AssignmentTests(unittest.TestCase):
    def test_capacity_and_reserved_seats(self):
        p = people(498) + people(9, '主席', '主席团') + people(11, '教师', '老师')
        a, s, _ = core.execute_assignment(p)
        self.assertEqual(len(a), 518)
        self.assertEqual(collections.Counter(s.values())['prohibited'], 4)
        self.assertEqual(collections.Counter(s.values())['桌子'], 10)

    def test_overcapacity(self):
        for n, role in [(499, ''), (10, '主席团'), (12, '老师')]:
            with self.subTest(role=role), self.assertRaises(ValueError):
                core.execute_assignment(people(n, role=role))

    def test_front_priority_does_not_jump_to_row_nine(self):
        a, _, _ = core.execute_assignment(people(18))
        self.assertEqual({k[0] for k in a}, {1})

    def test_compact_minimizes_fragments_for_nineteen(self):
        a, _, _ = core.execute_assignment(people(19), 'compact')
        self.assertEqual(len(core.contiguous_segments(a)), 2)

    def test_banned_seat_splits_segment(self):
        segments = core.contiguous_segments([(6, 'C', i) for i in range(14, 29) if i != 21])
        self.assertEqual(len(segments), 2)

    def test_ceremonial_input_order_preserved(self):
        p = people(1, 'Z', '主席团') + people(1, 'A', '主席团')
        a, _, _ = core.execute_assignment(p)
        self.assertEqual(a[(1, 'C', 17)]['class_name'], 'Z')

    def test_random_partitions_both_modes(self):
        rng = random.Random(29)
        for mode in ('front', 'compact'):
            for _ in range(100):
                n = rng.randint(0, 498)
                p, left, i = [], n, 0
                while left:
                    count = rng.randint(1, min(left, 90))
                    p += people(count, str(i))
                    left -= count
                    i += 1
                a, _, _ = core.execute_assignment(p, mode)
                self.assertEqual(len(a), n)
                core.validate_assignments(p, a)


class ExportTests(unittest.TestCase):
    def test_all_three_exports_precise_ranges_and_literal_names(self):
        p = people(5) + people(2, 'B')
        p[0]['name'] = '=1+1'
        a, s, _ = core.execute_assignment(p)
        with tempfile.TemporaryDirectory() as tmp:
            result = Path(core.export_results(a, s, tmp))
            self.assertEqual(len(list(result.glob('*.xlsx'))), 3)
            wb = load_workbook(result / '03_班级座位分配表.xlsx')
            rows = list(wb.active.values)
            self.assertIn((1, 'A', '1—5', 'A', 5, '普通'), rows)
            self.assertIn((1, 'A', '6—7', 'B', 2, '普通'), rows)
            wb.close()
            wb = load_workbook(result / '02_个人座位明细.xlsx')
            self.assertEqual(wb.active.max_row, 533)
            formula_name = next(c for row in wb.active for c in row if c.value == '=1+1')
            self.assertEqual(formula_name.data_type, 's')
            wb.close()
            wb = load_workbook(result / '01_可视化座位图.xlsx')
            self.assertIn('1号', wb.active['B5'].value)
            wb.close()
            other = Path(core.export_results(a, s, tmp))
            self.assertNotEqual(result, other)

    def test_failure_leaves_no_partial_result(self):
        a, s, _ = core.execute_assignment(people(1))
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(core, 'generate_detail_list', side_effect=OSError('file locked')):
                with self.assertRaises(OSError):
                    core.export_results(a, s, tmp)
            self.assertEqual(list(Path(tmp).iterdir()), [])


if __name__ == '__main__':
    unittest.main(verbosity=2)
