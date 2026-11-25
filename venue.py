"""Explicit venue geometry and structured Excel import."""
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import json
import os
from uuid import uuid4

from openpyxl import load_workbook

STATUS_LABELS = {'available': '普通', '主席团': '主席团', '老师': '老师',
                 '桌子': '桌子', 'prohibited': '禁排'}
STATUS_ALIASES = {**{v: k for k, v in STATUS_LABELS.items()},
                  '可用': 'available', '普通座位': 'available', '普通人员': 'available',
                  '教师': '老师', '桌子预留': '桌子', '禁排空位': 'prohibited',
                  '主席团预留': '主席团', '老师预留': '老师'}


@dataclass
class Venue:
    name: str
    seats: dict
    sections: tuple

    def __post_init__(self):
        if not isinstance(self.name, str) or not self.name.strip() or not isinstance(self.seats, dict) or not self.seats:
            raise ValueError('场地名称和座位表不能为空。')
        if (not isinstance(self.sections, tuple) or not all(isinstance(s, str) for s in self.sections)
                or len(self.sections) != len(set(self.sections)) or set(self.sections) != {k[1] for k in self.seats}):
            raise ValueError('座位区域信息不一致。')
        for (row, section, number), state in self.seats.items():
            if (isinstance(row, bool) or not isinstance(row, int) or row < 1
                    or isinstance(number, bool) or not isinstance(number, int) or number < 1
                    or not isinstance(section, str) or not section.strip() or state not in STATUS_LABELS):
                raise ValueError('座位排号、区域、座位号或用途不合法。')

    @property
    def rows(self):
        return sorted({k[0] for k in self.seats})

    @property
    def counts(self):
        return Counter(self.seats.values())

    def row_seats(self, row, section):
        return sorted(k for k in self.seats if k[:2] == (row, section))

    @property
    def widths(self):
        return {s: max(len(self.row_seats(r, s)) for r in self.rows) for s in self.sections}

    def sort_key(self, key):
        return key[0], self.sections.index(key[1]), key[2]


def _positive_integer(value, row, field):
    try:
        number = float(value)
        if isinstance(value, bool) or not number.is_integer() or number < 1:
            raise ValueError()
        return int(number)
    except (TypeError, ValueError, OverflowError):
        raise ValueError(f'第{row}行：{field}必须是正整数。') from None


def import_venue(filepath, sheet_name=None):
    wb = load_workbook(filepath, read_only=True, data_only=False)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active
        rows = [(i, row) for i, row in enumerate(ws.iter_rows(), 1)
                if any(c.value is not None and str(c.value).strip() for c in row)]
        if len(rows) < 2:
            raise ValueError('座位表需要表头和至少一个座位。')
        columns = {}
        aliases = {'排号': 'row', '排': 'row', '区域': 'section', '区块': 'section',
                   '座位号': 'number', '起始座位号': 'start', '结束座位号': 'end',
                   '用途': 'status', '状态': 'status', '场地名称': 'name'}
        for i, cell in enumerate(rows[0][1]):
            field = aliases.get(str(cell.value or '').replace(' ', '').strip())
            if field:
                if field in columns:
                    raise ValueError('座位表有重复字段，请保留一列：' + str(cell.value))
                columns[field] = i
        if not {'row', 'section'} <= columns.keys():
            raise ValueError('座位表必须包含“排号”和“区域”列。')
        single = 'number' in columns
        if single and ('start' in columns or 'end' in columns):
            raise ValueError('“座位号”与“起始/结束座位号”请只使用一种格式。')
        if not single and not {'start', 'end'} <= columns.keys():
            raise ValueError('请提供“座位号”，或同时提供“起始座位号”和“结束座位号”。')
        seats, sections, names, origins = {}, [], set(), {}
        for row_num, cells in rows[1:]:
            def value(field):
                col = columns.get(field)
                if col is None or col >= len(cells):
                    return None
                cell = cells[col]
                if cell.data_type == 'f':
                    raise ValueError(f'第{row_num}行：座位表请使用明确的数值或文本，不使用公式。')
                return cell.value
            row = _positive_integer(value('row'), row_num, '排号')
            section_value = value('section')
            section = '' if section_value is None else str(section_value).strip()
            if not section:
                raise ValueError(f'第{row_num}行：区域不能为空。')
            if section not in sections:
                sections.append(section)
            start = _positive_integer(value('number' if single else 'start'), row_num, '座位号')
            end = start if single else _positive_integer(value('end'), row_num, '结束座位号')
            if end < start:
                raise ValueError(f'第{row_num}行：结束座位号不能小于起始座位号。')
            label = str(value('status') or '普通').strip()
            if label not in STATUS_ALIASES:
                raise ValueError(f'第{row_num}行：不认识用途“{label}”，支持普通、主席团、老师、桌子、禁排。')
            state = STATUS_ALIASES[label]
            name = str(value('name') or '').strip()
            if name:
                names.add(name)
            for number in range(start, end + 1):
                key = (row, section, number)
                if key in seats:
                    raise ValueError(f'第{row_num}行与第{origins[key]}行重复或重叠：{row}排{section}区{number}号。')
                seats[key] = state
                origins[key] = row_num
        if len(names) > 1:
            raise ValueError('一个工作表只能描述一个场地，请统一“场地名称”。')
        name = next(iter(names), Path(filepath).stem)
        venue = Venue(name, seats, tuple(sections))
        venue.seats = dict(sorted(seats.items(), key=lambda item: venue.sort_key(item[0])))
        return venue
    finally:
        wb.close()


def save_venue(venue, directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / (uuid4().hex + '.json')
    temporary = path.with_suffix('.tmp')
    payload = {'name': venue.name, 'sections': list(venue.sections),
               'seats': [[*key, state] for key, state in venue.seats.items()]}
    try:
        temporary.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    return path


def load_venues(directory):
    venues, warnings = [], []
    for path in sorted(Path(directory).glob('*.json')):
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
            records = data['seats']
            seats = {(r, s, n): state for r, s, n, state in records}
            if len(seats) != len(records):
                raise ValueError('包含重复座位')
            venues.append(Venue(data['name'], seats, tuple(data['sections'])))
        except (ValueError, TypeError, KeyError, OSError) as exc:
            warnings.append(f'场地文件 {path.name} 读取失败：{exc}')
    return venues, warnings
