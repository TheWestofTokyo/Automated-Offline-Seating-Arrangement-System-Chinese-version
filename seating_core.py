"""自动化排座管理系统：场地、人员解析、排座与导出。"""
from venue import Venue, STATUS_LABELS, import_venue, save_venue, load_venues

from openpyxl import Workbook, load_workbook

from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from openpyxl.utils import get_column_letter

SECTION_MAX_WIDTH = {
    'A': 7,
    'B': 16,
    'C': 18,
    'D': 16,
    'E': 7,
}

SECTION_ORDER = ['A', 'B', 'C', 'D', 'E']

SEAT_RANGES = {
    # 对称布局：A区与E区镜像同宽，B区与D区镜像同宽，编号从左到右连续
    1:  {'A': (1, 7),   'B': (8, 16),  'C': (17, 25), 'D': (26, 34), 'E': (35, 41)},
    2:  {'A': (1, 6),   'B': (7, 16),  'C': (17, 26), 'D': (27, 36), 'E': (37, 42)},
    3:  {'A': (1, 6),   'B': (7, 17),  'C': (18, 28), 'D': (29, 39), 'E': (40, 45)},
    4:  {'A': (1, 5),   'B': (6, 17),  'C': (18, 30), 'D': (31, 42), 'E': (43, 47)},
    5:  {'A': (1, 3),   'B': (4, 16),  'C': (17, 30), 'D': (31, 43), 'E': (44, 46)},
    6:  {'B': (1, 13),  'C': (14, 28), 'D': (29, 41)},
    7:  {'B': (1, 14),  'C': (15, 30), 'D': (31, 44)},
    8:  {'B': (1, 15),  'C': (16, 32), 'D': (33, 47)},
    9:  {'B': (1, 16),  'C': (17, 34), 'D': (35, 50)},
    10: {'B': (1, 12),  'C': (13, 28), 'D': (29, 40)},
    11: {'B': (1, 13),  'C': (14, 30), 'D': (31, 43)},
    12: {'B': (1, 14),  'C': (15, 32), 'D': (33, 46)},
}

PROHIBITED_SEATS = {
    ('C', 6, 21),
    ('E', 1, 35),
    ('C', 10, 20),
    ('C', 10, 21),
}

FIXED_AREAS = {
    '主席团': {'section': 'C', 'row': 1, 'start': 17, 'end': 25},
    '桌子':   {'section': 'C', 'row': 2, 'start': 17, 'end': 26},
    '老师':   {'section': 'C', 'row': 3, 'start': 18, 'end': 28},
}

def _default_seats():
    """生成所有有效座位列表。"""
    seats = []
    for row in sorted(SEAT_RANGES.keys()):
        for section in SECTION_ORDER:
            if section in SEAT_RANGES[row]:
                start, end = SEAT_RANGES[row][section]
                for seat_num in range(start, end + 1):
                    seats.append((row, section, seat_num))
    return seats

def _default_status():
    """构建座位状态字典，key=(排号, 区块, 座位号)，value=状态标签。"""
    status = {}
    for row, section, seat_num in _default_seats():
        key = (row, section, seat_num)
        # 禁排空位优先
        if (section, row, seat_num) in PROHIBITED_SEATS:
            status[key] = 'prohibited'
            continue
        # 固定席位
        assigned = False
        for role, cfg in FIXED_AREAS.items():
            if (section == cfg['section'] and row == cfg['row']
                    and cfg['start'] <= seat_num <= cfg['end']):
                status[key] = role
                assigned = True
                break
        if not assigned:
            status[key] = 'available'
    return status

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

SECTION_COLORS = {
    'A': PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'),
    'B': PatternFill(start_color='BBDEFB', end_color='BBDEFB', fill_type='solid'),
    'C': PatternFill(start_color='FFE0B2', end_color='FFE0B2', fill_type='solid'),
    'D': PatternFill(start_color='BBDEFB', end_color='BBDEFB', fill_type='solid'),
    'E': PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'),
}

AISLE_FILL = PatternFill(start_color='EFEBE9', end_color='EFEBE9', fill_type='solid')

STAGE_FILL = PatternFill(start_color='90CAF9', end_color='90CAF9', fill_type='solid')

PROHIBITED_FILL = PatternFill(start_color='FFCDD2', end_color='FFCDD2', fill_type='solid')

FIXED_FILL = {
    '主席团': PatternFill(start_color='CE93D8', end_color='CE93D8', fill_type='solid'),
    '桌子':   PatternFill(start_color='BDBDBD', end_color='BDBDBD', fill_type='solid'),
    '老师':   PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid'),
}

from collections import Counter, defaultdict

from datetime import datetime

from pathlib import Path

from tempfile import TemporaryDirectory

from uuid import uuid4

import os

from openpyxl import load_workbook

def workbook_sheets(path):
    wb = load_workbook(path, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()

def parse_personnel(filepath, sheet_name=None):
    wb = load_workbook(filepath, data_only=True)
    formulas = load_workbook(filepath, data_only=False)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active
        raw = formulas[ws.title]
        rows = [(i, row) for i, row in enumerate(ws.iter_rows(values_only=True), 1)
                if any(v is not None and str(v).strip() for v in row)]
        if len(rows) < 2:
            raise ValueError('名单需要表头和至少一行人员数据。')
        aliases = {
            'name': {'姓名', '名字', 'name', 'fullname', 'xm'},
            'class': {'班级', '专业', '单位', 'class', 'classname', 'dept', 'department', 'bj'},
            'role': {'身份', '角色', '类型', '类别', 'role', 'type', 'category', 'sf'},
        }
        columns = {}
        for i, value in enumerate(rows[0][1]):
            header = str(value or '').lower().replace(' ', '').replace('\n', '').strip()
            for field, names in aliases.items():
                if header in names:
                    if field in columns:
                        raise ValueError(f'表头存在多个{field}列，请保留一个明确的姓名、班级或身份列。')
                    columns[field] = i
        if 'name' not in columns or 'class' not in columns:
            raise ValueError('第一行非空行必须包含“姓名”和“班级”（也支持“单位”）表头。')
        people, warnings = [], []
        duplicates = {}
        for row_num, values in rows[1:]:
            def value(field):
                col = columns.get(field)
                return str(values[col]).strip() if col is not None and col < len(values) and values[col] is not None else ''
            name, cls, original_role = value('name'), value('class'), value('role')
            uncached = [field for field in columns
                        if raw.cell(row_num, columns[field] + 1).data_type == 'f' and not value(field)]
            if uncached:
                warnings.append(f'第{row_num}行：存在没有计算结果的公式，请在 Excel 中重新计算并保存；该行已跳过。')
                continue
            if not name and not cls and not original_role:
                continue
            if '主席团' in original_role:
                role = '主席团'
            elif any(word in original_role for word in ('老师', '教师', '教员')):
                role = '老师'
            else:
                role = ''
            if not name or (not cls and not role):
                warnings.append(f'第{row_num}行：缺少{"姓名" if not name else "班级/单位"}，已跳过。')
                continue
            cls = cls or role
            if original_role and not role and original_role not in {'普通', '学生', '普通学生', '普通人员'}:
                warnings.append(f'第{row_num}行：身份“{original_role}”没有专用席位，按普通人员排座，保留原身份。')
            key = (name, cls, original_role)
            if key in duplicates:
                warnings.append(f'第{row_num}行与第{duplicates[key]}行人员信息完全重复；两条均保留，请核对。')
            else:
                duplicates[key] = row_num
            people.append({'name': name, 'class_name': cls, 'role': role,
                           'source_role': original_role, 'source_row': row_num})
        if not people:
            raise ValueError('没有可排座的有效人员。\n' + '\n'.join(warnings[:12]))
        return people, warnings
    finally:
        wb.close()
        formulas.close()

def contiguous_segments(seats, sections=None):
    sections = sections or DEFAULT_VENUE.sections
    segments = []
    for key in sorted(seats, key=lambda k: (k[0], sections.index(k[1]), k[2])):
        if segments and key[:2] == segments[-1][-1][:2] and key[2] == segments[-1][-1][2] + 1:
            segments[-1].append(key)
        else:
            segments.append([key])
    return segments

def execute_assignment(personnel, strategy='front', venue=None):
    venue = venue or DEFAULT_VENUE
    if strategy not in {'front', 'compact'}:
        raise ValueError('未知排座策略。')
    status = build_seat_status(venue)
    assignments, messages = {}, []
    groups = defaultdict(list)
    for p in personnel:
        groups[p['role'] or '普通'].append(p)
    for role in ('主席团', '老师'):
        keys = sorted((k for k, v in status.items() if v == role), key=venue.sort_key)
        if len(groups[role]) > len(keys):
            raise ValueError(f'{role}有{len(groups[role])}人，专用席位仅{len(keys)}个。')
        # Preserve the supplied ceremonial order.
        for key, person in zip(keys, groups[role]):
            assignments[key] = dict(person)
            status[key] = 'assigned'
    available = [k for k, v in status.items() if v == 'available']
    if len(groups['普通']) > len(available):
        raise ValueError(f'普通人员有{len(groups["普通"])}人，可用座位仅{len(available)}个。')
    classes = defaultdict(list)
    for p in groups['普通']:
        classes[p['class_name']].append(p)
    segments = contiguous_segments(available, venue.sections)
    for cls, members in sorted(classes.items(), key=lambda item: -len(item[1])):
        offset = 0
        used = []
        while offset < len(members):
            remaining = len(members) - offset
            candidates = segments
            if strategy == 'front':
                front = min(s[0][0] for s in candidates)
                candidates = [s for s in candidates if s[0][0] == front]
                fits = [s for s in candidates if len(s) >= remaining]
                segment = fits[0] if fits else max(candidates, key=len)
            else:
                # A complete continuous block first; otherwise largest remaining
                # segments minimize this class's number of fragments.
                fits = [s for s in candidates if len(s) >= remaining]
                segment = fits[0] if fits else max(candidates, key=len)
            take = min(remaining, len(segment))
            chosen = segment[:take]
            for key, person in zip(chosen, members[offset:offset + take]):
                assignments[key] = dict(person)
                status[key] = 'assigned'
            used.append(f'{chosen[0][0]}排{chosen[0][1]}区{chosen[0][2]}—{chosen[-1][2]}号（{take}人）')
            del segment[:take]
            if not segment:
                segments.remove(segment)
            offset += take
        messages.append(f'{cls} · {len(members)}人 → ' + '；'.join(used))
    validate_assignments(personnel, assignments, venue)
    messages.insert(0, '场地：' + venue.name + '\n策略：' + ('前排优先；同一排尽量减少拆分' if strategy == 'front' else '同班集中；可安排至后排'))
    messages.append(f'完成：{len(personnel)}人全部分配，无重复占座。')
    return assignments, status, messages

def validate_assignments(personnel, assignments, venue=None):
    def identity(p):
        return p['name'], p['class_name'], p['role'], p.get('source_role', ''), p.get('source_row')
    if Counter(map(identity, personnel)) != Counter(identity(p) for p in assignments.values()):
        raise ValueError('排座校验失败：人员与座位分配不一致。')
    original = build_seat_status(venue)
    for key, p in assignments.items():
        expected = p['role'] or 'available'
        if original.get(key) != expected:
            raise ValueError(f'排座校验失败：{key} 不允许安排该人员。')

def seat_ranges(numbers):
    groups = []
    for n in sorted(numbers):
        if groups and n == groups[-1][-1] + 1:
            groups[-1].append(n)
        else:
            groups.append([n])
    return '、'.join(str(g[0]) if len(g) == 1 else f'{g[0]}—{g[-1]}' for g in groups)

def generate_summary_table(assignments, seat_status, output_path, venue=None):
    venue = venue or DEFAULT_VENUE
    wb = Workbook()
    ws = wb.active
    ws.title = '班级座位分配表'
    headers = ['排号', '区域', '座位范围', '班级 / 用途', '人数', '身份 / 备注']
    ws.append([venue.name + ' · 班级座位分配'])
    ws.merge_cells('A1:F1')
    ws.append(headers)
    for row in venue.rows:
        for section in venue.sections:
            grouped = defaultdict(list)
            for key, state in seat_status.items():
                if key[:2] != (row, section):
                    continue
                if key in assignments:
                    p = assignments[key]
                    group = (p['class_name'], p.get('source_role') or p['role'] or '普通', True)
                else:
                    label = {'available': '未分配', 'prohibited': '禁排空位',
                             '桌子': '桌子预留', '老师': '老师预留', '主席团': '主席团预留'}[state]
                    group = (label, label, False)
                grouped[group].append(key[2])
            for (cls, role, occupied), numbers in grouped.items():
                ws.append([row, section, seat_ranges(numbers), cls, len(numbers) if occupied else 0, role])
    for row in ws:
        for cell in row:
            cell.font = Font(name='微软雅黑', size=10, color='263448')
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            if cell.row <= 2:
                cell.fill = PatternFill('solid', fgColor='E8EDF7')
                cell.font = Font(name='微软雅黑', size=11, bold=True)
            elif cell.row % 2:
                cell.fill = PatternFill('solid', fgColor='F5F8FC')
        ws.row_dimensions[row[0].row].height = 30
    for col, width in zip('ABCDEF', [8, 8, 30, 26, 8, 24]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'D3'
    ws.auto_filter.ref = f'A2:F{ws.max_row}'
    ws.print_title_rows = '1:2'
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    save_workbook(wb, output_path)

def save_workbook(wb, output_path):
    # Names and classes are literal text, including values starting with '='.
    for ws in wb:
        for row in ws:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
        ws.print_options.horizontalCentered = True
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_area = ws.dimensions
    wb.save(output_path)
    wb.close()

def export_results(assignments, seat_status, parent, venue=None):
    parent = Path(parent)
    target = parent / ('排座结果_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:6])
    with TemporaryDirectory(prefix='.seating-', dir=parent) as temp:
        temp = Path(temp)
        generate_visual_chart(assignments, seat_status, temp / '01_可视化座位图.xlsx', venue)
        generate_detail_list(assignments, seat_status, temp / '02_个人座位明细.xlsx', venue)
        generate_summary_table(assignments, seat_status, temp / '03_班级座位分配表.xlsx', venue)
        os.rename(temp, target)
    return str(target)

DEFAULT_VENUE = Venue('大学生活动中心礼堂', _default_status(), tuple(SECTION_ORDER))

"""Excel exporters shared by built-in and imported venues."""


"""Excel exporters shared by built-in and imported venues."""


def build_all_seats(venue=None):
    venue = venue or DEFAULT_VENUE
    return sorted(venue.seats, key=venue.sort_key)


def build_seat_status(venue=None):
    return dict((venue or DEFAULT_VENUE).seats)


def _format_table(ws, widths, title_rows='1:2'):
    for row in ws:
        for cell in row:
            cell.font = Font(name='微软雅黑', size=10, color='263448')
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            if cell.row <= 1:
                cell.fill = PatternFill('solid', fgColor='E8EDF7')
                cell.font = Font(name='微软雅黑', size=11, bold=True)
            elif cell.row % 2:
                cell.fill = PatternFill('solid', fgColor='F5F8FC')
        ws.row_dimensions[row[0].row].height = 32
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.print_title_rows = title_rows


def generate_visual_chart(assignments, seat_status, output_path, venue=None):
    venue = venue or DEFAULT_VENUE
    wb = Workbook()
    ws = wb.active
    ws.title = '可视化座位图'
    widths = venue.widths
    total_cols = 1 + sum(widths.values()) + len(venue.sections) - 1
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=total_cols)
    ws.cell(1, 2, venue.name + ' · 座位图')
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=total_cols)
    ws.cell(2, 2, '舞  台')
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 28
    ws.column_dimensions['A'].width = 9
    col = 2
    for index, section in enumerate(venue.sections):
        width = widths[section]
        fill = SECTION_COLORS.get(section, PatternFill('solid', fgColor='DCE5FF'))
        if width > 1:
            ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+width-1)
        ws.cell(4, col, section + '区')
        for display_row, row in enumerate(venue.rows, 5):
            ws.cell(display_row, 1, f'第{row}排')
            ws.row_dimensions[display_row].height = 76
            for offset, key in enumerate(venue.row_seats(row, section)):
                state = seat_status[key]
                p = assignments.get(key)
                number = key[2]
                if p:
                    text = f'{number}号\n{p["name"]}\n{p["class_name"]}'
                    if p['role']:
                        text += '\n' + p['role']
                    color = FIXED_FILL.get(p['role'], fill)
                else:
                    text = f'{number}号\n' + ('' if state == 'available' else STATUS_LABELS[state])
                    color = PROHIBITED_FILL if state == 'prohibited' else FIXED_FILL.get(state, fill)
                cell = ws.cell(display_row, col + offset, text)
                cell.fill = color
                cell.font = Font(name='微软雅黑', size=8)
                cell.border = THIN_BORDER
                cell.alignment = CENTER_ALIGN
        for c in range(col, col + width):
            ws.column_dimensions[get_column_letter(c)].width = 9
        col += width
        if index < len(venue.sections) - 1:
            ws.column_dimensions[get_column_letter(col)].width = 3
            for r in range(4, len(venue.rows) + 5):
                ws.cell(r, col).fill = AISLE_FILL
            col += 1
    for r in (1, 2, 4):
        for cell in ws[r]:
            cell.alignment = CENTER_ALIGN
            cell.font = Font(name='微软雅黑', size=14 if r == 1 else 11, bold=True)
    ws.freeze_panes = 'B5'
    ws.print_title_rows = '1:4'
    save_workbook(wb, output_path)


def generate_detail_list(assignments, seat_status, output_path, venue=None):
    venue = venue or DEFAULT_VENUE
    wb = Workbook()
    ws = wb.active
    ws.title = '个人座位明细'
    ws.append(['排号', '区块', '座位号', '姓名', '班级 / 单位', '身份', '备注'])
    for key in sorted(seat_status, key=venue.sort_key):
        p = assignments.get(key)
        if p:
            values = [*key, p['name'], p['class_name'], p.get('source_role') or p['role'] or '普通', '']
        else:
            label = '未分配' if seat_status[key] == 'available' else STATUS_LABELS[seat_status[key]]
            values = [*key, '（空）', '', label, label]
        ws.append(values)
    _format_table(ws, [8, 12, 10, 16, 22, 16, 22], '1:1')
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = ws.dimensions
    ws.oddHeader.center.text = venue.name
    save_workbook(wb, output_path)


def generate_seat_template(output_path, venue=None):
    venue = venue or DEFAULT_VENUE
    wb = Workbook()
    ws = wb.active
    ws.title = '座位表'
    ws.append(['场地名称', '排号', '区域', '座位号', '用途'])
    for key in sorted(venue.seats, key=venue.sort_key):
        ws.append([venue.name, *key, STATUS_LABELS[venue.seats[key]]])
    _format_table(ws, [32, 10, 14, 12, 18], '1:1')
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = ws.dimensions
    save_workbook(wb, output_path)
