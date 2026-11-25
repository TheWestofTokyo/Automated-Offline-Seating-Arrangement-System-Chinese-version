"""自动化排座管理系统。数据仅在本机处理。"""
import os
from pathlib import Path
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import seating_core as core

BG = '#F3F5FA'
INK = '#17233B'
MUTED = '#78859B'
BLUE = '#4357DC'
PALETTE = ['#DCE5FF', '#CDEEE5', '#FFE6C4', '#E5DDF8', '#D4ECF9', '#F9DCE5', '#E9EDC8']


class SeatingApp(ctk.CTk):
    def __init__(self, library_dir=None):
        ctk.set_appearance_mode('light')
        super().__init__()
        self.title('自动化排座管理系统 v6.0')
        self.geometry('1240x740')
        self.minsize(1080, 640)
        self.configure(fg_color=BG)
        self.people = []
        self.assignments = {}
        home = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        self.library_dir = Path(library_dir) if library_dir is not None else home / '场地库'
        self.venue = core.DEFAULT_VENUE
        self.venues = {self.venue.name: self.venue}
        saved, warnings = core.load_venues(self.library_dir)
        for venue in saved:
            self._unique_venue_name(venue)
            self.venues[venue.name] = venue
        self.seat_status = core.build_seat_status(self.venue)
        self.input_path = ''
        self.export_dir = ''
        self.busy = False
        self.jobs = queue.Queue()
        self.seat_items = {}
        self.search = tk.StringVar()
        self.strategy = tk.StringVar(value='前排优先')
        self._build_ui()
        if warnings:
            self._write_log('\n'.join(warnings))
        self.after(100, self._poll)
        self.protocol('WM_DELETE_WINDOW', self._close)
        self.search.trace_add('write', lambda *_: self._refresh_table())

    def label(self, parent, text, size=12, color=INK, bold=False, **kw):
        return ctk.CTkLabel(parent, text=text, text_color=color,
                            font=ctk.CTkFont('Microsoft YaHei UI', size, weight='bold' if bold else 'normal'), **kw)

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        sidebar = ctk.CTkFrame(self, width=270, corner_radius=0, fg_color='#FFFFFF')
        sidebar.grid(row=0, column=0, sticky='nsew')
        sidebar.grid_propagate(False)
        side = ctk.CTkScrollableFrame(sidebar, fg_color='transparent', width=238)
        side.pack(fill='both', expand=True, padx=10, pady=18)
        self.label(side, 'S /  SEATING', 22, BLUE, True).pack(anchor='w', padx=16, pady=(6, 2))
        self.label(side, '让每一个座位，都有安排。', 11, MUTED).pack(anchor='w', padx=16, pady=(0, 16))
        self.label(side, '当前场地', 13, bold=True).pack(anchor='w', padx=16)
        self.venue_select = ctk.CTkOptionMenu(side, values=list(self.venues), command=self._switch_venue,
                                             fg_color='#EEF1FA', button_color=BLUE, text_color=INK, height=34)
        self.venue_select.set(self.venue.name)
        self.venue_select.pack(fill='x', padx=16, pady=(8, 8))
        venue_buttons = ctk.CTkFrame(side, fg_color='transparent')
        venue_buttons.pack(fill='x', padx=16)
        self.venue_import_btn = ctk.CTkButton(venue_buttons, text='导入新座位表', width=106, height=31,
                                              fg_color=BLUE, command=self._on_venue_import)
        self.venue_import_btn.pack(side='left', padx=(0, 5))
        self.template_btn = ctk.CTkButton(venue_buttons, text='导出模板', width=88, height=31,
                                         fg_color='#EDF1FA', text_color=INK, hover_color='#DEE5F5', command=self._on_template)
        self.template_btn.pack(side='left')
        self.venue_summary = self.label(side, self._venue_description(), 10, MUTED, justify='left', wraplength=208)
        self.venue_summary.pack(anchor='w', padx=16, pady=(10, 20))
        self.label(side, '01   导入人员名单', 15, bold=True).pack(anchor='w', padx=16)
        self.label(side, '支持 .xlsx，可选择工作表', 11, MUTED).pack(anchor='w', padx=16, pady=(4, 12))
        self.import_btn = ctk.CTkButton(side, text='＋  选择 Excel 名单', height=43,
                                       fg_color=BLUE, hover_color='#3547C2', command=self._on_import)
        self.import_btn.pack(fill='x', padx=16)
        self.file_label = self.label(side, '尚未导入名单', 11, MUTED, wraplength=205, justify='left')
        self.file_label.pack(anchor='w', padx=16, pady=(10, 28))
        self.label(side, '02   选择排座方式', 15, bold=True).pack(anchor='w', padx=16)
        self.mode = ctk.CTkSegmentedButton(side, values=['前排优先', '同班集中'], variable=self.strategy,
                                          selected_color=BLUE, command=self._strategy_changed, height=35)
        self.mode.pack(fill='x', padx=16, pady=(14, 10))
        self.mode_hint = self.label(side, '从前往后排座，同一排尽量少拆分。\n允许同一区块容纳多个班级。',
                                    11, MUTED, wraplength=205, justify='left')
        self.mode_hint.pack(anchor='w', padx=16, pady=(0, 16))
        self.assign_btn = ctk.CTkButton(side, text='开始自动排座  →', height=44, fg_color=INK,
                                       hover_color='#293B5D', state='disabled', command=self._on_assign)
        self.assign_btn.pack(fill='x', padx=16, pady=(0, 18))
        self.label(side, '名单需含：姓名、班级（或单位）\n身份可选；老师可不填班级。\n数据仅在本机处理。',
                   10, MUTED, justify='left').pack(anchor='w', padx=16, pady=(2, 12))

        main = ctk.CTkFrame(self, fg_color='transparent')
        main.grid(row=0, column=1, sticky='nsew', padx=26, pady=22)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)
        header = ctk.CTkFrame(main, fg_color='transparent')
        header.grid(row=0, column=0, sticky='ew')
        self.label(header, '自动化排座管理系统', 27, bold=True).pack(anchor='w')
        self.venue_header = self.label(header, self.venue.name + '  /  SEATING MANAGEMENT', 11, MUTED)
        self.venue_header.pack(anchor='w', pady=(5, 18))
        cards = ctk.CTkFrame(main, fg_color='transparent')
        cards.grid(row=1, column=0, sticky='ew', pady=(0, 16))
        self.stats = []
        for i, (title, detail) in enumerate([('导入人数', '人'), ('普通班级', '个'), ('已安排', '席'), ('普通空余', '席')]):
            cards.grid_columnconfigure(i, weight=1, uniform='cards')
            card = ctk.CTkFrame(cards, fg_color='white', corner_radius=14)
            card.grid(row=0, column=i, sticky='ew', padx=(0 if i == 0 else 6, 0 if i == 3 else 6))
            self.label(card, title, 11, MUTED).pack(anchor='w', padx=18, pady=(12, 0))
            number = self.label(card, '0' if i < 3 else str(self.venue.counts['available']), 26, BLUE if i == 2 else INK, True)
            number.pack(anchor='w', padx=18, pady=(0, 12))
            self.stats.append(number)
        workspace = ctk.CTkFrame(main, fg_color='white', corner_radius=16)
        workspace.grid(row=2, column=0, sticky='nsew')
        workspace.grid_rowconfigure(1, weight=1)
        workspace.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(workspace, fg_color='transparent')
        top.grid(row=0, column=0, sticky='ew', padx=18, pady=(12, 0))
        self.label(top, '座位安排', 16, bold=True).pack(side='left')
        self.label(top, '点击座位查看人员信息  ·  滚动查看全图', 10, MUTED).pack(side='right')
        self.tabs = ctk.CTkTabview(workspace, fg_color='white', segmented_button_selected_color=BLUE,
                                    segmented_button_selected_hover_color=BLUE, corner_radius=10)
        self.tabs.grid(row=1, column=0, sticky='nsew', padx=10, pady=(0, 6))
        for name in ['座位全景', '人员明细', '处理日志']:
            self.tabs.add(name)
        self._build_map(self.tabs.tab('座位全景'))
        self._build_table(self.tabs.tab('人员明细'))
        self.log = ctk.CTkTextbox(self.tabs.tab('处理日志'), fg_color='#F7F9FD', text_color=INK,
                                 font=ctk.CTkFont('Microsoft YaHei UI', 12), wrap='word')
        self.log.pack(fill='both', expand=True)
        self.log.configure(state='disabled')
        footer = ctk.CTkFrame(main, fg_color='transparent')
        footer.grid(row=3, column=0, sticky='ew', pady=(14, 0))
        self.status_label = self.label(footer, '准备就绪 · 请先导入人员名单', 11, MUTED)
        self.status_label.pack(side='left')
        self.export_btn = ctk.CTkButton(footer, text='导出三个 Excel 文件  ↗', width=205, height=43,
                                       fg_color=BLUE, hover_color='#3547C2', state='disabled', command=self._on_export)
        self.export_btn.pack(side='right')
        self.progress = ctk.CTkProgressBar(main, height=3, progress_color=BLUE, mode='indeterminate')
        self.progress.grid(row=4, column=0, sticky='ew', pady=(12, 0))
        self.progress.set(0)
        self._draw_map()

    def _build_map(self, parent):
        legend = ctk.CTkFrame(parent, fg_color='transparent')
        legend.pack(fill='x', padx=8, pady=(4, 7))
        for label, color in [('普通座位', '#DCE5FF'), ('主席团', '#DED5F7'), ('老师', '#FFF0B9'),
                             ('桌子预留', '#DBE0E8'), ('禁排', '#F8D6DB')]:
            self.label(legend, '●', 15, color).pack(side='left', padx=(8, 3))
            self.label(legend, label, 10, MUTED).pack(side='left', padx=(0, 6))
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='both', expand=True)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(frame, bg='white', highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(frame, orient='vertical', command=self.canvas.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(frame, orient='horizontal', command=self.canvas.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        self.canvas.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
        self.canvas.bind('<Configure>', lambda event: self._draw_map())
        self.canvas.bind('<Button-1>', self._seat_click)
        self.canvas.bind('<MouseWheel>', lambda e: self.canvas.yview_scroll(int(-e.delta / 120), 'units'))
        self.seat_hint = self.label(parent, '未导入名单 · 图中数字为座位号', 11, MUTED)
        self.seat_hint.pack(anchor='w', padx=16, pady=(6, 2))

    def _build_table(self, parent):
        ctk.CTkEntry(parent, textvariable=self.search, placeholder_text='搜索姓名、班级或座位', height=34).pack(fill='x', padx=8, pady=8)
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('Desk.Treeview', rowheight=32, background='white', fieldbackground='white',
                        foreground=INK, font=('Microsoft YaHei UI', 10), borderwidth=0)
        style.configure('Desk.Treeview.Heading', background='#EDF1FA', foreground=INK,
                        font=('Microsoft YaHei UI', 10, 'bold'), padding=9)
        style.map('Desk.Treeview', background=[('selected', '#DBE4FF')], foreground=[('selected', INK)])
        self.table = ttk.Treeview(frame, columns=('name', 'class', 'role', 'seat'), show='headings', style='Desk.Treeview')
        for col, title, width in [('name', '姓名', 100), ('class', '班级 / 单位', 170), ('role', '身份', 110), ('seat', '座位', 180)]:
            self.table.heading(col, text=title)
            self.table.column(col, width=width, anchor='center')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.table.yview)
        scrollbar.pack(side='right', fill='y')
        self.table.configure(yscrollcommand=scrollbar.set)
        self.table.pack(fill='both', expand=True)

    def _draw_map(self):
        if not hasattr(self, 'canvas'):
            return
        canvas = self.canvas
        canvas.delete('all')
        self.seat_items.clear()
        width = max(canvas.winfo_width() - 20, 870)
        widths = self.venue.widths
        width = max(width, 70 + sum(widths.values()) * 12 + (len(widths) - 1) * 18)
        unit = (width - 70 - (len(widths) - 1) * 18) / sum(widths.values())
        canvas.create_rectangle(40, 8, width - 4, 43, fill='#EDF1FC', outline='')
        canvas.create_text(width / 2, 25, text='舞  台  /  STAGE', fill=BLUE, font=('Microsoft YaHei UI', 12, 'bold'))
        colors = {cls: PALETTE[i % len(PALETTE)] for i, cls in enumerate(dict.fromkeys(p['class_name'] for p in self.people if not p['role']))}
        x = 42
        for section in self.venue.sections:
            section_width = widths[section] * unit
            canvas.create_text(x + section_width / 2, 64, text=f'{section} 区', fill=MUTED, font=('Microsoft YaHei UI', 10, 'bold'))
            for row_index, row in enumerate(self.venue.rows):
                for i, key in enumerate(self.venue.row_seats(row, section)):
                    number = key[2]
                    state = self.seat_status[key]
                    p = self.assignments.get(key)
                    role = p['role'] if p else state
                    color = {'主席团': '#DED5F7', '老师': '#FFF0B9', '桌子': '#DBE0E8', 'prohibited': '#F8D6DB'}.get(role)
                    color = color or (colors.get(p['class_name'], '#DCE5FF') if p else '#F2F4F8')
                    left, top = x + i * unit, 82 + row_index * 33
                    rect = canvas.create_rectangle(left + 1, top, left + unit - 2, top + 27,
                                                   fill=color, outline='#FFFFFF', width=1)
                    label = '×' if state == 'prohibited' else str(number)
                    txt = canvas.create_text(left + unit / 2, top + 13, text=label, fill='#41506B', font=('Segoe UI', 8))
                    self.seat_items[rect] = self.seat_items[txt] = key
            x += section_width + 18
        for index, row in enumerate(self.venue.rows):
            canvas.create_text(18, 95 + index * 33, text=f'{row}排', fill=MUTED, font=('Microsoft YaHei UI', 9))
        canvas.configure(scrollregion=(0, 0, width + 10, 94 + len(self.venue.rows) * 33))

    def _seat_click(self, event):
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        items = self.canvas.find_overlapping(x - 1, y - 1, x + 1, y + 1)
        key = next((self.seat_items[i] for i in reversed(items) if i in self.seat_items), None)
        if key:
            person = self.assignments.get(key)
            position = f'{key[0]}排 {key[1]}区 {key[2]}号'
            if person:
                detail = f'{person["name"]}  ·  {person["class_name"]}  ·  {person.get("source_role") or person["role"] or "普通"}'
            else:
                detail = {'available': '尚未分配', 'prohibited': '禁排空位', '桌子': '桌子预留',
                          '老师': '老师预留', '主席团': '主席团预留'}[self.seat_status[key]]
            self.seat_hint.configure(text=position + '   |   ' + detail)

    def _refresh_table(self):
        self.table.delete(*self.table.get_children())
        query = self.search.get().strip().lower()
        rows = [(p, f'{k[0]}排{k[1]}区{k[2]}号') for k, p in sorted(self.assignments.items(), key=lambda item: self.venue.sort_key(item[0]))] if self.assignments else [(p, '待排座') for p in self.people]
        for p, seat in rows:
            values = (p['name'], p['class_name'], p.get('source_role') or p['role'] or '普通', seat)
            if not query or query in ' '.join(values).lower():
                self.table.insert('', 'end', values=values)

    def _refresh(self):
        normal_assigned = sum(not p['role'] for p in self.assignments.values())
        values = [len(self.people), len({p['class_name'] for p in self.people if not p['role']}),
                  len(self.assignments), self.venue.counts['available'] - normal_assigned]
        for label, value in zip(self.stats, values):
            label.configure(text=str(value))
        self.assign_btn.configure(state='normal' if self.people and not self.busy else 'disabled')
        self.export_btn.configure(state='normal' if self.assignments and not self.busy else 'disabled')
        self._draw_map()
        self._refresh_table()

    def _write_log(self, text, clear=False):
        self.log.configure(state='normal')
        if clear:
            self.log.delete('1.0', 'end')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _strategy_changed(self, value):
        hint = '从前往后排座，同一排尽量少拆分。' if value == '前排优先' else '优先同班连续；必要时使用后排。'
        self.mode_hint.configure(text=hint + '\n允许同一区块容纳多个班级。')
        if self.assignments:
            self.assignments = {}
            self.seat_status = core.build_seat_status(self.venue)
            self.seat_hint.configure(text='排座方式已更改，请重新排座。')
            self.status_label.configure(text='策略已更改 · 原结果已清除')
            self._refresh()

    def _run(self, label, job, on_success):
        if self.busy:
            return
        self.busy = True
        self.status_label.configure(text=label)
        self.import_btn.configure(state='disabled')
        self.venue_select.configure(state='disabled')
        self.venue_import_btn.configure(state='disabled')
        self.template_btn.configure(state='disabled')
        self.mode.configure(state='disabled')
        self._refresh()
        self.progress.start()
        def worker():
            try:
                self.jobs.put((True, job(), on_success))
            except Exception as exc:
                self.jobs.put((False, str(exc), on_success))
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            success, result, callback = self.jobs.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.progress.stop()
            self.progress.set(0)
            self.import_btn.configure(state='normal')
            self.venue_select.configure(state='normal')
            self.venue_import_btn.configure(state='normal')
            self.template_btn.configure(state='normal')
            self.mode.configure(state='normal')
            if success:
                callback(result)
            else:
                self.status_label.configure(text='操作未完成 · 请查看错误提示')
                self._write_log('操作失败：' + result)
                messagebox.showerror('操作未完成', result, parent=self)
            self._refresh()
        self.after(100, self._poll)

    def _on_import(self):
        path = filedialog.askopenfilename(parent=self, title='选择人员名单', filetypes=[('Excel 工作簿', '*.xlsx')])
        if not path:
            return
        self._choose_sheet(path, lambda sheet: self._load_sheet(path, sheet))

    def _choose_sheet(self, path, on_select):
        def choose(sheets):
            if len(sheets) == 1:
                on_select(sheets[0])
                return
            dialog = ctk.CTkToplevel(self)
            dialog.title('选择工作表')
            dialog.geometry('390x205')
            dialog.transient(self)
            dialog.grab_set()
            self.label(dialog, '请选择要导入的工作表', 14, bold=True).pack(pady=(22, 14))
            select = ctk.CTkOptionMenu(dialog, values=sheets, width=300)
            select.pack()
            def confirm():
                sheet = select.get()
                dialog.destroy()
                on_select(sheet)
            ctk.CTkButton(dialog, text='导入该工作表', command=confirm, fg_color=BLUE).pack(pady=18)
        self._run('正在读取工作表…', lambda: core.workbook_sheets(path), choose)

    def _load_sheet(self, path, sheet):
        def loaded(result):
            people, warnings = result
            self._write_log(f'读取工作表：{sheet}\n有效记录：{len(people)} 人', clear=True)
            if warnings:
                self._write_log('\n'.join(warnings))
                text = '\n'.join(warnings[:10]) + (f'\n其余 {len(warnings)-10} 条见处理日志。' if len(warnings)>10 else '')
                if not messagebox.askyesno('请核对名单', text + '\n\n继续使用以上有效记录？重复记录不会自动删除。', parent=self):
                    self.status_label.configure(text='本次导入已取消 · 可修正名单后重试')
                    self.tabs.set('处理日志')
                    return
            self.people, self.input_path = people, path
            self.assignments = {}
            self.seat_status = core.build_seat_status(self.venue)
            self.file_label.configure(text=os.path.basename(path) + '\n工作表：' + sheet)
            self.status_label.configure(text=f'名单已就绪 · {len(people)} 人')
            self.seat_hint.configure(text='名单已导入，请点击“开始自动排座”。')
            self.tabs.set('人员明细')
        self._run('正在校验人员名单…', lambda: core.parse_personnel(path, sheet), loaded)

    def _on_assign(self):
        if not self.people:
            return
        people = [dict(p) for p in self.people]
        venue = self.venue
        strategy = 'front' if self.strategy.get() == '前排优先' else 'compact'
        def assigned(result):
            self.assignments, self.seat_status, messages = result
            self._write_log('\n'.join(messages))
            self.status_label.configure(text=f'排座完成 · {len(self.assignments)} 人均已安排')
            self.seat_hint.configure(text='排座完成 · 点击座位查看姓名、班级和身份。')
            self.tabs.set('座位全景')
        self._run('正在安排座位…', lambda: core.execute_assignment(people, strategy, venue), assigned)

    def _on_export(self):
        if not self.assignments:
            return
        parent = filedialog.askdirectory(parent=self, title='选择结果保存位置（自动创建新文件夹）',
                                          initialdir=self.export_dir or os.path.dirname(self.input_path))
        if not parent:
            return
        assignments = {k: dict(v) for k, v in self.assignments.items()}
        status = dict(self.seat_status)
        venue = self.venue
        def exported(path):
            self.export_dir = parent
            self.status_label.configure(text='导出完成 · 已保存三个 Excel 文件')
            self._write_log('结果已保存：' + path)
            if messagebox.askyesno('导出完成', '已保存座位图、个人明细、班级分配表。\n\n' + path + '\n\n打开结果文件夹？', parent=self):
                os.startfile(path)
        self._run('正在生成三个 Excel 文件…', lambda: core.export_results(assignments, status, parent, venue), exported)

    def _unique_venue_name(self, venue):
        base, suffix = venue.name, 2
        if venue is core.DEFAULT_VENUE:
            return
        while venue.name in self.venues:
            venue.name = f'{base}（{suffix}）'
            suffix += 1

    def _venue_description(self):
        counts = self.venue.counts
        return (f'{len(self.venue.rows)} 排 · {len(self.venue.sections)} 区 · 共 {len(self.venue.seats)} 位\n'
                f'普通 {counts["available"]} 席 / 主席团 {counts["主席团"]} 席\n'
                f'老师 {counts["老师"]} 席 / 桌子 {counts["桌子"]} 位 / 禁排 {counts["prohibited"]} 位')

    def _switch_venue(self, name):
        if self.busy or name == self.venue.name:
            return
        self.venue = self.venues[name]
        self.assignments = {}
        self.seat_status = core.build_seat_status(self.venue)
        self.venue_select.set(name)
        self.venue_summary.configure(text=self._venue_description())
        self.venue_header.configure(text=name + '  /  SEATING MANAGEMENT')
        self.seat_hint.configure(text='场地已切换 · 请重新排座。')
        self.status_label.configure(text='当前场地：' + name)
        self._write_log('切换场地：' + name + '；原排座结果已清除。')
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        self._refresh()

    def _on_venue_import(self):
        path = filedialog.askopenfilename(parent=self, title='选择新座位表', filetypes=[('Excel 工作簿', '*.xlsx')])
        if path:
            self._choose_sheet(path, lambda sheet: self._load_venue(path, sheet))

    def _load_venue(self, path, sheet):
        def loaded(venue):
            self._unique_venue_name(venue)
            try:
                core.save_venue(venue, self.library_dir)
            except OSError as exc:
                messagebox.showerror('场地保存失败', '请将程序放在可写入的文件夹后重试。\n' + str(exc), parent=self)
                self.status_label.configure(text='场地导入未完成 · 原场地保持不变')
                return
            self.venues[venue.name] = venue
            self.venue_select.configure(values=list(self.venues))
            self._switch_venue(venue.name)
            self._write_log('新场地已保存，重启后可继续选择。')
        self._run('正在校验新座位表…', lambda: core.import_venue(path, sheet), loaded)

    def _on_template(self):
        path = filedialog.asksaveasfilename(parent=self, title='导出当前场地座位表，可修改后重新导入',
                                           initialfile=self.venue.name + '_座位表.xlsx', defaultextension='.xlsx',
                                           filetypes=[('Excel 工作簿', '*.xlsx')])
        if not path:
            return
        venue = self.venue
        def saved(_):
            self.status_label.configure(text='座位表已导出 · 修改后可作为新场地导入')
        self._run('正在导出座位表模板…', lambda: core.generate_seat_template(path, venue), saved)

    def _close(self):
        if self.busy:
            messagebox.showinfo('正在处理', '请等待当前导入、排座或导出操作完成后再退出。', parent=self)
            return
        self.destroy()


if __name__ == '__main__':
    SeatingApp().mainloop()
