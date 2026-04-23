import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import io
import traceback
import csv


def choose_sheet_for_file(parent, filename):
    # 如果不是 Excel 文件（例如 CSV），则没有 sheet 可选
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ('.xls', '.xlsx'):
        messagebox.showinfo('提示', f'文件 {os.path.basename(filename)} 不是 Excel 文件，无法选择 sheet。')
        return None
    try:
        ef_engine = 'xlrd' if ext == '.xls' else 'openpyxl'
        xl = pd.ExcelFile(filename, engine=ef_engine)
        sheets = xl.sheet_names
    except Exception as e:
        messagebox.showerror('读取错误', f'无法读取文件: {filename}\n{e}')
        return None

    dialog = tk.Toplevel(parent)
    dialog.title(f'为文件选择 sheet：{os.path.basename(filename)}')
    dialog.transient(parent)
    dialog.grab_set()
    dialog.resizable(False, False)

    var = tk.StringVar(value=sheets[0] if sheets else '')

    ttk.Label(dialog, text=os.path.basename(filename)).pack(padx=10, pady=(10, 4))
    combo = ttk.Combobox(dialog, values=sheets, textvariable=var, state='readonly')
    combo.pack(fill='x', padx=10)

    result = {'sheet': None}

    def on_ok():
        result['sheet'] = var.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text='确定', command=on_ok).pack(side='left', padx=6)
    ttk.Button(btn_frame, text='跳过/取消', command=on_cancel).pack(side='left')

    # center dialog over parent
    try:
        parent.update_idletasks()
        dialog.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        dialog.geometry(f'+{x}+{y}')
    except Exception:
        pass

    parent.wait_window(dialog)
    return result['sheet']


def _read_csv_with_encodings(path):
    """尝试使用多个编码读取 CSV，并尽量自动检测分隔符与容错不一致列。

    尝试顺序（对每种编码）：
    1. 使用 engine='python', sep=None 自动嗅探分隔符
    2. 使用 csv.Sniffer 在采样文本上嗅探分隔符并重读
    3. 使用 engine='python', sep=None 并设置 on_bad_lines='warn'（若可用）
    4. 使用 engine='python', sep=None 并设置 on_bad_lines='skip'（若可用）

    编码尝试顺序：utf-8, utf-8-sig, gbk, gb18030, cp1252, latin1, utf-16, utf-32
    成功返回 DataFrame，否则抛出最后的异常。
    """
    encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb18030', 'cp1252', 'latin1', 'utf-16', 'utf-32']
    last_exc = None
    for enc in encodings:
        try:
            # 1) 最先尝试 engine='python' + sep=None，让 pandas 尝试自动嗅探分隔符
            try:
                df = pd.read_csv(path, header=None, dtype=str, encoding=enc, sep=None, engine='python')
                return df
            except Exception as e1:
                last_exc = e1
                # 2) 再尝试用 csv.Sniffer 自行嗅探分隔符并重读（对复杂定界更稳健）
                try:
                    with open(path, 'r', encoding=enc, errors='replace') as fh:
                        sample = fh.read(8192)
                        # 如果样本太短，Sniffer 可能抛出异常
                        dialect = csv.Sniffer().sniff(sample)
                        delim = dialect.delimiter
                    df = pd.read_csv(path, header=None, dtype=str, encoding=enc, sep=delim, engine='python')
                    return df
                except Exception:
                    # 3) 尝试使用 on_bad_lines='warn'（pandas >=1.3）来跳过或警告不规则行
                    try:
                        df = pd.read_csv(path, header=None, dtype=str, encoding=enc, sep=None, engine='python', on_bad_lines='warn')
                        return df
                    except TypeError:
                        # 旧 pandas 版本可能不支持 on_bad_lines 参数，继续下一个尝试
                        pass
                    except Exception:
                        last_exc = Exception('on_bad_lines=warn 读取失败')
                    # 4) 最后尝试 on_bad_lines='skip' 跳过坏行
                    try:
                        df = pd.read_csv(path, header=None, dtype=str, encoding=enc, sep=None, engine='python', on_bad_lines='skip')
                        return df
                    except TypeError:
                        pass
                    except Exception:
                        last_exc = Exception('on_bad_lines=skip 读取失败')
                # 如果以上都没成功，继续尝试下一个编码
                continue
        except Exception as e:
            last_exc = e
            continue
    # 所有编码都失败，抛出最后一个异常
    raise last_exc if last_exc is not None else UnicodeError('无法检测编码或解析 CSV')


class ExcelMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel 合并工具')
        self.files = []
        self.sheets_map = {}
        self.rows_map = {}  # filename -> (start, end) inclusive, 1-based
        self.default_sheet_map = {}  # filename -> default sheet name (first sheet)

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=8)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='Excel 文件列表：').pack(anchor='w')
        tree_frame = ttk.Frame(frm)
        tree_frame.pack(fill='both', expand=True)
        # 使用 Treeview 显示文件名、sheet 和 行范围，便于对齐和即时刷新
        self.tree = ttk.Treeview(tree_frame, columns=('sheet', 'rows'), show='tree headings', selectmode='extended')
        # 第一列使用 #0 显示文件名
        self.tree.heading('#0', text='File')
        self.tree.column('#0', width=360, anchor='w')
        self.tree.heading('sheet', text='Sheet')
        self.tree.column('sheet', width=160, anchor='center')
        self.tree.heading('rows', text='Rows')
        self.tree.column('rows', width=140, anchor='center')
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True, side='left')

        btns = ttk.Frame(frm)
        btns.pack(fill='x', pady=6)
        ttk.Button(btns, text='添加文件', command=self.add_files).pack(side='left', padx=4)
        ttk.Button(btns, text='移除选中', command=self.remove_selected).pack(side='left', padx=4)
        ttk.Button(btns, text='清空列表', command=self.clear_list).pack(side='left', padx=4)
        ttk.Button(btns, text='为选中文件选择 sheet', command=self.select_sheets_for_selected).pack(side='left', padx=4)
        ttk.Button(btns, text='为选中文件设置行范围', command=self.set_rows_for_selected).pack(side='left', padx=4)
        ttk.Button(btns, text='清除所选文件行范围', command=self.clear_rows_for_selected).pack(side='left', padx=4)
        # 将选中的 CSV 文件转换为 Excel
        ttk.Button(btns, text='CSV -> Excel', command=self.convert_csv_to_excel).pack(side='left', padx=4)

        sep = ttk.Separator(frm, orient='horizontal')
        sep.pack(fill='x', pady=6)

        # 行范围设置提示区域（将显示当前已设置的行范围）
        rows_status = ttk.Frame(frm)
        rows_status.pack(fill='x', pady=(4, 8))
        ttk.Label(rows_status, text='已为部分文件设置行范围（1-based，包含结束行）。使用上方按钮设置/清除。').pack(anchor='w')

        sep2 = ttk.Separator(frm, orient='horizontal')
        sep2.pack(fill='x', pady=6)

        out_frame = ttk.Frame(frm)
        out_frame.pack(fill='x')
        ttk.Label(out_frame, text='输出文件：').pack(side='left')
        self.out_entry = ttk.Entry(out_frame)
        self.out_entry.pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(out_frame, text='选择', command=self.choose_output).pack(side='left')
        # 当前仅支持按行追加（简单追加每个文件的所有行，不处理表头）
        ttk.Label(out_frame, text='模式: 按行追加（简单追加每个文件的所有行）').pack(side='left', padx=(8,2))
        ttk.Label(out_frame, text='起始行(1-based，所有文件从该行开始复制):').pack(side='left', padx=(8,2))
        self.global_start_entry = ttk.Entry(out_frame, width=6)
        self.global_start_entry.insert(0, '1')
        self.global_start_entry.pack(side='left')
        # 当全局起始行改变时自动刷新列表显示（离开输入框或按回车）
        try:
            self.global_start_entry.bind('<FocusOut>', lambda e: self._refresh_list())
            self.global_start_entry.bind('<Return>', lambda e: self._refresh_list())
        except Exception:
            pass

        action_frame = ttk.Frame(frm)
        action_frame.pack(fill='x', pady=8)
        ttk.Button(action_frame, text='合并并保存', command=self.merge_and_save).pack(side='left')
        ttk.Button(action_frame, text='退出', command=self.root.quit).pack(side='left', padx=8)

    def add_files(self):
        files = filedialog.askopenfilenames(title='选择文件（Excel/CSV，可多选）', filetypes=[('Excel or CSV', '*.xlsx;*.xls;*.csv'), ('Excel Files', '*.xlsx;*.xls'), ('CSV Files', '*.csv')])
        if not files:
            return
        for f in files:
            if f not in self.files:
                self.files.append(f)
                # 预抓取默认 sheet 名称以便在列表显示（仅对 Excel 文件）
                ext = os.path.splitext(f)[1].lower()
                if ext in ('.xls', '.xlsx'):
                    try:
                        ef_engine = 'xlrd' if ext == '.xls' else 'openpyxl'
                        xl = pd.ExcelFile(f, engine=ef_engine)
                        self.default_sheet_map[f] = xl.sheet_names[0] if xl.sheet_names else ''
                    except Exception:
                        self.default_sheet_map[f] = ''
                else:
                    # 对于 CSV 文件，没有 sheet 概念
                    self.default_sheet_map[f] = ''
        self._refresh_list()

    def remove_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            return
        for iid in sel:
            filename = iid
            if filename in self.files:
                self.files.remove(filename)
            self.sheets_map.pop(filename, None)
            self.rows_map.pop(filename, None)
            self.default_sheet_map.pop(filename, None)
        self._refresh_list()

    def clear_list(self):
        self.files = []
        self.sheets_map = {}
        self.rows_map = {}
        self.default_sheet_map = {}
        self._refresh_list()

    def select_sheets_for_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo('提示', '请先选中要设置 sheet 的文件（在列表中选择）。')
            return
        skipped = []
        for iid in sel:
            filename = iid
            ext = os.path.splitext(filename)[1].lower()
            if ext not in ('.xls', '.xlsx'):
                skipped.append(os.path.basename(filename))
                continue
            sheet = choose_sheet_for_file(self.root, filename)
            if sheet:
                self.sheets_map[filename] = sheet
        self._refresh_list()
        msg = 'sheet 选择完成。未设置的文件将使用第一个 sheet。'
        if skipped:
            msg += '\n以下文件为 CSV，无法选择 sheet：\n' + '\n'.join(skipped)
        messagebox.showinfo('完成', msg)

    def set_rows_for_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo('提示', '请先选中要设置行范围的文件。')
            return
        for iid in sel:
            filename = iid
            cur = self.rows_map.get(filename, (None, None))
            res = self._ask_row_range_dialog(filename, cur[0], cur[1])
            if res is None:
                # 用户取消或开始为空 -> 不改变该文件设置
                continue
            start, end = res
            self.rows_map[filename] = (start, end)
        # 刷新列表并展示当前所有有设置的文件范围
        self._refresh_list()
        if self.rows_map:
            items = [f"{os.path.basename(k)}: {v[0]}-{v[1] if v[1] is not None else 'end'}" for k, v in self.rows_map.items()]
            messagebox.showinfo('已设置行范围', '\n'.join(items))

    def clear_rows_for_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo('提示', '请先选中要清除行范围的文件。')
            return
        removed = []
        for iid in sel:
            filename = iid
            if filename in self.rows_map:
                self.rows_map.pop(filename, None)
                removed.append(os.path.basename(filename))
        if removed:
            self._refresh_list()
            messagebox.showinfo('已清除', '已清除行范围：\n' + '\n'.join(removed))
        else:
            messagebox.showinfo('提示', '所选文件中没有设置行范围。')

    def choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv'), ('All Files', '*.*')])
        if path:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, path)

    def convert_csv_to_excel(self):
        """把当前在 Treeview 中选中的 CSV 文件转换为同目录下的 .xlsx 文件。
        若目标文件存在，会自动生成不冲突的文件名（追加 _converted 或数字后缀）。
        """
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo('提示', '请先在列表中选择要转换的 CSV 文件（可多选）。')
            return
        converted = []
        failed = []
        skipped = []
        for iid in sel:
            f = iid
            ext = os.path.splitext(f)[1].lower()
            if ext != '.csv':
                skipped.append(os.path.basename(f))
                continue
            try:
                # 使用多编码尝试的通用读取函数
                df = _read_csv_with_encodings(f)

                base = os.path.splitext(f)[0]
                outpath = base + '.xlsx'
                # 如果文件已存在，寻找一个不冲突的文件名
                if os.path.exists(outpath):
                    i = 1
                    candidate = f"{base}_converted.xlsx"
                    while os.path.exists(candidate):
                        candidate = f"{base}_converted_{i}.xlsx"
                        i += 1
                    outpath = candidate

                # 保存为 Excel，不写列名和索引
                df.to_excel(outpath, index=False, header=False, engine='openpyxl')
                converted.append(outpath)
            except Exception as e:
                failed.append((f, str(e)))

        msg_parts = []
        if converted:
            msg_parts.append('已转换文件：\n' + '\n'.join(converted))
        if skipped:
            msg_parts.append('跳过（非 CSV）：\n' + '\n'.join(skipped))
        if failed:
            msg_parts.append('转换失败：\n' + '\n'.join([f'{a}: {b}' for a, b in failed]))
        if not msg_parts:
            messagebox.showinfo('提示', '未找到可转换的 CSV 文件。')
        else:
            messagebox.showinfo('转换结果', '\n\n'.join(msg_parts))

    def _ask_row_range_dialog(self, filename, cur_start=None, cur_end=None):
        """自定义模态对话框：返回 (start:int, end:int|None)，
        如果用户取消或开始为空则返回 None。
        如果结束为空或留空则返回 end=None 表示到末尾。
        """
        dialog = tk.Toplevel(self.root)
        dialog.title(f'设置行范围：{os.path.basename(filename)}')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frm = ttk.Frame(dialog, padding=8)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='开始行 (1-based, 必填)').pack(anchor='w')
        start_var = tk.StringVar(value=str(cur_start) if cur_start is not None else '')
        start_entry = ttk.Entry(frm, textvariable=start_var)
        start_entry.pack(fill='x')

        ttk.Label(frm, text='结束行 (包含, 留空表示到末尾)').pack(anchor='w', pady=(6,0))
        end_var = tk.StringVar(value=str(cur_end) if cur_end is not None else '')
        end_entry = ttk.Entry(frm, textvariable=end_var)
        end_entry.pack(fill='x')

        result = {'ok': False, 'start': None, 'end': None}

        def on_ok():
            s = start_var.get().strip()
            e = end_var.get().strip()
            if s == '':
                # 开始为空视为取消/不设置
                result['ok'] = False
                dialog.destroy()
                return
            try:
                si = int(s)
                if si < 1:
                    raise ValueError()
            except Exception:
                messagebox.showerror('错误', '开始行必须是大于等于1的整数')
                return
            if e == '':
                ei = None
            else:
                try:
                    ei = int(e)
                    if ei < si:
                        messagebox.showerror('错误', '结束行不能小于开始行')
                        return
                except Exception:
                    messagebox.showerror('错误', '结束行必须为空或整数')
                    return
            result['ok'] = True
            result['start'] = si
            result['end'] = ei
            dialog.destroy()

        def on_cancel():
            result['ok'] = False
            dialog.destroy()

        btns = ttk.Frame(frm)
        btns.pack(pady=8)
        ttk.Button(btns, text='确定', command=on_ok).pack(side='left', padx=6)
        ttk.Button(btns, text='取消', command=on_cancel).pack(side='left')

        # center dialog over root
        try:
            self.root.update_idletasks()
            dialog.update_idletasks()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            dw = dialog.winfo_width()
            dh = dialog.winfo_height()
            x = rx + (rw - dw) // 2
            y = ry + (rh - dh) // 2
            dialog.geometry(f'+{x}+{y}')
        except Exception:
            pass

        self.root.wait_window(dialog)
        if result['ok']:
            return (result['start'], result['end'])
        return None

    def merge_and_save(self):
        if not self.files:
            messagebox.showerror('错误', '没有要合并的文件。请先添加文件。')
            return

        outpath = self.out_entry.get().strip()
        if not outpath:
            outpath = os.path.join(os.getcwd(), 'merged.xlsx')

        dfs = []
        failed = []
        # 解析全局起始行（1-based），若无效则使用 1
        gtext = self.global_start_entry.get().strip() if hasattr(self, 'global_start_entry') else '1'
        try:
            gstart = int(gtext) if gtext != '' else 1
            if gstart < 1:
                gstart = 1
        except ValueError:
            gstart = 1

        for f in self.files:
            try:
                ext = os.path.splitext(f)[1].lower()
                # 以原始行读取，不把任何行当作表头
                if ext in ('.xls', '.xlsx'):
                    sheet = self.sheets_map.get(f, None)
                    sheet_to_use = sheet if sheet is not None else 0
                    read_engine = 'xlrd' if ext == '.xls' else 'openpyxl'
                    df = pd.read_excel(f, sheet_name=sheet_to_use, engine=read_engine, header=None)
                else:
                    # CSV 文件
                        df = _read_csv_with_encodings(f)
                # 计算应当截取的起止行（基于原始文件的 0-based 索引）
                global_start_idx = gstart - 1
                if f in self.rows_map:
                    start, end = self.rows_map[f]
                    start_idx = max(global_start_idx, (start - 1) if start is not None else global_start_idx)
                    end_idx = end if end is not None else None
                    # end_idx 是 1-based 包含，需要在 iloc 中作为 exclusive end 使用（即 end_idx as-is）
                    df = df.iloc[start_idx:end_idx]
                else:
                    # 只应用全局起始行
                    if global_start_idx > 0:
                        df = df.iloc[global_start_idx:]
                dfs.append(df)
            except Exception as e:
                failed.append((f, str(e)))

        if not dfs:
            messagebox.showerror('错误', '所有文件读取失败，无法合并。')
            return

        try:
            # 按行追加：直接 concat，保存时不写列名（header=False）以避免产生数字列名行
            merged = pd.concat(dfs, ignore_index=True, sort=False, axis=0)
            outdir = os.path.dirname(outpath)
            if outdir and not os.path.exists(outdir):
                os.makedirs(outdir, exist_ok=True)
            # 根据输出路径后缀选择保存为 Excel 或 CSV
            if outpath.lower().endswith('.csv'):
                # 保存为 CSV，不写列名
                merged.to_csv(outpath, index=False, header=False, encoding='utf-8-sig')
            else:
                merged.to_excel(outpath, index=False, header=False, engine='openpyxl')
            msg = f'合并完成，已保存到：{outpath}'
            if failed:
                msg += '\n注意：部分文件读取失败：\n' + '\n'.join([f'{a}: {b}' for a, b in failed])
            messagebox.showinfo('完成', msg)
        except Exception as e:
            messagebox.showerror('保存失败', str(e))

    def _refresh_list(self):
        # 清空 Treeview 并重新插入所有文件行
        for item in self.tree.get_children():
            self.tree.delete(item)
        # 解析全局起始行以在无 per-file 设置时显示默认开始
        gtext = self.global_start_entry.get().strip() if hasattr(self, 'global_start_entry') else '1'
        try:
            gstart = int(gtext) if gtext != '' else 1
            if gstart < 1:
                gstart = 1
        except Exception:
            gstart = 1

        for f in self.files:
            fname = os.path.basename(f)
            # sheet（优先显示用户选择，其次默认已抓取的第一个 sheet）
            ext = os.path.splitext(f)[1].lower()
            if ext in ('.xls', '.xlsx'):
                sheet_name = self.sheets_map.get(f, None) or self.default_sheet_map.get(f, '') or ''
            else:
                sheet_name = 'CSV'
            # 行范围显示：优先 per-file，否则显示 global start 到 end
            if f in self.rows_map:
                s, e = self.rows_map[f]
                rows_str = f"{s}-{e if e is not None else 'end'}"
            else:
                rows_str = f"{gstart}-end"
            # 使用文件路径作为 iid，这样 Treeview selection 返回的就是文件路径
            try:
                self.tree.insert('', 'end', iid=f, text=fname, values=(sheet_name, rows_str))
            except Exception:
                # 如果插入失败（iid 非法），使用自动生成 iid
                self.tree.insert('', 'end', text=fname, values=(sheet_name, rows_str))
        # 更新完文件列表后，可在需要时使用 self.rows_map / self.sheets_map 查看各文件设置


def main():
    root = tk.Tk()
    app = ExcelMergeApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
