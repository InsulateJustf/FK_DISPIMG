#!/usr/bin/env python3
"""
FK_DISPIMG GUI - WPS 嵌入图片转换器
支持：
  1. 选择任意输入文件和输出位置进行转换
  2. 拖拽文件（支持多选）到窗口直接转换
  3. 拖拽多个文件到 .exe 批量转换（通过命令行参数）
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# 尝试加载拖拽支持
DND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    TkinterDnD = tk.Tk

from converter import wps_image_converter


def _parse_dnd_paths(data):
    """解析 tkinterdnd2 拖拽数据，支持多个文件路径（含空格）"""
    paths = []
    data = data.strip()
    if data.startswith('{') or data.startswith('"'):
        # tkinterdnd2 格式：多个路径用 } { 或 " " 分隔
        import re
        # 匹配 {path} 或无空格的裸路径
        paths = re.findall(r'\{([^}]+)\}|(\S+)', data)
        paths = [a or b for a, b in paths]
    else:
        paths = data.split()
    return [p.strip() for p in paths if p.strip()]


def _generate_output_path(input_path):
    """根据输入路径生成默认输出路径：同目录下 xxx_converted.xlsx"""
    base, ext = os.path.splitext(input_path)
    return f"{base}_converted{ext}"


class FKApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FK_DISPIMG - WPS嵌入图片转换器")
        self.root.geometry("660x560")
        self.root.minsize(560, 480)
        self.root.resizable(True, True)

        self.converting = False

        self._build_ui()
        self._setup_dnd()

        # 处理命令行参数（拖拽文件到 exe 的场景，支持多文件）
        cli_files = [
            a for a in sys.argv[1:]
            if os.path.isfile(a) and a.lower().endswith('.xlsx')
        ]
        if cli_files:
            self._batch_convert(cli_files)

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}

        # --- 文件列表 ---
        frame_files = tk.LabelFrame(self.root, text="待转换文件列表", padx=8, pady=6)
        frame_files.pack(fill=tk.BOTH, expand=True, **pad)

        btn_row = tk.Frame(frame_files)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        tk.Button(btn_row, text="添加文件...", command=self._add_files).pack(side=tk.LEFT)
        tk.Button(btn_row, text="添加文件夹...", command=self._add_folder).pack(side=tk.LEFT, padx=(6, 0))
        tk.Button(btn_row, text="清空列表", command=self._clear_list).pack(side=tk.RIGHT)

        list_frame = tk.Frame(frame_files)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(
            list_frame, font=("Menlo", 11),
            selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        self.file_count_label = tk.Label(frame_files, text="共 0 个文件", fg="#777")
        self.file_count_label.pack(anchor=tk.W)

        # --- 拖拽区域 ---
        self.drop_frame = tk.Frame(self.root, relief=tk.GROOVE, borderwidth=2, height=60)
        self.drop_frame.pack(fill=tk.X, padx=12, pady=(4, 4))
        self.drop_frame.pack_propagate(False)

        if DND_AVAILABLE:
            drop_text = "📂 将 .xlsx 文件拖拽到此处（支持多文件）"
        else:
            drop_text = "📂 拖拽不可用 (请安装 tkinterdnd2)"

        self.drop_label = tk.Label(
            self.drop_frame, text=drop_text,
            font=("", 12), fg="#555", justify=tk.CENTER
        )
        self.drop_label.pack(expand=True)

        # --- 操作按钮 ---
        frame_btn = tk.Frame(self.root)
        frame_btn.pack(fill=tk.X, padx=12, pady=4)

        self.convert_btn = tk.Button(
            frame_btn, text="开始转换", font=("", 13, "bold"),
            bg="#4CAF50", fg="white", command=self._start_convert
        )
        self.convert_btn.pack(fill=tk.X, ipady=4)

        # --- 日志 ---
        frame_log = tk.LabelFrame(self.root, text="转换日志", padx=8, pady=6)
        frame_log.pack(fill=tk.BOTH, expand=True, **pad)

        self.log_area = scrolledtext.ScrolledText(
            frame_log, height=6, font=("Menlo", 11), state=tk.DISABLED, wrap=tk.WORD
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def _setup_dnd(self):
        if not DND_AVAILABLE:
            return
        for widget in [self.root, self.drop_frame, self.drop_label]:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self._on_drop)
            widget.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            widget.dnd_bind('<<DragLeave>>', self._on_drag_leave)

    def _on_drag_enter(self, event):
        self.drop_frame.configure(bg="#E3F2FD")
        self.drop_label.configure(bg="#E3F2FD", fg="#1565C0")

    def _on_drag_leave(self, event):
        self.drop_frame.configure(bg=self.root.cget("bg"))
        self.drop_label.configure(bg=self.root.cget("bg"), fg="#555")

    def _on_drop(self, event):
        self.drop_frame.configure(bg=self.root.cget("bg"))
        self.drop_label.configure(bg=self.root.cget("bg"), fg="#555")

        paths = _parse_dnd_paths(event.data)
        xlsx_files = [p for p in paths if os.path.isfile(p) and p.lower().endswith('.xlsx')]

        # 如果拖入的是文件夹，扫描里面的 xlsx
        folders = [p for p in paths if os.path.isdir(p)]
        for folder in folders:
            for f in os.listdir(folder):
                if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                    xlsx_files.append(os.path.join(folder, f))

        if not xlsx_files:
            messagebox.showwarning("提示", "请拖入 .xlsx 文件")
            return

        self._add_to_list(xlsx_files)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 WPS XLSX 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if paths:
            self._add_to_list(list(paths))

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 XLSX 的文件夹")
        if folder:
            xlsx_files = []
            for f in os.listdir(folder):
                if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                    xlsx_files.append(os.path.join(folder, f))
            if xlsx_files:
                self._add_to_list(xlsx_files)
            else:
                messagebox.showinfo("提示", "该文件夹中没有找到 .xlsx 文件")

    def _add_to_list(self, paths):
        """添加文件到列表（去重）"""
        existing = set(self.file_listbox.get(0, tk.END))
        added = 0
        for p in paths:
            if p not in existing:
                self.file_listbox.insert(tk.END, p)
                added += 1
        self._update_count()

    def _clear_list(self):
        self.file_listbox.delete(0, tk.END)
        self._update_count()

    def _update_count(self):
        count = self.file_listbox.size()
        self.file_count_label.configure(text=f"共 {count} 个文件")

    def _log(self, msg):
        def _append():
            self.log_area.configure(state=tk.NORMAL)
            self.log_area.insert(tk.END, msg + "\n")
            self.log_area.see(tk.END)
            self.log_area.configure(state=tk.DISABLED)
        self.root.after(0, _append)

    def _batch_convert(self, file_list):
        """批量转换文件列表（命令行或拖拽到 exe）"""
        self._add_to_list(file_list)
        self.root.after(200, self._start_convert)

    def _start_convert(self):
        if self.converting:
            return

        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("提示", "请先添加要转换的文件")
            return

        # 验证所有文件
        invalid = [f for f in files if not os.path.isfile(f)]
        if invalid:
            messagebox.showerror("错误", f"以下文件不存在:\n" + "\n".join(invalid[:5]))
            return

        self.converting = True
        self.convert_btn.configure(state=tk.DISABLED, text="转换中...", bg="#9E9E9E")

        # 清空日志
        self.log_area.configure(state=tk.NORMAL)
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state=tk.DISABLED)

        def run():
            total = len(files)
            success_count = 0
            fail_count = 0

            for i, inp in enumerate(files, 1):
                out = _generate_output_path(inp)
                self._log(f"\n[{i}/{total}] {os.path.basename(inp)}")

                try:
                    success, msg, count = wps_image_converter(inp, out, log_callback=self._log)
                    if success:
                        success_count += 1
                    else:
                        fail_count += 1
                        self._log(f"  ✗ 失败: {msg}")
                except Exception as e:
                    fail_count += 1
                    self._log(f"  ✗ 异常: {e}")

            summary = f"\n{'=' * 50}\n批量转换完成！成功 {success_count} 个，失败 {fail_count} 个，共 {total} 个文件。"
            self._log(summary)
            self.root.after(0, lambda: self._batch_done(success_count, fail_count, total))

        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def _batch_done(self, success_count, fail_count, total):
        self.converting = False
        self.convert_btn.configure(state=tk.NORMAL, text="开始转换", bg="#4CAF50")

        if fail_count == 0:
            messagebox.showinfo("转换完成", f"全部 {success_count} 个文件转换成功！")
        else:
            messagebox.showwarning("转换完成", f"成功 {success_count} 个，失败 {fail_count} 个，共 {total} 个。\n请查看日志了解详情。")


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    FKApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
