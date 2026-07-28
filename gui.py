#!/usr/bin/env python3
"""
FK_DISPIMG GUI - WPS 嵌入图片转换器
支持：
  1. 选择任意输入文件和输出位置进行转换
  2. 拖拽文件到窗口直接转换
  3. 拖拽文件到 .exe 直接转换（通过命令行参数）
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


class FKApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FK_DISPIMG - WPS嵌入图片转换器")
        self.root.geometry("640x520")
        self.root.minsize(540, 460)
        self.root.resizable(True, True)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.converting = False

        self._build_ui()
        self._setup_dnd()

        # 处理命令行参数（拖拽文件到 exe 的场景）
        if len(sys.argv) > 1:
            arg = sys.argv[1]
            if os.path.isfile(arg) and arg.lower().endswith('.xlsx'):
                self.input_path.set(arg)
                self.root.after(100, self._auto_convert)

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}

        # --- 输入文件 ---
        frame_in = tk.LabelFrame(self.root, text="输入文件 (WPS XLSX)", padx=8, pady=6)
        frame_in.pack(fill=tk.X, **pad)

        tk.Entry(frame_in, textvariable=self.input_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_in, text="浏览...", command=self._browse_input).pack(side=tk.LEFT, padx=(6, 0))

        # --- 输出文件 ---
        frame_out = tk.LabelFrame(self.root, text="输出文件", padx=8, pady=6)
        frame_out.pack(fill=tk.X, **pad)

        tk.Entry(frame_out, textvariable=self.output_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_out, text="另存为...", command=self._browse_output).pack(side=tk.LEFT, padx=(6, 0))

        # --- 拖拽区域 ---
        self.drop_frame = tk.Frame(self.root, relief=tk.GROOVE, borderwidth=2, height=80)
        self.drop_frame.pack(fill=tk.X, padx=12, pady=(8, 4))
        self.drop_frame.pack_propagate(False)

        if DND_AVAILABLE:
            drop_text = "📂 将 .xlsx 文件拖拽到此处"
        else:
            drop_text = "📂 拖拽不可用 (请安装 tkinterdnd2)\n请使用上方浏览按钮选择文件"

        self.drop_label = tk.Label(
            self.drop_frame, text=drop_text,
            font=("", 13), fg="#555", justify=tk.CENTER
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
            frame_log, height=8, font=("Menlo", 11), state=tk.DISABLED, wrap=tk.WORD
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def _setup_dnd(self):
        if not DND_AVAILABLE:
            return
        # 绑定拖拽到整个窗口和 drop_frame
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

        # 解析拖入的文件路径（处理含空格的路径）
        path = event.data.strip()
        # tkinterdnd2 可能用 {} 包裹含空格的路径
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]

        if not os.path.isfile(path):
            # 可能拖入了多个文件，取第一个
            parts = event.data.split()
            for p in parts:
                p = p.strip('{}')
                if os.path.isfile(p) and p.lower().endswith('.xlsx'):
                    path = p
                    break
            else:
                messagebox.showwarning("提示", "请拖入 .xlsx 文件")
                return

        if not path.lower().endswith('.xlsx'):
            messagebox.showwarning("提示", "仅支持 .xlsx 文件")
            return

        self.input_path.set(path)
        # 自动生成输出路径
        base, ext = os.path.splitext(path)
        self.output_path.set(f"{base}_converted{ext}")
        # 自动开始转换
        self.root.after(100, self._start_convert)

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择 WPS XLSX 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.input_path.set(path)
            # 自动生成输出路径
            base, ext = os.path.splitext(path)
            self.output_path.set(f"{base}_converted{ext}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if path:
            self.output_path.set(path)

    def _log(self, msg):
        """线程安全地写入日志"""
        def _append():
            self.log_area.configure(state=tk.NORMAL)
            self.log_area.insert(tk.END, msg + "\n")
            self.log_area.see(tk.END)
            self.log_area.configure(state=tk.DISABLED)
        self.root.after(0, _append)

    def _auto_convert(self):
        """命令行传入文件时自动生成输出路径并转换"""
        inp = self.input_path.get()
        if inp:
            base, ext = os.path.splitext(inp)
            self.output_path.set(f"{base}_converted{ext}")
            self._start_convert()

    def _start_convert(self):
        if self.converting:
            return

        inp = self.input_path.get().strip()
        out = self.output_path.get().strip()

        if not inp:
            messagebox.showwarning("提示", "请先选择输入文件")
            return
        if not os.path.isfile(inp):
            messagebox.showerror("错误", f"输入文件不存在:\n{inp}")
            return
        if not inp.lower().endswith('.xlsx'):
            messagebox.showwarning("提示", "仅支持 .xlsx 文件")
            return
        if not out:
            messagebox.showwarning("提示", "请指定输出文件路径")
            return

        # 确保输出目录存在
        out_dir = os.path.dirname(out)
        if out_dir and not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录:\n{e}")
                return

        self.converting = True
        self.convert_btn.configure(state=tk.DISABLED, text="转换中...", bg="#9E9E9E")

        # 清空日志
        self.log_area.configure(state=tk.NORMAL)
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state=tk.DISABLED)

        def run():
            try:
                success, msg, count = wps_image_converter(inp, out, log_callback=self._log)
                self.root.after(0, lambda: self._convert_done(success, msg, count, out))
            except Exception as e:
                self.root.after(0, lambda: self._convert_done(False, str(e), 0, out))

        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def _convert_done(self, success, msg, count, output_path):
        self.converting = False
        self.convert_btn.configure(state=tk.NORMAL, text="开始转换", bg="#4CAF50")

        if success:
            result = messagebox.askyesno(
                "转换完成",
                f"成功修复 {count} 张图片！\n\n文件已保存到:\n{output_path}\n\n是否打开文件所在目录？"
            )
            if result:
                self._open_folder(output_path)
        else:
            messagebox.showerror("转换失败", msg)

    def _open_folder(self, filepath):
        """打开文件所在的目录"""
        folder = os.path.dirname(os.path.abspath(filepath))
        if sys.platform == "darwin":
            os.system(f'open "{folder}"')
        elif sys.platform == "win32":
            os.startfile(folder)
        else:
            os.system(f'xdg-open "{folder}"')


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = FKApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
