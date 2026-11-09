# -*- coding: utf-8 -*-
"""Word 原生公式转 EMF 图片工具。需要 Windows 和桌面版 Microsoft Word。"""
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

WD_PASTE_ENHANCED_METAFILE = 9
WD_FORMAT_DOCX = 12


def convert_doc(word, source_path, output_dir):
    source = Path(source_path).resolve()
    if source.name.startswith("~$"):
        raise RuntimeError("这是 Word 临时锁定文件，请关闭对应文档后再处理。")

    doc = word.Documents.Open(
        str(source),
        ReadOnly=False,
        AddToRecentFiles=False,
        ConfirmConversions=False,
    )
    results = []
    try:
        formula_count = int(doc.OMaths.Count)
        converted = 0
        errors = []

        # 从后往前处理，降低前面公式索引受文档编辑影响的风险。
        for formula_index in range(formula_count, 0, -1):
            try:
                formula_range = doc.OMaths.Item(formula_index).Range
                original_start = int(formula_range.Start)
                original_end = int(formula_range.End)
                original_length = original_end - original_start

                if original_length <= 0:
                    errors.append(f"公式 {formula_index}：公式范围为空，跳过。")
                    continue

                # 先复制为图片，再在原公式前插入 EMF 图片。
                formula_range.CopyAsPicture()
                inline_count_before = int(doc.InlineShapes.Count)

                insertion_range = doc.Range(original_start, original_start)
                insertion_range.PasteSpecial(DataType=WD_PASTE_ENHANCED_METAFILE)

                inline_count_after = int(doc.InlineShapes.Count)
                if inline_count_after <= inline_count_before:
                    errors.append(
                        f"公式 {formula_index}：未检测到新插入的图片，原公式保留。"
                    )
                    continue

                # 查找最接近原公式起始位置的行内图片。
                inserted_shape = None
                closest_distance = None
                for shape_index in range(1, inline_count_after + 1):
                    candidate = doc.InlineShapes.Item(shape_index)
                    distance = abs(int(candidate.Range.Start) - original_start)
                    if closest_distance is None or distance < closest_distance:
                        inserted_shape = candidate
                        closest_distance = distance
                    if distance == 0:
                        break

                if (
                    inserted_shape is None
                    or closest_distance is None
                    or closest_distance > 2
                ):
                    errors.append(
                        f"公式 {formula_index}：图片位置无法确认，原公式保留，请人工检查。"
                    )
                    continue

                # 图片插入在原公式之前，原公式随之向后移动。
                shifted_formula_start = int(inserted_shape.Range.End)
                shifted_formula_end = shifted_formula_start + original_length

                if shifted_formula_end > int(doc.Content.End):
                    errors.append(
                        f"公式 {formula_index}：无法安全定位原公式，原公式保留。"
                    )
                    continue

                # 只有确认图片已插入且原公式范围可定位时，才删除原公式。
                doc.Range(shifted_formula_start, shifted_formula_end).Delete()
                converted += 1

            except Exception as exc:
                errors.append(f"公式 {formula_index}：处理失败：{exc}")

        output = Path(output_dir).expanduser()
        output.mkdir(parents=True, exist_ok=True)
        target = output / f"{source.stem}_公式转图片_V2_1.docx"
        suffix = 2
        while target.exists():
            target = output / f"{source.stem}_公式转图片_V2_1_{suffix}.docx"
            suffix += 1

        doc.SaveAs2(str(target), FileFormat=WD_FORMAT_DOCX)
        return {
            "source": str(source),
            "target": str(target),
            "total": formula_count,
            "converted": converted,
            "remaining": int(doc.OMaths.Count),
            "errors": errors,
        }
    finally:
        # 即使转换中途报错，也关闭文档；源文件本身不会被覆盖。
        doc.Close(SaveChanges=False)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word公式图片转换助手 V2.1")
        self.root.geometry("760x550")
        self.root.minsize(700, 480)

        self.files = []
        self.output_dir = tk.StringVar(
            value=str(Path.home() / "Documents" / "公式图片转换结果")
        )
        self.status = tk.StringVar(value="请选择需要转换的 DOCX 文档。")

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Word公式图片转换助手 V2.1",
            font=("Microsoft YaHei UI", 17, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text="将 Word 原生公式转换为 EMF 图片。确认图片插入后才删除原公式，转换结果另存为新文件。",
            wraplength=710,
        ).pack(anchor="w", pady=(6, 12))

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="添加 Word 文件", command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="添加文件夹", command=self.add_folder).pack(
            side="left", padx=7
        )
        ttk.Button(toolbar, text="移除选中", command=self.remove_selected).pack(side="left")
        ttk.Button(toolbar, text="清空列表", command=self.clear_files).pack(
            side="left", padx=7
        )

        self.listbox = tk.Listbox(frame, height=11)
        self.listbox.pack(fill="both", expand=True, pady=9)

        output_row = ttk.Frame(frame)
        output_row.pack(fill="x")
        ttk.Label(output_row, text="输出目录：").pack(side="left")
        ttk.Entry(output_row, textvariable=self.output_dir).pack(
            side="left", fill="x", expand=True, padx=7
        )
        ttk.Button(output_row, text="浏览", command=self.choose_output).pack(side="left")

        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.pack(fill="x", pady=(10, 6))
        ttk.Label(frame, textvariable=self.status, wraplength=710).pack(anchor="w")

        bottom = ttk.Frame(frame)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Label(
            bottom,
            text="运行要求：Windows + Microsoft Word 桌面版。主要支持 Word 原生公式（OMath）；MathType、普通文本公式和公式截图不保证支持。",
            wraplength=535,
        ).pack(side="left", fill="x", expand=True)
        self.start_button = ttk.Button(
            bottom, text="开始转换", command=self.start_conversion
        )
        self.start_button.pack(side="right")

    def add_files(self):
        selected = filedialog.askopenfilenames(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx")],
        )
        self._add_paths(selected)

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择 Word 文档文件夹")
        if not folder:
            return
        paths = [
            str(path)
            for path in Path(folder).rglob("*.docx")
            if not path.name.startswith("~$")
        ]
        self._add_paths(paths)

    def _add_paths(self, paths):
        for path in paths:
            path = str(path)
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert("end", path)
        self.status.set(f"当前已添加 {len(self.files)} 个文件。")

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]
        self.status.set(f"当前已添加 {len(self.files)} 个文件。")

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.status.set("文件列表已清空。")

    def choose_output(self):
        folder = filedialog.askdirectory(title="选择结果保存目录")
        if folder:
            self.output_dir.set(folder)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加至少一个 DOCX 文件。")
            return
        if pythoncom is None or win32com is None:
            messagebox.showerror(
                "组件缺失",
                "未检测到 pywin32 组件。请使用 GitHub Actions 重新编译本项目。",
            )
            return

        self.start_button.config(state="disabled")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        pythoncom.CoInitialize()
        word = None
        results = []
        failures = []
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.ScreenUpdating = False

            for index, filename in enumerate(list(self.files), start=1):
                self.root.after(
                    0,
                    lambda i=index, f=filename: self.status.set(
                        f"正在处理（{i}/{len(self.files)}）：{Path(f).name}"
                    ),
                )
                try:
                    results.append(convert_doc(word, filename, self.output_dir.get()))
                except Exception as exc:
                    failures.append(f"{filename}\n错误：{exc}")
                self.root.after(
                    0, lambda value=index: self.progress.configure(value=value)
                )
        except Exception as exc:
            failures.append(
                "无法启动 Microsoft Word 自动化。请确认已安装桌面版 Word，"
                "并检查是否有 Word 对话框等待处理。\n"
                f"详细信息：{exc}"
            )
        finally:
            if word is not None:
                try:
                    word.ScreenUpdating = True
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self._finish(results, failures))

    def _finish(self, results, failures):
        self.start_button.config(state="normal")
        self.status.set(
            f"处理结束：完成 {len(results)} 个文档，文件级错误 {len(failures)} 个。"
        )
        messages = []
        for item in results:
            messages.append(
                f"{Path(item['source']).name}\n"
                f"识别公式：{item['total']}；转换成功：{item['converted']}；"
                f"剩余公式对象：{item['remaining']}\n"
                f"输出文件：{item['target']}"
            )
            if item["errors"]:
                messages.append("公式级异常：\n" + "\n".join(item["errors"][:10]))
        if failures:
            messages.append("文件级错误：\n" + "\n\n".join(failures))
        messagebox.showinfo("转换结果", "\n\n".join(messages)[:9000] or "没有生成结果。")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
