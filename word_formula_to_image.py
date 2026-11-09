# -*- coding: utf-8 -*-
"""Word 原生公式转 EMF 图片工具 V2.3。Windows + 桌面版 Microsoft Word。"""
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
WD_IN_LINE = 0
WD_FORMAT_DOCX = 12
WD_INLINE_SHAPE_PICTURE = 3
WD_INLINE_SHAPE_LINKED_PICTURE = 4


def _find_inserted_picture(doc, anchor_pos, before_inline_count, tolerance=8):
    """Return a newly created picture near the insertion point, if one exists."""
    candidates = []
    count = int(doc.InlineShapes.Count)
    for i in range(1, count + 1):
        try:
            shape = doc.InlineShapes.Item(i)
            shape_type = int(shape.Type)
            start = int(shape.Range.Start)
            if shape_type in (WD_INLINE_SHAPE_PICTURE, WD_INLINE_SHAPE_LINKED_PICTURE):
                if i > before_inline_count or abs(start - anchor_pos) <= tolerance:
                    candidates.append((abs(start - anchor_pos), shape))
        except Exception:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    if candidates[0][0] > tolerance and int(doc.InlineShapes.Count) <= before_inline_count:
        return None
    return candidates[0][1]


def _paste_as_emf(word, doc, formula_index, anchor_pos):
    """
    Strategy A mirrors the user's successful manual operation:
    select equation -> Ctrl+C equivalent -> Paste Special as Enhanced Metafile.
    Positional COM arguments are used because named arguments caused E_INVALIDARG
    on the user's Word automation setup.
    """
    errors = []

    # Strategy A: normal Copy, then PasteSpecial(Enhanced Metafile).
    try:
        formula_range = doc.OMaths.Item(formula_index).Range
        formula_range.Select()
        word.Selection.Copy()

        insertion = doc.Range(anchor_pos, anchor_pos)
        insertion.Select()
        before_count = int(doc.InlineShapes.Count)

        # Word Selection.PasteSpecial optional argument order:
        # IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel
        word.Selection.PasteSpecial(0, False, WD_IN_LINE, False,
                                    WD_PASTE_ENHANCED_METAFILE)
        picture = _find_inserted_picture(doc, anchor_pos, before_count)
        if picture is not None:
            return picture, errors
        errors.append("PasteSpecial 未检测到新建的行内图片。")
    except Exception as exc:
        errors.append(f"复制后按增强型图元文件粘贴失败：{exc}")

    # Strategy B: CopyAsPicture already places a picture representation on clipboard;
    # use ordinary Paste() with no format arguments, matching Word's default paste.
    try:
        formula_range = doc.OMaths.Item(formula_index).Range
        formula_range.Select()
        word.Selection.CopyAsPicture()

        insertion = doc.Range(anchor_pos, anchor_pos)
        insertion.Select()
        before_count = int(doc.InlineShapes.Count)
        word.Selection.Paste()
        picture = _find_inserted_picture(doc, anchor_pos, before_count)
        if picture is not None:
            return picture, errors
        errors.append("默认 Paste 未检测到新建的行内图片。")
    except Exception as exc:
        errors.append(f"CopyAsPicture 后默认粘贴失败：{exc}")

    # Strategy C: Copy normal equation content and use default Paste Special UI format
    # fallback through PasteSpecial with only DataType specified positionally.
    try:
        formula_range = doc.OMaths.Item(formula_index).Range
        formula_range.Select()
        word.Selection.Copy()
        insertion = doc.Range(anchor_pos, anchor_pos)
        insertion.Select()
        before_count = int(doc.InlineShapes.Count)
        word.Selection.PasteSpecial(0, False, WD_IN_LINE, False,
                                    WD_PASTE_ENHANCED_METAFILE)
        picture = _find_inserted_picture(doc, anchor_pos, before_count)
        if picture is not None:
            return picture, errors
        errors.append("备用 PasteSpecial 未检测到新建的行内图片。")
    except Exception as exc:
        errors.append(f"备用粘贴失败：{exc}")

    return None, errors


def convert_doc(word, source_path, output_dir):
    source = Path(source_path).resolve()
    if source.name.startswith("~$"):
        raise RuntimeError("这是 Word 临时锁定文件，请关闭对应文档后再处理。")

    doc = word.Documents.Open(
        str(source), ReadOnly=False, AddToRecentFiles=False,
        ConfirmConversions=False,
    )
    try:
        initial_count = int(doc.OMaths.Count)
        converted = 0
        errors = []

        # Reverse order protects the positions/indexes of formulas not yet processed.
        for formula_index in range(initial_count, 0, -1):
            try:
                formula_range = doc.OMaths.Item(formula_index).Range
                anchor_pos = int(formula_range.Start)
                before_omath_count = int(doc.OMaths.Count)
                before_inline_count = int(doc.InlineShapes.Count)

                picture, paste_errors = _paste_as_emf(
                    word, doc, formula_index, anchor_pos
                )
                if picture is None:
                    errors.append(
                        f"公式 {formula_index}：未能确认图片插入，原公式保留。"
                        + "；".join(paste_errors)
                    )
                    continue

                # Verify a newly created picture object is still present.
                try:
                    picture_type = int(picture.Type)
                    picture_start = int(picture.Range.Start)
                    picture_end = int(picture.Range.End)
                    if picture_type not in (
                        WD_INLINE_SHAPE_PICTURE, WD_INLINE_SHAPE_LINKED_PICTURE
                    ) or picture_end <= picture_start:
                        raise RuntimeError("图片对象类型或范围验证失败")
                except Exception as exc:
                    try:
                        picture.Delete()
                    except Exception:
                        pass
                    errors.append(
                        f"公式 {formula_index}：图片对象验证失败，原公式保留。{exc}"
                    )
                    continue

                # Only after a confirmed picture insertion, delete the source equation.
                try:
                    doc.OMaths.Item(formula_index).Range.Delete()
                except Exception as exc:
                    try:
                        picture.Delete()
                    except Exception:
                        pass
                    errors.append(
                        f"公式 {formula_index}：图片已插入但原公式删除失败；已尝试撤销图片。{exc}"
                    )
                    continue

                after_omath_count = int(doc.OMaths.Count)
                if after_omath_count != before_omath_count - 1:
                    # Safety rollback: do not leave a duplicate picture if source equation remains.
                    try:
                        picture.Delete()
                    except Exception:
                        pass
                    errors.append(
                        f"公式 {formula_index}：未确认原公式删除，已尝试撤销图片。"
                    )
                    continue

                converted += 1
            except Exception as exc:
                errors.append(
                    f"公式 {formula_index}：处理异常，原公式尽量保留。原因：{exc}"
                )

        output = Path(output_dir).expanduser()
        output.mkdir(parents=True, exist_ok=True)
        target = output / f"{source.stem}_公式转图片_V2_3.docx"
        suffix = 2
        while target.exists():
            target = output / f"{source.stem}_公式转图片_V2_3_{suffix}.docx"
            suffix += 1

        doc.SaveAs2(str(target), FileFormat=WD_FORMAT_DOCX)
        return {
            "source": str(source), "target": str(target),
            "total": initial_count, "converted": converted,
            "remaining": int(doc.OMaths.Count),
            "pictures_before": before_inline_count,
            "pictures_after": int(doc.InlineShapes.Count),
            "errors": errors,
        }
    finally:
        doc.Close(SaveChanges=False)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word公式图片转换助手 V2.3")
        self.root.geometry("800x590")
        self.root.minsize(720, 500)
        self.files = []
        self.output_dir = tk.StringVar(
            value=str(Path.home() / "Documents" / "公式图片转换结果")
        )
        self.status = tk.StringVar(value="请选择需要转换的 DOCX 文档。")

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame, text="Word公式图片转换助手 V2.3",
            font=("Microsoft YaHei UI", 17, "bold")
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text="V2.3 按照 Word 中手动成功的流程改造：复制公式后优先粘贴为增强型图元文件；同时提供默认图片粘贴备用路径。确认图片对象存在且原公式删除后才计为成功。",
            wraplength=750
        ).pack(anchor="w", pady=(6, 12))

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="添加 Word 文件", command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="添加文件夹", command=self.add_folder).pack(side="left", padx=7)
        ttk.Button(toolbar, text="移除选中", command=self.remove_selected).pack(side="left")
        ttk.Button(toolbar, text="清空列表", command=self.clear_files).pack(side="left", padx=7)

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
        ttk.Label(frame, textvariable=self.status, wraplength=750).pack(anchor="w")

        bottom = ttk.Frame(frame)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Label(
            bottom,
            text="要求：Windows + Microsoft Word 桌面版。适用于 Word 原生公式（OMath）。MathType、普通文本公式和公式截图不保证支持。请先使用文档副本测试。",
            wraplength=570
        ).pack(side="left", fill="x", expand=True)
        self.start_button = ttk.Button(
            bottom, text="开始转换", command=self.start_conversion
        )
        self.start_button.pack(side="right")

    def add_files(self):
        self._add_paths(filedialog.askopenfilenames(
            title="选择 Word 文档", filetypes=[("Word 文档", "*.docx")]
        ))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择 Word 文档文件夹")
        if folder:
            self._add_paths(
                str(x) for x in Path(folder).rglob("*.docx")
                if not x.name.startswith("~$")
            )

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
            messagebox.showerror("组件缺失", "未检测到 pywin32，请重新编译完整项目。")
            return
        self.start_button.config(state="disabled")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        pythoncom.CoInitialize()
        word = None
        results, failures = [], []
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            word.DisplayAlerts = 0
            word.ScreenUpdating = True
            try:
                word.WindowState = 2
            except Exception:
                pass

            for index, filename in enumerate(list(self.files), start=1):
                self.root.after(
                    0,
                    lambda i=index, f=filename: self.status.set(
                        f"正在处理（{i}/{len(self.files)}）：{Path(f).name}"
                    )
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
                "无法启动 Microsoft Word 自动化。请确认已安装桌面版 Word。\n" + str(exc)
            )
        finally:
            if word is not None:
                try:
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
                f"图片数量：处理前 {item['pictures_before']}，处理后 {item['pictures_after']}\n"
                f"输出文件：{item['target']}"
            )
            if item["errors"]:
                messages.append("公式级异常：\n" + "\n".join(item["errors"][:15]))
        if failures:
            messages.append("文件级错误：\n" + "\n\n".join(failures))
        messagebox.showinfo(
            "转换结果", "\n\n".join(messages)[:10000] or "没有生成结果。"
        )


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
