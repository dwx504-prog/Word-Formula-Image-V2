# -*- coding: utf-8 -*-
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

def convert_doc(word, source, output_dir):
    source = Path(source).resolve()
    doc = word.Documents.Open(str(source), ReadOnly=False, AddToRecentFiles=False,
                              ConfirmConversions=False)
    total = int(doc.OMaths.Count)
    converted, errors = 0, []
    try:
        # 从文档末尾向前处理，减少替换对前面公式索引的影响
        for n in range(total, 0, -1):
            try:
                r = doc.OMaths.Item(n).Range
                start, end = int(r.Start), int(r.End)
                length = end - start
                if length <= 0:
                    errors.append(f"公式{n}：范围为空，跳过")
                    continue

                # 复制公式为图片，在原公式前粘贴 EMF
                r.CopyAsPicture()
                before = int(doc.InlineShapes.Count)
                doc.Range(start, start).PasteSpecial(DataType=WD_PASTE_ENHANCED_METAFILE)
                after = int(doc.InlineShapes.Count)

                # 未确认插入图片时绝不删除原公式
                if after <= before:
                    errors.append(f"公式{n}：未检测到新图片，原公式保留")
                    continue

                # 找到刚插入在原公式位置附近的行内图片
                shape, distance = None, None
                for j in range(1, after + 1):
                    candidate = doc.InlineShapes.Item(j)
                    d = abs(int(candidate.Range.Start) - start)
                    if distance is None or d < distance:
                        shape, distance = candidate, d
                    if d == 0:
                        break

                if shape is None or distance is None or distance > 2:
                    errors.append(f"公式{n}：图片位置无法确认，原公式保留，请检查输出")
                    continue

                # 图片占据的范围位于原公式之前；原公式随之向后移动
                shifted_start = int(shape.Range.End)
                shifted_end = shifted_start + length
                if shifted_end > int(doc.Content.End):
                    errors.append(f"公式{n}：无法安全定位原公式，原公式保留")
                    continue

                doc.Range(shifted_start, shifted_end).Delete()
                converted += 1

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / f"{source.stem}_公式转图片_V2.docx"
        k = 2
        while target.exists():
            target = output_dir / f"{source.stem}_公式转图片_V2_{k}.docx"
            k += 1
        doc.SaveAs2(str(target), FileFormat=WD_FORMAT_DOCX)
        return {"name": source.name, "total": total, "converted": converted,
                "remaining": int(doc.OMaths.Count), "target": str(target), "errors": errors}
    finally:
        doc.Close(SaveChanges=False)

class App:
    def __init__(self, root):
        self.root = root
        root.title("Word公式图片转换助手 V2")
        root.geometry("740x530")
        self.files = []
        self.out = tk.StringVar(value=str(Path.home() / "Documents" / "公式图片转换结果"))
        self.status = tk.StringVar(value="请选择 DOCX 文件。")
        f = ttk.Frame(root, padding=16); f.pack(fill="both", expand=True)
        ttk.Label(f, text="Word公式图片转换助手 V2",
                  font=("Microsoft YaHei UI", 17, "bold")).pack(anchor="w")
        ttk.Label(f, text="将 Word 原生公式转换为 EMF 图片；确认图片插入后再删除原公式；原文件不会被覆盖。",
                  wraplength=690).pack(anchor="w", pady=8)
        bar = ttk.Frame(f); bar.pack(fill="x")
        ttk.Button(bar, text="添加 Word 文件", command=self.add).pack(side="left")
        ttk.Button(bar, text="添加文件夹", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(bar, text="移除选中", command=self.remove).pack(side="left")
        ttk.Button(bar, text="清空列表", command=self.clear).pack(side="left", padx=6)
        self.lb = tk.Listbox(f, height=10); self.lb.pack(fill="both", expand=True, pady=8)
        row = ttk.Frame(f); row.pack(fill="x")
        ttk.Label(row, text="输出目录：").pack(side="left")
        ttk.Entry(row, textvariable=self.out).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="浏览", command=self.browse).pack(side="left")
        self.pb = ttk.Progressbar(f, mode="determinate"); self.pb.pack(fill="x", pady=8)
        ttk.Label(f, textvariable=self.status, wraplength=690).pack(anchor="w")
        row2 = ttk.Frame(f); row2.pack(fill="x", pady=10)
        ttk.Label(row2, text="需要 Windows + Microsoft Word 桌面版。主要支持 Word 原生公式（OMath）；不自动识别 MathType、普通文本公式或公式截图。",
                  wraplength=510).pack(side="left", fill="x", expand=True)
        self.btn = ttk.Button(row2, text="开始转换", command=self.start); self.btn.pack(side="right")

    def add(self):
        for p in filedialog.askopenfilenames(filetypes=[("Word 文档", "*.docx")]):
            if p not in self.files:
                self.files.append(p); self.lb.insert("end", p)
        self.status.set(f"已添加 {len(self.files)} 个文件。")

    def add_folder(self):
        d = filedialog.askdirectory()
        if d:
            for p in Path(d).rglob("*.docx"):
                if not p.name.startswith("~$") and str(p) not in self.files:
                    self.files.append(str(p)); self.lb.insert("end", str(p))
            self.status.set(f"已添加 {len(self.files)} 个文件。")

    def remove(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i); del self.files[i]

    def clear(self):
        self.files.clear(); self.lb.delete(0, "end")
        self.status.set("列表已清空。")

    def browse(self):
        d = filedialog.askdirectory()
        if d: self.out.set(d)

    def start(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加 DOCX 文件。"); return
        if pythoncom is None or win32com is None:
            messagebox.showerror("组件缺失", "缺少 Word 自动化组件，请重新编译 V2。"); return
        self.btn.config(state="disabled")
        self.pb["maximum"] = len(self.files); self.pb["value"] = 0
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):
        pythoncom.CoInitialize()
        word, results, failures = None, [], []
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for i, p in enumerate(list(self.files), 1):
                self.root.after(0, lambda i=i, p=p: self.status.set(f"正在处理 {i}/{len(self.files)}：{Path(p).name}"))
                try:
                    results.append(convert_doc(word, p, self.out.get()))
                except Exception as e:
                    failures.append(f"{p}\n错误：{e}")
                self.root.after(0, lambda i=i: self.pb.configure(value=i))
        except Exception as e:
            failures.append(f"无法启动 Microsoft Word：{e}")
        finally:
            if word is not None:
                try: word.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.finish(results, failures))

    def finish(self, results, failures):
        self.btn.config(state="normal")
        self.status.set(f"处理结束：{len(results)} 个文档完成，{len(failures)} 个文件错误。")
        lines = []
        for r in results:
            lines.append(f"{r['name']}\n识别公式：{r['total']}；转换成功：{r['converted']}；剩余公式对象：{r['remaining']}\n输出：{r['target']}")
            if r["errors"]: lines.append("公式异常：\n" + "\n".join(r["errors"][:8]))
        if failures: lines.append("文件错误：\n" + "\n\n".join(failures))
        messagebox.showinfo("转换结果", "\n\n".join(lines)[:8500] or "没有生成结果。")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        if "vista" in ttk.Style().theme_names(): ttk.Style().theme_use("vista")
    except Exception: pass
    App(root)
    root.mainloop()
