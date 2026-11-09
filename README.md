# Word公式图片转换助手 V2 完整项目

## 项目文件
- `word_formula_to_image.py`：软件主体
- `requirements.txt`：编译依赖
- `.github/workflows/build.yml`：GitHub 云端 Windows 编译流程
- `README.md`、`使用说明.txt`：使用说明

## 功能
- 批量选择 DOCX 或添加整个文件夹
- 通过 Windows 桌面版 Microsoft Word 自动化识别 Word 原生公式（OMath）
- 将公式复制为 EMF 图片，在原位置插入
- 检测到图片插入后才尝试删除原公式
- 输出为新 DOCX，不覆盖原文件
- 显示识别数、转换数、剩余公式数及异常

## 环境要求
- Windows 10/11
- 已安装桌面版 Microsoft Word
- 输入文件为 DOCX

当前版本主要支持 Word 原生公式（OMath），不自动识别 MathType OLE、普通文本公式、公式截图或 PDF 中的公式。

## GitHub 云端编译（本机无需安装 Python）
1. 在 GitHub 创建一个空仓库。
2. 将 ZIP 解压，把解压目录里的**所有内容**上传到仓库根目录。
3. 必须保留 `.github/workflows/build.yml` 这个目录结构。
4. 打开仓库顶部 Actions。
5. 左侧点击 `Build Windows EXE V2`。
6. 点击 `Run workflow` 并确认。
7. 成功后进入本次运行的 Summary 页面，滚动到页面底部的 Artifacts。
8. 下载 `Word公式图片转换助手-V2-Windows`，解压后得到 `Word公式图片转换助手_V2.exe`。

## 使用
1. 双击 EXE。
2. 添加一份或多份 DOCX，或选择文件夹。
3. 设置输出目录。
4. 点击开始转换。
5. 在输出目录打开 `_公式转图片_V2.docx` 检查结果。

## 重要说明
- 请先用原文件副本测试，确认图片、公式位置和排版符合预期后再批量处理。
- Office 自动化和剪贴板行为可能受 Word 版本、公式结构、文档保护和其他程序占用剪贴板影响。
- 若某条公式不能确认图片插入位置，程序会尽量保留原公式，并在结果中报告异常；该文件应人工检查。
- 转换时不要同时复制其他内容或操作 Word 剪贴板。
