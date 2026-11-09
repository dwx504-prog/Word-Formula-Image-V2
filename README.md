# Word公式图片转换助手 V2.2

本版针对 V2.1 的“识别到公式但没有检测到图片”问题调整了 Word 自动化流程：使用 Word 当前选区复制公式为图片，优先粘贴增强型图元文件（EMF），并检查行内图片和浮动图片；只有确认图片存在且对应原公式对象数量减少后才计为成功。

## 重新上传
将解压后的项目文件上传到 GitHub 仓库，覆盖旧版本文件，保持 `.github/workflows/build.yml` 目录结构。

## 编译
1. 打开仓库 Actions。
2. 选择 `Build Windows EXE V2.2`。
3. 点击 `Run workflow`。
4. 等待语法检查和 Build EXE 均通过。
5. 在本次运行 Summary 页面下载 Artifacts：`Word公式图片转换助手-V2.2-Windows`。

## 使用限制
需要 Windows 10/11 和 Microsoft Word 桌面版。主要支持 DOCX 中的 Word 原生公式（OMath）。MathType、普通文本公式、公式截图及其他嵌入对象不保证支持。程序另存为新 DOCX，不覆盖源文件。建议先在文档副本上测试，并人工检查转换后的公式排版。

## 关于本次上传的测试文档
已检查 `01_权利要求书_系统修订版.docx` 的 DOCX 内部结构，确认其中有 41 个 OMML 公式对象，且主文档 XML 中没有现存图片或嵌入对象。该文件适合作为本软件的首个实测样本。
