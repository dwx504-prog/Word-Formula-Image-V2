# Word公式图片转换助手 V2.3

## 本版针对的问题
用户已确认：在 Word 中手动选择公式，按 Ctrl+C，再按 Ctrl+Alt+V，可以选择“图片（增强型图元文件）”并成功粘贴。
V2.2 的 COM 自动化调用 PasteSpecial 时出现 E_INVALIDARG（-2147024809）。

V2.3 调整为：
1. 先执行 Word 公式范围选择与普通 Copy，再使用 Word PasteSpecial 的位置参数传递方式粘贴为增强型图元文件；
2. 若该路径未能生成可识别的行内图片，则回退到 CopyAsPicture + 默认 Paste；
3. 再提供一次备用 PasteSpecial 路径；
4. 检测实际插入的图片对象，并且只在确认图片插入后才删除原公式；
5. 处理结果另存为新 DOCX，不覆盖源文件。

## 上传 GitHub
将 ZIP 解压后的全部文件覆盖仓库中的同名文件，确保：
- word_formula_to_image.py 位于仓库根目录
- requirements.txt 位于仓库根目录
- .github/workflows/build.yml 位于该路径

## 编译
进入 GitHub Actions，选择 “Build Windows EXE V2.3”，点击 Run workflow。
编译完成后在运行 Summary 的 Artifacts 区域下载 “Word公式图片转换助手-V2.3-Windows”。

## 使用
需要 Windows 和 Microsoft Word 桌面版。当前版本面向 Word 原生公式（OMath）。建议先用副本测试，并检查转换后公式图片的显示、位置和排版。
注意：当前运行环境无法直接启动用户本机的 Microsoft Word，因此语法检查通过不等于已完成 Office 实机验证。
