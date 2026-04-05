# RTF Tools v0.3 User Guide

> Prerequisites: Windows with Microsoft Word installed. **Close all open Word documents** before running any function.

---

## 1. RTF Page Check

Validates page-count consistency across all RTF files in a folder (including subfolders).

**How it works:** Compares the actual page count rendered by Word (via COM) against the `Page 1 of N` marker parsed from the RTF text. Mismatches trigger a deep section-level scan to locate the discrepancy.

**Steps:**
**Close all open Word documents first!**
1. Select the target folder (browse or type the path manually).
2. Click **Start to Check!**.
3. The program checks files in parallel (one Word instance per CPU thread). Results appear in the log area.

---

## 2. RTF Combine (Specify)

Merges multiple RTF files into one, with optional Table of Contents and global page-number refresh.

**Steps:**
1. **Add Folder** → check the files you need → **Confirm >** to push them to the merge queue on the right.
2. Use **Move Up / Move Down** to reorder.
3. Options:
   - **Add TOC** – insert a hyperlinked Table of Contents at the beginning.
   - **Rows of Page** – TOC rows per page (default 23).
   - **Refresh Page Number** – rewrite all `Page X of Y` footers to reflect the merged document.
4. Set output path and filename, then click **Combine Now!**.

### RTF input standard
To ensure the program correctly merges RTF files and generates a Table of Contents (TOC), inputted RTF outputs need to follow requirements:
1. Title and TOC Format
An IDX bookmark (e.g., IDX1) must be present. The title text must strictly follow this exact format: \s999 \b [Your Title Text] \b0. Missing the style name or bold tags will result in extraction failure.
2. Pagination Format
The RTF syntax for pagination must support [of] or [/] as page connectors (e.g., Page 1 of 5 or Page 1 / 5). The first page of every file must explicitly contain Page 1 and its total page count.
3. Document Boundary Control Tags
The RTF document must contain at least one of the following tags: \widowctrl, \sectd, or \info. The program relies on these tags to properly split the Header from the Body.
4. Page Dimensions
The header of the first inputted RTF file must include the \pgwsxn (width) and \pghsxn (height) tags. These are used to determine and set the dimensions for the dynamically generated TOC page.
---

## 3. Docx/RTF Combine (General)

Merges multiple DOCX / RTF files into a single DOCX via Word COM. Does not support TOC insertion.

**Steps:**
**Close all open Word documents first!**
1. Add and reorder files the same way as Tab 2 (accepts `.docx` and `.rtf`).
2. Set output path and filename, then click **Combine Docx Now!**.

---

## 4. RTF Converter

### RTF → PDF / DOCX
1. Check the desired output format(s).
2. Select the source RTF file.
3. Click **Run Conversion**. PDFs are automatically optimized (bookmark expansion + Fast Web View).

### DOCX → RTF
**Close all open Word documents first!**
1. Select a single `.docx` file **or** a folder for batch conversion.
2. Click **Run Conversion (Docx to RTF)**. Existing RTF files with the same name are skipped.

---

## Notes

- All paths and settings are saved automatically on exit and restored on next launch.
- If the program hangs, manually kill `WINWORD.EXE` in Task Manager and retry.


---
---


# RTF Tools v0.3 使用说明

> 前置条件：Windows 系统，已安装 Microsoft Word。运行任何功能前请先**关闭所有已打开的 Word 文档**。

---

## 一、RTF Page Check（RTF 页码校验）

校验指定文件夹（含子文件夹）内所有 RTF 文件的页码一致性。

**原理：** 通过 Word COM 获取实际渲染页数，同时从 RTF 文本中解析 `Page 1 [of] or [/] N` 类页码标记，对比两者是否一致。不一致时自动定位到出错的节。

**步骤：**
1. 请确保运行之前关闭任何Word文档！
2. 选择包含 RTF 文件的文件夹（支持手动输入路径或浏览选择）。
3. 点击 **Start to Check!**。
4. 程序按 CPU 线程数并行检查，结果输出到日志区。

---

## 二、RTF Combine - Specify（RTF 合并 - 指定样式）

将多个 RTF 文件合并为一个，支持自动生成目录（TOC）和刷新全局页码。

**步骤：**
1. **Add Folder** 添加文件夹，勾选需要的文件后点击 **Confirm >** 提交到右侧队列。
2. 右侧队列支持 **Move Up / Move Down** 调整顺序。
3. 配置选项：
   - **Add TOC**：是否在合并文件头部插入目录。
   - **Rows of Page**：目录每页行数（默认 23）。
   - **Refresh Page Number**：根据合并后的实际页码重写所有 `Page X of Y`。
4. 设置输出路径和文件名，点击 **Combine Now!**。

### RTF 输入规范
为确保 Golang 程序能正确合并 RTF 并生成目录，上游（SAS/Python）输出的 RTF 必须严格满足以下要求：
1. 标题与目录格式
   必须存在 IDX（如 IDX1）书签标记。标题文本必须严格采用该格式：\s999 \b [你的标题文本] \b0。缺失样式名或加粗标记会导致提取失败。
2. 页码格式
   页码RTF语法必须支持 of 或 / 作为页码连接符（如 Page 1 of 5 或 Page 1 / 5），每份文件的第一页必须包含 Page 1 及其总页数，用于统计单表总页数。
3. 文档边界控制符
   RTF 文档必须包含 \widowctrl、\sectd 或 \info 之一。程序依赖它们来切割 Header（头部）和 Body（正文）。
4. 页面尺寸
   传入的第一个 RTF 文件头部必须包含 \pgwsxn (宽) 和 \pghsxn (高) 标签，用于设定生成目录页的尺寸。

---

## 三、Docx/RTF Combine - General（Docx/RTF 通用合并）

通过 Word COM 将多个 DOCX 或 RTF 文件按顺序合并为一个 DOCX 文件。不支持 TOC。

**步骤：**
1. 请确保运行之前关闭任何Word文档！
2. 与 Tab 2 相同方式添加和排序文件（支持 `.docx` 和 `.rtf`）。
3. 设置输出路径和文件名，点击 **Combine Docx Now!**。

---

## 四、RTF Converter（格式转换）

提供两个方向的转换：

### RTF → PDF / DOCX
1. 勾选目标格式（PDF / DOCX，可同时勾选）。
2. 选择源 RTF 文件。
3. 点击 **Run Conversion**。PDF 会自动优化（展开书签 + Fast Web View）。

### DOCX → RTF
1. 请确保运行之前关闭任何Word文档！
2. 选择单个 `.docx` 文件或包含多个 `.docx` 的文件夹（批量模式）。
3. 点击 **Run Conversion (Docx to RTF)**。已存在同名 RTF 的文件会跳过。

---

## 其他说明

- 所有填写过的路径和配置会在关闭窗口时自动保存，下次启动自动恢复。
- 如程序卡死，可手动在任务管理器中结束 `WINWORD.EXE` 后重试。
