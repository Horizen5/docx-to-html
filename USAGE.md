# docx2html 使用指南

把 `.docx`（Word/WPS 简历模板）**离线**转换成独立 HTML，保留表格、矢量形状、填充、分割线、浮动照片、旋转/翻转与真实制表位，**不联网、不调用 AI、零第三方依赖**。

> 转换器本体仅用 Python 标准库，开箱即用。

---

## 一、环境要求

| 项目 | 要求 |
| --- | --- |
| Python | 3.6 及以上（仅用标准库，无需 pip install 任何包） |
| 操作系统 | Windows / macOS / Linux 均可 |
| 浏览器 | 用于打开/预览/打印生成的 HTML（推荐 Chrome / Edge） |

无需安装任何第三方库，拿到 `docx2html.py` 即可直接运行。

---

## 二、快速开始

### 方式 1：命令行单文件转换（最常用）

```bash
python docx2html.py 简历.docx
```

- 默认在 `.docx` 同目录生成同名的 `.html`（如 `简历.html`）。
- 想指定输出路径：

```bash
python docx2html.py 简历.docx -o 输出.html
```

### 方式 2：Web 上传界面（图形化批量转换 + 预览，适合不熟命令行）

```bash
python docx2html.py --web --port 8766
```

- 启动后浏览器打开 `http://127.0.0.1:8766`。
- ⚠️ **务必用 `127.0.0.1` 而不是 `localhost`**：Windows 双栈下 `localhost` 会解析到 IPv6 `::1`，连接超时重试 IPv4，每次请求会慢约 2 秒；用 `127.0.0.1` 直接走 IPv4，毫秒级响应。
- 拖拽 / 多选 `.docx` 加入队列，自动逐个转换，每个文件显示「排队中 / 转换中 / 已完成 / 失败」状态徽章。
- 点文件卡片上的「👁 眼睛」按钮以居中弹窗预览转换结果；点「⬇ 下载」单独存盘；底部「🗜 打包下载全部」用 JSZip 一次性导出 zip。
- 还支持 🌙 暗色模式切换、失败「↻ 重试」、删除单文件、复制单文件 HTML 源码。
- 想换端口：`python docx2html.py --web --port 9000`

### 方式 3：批量转换（命令行循环）

把一个文件夹里所有 `.docx` 一次性转成 HTML，输出到 `out/` 目录：

**Windows (PowerShell)：**
```powershell
New-Item -ItemType Directory -Force out | Out-Null
Get-ChildItem -Path . -Filter *.docx | ForEach-Object {
    python docx2html.py $_.FullName -o "out/$($_.BaseName).html"
}
```

**macOS / Linux (bash)：**
```bash
mkdir -p out
for f in *.docx; do
    python docx2html.py "$f" -o "out/${f%.docx}.html"
done
```

---

## 三、命令行参数详解

```
python docx2html.py [input.docx] [选项]
```

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `input` | 待转换的 `.docx` 文件路径（位置参数，可省略） | `简历.docx` |
| `-o, --output` | 输出 HTML 路径；省略则输出到同目录同名 `.html` | `-o result.html` |
| `--web` | 启动 Web 上传界面（与 `input` 二选一） | `--web` |
| `--port` | Web 服务端口，默认 8765 | `--port 9000` |
| `--debug` | 在生成的 HTML 上叠加页面尺寸/页边距调试水印（开发核对用，默认关闭） | `--debug` |

> 不带任何参数运行 `python docx2html.py` 会打印帮助信息。

---

## 四、输出说明

生成的 HTML 是**单文件独立网页**：

- **内联样式**：所有 CSS 写在 `<style>` 里，不依赖外部文件。
- **内联图片**：文档中的图片以 base64 嵌入，单个 `.html` 即可打开，无需附带素材。
- **1:1 还原**：按文档真实页面尺寸（A4 等）、页边距、字号、行距还原版式。

直接用浏览器打开即可预览。后续可在浏览器里：

1. **填内容**：用任意编辑器（或浏览器开发者工具）修改 HTML 文本；
2. **打印成 PDF**：浏览器 `Ctrl+P` → 目标选「另存为 PDF」→ 边距选「无」→ 保存。矢量输出，清晰不糊。

---

## 五、目录结构

整理后的项目目录如下：

```
docx/
├── docx2html.py                  # 核心转换脚本（唯一源码，转换器本体）
├── README.md                     # 项目简介
├── USAGE.md                      # 本文件（详细使用指南）
├── LICENSE                       # MIT 许可证
├── .gitignore
│
├── converted_html/               # 转换输出示例（8 份简历 HTML + 对应 PDF）
├── PDF/                          # 8 份 PDF 真值（用于开发期对比验证）
├── PDF-300/                      # 300 套批量 PDF（开发期数据集）
├── A→【推荐】300+套单页简历/      # 简历模板源（.docx）
├── A→【推荐】300+套单页简历.zip   # 模板压缩包
│
└── archive/                      # 开发期调试产物暂存（非转换器必需）
    ├── debug_scripts/            # 调试脚本（75 个 _*.py / measure*.py 等）
    ├── debug_images/             # 调试截图（53 张对比/渲染 PNG）
    ├── debug_html/               # 中间产物 HTML（25 个）
    ├── debug_temp/               # 临时文本 + 缓存目录
    ├── inspect_dirs/             # 解压检查目录（docx XML 结构分析）
    └── misc/                     # 杂项（单文件 docx、渲染 PDF）
```

> **日常使用只需 `docx2html.py` 一个文件。** `archive/` 是开发调试过程的历史产物，已全部暂存归档，不影响转换器运行，可随时删除或忽略。

---

## 六、常见问题

### Q1：转换后打开 HTML 是一片空白 / 样式丢失？
确保用**现代浏览器**（Chrome / Edge / Firefox）打开。生成的 HTML 是标准网页，不支持 IE。

### Q2：打印成 PDF 时内容超出了一页？
在浏览器打印设置里把**边距选「无」或「默认」**，并关闭「页眉页脚」。转换器已按原文档页面尺寸排版，打印时不要再加额外边距。

### Q3：某些特殊模板转换后版式有偏差？
本转换器面向**单页中文简历模板族**校准。风格差异极大的模板（多栏报刊、复杂图表）可能需要微调。代码顶部有集中常量（默认段后距、小节条段后距等）可调。

### Q4：`--debug` 水印是什么？
开启后会在 HTML 上叠加页面宽高、页边距等参考线，**仅供开发期核对版式**用。正常使用不要加此参数。

### Q5：需要安装 Pillow / PyMuPDF 吗？
**不需要。** 转换器本体零依赖。Pillow / PyMuPDF / 无头浏览器只用于 `archive/` 里的开发期校准脚本，不属于转换流程。

### Q6：如何把转换好的 HTML 改成自己的简历？
1. 用浏览器打开 HTML 确认版式无误；
2. 用 VS Code 等编辑器打开 HTML，直接改文字内容（结构已语义化）；
3. 保存后在浏览器 `Ctrl+P` 导出 PDF。

---

## 七、一行速查

```bash
# 最快上手：转一个 docx
python docx2html.py 我的简历.docx

# 指定输出
python docx2html.py 我的简历.docx -o 简历.html

# 图形界面
python docx2html.py --web
```

---

*许可证：MIT。详见 [LICENSE](LICENSE)。*
