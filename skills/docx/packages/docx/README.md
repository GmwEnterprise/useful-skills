# docx

读取 Word 文档（`.docx`）为 Markdown，并从本地 Markdown 生成带样式的 Word 文档。

## 安装

```bash
npx skills add https://github.com/GmwEnterprise/useful-skills --skill docx
```

## 读取

```bash
scripts/docx-read document.docx
scripts/docx-read document.docx output
```

## 创建

```bash
scripts/docx-write report.md report.docx
scripts/docx-write report.md report.docx --body-size 12 --h1-size 22
scripts/docx-write report.md report.docx --style-config .tmp/docx-style.json
```

## 创建支持内容

- 标题 1-6
- 普通段落
- 加粗、斜体、删除线、行内代码
- 引用块
- 列表
- 分隔线
- 代码块
- 表格
- 图片

## 默认样式

- A4 页面
- 正文 `Microsoft YaHei` 11pt
- 标题逐级递减字号并加粗
- 代码块浅灰底色
- 表格首行强调
- 图片自动缩放到页面可用宽度
