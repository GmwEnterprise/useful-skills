# docx-reader

读取 Word 文档（.docx）并提取内容为 Markdown 格式，同时导出嵌入的图片和附件。

## 安装

```bash
npx skills add https://github.com/GmwEnterprise/useful-skills --skill docx-reader
```

## 使用

```bash
# 默认输出到 ./docs/extracted/
./skills/docx-reader/scripts/docx-reader document.docx

# 指定输出目录
./skills/docx-reader/scripts/docx-reader document.docx ./output
```

## 输出

```
docs/extracted/
├── document.md              # Markdown 内容
└── document/                # 图片和附件
    ├── image1.png
    └── embedded.xlsx
```

## 功能

- 段落文本（粗体、斜体、删除线、下划线）
- 标题层级（H1-H6）
- 表格转 Markdown 表格
- 列表（有序/无序）
- 文本对齐（居中/右对齐）
- 嵌入图片提取与引用
- 嵌入附件提取（xlsx、pdf 等）
- 文档元数据（标题、作者、时间等）
