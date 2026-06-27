# 公文模板库

本目录存放 GB/T 9704-2012 公文生成脚本的模板文件（JSON 格式）。
内置模板为**参考样例**，用户可自行创建、替换或扩展。

---

## 模板文件格式

每个模板是一个独立 JSON 文件，字段说明：

| 字段 | 必填 | 说明 |
|------|------|------|
| `template_id` | 是 | 唯一标识符，**必须与文件名一致**（不含 `.json` 后缀） |
| `name` | 是 | 模板显示名称 |
| `description` | 是 | 一句话描述适用场景 |
| `document_type` | 是 | 文种（通知、报告、函、纪要 等） |
| `source_refs` | 是 | 来源依据列表，见下方说明 |
| `defaults` | 否 | 预置 config 字段默认值（如 `main_recipient`） |
| `body_sections` | 是 | 正文段落结构，同 `generate_gongwen()` 的 body_sections 格式 |
| `hints` | 否 | 写作提示信息 |

### source_refs 格式

每条来源记录包含 type 字段，取值：

| type | 含义 | 示例 |
|------|------|------|
| `regulation` | 法规依据（须注明法规名称+条款） | 《党政机关公文处理工作条例》中办发〔2012〕14号 |
| `common_practice` | 常见做法（无单一法规硬性规定） | 通知的惯用正文结构 |
| `example` | 实例参考（参考某具体公文） | XX省2023年度职称评审通知 |

示例：

```json
"source_refs": [
  {
    "type": "regulation",
    "title": "《党政机关公文处理工作条例》",
    "ref": "中办发〔2012〕14号",
    "section": "第八条 通知"
  },
  {
    "type": "common_practice",
    "note": "人事任免类通知惯用结构：'任职—免职'两节"
  }
]
```

### body_sections 格式

```json
{"level": 0, "text": "一级标题，格式为'一、'"}
{"level": 1, "text": "二级标题，格式为'（一）'"}
{"level": null, "text": "正文段落（有首行缩进）"}
```

---

## 使用自定义模板

有三种方式覆盖或扩展内置模板：

### 方式一：--template-dir（推荐）

把自定义模板 JSON 文件放在同一目录中，同名 template_id 将覆盖内置模板：

```bash
python scripts/generate_gongwen_docx.py --template-dir ~/my-gongwen-templates/ --list-templates
```

### 方式二：GONGWEN_TEMPLATE_DIR 环境变量

```bash
export GONGWEN_TEMPLATE_DIR=~/my-gongwen-templates/
python scripts/generate_gongwen_docx.py --list-templates
```

### 方式三：--template-file

直接从任意路径加载一个模板文件查看或修改：

```bash
python scripts/generate_gongwen_docx.py --template-file ~/my-templates/special-notice.json
```

### 搜索优先级

1. 用户自定义 `--template-dir`（同名覆盖内置）
2. 内置 `templates/` 目录
