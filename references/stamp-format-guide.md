# 印章/署名/日期格式指南

## 印章（盖章）规范

### 单一机关行文
- **印章位置**: 端正、居中下压成文日期
- **骑车盖月**: 印章的上边缘压在成文日期上（不压正文），下边缘盖在成文日期上
- 印章用红色，不得使用复印件

### 联合行文（多个机关）
- 各机关署名按发文机关顺序排列
- 各印章应加盖在对应的署名之上
- 最后一个印章的下边缘压在成文日期上
- 印章之间互不重叠
- **自动生成**：使用 `joint_issuing_orgs` 字段（list[str]），脚本自动按机关数量分行排列署名和日期
  - 示例：`--joint-sample` 生成双机关联合行文

## 发文机关署名

### 署名格式
- 署名在正文（或附件说明）下空一行
- 使用机关全称或规范化简称
- 与成文日期居中对齐

## 成文日期格式

- 用阿拉伯数字标全年、月、日
- 右空四字
- 示例: `2024年1月15日`（非"贰零贰肆年壹月拾伍日"）

### 日期位置规则
| 情况 | 位置 |
|------|------|
| 加盖印章的公文 | 印章下压日期 |
| 不加盖印章的公文 | 正文下右空四字排印 |

## 附注（联系信息）

- 如需标注联系人和联系电话，在成文日期下一行左空二字
- 用圆括号括入：`（联系人：XXX，电话：XXXXXX）`

## 附件说明格式

```
附件：1. [附件名称一]
       2. [附件名称二]
```

- 正文下空一行，左空二字
- 附件名称后不加标点符号
- 附件名称较长需回行时，与上一行附件名称的首字对齐

## 抄送机关格式

```
抄送：[机关名称一]，[机关名称二]，[机关名称三]。
```

- 版记中首条分隔线之下
- 仿宋_GB2312 三号
- 左空一字
- 回行时与冒号后的首字对齐
- 最后一个抄送机关名称后标句号

## python-docx 实现参考

```python
# 盖章位置: 通过段落对齐实现右空四字
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_signature_and_date(doc, org_name, date_str):
    """发文机关署名+成文日期"""
    # 署名: 右空四字
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.right_indent = Cm(1.5)  # 约4个字符
    run1 = p1.add_run(org_name)
    set_run_font(run1, "仿宋_GB2312", size=Pt(16))
    
    # 成文日期: 右空四字
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.right_indent = Cm(1.5)
    run2 = p2.add_run(date_str)
    set_run_font(run2, "仿宋_GB2312", size=Pt(16))

def add_attachment(doc, attachments):
    """附件说明"""
    p = doc.add_paragraph()
    for i, att in enumerate(attachments, 1):
        if i == 1:
            run = p.add_run(f"附件：{att}")
        else:
            run = p.add_run(f"\n       {att}")
        set_run_font(run, "仿宋_GB2312", size=Pt(16))
```

## 常见错误

| 错误 | 正确 |
|------|------|
| 印章盖在空白处 | 印章必须压住成文日期 |
| 成文日期用汉字数字 | 用阿拉伯数字：2024年1月15日 |
| 发文字号中括号用[] | 用六角括号〔〕 |
| 附件名称后加标点 | 附件名称后不加标点 |
| 署名与日期分左右排列 | 署名与日期居中对齐（右空四字） |
