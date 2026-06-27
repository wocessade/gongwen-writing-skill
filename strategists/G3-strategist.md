# Stage G3: 格式 + 内容审查 [Strategist]

**Gate:** QG3 (BLOCK)
**Goal:** 4 代理并行审查文档
**Needs Composers:** none

## 代理分发

申明: 以下 agent 均为 LLM 编码，通过 evaluate 模块加载对应配置。各 agent 的输出包含结构化检查项 + 严重度标记。

### Agent 1: format_compliance_gongwen (1x)

GB/T 9704 格式自动检查 + 运行 `gongwen_format_check.py`。

```
输入:
  - 生成的 .docx 文件路径
  - GB/T 9704 规范 (references/gb-t9704-2012-quick-ref.md)
  - 印章格式指南 (references/stamp-format-guide.md) — 如涉及

检查项:
  1. 页面设置: A4, 版心尺寸正确
  2. 红头: 发行机关标志字体/颜色/居中，分隔线
  3. 发文字号: 格式 {机关代字}[{年份}]{顺序号}
  4. 标题: 小标宋/黑体, 居中, 2.0 行距
  5. 主送机关: 仿宋, 顶格
  6. 正文: 仿宋三号 (16pt), 1.5 倍行距 (28pt), 首行缩进 2 字符
  7. 一级标题: 黑体三号
  8. 二级标题: 楷体三号
  9. 附件标记: 格式正确
  10. 落款: 发文机关 + 成文日期右空 2 字符
  11. 页码: 标注格式正确

评分: 每项 0-10 分，总分 = 平均分 × 10

输出格式:
  - 每项: [PASS/FAIL] {检查项名称} — {描述} — 得分
  - 总体评分: {0-100}
  - 如 FAIL 项，标注严重度 (Critical/Major/Minor)
```

另外，在 LLM agent 检查之外，同时调用 `evaluate/scripts/gongwen_format_check.py` 进行自动化格式检查。如其返回 FAIL，则视为对应项 FAIL。

### Agent 2: language_formality (3x)

语言规范性审查，3 个独立 agent 并行运行，取中位数评分。

```
输入:
  - 从 .docx 提取的纯文本

检查项:
  1. 语体规范: 是否使用正式公文语体 (无口语化表达)
  2. 专用术语: 党政机关公文用语是否准确
  3. 标点符号: 符合中文标点规范 (顿号/逗号/分号/句号使用正确)
  4. 数字用法: 符合 GB/T 15835 (数字用法)
  5. 人名/地名/机构名: 全称使用，无简称错误
  6. 语气适当: 下行文/上行文/平行文语气符合文种规范

评分: agent 独立 0-10 分每项

聚合: 取 3 个 agent 的中位数作为最终得分
  总分 = 中位数平均分 × 10

输出格式:
  - agent_{1..3} 原始评分
  - 中位数聚合后评分
  - 不一致项目标记（分差 > 2 的项目）
```

### Agent 3: structure_completeness (3x)

结构完整性审查，3 个独立 agent 并行运行，取中位数评分。

```
输入:
  - 从 .docx 提取的纯文本
  - G1 阶段规划的章节结构大纲
  - 文种模板

检查项:
  1. 必备要素完整: 标题、主送、正文、发文机关、成文日期齐全
  2. 章节顺序规范: 符合文种模板顺序
  3. 内容充实度: 每节有实质内容，非空泛表述
  4. 逻辑连贯性: 段落间过渡自然，逻辑清晰
  5. 文种要素匹配: 上行文必备要素(签发人)、下行文必备要素(红头)

评分: agent 独立 0-10 分每项

聚合: 取 3 个 agent 的中位数作为最终得分
  总分 = 中位数平均分 × 10

输出格式:
  - agent_{1..3} 原始评分
  - 中位数聚合后评分
  - 不一致项目标记（分差 > 2 的项目）
```

### Agent 4: attachment_seal_check (1x, scoreless)

附件/印章/日期二元检查，无评分。

```
输入:
  - 生成的 .docx 文件
  - G1 附件规划信息
  - 印章格式指南 (references/stamp-format-guide.md)

检查项:
  1. 附件: 如有附件 → 正文中已标注、附件内容齐全
  2. 印章: 如文种需要 → 套红印章位置正确
  3. 成文日期: 格式正确 (XXXX年XX月XX日)
  4. 发文机关署名: 与 G0 确认的发文机关一致
  5. 联系人与电话: 上行文附注联系人信息完整

输出: PASS / FAIL + 具体说明
  - FAIL 视为门控不通过 (不参与加权总分，但独立门控)
```

## 评分聚合

### 加权方案

| 维度 | 权重 | 说明 |
|------|------|------|
| format_compliance | 0.35 | GB/T 9704 格式合规 |
| language_formality | 0.35 | 语言规范性 (3 agent 中位数) |
| structure_completeness | 0.30 | 结构完整性 (3 agent 中位数) |
| attachment_seal | 无分门控 | 附件/印章/日期二元检查，不参与加权 |

### 计算方法

```
weighted_total = (
    format_compliance.score × 0.35 +
    language_formality.score × 0.35 +
    structure_completeness.score × 0.30
)
```

### 严重度等级

| 等级 | 含义 | 处理要求 |
|------|------|---------|
| Critical | 违反公文规范底线，影响公文效力 | 必须修复 |
| Major | 显著偏离标准或规范 | 必须修复或标注为已知 |
| Minor | 轻微格式/措辞问题 | 建议修复 |

## Gate

### QG3 Gate Checklist

- [ ] format_compliance 评分完成
- [ ] gongwen_format_check.py 运行完成 (如存在)
- [ ] language_formality 3 agent 均完成，中位数已聚合
- [ ] structure_completeness 3 agent 均完成，中位数已聚合
- [ ] attachment_seal_check 完成 (PASS/FAIL)
- [ ] 加权总分已计算
- [ ] 加权总分 >= 70
- [ ] attachment_seal 检查结果为 PASS
- [ ] 无 Critical 级别问题

### Gate Failure Route

```
if 加权总分 < 70:
    → BLOCK — 加权总分不达标
    → 报告各维度得分
    → 自动进入 G4 修订循环
if attachment_seal 检查 FAIL:
    → BLOCK — 附件/印章/日期问题
    → 报告具体 FAIL 项
    → 自动进入 G4 修订循环
if 存在 Critical 级别问题:
    → BLOCK — 存在底线问题
    → 列出所有 Critical 项
    → 自动进入 G4 修订循环
if 加权总分 >= 70 AND attachment_seal PASS AND 无 Critical:
    → PASS — 文档审查通过
    → 输出审查报告
    → 进入最终交付
```

**To Next Stage:** 各 agent 审查报告、评分聚合结果、问题列表 (按严重度排序)、退化的审查基准 → G4 (若失败) / 最终交付 (若通过)

## 触发内容审查

QG3 通过后，读取用户会话的 trigger_intent 检查是否含内容审查触发词。
如果命中 → 路由至 G3.5。
