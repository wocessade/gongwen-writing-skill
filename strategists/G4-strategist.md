# Stage G4: 修订 + 收敛循环 [Strategist]

**Gate:** QG4 (BLOCK)
**Goal:** 修复 G3 审查发现的问题，经最多 3 轮收敛循环后达到门控标准
**Needs Composers:** [gongwen-docx-assembly]

## Convergence Loop

```
max_rounds = 3
for round in 1..max_rounds:
    1. 汇总本轮审查问题 → 按严重度排序
    2. 逐项修复 (内容修改 → 修改 G2 脚本 → 重新生成 .docx)
    3. 重新运行 G3 审查 (完整 4 agent)
    4. 评估收敛条件:
        - 无新的 Critical 项
        - Major 项数减少 (或 ≤ 2 个新的)
        - 无退化 (已修复问题未再次出现)
    5. if 收敛:
        → PASS → 退出循环
    6. elif round == max_rounds:
        → STOP-AND-ASK — 报告剩余问题，用户决定是否接受
    7. else:
        → 继续下一轮
```

## Decisions

### Axis 1: 问题汇总与排序

```
从 G3 审查报告汇总所有问题，按严重度排序:

优先级队列:
  1. Critical 问题 (必须修复)
  2. Major 问题 (必须修复或标注)
  3. Minor 问题 (建议修复，时间允许时修复)

排序规则:
  - 同严重度: format_compliance → content_compliance → language_formality → structure_completeness
  - 跨 agent 冲突: 若 3 个 language_formality agent 分差 > 2，优先核查不一致项目
```

### Axis 2: 修复策略

```
修复方式分类:
  格式问题 (format_compliance):
      → 修改 G2 脚本中对应参数
      → 重新执行脚本生成 .docx
      → 不直接编辑 .docx XML (除非极端情况)

  语言问题 (language_formality):
      → 修改脚本中的正文文本
      → 按 formal-language-guide.md 修正措辞
      → 重新生成 .docx

  结构问题 (structure_completeness):
      → 按 G1 模板调整章节结构
      → 补充缺失的内容要素
      → 重新生成 .docx

  附件问题 (attachment_seal):
      → 补充或调整附件标记
      → 修正印章/日期信息
      → 重新生成 .docx

  内容合规问题 (content_compliance):
      → 对 UNVERIFIED 声明: 替换为已确认信息，或删除无法验证的声称
      → 对 VERIFIED-SOURCE 标注: 保留现有内容，可添加来源标记
      → 修改 G2 脚本中的文本内容 → 重新生成 .docx
      → 涉及机构/人员/政策名称错误的，替换为正确名称后重生成
```

### Axis 3: 退化检测

使用 `academic-shared/evaluate/scripts/diff_issues.py` 对比相邻两轮审查结果。

```
退化定义:
  - 上一轮已修复的问题在本轮重新出现 (完全相同)
  - 上一轮 P ASS 的项目在本轮变为 FAIL

退化处理:
  if 检测到退化:
      → STOP-AND-ASK — 报告退化详情
      → 分析 root cause:
          - 修复未正确应用
          - 修复引入关联问题
          - 审查 agent 评分不一致
      → 根据 root cause 调整修复策略
      → 重置该问题修复状态，重新尝试
  if 同一问题两次退化:
      → STOP-AND-ASK — 建议手动介入或接受现状
```

### Axis 4: 停止条件与电路熔断

```
正常收敛:
  - 无新的 Critical 项
  - Major 项数 ≤ 2 个新增
  - 无退化
  → PASS, 退出循环

电路熔断 (Circuit Breakers):
  1. 达到 max_rounds (3轮) 仍不收敛:
      → STOP-AND-ASK
      → 报告剩余问题
      → 用户决定: 继续修订 / 接受现状 / 重新起草
  2. 检测到退化:
      → STOP-AND-ASK (首次退化警告)
      → 两次退化 → 建议手动介入
  3. G3 审查 agent 不可用或返回空结果:
      → 重试 1 次
      → 第二次失败 → STOP-AND-ASK
```

## Gate

### QG4 Gate Checklist

- [ ] 所有 Critical 问题已修复
- [ ] 所有 Major 问题已修复或标注为已知
- [ ] 退化检测结果: 无退化
- [ ] 收敛轮次 ≤ 3
- [ ] 加权总分 ≥ 70 (最新一轮 G3)
- [ ] attachment_seal 检查 PASS (最新一轮)
- [ ] 如果 3 轮未收敛: STOP-AND-ASK 已完成

### Gate Failure Route

```
if 有未修复的 Critical 问题且轮次已耗尽:
    → BLOCK — 报告剩余 Critical 问题
    → 用户决定: 额外修订 / 接受风险
    → 如果用户接受，标记为[用户接受: 风险已告知]
if 退化问题未解决:
    → BLOCK — 建议手动审查和修正
    → 用户确认后继续
if 3 轮尚未耗尽:
    → 继续下一轮修订
if 3 轮耗尽且未收敛:
    → SOFT BLOCK — 报告剩余问题
    → 用户决定:
        - 超出轮次继续修订 → 加 1 轮
        - 接受现状 → 标记为 [未完全收敛: {剩余问题摘要}]
        - 放弃当前版本 → 从 G1 重新开始
```

**To Next Stage:** 最终 .docx 文件路径、修订轮次统计、修复问题清单、剩余已知问题(如有)、退化日志(如有) → 最终交付
