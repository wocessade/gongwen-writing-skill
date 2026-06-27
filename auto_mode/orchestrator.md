# 全自动模式编排器 (Auto Mode Orchestrator)

**Purpose:** 定义公文写作全自动模式下的端到端编排流程。用户在 G0 提供文种和背景后，无需中间确认即可获得符合 GB/T 9704 标准的 .docx 公文。

**适用场景:**
- 文种和背景明确，无需逐阶段确认
- 常见公文写作任务（通知、报告、请示等）
- 时间紧迫，需要快速出稿

---

## 自动编排流程

```
┌─────────────────────────────────────────────────────────┐
│  Step 0: 接收用户指令                                   │
│  "帮我写一份关于...的通知"                              │
│  -> 提取 documentType, 发文机关, 标题主题, 主送机关      │
│  -> 设置 mode=auto                                      │
│  -> 加载 manifest.yaml                                   │
│  -> 检测 prerequisites (python-docx 等)                   │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 1: G0 Auto — 自动文种识别 + 上下文收集             │
│  -> 从用户输入识别文种                                   │
│  -> 提取发文机关、标题主题、主送机关                      │
│  -> 检测是否上行文 -> 询问签发人                         │
│  -> 检测密级/紧急程度提及                               │
│  ┌──────────────────────────────────────────────────┐   │
│  │  熔断: 文种不明确时 STOP-AND-ASK                  │   │
│  │  不自动猜测文种                                   │   │
│  └──────────────────────────────────────────────────┘   │
│  -> Auto-check QG0                                       │
│  -> 写入 passport                                        │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 2: G1 Auto — 自动结构规划 + 内容大纲               │
│  -> 按 documentType 加载 references/模板                 │
│  -> 自动规划章节结构                                     │
│  -> 自动列出每节核心信息点                                │
│  -> 自动规划附件 (如需)                                  │
│  -> 标注签发人 (上行文)                                  │
│  -> Auto-check QG1                                       │
│  -> 仅门控失败时 STOP-AND-ASK                            │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 3: G2 Auto — 自动起草 .docx                       │
│  -> 加载 composers/gongwen-docx-assembly.md             │
│  -> 加载 references/gb-t9704-2012-quick-ref.md          │
│  -> 构建 config dict                                     │
│  -> 生成并执行 generate_gongwen_docx.py                  │
│  -> 确认 .docx 成功生成                                  │
│  -> Auto-check QG2                                       │
│  -> 仅门控失败时 STOP-AND-ASK                            │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 4: G3 Auto — 自动审查                             │
│  -> 运行 format_compliance_gongwen (1x)                  │
│  -> 并行运行 language_formality (3x)                     │
│  -> 并行运行 structure_completeness (3x)                  │
│  -> 运行 attachment_seal_check (1x)                      │
│  -> 聚合评分 + 计算加权总分                              │
│  ┌──────────────────────────────────────────────────┐   │
│  │  分支:                                           │   │
│  │  总分 >= 70 + attachment_seal PASS + 无 Critical  │   │
│  │    -> 自动通过 -> 进入最终交付                     │   │
│  │  总分 < 70 或 attachment_seal FAIL 或有 Critical  │   │
│  │    -> 自动进入 G4 修订循环                        │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 5: G4 Auto — 自动修订循环 (最多 3 轮)              │
│  -> 汇总审查问题按严重度排序                              │
│  -> 逐项修复 (修改脚本 + 重新生成 .docx)                  │
│  -> 重新运行 G3 完整审查                                 │
│  -> 检测退化                                            │
│  ┌──────────────────────────────────────────────────┐   │
│  │  if 收敛: 自动通过                               │   │
│  │  if 3 轮耗尽: STOP-AND-ASK 报告剩余问题           │   │
│  │  if 退化出现: STOP-AND-ASK 报告退化详情           │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────┬───────────────────────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────┐
│  Step 6: 最终交付                                       │
│  -> 输出 final_{slug}.docx                              │
│  -> 生成交付报告:                                        │
│     - 文件名、文种、发文机关                              │
│     - 审查轮次、加权总分、问题统计                         │
│     - 已知剩余问题 (如有)                                │
│  -> 交付完成                                            │
└─────────────────────────────────────────────────────────┘
```

## 自动决策规则

### G0 Auto — 文种识别

```
if 用户明确提到文种名称:
    → 直接确认，不询问
elif 用户描述场景但文种模糊:
    → STOP-AND-ASK — 列出 2-3 个候选文种
    → 用户选择后继续 — 不自动猜测
else:
    → STOP-AND-ASK — 无法识别文种

密级/紧急程度: 仅当用户提及时才设置
  用户未提及 → 默认 none/normal，不询问
```

### G0 Auto — 上下文收集

```
if 用户一次性提供了所有字段:
    → 自动全部提取，不中断
elif 缺少字段:
    → 仅在 G0 门控检查时询问缺失字段
    → 字段全部确认后自动推进

缺失字段自动询问规则 (一次性全部问完，不逐个打断):
    "请确认以下信息:
     - 发文机关: ?
     - 标题主题: ?
     - 主送机关: ?
     - (上行文) 签发人: ?"
```

### G3 Auto — 审查结果判断

```
if 加权总分 >= 70 AND attachment_seal PASS AND 无 Critical:
    → 自动通过 → 进入最终交付
else:
    → 自动进入 G4 修订循环
    → 不中断用户
```

### G4 Auto — 修订循环

```
Critical 问题: 全部自动修复
Major 问题: 自动修复，记录修复方式
Minor 问题: 时间允许时修复

收敛检测:
  if 无新的 Critical AND Major 项数 <= 2 新增 AND 无退化:
      → 自动通过
  else:
      → 继续下一轮

3 轮耗尽:
  → STOP-AND-ASK
  → 报告: "已自动修订 {N} 轮，仍有 {M} 个问题未解决"
  → 用户决定: 继续 / 接受 / 重写
```

## 故障处理

| 问题 | 自动行为 | 升级条件 |
|------|---------|---------|
| G0 文种不明确 | STOP-AND-ASK，不自动猜测 | — |
| G1 模板加载失败 | 尝试相近文种模板 | 无模板可用 -> STOP-AND-ASK |
| G2 脚本执行失败 | 重试 1 次 (修复语法错误) | 第二次失败 -> STOP-AND-ASK |
| G3 agent 失败 | 自动重试 1 次 | 2+ agents 失败 -> STOP-AND-ASK |
| G4 收敛退化 | 自动重置到上一轮 + 重新修复 | 两次退化 -> STOP-AND-ASK |
| G4 3 轮耗尽 | 报告剩余问题 + STOP-AND-ASK | 用户决定后续方向 |
| python-docx 未安装 | 自动提示安装命令 | 用户拒绝 -> STOP |

## 自动日志

全自动模式下，每步输出记录在 `{output_dir}/auto_decision_log.md`:

```markdown
# Auto Decision Log

## Step 1: G0 Auto
- Document Type: 通知
- Organ: XX省人民政府办公厅
- Urgency: normal
- Security Level: none
- Mode: auto

## Step 2: G1 Auto
- Template: 通知 (references/document-type-templates.md)
- Sections: [标题, 主送, 缘由, 事项, 要求, 发文机关, 成文日期]
- Attachments: none

## Step 3: G2 Auto
- Script: {output_dir}/generate_gongwen_docx.py
- Output: {output_dir}/final_{slug}.docx
- Font warnings: [仿宋_GB2312 not found, fallback to 仿宋]

## Step 4: G3 Auto
- format_compliance: 85
- language_formality: 78 (median of 3)
- structure_completeness: 82 (median of 3)
- attachment_seal: PASS
- Weighted total: 81.2
- Critical: 0, Major: 3, Minor: 5
- Result: < 70 threshold → auto-routing to G4

## Step 5: G4 Auto (Round 1/3)
- Fixed: 3 Major (format), 2 Minor
- Regenerated: yes
- Degradation detected: no
- Re-review total: 84.5
- Convergence: yes → exit

## Step 6: Final Deliverable
- Output: {output_dir}/final_{slug}.docx
- Total rounds: 1
- Remaining issues: 3 Minor (deferred)
```
