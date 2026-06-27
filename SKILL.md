---
name: gongwen-writing
description: 党政机关公文写作（GB/T 9704-2012 标准）。Use when writing Chinese government official documents — 通知, 报告, 请示, 批复, 函, 纪要, 通报, 通告, 公告, 决定, 决议, 命令, 公报, 意见, 议案 — or when checking an existing .docx for 公文 format compliance.
---

# 公文写作 — 党政机关公文 GB/T 9704-2012

独立完整的公文写作模块，可从用户指令直接调用。

## 能力范围

适用于以下场景：
- **撰写新公文**: 15种法定文种（通知、报告、请示、批复、函、纪要等）
- **格式审查**: 检查现有 .docx 是否符合 GB/T 9704-2012 标准

## 两种模式

| 模式 | 适用 | 流程 |
|------|------|------|
| Auto (全自动) | 文种+背景明确 | G1(规划)→G2(起草)→G3(审查)→G4(修订)，详见 `auto_mode/orchestrator.md` |
| Semi-Auto (半自动) | 需逐阶段确认 | 每阶段门控暂停，等用户确认 |

## 管道阶段

```
G0(前置:文种确认) → G1(结构规划) → G2(起草.docx) → G3(格式审查) → G4(修订收敛)
```

**会话变量:** `output_dir`（输出目录）、`doc_slug`（文档标识）、`doc_dir`（当前文档目录）在 `manifest.yaml` 中定义，管道启动时自动解析。

## 前置条件

- Python 3.8+ with `python-docx`
- 推荐安装系统字体:
  - 方正小标宋简体 (发行机关标志/标题用)
  - 仿宋_GB2312 (正文用)
  - 楷体_GB2312 (二级标题用)
  - 黑体 (一级标题用)
  - 缺省时脚本会回退到宋体并给出警告

## 快速开始

```
请帮我写一份关于做好2024年元旦春节期间有关工作的通知，发文机关是XX省人民政府办公厅

请检查这份公文是否符合GB/T 9704-2012格式标准: /path/to/document.docx
```

## 文件结构

| 文件/目录 | 用途 |
|----------|------|
| `SKILL.md` | 本文件 — 模块入口 |
| `manifest.yaml` | 声明式加载清单 — stages、axes、references、session variables |
| `strategists/` | G0-G4 各阶段策略指令 |
| `auto_mode/` | 全自动模式编排流程定义 |
| `static/core/` | 22 个核心规则文件（共享自 academic-shared） |
| `references/` | GB/T 9704 规范、文种模板、语言规范、版记指南 |
| `evaluate/` | 格式审查子系统（agent 提示词 + 评分脚本 + 配置） |
| `scripts/` | docx 生成、格式转换、发文处理笺、格式检查脚本 |
| `composers/` | 公文 docx 组装规范 |
| `skills-embedded/` | python-docx 参考快照 |

## 与 academic-shared 的依赖关系

本模块通过 `manifest.yaml` 引用以下共享组件：
- `../academic-shared/composers/*.md` — 共享 composer
- `../academic-shared/evaluate/scripts/scoring.py` — 评分聚合脚本
- `../academic-shared/evaluate/scripts/diff_issues.py` — 退化检测脚本
- `academic-shared/GLOSSARY.md` — 通用管道术语表
