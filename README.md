# course-report-skill

高校《课程报告》Word 排版技能的公开仓库。

把学校模板里那些琐碎的格式要求（页面分节、标题编号、图表题注、页码页眉、参考文献）固化成一个可复用的 agent skill，让写报告时直接套规范，而不是写完再逐项返工。

## 内容

```
skills/
└── course-report-docx/
    ├── SKILL.md                  技能说明与工作流程
    ├── references/
    │   ├── format-spec.md        完整格式规范（20 节，凡涉及具体数值以它为准）
    │   └── writing-guide.md      章节骨架与写作要求
    ├── assets/
    │   └── course-report-template.docx   主模板（含固化封面、内嵌字体、5 个分节）
    └── scripts/
        ├── prepare_report.py     从模板生成"干净起点"
        ├── check_report.py       XML 级格式自检
        └── render_check.ps1      渲染级自检（可选，需 Word）
```

## 安装

### 方式一：CC Switch（推荐，可一键更新）

仓库管理 → 添加仓库，填四项：

| 字段 | 值 |
|---|---|
| Owner | 你的 GitHub 用户名 |
| Name | `course-report-skill` |
| Branch | `main` |
| Subdirectory | `skills` |

添加后点「刷新」，技能卡片出现，点「安装」。

### 方式二：手动复制

把 `skills/course-report-docx/` 整个目录复制到目标应用的 skills 目录，例如：

- Claude Code：`~/.claude/skills/`
- Codex：`~/.codex/skills/`
- Gemini CLI：`~/.gemini/skills/`
- WorkBuddy：`~/.workbuddy/skills/`

## 更新机制

CC Switch 用 **SHA-256 内容哈希**比对本地已安装目录与远端仓库，不读版本号。只要远端任何文件字节有变化，卡片就会显示「有新版本」，点一下即可覆盖更新。

由此带来的两条纪律：

- **不要手改本地副本**。手改会让哈希对不上，显示"有更新"；点更新又会把手改覆盖掉。
- **不要在 main 上推半成品**。改一个标点都会立刻变成一次"更新"。

## 开发流程

日常只在 WorkBuddy 的副本里改（`~/.workbuddy/skills/course-report-docx`），改完用脚本同步到本仓库：

```powershell
powershell -ExecutionPolicy Bypass -File tools/sync-from-workbuddy.ps1
```

脚本会做镜像复制 → 显示变更 → 提交 → 推送。加 `-Message "..."` 自定义提交信息，加 `-NoPush` 只提交不推送。

## 环境要求

- `prepare_report.py` / `check_report.py`：Python 3，只用标准库，无需 pip 安装
- `render_check.ps1`：Windows + Microsoft Word（可选，检测不到 Word 会跳过并返回 0）

## 版本

见 [CHANGELOG.md](CHANGELOG.md)。当前 `1.0.0`。
