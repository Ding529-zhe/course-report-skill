# AGENTS.md — course-report-skill 仓库规则

## 改动流程

技能本体只在 `~/.workbuddy/skills/course-report-docx` 修改，改完运行：

```powershell
powershell -ExecutionPolicy Bypass -File tools/sync-from-workbuddy.ps1
```

脚本把技能目录镜像到本仓库根目录的 `course-report-docx/`，然后提交并推送。

## 发布要求

- 同步后更新 `CHANGELOG.md` 顶部条目与 `course-report-docx/SKILL.md` 的 `metadata.version`
- 版本号递增后打 tag 并推送：`git push --tags`
- 版本号供人阅读与回滚；CC Switch 按内容哈希检测更新，不读版本号

## 布局约束

- 技能目录必须在仓库根目录一级（`course-report-docx/`），不得移入 `skills/` 等子目录
- 不使用 Git LFS；`.gitattributes` 保持 `* text=auto eol=lf` 加二进制文件标注
- 本仓库公开，禁止写入内网地址、账号、学号等敏感信息
