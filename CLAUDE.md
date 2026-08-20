# 專案筆記（給未來處理這個 repo 的 Claude 看）

## Ybot 已獨立成專屬 repo

- Ybot 個人助理（原 `ybot.html` / `ybot-backend.gs` / `ybot-icon.svg` /
  `ybot-manifest.webmanifest`）已搬到 https://github.com/cyysongyy/ybot ，
  本 repo 不再存放這些檔案。相關的 AI 模型棄用注意事項（Gemini
  `gemini-flash-latest`、NVIDIA `NV_DEFAULT_MODEL`、API Key 處理原則）已同步
  寫進該 repo 自己的 CLAUDE.md，去那邊找。
- `portal.html` 對 Ybot 的卡片現在是外部連結（`https://cyysongyy.github.io/ybot/`），
  跟 principal-assistant／micro-politics 等其他獨立 repo 的處理方式一致。

## Git 工作流程

- 這個使用者常常直接在 GitHub 上把 PR 合併掉，不會等在這邊確認。
  每次要 push 前，先 `git fetch origin main` 再
  `git merge-base --is-ancestor <上一個commit> origin/main` 檢查是否已被合併；
  已合併的話要 `git reset --hard origin/main` 後把還沒推的變更疊上去，
  重新開一個 PR，不要假設舊 PR 還開著。
