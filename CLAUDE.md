# 專案筆記（給未來處理這個 repo 的 Claude 看）

## Ybot 已搬到獨立 repo

- Ybot 個人助理（`ybot.html`／`ybot-backend.gs`）已經**不在這個 repo**了，獨立成
  [`cyysongyy/Ybot`](https://github.com/cyysongyy/Ybot)（大寫 Y；線上網址
  `https://cyysongyy.github.io/Ybot/`）。若使用者提到「Ybot」或貼 Ybot 的網址，
  去那個 repo 處理，不要在這裡找 `ybot.html`（已刪除）或誤以為要重建它。
- 那個 repo 有自己的 CLAUDE.md，記錄了同樣的 AI 模型棄用注意事項（見下方）。

## AI 模型設定（適用本 repo 現存的 `index.html`／`health.html`／`health-backend.gs`）

- 呼叫 Gemini API 時，優先使用穩定別名 `gemini-flash-latest`，**避免**寫死成
  `gemini-2.5-flash` 這種帶版本號的字串。Google 會定期棄用舊版本（曾在官方
  下架日之前就開始回傳「model not found」），寫死版本號會讓使用者莫名其妙
  「AI 突然都失效」。若之後又要指定特定世代（例如效能/價格考量），至少也要
  先查證該版本目前仍受支援。
  > 目前 `index.html`（考卷批改）、`health.html`、`health-backend.gs` 仍寫死
  > `gemini-2.5-flash`，尚未套用這個修正——之前只在 Ybot 修過，處理這幾個
  > 檔案時記得一併檢查。
- NVIDIA 的 `NV_DEFAULT_MODEL` 也會被 NVIDIA 常態性換掉／棄用。使用者回報
  「AI 呼叫失效」時，先懷疑是不是模型 id 過期，可上
  https://build.nvidia.com/models 查目前可用的模型 id 再更新。
- 使用者貼過來的 AI API Key（Gemini／OpenAI／NVIDIA）**絕對不要**寫進程式碼、
  commit、PR 說明或任何會進 repo 的地方。Key 只存在使用者自己瀏覽器的
  localStorage（前端設定頁）或 Apps Script 的 PropertiesService（後台，
  透過 `setAiConfig()` 由使用者自己在 Apps Script 編輯器執行）。只需要指引
  使用者去哪個欄位貼，不要代為儲存。

## Git 工作流程

- 這個使用者常常直接在 GitHub 上把 PR 合併掉，不會等在這邊確認。
  每次要 push 前，先 `git fetch origin main` 再
  `git merge-base --is-ancestor <上一個commit> origin/main` 檢查是否已被合併；
  已合併的話要 `git reset --hard origin/main` 後把還沒推的變更疊上去，
  重新開一個 PR，不要假設舊 PR 還開著。
