# 🎯 智慧補救教學系統 SRAS 3.0
**Smart Remediation and Assessment System**

> 拍照上傳考卷 → AI 自動批改 → 找出迷思概念 → 生成補救題目

---

## 功能概覽

### 📷 批改考卷
- 拍照或選取學生考卷照片，**一次可上傳多張**（批次批改）
- 可選填「答案卷照片」讓 AI 自動比對正確答案
- AI 逐張分析：辨識學生姓名/座號、答錯題號、得分、迷思概念
- 支援相機即時拍攝（桌機 / 手機均可）

### 📊 班級分析
- 自動統計班級平均、及格率、最高/最低分
- 成績分佈圖、迷思概念熱點圖、學生排名圖
- 一鍵跳轉生成對應補救題目

### 💡 補救題庫
- **自動偵測**：從所有批改記錄統計迷思概念，按出現頻率排列，一鍵生成
- **手動指定**：自選科目、年級、迷思概念，生成選擇題或填充題
- 附帶成因分析與詳細解析

### 🏠 首頁 Dashboard
- 批改人次、班級平均、最常見迷思等統計
- 成績分佈圓餅圖、趨勢折線圖、迷思概念長條圖
- 歷史記錄搜尋與 CSV 匯出

---

## 快速開始

### 1. 設定 AI API Key
進入「⚙️ 設定」頁面，輸入至少一組 API Key：

| 提供者 | 取得方式 | 推薦模型 |
|--------|----------|----------|
| Google Gemini | [ai.google.dev](https://ai.google.dev)（免費） | gemini-2.0-flash-exp |
| OpenAI GPT-4o | [platform.openai.com](https://platform.openai.com) | gpt-4o-mini |

兩組都設定時系統自動備援，一個失敗換另一個。

### 2. 建立考卷
「📷 批改考卷」→ 填入：
- 學年度（114 / 115 / 116）+ 學期（上 / 下）
- 考卷名稱、科目、年級
- 選填：上傳答案卷照片（讓 AI 知道正確答案）
- 選填：輸入班級名單（座號 + 姓名）

### 3. 批次批改
選取考卷 → 上傳多張學生考卷照片 → 點「🤖 開始 AI 批改」
- 每張照片獨立分析，即時顯示進度
- 完成後自動儲存至本機與雲端

### 4. 生成補救題目
從批改結果頁點「💡 生成補救題目」，或進入「💡 補救題庫」查看自動偵測的迷思清單，一鍵生成練習題。

---

## 雲端後台設定（選填）

使用 Google Apps Script 將資料儲存至 Google 試算表，跨裝置共享。

1. 開啟 Google 試算表 → 擴充功能 → Apps Script
2. 貼上 `remediation-backend.gs` 的程式碼
3. 部署 → 新增部署作業 → 網頁應用程式
   - 以下列身分執行：**我（Me）**
   - 誰可以存取：**所有人（Anyone）**
4. 複製部署網址，貼入系統「設定」→「Google Apps Script 網址」
5. 點「🔌 測試連線」確認成功

---

## 檔案說明

```
Abot/
├── remediation-system.html   # 主程式（單一 HTML，全功能）
├── remediation-backend.gs    # Google Apps Script 後台程式碼
└── README.md                 # 本說明文件
```

---

## 技術規格

- **前端**：純 HTML / CSS / JavaScript（無框架，單一檔案）
- **AI**：Google Gemini Vision API + OpenAI GPT-4o Vision API（自動備援）
- **圖表**：Chart.js 4.4.1
- **本機儲存**：localStorage
- **雲端儲存**：Google Apps Script + Google Sheets
- **部署**：GitHub Pages（無需伺服器）

---

## 支援科目

數學 / 英文 / 國語 / 自然理化 / 社會

---

## 注意事項

- API Key 僅儲存於本機瀏覽器，不會傳送至任何第三方伺服器
- 考卷照片經 Base64 編碼後直接傳送至 Google / OpenAI API 進行分析
- 建議使用 Chrome / Edge 瀏覽器以獲得最佳體驗
- 手機使用時相機功能透過 `capture="environment"` 調用後置鏡頭
