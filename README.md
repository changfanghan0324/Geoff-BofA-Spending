# 爸爸的卡片消費紀錄 — 自動更新網站

## 🔗 常用連結（放這裡以後好找）

| | 網址 |
|---|---|
| 📊 **看網站（Geoff BofA Spending）** | **https://changfanghan0324.github.io/Geoff-BofA-Spending/** |
| 📝 記帳的 Google Sheet | https://docs.google.com/spreadsheets/d/1i9j_XCqT70ZGFq-6IPMvT035l5XjkXtLj0xSIDCy9fQ/edit |
| 💻 GitHub 專案 | https://github.com/changfanghan0324/Geoff-BofA-Spending |
| ⚙️ 手動立刻更新 | 到 GitHub 專案的 **Actions** 分頁 → 「每月更新資料」→ **Run workflow** |

---

把記帳資料放在 **Google Sheet**，網站每個月由 **GitHub Action** 自動去抓最新資料、
產生靜態網頁，放在 **GitHub Pages** 上就能用網址打開看。

- 你只要維護 Google Sheet，網站會自動跟著更新（每月一次，也可以手動立刻更新）。
- 網站是純靜態網頁，沒有伺服器、不用錢。

---

## 檔案說明

| 檔案 | 作用 |
|---|---|
| `index.html` | 網站本體（總覽、逐年比較、分類支出、每月收支、可搜尋排序的交易表） |
| `data.js` | 資料檔，**由程式自動產生**，不用手改 |
| `config.json` | 設定：你的 Google Sheet 網址代碼、網站標題、幣別符號 |
| `scripts/build_data.py` | GitHub Action 用它去抓 Google Sheet 產生 `data.js` |
| `scripts/gen_local_preview.py` | 用桌面現有的 `.xlsx` 先產生預覽（連 Google 之前先看成品） |
| `.github/workflows/update-data.yml` | 每月自動更新的排程 |

---

## 先在本機看看成品（可選）

直接用瀏覽器打開 `index.html` 就能看到目前用桌面 6 個 `.xlsx` 產生的網站。
之後改了 `.xlsx` 想重新產生預覽，在這個資料夾執行：

```bash
pip install openpyxl
python scripts/gen_local_preview.py
```

---

## 設定 Google Sheet（正式上線用）

### 1. 建立試算表，每年一個分頁
- 分頁名稱用年份：`2021`、`2022`、`2023`…（程式會自動把每個分頁當成一個年份）
- 每個分頁第一列是欄位名稱，欄位名稱中英文都可以，順序不拘。程式認得這些：

| 你要放的資料 | 欄位名稱可以寫成（任一種都行） |
|---|---|
| 日期 | `日期` 或 `Date` |
| 種類 | `種類` / `類別` / `分類` / `Category` |
| 詳細說明 | `詳細說明` / `說明` / `Description` |
| 花費價格 | `花費` / `支出` / `Expense` |
| 入金＆轉帳 | `入金` / `轉帳` / `收入` / `Credit` |
| 餘額 | `餘額` / `Balance` |

> 小技巧：如果你只有一欄「金額」（花費記負數、入金記正數），欄位取名 `金額` 或 `Amount` 也可以，
> 程式會自動幫你拆成「花費」和「入金」。

### 2. 把試算表設成「知道連結的人可讀」
右上角「共用」→ 一般存取權改成 **「知道連結的任何人」→「檢視者」**。
（你選了公開沒關係，這步是必要的，Action 才抓得到。）

### 3. 複製試算表代碼填進 `config.json`
你的試算表網址長這樣：
```
https://docs.google.com/spreadsheets/d/【這一段就是代碼】/edit
```
把中間那段代碼貼進 `config.json`：
```json
{
  "sheet_id": "把代碼貼在這裡",
  "title": "爸爸的卡片消費紀錄",
  "currency_symbol": "$"
}
```

---

## 上傳到 GitHub 並開啟網站

### 1. 建一個 repo，把這個資料夾所有檔案上傳
（`index.html`、`data.js`、`config.json`、`scripts/`、`.github/` 都要上傳。
`.github` 是隱藏資料夾，用 GitHub 網頁拖曳上傳有時會漏掉，建議用 GitHub Desktop 或 git 指令。）

### 2. 開啟 GitHub Pages
repo 的 **Settings → Pages** → Source 選 **`main` 分支 / `/ (root)`** → Save。
稍等一下就會給你一個網址（`https://你的帳號.github.io/repo名稱/`），那就是你的網站。

### 3. 讓 Action 有權限自動更新
**Settings → Actions → General → Workflow permissions** →
選 **「Read and write permissions」** → Save。
（這樣每月更新後才能把新的 `data.js` 存回 repo。）

### 4. 第一次先手動跑一次
**Actions 分頁 → 左邊「每月更新資料」→ 右邊「Run workflow」**。
跑完後 `data.js` 會換成 Google Sheet 的最新資料，網站也跟著更新。

---

## 之後怎麼用

- **平常**：只要在 Google Sheet 記帳就好。
- **每月 1 號**：Action 自動抓一次，網站自動更新，你什麼都不用做。
- **想馬上更新**：到 Actions 分頁按一下「Run workflow」即可。
- **新增年份**：在 Google Sheet 多開一個分頁（例如 `2027`），網站下次更新就會自動多一個年份分頁。

---

## 常見問題

- **網站說「還沒有資料」**：代表 `data.js` 沒產生成功。先確認 `config.json` 的 `sheet_id` 正確、
  試算表已設成「知道連結的人可讀」，再到 Actions 手動跑一次看紅字錯誤訊息。
- **金額顯示怪怪的**：確認花費／餘額欄位裡是數字，不要夾雜文字。
- **想改標題或幣別**：改 `config.json` 的 `title` / `currency_symbol`，下次 Action 更新後生效。
