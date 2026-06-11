# 折扣率分析儀表板

新地國際 · 實體門市折扣率自動分析 · GitHub Pages 部署

---

## Repo 結構

```
discount-dashboard/
├── data/
│   ├── 門市基本資料.xlsx          ← 固定參照，更新時直接覆蓋
│   ├── AIRSPACE_線下通路_行銷排程_含甘特圖_2025-2026.xlsx
│   ├── 銷售_2026_3月.xlsx
│   ├── 銷售_2026_4月.xlsx
│   └── 銷售_2026_5月.xlsx        ← 每月新增
├── scripts/
│   └── generate.py               ← 核心腳本
├── template.html                 ← HTML 模板（不手動修改）
├── index.html                    ← 自動產出，GitHub Pages 對外頁面
├── requirements.txt
└── .github/workflows/
    └── update.yml
```

---

## 第一次設定

### 1. 建立 Repo

```bash
# GitHub 上建立新 repo：bunny0405/discount-dashboard
# 建議設為 public（公司內部可看）
```

### 2. 上傳初始檔案

把以下檔案放進 `data/` 資料夾：
- `門市基本資料.xlsx`
- `AIRSPACE_線下通路_行銷排程_含甘特圖_2025-2026.xlsx`
- 所有歷史月份銷售資料（`銷售_YYYY_M月.xlsx`）

### 3. 開啟 GitHub Pages

1. Repo → **Settings** → **Pages**
2. Source 選 **Deploy from a branch**
3. Branch 選 `main`，資料夾選 `/ (root)`
4. Save

網址會是：`https://bunny0405.github.io/discount-dashboard/`

### 4. 確認 Actions 權限

Repo → **Settings** → **Actions** → **General**
→ Workflow permissions 選 **Read and write permissions** → Save

---

## 每月更新流程

### 自動更新（每月5號）

什麼都不用做，GitHub Actions 會在每月5號 00:00（台灣時間）自動執行。

### 手動立即更新

**方法一：上傳新資料自動觸發**
1. 把新月份 `銷售_2026_X月.xlsx` 上傳到 `data/`
2. Push 到 main → Actions 自動觸發

**方法二：手動跑 Actions**
1. Repo → **Actions** → **每月折扣率儀表板更新**
2. 右上角 **Run workflow** → **Run workflow**

---

## 銷售資料命名規則

| 月份 | 檔名 |
|------|------|
| 2026年1月 | `銷售_2026_1月.xlsx` |
| 2026年10月 | `銷售_2026_10月.xlsx` |

> 注意：月份不補零（`1月` 而非 `01月`）

---

## 資料欄位需求

銷售資料需包含以下欄位（與現行匯出格式一致）：

| 欄位 | 說明 |
|------|------|
| 訂單編號 | 唯一訂單識別 |
| 購買當下會員等級 | 00/01/02/03/09 |
| 倉庫名稱 | 對應門市基本資料 |
| 零售價 | 單品售價 |
| 小計 | 實收金額 |
| 大分類 | 品類 |
| 品別 | 正品/續賣品/出清品等 |
| 行促類型 | 任選/區間折扣等 |
| 行促名稱 | 活動名稱 |
| 任選折扣金額攤算 | 件數折金額 |
| 區間折扣金額攤算 | 單件折金額 |
| 紅利金攤算 | 紅利折抵 |
| 滿額折現攤算 | 滿額折現 |
| 折扣碼攤算 | 折扣碼 |
| 會員折扣攤算 | 會員折 |
| 折價券攤算 | 折價券 |
| 會員生日年度折扣攤算 | 生日折 |

---

## 行銷排程更新

每年底把新年度行銷排程覆蓋 `data/` 裡的排程 xlsx，
下次 generate.py 執行時會自動讀取新年度活動。

---

## 本地端手動執行

```bash
# 安裝套件
pip install -r requirements.txt

# 執行腳本
python scripts/generate.py

# 開啟預覽
open index.html
```

---

## 主指標定義

| 指標 | 定義 |
|------|------|
| 正價折扣率 | 正價門市 × 非員購（00–03）|
| 含OUTLET折扣率 | 全館（含OUTLET）× 非員購 |
| 員購（09） | 參考區，不納入主指標 |
| OUTLET | 參考區，不納入主指標 |
| 出清品/特惠品 | 品別表中標示「排除」|

---

## 聯絡

如有問題請聯絡 Bunny。
