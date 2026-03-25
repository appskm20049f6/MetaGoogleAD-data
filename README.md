# MetaGoogleAD-data

這是一個用來查詢與整理 Meta、Google Ads 廣告資料的本機儀表板，後端使用 Node.js + Express，前端使用原生 HTML、CSS、JavaScript。

目前前台提供以下頁面：

- Meta
- Google
- 數據統整
- 投放簡表

Appsflyer 前端頁面已移除，但後端 Appsflyer 匯入與查詢 API 仍保留在專案內。

## 功能概要

### Meta 頁

- 依日期區間查詢 Meta 廣告資料
- 支援 campaign、adset、ad 層級
- 顯示曝光、點擊、CTR、AF 下載、CPI、花費
- 可儲存多組帳號設定
- 支援 Meta token 重新授權流程

### Google 頁

- 依日期區間查詢 Google Ads 資料
- 顯示曝光、點擊、CTR、AF 下載、CPI、花費
- 支援多組 Google Customer ID 設定

### 數據統整

- 整合 Meta 與 Google 最近一次查詢結果
- 依 Android、iOS、網頁預約、未分類做花費與成果對照
- 顯示來源明細與整體加總

### 投放簡表

- 依最近一次 Meta / Google 查詢結果產出簡表
- 產生 FB / GA 與 Android / iOS 的成本對照
- 可將結果存成便利貼卡片留存

## 專案結構

- [index.js](index.js)：Express 伺服器與 API
- [public/index.html](public/index.html)：前端頁面骨架
- [public/app.js](public/app.js)：前端互動與資料渲染
- [public/styles.css](public/styles.css)：樣式
- [.env.example](.env.example)：環境變數範例
- [get_refresh_token.js](get_refresh_token.js)：Google Ads OAuth refresh token 輔助腳本
- [data/](data/)：本機資料目錄

## 環境需求

- Node.js 18 以上為佳
- npm
- Meta 或 Google Ads 至少一組有效憑證

## 安裝方式

1. 安裝套件

```bash
npm install
```

2. 建立本機環境變數檔

將 [.env.example](.env.example) 複製成 .env，並填入你的帳號資訊。

必要規則：

- 若要使用 Meta，至少要有 ACCESS_TOKEN 與 AD_ACCOUNT_ID
- 若要使用 Google，至少要有 GOOGLE_ADS_CLIENT_ID、GOOGLE_ADS_CLIENT_SECRET、GOOGLE_ADS_DEVELOPER_TOKEN、GOOGLE_ADS_REFRESH_TOKEN、GOOGLE_ADS_CUSTOMER_ID
- 兩邊都沒設時，服務不會啟動

## 環境變數

### Meta

- ACCESS_TOKEN：Meta 長效 access token
- AD_ACCOUNT_ID：廣告帳號，格式需為 act_1234567890
- META_APP_ID：Meta App ID
- META_APP_SECRET：Meta App Secret
- META_OAUTH_REDIRECT_URI：Meta OAuth callback URL

### Google Ads

- GOOGLE_ADS_CLIENT_ID
- GOOGLE_ADS_CLIENT_SECRET
- GOOGLE_ADS_DEVELOPER_TOKEN
- GOOGLE_ADS_REFRESH_TOKEN
- GOOGLE_ADS_CUSTOMER_ID
- GOOGLE_ADS_LOGIN_CUSTOMER_ID
- GOOGLE_ADS_API_VERSION

### Server

- PORT：預設 3000

## 啟動方式

```bash
npm start
```

啟動成功後，打開：

```text
http://localhost:3000
```

如果 3000 port 已被占用，可以改用其他 port：

PowerShell：

```powershell
$env:PORT=3001
npm start
```

## 集中分享部署（你控 env，其他人連網址）

如果你要讓同事直接連你的服務，建議使用「單一主機集中控管」模式：

1. 在主機設定完整 `.env`（Meta/Google/APPS_SCRIPT_URL）
2. 啟用分享保護（建議）

```env
SHARE_USERNAME=admin
SHARE_PASSWORD=請改成強密碼
```

3. 啟動服務

```bash
npm start
```

4. 對外提供網址（例如反向代理 HTTPS 網域，或暫時用 tunnel）

注意：

- 若有設定 `SHARE_USERNAME/SHARE_PASSWORD`，系統會啟用 Basic Auth
- `/health`、`/meta-token-callback.html`、`/api/meta/token`、`/api/meta/reauth-url` 不會被攔截
- `.env` 不要上傳到版本控制
- 對外分享時建議一定要走 HTTPS

### Windows 常駐（PM2）

1. 安裝相依套件

```bash
npm install
```

2. 啟動 PM2 常駐程序

```bash
npm run pm2:start
```

3. 查看執行日誌

```bash
npm run pm2:logs
```

4. 更新程式後重啟

```bash
npm run pm2:restart
```

5. 停止服務

```bash
npm run pm2:stop
```

備註：

- PM2 讀取 [ecosystem.config.cjs](ecosystem.config.cjs)
- 環境變數仍以主機上的 `.env` 為主

## API 概要

### 廣告查詢

- GET /api/insights：Meta 資料查詢
- GET /api/google/insights：Google Ads 資料查詢

### Meta 授權

- GET /api/meta/reauth-url：取得重新授權網址
- POST /api/meta/token：更新 ACCESS_TOKEN

### Appsflyer 後端 API

目前前端未使用，但 API 還保留：

- POST /api/appsflyer/import：匯入 Appsflyer CSV
- GET /api/appsflyer/files：列出已建檔資料
- GET /api/appsflyer/insights：查詢 Appsflyer 統整結果

### 健康檢查

- GET /health

## 分類邏輯補充

系統會依名稱與資料欄位推斷分類：

- 名稱中含 Android、AOS、AND\_，或獨立 AND 關鍵字，視為 Android
- 名稱中含 iOS、iPhone、iPad、ios\_，視為 iOS
- Meta 名稱中若符合 A+SC 類型，會歸為網頁預約
- 其餘歸為未分類

## 注意事項

- 這個專案以本機使用為主，未做完整權限管理
- .env 不要提交到版本控制
- data 目錄內可能包含本機查詢或匯入資料，提交前請先確認
- 若 Meta token 過期，前端會引導重新授權

## 常見問題

### 啟動失敗

若看到 port 被占用，代表已有其他程式使用同連接埠，請改用其他 PORT 或先關閉舊程序。

### Google 查不到資料

請先確認：

- OAuth 憑證是否正確
- refresh token 是否有效
- developer token 是否可用
- customer ID 格式是否正確

### Meta 查不到資料

請先確認：

- ACCESS_TOKEN 是否有效
- AD_ACCOUNT_ID 是否正確
- 帳號是否有 ads_read 權限

## License

ISC
