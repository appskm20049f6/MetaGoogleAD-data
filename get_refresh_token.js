const { google } = require('googleapis');
const readline = require('readline');

// 自動讀取同資料夾下的 .env 檔案
try {
    require('dotenv').config();
} catch (e) {
    console.error('❌ 找不到 dotenv 套件，請執行 npm install dotenv');
}

// 直接指定從環境變數讀取
const CLIENT_ID = process.env.GOOGLE_ADS_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_ADS_CLIENT_SECRET;
const REDIRECT_URI = 'http://localhost'; 

// 檢查 .env 是否已填寫
if (!CLIENT_ID || !CLIENT_SECRET) {
    console.error('❌ 錯誤：請先在 .env 檔案中填寫 GOOGLE_ADS_CLIENT_ID 與 GOOGLE_ADS_CLIENT_SECRET');
    process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline', // 確保能拿到 Refresh Token
    prompt: 'consent',      // 強制顯示同意畫面
    scope: ['https://www.googleapis.com/auth/adwords'],
});

console.log('請打開此網址進行授權:\n', authUrl);

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.question('\n授權完後，網址會變成 http://localhost/?code=xxxx...\n請複製 code= 後面那串字並貼在此處: ', async (code) => {
    try {
        const { tokens } = await oauth2Client.getToken(code);
        console.log('\n✅ 成功拿到 REFRESH_TOKEN:');
        console.log('------------------------------------');
        console.log(tokens.refresh_token);
        console.log('------------------------------------');
        console.log('\n請將上面這串代碼填入 .env 的 GOOGLE_ADS_REFRESH_TOKEN 欄位。');
    } catch (err) {
        console.error('❌ 取得 Token 失敗:', err.message);
    }
    process.exit();
});