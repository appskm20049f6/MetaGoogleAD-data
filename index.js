const axios = require('axios');
const express = require('express');
const path = require('path');

// 若有安裝 dotenv，嘗試載入本機 .env
try {
    require('dotenv').config();
} catch (e) {
    // dotenv not installed — 繼續，改用系統環境變數
}

// 以環境變數為主，必要時中止並提示
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const AD_ACCOUNT_ID = process.env.AD_ACCOUNT_ID; // 記得保留 act_ 開頭
const PORT = process.env.PORT || 3000;

if (!ACCESS_TOKEN || !AD_ACCOUNT_ID) {
    console.error('❌ 環境變數 ACCESS_TOKEN 或 AD_ACCOUNT_ID 未設定。請建立 .env 或設定系統環境變數。');
    console.error('範例請參考 .env.example');
    process.exit(1);
}

const app = express();

app.use(express.static(path.join(__dirname, 'public')));

function isValidDateFormat(dateText) {
    return /^\d{4}-\d{2}-\d{2}$/.test(dateText);
}

function isValidAdAccountId(accountId) {
    return /^act_\d+$/.test(accountId);
}

async function fetchAdsReport(options = {}) {
    const {
        datePreset = 'last_7d',
        since,
        until,
        level = 'campaign',
        accountId = AD_ACCOUNT_ID
    } = options;

    const url = `https://graph.facebook.com/v19.0/${accountId}/insights`;
    const params = {
        access_token: ACCESS_TOKEN,
        level,
        fields: 'campaign_name,adset_name,ad_name,date_start,date_stop,spend,impressions,clicks,actions'
    };

    if (since && until) {
        params.time_range = JSON.stringify({ since, until });
    } else {
        params.date_preset = datePreset;
    }

    const response = await axios.get(url, { params });
    return response.data.data || [];
}

app.get('/api/insights', async (req, res) => {
    try {
        const { datePreset, since, until, level, accountId } = req.query;

        if ((since && !until) || (!since && until)) {
            return res.status(400).json({
                error: '請同時提供 since 與 until，或僅使用 datePreset。'
            });
        }

        if (accountId && !isValidAdAccountId(accountId)) {
            return res.status(400).json({
                error: 'accountId 格式錯誤，需為 act_ 開頭加數字。'
            });
        }

        if (since && until) {
            if (!isValidDateFormat(since) || !isValidDateFormat(until)) {
                return res.status(400).json({
                    error: '日期格式需為 YYYY-MM-DD。'
                });
            }
        }

        const data = await fetchAdsReport({
            datePreset: datePreset || 'last_7d',
            since,
            until,
            level: level || 'campaign',
            accountId: accountId || AD_ACCOUNT_ID
        });

        return res.json({ data });
    } catch (error) {
        const errorMessage = error.response?.data || { message: error.message };
        return res.status(500).json({
            error: '抓取失敗',
            details: errorMessage
        });
    }
});

app.get('/health', (req, res) => {
    res.json({ ok: true });
});

const server = app.listen(PORT, () => {
    console.log(`✅ 服務啟動：http://localhost:${PORT}`);
});

server.on('error', (error) => {
    if (error.code === 'EADDRINUSE') {
        console.error(`❌ Port ${PORT} 已被占用，請先關閉舊服務或改用其他連接埠。`);
        console.error(`👉 可用：set PORT=3001; node index.js`);
        process.exit(1);
    }

    console.error('❌ 服務啟動失敗:', error.message);
    process.exit(1);
});