const axios = require('axios');
const express = require('express');
const path = require('path');
const fs = require('fs');

// 若有安裝 dotenv，嘗試載入本機 .env
try {
    require('dotenv').config();
} catch (e) {
    // dotenv not installed — 繼續，改用系統環境變數
}

// 以環境變數為主，必要時中止並提示
let ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const AD_ACCOUNT_ID = process.env.AD_ACCOUNT_ID; // 記得保留 act_ 開頭
const GOOGLE_ADS_CLIENT_ID = process.env.GOOGLE_ADS_CLIENT_ID;
const GOOGLE_ADS_CLIENT_SECRET = process.env.GOOGLE_ADS_CLIENT_SECRET;
const GOOGLE_ADS_DEVELOPER_TOKEN = process.env.GOOGLE_ADS_DEVELOPER_TOKEN;
const GOOGLE_ADS_REFRESH_TOKEN = process.env.GOOGLE_ADS_REFRESH_TOKEN;
const GOOGLE_ADS_CUSTOMER_ID = process.env.GOOGLE_ADS_CUSTOMER_ID;
const GOOGLE_ADS_LOGIN_CUSTOMER_ID = process.env.GOOGLE_ADS_LOGIN_CUSTOMER_ID;
const GOOGLE_ADS_API_VERSION = process.env.GOOGLE_ADS_API_VERSION || 'v22';
const META_APP_ID = process.env.META_APP_ID;
const META_OAUTH_REDIRECT_URI = process.env.META_OAUTH_REDIRECT_URI;
const META_SCOPE = process.env.META_SCOPE || 'ads_read';
const PORT = process.env.PORT || 3000;
const ENV_FILE = path.join(__dirname, '.env');

function hasMetaConfig() {
    return Boolean(ACCESS_TOKEN && AD_ACCOUNT_ID);
}

const hasGoogleConfig = Boolean(
    GOOGLE_ADS_CLIENT_ID
    && GOOGLE_ADS_CLIENT_SECRET
    && GOOGLE_ADS_DEVELOPER_TOKEN
    && GOOGLE_ADS_REFRESH_TOKEN
    && GOOGLE_ADS_CUSTOMER_ID
);

if (!hasMetaConfig() && !hasGoogleConfig) {
    console.error('❌ 尚未設定可用的平台憑證（Meta 或 Google）。');
    console.error('請在 .env 至少完成其中一組設定，範例請參考 .env.example');
    process.exit(1);
}

const app = express();

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '10mb' }));

function isValidDateFormat(dateText) {
    return /^\d{4}-\d{2}-\d{2}$/.test(dateText);
}

function isValidAdAccountId(accountId) {
    return /^act_\d+$/.test(accountId);
}

function isValidGoogleCustomerId(customerId) {
    return /^\d{3}-?\d{3}-?\d{4}$/.test(customerId);
}

function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function resolveDateRange({ datePreset = 'last_7d', since, until }) {
    if (since && until) {
        return { since, until };
    }

    const today = new Date();
    const end = new Date(today);
    let start = new Date(today);

    switch (datePreset) {
    case 'today':
        break;
    case 'yesterday':
        start.setDate(start.getDate() - 1);
        end.setDate(end.getDate() - 1);
        break;
    case 'last_30d':
        start.setDate(start.getDate() - 29);
        break;
    case 'this_month':
        start = new Date(today.getFullYear(), today.getMonth(), 1);
        break;
    case 'last_month':
        start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        end.setFullYear(today.getFullYear(), today.getMonth(), 0);
        break;
    case 'last_7d':
    default:
        start.setDate(start.getDate() - 6);
        break;
    }

    return { since: formatDate(start), until: formatDate(end) };
}

function cleanGoogleCustomerId(customerId) {
    return String(customerId || '').replace(/-/g, '');
}

function upsertEnvVar(key, value) {
    const line = `${key}=${value}`;
    let content = '';

    if (fs.existsSync(ENV_FILE)) {
        content = fs.readFileSync(ENV_FILE, 'utf8');
    }

    const lines = content ? content.split(/\r?\n/) : [];
    let replaced = false;
    const nextLines = lines.map((item) => {
        if (new RegExp(`^${key}=`).test(item)) {
            replaced = true;
            return line;
        }
        return item;
    });

    if (!replaced) {
        nextLines.push(line);
    }

    const nextContent = `${nextLines.filter((item) => item !== '').join('\n')}\n`;
    fs.writeFileSync(ENV_FILE, nextContent, 'utf8');
}

function persistAccessToken(accessToken) {
    ACCESS_TOKEN = accessToken;
    upsertEnvVar('ACCESS_TOKEN', accessToken);
}

function getRequestOrigin(req) {
    const forwardedProto = req.headers['x-forwarded-proto'];
    const protocol = forwardedProto ? String(forwardedProto).split(',')[0].trim() : req.protocol;
    return `${protocol}://${req.get('host')}`;
}

function buildMetaReauthUrl(options = {}) {
    const { origin } = options;
    const fallbackLocalRedirect = `http://localhost:${PORT}/meta-token-callback.html`;
    const redirectUri = origin
        ? `${origin}/meta-token-callback.html`
        : (META_OAUTH_REDIRECT_URI || fallbackLocalRedirect);

    if (!META_APP_ID) {
        return 'https://developers.facebook.com/tools/explorer/';
    }

    const authUrl = new URL('https://www.facebook.com/v19.0/dialog/oauth');
    authUrl.searchParams.set('client_id', META_APP_ID);
    authUrl.searchParams.set('redirect_uri', redirectUri);
    authUrl.searchParams.set('response_type', 'token');
    authUrl.searchParams.set('scope', META_SCOPE);
    authUrl.searchParams.set('state', String(Date.now()));
    return authUrl.toString();
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

async function fetchGoogleAdsReport(options = {}) {
    if (!hasGoogleConfig) {
        throw new Error('Google Ads 憑證未設定，請先更新 .env。');
    }

    const { datePreset = 'last_7d', since, until, customerId } = options;
    const dateRange = resolveDateRange({ datePreset, since, until });

    const normalizedCustomerId = cleanGoogleCustomerId(customerId || GOOGLE_ADS_CUSTOMER_ID);
    if (!/^\d{10}$/.test(normalizedCustomerId)) {
        throw new Error('GOOGLE_ADS_CUSTOMER_ID 必須是 10 位數字（可在前端輸入 123-456-7890）。');
    }

    const tokenResponse = await axios.post(
        'https://oauth2.googleapis.com/token',
        new URLSearchParams({
            client_id: GOOGLE_ADS_CLIENT_ID,
            client_secret: GOOGLE_ADS_CLIENT_SECRET,
            refresh_token: GOOGLE_ADS_REFRESH_TOKEN,
            grant_type: 'refresh_token'
        }),
        {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        }
    );

    const accessToken = tokenResponse.data?.access_token;
    if (!accessToken) {
        throw new Error('無法取得 Google OAuth access token，請檢查 Refresh Token 與 OAuth 憑證。');
    }

    const query = `
        SELECT
            campaign.name,
            campaign.app_campaign_setting.app_store,
            segments.date,
            metrics.impressions,
            metrics.clicks,
            metrics.conversions,
            metrics.cost_micros
        FROM campaign
        WHERE segments.date BETWEEN '${dateRange.since}' AND '${dateRange.until}'
        ORDER BY segments.date DESC
        LIMIT 1000
    `;

    const loginCustomerId = cleanGoogleCustomerId(GOOGLE_ADS_LOGIN_CUSTOMER_ID || '');
    if (loginCustomerId && !/^\d{10}$/.test(loginCustomerId)) {
        throw new Error('GOOGLE_ADS_LOGIN_CUSTOMER_ID 格式錯誤，需為 10 位數字。');
    }
    const headers = {
        Authorization: `Bearer ${accessToken}`,
        'developer-token': GOOGLE_ADS_DEVELOPER_TOKEN,
        'Content-Type': 'application/json'
    };
    if (loginCustomerId) {
        headers['login-customer-id'] = loginCustomerId;
    }

    const versionCandidates = Array.from(new Set([
        GOOGLE_ADS_API_VERSION,
        'v22',
        'v21',
        'v20'
    ].filter(Boolean)));

    let rows = null;
    let lastError = null;

    for (const version of versionCandidates) {
        const endpoint = `https://googleads.googleapis.com/${version}/customers/${normalizedCustomerId}/googleAds:search`;
        const collectedRows = [];
        let pageToken;

        try {
            do {
                const response = await axios.post(
                    endpoint,
                    {
                        query,
                        ...(pageToken ? { pageToken } : {})
                    },
                    { headers }
                );

                const resultRows = response.data?.results || [];
                collectedRows.push(...resultRows);
                pageToken = response.data?.nextPageToken;
            } while (pageToken);

            rows = collectedRows;
            break;
        } catch (error) {
            const status = error.response?.status;
            const message = error.response?.data?.error?.message || '';
            const isVersionIssue = status === 404 || status === 501 || /not implemented|not found/i.test(message);
            if (!isVersionIssue) {
                throw error;
            }
            lastError = error;
        }
    }

    if (!rows) {
        throw lastError || new Error('Google Ads API 版本不可用，請檢查 GOOGLE_ADS_API_VERSION。');
    }

    return rows.map((row) => {
        const conversions = Number(row.metrics?.conversions || 0);
        const appStore = String(row.campaign?.appCampaignSetting?.appStore || '').toUpperCase();
        let os = '';
        if (appStore.includes('GOOGLE_APP_STORE') || appStore.includes('GOOGLE_PLAY')) {
            os = 'Android';
        } else if (appStore.includes('APPLE_APP_STORE')) {
            os = 'iOS';
        }

        return {
            campaign_name: row.campaign?.name || '未命名',
            date_start: row.segments?.date || dateRange.since,
            date_stop: row.segments?.date || dateRange.until,
            spend: Number(row.metrics?.costMicros || row.metrics?.cost_micros || 0) / 1000000,
            currency: 'TWD',
            impressions: Number(row.metrics?.impressions || 0),
            clicks: Number(row.metrics?.clicks || 0),
            os,
            actions: [{ action_type: 'app_install', value: conversions }]
        };
    });
}

app.get('/api/insights', async (req, res) => {
    try {
        if (!hasMetaConfig()) {
            return res.status(400).json({ error: 'Meta 憑證未設定，請先更新 .env。' });
        }

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
        const metaError = error.response?.data?.error;
        if (metaError?.code === 190) {
            return res.status(401).json({
                error: 'Meta Access Token 已過期，請更新 ACCESS_TOKEN 後重試。',
                reauthUrl: buildMetaReauthUrl({ origin: getRequestOrigin(req) }),
                details: {
                    message: metaError.message,
                    code: metaError.code,
                    subcode: metaError.error_subcode
                }
            });
        }

        const errorMessage = error.response?.data || { message: error.message };
        return res.status(500).json({
            error: '抓取失敗',
            details: errorMessage
        });
    }
});

app.get('/api/google/insights', async (req, res) => {
    try {
        if (!hasGoogleConfig) {
            return res.status(400).json({ error: 'Google Ads 憑證未設定，請先更新 .env。' });
        }

        const { datePreset, since, until, accountId } = req.query;

        if ((since && !until) || (!since && until)) {
            return res.status(400).json({
                error: '請同時提供 since 與 until，或僅使用 datePreset。'
            });
        }

        if (since && until) {
            if (!isValidDateFormat(since) || !isValidDateFormat(until)) {
                return res.status(400).json({
                    error: '日期格式需為 YYYY-MM-DD。'
                });
            }
        }

        if (accountId && !isValidGoogleCustomerId(accountId)) {
            return res.status(400).json({
                error: 'Google accountId 格式錯誤，請填 123-456-7890 或 1234567890。'
            });
        }

        const data = await fetchGoogleAdsReport({
            datePreset: datePreset || 'last_7d',
            since,
            until,
            customerId: accountId || GOOGLE_ADS_CUSTOMER_ID
        });

        return res.json({ data });
    } catch (error) {
        const googleError = error.response?.data?.error;
        return res.status(500).json({
            error: '抓取 Google Ads 失敗',
            details: googleError ? `${googleError.code || ''} ${googleError.message || ''}`.trim() : error.message
        });
    }
});

app.get('/api/meta/reauth-url', (req, res) => {
    return res.json({
        reauthUrl: buildMetaReauthUrl({ origin: getRequestOrigin(req) })
    });
});

app.post('/api/meta/token', (req, res) => {
    try {
        const { accessToken } = req.body || {};
        const token = String(accessToken || '').trim();
        if (!token || token.length < 20) {
            return res.status(400).json({ error: 'accessToken 格式不正確。' });
        }

        persistAccessToken(token);
        return res.json({
            ok: true,
            message: 'ACCESS_TOKEN 已更新並寫入 .env'
        });
    } catch (error) {
        return res.status(500).json({
            error: '寫入 ACCESS_TOKEN 失敗',
            details: error.message
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