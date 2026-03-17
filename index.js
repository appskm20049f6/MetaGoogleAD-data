const axios = require('axios');
const express = require('express');
const path = require('path');
const fs = require('fs');
const { GoogleAdsApi } = require('google-ads-api');
const { parse } = require('csv-parse/sync');
const { open } = require('sqlite');
const sqlite3 = require('sqlite3');

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

const DB_DIR = path.join(__dirname, 'data');
const DB_FILE = path.join(DB_DIR, 'reports.db');

async function initDb() {
    fs.mkdirSync(DB_DIR, { recursive: true });
    const db = await open({
        filename: DB_FILE,
        driver: sqlite3.Database
    });

    await db.exec(`
        CREATE TABLE IF NOT EXISTS google_campaign_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_id TEXT NOT NULL,
            game_name TEXT NOT NULL DEFAULT '',
            report_date TEXT NOT NULL,
            campaign_name TEXT NOT NULL,
            impressions INTEGER NOT NULL DEFAULT 0,
            clicks INTEGER NOT NULL DEFAULT 0,
            spend_twd REAL NOT NULL DEFAULT 0,
            ctr REAL NOT NULL DEFAULT 0,
            installs REAL NOT NULL DEFAULT 0,
            conversions REAL NOT NULL DEFAULT 0,
            source_file TEXT NOT NULL,
            raw_status TEXT,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(account_id, report_date, campaign_name, source_file) ON CONFLICT REPLACE
        );

        CREATE INDEX IF NOT EXISTS idx_google_report_date ON google_campaign_reports(report_date);
        CREATE INDEX IF NOT EXISTS idx_google_account_date ON google_campaign_reports(account_id, report_date);
    `);

    const tableInfo = await db.all('PRAGMA table_info(google_campaign_reports)');
    const hasGameName = tableInfo.some((col) => col.name === 'game_name');
    if (!hasGameName) {
        await db.exec("ALTER TABLE google_campaign_reports ADD COLUMN game_name TEXT NOT NULL DEFAULT ''");
    }

    return db;
}

const dbPromise = initDb();

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

function parseLocalizedNumber(raw) {
    if (raw === null || raw === undefined) {
        return 0;
    }

    const text = String(raw).trim();
    if (!text || text === '--') {
        return 0;
    }

    const normalized = text.replace(/[,%"\s]/g, '').replace(/,/g, '');
    const value = Number(normalized);
    return Number.isFinite(value) ? value : 0;
}

function parseReportDateFromText(csvText) {
    const dateLine = csvText.split(/\r?\n/).find((line) => /\d{4}年\d{1,2}月\d{1,2}日/.test(line));
    if (!dateLine) {
        return null;
    }

    const match = dateLine.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (!match) {
        return null;
    }

    const year = match[1];
    const month = String(match[2]).padStart(2, '0');
    const day = String(match[3]).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function parseGoogleCampaignCsv(csvText) {
    const lines = csvText.split(/\r?\n/);
    const headerLine = lines.find((line) => line.includes('廣告活動狀態') && line.includes('廣告活動'));
    if (!headerLine) {
        throw new Error('找不到 Google 報表欄位標題列，請確認檔案格式。');
    }

    const delimiter = headerLine.includes('\t') ? '\t' : ',';
    const records = parse(csvText, {
        delimiter,
        skip_empty_lines: true,
        relax_quotes: true,
        relax_column_count: true,
        bom: true
    });

    const headerIndex = records.findIndex((row) => row.includes('廣告活動狀態') && row.includes('廣告活動'));
    if (headerIndex < 0) {
        throw new Error('無法解析 Google 報表標題列。');
    }

    const header = records[headerIndex].map((item) => String(item || '').trim());
    const rows = records.slice(headerIndex + 1).filter((row) => row.some((cell) => String(cell || '').trim() !== ''));
    const reportDate = parseReportDateFromText(csvText);

    const getValue = (row, columnName) => {
        const idx = header.indexOf(columnName);
        if (idx < 0) {
            return '';
        }
        return String(row[idx] || '').trim();
    };

    const normalized = [];
    rows.forEach((row) => {
        const firstCell = String(row[0] || '').trim();
        if (!firstCell || firstCell.startsWith('總計：')) {
            return;
        }

        const campaignName = getValue(row, '廣告活動');
        if (!campaignName || campaignName === '--') {
            return;
        }

        normalized.push({
            reportDate,
            campaignName,
            status: getValue(row, '廣告活動狀態'),
            impressions: parseLocalizedNumber(getValue(row, '曝光')),
            clicks: parseLocalizedNumber(getValue(row, '點擊')),
            spendTwd: parseLocalizedNumber(getValue(row, '費用')),
            ctr: parseLocalizedNumber(getValue(row, '點閱率')),
            installs: parseLocalizedNumber(getValue(row, '安裝')),
            conversions: parseLocalizedNumber(getValue(row, '轉換'))
        });
    });

    return normalized;
}

function decodeCsvPayload({ csvText, csvBase64 }) {
    if (csvBase64) {
        const buffer = Buffer.from(csvBase64, 'base64');
        if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
            return buffer.toString('utf16le');
        }
        if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
            const swapped = Buffer.allocUnsafe(buffer.length);
            for (let i = 0; i < buffer.length; i += 2) {
                swapped[i] = buffer[i + 1];
                swapped[i + 1] = buffer[i];
            }
            return swapped.toString('utf16le');
        }
        if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
            return buffer.toString('utf8');
        }
        return buffer.toString('utf8');
    }

    return typeof csvText === 'string' ? csvText : '';
}

async function importGoogleCsvToDb({ csvText, accountId, sourceFile, gameName }) {
    const parsedRows = parseGoogleCampaignCsv(csvText);
    if (!parsedRows.length) {
        throw new Error('報表沒有可匯入的廣告活動資料。');
    }

    const db = await dbPromise;
    const normalizedAccountId = cleanGoogleCustomerId(accountId || GOOGLE_ADS_CUSTOMER_ID || 'manual');
    const fileName = sourceFile || 'manual_upload.csv';
    const normalizedGameName = String(gameName || '').trim() || '未命名遊戲';

    await db.exec('BEGIN TRANSACTION');
    try {
        for (const row of parsedRows) {
            await db.run(
                `
                INSERT INTO google_campaign_reports (
                    account_id, game_name, report_date, campaign_name, impressions, clicks,
                    spend_twd, ctr, installs, conversions, source_file, raw_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                `,
                [
                    normalizedAccountId,
                    normalizedGameName,
                    row.reportDate || formatDate(new Date()),
                    row.campaignName,
                    row.impressions,
                    row.clicks,
                    row.spendTwd,
                    row.ctr,
                    row.installs,
                    row.conversions,
                    fileName,
                    row.status
                ]
            );
        }
        await db.exec('COMMIT');
    } catch (error) {
        await db.exec('ROLLBACK');
        throw error;
    }

    const reportDates = Array.from(new Set(parsedRows.map((item) => item.reportDate).filter(Boolean)));
    return {
        imported: parsedRows.length,
        reportDates
    };
}

async function fetchGoogleAdsReportFromDb(options = {}) {
    const { datePreset = 'last_7d', since, until, customerId, reportDate, gameName } = options;
    const dateRange = resolveDateRange({ datePreset, since, until });
    const db = await dbPromise;
    const normalizedCustomerId = cleanGoogleCustomerId(customerId || GOOGLE_ADS_CUSTOMER_ID || '');

    const params = [];
    let sql = `
        SELECT report_date, campaign_name, impressions, clicks, spend_twd, installs, conversions, game_name
        FROM google_campaign_reports
        WHERE 1=1
    `;

    if (reportDate) {
        sql += ' AND report_date = ?';
        params.push(reportDate);
    } else {
        sql += ' AND report_date BETWEEN ? AND ?';
        params.push(dateRange.since, dateRange.until);
    }

    if (normalizedCustomerId) {
        sql += ' AND account_id = ?';
        params.push(normalizedCustomerId);
    }

    if (gameName) {
        sql += ' AND game_name = ?';
        params.push(gameName);
    }

    sql += ' ORDER BY report_date DESC, game_name ASC, campaign_name ASC';

    const rows = await db.all(sql, params);
    return rows.map((row) => {
        const installs = Number(row.installs || 0);
        const conversions = Number(row.conversions || 0);
        return {
            campaign_name: row.campaign_name || '未命名',
            date_start: row.report_date,
            date_stop: row.report_date,
            spend: Number(row.spend_twd || 0),
            impressions: Number(row.impressions || 0),
            clicks: Number(row.clicks || 0),
            currency: 'TWD',
            game_name: row.game_name || '',
            actions: [{ action_type: 'app_install', value: installs || conversions || 0 }]
        };
    });
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

    const api = new GoogleAdsApi({
        client_id: GOOGLE_ADS_CLIENT_ID,
        client_secret: GOOGLE_ADS_CLIENT_SECRET,
        developer_token: GOOGLE_ADS_DEVELOPER_TOKEN
    });

    const normalizedCustomerId = cleanGoogleCustomerId(customerId || GOOGLE_ADS_CUSTOMER_ID);

    const customer = api.Customer({
        customer_id: normalizedCustomerId,
        refresh_token: GOOGLE_ADS_REFRESH_TOKEN,
        login_customer_id: cleanGoogleCustomerId(GOOGLE_ADS_LOGIN_CUSTOMER_ID || '') || undefined
    });

    const rows = await customer.query(`
        SELECT
            campaign.name,
            segments.date,
            metrics.impressions,
            metrics.clicks,
            metrics.conversions,
            metrics.cost_micros
        FROM campaign
        WHERE segments.date BETWEEN '${dateRange.since}' AND '${dateRange.until}'
        ORDER BY segments.date DESC
        LIMIT 1000
    `);

    return rows.map((row) => {
        const conversions = Number(row.metrics?.conversions || 0);
        return {
            campaign_name: row.campaign?.name || '未命名',
            date_start: row.segments?.date || dateRange.since,
            date_stop: row.segments?.date || dateRange.until,
            spend: Number(row.metrics?.cost_micros || 0) / 1000000,
            impressions: Number(row.metrics?.impressions || 0),
            clicks: Number(row.metrics?.clicks || 0),
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
        const { datePreset, since, until, accountId, source, reportDate, gameName } = req.query;

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

        if (reportDate && !isValidDateFormat(reportDate)) {
            return res.status(400).json({
                error: 'reportDate 格式需為 YYYY-MM-DD。'
            });
        }

        if (accountId && !isValidGoogleCustomerId(accountId)) {
            return res.status(400).json({
                error: 'Google accountId 格式錯誤，請填 123-456-7890 或 1234567890。'
            });
        }

        const shouldUseDb = source === 'db' || !hasGoogleConfig;
        const data = shouldUseDb
            ? await fetchGoogleAdsReportFromDb({
                datePreset: datePreset || 'last_7d',
                since,
                until,
                customerId: accountId || GOOGLE_ADS_CUSTOMER_ID,
                reportDate,
                gameName
            })
            : await fetchGoogleAdsReport({
                datePreset: datePreset || 'last_7d',
                since,
                until,
                customerId: accountId || GOOGLE_ADS_CUSTOMER_ID
            });

        return res.json({ data });
    } catch (error) {
        return res.status(500).json({
            error: '抓取 Google Ads 失敗',
            details: error.message
        });
    }
});

app.post('/api/google/import-csv', async (req, res) => {
    try {
        const { csvText, csvBase64, accountId, fileName, gameName } = req.body || {};
        const decodedCsvText = decodeCsvPayload({ csvText, csvBase64 });
        if (!decodedCsvText) {
            return res.status(400).json({ error: '請提供 csvText 或 csvBase64 內容。' });
        }

        if (accountId && !isValidGoogleCustomerId(accountId)) {
            return res.status(400).json({
                error: 'Google accountId 格式錯誤，請填 123-456-7890 或 1234567890。'
            });
        }

        const result = await importGoogleCsvToDb({
            csvText: decodedCsvText,
            accountId,
            sourceFile: fileName,
            gameName
        });

        return res.json({
            ok: true,
            imported: result.imported,
            reportDates: result.reportDates
        });
    } catch (error) {
        return res.status(500).json({
            error: '匯入 Google CSV 失敗',
            details: error.message
        });
    }
});

app.get('/api/google/db-filters', async (req, res) => {
    try {
        const { accountId } = req.query;
        if (accountId && !isValidGoogleCustomerId(accountId)) {
            return res.status(400).json({
                error: 'Google accountId 格式錯誤，請填 123-456-7890 或 1234567890。'
            });
        }

        const db = await dbPromise;
        const normalizedCustomerId = cleanGoogleCustomerId(accountId || GOOGLE_ADS_CUSTOMER_ID || '');
        const params = [];
        let baseWhere = 'WHERE 1=1';
        if (normalizedCustomerId) {
            baseWhere += ' AND account_id = ?';
            params.push(normalizedCustomerId);
        }

        const dateRows = await db.all(
            `SELECT DISTINCT report_date FROM google_campaign_reports ${baseWhere} ORDER BY report_date DESC`,
            params
        );
        const gameRows = await db.all(
            `SELECT DISTINCT game_name FROM google_campaign_reports ${baseWhere} ORDER BY game_name ASC`,
            params
        );

        return res.json({
            reportDates: dateRows.map((row) => row.report_date).filter(Boolean),
            gameNames: gameRows.map((row) => row.game_name).filter(Boolean)
        });
    } catch (error) {
        return res.status(500).json({
            error: '抓取 Google DB 篩選條件失敗',
            details: error.message
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