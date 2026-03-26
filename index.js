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
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;
const SHARE_USERNAME = process.env.SHARE_USERNAME;
const SHARE_PASSWORD = process.env.SHARE_PASSWORD;
const ENV_FILE = path.join(__dirname, '.env');
const DATA_DIR = path.join(__dirname, 'data');
const APPSFLYER_DIR = path.join(DATA_DIR, 'appsflyer');
const ACCOUNT_PRESETS_FILE = path.join(DATA_DIR, 'account-presets.json');
const MANUAL_ENTRIES_FILE = path.join(DATA_DIR, 'manual-entries.json');

function readManualEntriesFile() {
    try {
        if (!fs.existsSync(MANUAL_ENTRIES_FILE)) return [];
        return JSON.parse(fs.readFileSync(MANUAL_ENTRIES_FILE, 'utf8'));
    } catch (e) { return []; }
}
function writeManualEntriesFile(entries) {
    fs.writeFileSync(MANUAL_ENTRIES_FILE, JSON.stringify(entries, null, 2), 'utf8');
}

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

function shouldBypassShareAuth(reqPath) {
    return reqPath === '/health'
        || reqPath === '/meta-token-callback.html'
        || reqPath === '/api/meta/token'
        || reqPath === '/api/meta/reauth-url';
}

function shareAuthMiddleware(req, res, next) {
    if (!SHARE_USERNAME || !SHARE_PASSWORD) {
        return next();
    }

    if (shouldBypassShareAuth(req.path)) {
        return next();
    }

    const authHeader = String(req.headers.authorization || '');
    if (!authHeader.startsWith('Basic ')) {
        res.set('WWW-Authenticate', 'Basic realm="MetaGoogleAD"');
        return res.status(401).send('Authentication required');
    }

    const encoded = authHeader.slice(6);
    let decoded = '';
    try {
        decoded = Buffer.from(encoded, 'base64').toString('utf8');
    } catch (_) {
        res.set('WWW-Authenticate', 'Basic realm="MetaGoogleAD"');
        return res.status(401).send('Invalid authentication');
    }

    const idx = decoded.indexOf(':');
    const username = idx >= 0 ? decoded.slice(0, idx) : '';
    const password = idx >= 0 ? decoded.slice(idx + 1) : '';

    if (username !== SHARE_USERNAME || password !== SHARE_PASSWORD) {
        res.set('WWW-Authenticate', 'Basic realm="MetaGoogleAD"');
        return res.status(401).send('Invalid credentials');
    }

    return next();
}

app.use(shareAuthMiddleware);
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

function ensureDir(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
}

function normalizeAccountPresetsPayload(payload = {}) {
    const selectedPlatform = payload.selectedPlatform === 'google' ? 'google' : 'meta';
    const selectedAccountIdByPlatform = {
        meta: String(payload.selectedAccountIdByPlatform?.meta || '').trim(),
        google: String(payload.selectedAccountIdByPlatform?.google || '').trim()
    };

    const accounts = Array.isArray(payload.accounts)
        ? payload.accounts
            .map((item) => ({
                name: String(item?.name || '').trim(),
                accountId: String(item?.accountId || '').trim(),
                platform: item?.platform === 'google' ? 'google' : 'meta'
            }))
            .filter((item) => item.name && item.accountId)
        : [];

    return {
        selectedPlatform,
        selectedAccountIdByPlatform,
        accounts
    };
}

function readAccountPresetsFile() {
    try {
        if (!fs.existsSync(ACCOUNT_PRESETS_FILE)) {
            return normalizeAccountPresetsPayload({});
        }
        const raw = fs.readFileSync(ACCOUNT_PRESETS_FILE, 'utf8');
        const parsed = JSON.parse(raw);
        return normalizeAccountPresetsPayload(parsed);
    } catch (_) {
        return normalizeAccountPresetsPayload({});
    }
}

function writeAccountPresetsFile(payload) {
    ensureDir(DATA_DIR);
    fs.writeFileSync(ACCOUNT_PRESETS_FILE, JSON.stringify(payload, null, 2), 'utf8');
}

function parseCsvText(csvText) {
    const rows = [];
    let current = '';
    let row = [];
    let inQuotes = false;

    for (let i = 0; i < csvText.length; i += 1) {
        const char = csvText[i];
        const next = csvText[i + 1];

        if (char === '"') {
            if (inQuotes && next === '"') {
                current += '"';
                i += 1;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }

        if (char === ',' && !inQuotes) {
            row.push(current);
            current = '';
            continue;
        }

        if ((char === '\n' || char === '\r') && !inQuotes) {
            if (char === '\r' && next === '\n') {
                i += 1;
            }

            row.push(current);
            rows.push(row);
            row = [];
            current = '';
            continue;
        }

        current += char;
    }

    if (current.length > 0 || row.length > 0) {
        row.push(current);
        rows.push(row);
    }

    return rows;
}

function toNumber(value) {
    const text = String(value ?? '').trim();
    if (!text) {
        return 0;
    }

    const normalized = text.replace(/,/g, '');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeAppsflyerRows(headers, rows) {
    const headerIndex = new Map();
    headers.forEach((item, idx) => {
        headerIndex.set(String(item || '').trim(), idx);
    });

    const getValue = (row, key) => {
        const idx = headerIndex.get(key);
        return idx === undefined ? '' : String(row[idx] ?? '').trim();
    };

    return rows
        .filter((row) => row.some((item) => String(item || '').trim() !== ''))
        .map((row) => {
            const impressions = toNumber(getValue(row, 'Impressions'));
            const clicks = toNumber(getValue(row, 'Clicks'));
            const totalAttributions = toNumber(getValue(row, 'Total attributions appsflyer'));
            const activated = toNumber(getValue(row, 'Installs appsflyer'));
            const reattributions = toNumber(getValue(row, 'Re-attributions appsflyer'));
            const reengagements = toNumber(getValue(row, 'Re-engagements appsflyer'));
            
            // 嘗試提取自然歸因和非自然歸因
            const organicInstalls = toNumber(getValue(row, 'Organic installs appsflyer')) || 
                                    toNumber(getValue(row, 'Organic Installs appsflyer')) || 
                                    toNumber(getValue(row, 'Organic'));
            const nonOrganicInstalls = toNumber(getValue(row, 'Non-organic installs appsflyer')) || 
                                       toNumber(getValue(row, 'Non-organic Installs appsflyer')) || 
                                       toNumber(getValue(row, 'Non-organic'));

            // 提取 OS/Platform 欄位
            const os = getValue(row, 'OS') || getValue(row, 'Platform') || getValue(row, 'Operating system');

            return {
                os,
                mediaSource: getValue(row, 'Media source') || 'None',
                campaign: getValue(row, 'Campaign') || 'None',
                impressions,
                clicks,
                totalAttributions,
                activated,
                reattributions,
                reengagements,
                afTotal: activated + reattributions + reengagements,
                cost: toNumber(getValue(row, 'Cost')),
                ecpi: toNumber(getValue(row, 'eCPI appsflyer')),
                organicInstalls,
                nonOrganicInstalls
            };
        });
}

function summarizeAppsflyerRows(rows) {
    const totalCost = rows.reduce((sum, row) => sum + Number(row.cost || 0), 0);
    const totalActivated = rows.reduce((sum, row) => sum + Number(row.activated || 0), 0);
    const totalClicks = rows.reduce((sum, row) => sum + Number(row.clicks || 0), 0);
    const totalAttributions = rows.reduce((sum, row) => sum + Number(row.totalAttributions || 0), 0);
    const totalAf = rows.reduce((sum, row) => sum + Number(row.afTotal || 0), 0);
    return {
        rowCount: rows.length,
        totalCost,
        totalActivated,
        totalClicks,
        totalAttributions,
        totalAf
    };
}

function extractDateRangeFromFileName(fileName) {
    const text = String(fileName || '');
    const match = text.match(/([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})\s*-\s*([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})/);
    if (!match) {
        return null;
    }

    const start = new Date(match[1]);
    const end = new Date(match[2]);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
        return null;
    }

    return {
        since: formatDate(start),
        until: formatDate(end)
    };
}

function overlapDateRange(aSince, aUntil, bSince, bUntil) {
    return aSince <= bUntil && aUntil >= bSince;
}

function isContainedDateRange(innerSince, innerUntil, outerSince, outerUntil) {
    return innerSince >= outerSince && innerUntil <= outerUntil;
}

function parseIsoDate(dateText) {
    if (!isValidDateFormat(dateText)) {
        return null;
    }

    const [year, month, day] = dateText.split('-').map((value) => Number(value));
    const date = new Date(year, month - 1, day);
    return Number.isNaN(date.getTime()) ? null : date;
}

function getInclusiveDateCount(since, until) {
    const start = parseIsoDate(since);
    const end = parseIsoDate(until);
    if (!start || !end || start > end) {
        return 0;
    }

    const diff = end.getTime() - start.getTime();
    return Math.floor(diff / 86400000) + 1;
}

function getFileTimestampKey(fileName) {
    const text = String(fileName || '');
    const match = text.match(/^(\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z)/);
    return match ? match[1] : text;
}

function compareFileFreshness(left, right) {
    return getFileTimestampKey(right.fileName).localeCompare(getFileTimestampKey(left.fileName));
}

function buildAppsflyerFileRecords(targetRange) {
    ensureDir(APPSFLYER_DIR);

    const jsonFiles = fs.readdirSync(APPSFLYER_DIR)
        .filter((file) => file.toLowerCase().endsWith('.json'));

    const deduped = new Map();
    jsonFiles.forEach((fileName) => {
        const filePath = path.join(APPSFLYER_DIR, fileName);
        const content = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        const rows = Array.isArray(content.rows) ? content.rows : [];
        const metaSince = String(content.meta?.reportSince || content.meta?.since || '').trim();
        const metaUntil = String(content.meta?.reportUntil || content.meta?.until || '').trim();
        const detectedOS = String(content.meta?.detectedOS || 'Unknown').trim() || 'Unknown';
        const recordSince = isValidDateFormat(metaSince) ? metaSince : targetRange.since;
        const recordUntil = isValidDateFormat(metaUntil) ? metaUntil : recordSince;

        const dedupeKey = `${recordSince}|${recordUntil}|${detectedOS.toLowerCase()}`;
        const record = {
            fileName,
            filePath,
            rows,
            detectedOS,
            since: recordSince,
            until: recordUntil,
            coveredDays: getInclusiveDateCount(recordSince, recordUntil)
        };

        const existing = deduped.get(dedupeKey);
        if (!existing || compareFileFreshness(record, existing) < 0) {
            deduped.set(dedupeKey, record);
        }
    });

    return {
        totalFileCount: jsonFiles.length,
        records: Array.from(deduped.values())
    };
}

function chooseBestNonOverlappingAppsflyerRecords(records) {
    const sorted = records
        .filter((record) => record.coveredDays > 0)
        .sort((left, right) => {
            if (left.until !== right.until) {
                return left.until.localeCompare(right.until);
            }
            if (left.since !== right.since) {
                return left.since.localeCompare(right.since);
            }
            return compareFileFreshness(left, right);
        });

    if (!sorted.length) {
        return [];
    }

    const previousCompatible = sorted.map((current, index) => {
        for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
            if (sorted[cursor].until < current.since) {
                return cursor;
            }
        }
        return -1;
    });

    const best = new Array(sorted.length);

    const comparePlan = (left, right) => {
        if (!right) return left;
        if (!left) return right;

        if (left.coveredDays !== right.coveredDays) {
            return left.coveredDays > right.coveredDays ? left : right;
        }
        if (left.fileCount !== right.fileCount) {
            return left.fileCount < right.fileCount ? left : right;
        }
        return left.freshnessKey >= right.freshnessKey ? left : right;
    };

    const buildPlan = (basePlan, record) => ({
        coveredDays: (basePlan?.coveredDays || 0) + record.coveredDays,
        fileCount: (basePlan?.fileCount || 0) + 1,
        freshnessKey: [basePlan?.freshnessKey || '', getFileTimestampKey(record.fileName)].sort().join('|'),
        items: [...(basePlan?.items || []), record]
    });

    sorted.forEach((record, index) => {
        const includeBase = previousCompatible[index] >= 0 ? best[previousCompatible[index]] : null;
        const includePlan = buildPlan(includeBase, record);
        const excludePlan = index > 0 ? best[index - 1] : null;
        best[index] = comparePlan(includePlan, excludePlan);
    });

    return best[best.length - 1]?.items || [];
}

function detectOSFromRows(rows) {
    if (!rows || !rows.length) return 'Unknown';
    
    const osSet = new Set();
    rows.forEach(row => {
        if (row.os && row.os.trim()) {
            osSet.add(row.os.toLowerCase());
        }
    });
    
    if (osSet.size === 0) return 'Unknown';
    if (osSet.size === 1) return Array.from(osSet)[0];
    
    // 混合多個OS，判斷主要OS
    const osArray = Array.from(osSet);
    if (osArray.includes('android') && osArray.includes('ios')) return 'Mixed';
    if (osArray.includes('android')) return 'Android';
    if (osArray.includes('ios')) return 'iOS';
    return 'Unknown';
}

function removeOldDuplicateFiles(newReportSince, newReportUntil, newOS) {
    try {
        ensureDir(APPSFLYER_DIR);
        const files = fs.readdirSync(APPSFLYER_DIR)
            .filter((file) => file.toLowerCase().endsWith('.json'))
            .map((file) => ({
                name: file,
                path: path.join(APPSFLYER_DIR, file)
            }));
        
        files.forEach(({ name, path: filePath }) => {
            try {
                const content = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                const existingSince = content.meta?.reportSince;
                const existingUntil = content.meta?.reportUntil;
                const existingOS = content.meta?.detectedOS || 'Unknown';
                
                // 如果日期範圍完全相同且 OS 相同，刪除舊檔案
                if (
                    existingSince === newReportSince &&
                    existingUntil === newReportUntil &&
                    existingOS.toLowerCase() === newOS.toLowerCase()
                ) {
                    const csvFile = name.replace(/\.json$/, '.csv');
                    const csvPath = path.join(APPSFLYER_DIR, csvFile);
                    
                    // 刪除舊的 JSON 和 CSV
                    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
                    if (fs.existsSync(csvPath)) fs.unlinkSync(csvPath);
                    
                    console.log(`✓ 已移除重複檔案: ${name}`);
                }
            } catch (err) {
                console.error(`讀取檔案 ${name} 失敗:`, err.message);
            }
        });
    } catch (error) {
        console.error('移除重複檔案失敗:', error.message);
    }
}

function buildAppsflyerInsights(options = {}) {
    const { datePreset = 'last_7d', since, until } = options;
    const targetRange = resolveDateRange({ datePreset, since, until });

    const { totalFileCount, records } = buildAppsflyerFileRecords(targetRange);
    const containedRecords = records.filter((record) => {
        return isContainedDateRange(record.since, record.until, targetRange.since, targetRange.until);
    });
    const selectedRecords = chooseBestNonOverlappingAppsflyerRecords(containedRecords);
    const matchedFileCount = selectedRecords.length;
    const coveredDays = selectedRecords.reduce((sum, record) => sum + record.coveredDays, 0);
    const targetDays = getInclusiveDateCount(targetRange.since, targetRange.until);

    const grouped = new Map();
    selectedRecords.forEach((record) => {
        record.rows.forEach((row) => {
            const mediaSource = String(row.mediaSource || 'None');
            const campaign = String(row.campaign || 'None');
            const key = `${mediaSource}|${campaign}`;
            const current = grouped.get(key) || {
                mediaSource,
                campaign,
                impressions: 0,
                clicks: 0,
                totalAttributions: 0,
                afTotal: 0,
                activated: 0,
                reattributions: 0,
                reengagements: 0
            };

            const impressions = Number(row.impressions || 0);
            const clicks = Number(row.clicks || 0);
            const totalAttributions = Number(row.totalAttributions || 0);
            const activated = Number(row.activated || row.installs || 0);
            const reattributions = Number(row.reattributions || 0);
            const reengagements = Number(row.reengagements || 0);
            const afTotal = Number(row.afTotal || (activated + reattributions + reengagements));

            current.impressions += impressions;
            current.clicks += clicks;
            current.totalAttributions += totalAttributions;
            current.afTotal += afTotal;
            current.activated += activated;
            current.reattributions += reattributions;
            current.reengagements += reengagements;
            grouped.set(key, current);
        });
    });

    const rows = Array.from(grouped.values()).sort((a, b) => {
        if (a.mediaSource !== b.mediaSource) {
            return a.mediaSource.localeCompare(b.mediaSource, 'zh-Hant');
        }

        return a.campaign.localeCompare(b.campaign, 'zh-Hant');
    });

    return {
        rows,
        totalFileCount,
        matchedFileCount,
        coverage: {
            requestedSince: targetRange.since,
            requestedUntil: targetRange.until,
            requestedDays: targetDays,
            coveredDays,
            isComplete: targetDays > 0 && coveredDays >= targetDays,
            selectedFiles: selectedRecords.map((record) => ({
                fileName: record.fileName,
                since: record.since,
                until: record.until,
                detectedOS: record.detectedOS
            }))
        }
    };
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
        const responseData = error.response?.data || {};
        const googleError = responseData.error;
        const oauthDescription = responseData.error_description;

        let details = error.message;
        if (typeof googleError === 'string') {
            details = oauthDescription
                ? `${googleError}: ${oauthDescription}`
                : googleError;
        } else if (googleError && typeof googleError === 'object') {
            details = `${googleError.code || ''} ${googleError.message || ''}`.trim() || error.message;
        }

        const isInvalidGrant = typeof googleError === 'string' && googleError === 'invalid_grant';
        return res.status(isInvalidGrant ? 401 : 500).json({
            error: '抓取 Google Ads 失敗',
            details
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

app.post('/api/appsflyer/import', (req, res) => {
    try {
        const { fileName, csvText } = req.body || {};
        const safeFileName = String(fileName || '').trim();
        const rawCsvText = String(csvText || '');

        if (!safeFileName.toLowerCase().endsWith('.csv')) {
            return res.status(400).json({ error: '請上傳 .csv 檔案。' });
        }

        if (!rawCsvText.trim()) {
            return res.status(400).json({ error: 'CSV 內容為空。' });
        }

        const parsed = parseCsvText(rawCsvText);
        if (!parsed.length) {
            return res.status(400).json({ error: 'CSV 解析失敗，找不到資料列。' });
        }

        const headers = parsed[0].map((item) => String(item || '').trim());
        const bodyRows = parsed.slice(1);
        const normalizedRows = normalizeAppsflyerRows(headers, bodyRows);
        const summary = summarizeAppsflyerRows(normalizedRows);
        const parsedRange = extractDateRangeFromFileName(safeFileName);
        const fallbackDate = formatDate(new Date());
        const reportSince = parsedRange?.since || fallbackDate;
        const reportUntil = parsedRange?.until || reportSince;
        
        // 自動偵測 OS
        const detectedOS = detectOSFromRows(normalizedRows);
        
        // 檢查並移除相同日期+OS 的舊檔案（保留最新的）
        removeOldDuplicateFiles(reportSince, reportUntil, detectedOS);

        ensureDir(APPSFLYER_DIR);

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseName = safeFileName.replace(/\.csv$/i, '').replace(/[^a-zA-Z0-9_\-\u4e00-\u9fa5]/g, '_');
        const osTag = detectedOS === 'Mixed' ? '_mixed' : detectedOS === 'Unknown' ? '_unknown' : `_${detectedOS.toLowerCase()}`;
        const savedCsvFile = `${timestamp}${osTag}_${baseName}.csv`;
        const savedJsonFile = `${timestamp}${osTag}_${baseName}.json`;

        const savedCsvPath = path.join(APPSFLYER_DIR, savedCsvFile);
        const savedJsonPath = path.join(APPSFLYER_DIR, savedJsonFile);

        fs.writeFileSync(savedCsvPath, rawCsvText, 'utf8');
        fs.writeFileSync(savedJsonPath, JSON.stringify({
            meta: {
                originalFileName: safeFileName,
                createdAt: new Date().toLocaleString('zh-TW', { hour12: false }),
                reportSince,
                reportUntil,
                detectedOS
            },
            headers,
            summary,
            rows: normalizedRows
        }, null, 2), 'utf8');

        return res.json({
            ok: true,
            savedCsvFile,
            savedJsonFile,
            summary,
            detectedOS,
            previewRows: normalizedRows.slice(0, 300)
        });
    } catch (error) {
        return res.status(500).json({
            error: 'Appsflyer 匯入建檔失敗',
            details: error.message
        });
    }
});

app.get('/api/appsflyer/files', (req, res) => {
    try {
        ensureDir(APPSFLYER_DIR);
        const files = fs.readdirSync(APPSFLYER_DIR)
            .filter((file) => file.toLowerCase().endsWith('.json'))
            .sort((a, b) => b.localeCompare(a))
            .slice(0, 100)
            .map((file) => {
                const filePath = path.join(APPSFLYER_DIR, file);
                const content = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                return {
                    fileName: file,
                    createdAt: content.meta?.createdAt || '-',
                    reportSince: content.meta?.reportSince || '-',
                    reportUntil: content.meta?.reportUntil || '-',
                    rowCount: content.summary?.rowCount || 0,
                    totalCost: content.summary?.totalCost || 0,
                    totalInstalls: content.summary?.totalInstalls || 0
                };
            });

        return res.json({ files });
    } catch (error) {
        return res.status(500).json({
            error: '讀取 Appsflyer 建檔清單失敗',
            details: error.message
        });
    }
});

app.get('/api/appsflyer/insights', (req, res) => {
    try {
        const { datePreset, since, until } = req.query;

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

        const result = buildAppsflyerInsights({
            datePreset: String(datePreset || 'last_7d'),
            since: String(since || ''),
            until: String(until || '')
        });

        return res.json({
            data: result.rows,
            meta: {
                totalFileCount: result.totalFileCount,
                matchedFileCount: result.matchedFileCount,
                coverage: result.coverage
            }
        });
    } catch (error) {
        return res.status(500).json({
            error: '查詢 Appsflyer 失敗',
            details: error.message
        });
    }
});

app.get('/api/sheet/installs', async (req, res) => {
    try {
        if (!APPS_SCRIPT_URL) {
            return res.status(400).json({ error: 'APPS_SCRIPT_URL 未設定，請更新 .env。' });
        }

        const { game, startDate, endDate } = req.query;

        if (game && !/^[\w\u4e00-\u9fff\-_ ]{1,50}$/.test(game)) {
            return res.status(400).json({ error: '請提供有效的 game 參數（工作表名稱，1-50字）。' });
        }
        if (startDate && !/^\d{4}-\d{2}-\d{2}$/.test(startDate)) {
            return res.status(400).json({ error: 'startDate 格式需為 YYYY-MM-DD。' });
        }
        if (endDate && !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
            return res.status(400).json({ error: 'endDate 格式需為 YYYY-MM-DD。' });
        }

        const params = new URLSearchParams();
        if (game) params.append('game', game);
        if (startDate) params.append('startDate', startDate);
        if (endDate) params.append('endDate', endDate);

        const url = `${APPS_SCRIPT_URL}?${params.toString()}`;
        const response = await axios.get(url, { timeout: 20000, maxRedirects: 5 });
        return res.json(response.data);
    } catch (error) {
        return res.status(500).json({
            error: '取得 Sheet 資料失敗',
            details: error.message
        });
    }
});

app.get('/api/account-presets', (req, res) => {
    const data = readAccountPresetsFile();
    return res.json(data);
});

app.post('/api/account-presets', (req, res) => {
    try {
        const normalized = normalizeAccountPresetsPayload(req.body || {});
        writeAccountPresetsFile(normalized);
        return res.json({ ok: true, saved: normalized });
    } catch (error) {
        return res.status(500).json({ error: `儲存失敗：${error.message}` });
    }
});

app.get('/api/manual-entries', (req, res) => {
    res.json(readManualEntriesFile());
});

app.post('/api/manual-entries', (req, res) => {
    try {
        const entry = req.body;
        if (!entry || !entry.id || !entry.name) return res.status(400).json({ error: '缺少必要欄位' });
        const entries = readManualEntriesFile();
        const idx = entries.findIndex(e => e.id === entry.id);
        if (idx >= 0) entries[idx] = entry; else entries.push(entry);
        writeManualEntriesFile(entries);
        res.json({ ok: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/manual-entries/:id', (req, res) => {
    try {
        const { id } = req.params;
        const entries = readManualEntriesFile().filter(e => e.id !== id);
        writeManualEntriesFile(entries);
        res.json({ ok: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
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