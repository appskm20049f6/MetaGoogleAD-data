const tabMetaEl = document.getElementById('tabMeta');
const tabGoogleEl = document.getElementById('tabGoogle');
const tabIntegratedEl = document.getElementById('tabIntegrated');
const tabIntegrated2El = document.getElementById('tabIntegrated2');
const queryPanelEl = document.getElementById('queryPanel');
const integratedPanelEl = document.getElementById('integratedPanel');
const integratedSheetPanelEl = document.getElementById('integratedSheetPanel');
const modeEl = document.getElementById('mode');
const datePresetEl = document.getElementById('datePreset');
const sinceEl = document.getElementById('since');
const untilEl = document.getElementById('until');
const levelEl = document.getElementById('level');
const fxRateEl = document.getElementById('fxRate');
const accountSelectEl = document.getElementById('accountSelect');
const accountSelectLabelEl = document.getElementById('accountSelectLabel');
const pageNameEl = document.getElementById('pageName');
const pageNameLabelEl = document.getElementById('pageNameLabel');
const pageAccountIdEl = document.getElementById('pageAccountId');
const pageAccountIdLabelEl = document.getElementById('pageAccountIdLabel');
const addPageBtn = document.getElementById('addPageBtn');
const removePageBtn = document.getElementById('removePageBtn');
const queryBtn = document.getElementById('queryBtn');
const checkAllEl = document.getElementById('checkAll');
const statusEl = document.getElementById('status');
const summaryEl = document.getElementById('summary');
const tableWrapEl = document.getElementById('tableWrap');
const tbodyEl = document.getElementById('tbody');
const integratedStatusEl = document.getElementById('integratedStatus');
const integratedMatrixWrapEl = document.getElementById('integratedMatrixWrap');
const integratedMatrixTbodyEl = document.getElementById('integratedMatrixTbody');
const integratedTableWrapEl = document.getElementById('integratedTableWrap');
const integratedTbodyEl = document.getElementById('integratedTbody');
const integratedSheetStatusEl = document.getElementById('integratedSheetStatus');
const integratedSheetDateEl = document.getElementById('integratedSheetDate');
const integratedSheetSourceWrapEl = document.getElementById('integratedSheetSourceWrap');
const integratedSheetSourceTbodyEl = document.getElementById('integratedSheetSourceTbody');
const integratedSheetSummaryWrapEl = document.getElementById('integratedSheetSummaryWrap');
const integratedSheetSummaryTbodyEl = document.getElementById('integratedSheetSummaryTbody');
const integratedSheetTitleEl = document.getElementById('integratedSheetTitle');
const integratedSheetSaveBtnEl = document.getElementById('integratedSheetSaveBtn');
const integratedSheetNotesEl = document.getElementById('integratedSheetNotes');

const sumImpressionsEl = document.getElementById('sumImpressions');
const sumClicksEl = document.getElementById('sumClicks');
const sumCtrEl = document.getElementById('sumCtr');
const sumResultsEl = document.getElementById('sumResults');
const avgCpiEl = document.getElementById('avgCpi');
const sumSpendEl = document.getElementById('sumSpend');
const integratedMetaSpendEl = document.getElementById('integratedMetaSpend');
const integratedGoogleSpendEl = document.getElementById('integratedGoogleSpend');
const integratedAndroidSpendEl = document.getElementById('integratedAndroidSpend');
const integratedIosSpendEl = document.getElementById('integratedIosSpend');
const integratedWebBookingSpendEl = document.getElementById('integratedWebBookingSpend');
const integratedUnknownSpendEl = document.getElementById('integratedUnknownSpend');
const integratedTotalSpendEl = document.getElementById('integratedTotalSpend');

const nf = new Intl.NumberFormat('zh-TW');
const twdFormatter = new Intl.NumberFormat('zh-TW', { style: 'currency', currency: 'TWD', maximumFractionDigits: 0 });
const SETTINGS_KEY = 'meta_dashboard_settings_v1';
const SHEET_SNAPSHOTS_KEY = 'integrated_sheet_snapshots_v1';
const defaultSettings = {
  fxRate: 32,
  selectedPlatform: 'meta',
  selectedAccountIdByPlatform: {
    meta: '',
    google: ''
  },
  accounts: []
};

let settings = loadSettings();
let activeTab = settings.selectedPlatform === 'google' ? 'google' : 'meta';
let currentRows = [];
let selectedRowIds = new Set();
let metaAuthPopup = null;
const platformRows = {
  meta: [],
  google: []
};
let integratedSheetSnapshots = loadIntegratedSheetSnapshots();

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) {
      return { ...defaultSettings };
    }

    const parsed = JSON.parse(raw);
    const migratedSelectedMeta = parsed.selectedAccountId || '';
    return {
      fxRate: Number(parsed.fxRate) > 0 ? Number(parsed.fxRate) : defaultSettings.fxRate,
      selectedPlatform: parsed.selectedPlatform || 'meta',
      selectedAccountIdByPlatform: {
        meta: parsed.selectedAccountIdByPlatform?.meta || migratedSelectedMeta,
        google: parsed.selectedAccountIdByPlatform?.google || ''
      },
      accounts: Array.isArray(parsed.accounts) ? parsed.accounts : []
    };
  } catch (error) {
    return { ...defaultSettings };
  }
}

function saveSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

function loadIntegratedSheetSnapshots() {
  try {
    const raw = localStorage.getItem(SHEET_SNAPSHOTS_KEY);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function saveIntegratedSheetSnapshots() {
  localStorage.setItem(SHEET_SNAPSHOTS_KEY, JSON.stringify(integratedSheetSnapshots));
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getCurrentPlatform() {
  return activeTab === 'google' ? 'google' : 'meta';
}

function getPlatformText(platform) {
  return platform === 'google' ? 'Google' : 'Meta';
}

function isValidAccountId(text) {
  if (getCurrentPlatform() === 'google') {
    return /^\d{3}-?\d{3}-?\d{4}$/.test(text);
  }
  return /^act_\d+$/.test(text);
}

function getCurrentSelectedAccountId() {
  return settings.selectedAccountIdByPlatform?.[getCurrentPlatform()] || '';
}

function setCurrentSelectedAccountId(accountId) {
  settings.selectedAccountIdByPlatform[getCurrentPlatform()] = accountId;
}

function setModeUI() {
  const isCustom = modeEl.value === 'custom';
  sinceEl.disabled = !isCustom;
  untilEl.disabled = !isCustom;
  datePresetEl.disabled = isCustom;
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}

function setIntegratedStatus(message) {
  integratedStatusEl.textContent = message;
}

function setIntegratedSheetStatus(message) {
  integratedSheetStatusEl.textContent = message;
}

function scheduleMetaReauthRedirect(reauthUrl) {
  if (!reauthUrl) {
    return;
  }

  setStatus('Meta Token 已失效，正在開啟授權視窗...', true);

  const width = 560;
  const height = 760;
  const left = Math.max((window.screen.width - width) / 2, 0);
  const top = Math.max((window.screen.height - height) / 2, 0);
  const features = `popup=yes,width=${width},height=${height},left=${left},top=${top}`;

  metaAuthPopup = window.open(reauthUrl, 'meta_oauth_popup', features);
  if (!metaAuthPopup) {
    setStatus('瀏覽器阻擋彈出視窗，改為頁面跳轉授權...', true);
    window.location.href = reauthUrl;
    return;
  }

  metaAuthPopup.focus();
}

async function onMetaAuthMessage(event) {
  if (event.origin !== window.location.origin) {
    return;
  }

  const data = event.data || {};
  if (data.type !== 'meta-token-updated') {
    return;
  }

  if (data.ok) {
    setStatus('Meta 授權成功，ACCESS_TOKEN 已更新，正在驗證新權杖...');
    if (metaAuthPopup && !metaAuthPopup.closed) {
      metaAuthPopup.close();
    }

    if (getCurrentPlatform() === 'meta') {
      await queryData();
    } else {
      setStatus('Meta 權杖已更新完成。');
    }
  } else {
    setStatus(data.message || 'Meta 授權失敗，請重試。', true);
  }
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function getFxRate() {
  const fx = Number(fxRateEl.value);
  return Number.isFinite(fx) && fx > 0 ? fx : 0;
}

function toTwd(valueInUsd) {
  return toNumber(valueInUsd) * getFxRate();
}

function toUsdFromRow(row) {
  if (String(row.currency || '').toUpperCase() === 'TWD') {
    const fx = getFxRate();
    return fx > 0 ? toNumber(row.spend) / fx : 0;
  }
  return toNumber(row.spend);
}

function rowName(row) {
  return row.campaign_name || row.adset_name || row.ad_name || '未命名';
}

function getAfDownload(actions) {
  if (!Array.isArray(actions)) {
    return 0;
  }

  const preferredTypes = [
    'mobile_app_install',
    'app_install',
    'omni_app_install',
    'app_custom_event.fb_mobile_activate_app'
  ];

  for (const type of preferredTypes) {
    const found = actions.find((item) => item.action_type === type);
    if (found) {
      return toNumber(found.value);
    }
  }

  const fuzzy = actions.find((item) => {
    const actionType = String(item.action_type || '').toLowerCase();
    return actionType.includes('install') || actionType.includes('download') || actionType.includes('af');
  });

  return fuzzy ? toNumber(fuzzy.value) : 0;
}

function detectCategory(row, name, source) {
  const normalizedName = String(name || '').toLowerCase().replace(/\s+/g, '');
  if (source === 'meta' && /a[+＋]sc/.test(normalizedName)) {
    return '網頁預約';
  }

  const apiOs = String(row?.os || '').trim().toLowerCase();
  if (apiOs === 'android') {
    return 'Android';
  }
  if (apiOs === 'ios') {
    return 'iOS';
  }

  const text = String(name || '').toLowerCase();
  const hasAndToken = /(^|[^a-z0-9])and([^a-z0-9]|$)/i.test(text);
  if (text.includes('android') || text.includes('aos') || text.includes('and_') || text.includes('安卓') || hasAndToken) {
    return 'Android';
  }
  if (text.includes('ios') || text.includes('iphone') || text.includes('ipad') || text.includes('ios_')) {
    return 'iOS';
  }
  return '未分類';
}

function normalizeRows(rows, source) {
  return rows.map((row, index) => {
    const results = getAfDownload(row.actions);
    const normalizedSpendUsd = toUsdFromRow(row);
    const impressions = toNumber(row.impressions);
    const clicks = toNumber(row.clicks);
    const ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
    const cpiUsd = results > 0 ? normalizedSpendUsd / results : 0;
    const name = rowName(row);
    const date = row.date_start || row.date_stop || '-';

    return {
      id: `${source}-${name}-${row.date_start || ''}-${row.date_stop || ''}-${index}`,
      source,
      name,
      date,
      os: detectCategory(row, name, source),
      spendUsd: normalizedSpendUsd,
      impressions,
      clicks,
      ctr,
      results,
      cpiUsd
    };
  });
}

function syncAccountForm() {
  const selected = settings.accounts.find((item) => {
    return (item.platform || 'meta') === getCurrentPlatform() && item.accountId === accountSelectEl.value;
  });

  if (!selected) {
    pageNameEl.value = '';
    pageAccountIdEl.value = '';
    return;
  }

  pageNameEl.value = selected.name;
  pageAccountIdEl.value = selected.accountId;
}

function renderAccountOptions() {
  const platform = getCurrentPlatform();
  const accountList = settings.accounts.filter((item) => (item.platform || 'meta') === platform);
  const selectedAccountId = settings.selectedAccountIdByPlatform[platform] || '';
  const optionsHtml = accountList.map((item) => {
    const selected = item.accountId === selectedAccountId ? 'selected' : '';
    return `<option value="${item.accountId}" ${selected}>${item.name} (${item.accountId})</option>`;
  }).join('');

  accountSelectEl.innerHTML = optionsHtml || '<option value="">未設定帳號</option>';

  if (!selectedAccountId && accountList.length) {
    settings.selectedAccountIdByPlatform[platform] = accountList[0].accountId;
    saveSettings();
  }

  if (settings.selectedAccountIdByPlatform[platform]) {
    accountSelectEl.value = settings.selectedAccountIdByPlatform[platform];
  }

  syncAccountForm();
}

function setPlatformUI() {
  const isGoogle = getCurrentPlatform() === 'google';
  accountSelectLabelEl.textContent = isGoogle ? 'Google 帳號設定' : 'Meta 帳號設定';
  pageNameLabelEl.textContent = isGoogle ? 'Google 帳號名稱' : 'Meta 帳號名稱';
  pageAccountIdLabelEl.textContent = isGoogle ? 'Google Customer ID' : 'Meta 廣告帳號 ID';
  pageAccountIdEl.placeholder = isGoogle ? '123-456-7890 或 1234567890' : 'act_1234567890';
  levelEl.disabled = isGoogle;
  renderAccountOptions();
}

function upsertAccount() {
  const name = pageNameEl.value.trim();
  const accountId = pageAccountIdEl.value.trim();

  if (!name) {
    setStatus('請輸入帳號名稱。', true);
    return;
  }

  if (!isValidAccountId(accountId)) {
    const example = getCurrentPlatform() === 'google' ? '123-456-7890' : 'act_1234567890';
    setStatus(`帳號 ID 格式錯誤，範例：${example}`, true);
    return;
  }

  const platform = getCurrentPlatform();
  const idx = settings.accounts.findIndex((item) => {
    return (item.platform || 'meta') === platform && item.accountId === accountId;
  });

  if (idx >= 0) {
    settings.accounts[idx] = { name, accountId, platform };
  } else {
    settings.accounts.push({ name, accountId, platform });
  }

  settings.selectedAccountIdByPlatform[platform] = accountId;
  saveSettings();
  renderAccountOptions();
  setStatus(`${getPlatformText(platform)} 帳號設定已儲存。`);
}

function removeCurrentAccount() {
  const accountId = accountSelectEl.value;
  if (!accountId) {
    setStatus('目前沒有可刪除的帳號設定。', true);
    return;
  }

  const platform = getCurrentPlatform();
  settings.accounts = settings.accounts.filter((item) => {
    return !((item.platform || 'meta') === platform && item.accountId === accountId);
  });

  const remainList = settings.accounts.filter((item) => (item.platform || 'meta') === platform);
  settings.selectedAccountIdByPlatform[platform] = remainList[0]?.accountId || '';
  saveSettings();
  renderAccountOptions();
  setStatus(`${getPlatformText(platform)} 帳號設定已刪除。`);
}

function getSelectedRows() {
  return currentRows.filter((row) => selectedRowIds.has(row.id));
}

function updateSummaryFromSelection() {
  const selectedRows = getSelectedRows();
  const totalSpendUsd = selectedRows.reduce((sum, row) => sum + row.spendUsd, 0);
  const totalImpressions = selectedRows.reduce((sum, row) => sum + row.impressions, 0);
  const totalClicks = selectedRows.reduce((sum, row) => sum + row.clicks, 0);
  const totalResults = selectedRows.reduce((sum, row) => sum + row.results, 0);
  const totalCtr = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
  const avgCpi = totalResults > 0 ? toTwd(totalSpendUsd) / totalResults : 0;

  sumImpressionsEl.textContent = nf.format(totalImpressions);
  sumClicksEl.textContent = nf.format(totalClicks);
  sumCtrEl.textContent = `${totalCtr.toFixed(2)}%`;
  sumResultsEl.textContent = nf.format(totalResults);
  avgCpiEl.textContent = totalResults > 0 ? twdFormatter.format(avgCpi) : '-';
  sumSpendEl.textContent = twdFormatter.format(toTwd(totalSpendUsd));
  summaryEl.hidden = false;
  checkAllEl.checked = currentRows.length > 0 && selectedRows.length === currentRows.length;
  checkAllEl.indeterminate = selectedRows.length > 0 && selectedRows.length < currentRows.length;
}

function bindRowCheckboxEvents() {
  const checkboxes = document.querySelectorAll('.row-check');
  checkboxes.forEach((checkbox) => {
    checkbox.addEventListener('change', (event) => {
      const rowId = event.target.dataset.rowId;
      if (event.target.checked) {
        selectedRowIds.add(rowId);
      } else {
        selectedRowIds.delete(rowId);
      }

      updateSummaryFromSelection();
    });
  });
}

function renderTable() {
  if (!currentRows.length) {
    tableWrapEl.hidden = true;
    summaryEl.hidden = true;
    return;
  }

  tbodyEl.innerHTML = currentRows.map((row) => {
    const checked = selectedRowIds.has(row.id) ? 'checked' : '';
    return `
      <tr>
        <td class="checkbox-cell"><input class="row-check" type="checkbox" data-row-id="${row.id}" ${checked} /></td>
        <td>${row.name}</td>
        <td>${row.date}</td>
        <td>${row.os}</td>
        <td>${nf.format(row.impressions)}</td>
        <td>${nf.format(row.clicks)}</td>
        <td>${row.ctr.toFixed(2)}%</td>
        <td>${nf.format(row.results)}</td>
        <td>${row.results > 0 ? twdFormatter.format(toTwd(row.cpiUsd)) : '-'}</td>
        <td>${twdFormatter.format(toTwd(row.spendUsd))}</td>
      </tr>
    `;
  }).join('');

  tableWrapEl.hidden = false;
  bindRowCheckboxEvents();
  updateSummaryFromSelection();
}

function selectAllRows(checked) {
  selectedRowIds = checked ? new Set(currentRows.map((row) => row.id)) : new Set();
  renderTable();
}

function renderCurrentPlatformState() {
  currentRows = [...platformRows[getCurrentPlatform()]];
  selectedRowIds = new Set(currentRows.map((row) => row.id));
  renderTable();

  if (!currentRows.length) {
    setStatus(`${getPlatformText(getCurrentPlatform())} 尚未查詢資料。`);
  }
}

function renderIntegratedView() {
  const combinedRows = [...platformRows.meta, ...platformRows.google];
  if (!combinedRows.length) {
    integratedMetaSpendEl.textContent = '-';
    integratedGoogleSpendEl.textContent = '-';
    integratedAndroidSpendEl.textContent = '-';
    integratedIosSpendEl.textContent = '-';
    integratedWebBookingSpendEl.textContent = '-';
    integratedUnknownSpendEl.textContent = '-';
    integratedTotalSpendEl.textContent = '-';
    integratedMatrixWrapEl.hidden = true;
    integratedTableWrapEl.hidden = true;
    setIntegratedStatus('請先到 Meta 與 Google 分頁各自查詢資料。');
    return;
  }

  const sumSpendBy = (predicate) => {
    return combinedRows.filter(predicate).reduce((sum, row) => sum + row.spendUsd, 0);
  };

  integratedMetaSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.source === 'meta')));
  integratedGoogleSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.source === 'google')));
  integratedAndroidSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.os === 'Android')));
  integratedIosSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.os === 'iOS')));
  integratedWebBookingSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.os === '網頁預約')));
  integratedUnknownSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy((row) => row.os === '未分類')));
  integratedTotalSpendEl.textContent = twdFormatter.format(toTwd(sumSpendBy(() => true)));

  const categoryOrder = ['Android', 'iOS', '網頁預約', '未分類'];
  const categories = Array.from(new Set([
    ...categoryOrder,
    ...combinedRows.map((row) => row.os || '未分類')
  ])).filter(Boolean);

  const compareByCategory = new Map();
  combinedRows.forEach((row) => {
    const category = row.os || '未分類';
    const current = compareByCategory.get(category) || {
      category,
      metaSpendUsd: 0,
      googleSpendUsd: 0,
      metaResults: 0,
      googleResults: 0
    };

    if (row.source === 'meta') {
      current.metaSpendUsd += row.spendUsd;
      current.metaResults += row.results;
    } else {
      current.googleSpendUsd += row.spendUsd;
      current.googleResults += row.results;
    }

    compareByCategory.set(category, current);
  });

  integratedMatrixTbodyEl.innerHTML = categories
    .filter((category) => compareByCategory.has(category))
    .map((category) => compareByCategory.get(category))
    .sort((left, right) => {
      return categoryOrder.indexOf(left.category) - categoryOrder.indexOf(right.category);
    })
    .map((row) => {
      const totalSpendUsd = row.metaSpendUsd + row.googleSpendUsd;
      const totalResults = row.metaResults + row.googleResults;
      return `
        <tr>
          <td>${row.category}</td>
          <td>${twdFormatter.format(toTwd(row.metaSpendUsd))}</td>
          <td>${twdFormatter.format(toTwd(row.googleSpendUsd))}</td>
          <td>${twdFormatter.format(toTwd(totalSpendUsd))}</td>
          <td>${nf.format(row.metaResults)}</td>
          <td>${nf.format(row.googleResults)}</td>
          <td>${nf.format(totalResults)}</td>
        </tr>
      `;
    }).join('');

  const grouped = new Map();
  combinedRows.forEach((row) => {
    const key = `${row.source}|${row.os}`;
    const current = grouped.get(key) || {
      source: row.source,
      os: row.os,
      count: 0,
      impressions: 0,
      clicks: 0,
      results: 0,
      spendUsd: 0
    };

    current.count += 1;
    current.impressions += row.impressions;
    current.clicks += row.clicks;
    current.results += row.results;
    current.spendUsd += row.spendUsd;
    grouped.set(key, current);
  });

  integratedTbodyEl.innerHTML = Array.from(grouped.values()).map((row) => {
    const cpiTwd = row.results > 0 ? toTwd(row.spendUsd) / row.results : 0;
    return `
      <tr>
        <td>${getPlatformText(row.source)}</td>
        <td>${row.os}</td>
        <td>${nf.format(row.count)}</td>
        <td>${nf.format(row.impressions)}</td>
        <td>${nf.format(row.clicks)}</td>
        <td>${nf.format(row.results)}</td>
        <td>${twdFormatter.format(toTwd(row.spendUsd))}</td>
        <td>${row.results > 0 ? twdFormatter.format(cpiTwd) : '-'}</td>
      </tr>
    `;
  }).join('');

  integratedMatrixWrapEl.hidden = false;
  integratedTableWrapEl.hidden = false;
  setIntegratedStatus('整合頁已依分類先完成 Meta / Google 對照，再提供來源明細。');
}

function buildIntegratedSheetMetrics(combinedRows) {
  if (!combinedRows.length) {
    return null;
  }

  const dateValues = combinedRows
    .map((row) => row.date)
    .filter((text) => /^\d{4}-\d{2}-\d{2}$/.test(String(text || '')))
    .sort();
  const firstDate = dateValues[0] || '-';
  const lastDate = dateValues[dateValues.length - 1] || '-';
  const rangeText = firstDate === '-' ? '-' : `${firstDate} ~ ${lastDate}`;

  const summarize = (source, category) => {
    const matched = combinedRows.filter((row) => row.source === source && row.os === category);
    const spendUsd = matched.reduce((sum, row) => sum + row.spendUsd, 0);
    const results = matched.reduce((sum, row) => sum + row.results, 0);
    return {
      spendTwd: toTwd(spendUsd),
      results,
      cpa: results > 0 ? toTwd(spendUsd) / results : null
    };
  };

  const fbAnd = summarize('meta', 'Android');
  const fbIos = summarize('meta', 'iOS');
  const fbWeb = summarize('meta', '網頁預約');
  const gaAnd = summarize('google', 'Android');
  const gaIos = summarize('google', 'iOS');
  const gaWeb = summarize('google', '網頁預約');

  const andSpend = fbAnd.spendTwd + gaAnd.spendTwd;
  const iosSpend = fbIos.spendTwd + gaIos.spendTwd;
  const webSpend = fbWeb.spendTwd + gaWeb.spendTwd;
  const andResults = fbAnd.results + gaAnd.results;
  const iosResults = fbIos.results + gaIos.results;
  const webResults = fbWeb.results + gaWeb.results;
  const excludedRows = combinedRows.filter((row) => row.os !== 'Android' && row.os !== 'iOS' && row.os !== '網頁預約');
  const excludedSpend = excludedRows.reduce((sum, row) => sum + toTwd(row.spendUsd), 0);
  const excludedResults = excludedRows.reduce((sum, row) => sum + row.results, 0);
  const totalSpend = andSpend + iosSpend + webSpend + excludedSpend;
  const totalResults = andResults + iosResults + webResults + excludedResults;
  const andCpa = andResults > 0 ? andSpend / andResults : null;
  const iosCpa = iosResults > 0 ? iosSpend / iosResults : null;
  const webCpa = webResults > 0 ? webSpend / webResults : null;
  const totalCpa = totalResults > 0 ? totalSpend / totalResults : null;

  return {
    rangeText,
    fbAnd,
    fbIos,
    fbWeb,
    gaAnd,
    gaIos,
    gaWeb,
    andSpend,
    iosSpend,
    webSpend,
    totalSpend,
    andResults,
    iosResults,
    webResults,
    totalResults,
    andCpa,
    iosCpa,
    webCpa,
    totalCpa,
    sourceRowCount: combinedRows.length,
    includedRowCount: combinedRows.length - excludedRows.length,
    excludedRowCount: excludedRows.length,
    excludedSpend,
    excludedResults
  };
}

function renderIntegratedSheetTables(metrics) {
  const cpaText = (value) => value === null ? '-' : nf.format(Math.round(value));
  const amountText = (value) => nf.format(Math.round(value));

  integratedSheetSourceTbodyEl.innerHTML = `
    <tr>
      <td>花費</td>
      <td>${amountText(metrics.fbAnd.spendTwd)}</td>
      <td>${amountText(metrics.fbIos.spendTwd)}</td>
      <td>${amountText(metrics.fbWeb.spendTwd)}</td>
      <td>${amountText(metrics.gaAnd.spendTwd)}</td>
      <td>${amountText(metrics.gaIos.spendTwd)}</td>
    </tr>
    <tr>
      <td>人數(AF)</td>
      <td>${nf.format(metrics.fbAnd.results)}</td>
      <td>${nf.format(metrics.fbIos.results)}</td>
      <td>${nf.format(metrics.fbWeb.results)}</td>
      <td>${nf.format(metrics.gaAnd.results)}</td>
      <td>${nf.format(metrics.gaIos.results)}</td>
    </tr>
    <tr>
      <td>成本</td>
      <td>${cpaText(metrics.fbAnd.cpa)}</td>
      <td>${cpaText(metrics.fbIos.cpa)}</td>
      <td>${cpaText(metrics.fbWeb.cpa)}</td>
      <td>${cpaText(metrics.gaAnd.cpa)}</td>
      <td>${cpaText(metrics.gaIos.cpa)}</td>
    </tr>
  `;

  integratedSheetSummaryTbodyEl.innerHTML = `
    <tr>
      <td>花費</td>
      <td>${amountText(metrics.andSpend)}</td>
      <td>${amountText(metrics.iosSpend)}</td>
      <td>${amountText(metrics.webSpend)}</td>
      <td>${amountText(metrics.totalSpend)}</td>
    </tr>
    <tr>
      <td>總人數</td>
      <td>${nf.format(metrics.andResults)}</td>
      <td>${nf.format(metrics.iosResults)}</td>
      <td>${nf.format(metrics.webResults)}</td>
      <td>${nf.format(metrics.totalResults)}</td>
    </tr>
    <tr>
      <td>成本</td>
      <td>${cpaText(metrics.andCpa)}</td>
      <td>${cpaText(metrics.iosCpa)}</td>
      <td>${cpaText(metrics.webCpa)}</td>
      <td>${cpaText(metrics.totalCpa)}</td>
    </tr>
  `;
}

function renderIntegratedSheetSnapshots() {
  if (!integratedSheetNotesEl) {
    return;
  }

  if (!integratedSheetSnapshots.length) {
    integratedSheetNotesEl.innerHTML = '<div class="sticky-empty">尚未建立便利貼，查詢後可把這次投放簡表存成一張卡片。</div>';
    return;
  }

  const cpaText = (value) => value === null ? '-' : nf.format(Math.round(value));
  const amountText = (value) => nf.format(Math.round(value));

  integratedSheetNotesEl.innerHTML = integratedSheetSnapshots.map((snapshot) => {
    const metrics = snapshot.metrics || {};
    const fbAnd = metrics.fbAnd || { spendTwd: 0, results: 0, cpa: null };
    const fbIos = metrics.fbIos || { spendTwd: 0, results: 0, cpa: null };
    const fbWeb = metrics.fbWeb || { spendTwd: 0, results: 0, cpa: null };
    const gaAnd = metrics.gaAnd || { spendTwd: 0, results: 0, cpa: null };
    const gaIos = metrics.gaIos || { spendTwd: 0, results: 0, cpa: null };
    const title = escapeHtml(snapshot.title || metrics.rangeText || '-');
    const createdAt = escapeHtml(snapshot.createdAt || '-');

    return `
      <article class="sticky-note">
        <div class="sticky-note-head">
          <div>
            <p class="sticky-note-title">${title}</p>
            <div class="sticky-note-time">建立時間：${createdAt}</div>
          </div>
          <button type="button" class="sticky-note-delete" data-snapshot-id="${snapshot.id}">刪除</button>
        </div>

        <div class="sticky-mini">
          <table>
            <thead>
              <tr>
                <th></th>
                <th>FB AND</th>
                <th>FB iOS</th>
                <th>FB WEB</th>
                <th>GA AND</th>
                <th>GA iOS</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>花費</td>
                <td>${amountText(fbAnd.spendTwd)}</td>
                <td>${amountText(fbIos.spendTwd)}</td>
                <td>${amountText(fbWeb.spendTwd)}</td>
                <td>${amountText(gaAnd.spendTwd)}</td>
                <td>${amountText(gaIos.spendTwd)}</td>
              </tr>
              <tr>
                <td>人數</td>
                <td>${nf.format(fbAnd.results)}</td>
                <td>${nf.format(fbIos.results)}</td>
                <td>${nf.format(fbWeb.results)}</td>
                <td>${nf.format(gaAnd.results)}</td>
                <td>${nf.format(gaIos.results)}</td>
              </tr>
              <tr>
                <td>成本</td>
                <td>${cpaText(fbAnd.cpa)}</td>
                <td>${cpaText(fbIos.cpa)}</td>
                <td>${cpaText(fbWeb.cpa)}</td>
                <td>${cpaText(gaAnd.cpa)}</td>
                <td>${cpaText(gaIos.cpa)}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="sticky-total">
          <div class="sticky-total-item">
            <span>總花費</span>
            <strong>${amountText(metrics.totalSpend || 0)}</strong>
          </div>
          <div class="sticky-total-item">
            <span>總人數</span>
            <strong>${nf.format(metrics.totalResults || 0)}</strong>
          </div>
          <div class="sticky-total-item">
            <span>總成本</span>
            <strong>${cpaText(metrics.totalCpa ?? null)}</strong>
          </div>
        </div>
      </article>
    `;
  }).join('');
}

function saveIntegratedSheetSnapshot() {
  const combinedRows = [...platformRows.meta, ...platformRows.google];
  const metrics = buildIntegratedSheetMetrics(combinedRows);
  if (!metrics) {
    setIntegratedSheetStatus('目前沒有資料可存成便利貼，請先查詢 Meta / Google。');
    return;
  }

  const customTitle = integratedSheetTitleEl.value.trim();
  const snapshot = {
    id: `snap_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    title: customTitle || metrics.rangeText,
    createdAt: new Date().toLocaleString('zh-TW', { hour12: false }),
    metrics
  };

  integratedSheetSnapshots.unshift(snapshot);
  integratedSheetSnapshots = integratedSheetSnapshots.slice(0, 50);
  saveIntegratedSheetSnapshots();
  renderIntegratedSheetSnapshots();
  integratedSheetTitleEl.value = '';
  setIntegratedSheetStatus('已新增一張投放簡表便利貼。');
}

function deleteIntegratedSheetSnapshot(snapshotId) {
  integratedSheetSnapshots = integratedSheetSnapshots.filter((item) => item.id !== snapshotId);
  saveIntegratedSheetSnapshots();
  renderIntegratedSheetSnapshots();
  setIntegratedSheetStatus('已刪除便利貼。');
}

function renderIntegratedSheetView() {
  const combinedRows = [...platformRows.meta, ...platformRows.google];
  const metrics = buildIntegratedSheetMetrics(combinedRows);
  if (!metrics) {
    integratedSheetDateEl.textContent = '報表區間：-';
    integratedSheetSourceWrapEl.hidden = true;
    integratedSheetSummaryWrapEl.hidden = true;
    setIntegratedSheetStatus('請先到 Meta 與 Google 分頁各自查詢資料。');
    renderIntegratedSheetSnapshots();
    return;
  }

  integratedSheetDateEl.textContent = `報表區間：${metrics.rangeText}`;
  renderIntegratedSheetTables(metrics);

  integratedSheetSourceWrapEl.hidden = false;
  integratedSheetSummaryWrapEl.hidden = false;
  if (metrics.includedRowCount === 0 && metrics.sourceRowCount > 0) {
    setIntegratedSheetStatus('已有查詢資料；目前都屬於未分類，已先納入「總計」，AND/iOS/網頁預約 欄位會是 0。');
  } else if (metrics.excludedRowCount > 0) {
    setIntegratedSheetStatus(`整合頁2已建立 FB/GA 與 AND/iOS/網頁預約 成本對照（另有 ${nf.format(metrics.excludedRowCount)} 筆未納入；花費 ${nf.format(Math.round(metrics.excludedSpend))}，成果 ${nf.format(metrics.excludedResults)}）。`);
  } else {
    setIntegratedSheetStatus('整合頁2已建立 FB/GA 與 AND/iOS/網頁預約 成本對照。');
  }
  renderIntegratedSheetSnapshots();
}

function setActiveTab(tab) {
  activeTab = tab;

  tabMetaEl.classList.toggle('active', tab === 'meta');
  tabGoogleEl.classList.toggle('active', tab === 'google');
  tabIntegratedEl.classList.toggle('active', tab === 'integrated');
  tabIntegrated2El.classList.toggle('active', tab === 'integrated2');

  queryPanelEl.hidden = tab === 'integrated' || tab === 'integrated2';
  integratedPanelEl.hidden = tab !== 'integrated';
  integratedSheetPanelEl.hidden = tab !== 'integrated2';

  if (tab === 'integrated') {
    renderIntegratedView();
    return;
  }

  if (tab === 'integrated2') {
    renderIntegratedSheetView();
    return;
  }

  settings.selectedPlatform = tab;
  saveSettings();
  setPlatformUI();
  renderCurrentPlatformState();
}

async function queryData() {
  if (activeTab === 'integrated') {
    renderIntegratedView();
    return;
  }

  queryBtn.disabled = true;
  tableWrapEl.hidden = true;
  summaryEl.hidden = true;
  setStatus('查詢中，請稍候...');

  try {
    const fx = getFxRate();
    if (!fx) {
      throw new Error('匯率需大於 0。');
    }

    settings.fxRate = fx;
    saveSettings();

    const platform = getCurrentPlatform();
    const endpoint = platform === 'google' ? '/api/google/insights' : '/api/insights';
    const params = new URLSearchParams();

    if (platform === 'meta') {
      params.set('level', levelEl.value);
    }

    const selectedAccountId = getCurrentSelectedAccountId();
    if (selectedAccountId) {
      params.set('accountId', selectedAccountId);
    }

    if (modeEl.value === 'custom') {
      if (!sinceEl.value || !untilEl.value) {
        throw new Error('自訂日期需要同時選擇開始與結束日期。');
      }
      params.set('since', sinceEl.value);
      params.set('until', untilEl.value);
    } else {
      params.set('datePreset', datePresetEl.value);
    }

    const response = await fetch(`${endpoint}?${params.toString()}`);
    const payload = await response.json();

    if (!response.ok) {
      if (platform === 'meta' && response.status === 401 && payload?.details?.code === 190) {
        scheduleMetaReauthRedirect(payload.reauthUrl);
      }

      const details = payload.details ? JSON.stringify(payload.details) : '';
      throw new Error(payload.error ? `${payload.error} ${details}` : '查詢失敗');
    }

    const rows = payload.data || [];
    if (!rows.length) {
      platformRows[platform] = [];
      currentRows = [];
      selectedRowIds = new Set();
      renderTable();
      renderIntegratedView();
      setStatus('此區間無資料。');
      return;
    }

    const normalized = normalizeRows(rows, platform);
    platformRows[platform] = normalized;
    currentRows = [...normalized];
    selectedRowIds = new Set(currentRows.map((row) => row.id));
    renderTable();
    renderIntegratedView();
    setStatus(`${getPlatformText(platform)} 查詢完成，共 ${rows.length} 筆。可用勾選框加總上方數據。`);
  } catch (error) {
    setStatus(error.message || '查詢發生錯誤。', true);
  } finally {
    queryBtn.disabled = false;
  }
}

modeEl.addEventListener('change', setModeUI);
window.addEventListener('message', onMetaAuthMessage);
tabMetaEl.addEventListener('click', () => setActiveTab('meta'));
tabGoogleEl.addEventListener('click', () => setActiveTab('google'));
tabIntegratedEl.addEventListener('click', () => setActiveTab('integrated'));
tabIntegrated2El.addEventListener('click', () => setActiveTab('integrated2'));
integratedSheetSaveBtnEl.addEventListener('click', saveIntegratedSheetSnapshot);
integratedSheetNotesEl.addEventListener('click', (event) => {
  const deleteBtn = event.target.closest('.sticky-note-delete');
  if (!deleteBtn) {
    return;
  }

  deleteIntegratedSheetSnapshot(deleteBtn.dataset.snapshotId);
});
accountSelectEl.addEventListener('change', () => {
  setCurrentSelectedAccountId(accountSelectEl.value);
  saveSettings();
  syncAccountForm();
});
addPageBtn.addEventListener('click', upsertAccount);
removePageBtn.addEventListener('click', removeCurrentAccount);
fxRateEl.addEventListener('change', () => {
  const fx = getFxRate();
  if (!fx) {
    setStatus('匯率需大於 0。', true);
    return;
  }

  settings.fxRate = fx;
  saveSettings();
  renderTable();
  renderIntegratedView();
  setStatus('匯率已更新。');
});
checkAllEl.addEventListener('change', (event) => {
  selectAllRows(event.target.checked);
});
queryBtn.addEventListener('click', queryData);

fxRateEl.value = settings.fxRate;
setModeUI();
renderIntegratedSheetSnapshots();
setActiveTab(settings.selectedPlatform === 'google' ? 'google' : 'meta');
