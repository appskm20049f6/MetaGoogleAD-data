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
const savePresetBtn = document.getElementById('savePresetBtn');
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
const afDbGameSelectEl = document.getElementById('afDbGameSelect');
const afDbReloadBtnEl = document.getElementById('afDbReloadBtn');
const afDbListStatusEl = document.getElementById('afDbListStatus');
const afDbSinceEl = document.getElementById('afDbSince');
const afDbUntilEl = document.getElementById('afDbUntil');
const afDbLabelEl = document.getElementById('afDbLabel');
const afDbFetchBtnEl = document.getElementById('afDbFetchBtn');
const afDbStatusEl = document.getElementById('afDbStatus');
const afDbWrapEl = document.getElementById('afDbWrap');
const afDbTbodyEl = document.getElementById('afDbTbody');
const afDbSourceSectionEl = document.getElementById('afDbSourceSection');
const afDbSourceTbodyEl = document.getElementById('afDbSourceTbody');
const afDbDetailSectionEl = document.getElementById('afDbDetailSection');
const afDbDetailTbodyEl = document.getElementById('afDbDetailTbody');
const afDbMappedSectionEl = document.getElementById('afDbMappedSection');
const afDbMappedTbodyEl = document.getElementById('afDbMappedTbody');
const tabCopySheetEl = document.getElementById('tabCopySheet');
const copySheetPanelEl = document.getElementById('copySheetPanel');
const copySheetDateLabelEl = document.getElementById('copySheetDateLabel');
const copySheetOfflineAndEl = document.getElementById('copySheetOfflineAnd');
const copySheetOfflineIosEl = document.getElementById('copySheetOfflineIos');
const copySheetStatusEl = document.getElementById('copySheetStatus');
const copySheetTableWrapEl = document.getElementById('copySheetTableWrap');
const copySheetCopyBtnEl = document.getElementById('copySheetCopyBtn');

const sumImpressionsEl = document.getElementById('sumImpressions');
const sumClicksEl = document.getElementById('sumClicks');
const sumCtrEl = document.getElementById('sumCtr');
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
let afMappedMetrics = null;
let afTotalCounts = null;
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

const {
  resolveCompanionSheets,
  normalizeSheetPayload,
  mergeSheetPayloads,
  mapAfSourcesToPlatforms
} = window.AppAf || {};

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

  // Keep a conservative fallback to avoid counting non-install events as AF downloads.
  const fallbackInstallTypes = [
    'app_custom_event.install',
    'app_custom_event.app_install',
    'app_custom_event.mobile_app_install'
  ];

  const fuzzy = actions.find((item) => {
    const actionType = String(item.action_type || '').toLowerCase();
    if (!actionType) {
      return false;
    }

    if (fallbackInstallTypes.includes(actionType)) {
      return true;
    }

    return /(^|[._])app_install($|[._])/.test(actionType)
      || /(^|[._])mobile_app_install($|[._])/.test(actionType)
      || /(^|[._])omni_app_install($|[._])/.test(actionType)
      || actionType === 'app_custom_event.fb_mobile_activate_app';
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

async function saveAccountPresetsToServer() {
  try {
    const payload = {
      selectedPlatform: settings.selectedPlatform,
      selectedAccountIdByPlatform: settings.selectedAccountIdByPlatform,
      accounts: settings.accounts
    };

    const response = await fetch('/api/account-presets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || '儲存程式預設失敗');
    }

    setStatus(`已儲存程式預設（共 ${result.saved?.accounts?.length || 0} 組帳號）。`);
  } catch (error) {
    setStatus(error.message || '儲存程式預設失敗。', true);
  }
}

async function hydrateSettingsFromServerPresets() {
  try {
    const response = await fetch('/api/account-presets');
    if (!response.ok) {
      return;
    }

    const data = await response.json();
    const serverAccounts = Array.isArray(data.accounts) ? data.accounts : [];
    if (!serverAccounts.length) {
      return;
    }

    const hasLocalAccounts = Array.isArray(settings.accounts) && settings.accounts.length > 0;
    if (hasLocalAccounts) {
      return;
    }

    settings.accounts = serverAccounts;
    settings.selectedAccountIdByPlatform = {
      meta: data.selectedAccountIdByPlatform?.meta || '',
      google: data.selectedAccountIdByPlatform?.google || ''
    };

    if (data.selectedPlatform === 'meta' || data.selectedPlatform === 'google') {
      settings.selectedPlatform = data.selectedPlatform;
    }

    saveSettings();
    renderAccountOptions();
    setStatus('已套用程式預設帳號設定。');
  } catch (_) {
    // Keep local settings when preset loading fails.
  }
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

  const hasAfCounts = Boolean(afMappedMetrics && afTotalCounts);
  const gaAndBase = hasAfCounts ? Number(afMappedMetrics.gaAnd || 0) : 0;
  const gaIosBase = hasAfCounts ? Number(afMappedMetrics.gaIos || 0) : 0;
  const gaAndResults = hasAfCounts ? Math.max(0, Math.min(Number(afTotalCounts.andTotal || 0), gaAndBase)) : 0;
  const gaIosResults = hasAfCounts ? Math.max(0, Math.min(Number(afTotalCounts.iosTotal || 0), gaIosBase)) : 0;
  const fbAndResults = hasAfCounts ? Math.max(0, Number(afTotalCounts.andTotal || 0) - gaAndResults) : 0;
  const fbIosResults = hasAfCounts ? Math.max(0, Number(afTotalCounts.iosTotal || 0) - gaIosResults) : 0;

  if (hasAfCounts) {
    const androidRow = compareByCategory.get('Android') || {
      category: 'Android',
      metaSpendUsd: 0,
      googleSpendUsd: 0,
      metaResults: 0,
      googleResults: 0
    };
    const iosRow = compareByCategory.get('iOS') || {
      category: 'iOS',
      metaSpendUsd: 0,
      googleSpendUsd: 0,
      metaResults: 0,
      googleResults: 0
    };

    androidRow.metaResults = fbAndResults;
    androidRow.googleResults = gaAndResults;
    iosRow.metaResults = fbIosResults;
    iosRow.googleResults = gaIosResults;

    compareByCategory.set('Android', androidRow);
    compareByCategory.set('iOS', iosRow);

    ['網頁預約', '未分類'].forEach((category) => {
      if (!compareByCategory.has(category)) {
        return;
      }
      const row = compareByCategory.get(category);
      row.metaResults = 0;
      row.googleResults = 0;
      compareByCategory.set(category, row);
    });
  }

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

  const getDetailAfResults = (source, os, fallbackResults) => {
    if (!hasAfCounts) {
      return fallbackResults;
    }

    if (os === 'Android') {
      if (source === 'meta') return fbAndResults;
      if (source === 'google') return gaAndResults;
    }

    if (os === 'iOS') {
      if (source === 'meta') return fbIosResults;
      if (source === 'google') return gaIosResults;
    }

    return 0;
  };

  integratedTbodyEl.innerHTML = Array.from(grouped.values()).map((row) => {
    const detailResults = getDetailAfResults(row.source, row.os, row.results);
    const cpiTwd = detailResults > 0 ? toTwd(row.spendUsd) / detailResults : 0;
    return `
      <tr>
        <td>${getPlatformText(row.source)}</td>
        <td>${row.os}</td>
        <td>${nf.format(row.count)}</td>
        <td>${nf.format(row.impressions)}</td>
        <td>${nf.format(row.clicks)}</td>
        <td>${nf.format(detailResults)}</td>
        <td>${twdFormatter.format(toTwd(row.spendUsd))}</td>
        <td>${detailResults > 0 ? twdFormatter.format(cpiTwd) : '-'}</td>
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
    const adResults = matched.reduce((sum, row) => sum + row.results, 0);
    return {
      spendTwd: toTwd(spendUsd),
      adResults
    };
  };

  const fbAnd = summarize('meta', 'Android');
  const fbIos = summarize('meta', 'iOS');
  const fbWeb = summarize('meta', '網頁預約');
  const gaAnd = summarize('google', 'Android');
  const gaIos = summarize('google', 'iOS');
  const gaWeb = summarize('google', '網頁預約');

  const hasAfMapped = Boolean(afMappedMetrics && afTotalCounts);
  const afAndTotal = hasAfMapped ? Math.max(0, Number(afTotalCounts.andTotal || 0)) : 0;
  const afIosTotal = hasAfMapped ? Math.max(0, Number(afTotalCounts.iosTotal || 0)) : 0;
  const gaAndBase = hasAfMapped ? Math.max(0, Number(afMappedMetrics.gaAnd || 0)) : 0;
  const gaIosBase = hasAfMapped ? Math.max(0, Number(afMappedMetrics.gaIos || 0)) : 0;
  const gaAndResults = hasAfMapped ? Math.min(afAndTotal, gaAndBase) : gaAnd.adResults;
  const gaIosResults = hasAfMapped ? Math.min(afIosTotal, gaIosBase) : gaIos.adResults;
  const fbAndResults = hasAfMapped ? Math.max(0, afAndTotal - gaAndResults) : fbAnd.adResults;
  const fbIosResults = hasAfMapped ? Math.max(0, afIosTotal - gaIosResults) : fbIos.adResults;
  const fbWebResults = fbWeb.adResults;
  const gaWebResults = gaWeb.adResults;

  fbAnd.results = fbAndResults;
  fbIos.results = fbIosResults;
  gaAnd.results = gaAndResults;
  gaIos.results = gaIosResults;
  fbWeb.results = fbWebResults;
  gaWeb.results = gaWebResults;

  fbAnd.cpa = fbAnd.results > 0 ? fbAnd.spendTwd / fbAnd.results : null;
  fbIos.cpa = fbIos.results > 0 ? fbIos.spendTwd / fbIos.results : null;
  fbWeb.cpa = fbWeb.results > 0 ? fbWeb.spendTwd / fbWeb.results : null;
  gaAnd.cpa = gaAnd.results > 0 ? gaAnd.spendTwd / gaAnd.results : null;
  gaIos.cpa = gaIos.results > 0 ? gaIos.spendTwd / gaIos.results : null;
  gaWeb.cpa = gaWeb.results > 0 ? gaWeb.spendTwd / gaWeb.results : null;

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
  if (!integratedSheetSourceTbodyEl || !integratedSheetSummaryTbodyEl) {
    return;
  }

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
    if (integratedSheetSourceWrapEl) integratedSheetSourceWrapEl.hidden = true;
    if (integratedSheetSummaryWrapEl) integratedSheetSummaryWrapEl.hidden = true;
    setIntegratedSheetStatus('請先到 Meta 與 Google 分頁各自查詢資料。');
    renderIntegratedSheetSnapshots();
    return;
  }

  integratedSheetDateEl.textContent = `報表區間：${metrics.rangeText}`;
  renderIntegratedSheetTables(metrics);

  if (integratedSheetSourceWrapEl) integratedSheetSourceWrapEl.hidden = false;
  if (integratedSheetSummaryWrapEl) integratedSheetSummaryWrapEl.hidden = false;
  if (metrics.includedRowCount === 0 && metrics.sourceRowCount > 0) {
    setIntegratedSheetStatus('已有查詢資料；目前都屬於未分類，已先納入「總計」，AND/iOS/網頁預約 欄位會是 0。');
  } else if (afMappedMetrics) {
    setIntegratedSheetStatus('投放簡表已使用 AF 下載表，並依來源(Meta/Google)與 OS(Android/iOS) 分類套用到各區塊。');
  } else if (metrics.excludedRowCount > 0) {
    setIntegratedSheetStatus(`整合頁2已建立 FB/GA 與 AND/iOS/網頁預約 成本對照（另有 ${nf.format(metrics.excludedRowCount)} 筆未納入；花費 ${nf.format(Math.round(metrics.excludedSpend))}，成果 ${nf.format(metrics.excludedResults)}）。`);
  } else {
    setIntegratedSheetStatus('整合頁2已建立 FB/GA 與 AND/iOS/網頁預約 成本對照。');
  }
  renderIntegratedSheetSnapshots();
}

async function loadAfSheetList() {
  afDbListStatusEl.textContent = '載入中…';
  afDbReloadBtnEl.disabled = true;
  try {
    const res = await fetch('/api/sheet/installs');
    const data = await res.json();
    if (!data.availableSheets) {
      afDbListStatusEl.textContent = data.error || '無法取得清單';
      return;
    }
    const current = afDbGameSelectEl.value;
    afDbGameSelectEl.innerHTML = '<option value="">— 請選擇 —</option>' +
      data.availableSheets.map((s) =>
        `<option value="${s}"${s === current ? ' selected' : ''}>${s}</option>`
      ).join('');
    afDbListStatusEl.textContent = `共 ${data.availableSheets.length} 個工作表`;
  } catch (e) {
    afDbListStatusEl.textContent = `失敗：${e.message}`;
  } finally {
    afDbReloadBtnEl.disabled = false;
  }
}


async function fetchAfSheetData() {
  if (!resolveCompanionSheets || !normalizeSheetPayload || !mergeSheetPayloads || !mapAfSourcesToPlatforms) {
    afDbStatusEl.textContent = 'AF helper 載入失敗，請重新整理頁面。';
    return;
  }

  const game = afDbGameSelectEl.value;
  if (!game) {
    afDbStatusEl.textContent = '請先選擇工作表';
    return;
  }

  afDbStatusEl.textContent = '抓取中…';
  afDbFetchBtnEl.disabled = true;

  const startDate = afDbSinceEl.value;
  const endDate = afDbUntilEl.value;
  const label = afDbLabelEl.value.trim() || game;

  try {
    const allSheets = Array.from(afDbGameSelectEl.options).map((opt) => opt.value).filter(Boolean);
    const targetSheets = resolveCompanionSheets(game, allSheets);

    const payloads = await Promise.all(targetSheets.map(async (sheetName) => {
      const params = new URLSearchParams({ game: sheetName });
      if (startDate) params.append('startDate', startDate);
      if (endDate) params.append('endDate', endDate);
      const res = await fetch(`/api/sheet/installs?${params}`);
      const data = await res.json();
      if (data.error) {
        throw new Error(`[${sheetName}] ${data.error}`);
      }
      return normalizeSheetPayload(data, sheetName);
    }));

    const merged = mergeSheetPayloads(payloads);
    afMappedMetrics = mapAfSourcesToPlatforms(merged);
    afTotalCounts = {
      andTotal: merged.android.total,
      iosTotal: merged.ios.total,
      andInstalls: merged.android.installs,
      iosInstalls: merged.ios.installs,
      andReattribution: merged.android.reattribution,
      iosReattribution: merged.ios.reattribution,
      andReengagement: merged.android.reengagement,
      iosReengagement: merged.ios.reengagement
    };

    const rangeLabel = startDate ? `${startDate} ~ ${endDate}` : '全部資料';
    afDbStatusEl.textContent = `${label}｜${rangeLabel}｜合併 ${targetSheets.length} 張表（${targetSheets.join('、')}）`;
    renderAfSheetResult(merged, label, rangeLabel);
    // 逐日明細：若有自訂區間則為每一天呼叫 API 並顯示每日統計（含安裝/再歸因/再互動）
    try {
      await fetchAndRenderAfDaily(game, allSheets, startDate, endDate, label);
    } catch (errDaily) {
      console.warn('fetch daily failed', errDaily);
    }
    renderIntegratedSheetView();
    renderCopySheetView();
  } catch (e) {
    afMappedMetrics = null;
      afTotalCounts = null;
    afDbStatusEl.textContent = `請求失敗：${e.message}`;
  } finally {
    afDbFetchBtnEl.disabled = false;
  }
}

function formatYmd(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

async function fetchAndRenderAfDaily(game, allSheets, startDate, endDate, label) {
  if (!game) return;
  // if no startDate provided, do nothing
  if (!startDate) return;
  const s = new Date(startDate + 'T00:00:00');
  const e = new Date(endDate + 'T00:00:00');
  if (isNaN(s) || isNaN(e) || s > e) return;

  const days = [];
  for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    days.push(new Date(d));
  }

  const dailyResults = [];
  for (const day of days) {
    const dayStr = formatYmd(day);
    try {
      const payloads = await Promise.all(resolveCompanionSheets(game, allSheets).map(async (sheetName) => {
        const params = new URLSearchParams({ game: sheetName, startDate: dayStr, endDate: dayStr });
        const res = await fetch(`/api/sheet/installs?${params}`);
        const data = await res.json();
        if (data.error) throw new Error(`[${sheetName}] ${data.error}`);
        return normalizeSheetPayload(data, sheetName);
      }));

      const merged = mergeSheetPayloads(payloads);
      dailyResults.push({
        date: dayStr,
        total: Number(merged.total || 0),
        android: Number(merged.android?.total || 0),
        ios: Number(merged.ios?.total || 0),
        installs: Number(merged.installs || 0),
        reattribution: Number(merged.reattribution || 0),
        reengagement: Number(merged.reengagement || 0)
      });
    } catch (err) {
      dailyResults.push({ date: dayStr, total: 0, android: 0, ios: 0, installs: 0, reattribution: 0, reengagement: 0 });
    }
  }

  renderAfDailyTable(dailyResults, label);
}

function renderAfDailyTable(dailyResults, label) {
  if (!afDbWrapEl) return;
  let section = document.getElementById('afDbDailySection');
  if (!section) {
    section = document.createElement('div');
    section.id = 'afDbDailySection';
    section.className = 'af-daily-section';
    section.style.marginTop = '16px';
    afDbWrapEl.appendChild(section);
  }

  const heading = `<div style="font-weight:600;margin-bottom:8px;">日期明細（${escapeHtml(label || '')}）</div>`;

  if (!Array.isArray(dailyResults) || dailyResults.length === 0) {
    section.innerHTML = heading + '<div class="hint">無每日資料</div>';
    return;
  }

  const rowsHtml = dailyResults.map((d) => {
    const total = d.total || 0;
    return `
      <tbody class="daily-block">
        <tr style="background:#f7f2e6;font-weight:700;"><td>${d.date}</td><td>${total}</td><td>總覽</td><td>${d.android}</td><td>${d.ios}</td></tr>
        <tr><td></td><td>${d.installs}</td><td>安裝</td><td></td><td></td></tr>
        <tr><td></td><td>${d.reattribution}</td><td>再歸因</td><td></td><td></td></tr>
        <tr><td></td><td>${d.reengagement}</td><td>再互動</td><td></td><td></td></tr>
      </tbody>
    `;
  }).join('');

  // Note: we only have per-OS totals in merged payload; installs/reattribution/reengagement per-OS may be available in detailRows but to keep this simple we display totals and OS totals.
  section.innerHTML = heading + `
    <table style="width:100%;border-collapse:collapse;font-size:0.95rem;">
      <thead>
        <tr>
          <th style="text-align:left;padding:6px;border-bottom:1px solid #e0d6c0;">日期</th>
          <th style="text-align:right;padding:6px;border-bottom:1px solid #e0d6c0;">總計</th>
          <th style="text-align:left;padding:6px;border-bottom:1px solid #e0d6c0;">類型</th>
          <th style="text-align:right;padding:6px;border-bottom:1px solid #e0d6c0;">Android</th>
          <th style="text-align:right;padding:6px;border-bottom:1px solid #e0d6c0;">iOS</th>
        </tr>
      </thead>
      ${rowsHtml}
    </table>
  `;
}

function renderAfSheetResult(data, label, rangeLabel) {
  afDbWrapEl.hidden = false;
  const headingEl = afDbWrapEl.querySelector('.af-result-heading');
  const headingText = label ? `${label}　${rangeLabel || ''}` : (rangeLabel || '');
  if (headingEl) {
    headingEl.textContent = headingText;
  } else {
    const div = document.createElement('div');
    div.className = 'af-result-heading hint';
    div.style.cssText = 'margin-bottom:6px;font-weight:600;';
    div.textContent = headingText;
    afDbWrapEl.insertBefore(div, afDbWrapEl.firstChild);
  }
  afDbTbodyEl.innerHTML = `
    <tr><td>自然量 (Organic)</td><td>${data.organic}</td><td>${data.android?.organic ?? '-'}</td><td>${data.ios?.organic ?? '-'}</td></tr>
    <tr><td>付費量 (Non-organic)</td><td>${data.nonOrganic}</td><td>${data.android?.nonOrganic ?? '-'}</td><td>${data.ios?.nonOrganic ?? '-'}</td></tr>
    <tr><td>再歸因</td><td>${data.reattribution ?? 0}</td><td>${data.android?.reattribution ?? 0}</td><td>${data.ios?.reattribution ?? 0}</td></tr>
    <tr><td>再互動</td><td>${data.reengagement ?? 0}</td><td>${data.android?.reengagement ?? 0}</td><td>${data.ios?.reengagement ?? 0}</td></tr>
    <tr><td><strong>總下載</strong></td><td><strong>${data.total}</strong></td><td><strong>${data.android?.total ?? '-'}</strong></td><td><strong>${data.ios?.total ?? '-'}</strong></td></tr>
  `;

  const sources = Object.entries(data.bySource || {});
  if (sources.length > 0) {
    afDbSourceSectionEl.hidden = false;
    afDbSourceTbodyEl.innerHTML = sources
      .sort((a, b) => b[1].total - a[1].total)
      .map(([source, info]) =>
        `<tr><td>${source}</td><td>${info.total}</td><td>${info.android}</td><td>${info.ios}</td></tr>`
      ).join('');
  } else {
    afDbSourceSectionEl.hidden = true;
  }

  if (Array.isArray(data.detailRows) && data.detailRows.length > 0) {
    afDbDetailSectionEl.hidden = false;
    afDbDetailTbodyEl.innerHTML = data.detailRows
      .map((row) => `
        <tr>
          <td>${row.eventType}</td>
          <td>${row.mediaSource}</td>
          <td>${row.sourceName}</td>
          <td>${row.os}</td>
          <td>${row.count}</td>
        </tr>
      `)
      .join('');
  } else {
    afDbDetailSectionEl.hidden = true;
  }

  if (afMappedMetrics && afDbMappedSectionEl && afDbMappedTbodyEl) {
    afDbMappedSectionEl.hidden = false;
    const fbTotal = afMappedMetrics.fbAnd + afMappedMetrics.fbIos;
    const gaTotal = afMappedMetrics.gaAnd + afMappedMetrics.gaIos;
    const otherTotal = afMappedMetrics.otherAnd + afMappedMetrics.otherIos;
    afDbMappedTbodyEl.innerHTML = `
      <tr><td>Meta</td><td>${afMappedMetrics.fbAnd}</td><td>${afMappedMetrics.fbIos}</td><td>${fbTotal}</td></tr>
      <tr><td>Google</td><td>${afMappedMetrics.gaAnd}</td><td>${afMappedMetrics.gaIos}</td><td>${gaTotal}</td></tr>
      <tr><td>其他</td><td>${afMappedMetrics.otherAnd}</td><td>${afMappedMetrics.otherIos}</td><td>${otherTotal}</td></tr>
    `;
  } else if (afDbMappedSectionEl) {
    afDbMappedSectionEl.hidden = true;
  }
}

function tableToTsv(tableEl) {
  const rows = Array.from(tableEl.querySelectorAll('tr'));
  return rows.map((tr) => {
    const cells = Array.from(tr.querySelectorAll('td'));
    return cells.map((cell) => {
      const text = cell.textContent.trim().replace(/\t|\n/g, ' ');
      const span = parseInt(cell.getAttribute('colspan') || '1', 10);
      return [text, ...Array(span - 1).fill('')].join('\t');
    }).join('\t');
  }).join('\n');
}

async function copyTableToClipboard(tableEl) {
  if (!tableEl) return false;
  const tsv = tableToTsv(tableEl);
  const htmlStr = `<meta charset="utf-8">${tableEl.outerHTML}`;
  try {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([htmlStr], { type: 'text/html' }),
        'text/plain': new Blob([tsv], { type: 'text/plain' })
      })
    ]);
    return true;
  } catch (_) {
    try {
      await navigator.clipboard.writeText(tsv);
      return true;
    } catch (_2) {
      return false;
    }
  }
}

function renderCopySheetView() {
  const combinedRows = [...platformRows.meta, ...platformRows.google];
  const metrics = buildIntegratedSheetMetrics(combinedRows);

  if (!metrics) {
    copySheetStatusEl.textContent = '請先到 Meta 與 Google 分頁各自查詢資料。';
    copySheetTableWrapEl.innerHTML = '';
    return;
  }

  if (!afMappedMetrics) {
    copySheetStatusEl.textContent = '⚠ 尚未載入 AF 下載資料，人數與成本欄位顯示「-」。請先到「投放簡表」頁面抓取 AF 下載數。';
  } else {
    const andT = afTotalCounts?.andTotal ?? 0;
    const iosT = afTotalCounts?.iosTotal ?? 0;
    copySheetStatusEl.textContent = `AF 資料已載入　Android 總下載：${nf.format(andT)}　iOS 總下載：${nf.format(iosT)}（上表人數已含自然量 + 付費量）`;
  }

  const offAnd = Math.max(0, Number(copySheetOfflineAndEl.value) || 0);
  const rawOffIos = copySheetOfflineIosEl.value.trim();
  const offIos = rawOffIos === '' ? null : Math.max(0, Number(rawOffIos) || 0);
  const offIosNum = offIos !== null ? offIos : 0;
  const dateLabel = copySheetDateLabelEl.value.trim() || metrics.rangeText;

  const fmtVal = (val) => {
    if (val === null || val === undefined) return '-';
    if (typeof val === 'number' && !isFinite(val)) return '-';
    return nf.format(Math.round(val));
  };

  const hasAfCounts = Boolean(afMappedMetrics && afTotalCounts);
  const gaAndBase = hasAfCounts ? Number(afMappedMetrics.gaAnd || 0) : 0;
  const gaIosBase = hasAfCounts ? Number(afMappedMetrics.gaIos || 0) : 0;
  const gaAndRes = hasAfCounts ? Math.max(0, Math.min(Number(afTotalCounts.andTotal || 0), gaAndBase)) : null;
  const gaIosRes = hasAfCounts ? Math.max(0, Math.min(Number(afTotalCounts.iosTotal || 0), gaIosBase)) : null;
  const fbAndRes = hasAfCounts ? Math.max(0, Number(afTotalCounts.andTotal || 0) - gaAndRes) : null;
  const fbIosRes = hasAfCounts ? Math.max(0, Number(afTotalCounts.iosTotal || 0) - gaIosRes) : null;

  const safeCpa = (spend, results) => (results ? spend / results : null);
  const fbAndCpa = safeCpa(metrics.fbAnd.spendTwd, fbAndRes);
  const fbIosCpa = safeCpa(metrics.fbIos.spendTwd, fbIosRes);
  const gaAndCpa = safeCpa(metrics.gaAnd.spendTwd, gaAndRes);
  const gaIosCpa = safeCpa(metrics.gaIos.spendTwd, gaIosRes);

  const andTotalInstalls = hasAfCounts ? (fbAndRes + gaAndRes) : null;
  const iosTotalInstalls = hasAfCounts ? (fbIosRes + gaIosRes) : null;
  const andPlatformCpa = safeCpa(metrics.andSpend, andTotalInstalls);
  const iosPlatformCpa = safeCpa(metrics.iosSpend, iosTotalInstalls);

  const grandTotalSpend = metrics.andSpend + metrics.iosSpend + offAnd + offIosNum;
  const grandTotalResults = (andTotalInstalls ?? 0) + (iosTotalInstalls ?? 0);
  const grandTotalCpa = grandTotalResults > 0 ? grandTotalSpend / grandTotalResults : null;

  const S = {
    hdr:     'background:#ddb83a;color:#2d2200;font-weight:800;text-align:center;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    sub:     'background:#f0d060;color:#2d2200;font-weight:700;text-align:center;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    lbl:     'background:#f5edbc;font-weight:600;text-align:center;min-width:80px;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    corner:  'background:#f0e8cc;border:1px solid #b09450;padding:5px 13px;',
    data:    'background:#fffef4;min-width:75px;text-align:center;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    totLbl:  'background:#fef3b0;font-weight:700;text-align:center;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    totData: 'background:#fef3b0;font-weight:700;text-align:center;border:1px solid #b09450;padding:5px 13px;white-space:nowrap;',
    totEmp:  'background:#d8d6cc;border:1px solid #b09450;padding:5px 13px;',
  };

  copySheetTableWrapEl.innerHTML = `
    <table id="excelMergedTable" style="border-collapse:collapse;font-size:0.875rem;">
      <tbody>
        <tr>
          <td rowspan="2" style="${S.hdr}">${escapeHtml(dateLabel)}</td>
          <td colspan="2" style="${S.hdr}">FB</td>
          <td colspan="2" style="${S.hdr}">GA</td>
        </tr>
        <tr>
          <td style="${S.sub}">AND</td>
          <td style="${S.sub}">IOS</td>
          <td style="${S.sub}">AND</td>
          <td style="${S.sub}">IOS</td>
        </tr>
        <tr>
          <td style="${S.lbl}">花費</td>
          <td style="${S.data}">${fmtVal(metrics.fbAnd.spendTwd)}</td>
          <td style="${S.data}">${fmtVal(metrics.fbIos.spendTwd)}</td>
          <td style="${S.data}">${fmtVal(metrics.gaAnd.spendTwd)}</td>
          <td style="${S.data}">${fmtVal(metrics.gaIos.spendTwd)}</td>
        </tr>
        <tr>
          <td style="${S.lbl}">人數(AF)</td>
          <td style="${S.data}">${fmtVal(fbAndRes)}</td>
          <td style="${S.data}">${fmtVal(fbIosRes)}</td>
          <td style="${S.data}">${fmtVal(gaAndRes)}</td>
          <td style="${S.data}">${fmtVal(gaIosRes)}</td>
        </tr>
        <tr>
          <td style="${S.lbl}">成本</td>
          <td style="${S.data}">${fmtVal(fbAndCpa)}</td>
          <td style="${S.data}">${fmtVal(fbIosCpa)}</td>
          <td style="${S.data}">${fmtVal(gaAndCpa)}</td>
          <td style="${S.data}">${fmtVal(gaIosCpa)}</td>
        </tr>
        <tr>
          <td style="${S.corner}"></td>
          <td colspan="2" style="${S.sub}">AND</td>
          <td colspan="2" style="${S.sub}">IOS</td>
        </tr>
        <tr>
          <td style="${S.lbl}">花費</td>
          <td colspan="2" style="${S.data}">${fmtVal(metrics.andSpend)}</td>
          <td colspan="2" style="${S.data}">${fmtVal(metrics.iosSpend)}</td>
        </tr>
        <tr>
          <td style="${S.lbl}">總人數</td>
          <td colspan="2" style="${S.data}">${fmtVal(andTotalInstalls)}</td>
          <td colspan="2" style="${S.data}">${fmtVal(iosTotalInstalls)}</td>
        </tr>
        <tr>
          <td style="${S.lbl}">平台成本</td>
          <td colspan="2" style="${S.data}">${fmtVal(andPlatformCpa)}</td>
          <td colspan="2" style="${S.data}">${fmtVal(iosPlatformCpa)}</td>
        </tr>
        <tr>
          <td style="${S.lbl}">線下花費</td>
          <td colspan="2" style="${S.data}">${fmtVal(offAnd)}</td>
          <td colspan="2" style="${S.data}">${offIos !== null ? fmtVal(offIos) : ''}</td>
        </tr>
        <tr>
          <td style="${S.totLbl}">總花費</td>
          <td colspan="4" style="${S.totData}">${fmtVal(grandTotalSpend)}</td>
        </tr>
        <tr>
          <td style="${S.totLbl}">總人數</td>
          <td colspan="4" style="${S.totData}">${grandTotalResults > 0 ? nf.format(grandTotalResults) : '-'}</td>
        </tr>
        <tr>
          <td style="${S.totLbl}">總成本</td>
          <td colspan="4" style="${S.totData}">${fmtVal(grandTotalCpa)}</td>
        </tr>
      </tbody>
    </table>
  `;
}


function setActiveTab(tab) {
  activeTab = tab;

  tabMetaEl.classList.toggle('active', tab === 'meta');
  tabGoogleEl.classList.toggle('active', tab === 'google');
  tabIntegratedEl.classList.toggle('active', tab === 'integrated');
  tabIntegrated2El.classList.toggle('active', tab === 'integrated2');
  tabCopySheetEl.classList.toggle('active', tab === 'copySheet');

  queryPanelEl.hidden = tab === 'integrated' || tab === 'integrated2' || tab === 'copySheet';
  integratedPanelEl.hidden = tab !== 'integrated';
  integratedSheetPanelEl.hidden = tab !== 'integrated2';
  copySheetPanelEl.hidden = tab !== 'copySheet';

  if (tab === 'integrated') {
    renderIntegratedView();
    return;
  }

  if (tab === 'copySheet') {
    renderCopySheetView();
    return;
  }

  if (tab === 'integrated2') {
    renderIntegratedSheetView();
    if (!afDbSinceEl.value) {
      const combinedRows2 = [...platformRows.meta, ...platformRows.google];
      const sheetMetrics = buildIntegratedSheetMetrics(combinedRows2);
      if (sheetMetrics) {
        const match = sheetMetrics.rangeText.match(/(\d{4}-\d{2}-\d{2}) ~ (\d{4}-\d{2}-\d{2})/);
        if (match) { afDbSinceEl.value = match[1]; afDbUntilEl.value = match[2]; }
      }
    }
    if (afDbGameSelectEl.options.length <= 1) {
      loadAfSheetList();
    }
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
    copySheetDateLabelEl.value = new Date().toLocaleString('zh-TW', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false
    }).replace(/\//g, '/').replace(',', '');
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
tabCopySheetEl.addEventListener('click', () => setActiveTab('copySheet'));
copySheetCopyBtnEl.addEventListener('click', async () => {
  const tableEl = document.getElementById('excelMergedTable');
  const ok = await copyTableToClipboard(tableEl);
  copySheetStatusEl.textContent = ok ? '✓ 報表已複製，可直接貼到 Excel！' : '複製失敗，請手動選取表格。';
  if (ok) setTimeout(() => { copySheetStatusEl.textContent = ''; }, 3000);
});
copySheetOfflineAndEl.addEventListener('input', renderCopySheetView);
copySheetOfflineIosEl.addEventListener('input', renderCopySheetView);
copySheetDateLabelEl.addEventListener('input', renderCopySheetView);
if (integratedSheetSaveBtnEl) {
  integratedSheetSaveBtnEl.addEventListener('click', saveIntegratedSheetSnapshot);
}
afDbReloadBtnEl.addEventListener('click', loadAfSheetList);
afDbFetchBtnEl.addEventListener('click', fetchAfSheetData);
if (integratedSheetNotesEl) {
  integratedSheetNotesEl.addEventListener('click', (event) => {
    const deleteBtn = event.target.closest('.sticky-note-delete');
    if (!deleteBtn) {
      return;
    }

    deleteIntegratedSheetSnapshot(deleteBtn.dataset.snapshotId);
  });
}
accountSelectEl.addEventListener('change', () => {
  setCurrentSelectedAccountId(accountSelectEl.value);
  saveSettings();
  syncAccountForm();
});
addPageBtn.addEventListener('click', upsertAccount);
removePageBtn.addEventListener('click', removeCurrentAccount);
savePresetBtn.addEventListener('click', saveAccountPresetsToServer);
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
hydrateSettingsFromServerPresets();
