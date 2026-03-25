const axios = require('axios');

try {
  require('dotenv').config();
} catch (error) {
  // Ignore when dotenv is not available.
}

const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

function parseArgs(argv) {
  const result = {
    game: '',
    startDate: '',
    endDate: ''
  };

  for (let i = 0; i < argv.length; i += 1) {
    const current = argv[i];
    const next = argv[i + 1];

    if (current === '--game' && next) {
      result.game = next;
      i += 1;
      continue;
    }

    if (current === '--start' && next) {
      result.startDate = next;
      i += 1;
      continue;
    }

    if (current === '--end' && next) {
      result.endDate = next;
      i += 1;
      continue;
    }
  }

  return result;
}

function inferSource(mediaSource, campaign) {
  const ms = String(mediaSource || '').toLowerCase();
  const cp = String(campaign || '').toLowerCase();
  const text = `${ms} ${cp}`;

  if (text.includes('facebook') || text.includes('meta') || text.includes('fb_')) {
    return 'Meta';
  }
  if (text.includes('google') || text.includes('adwords') || text.includes('uac')) {
    return 'Google';
  }
  if (text.includes('sms')) {
    return 'SMS';
  }
  if (ms === 'restricted') {
    return 'Restricted(Unknown)';
  }
  if (ms === 'organic') {
    return 'Organic';
  }
  return 'Other';
}

async function fetchSheetList() {
  const response = await axios.get(APPS_SCRIPT_URL, { timeout: 20000 });
  const availableSheets = response.data?.availableSheets;
  if (!Array.isArray(availableSheets)) {
    return [];
  }
  return availableSheets;
}

async function fetchSheetData({ game, startDate, endDate }) {
  const params = new URLSearchParams();
  if (game) params.set('game', game);
  if (startDate) params.set('startDate', startDate);
  if (endDate) params.set('endDate', endDate);

  const url = `${APPS_SCRIPT_URL}?${params.toString()}`;
  const response = await axios.get(url, { timeout: 20000 });
  return response.data;
}

function printCoreSummary(data) {
  console.log('=== Summary ===');
  console.log('game:', data.game || 'N/A');
  console.log('total:', Number(data.total || 0));
  console.log('organic:', Number(data.organic || 0));
  console.log('nonOrganic:', Number(data.nonOrganic || 0));
  console.log('android.total:', Number(data.android?.total || 0));
  console.log('ios.total:', Number(data.ios?.total || 0));
  console.log('');
}

function printBySourceSummary(data) {
  const entries = Object.entries(data.bySource || {}).sort((left, right) => {
    return Number(right[1]?.total || 0) - Number(left[1]?.total || 0);
  });

  console.log('=== bySource ===');
  if (!entries.length) {
    console.log('No bySource data.');
    console.log('');
    return;
  }

  entries.forEach(([source, info]) => {
    console.log(
      `${source}: total=${Number(info?.total || 0)}, android=${Number(info?.android || 0)}, ios=${Number(info?.ios || 0)}`
    );
  });
  console.log('');
}

function printDetailDiagnostics(data) {
  const details = Array.isArray(data.details) ? data.details : [];
  console.log('=== details diagnostics ===');
  console.log('details row count:', details.length);

  if (!details.length) {
    console.log('No details found. Current Apps Script response may only include aggregated data.');
    console.log('To verify hidden fields, include details/raw payload in Apps Script doGet response.');
    console.log('');
    return;
  }

  const keySet = new Set();
  details.forEach((row) => Object.keys(row || {}).forEach((key) => keySet.add(key)));
  console.log('detail keys:', Array.from(keySet).sort().join(', '));

  const restrictedRows = details.filter((row) => String(row.mediaSource || '').toLowerCase() === 'restricted');
  const restrictedWithCampaign = restrictedRows.filter((row) => {
    const campaign = String(row.campaign || row.rawCampaign || '').trim();
    return Boolean(campaign) && campaign.toLowerCase() !== 'n/a';
  });

  console.log('restricted count:', restrictedRows.length);
  console.log('restricted with campaign-like value:', restrictedWithCampaign.length);

  const inferred = {};
  details.forEach((row) => {
    const source = inferSource(row.mediaSource, row.campaign || row.rawCampaign);
    const count = Number(row.count || 1);
    inferred[source] = (inferred[source] || 0) + (Number.isFinite(count) ? count : 1);
  });

  console.log('inferred source totals:', inferred);
  console.log('');
}

async function main() {
  if (!APPS_SCRIPT_URL) {
    throw new Error('APPS_SCRIPT_URL is missing in .env');
  }

  const args = parseArgs(process.argv.slice(2));
  let selectedGame = args.game;

  if (!selectedGame) {
    const availableSheets = await fetchSheetList();
    if (!availableSheets.length) {
      throw new Error('No available sheets returned. Please pass --game manually.');
    }
    selectedGame = availableSheets[0];
    console.log(`No --game provided, using first sheet: ${selectedGame}`);
  }

  const data = await fetchSheetData({
    game: selectedGame,
    startDate: args.startDate,
    endDate: args.endDate
  });

  if (data.error) {
    throw new Error(data.error);
  }

  printCoreSummary(data);
  printBySourceSummary(data);
  printDetailDiagnostics(data);

  console.log('Done.');
}

main().catch((error) => {
  console.error('[test_af_source] failed:', error.message);
  process.exit(1);
});
