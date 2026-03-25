(function attachAfHelpers(global) {
  function getProjectKeyFromSheetName(sheetName) {
    return String(sheetName || '')
      .replace(/[_\-\s]?(android|aos|ios)$/i, '')
      .trim();
  }

  function resolveCompanionSheets(selectedSheet, allSheets) {
    const current = String(selectedSheet || '').trim();
    if (!current) {
      return [];
    }

    const rows = Array.isArray(allSheets) ? allSheets.filter(Boolean) : [];
    const lowerRows = rows.map((name) => String(name).toLowerCase());
    const addUnique = (arr, value) => {
      if (value && !arr.includes(value)) arr.push(value);
    };

    const result = [];
    addUnique(result, current);

    const toMatch = current.match(/^(.*?)([_\-\s]?)(android|aos|ios)$/i);
    if (!toMatch) {
      return result;
    }

    const prefix = toMatch[1];
    const connector = toMatch[2] || '_';
    const suffix = toMatch[3].toLowerCase();
    const targetSuffix = suffix === 'ios' ? 'android' : 'ios';
    const candidate = `${prefix}${connector}${targetSuffix}`.toLowerCase();
    const foundIdx = lowerRows.findIndex((name) => name === candidate);
    if (foundIdx >= 0) {
      addUnique(result, rows[foundIdx]);
    }

    // Additional compatibility for Android <-> AOS naming differences.
    if (targetSuffix === 'android') {
      const aosCandidate = `${prefix}${connector}aos`.toLowerCase();
      const aosIdx = lowerRows.findIndex((name) => name === aosCandidate);
      if (aosIdx >= 0) {
        addUnique(result, rows[aosIdx]);
      }
    }

    return result;
  }

  function detectSheetOs(sheetName) {
    const text = String(sheetName || '').toLowerCase();
    if (text.includes('android') || text.includes('aos') || text.includes('安卓')) {
      return 'android';
    }
    if (text.includes('ios') || text.includes('iphone') || text.includes('ipad')) {
      return 'ios';
    }
    return '';
  }

  function normalizeAfDetailOs(value, sheetName) {
    const text = String(value || '').trim().toLowerCase();
    if (text.includes('android') || text.includes('aos') || text.includes('google play') || text.includes('安卓')) {
      return 'Android';
    }
    if (text.includes('ios') || text.includes('iphone') || text.includes('ipad') || text.includes('apple')) {
      return 'iOS';
    }

    const fallback = detectSheetOs(sheetName);
    if (fallback === 'android') {
      return 'Android';
    }
    if (fallback === 'ios') {
      return 'iOS';
    }
    return 'N/A';
  }

  function normalizeAfEventType(value) {
    const text = String(value || '').trim().toLowerCase();
    if (!text) {
      return 'install';
    }
    if (text === 'reattribution' || text === 're_attribution') {
      return 're-attribution';
    }
    return text;
  }

  function isPaidMediaEventType(eventType) {
    const normalized = normalizeAfEventType(eventType);
    return normalized === 'install' || normalized === 're-attribution';
  }

  function collectAfDetailRows(data, sheetName) {
    const rawRows = Array.isArray(data.details)
      ? data.details
      : Array.isArray(data.rows)
        ? data.rows
        : Array.isArray(data.records)
          ? data.records
          : [];

    const grouped = new Map();
    rawRows.forEach((item) => {
      const eventType = normalizeAfEventType(item?.eventType || item?.type || item?.installType || item?.conversionType || item?.actionType);
      const mediaSource = String(item?.mediaSource || item?.source || item?.media || item?.media_source || 'N/A').trim() || 'N/A';
      const rawCampaign = String(item?.campaign || item?.campaignName || item?.campaign_name || item?.name || item?.label || '').trim();
      const sourceName = /^(restricted|organic)$/i.test(mediaSource) ? 'N/A' : (rawCampaign || 'N/A');
      const os = normalizeAfDetailOs(item?.os || item?.platform || item?.deviceOs || item?.store, sheetName);
      const count = Number(item?.count || item?.total || item?.value || 1);

      if (!Number.isFinite(count) || count <= 0) {
        return;
      }

      const key = [eventType, mediaSource, sourceName, os].join('|');
      const current = grouped.get(key) || {
        eventType,
        mediaSource,
        sourceName,
        os,
        count: 0
      };

      current.count += count;
      grouped.set(key, current);
    });

    return Array.from(grouped.values()).sort((left, right) => {
      if (left.mediaSource !== right.mediaSource) {
        return left.mediaSource.localeCompare(right.mediaSource, 'zh-Hant');
      }
      if (left.sourceName !== right.sourceName) {
        return left.sourceName.localeCompare(right.sourceName, 'zh-Hant');
      }
      if (left.eventType !== right.eventType) {
        return left.eventType.localeCompare(right.eventType, 'zh-Hant');
      }
      return left.os.localeCompare(right.os, 'zh-Hant');
    });
  }

  function normalizeSheetPayload(data, sheetName) {
    const detailRows = collectAfDetailRows(data, sheetName);
    const payload = {
      total: Number(data.total || 0),
      organic: Number(data.organic || 0),
      nonOrganic: Number(data.nonOrganic || 0),
      android: {
        total: Number(data.android?.total || 0),
        organic: Number(data.android?.organic || 0),
        nonOrganic: Number(data.android?.nonOrganic || 0)
      },
      ios: {
        total: Number(data.ios?.total || 0),
        organic: Number(data.ios?.organic || 0),
        nonOrganic: Number(data.ios?.nonOrganic || 0)
      },
      bySource: Object.fromEntries(Object.entries(data.bySource || {}).map(([k, v]) => [k, {
        total: Number(v?.total || 0),
        android: Number(v?.android || 0),
        ios: Number(v?.ios || 0)
      }])),
      detailRows
    };

    const os = detectSheetOs(sheetName);
    const osEmpty = payload.android.total === 0 && payload.ios.total === 0;
    if (payload.total > 0 && os && osEmpty) {
      payload[os].total = payload.total;
      payload[os].organic = payload.organic;
      payload[os].nonOrganic = payload.nonOrganic;
      Object.values(payload.bySource).forEach((item) => {
        if (item.android === 0 && item.ios === 0 && item.total > 0) {
          item[os] = item.total;
        }
      });
    }

    return payload;
  }

  function mergeSheetPayloads(payloads) {
    const merged = {
      total: 0,
      organic: 0,
      nonOrganic: 0,
      android: { total: 0, organic: 0, nonOrganic: 0 },
      ios: { total: 0, organic: 0, nonOrganic: 0 },
      bySource: {},
      detailRows: []
    };

    payloads.forEach((data) => {
      merged.total += Number(data.total || 0);
      merged.organic += Number(data.organic || 0);
      merged.nonOrganic += Number(data.nonOrganic || 0);
      merged.android.total += Number(data.android?.total || 0);
      merged.android.organic += Number(data.android?.organic || 0);
      merged.android.nonOrganic += Number(data.android?.nonOrganic || 0);
      merged.ios.total += Number(data.ios?.total || 0);
      merged.ios.organic += Number(data.ios?.organic || 0);
      merged.ios.nonOrganic += Number(data.ios?.nonOrganic || 0);

      Object.entries(data.bySource || {}).forEach(([source, info]) => {
        if (!merged.bySource[source]) {
          merged.bySource[source] = { total: 0, android: 0, ios: 0 };
        }
        merged.bySource[source].total += Number(info?.total || 0);
        merged.bySource[source].android += Number(info?.android || 0);
        merged.bySource[source].ios += Number(info?.ios || 0);
      });

      merged.detailRows.push(...(Array.isArray(data.detailRows) ? data.detailRows : []));
    });

    return merged;
  }

  function mapAfSourcesToPlatforms(mergedPayload) {
    const mapped = { fbAnd: 0, fbIos: 0, gaAnd: 0, gaIos: 0, otherAnd: 0, otherIos: 0 };
    const paidAnd = Math.max(0, Number(mergedPayload?.android?.nonOrganic || 0));
    const paidIos = Math.max(0, Number(mergedPayload?.ios?.nonOrganic || 0));

    let googleAnd = 0;

    const detailRows = Array.isArray(mergedPayload?.detailRows) ? mergedPayload.detailRows : [];
    if (detailRows.length) {
      detailRows
        .filter((row) => isPaidMediaEventType(row.eventType))
        .forEach((row) => {
          const key = String(row.mediaSource || '').toLowerCase();
          const count = Number(row.count || 0);
          if (!Number.isFinite(count) || count <= 0) {
            return;
          }
          if (key.includes('google') || key.includes('adwords') || key.includes('uac')) {
            if (row.os === 'Android') {
              googleAnd += count;
            }
          }
        });
    } else {
      Object.entries(mergedPayload?.bySource || {}).forEach(([source, info]) => {
        const key = String(source || '').toLowerCase();
        if (key.includes('google') || key.includes('adwords') || key.includes('uac')) {
          googleAnd += Number(info?.android || 0);
        }
      });
    }

    mapped.gaAnd = Math.min(paidAnd, Math.max(0, Math.round(googleAnd)));
    mapped.fbAnd = Math.max(0, paidAnd - mapped.gaAnd);
    mapped.fbIos = Math.max(0, paidIos);
    mapped.gaIos = 0;

    return mapped;
  }

  global.AppAf = {
    getProjectKeyFromSheetName,
    resolveCompanionSheets,
    detectSheetOs,
    normalizeAfDetailOs,
    normalizeAfEventType,
    isPaidMediaEventType,
    collectAfDetailRows,
    normalizeSheetPayload,
    mergeSheetPayloads,
    mapAfSourcesToPlatforms
  };
}(window));
