const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const cwd = process.cwd();
const forecastYear = 2026;
const requiredHeaders = ['UNIT', 'NDA ITEM DESCRIPTION', 'INVOICE DATE', 'ORDERED'];
const historyYears = Array.from({ length: 6 }, (_, index) => forecastYear - 6 + index);
const weatherRainThresholdInches = 0.1;
const weatherFallbackDropRate = 0.1;
const weatherCacheVersion = 2;
const weatherAnalogYear = forecastYear - 1;

const stateNames = {
  AL: 'Alabama',
  AK: 'Alaska',
  AZ: 'Arizona',
  AR: 'Arkansas',
  CA: 'California',
  CO: 'Colorado',
  CT: 'Connecticut',
  DE: 'Delaware',
  FL: 'Florida',
  GA: 'Georgia',
  HI: 'Hawaii',
  ID: 'Idaho',
  IL: 'Illinois',
  IN: 'Indiana',
  IA: 'Iowa',
  KS: 'Kansas',
  KY: 'Kentucky',
  LA: 'Louisiana',
  ME: 'Maine',
  MD: 'Maryland',
  MA: 'Massachusetts',
  MI: 'Michigan',
  MN: 'Minnesota',
  MS: 'Mississippi',
  MO: 'Missouri',
  MT: 'Montana',
  NE: 'Nebraska',
  NV: 'Nevada',
  NH: 'New Hampshire',
  NJ: 'New Jersey',
  NM: 'New Mexico',
  NY: 'New York',
  NC: 'North Carolina',
  ND: 'North Dakota',
  OH: 'Ohio',
  OK: 'Oklahoma',
  OR: 'Oregon',
  PA: 'Pennsylvania',
  RI: 'Rhode Island',
  SC: 'South Carolina',
  SD: 'South Dakota',
  TN: 'Tennessee',
  TX: 'Texas',
  UT: 'Utah',
  VT: 'Vermont',
  VA: 'Virginia',
  WA: 'Washington',
  WV: 'West Virginia',
  WI: 'Wisconsin',
  WY: 'Wyoming',
  DC: 'District of Columbia',
};

const locationAliases = {
  'CA||TERRA VISTA TOWN CENTER': 'Rancho Cucamonga',
  'CA||WESTGATE WEST': 'San Jose',
  'CA||FLOWER HILL DEL MAR': 'Del Mar',
  'CA||PAVILION AT LA QUINTA': 'La Quinta',
};

function normalizeHeader(value) {
  return String(value ?? '')
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function escapeCsv(value) {
  const s = String(value ?? '');
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeCsv(filePath, rows, headers) {
  const lines = [headers.map(escapeCsv).join(',')];
  for (const row of rows) {
    lines.push(headers.map((h) => escapeCsv(row[h])).join(','));
  }
  fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
}

function getStoreCode(unit) {
  const text = String(unit || '');
  let match = text.match(/(?<!\d)(\d{3,5})(?!\d)/);
  if (match) return match[1];
  match = text.match(/(?<![A-Z0-9])([A-Z]{1,3}\d{2,5}[A-Z0-9]*)(?![A-Z0-9])/);
  return match ? match[1] : '';
}

function weekOfMonth(date) {
  return Math.floor((date.getDate() - 1) / 7) + 1;
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, '0')}`;
}

function monthName(monthKeyValue) {
  const [year, month] = String(monthKeyValue).split('-').map(Number);
  return new Date(year, month - 1, 1).toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function round4(value) {
  return Number(Number(value || 0).toFixed(4));
}

function round2(value) {
  return Number(Number(value || 0).toFixed(2));
}

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function dateKey(year, month, day) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function weekDateKeys(year, month, week) {
  const startDay = (week - 1) * 7 + 1;
  const endDay = week < 5 ? Math.min(week * 7, daysInMonth(year, month)) : daysInMonth(year, month);
  const days = [];
  for (let day = startDay; day <= endDay; day += 1) {
    days.push(dateKey(year, month, day));
  }
  return days;
}

function normalizeItemName(value) {
  return normalizeHeader(value);
}

function uniqueValues(values) {
  return Array.from(new Set(values.filter((value) => String(value || '').trim()).map((value) => String(value).trim())));
}

function readClassificationData(files) {
  const emptyClassification = {
    source_file: '',
    items: new Map(),
    categories: [],
    flavor_types: [],
  };

  for (const file of [...files].reverse()) {
    const fullPath = path.join(cwd, file);
    let sheetNames;
    try {
      sheetNames = XLSX.readFile(fullPath, { bookSheets: true }).SheetNames;
    } catch (error) {
      continue;
    }

    if (!sheetNames.includes('Categorization')) continue;

    const workbook = XLSX.readFile(fullPath, { sheetRows: 1000 });
    const categorizationSheet = workbook.Sheets.Categorization;
    if (!categorizationSheet) continue;

    const categorizationRows = XLSX.utils.sheet_to_json(categorizationSheet, { header: 1, raw: false, defval: '' });
    if (!categorizationRows.length) continue;

    const header = categorizationRows[0].map(normalizeHeader);
    const itemIndex = header.indexOf('NDA ITEM DESCRIPTION');
    const categoryIndex = header.indexOf('CATEGORY');
    const flavorTypeIndex = header.indexOf('FLAVOR TYPE');
    const seasonalIndex = header.findIndex((value) => value === 'SEASONAL');
    const fallbackSeasonalIndex = Math.max(itemIndex, categoryIndex, flavorTypeIndex) + 1;

    if (itemIndex < 0) continue;

    const items = new Map();
    for (let rowIndex = 1; rowIndex < categorizationRows.length; rowIndex += 1) {
      const row = categorizationRows[rowIndex];
      const item = String(row[itemIndex] || '').trim();
      if (!item) continue;

      const category = categoryIndex >= 0 ? String(row[categoryIndex] || '').trim() : '';
      const flavorType = flavorTypeIndex >= 0 ? String(row[flavorTypeIndex] || '').trim() : '';
      const seasonal = seasonalIndex >= 0
        ? String(row[seasonalIndex] || '').trim()
        : String(row[fallbackSeasonalIndex] || '').trim();

      items.set(normalizeItemName(item), {
        item_group: 'Flavor',
        category: category || 'Uncategorized',
        flavor_type: flavorType || 'Uncategorized',
        seasonal,
      });
    }

    const listRows = workbook.Sheets.List
      ? XLSX.utils.sheet_to_json(workbook.Sheets.List, { header: 1, raw: false, defval: '' })
      : [];
    const listHeader = (listRows[0] || []).map(normalizeHeader);
    const listCategoryIndex = listHeader.indexOf('CATEGORY');
    const listFlavorTypeIndex = listHeader.indexOf('FLAVOR TYPE');
    const categories = uniqueValues([
      ...listRows.slice(1).map((row) => (listCategoryIndex >= 0 ? row[listCategoryIndex] : '')),
      ...Array.from(items.values()).map((item) => item.category),
      'Topping',
    ]);
    const flavorTypes = uniqueValues([
      ...listRows.slice(1).map((row) => (listFlavorTypeIndex >= 0 ? row[listFlavorTypeIndex] : '')),
      ...Array.from(items.values()).map((item) => item.flavor_type),
      'Topping',
    ]);

    return {
      source_file: file,
      items,
      categories,
      flavor_types: flavorTypes,
    };
  }

  return emptyClassification;
}

function classifyItem(description, classification) {
  const normalized = normalizeItemName(description);
  const knownFlavor = classification.items.get(normalized);
  if (knownFlavor) return knownFlavor;

  if (/^TOPPING\b/.test(normalized)) {
    return {
      item_group: 'Topping',
      category: 'Topping',
      flavor_type: 'Topping',
      seasonal: '',
    };
  }

  if (/^(YOGURT|SORBET|YL)\b/.test(normalized)) {
    return {
      item_group: 'Flavor',
      category: 'Uncategorized',
      flavor_type: 'Uncategorized',
      seasonal: '',
    };
  }

  return {
    item_group: 'Other',
    category: 'Uncategorized',
    flavor_type: 'Uncategorized',
    seasonal: '',
  };
}

function extractStateFromStoreToken(token) {
  const text = String(token || '').toUpperCase();
  let match = text.match(/^([A-Z]{2})\d/);
  if (match && stateNames[match[1]]) return match[1];
  match = text.match(/\d([A-Z]{2})\d*$/);
  if (match && stateNames[match[1]]) return match[1];
  match = text.match(/([A-Z]{2})/);
  if (match && stateNames[match[1]]) return match[1];
  return '';
}

function parseStoreLocation(unit) {
  const text = String(unit || '').replace(/^YOGURTLAND\s+/i, '').trim();
  const parts = text.split(/\s+/).filter(Boolean);
  let storeToken = parts.shift() || '';
  if (storeToken && !/\d/.test(storeToken)) {
    parts.unshift(storeToken);
    storeToken = '';
  }
  const stateCode = extractStateFromStoreToken(storeToken);
  const area = parts.join(' ').replace(/\s+/g, ' ').trim();
  const alias = locationAliases[`${stateCode}||${area.toUpperCase()}`] || '';
  const place = alias || area;
  const stateName = stateNames[stateCode] || stateCode;
  const query = [place, stateName, 'United States'].filter(Boolean).join(', ');
  return {
    unit,
    store_token: storeToken,
    state_code: stateCode,
    state_name: stateName,
    area,
    query,
  };
}

function locationCandidates(location) {
  const alias = locationAliases[`${location.state_code}||${location.area.toUpperCase()}`] || '';
  const normalizedArea = location.area
    .replace(/[().]/g, ' ')
    .replace(/&/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const withoutCenterWords = normalizedArea
    .replace(/\b(TOWN CENTER|SHOPPING CENTER|MARKETPLACE|MARKET PLACE|TOWNE SQUARE|TOWN SQUARE|MALL|PLAZA|PAVILION AT|PAVILLION|PAVILION|CENTER|CENTRE|VILLAGE|COMMONS|PROMENADE|GATEWAY|CROSSROADS|RIVERWALK|RIVER PARK|SPECTRUM|DOWNTOWN)\b/gi, ' ')
    .replace(/\b(BLVD|BOULEVARD|STREET|ST|AVENUE|AVE|ROAD|RD|DRIVE|DR|LANE|LN|FOOTHILL|CAMPUS|STATE|UNIVERSITY)\b/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const withoutTrailingDirection = withoutCenterWords
    .replace(/\s+\b(WEST|EAST|NORTH|SOUTH)\b$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  const atMatch = normalizedArea.match(/\bAT\s+(.+)$/i);
  const afterAt = atMatch ? atMatch[1].replace(/\s+/g, ' ').trim() : '';
  const words = withoutTrailingDirection.split(/\s+/).filter(Boolean);
  const firstOne = words.slice(0, 1).join(' ');
  const firstTwo = words.slice(0, 2).join(' ');
  const firstThree = words.slice(0, 3).join(' ');
  const lastOne = words.slice(-1).join(' ');
  const lastTwo = words.slice(-2).join(' ');
  const lastThree = words.slice(-3).join(' ');
  return uniqueValues([
    alias,
    normalizedArea,
    withoutCenterWords,
    withoutTrailingDirection,
    afterAt,
    firstThree,
    firstTwo,
    firstOne,
    lastThree,
    lastTwo,
    lastOne,
  ]).filter((candidate) => candidate && candidate.length > 2 && !['THE', 'AND'].includes(candidate.toUpperCase()));
}

async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} for ${url}`);
  }
  return response.json();
}

function readWeatherCache(cachePath) {
  if (!fs.existsSync(cachePath)) {
    return { version: weatherCacheVersion, geocodes: {}, daily: {} };
  }
  try {
    const cache = JSON.parse(fs.readFileSync(cachePath, 'utf8'));
    if (cache.version !== weatherCacheVersion) {
      return { version: weatherCacheVersion, geocodes: {}, daily: {} };
    }
    cache.geocodes = cache.geocodes || {};
    cache.daily = cache.daily || {};
    return cache;
  } catch (error) {
    return { version: weatherCacheVersion, geocodes: {}, daily: {} };
  }
}

function writeWeatherCache(cachePath, cache) {
  ensureDir(path.dirname(cachePath));
  fs.writeFileSync(cachePath, JSON.stringify({
    ...cache,
    version: weatherCacheVersion,
    updated_at: new Date().toISOString(),
  }), 'utf8');
}

async function geocodeStoreLocation(location, cache) {
  if (cache.geocodes[location.unit]) return cache.geocodes[location.unit];
  if (!location.area) {
    cache.geocodes[location.unit] = null;
    return null;
  }

  const queries = locationCandidates(location);

  for (const query of queries) {
    if (!query) continue;
    const url = `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(query)}&count=10&language=en&format=json`;
    try {
      const json = await fetchJson(url);
      const results = (json.results || []).filter((result) => result.country_code === 'US');
      const result = results.find((candidate) => {
        if (!location.state_name) return true;
        return String(candidate.admin1 || '').toLowerCase() === location.state_name.toLowerCase();
      }) || results[0];
      if (result && Number.isFinite(Number(result.latitude)) && Number.isFinite(Number(result.longitude))) {
        const geocode = {
          query,
          latitude: Number(result.latitude),
          longitude: Number(result.longitude),
          name: result.name || '',
          admin1: result.admin1 || '',
          country_code: result.country_code || '',
        };
        cache.geocodes[location.unit] = geocode;
        return geocode;
      }
    } catch (error) {
      // Try the next candidate query before giving up on this store.
    }
    await sleep(120);
  }

  cache.geocodes[location.unit] = null;
  return null;
}

function appendDailyWeather(cache, unit, daily) {
  cache.daily[unit] = cache.daily[unit] || {};
  const times = daily.time || [];
  const precipitation = daily.precipitation_sum || [];
  for (let index = 0; index < times.length; index += 1) {
    const value = precipitation[index];
    cache.daily[unit][times[index]] = Number.isFinite(Number(value)) ? round4(Number(value)) : 0;
  }
}

async function fetchArchiveWeather(locations, cache, startDate, endDate) {
  const batchSize = 25;
  for (let offset = 0; offset < locations.length; offset += batchSize) {
    const batch = locations.slice(offset, offset + batchSize);
    const latitude = batch.map((item) => item.geocode.latitude).join(',');
    const longitude = batch.map((item) => item.geocode.longitude).join(',');
    const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${latitude}&longitude=${longitude}&start_date=${startDate}&end_date=${endDate}&daily=precipitation_sum&precipitation_unit=inch&timezone=auto`;
    let json;
    try {
      json = await fetchJson(url);
    } catch (error) {
      console.log(`Weather archive batch failed (${offset + 1}-${offset + batch.length}); retrying individually...`);
      for (const item of batch) {
        const singleUrl = `https://archive-api.open-meteo.com/v1/archive?latitude=${item.geocode.latitude}&longitude=${item.geocode.longitude}&start_date=${startDate}&end_date=${endDate}&daily=precipitation_sum&precipitation_unit=inch&timezone=auto`;
        try {
          const singleJson = await fetchJson(singleUrl);
          appendDailyWeather(cache, item.unit, singleJson.daily || {});
        } catch (singleError) {
          item.weather_error = singleError.message;
        }
        await sleep(120);
      }
      continue;
    }

    const results = Array.isArray(json) ? json : [json];
    results.forEach((result, index) => {
      const item = batch[index];
      if (item && result && result.daily) appendDailyWeather(cache, item.unit, result.daily);
    });
    console.log(`Fetched weather archive for ${Math.min(offset + batch.length, locations.length)} of ${locations.length} locations...`);
    await sleep(250);
  }
}

function hasWeatherRange(cache, unit, startDate, endDate) {
  const daily = cache.daily[unit];
  return Boolean(daily && Object.prototype.hasOwnProperty.call(daily, startDate) && Object.prototype.hasOwnProperty.call(daily, endDate));
}

async function loadWeatherData(units, cacheDir) {
  if (process.env.SKIP_WEATHER === '1') {
    return { enabled: false, cache: { geocodes: {}, daily: {} }, locations: [], source: 'disabled' };
  }
  if (typeof fetch !== 'function') {
    return { enabled: false, cache: { geocodes: {}, daily: {} }, locations: [], source: 'fetch unavailable' };
  }

  const cachePath = path.join(cacheDir, 'open_meteo_weather.json');
  const cache = readWeatherCache(cachePath);
  const startDate = `${historyYears[0]}-01-01`;
  const endDate = `${weatherAnalogYear}-12-31`;
  const locations = units.map(parseStoreLocation);
  const geocodedLocations = [];

  console.log(`Resolving weather locations for ${locations.length} stores...`);
  for (let index = 0; index < locations.length; index += 1) {
    const location = locations[index];
    const geocode = await geocodeStoreLocation(location, cache);
    if (geocode) {
      geocodedLocations.push({ ...location, geocode });
    }
    if ((index + 1) % 25 === 0) {
      console.log(`Resolved ${index + 1} of ${locations.length} weather locations...`);
      writeWeatherCache(cachePath, cache);
    }
  }

  const needsWeather = geocodedLocations.filter((location) => !hasWeatherRange(cache, location.unit, startDate, endDate));
  if (needsWeather.length) {
    console.log(`Fetching historical precipitation for ${needsWeather.length} stores from Open-Meteo...`);
    await fetchArchiveWeather(needsWeather, cache, startDate, endDate);
  } else {
    console.log('Using cached historical precipitation for all resolved stores...');
  }
  writeWeatherCache(cachePath, cache);

  return {
    enabled: true,
    cache,
    locations: geocodedLocations,
    source: 'Open-Meteo archive precipitation_sum',
    cache_path: cachePath,
    unresolved_locations: locations.length - geocodedLocations.length,
  };
}

function weeklyRainFromWeather(weatherData, unit, year, month, week, options = {}) {
  const daily = weatherData.cache.daily[unit];
  if (!weatherData.enabled || !daily) {
    return { precipitation_inches: 0, rainy: false, source: 'weather unavailable' };
  }

  const sourceYear = options.analog ? weatherAnalogYear : year;
  const dates = weekDateKeys(sourceYear, month, week);
  let missing = false;
  const total = dates.reduce((sum, key) => {
    if (!Object.prototype.hasOwnProperty.call(daily, key)) {
      missing = true;
      return sum;
    }
    return sum + Number(daily[key] || 0);
  }, 0);

  if (missing) {
    return { precipitation_inches: 0, rainy: false, source: 'weather unavailable' };
  }

  const precipitation = round4(total);
  return {
    precipitation_inches: precipitation,
    rainy: precipitation >= weatherRainThresholdInches,
    source: options.analog ? `${weatherAnalogYear} analog` : 'historical',
  };
}

function fallbackWeatherDrop(baseCases, precipitationInches) {
  if (baseCases <= 0) return 0;
  const rainSeverity = Math.max(0, precipitationInches - weatherRainThresholdInches);
  const rate = Math.min(0.25, weatherFallbackDropRate + rainSeverity * 0.03);
  return baseCases * rate;
}

function calculateWeatherAdjustment(baseCases, activeValues, historicalWeather, targetWeather) {
  if (!targetWeather || targetWeather.source === 'weather unavailable' || baseCases <= 0) {
    return { adjustment_cases: 0, adjusted_cases: baseCases, method: 'no weather adjustment' };
  }

  const targetRainy = Boolean(targetWeather.rainy);
  const matchingValues = activeValues.filter((_, index) => historicalWeather[index] && historicalWeather[index].rainy === targetRainy);
  let adjustment = 0;
  let method = targetRainy ? 'rain fallback drop' : 'dry no adjustment';

  if (matchingValues.length >= 2) {
    const weatherAverage = average(matchingValues);
    adjustment = weatherAverage - baseCases;
    method = targetRainy ? 'historical rainy average' : 'historical dry average';

    if (targetRainy && adjustment >= 0) {
      adjustment = -fallbackWeatherDrop(baseCases, targetWeather.precipitation_inches);
      method = 'rain fallback drop';
    }
    if (!targetRainy && adjustment < 0) {
      adjustment = 0;
      method = 'dry no adjustment';
    }
  } else if (targetRainy) {
    adjustment = -fallbackWeatherDrop(baseCases, targetWeather.precipitation_inches);
  }

  const roundedAdjustment = round4(adjustment);
  return {
    adjustment_cases: roundedAdjustment,
    adjusted_cases: round4(Math.max(0, baseCases + roundedAdjustment)),
    method,
  };
}

async function readWorkbookRowsExcelJs(filePath) {
  let headerRowNumber = -1;
  let headerMap = null;
  const output = [];
  let seenWorksheet = false;

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'cache',
  });

  for await (const worksheetReader of workbookReader) {
    if (worksheetReader.name !== 'EXPORT') {
      continue;
    }
    seenWorksheet = true;

    for await (const row of worksheetReader) {
      const rowNumber = row.number;
      const values = row.values || [];

      if (headerRowNumber < 0) {
        const candidate = {};
        for (let col = 1; col < values.length; col += 1) {
          const cell = row.getCell(col);
          const normalized = normalizeHeader(cell.text);
          if (requiredHeaders.includes(normalized)) candidate[normalized] = col;
        }
        if (requiredHeaders.every((name) => Number.isInteger(candidate[name]))) {
          headerRowNumber = rowNumber;
          headerMap = candidate;
        }
        continue;
      }

      const unit = String(row.getCell(headerMap['UNIT']).text || '').trim();
      const flavor = String(row.getCell(headerMap['NDA ITEM DESCRIPTION']).text || '').trim();
      const invoiceCell = row.getCell(headerMap['INVOICE DATE']);
      const orderedCell = row.getCell(headerMap['ORDERED']);
      const invoiceRaw = invoiceCell.value;
      const orderedRaw = orderedCell.value;

      if (!unit && !flavor && invoiceRaw == null && orderedRaw == null) continue;

      let date = null;
      if (invoiceRaw instanceof Date) date = invoiceRaw;
      else if (invoiceRaw && typeof invoiceRaw === 'object' && invoiceRaw.result instanceof Date) date = invoiceRaw.result;
      else date = parseDate(invoiceCell.text || invoiceRaw);
      if (!date) continue;

      let ordered = null;
      if (typeof orderedRaw === 'number') ordered = orderedRaw;
      else if (orderedRaw && typeof orderedRaw === 'object' && typeof orderedRaw.result === 'number') ordered = orderedRaw.result;
      else ordered = Number(String(orderedCell.text || orderedRaw || '').replace(/,/g, ''));
      if (!Number.isFinite(ordered)) continue;

      output.push({
        UNIT: unit,
        STORE_CODE: getStoreCode(unit),
        'NDA ITEM DESCRIPTION': flavor,
        'INVOICE DATE': monthKey(date.getFullYear(), date.getMonth() + 1) + `-${String(date.getDate()).padStart(2, '0')}`,
        ORDERED: Number(ordered.toFixed(4)),
        YEAR: date.getFullYear(),
        MONTH: date.getMonth() + 1,
        WEEK_OF_MONTH: weekOfMonth(date),
      });
    }
  }

  if (!seenWorksheet) throw new Error(`EXPORT worksheet not found in ${path.basename(filePath)}`);
  if (headerRowNumber < 0) throw new Error(`Could not find required headers in ${path.basename(filePath)}`);
  return output;
}

async function readWorkbookRows(filePath) {
  try {
    return readWorkbookRowsXlsx(filePath);
  } catch (error) {
    console.log(`Falling back to exceljs for ${path.basename(filePath)}...`);
    return readWorkbookRowsExcelJs(filePath);
  }
}

function readWorkbookRowsXlsx(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true, dense: true });
  const sheetName = workbook.SheetNames.includes('EXPORT') ? 'EXPORT' : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

  let headerIndex = -1;
  let headerMap = null;

  for (let i = 0; i < rows.length; i += 1) {
    const candidate = {};
    rows[i].forEach((cell, idx) => {
      const normalized = normalizeHeader(cell);
      if (requiredHeaders.includes(normalized)) candidate[normalized] = idx;
    });
    if (requiredHeaders.every((name) => Number.isInteger(candidate[name]))) {
      headerIndex = i;
      headerMap = candidate;
      break;
    }
  }

  if (headerIndex < 0) throw new Error(`Could not find required headers in ${path.basename(filePath)}`);

  const output = [];
  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i];
    const unit = String(row[headerMap['UNIT']] ?? '').trim();
    const flavor = String(row[headerMap['NDA ITEM DESCRIPTION']] ?? '').trim();
    const invoiceRaw = row[headerMap['INVOICE DATE']];
    const orderedRaw = row[headerMap['ORDERED']];

    if (!unit && !flavor && invoiceRaw == null && orderedRaw == null) continue;

    const date = parseDate(invoiceRaw);
    if (!date) continue;

    const ordered = Number(String(orderedRaw ?? '').replace(/,/g, ''));
    if (!Number.isFinite(ordered)) continue;

    output.push({
      UNIT: unit,
      STORE_CODE: getStoreCode(unit),
      'NDA ITEM DESCRIPTION': flavor,
      'INVOICE DATE': monthKey(date.getFullYear(), date.getMonth() + 1) + `-${String(date.getDate()).padStart(2, '0')}`,
      ORDERED: Number(ordered.toFixed(4)),
      YEAR: date.getFullYear(),
      MONTH: date.getMonth() + 1,
      WEEK_OF_MONTH: weekOfMonth(date),
    });
  }

  return output;
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function buildHtmlReport(outPath) {
  const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Weekly Average Forecast ${forecastYear}</title>
<style>
:root{--ink:#17212b;--muted:#5c6b73;--line:#d6e2e7;--panel:#f7fbfc;--accent:#0f766e;--forecast:#1d4ed8;--bg:#fbf8f1}
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--ink);background:linear-gradient(135deg,#fbf8f1 0%,#eef8f6 52%,#f8fafc 100%)}
header{padding:24px 30px 20px;border-bottom:1px solid var(--line);background:rgba(255,255,255,.72)}
.brand{display:flex;align-items:center;gap:12px}
.brand-mark{display:inline-flex;align-items:center;gap:7px;font-weight:800;color:#64a70b;font-size:28px;line-height:1}
.brand-mark::before{content:"";width:24px;height:24px;border-radius:5px 5px 13px 13px;background:#64a70b}
.brand-divider{width:1px;height:30px;background:var(--line)}
.brand-title{font-size:28px;font-weight:800;color:#17212b;line-height:1}
main{padding:22px 30px 38px}
.filters,.grid{display:grid;gap:12px;margin-bottom:18px}
.filters{grid-template-columns:repeat(6,minmax(150px,1fr))}
.grid{grid-template-columns:repeat(9,minmax(120px,1fr))}
label{display:block;font-size:11px;color:var(--muted);text-transform:uppercase;margin-bottom:4px;letter-spacing:.04em}
select{width:100%;padding:10px 11px;border:1px solid var(--line);border-radius:7px;background:#fff;font-size:14px}
.card{border:1px solid var(--line);border-radius:10px;background:rgba(255,255,255,.82);padding:12px}
.label{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.04em}
.value{margin-top:6px;font-size:22px;font-weight:700}
.forecast-card{border-color:#b8c9f3;background:#f5f8ff}
.weather-card{border-color:#c9d8ba;background:#f7fbf3}
.selected{border:1px solid var(--line);border-radius:10px;background:rgba(255,255,255,.82);padding:12px 14px;margin-bottom:18px}
.selected strong{display:block;margin-bottom:3px}
h2{font-size:18px;margin:22px 0 10px}
.table-wrap{border:1px solid var(--line);border-radius:10px;overflow:auto;background:#fff}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border-bottom:1px solid var(--line);padding:9px 10px;text-align:right;white-space:nowrap}
th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}
th{background:#edf6f7;font-size:12px}
td.forecast,th.forecast{color:var(--forecast);font-weight:700;background:#f5f8ff}
.empty{display:none;border:1px solid var(--line);border-radius:10px;background:#fff;padding:16px;color:var(--muted)}
@media (max-width:1100px){.grid{grid-template-columns:repeat(2,minmax(140px,1fr))}}
@media (max-width:900px){.filters,.grid{grid-template-columns:1fr}header,main{padding-left:16px;padding-right:16px}}
</style>
</head>
<body>
<header>
<div class="brand" aria-label="Yogurtland Forecasting">
<span class="brand-mark">Yogurtland</span>
<span class="brand-divider" aria-hidden="true"></span>
<span class="brand-title">Forecasting</span>
</div>
</header>
<main>
<div class="filters">
<div><label for="storeFilter">Store / Unit</label><select id="storeFilter"></select></div>
<div><label for="monthFilter">Month</label><select id="monthFilter"></select></div>
<div><label for="groupFilter">Item Group</label><select id="groupFilter"></select></div>
<div><label for="categoryFilter">Category</label><select id="categoryFilter"></select></div>
<div><label for="flavorTypeFilter">Flavor Type</label><select id="flavorTypeFilter"></select></div>
<div><label for="flavorFilter">Item</label><select id="flavorFilter"></select></div>
</div>
<div class="selected" id="selectedSummary"></div>
<div class="grid">
${historyYears.map((year, index) => `<div class="card"><div class="label">${year} Actual</div><div class="value" id="kpiActual${index}">0</div></div>`).join('\n')}
<div class="card forecast-card"><div class="label">${forecastYear} Forecast</div><div class="value" id="kpiForecast">0</div></div>
<div class="card weather-card"><div class="label">Weather +/-</div><div class="value" id="kpiWeatherAdjustment">0</div></div>
<div class="card weather-card"><div class="label">Weather Adjusted</div><div class="value" id="kpiWeatherAdjusted">0</div></div>
</div>
<div class="empty" id="emptyState">No historical rows are available for this store, month, and item.</div>
<h2>Week-on-Week Forecast</h2>
<div class="table-wrap"><table><thead><tr><th>Week</th><th>Days</th>${historyYears.map((year) => `<th>${year} Actual</th>`).join('')}<th class="forecast">${forecastYear} Forecast Avg</th><th>Weather Rain</th><th>Weather +/-</th><th>Weather Adjusted</th></tr></thead><tbody id="weeklyRows"></tbody></table></div>
</main>
<script src="./two_step_forecast_2026_rows.js"></script>
<script>
const weekDays = ['1-7', '8-14', '15-21', '22-28', '29-end'];
const allValue = '__ALL__';
const stores = Array.from(new Set(rows.map(function(r) { return r[0]; }))).sort();
function fmt(n) { return Number(n || 0).toFixed(2); }
function fmtSigned(n) {
  const value = Number(n || 0);
  return (value > 0 ? '+' : '') + fmt(value);
}
function fmtRain(n) {
  return fmt(n) + ' in';
}
function esc(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(ch) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch];
  });
}
function formatMonth(monthKey) {
  const parts = String(monthKey).split('-');
  const date = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  return date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
}
function uniqueSorted(values) {
  return Array.from(new Set(values.filter(function(value) { return value != null && String(value).trim() !== ''; }))).sort();
}
function preferredSorted(values, preferred) {
  const present = new Set(uniqueSorted(values));
  const preferredValues = Array.isArray(preferred) ? preferred : [];
  const ordered = preferredValues.filter(function(value) { return present.has(value); });
  const rest = Array.from(present).filter(function(value) { return !preferredValues.includes(value); }).sort();
  return ordered.concat(rest);
}
function addOption(select, value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}
function setOptions(id, values, formatter, allLabel) {
  const select = document.getElementById(id);
  const previous = select.value;
  select.innerHTML = '';
  if (allLabel) addOption(select, allValue, allLabel);
  values.forEach(function(value) {
    addOption(select, value, formatter ? formatter(value) : value);
  });
  if (Array.from(select.options).some(function(option) { return option.value === previous; })) {
    select.value = previous;
  }
}
function rowsForStore() {
  const store = document.getElementById('storeFilter').value;
  return rows.filter(function(r) { return r[0] === store; });
}
function rowsForStoreMonth() {
  const store = document.getElementById('storeFilter').value;
  const month = document.getElementById('monthFilter').value;
  return rows.filter(function(r) { return r[0] === store && r[2] === month; });
}
function rowsForStoreMonthGroup() {
  const group = document.getElementById('groupFilter').value;
  return rowsForStoreMonth().filter(function(r) { return r[8] === group; });
}
function rowsForCategory() {
  const category = document.getElementById('categoryFilter').value;
  return rowsForStoreMonthGroup().filter(function(r) {
    return category === allValue || r[9] === category;
  });
}
function rowsForSelectableItems() {
  const flavorType = document.getElementById('flavorTypeFilter').value;
  return rowsForCategory().filter(function(r) {
    return flavorType === allValue || r[10] === flavorType;
  });
}
function refreshMonthOptions() {
  setOptions('monthFilter', uniqueSorted(rowsForStore().map(function(r) { return r[2]; })), formatMonth);
}
function refreshGroupOptions() {
  setOptions('groupFilter', uniqueSorted(rowsForStoreMonth().map(function(r) { return r[8]; })));
}
function refreshCategoryOptions() {
  const preferred = typeof categoryOptions === 'undefined' ? [] : categoryOptions;
  setOptions('categoryFilter', preferredSorted(rowsForStoreMonthGroup().map(function(r) { return r[9]; }), preferred), null, 'All Categories');
}
function refreshFlavorTypeOptions() {
  const preferred = typeof flavorTypeOptions === 'undefined' ? [] : flavorTypeOptions;
  setOptions('flavorTypeFilter', preferredSorted(rowsForCategory().map(function(r) { return r[10]; }), preferred), null, 'All Types');
}
function refreshFlavorOptions() {
  setOptions('flavorFilter', uniqueSorted(rowsForSelectableItems().map(function(r) { return r[3]; })));
}
function selectedRow() {
  const item = document.getElementById('flavorFilter').value;
  return rowsForSelectableItems().find(function(r) { return r[3] === item; });
}
function setEmpty(isEmpty) {
  document.getElementById('emptyState').style.display = isEmpty ? 'block' : 'none';
}
function onStoreChange() {
  refreshMonthOptions();
  refreshGroupOptions();
  refreshCategoryOptions();
  refreshFlavorTypeOptions();
  refreshFlavorOptions();
  render();
}
function onMonthChange() {
  refreshGroupOptions();
  refreshCategoryOptions();
  refreshFlavorTypeOptions();
  refreshFlavorOptions();
  render();
}
function onGroupChange() {
  refreshCategoryOptions();
  refreshFlavorTypeOptions();
  refreshFlavorOptions();
  render();
}
function onCategoryChange() {
  refreshFlavorTypeOptions();
  refreshFlavorOptions();
  render();
}
function onFlavorTypeChange() {
  refreshFlavorOptions();
  render();
}
function render() {
  const row = selectedRow();
  if (!row) {
    setEmpty(true);
    document.getElementById('selectedSummary').innerHTML = '<strong>No selection available</strong>';
    historyYears.forEach(function(_, index) { document.getElementById('kpiActual' + index).textContent = '0.00'; });
    document.getElementById('kpiForecast').textContent = '0.00';
    document.getElementById('kpiWeatherAdjustment').textContent = '0.00';
    document.getElementById('kpiWeatherAdjusted').textContent = '0.00';
    document.getElementById('weeklyRows').innerHTML = '';
    return;
  }
  setEmpty(false);
  const details = [formatMonth(row[2]), row[8], row[9], row[10], row[11]]
    .filter(function(value) { return value && value !== 'Uncategorized'; })
    .filter(function(value, index, array) { return array.indexOf(value) === index; })
    .join(' - ');
  document.getElementById('selectedSummary').innerHTML =
    '<strong>' + esc(row[0]) + '</strong>' + esc(details) + ' - ' + esc(row[3]);
  historyYears.forEach(function(_, index) {
    document.getElementById('kpiActual' + index).textContent = fmt(row[6][index]);
  });
  document.getElementById('kpiForecast').textContent = fmt(row[4]);
  document.getElementById('kpiWeatherAdjustment').textContent = fmtSigned(row[12]);
  document.getElementById('kpiWeatherAdjusted').textContent = fmt(row[13]);
  document.getElementById('weeklyRows').innerHTML = row[7].map(function(yearValues, weekIndex) {
    const actualCells = yearValues.map(function(value) { return '<td>' + fmt(value) + '</td>'; }).join('');
    return '<tr><td>Week ' + (weekIndex + 1) + '</td><td>' + weekDays[weekIndex] + '</td>' +
      actualCells + '<td class="forecast">' + fmt(row[5][weekIndex]) + '</td>' +
      '<td>' + fmtRain(row[16][weekIndex]) + '</td>' +
      '<td>' + fmtSigned(row[14][weekIndex]) + '</td>' +
      '<td class="forecast">' + fmt(row[15][weekIndex]) + '</td></tr>';
  }).join('');
}
setOptions('storeFilter', stores);
refreshMonthOptions();
refreshGroupOptions();
refreshCategoryOptions();
refreshFlavorTypeOptions();
refreshFlavorOptions();
document.getElementById('storeFilter').addEventListener('change', onStoreChange);
document.getElementById('monthFilter').addEventListener('change', onMonthChange);
document.getElementById('groupFilter').addEventListener('change', onGroupChange);
document.getElementById('categoryFilter').addEventListener('change', onCategoryChange);
document.getElementById('flavorTypeFilter').addEventListener('change', onFlavorTypeChange);
document.getElementById('flavorFilter').addEventListener('change', render);
render();
</script>
</body>
</html>`;

  fs.writeFileSync(outPath, html, 'utf8');
}

async function main() {
  const files = fs.readdirSync(cwd)
    .filter((name) => name.toLowerCase().endsWith('.xlsx') && !name.startsWith('~$'))
    .sort();
  if (!files.length) throw new Error('No .xlsx files found.');
  const lookbackYears = historyYears;
  const lookbackYearSet = new Set(lookbackYears);
  const sourceFiles = files.filter((file) => {
    const match = file.match(/^(\d{4})\b/);
    return !match || lookbackYearSet.has(Number(match[1]));
  });
  if (!sourceFiles.length) throw new Error(`No .xlsx files found for ${lookbackYears.join(', ')}.`);
  const classification = readClassificationData(files);
  if (classification.source_file) {
    console.log(`Using classification from ${classification.source_file}...`);
  } else {
    console.log('No Categorization sheet found; using item-name rules for flavors and toppings...');
  }

  const cacheDir = path.join(cwd, '.cache_extracted_history_node');
  ensureDir(cacheDir);

  const historyRows = [];
  const fileSummaries = [];

  for (const file of sourceFiles) {
    const fullPath = path.join(cwd, file);
    const cacheFile = path.join(cacheDir, file.replace(/\.xlsx$/i, '.json'));
    let rows;
    if (fs.existsSync(cacheFile)) {
      rows = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
      console.log(`Using cache for ${file}...`);
    } else {
      console.log(`Reading ${file}...`);
        rows = await readWorkbookRows(fullPath);
        fs.writeFileSync(cacheFile, JSON.stringify(rows), 'utf8');
    }
    let usedRows = 0;
    for (const row of rows) {
      if (!lookbackYearSet.has(Number(row.YEAR))) continue;
      historyRows.push(row);
      usedRows += 1;
    }
    fileSummaries.push({ file, rows: rows.length, used_rows: usedRows });
  }

  const historyCsvPath = path.join(cwd, 'purchase_history_4cols.csv');
  writeCsv(historyCsvPath, historyRows, ['UNIT', 'STORE_CODE', 'NDA ITEM DESCRIPTION', 'INVOICE DATE', 'ORDERED', 'YEAR', 'MONTH', 'WEEK_OF_MONTH']);

  const monthlyMap = new Map();
  const weeklyMap = new Map();
  const pairSet = new Map();
  const classificationCounts = { Flavor: 0, Topping: 0, Other: 0 };

  for (const row of historyRows) {
    const itemDescription = row['NDA ITEM DESCRIPTION'];
    const itemInfo = classifyItem(itemDescription, classification);
    classificationCounts[itemInfo.item_group] = (classificationCounts[itemInfo.item_group] || 0) + 1;
    if (itemInfo.item_group === 'Other') continue;

    const monthlyKey = `${row.UNIT}||${itemDescription}||${row.YEAR}||${row.MONTH}`;
    monthlyMap.set(monthlyKey, (monthlyMap.get(monthlyKey) || 0) + Number(row.ORDERED));
    const weeklyKey = `${row.UNIT}||${itemDescription}||${row.YEAR}||${row.MONTH}||${row.WEEK_OF_MONTH}`;
    weeklyMap.set(weeklyKey, (weeklyMap.get(weeklyKey) || 0) + Number(row.ORDERED));
    if (lookbackYears.includes(Number(row.YEAR))) {
      const pairKey = `${row.UNIT}||${itemDescription}`;
      if (!pairSet.has(pairKey)) {
        pairSet.set(pairKey, {
          unit: row.UNIT,
          store_code: row.STORE_CODE,
          flavor: itemDescription,
          item_group: itemInfo.item_group,
          category: itemInfo.category,
          flavor_type: itemInfo.flavor_type,
          seasonal: itemInfo.seasonal,
        });
      }
    }
  }

  const weatherUnits = [...new Set([...pairSet.values()].map((pair) => pair.unit))].sort();
  const weatherData = await loadWeatherData(weatherUnits, path.join(cwd, '.cache_weather'));
  console.log(`Weather data: ${weatherData.enabled ? `${weatherData.locations.length} resolved, ${weatherData.unresolved_locations || 0} unresolved` : weatherData.source}.`);

  const forecastRows = [];
  const pairs = [...pairSet.values()].sort((a, b) => a.unit.localeCompare(b.unit) || a.flavor.localeCompare(b.flavor));
  for (const pair of pairs) {
    for (let month = 1; month <= 12; month += 1) {
      const yearlyTotals = lookbackYears.map((year) => monthlyMap.get(`${pair.unit}||${pair.flavor}||${year}||${month}`) || 0);
      const activeYearIndexes = yearlyTotals
        .map((value, index) => (value > 0 ? index : -1))
        .filter((index) => index >= 0);
      if (!activeYearIndexes.length) continue;

      const weeklyActuals = [];
      const weekValues = [];
      const weeklyWeatherAdjustments = [];
      const weeklyWeatherAdjusted = [];
      const weeklyWeatherRain = [];
      const weeklyWeatherSources = [];
      const weeklyWeatherMethods = [];
      for (let week = 1; week <= 5; week += 1) {
        const yearWeekValues = lookbackYears.map((year) => weeklyMap.get(`${pair.unit}||${pair.flavor}||${year}||${month}||${week}`) || 0);
        const activeValues = activeYearIndexes.map((yearIndex) => yearWeekValues[yearIndex]);
        const weekAverage = average(activeValues);
        const historicalWeather = activeYearIndexes.map((yearIndex) => weeklyRainFromWeather(weatherData, pair.unit, lookbackYears[yearIndex], month, week));
        const targetWeather = weeklyRainFromWeather(weatherData, pair.unit, forecastYear, month, week, { analog: true });
        const weatherResult = calculateWeatherAdjustment(weekAverage, activeValues, historicalWeather, targetWeather);

        weeklyActuals.push(yearWeekValues.map((value) => Number(value.toFixed(4))));
        weekValues.push(Number(weekAverage.toFixed(4)));
        weeklyWeatherAdjustments.push(weatherResult.adjustment_cases);
        weeklyWeatherAdjusted.push(weatherResult.adjusted_cases);
        weeklyWeatherRain.push(targetWeather.precipitation_inches);
        weeklyWeatherSources.push(targetWeather.source);
        weeklyWeatherMethods.push(weatherResult.method);
      }
      const monthlyForecast = Number(weekValues.reduce((a, b) => a + b, 0).toFixed(4));
      const monthlyWeatherAdjustment = round4(weeklyWeatherAdjustments.reduce((a, b) => a + b, 0));
      const monthlyWeatherAdjusted = round4(Math.max(0, monthlyForecast + monthlyWeatherAdjustment));
      const monthlyWeatherRain = round4(weeklyWeatherRain.reduce((a, b) => a + b, 0));
      const weatherSource = uniqueValues(weeklyWeatherSources).join(', ') || 'weather unavailable';
      const weatherMethod = uniqueValues(weeklyWeatherMethods).join(', ') || 'no weather adjustment';
      const forecastMonthKey = monthKey(forecastYear, month);
      const outputRow = {
        unit: pair.unit,
        store_code: pair.store_code,
        month_key: forecastMonthKey,
        flavor: pair.flavor,
        item_group: pair.item_group,
        category: pair.category,
        flavor_type: pair.flavor_type,
        seasonal: pair.seasonal,
        monthly_forecast_cases: monthlyForecast,
        weather_adjustment_cases: monthlyWeatherAdjustment,
        weather_adjusted_cases: monthlyWeatherAdjusted,
        weather_rain_inches: monthlyWeatherRain,
        weather_source: weatherSource,
        weather_method: weatherMethod,
        history_years_used: activeYearIndexes.length,
        week_1_cases: weekValues[0],
        week_2_cases: weekValues[1],
        week_3_cases: weekValues[2],
        week_4_cases: weekValues[3],
        week_5_cases: weekValues[4],
        week_1_weather_adjustment_cases: weeklyWeatherAdjustments[0],
        week_2_weather_adjustment_cases: weeklyWeatherAdjustments[1],
        week_3_weather_adjustment_cases: weeklyWeatherAdjustments[2],
        week_4_weather_adjustment_cases: weeklyWeatherAdjustments[3],
        week_5_weather_adjustment_cases: weeklyWeatherAdjustments[4],
        week_1_weather_adjusted_cases: weeklyWeatherAdjusted[0],
        week_2_weather_adjusted_cases: weeklyWeatherAdjusted[1],
        week_3_weather_adjusted_cases: weeklyWeatherAdjusted[2],
        week_4_weather_adjusted_cases: weeklyWeatherAdjusted[3],
        week_5_weather_adjusted_cases: weeklyWeatherAdjusted[4],
        week_1_weather_rain_inches: weeklyWeatherRain[0],
        week_2_weather_rain_inches: weeklyWeatherRain[1],
        week_3_weather_rain_inches: weeklyWeatherRain[2],
        week_4_weather_rain_inches: weeklyWeatherRain[3],
        week_5_weather_rain_inches: weeklyWeatherRain[4],
        _monthly_actuals: yearlyTotals.map((value) => Number(value.toFixed(4))),
        _weekly_actuals: weeklyActuals,
        _weekly_weather_adjustments: weeklyWeatherAdjustments,
        _weekly_weather_adjusted: weeklyWeatherAdjusted,
        _weekly_weather_rain: weeklyWeatherRain,
      };
      lookbackYears.forEach((year, yearIndex) => {
        outputRow[`actual_${year}_month_cases`] = outputRow._monthly_actuals[yearIndex];
        for (let week = 1; week <= 5; week += 1) {
          outputRow[`actual_${year}_week_${week}_cases`] = outputRow._weekly_actuals[week - 1][yearIndex];
        }
      });
      forecastRows.push(outputRow);
    }
  }

  const forecastCsvPath = path.join(cwd, 'two_step_forecast_2026.csv');
  const forecastCsvHeaders = [
    'unit',
    'store_code',
    'month_key',
    'flavor',
    'item_group',
    'category',
    'flavor_type',
    'seasonal',
    'monthly_forecast_cases',
    'weather_adjustment_cases',
    'weather_adjusted_cases',
    'weather_rain_inches',
    'weather_source',
    'weather_method',
    'history_years_used',
    'week_1_cases',
    'week_2_cases',
    'week_3_cases',
    'week_4_cases',
    'week_5_cases',
    'week_1_weather_adjustment_cases',
    'week_2_weather_adjustment_cases',
    'week_3_weather_adjustment_cases',
    'week_4_weather_adjustment_cases',
    'week_5_weather_adjustment_cases',
    'week_1_weather_adjusted_cases',
    'week_2_weather_adjusted_cases',
    'week_3_weather_adjusted_cases',
    'week_4_weather_adjusted_cases',
    'week_5_weather_adjusted_cases',
    'week_1_weather_rain_inches',
    'week_2_weather_rain_inches',
    'week_3_weather_rain_inches',
    'week_4_weather_rain_inches',
    'week_5_weather_rain_inches',
    ...lookbackYears.map((year) => `actual_${year}_month_cases`),
    ...lookbackYears.flatMap((year) => [1, 2, 3, 4, 5].map((week) => `actual_${year}_week_${week}_cases`)),
  ];
  writeCsv(forecastCsvPath, forecastRows, forecastCsvHeaders);
  fs.writeFileSync(
    path.join(cwd, 'two_step_forecast_2026_rows.js'),
    `const historyYears = ${JSON.stringify(lookbackYears)};\nconst categoryOptions = ${JSON.stringify(classification.categories)};\nconst flavorTypeOptions = ${JSON.stringify(classification.flavor_types)};\nconst rows = ${JSON.stringify(forecastRows.map((r) => [
      r.unit,
      r.store_code,
      r.month_key,
      r.flavor,
      r.monthly_forecast_cases,
      [r.week_1_cases, r.week_2_cases, r.week_3_cases, r.week_4_cases, r.week_5_cases],
      r._monthly_actuals,
      r._weekly_actuals,
      r.item_group,
      r.category,
      r.flavor_type,
      r.seasonal,
      r.weather_adjustment_cases,
      r.weather_adjusted_cases,
      r._weekly_weather_adjustments,
      r._weekly_weather_adjusted,
      r._weekly_weather_rain,
      r.weather_source,
      r.weather_method,
    ]))};`,
    'utf8',
  );
  const reportHtmlPath = path.join(cwd, 'two_step_forecast_2026.html');
  buildHtmlReport(reportHtmlPath);

  const summary = {
    generated_at: new Date().toISOString(),
    forecast_year: forecastYear,
    files: fileSummaries,
    history_rows: historyRows.length,
    forecast_rows: forecastRows.length,
    classification_source: classification.source_file,
    classification_counts: classificationCounts,
    weather: {
      enabled: weatherData.enabled,
      source: weatherData.source,
      cache_path: weatherData.cache_path || '',
      resolved_locations: weatherData.locations ? weatherData.locations.length : 0,
      unresolved_locations: weatherData.unresolved_locations || 0,
      rain_threshold_inches: weatherRainThresholdInches,
      analog_year_for_2026_weather: weatherAnalogYear,
      fallback_rain_drop_rate: weatherFallbackDropRate,
    },
    cache_directory: cacheDir,
    history_csv: historyCsvPath,
    forecast_csv: forecastCsvPath,
    report_html: reportHtmlPath,
    method: {
      source_years: lookbackYears,
      weekly_forecast: `Each ${forecastYear} week is averaged from the same store, item, month, and week in ${lookbackYears.join(', ')}, excluding years where that store/item/month had zero cases.`,
      monthly_forecast: 'Monthly forecast is the sum of the five weekly forecast averages after excluding zero monthly years from the denominator.',
      weather_adjustment: `Weather adjustment uses location precipitation from Open-Meteo archive data. For ${forecastYear} weather, it uses ${weatherAnalogYear} same-month/same-week precipitation as the future-weather analog, then compares rainy/dry historical weeks for that store and item; rainy weeks fall back to a ${Math.round(weatherFallbackDropRate * 100)}% drop when history is sparse.`,
      week_definition: 'Week 1 = days 1-7, Week 2 = days 8-14, Week 3 = days 15-21, Week 4 = days 22-28, Week 5 = days 29-end.',
    },
  };
  fs.writeFileSync(path.join(cwd, 'two_step_forecast_2026_summary.json'), JSON.stringify(summary, null, 2), 'utf8');

  console.log(`Wrote history: ${historyCsvPath}`);
  console.log(`Wrote forecast: ${forecastCsvPath}`);
  console.log(`Wrote report: ${reportHtmlPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
