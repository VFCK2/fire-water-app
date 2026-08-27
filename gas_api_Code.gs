/**
 * 🚒 消防団 現場支援マップ - GAS API 完全版 v17 Google Maps版
 *
 * GitHub Pages から JSONP で呼び出す API。
 * Supabase は使用しません。
 *
 * 使い方
 * 1. このコードを Apps Script の Code.gs に全置換
 * 2. SPREADSHEET_ID を確認
 * 3. setupAll() を1回実行
 * 4. ウェブアプリとしてデプロイ
 *    - 次のユーザーとして実行：自分
 *    - アクセスできるユーザー：必要な範囲に設定
 * 5. GitHub の index.html をこの完全版に置換
 *
 * 改善点
 * - 水利1件ごとに Jurisdiction シートを読み直す処理を廃止
 * - 現場データ取得を1回の snapshot API に集約
 * - 水利データは CacheService を優先
 * - 水利状態保存時に不要な全画面再読込をしない
 * - JSONP のタイムアウト/エラー処理を統一
 */

const SPREADSHEET_ID = '1B7Au6B65DkvYhTSBVGyRNsajCfbKXY5ByRTXzN_vnOc';

const SEARCH_RADIUS_M = 200;
const SOURCE_DATE = '令和7年4月1日';
const IMPORT_SHEETS = ['IMPORT_消火栓','IMPORT_防火水槽','消火栓','防火水槽等'];

const TARGET_AREA = {
  south: 35.5450,
  north: 35.5665,
  west: 139.7150,
  east: 139.7555,
  bufferM: 200
};

const WATER_CACHE_KEY = 'TARGET_WATER_V14';
const WATER_CACHE_TTL = 21600; // 6時間
const JUR_CACHE_KEY = 'JURISDICTION_V17';
const JUR_CACHE_TTL = 21600;

const SHEETS = {
  WATER_LOG: 'Water_Update_Log',
  MEMBERS: 'Members',
  INCIDENT: 'Incident',
  INCIDENT_WATER: 'IncidentWater',
  MARKERS: 'Markers',
  JURISDICTION: 'Jurisdiction'
};

function ss_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'health';
  const callback = p.callback;

  try {
    const out = dispatch_(action, p);
    return respond_(callback, {ok:true, data:out});
  } catch (err) {
    return respond_(callback, {
      ok:false,
      error:String(err && err.message || err)
    });
  }
}

function respond_(callback, obj) {
  const json = JSON.stringify(obj).replace(/</g, '\\u003c');
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatch_(action, p) {
  switch (action) {
    case 'health':
      return {ok:true, time:new Date().toISOString()};

    case 'setupAll':
      return setupAll();

    case 'getWaterDataInfo':
      return getWaterDataInfo();

    case 'getInitialData':
      return getInitialData();

    case 'getIncidentSnapshot':
      return getIncidentSnapshot_(
        p.lat,
        p.lng,
        p.radiusM
      );

    case 'getNearbyWater':
      return getNearbyWater_(p.lat, p.lng, p.radiusM);

    case 'refreshWaterCache':
      return refreshWaterCache();

    case 'geocodeAddress':
      return geocodeAddress_(p.address);

    case 'computeRoute':
      return computeRoute_(
        p.originLat,
        p.originLng,
        p.destLat,
        p.destLng,
        p.mode || 'WALK'
      );

    case 'startIncident':
      return startIncident_(
        p.address,
        p.lat,
        p.lng,
        p.radiusM
      );

    case 'saveWaterStatus':
      return saveWaterStatus_(
        p.waterId,
        p.status
      );

    case 'addMarker':
      return addMarker_(
        p.type,
        p.name,
        p.lat,
        p.lng,
        p.status
      );

    case 'deleteMarker':
      return deleteMarker_(p.markerId);

    case 'endIncident':
      return endIncident_();

    default:
      throw new Error('Unknown action: ' + action);
  }
}

/* ---------- 初期設定 ---------- */

function setupAll() {
  const ss = ss_();

  const defs = {
    Water_Update_Log: ['更新日時','件数','基準日','結果'],
    Members: ['団員名'],
    Incident: ['incidentId','address','lat','lng','radiusM','createdAt','active'],
    IncidentWater: ['incidentId','waterId','status','updatedAt'],
    Markers: ['markerId','incidentId','type','name','lat','lng','status','geometry','createdAt'],
    Jurisdiction: ['名称','種別','南緯度','西経度','北緯度','東経度','周辺拡張m','備考']
  };

  Object.keys(defs).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,defs[name].length).setValues([defs[name]]);
    }
    sh.setFrozenRows(1);
  });

  const jur = ss.getSheetByName(SHEETS.JURISDICTION);

  if (jur.getLastRow() < 2) {
    const rows = [
      ['南蒲田3丁目','管轄内','','','','',200,''],
      ['西糀谷1丁目','管轄内','','','','',200,''],
      ['西糀谷2丁目','管轄内','','','','',200,''],
      ['西糀谷3丁目','管轄内','','','','',200,''],
      ['西糀谷4丁目','管轄内','','','','',200,''],
      ['北糀谷','管轄内','','','','',200,'']
    ];
    jur.getRange(2,1,rows.length,8).setValues(rows);
  }

  CacheService.getScriptCache().removeAll([
    WATER_CACHE_KEY,
    WATER_CACHE_KEY + '_COUNT',
    JUR_CACHE_KEY
  ]);

  return '初期設定完了';
}

/* ---------- 水利基本データ ---------- */

function getImportSheets_() {
  return ss_().getSheets().filter(sh => {
    const n = sh.getName();
    return IMPORT_SHEETS.indexOf(n) >= 0 || /^IMPORT_/.test(n);
  });
}

function getWaterDataInfo() {
  const cached = getCachedWater_();
  const log = ss_().getSheetByName(SHEETS.WATER_LOG);

  let updatedAt = '';
  let basis = SOURCE_DATE;

  if (log && log.getLastRow() >= 2) {
    const r = log.getRange(log.getLastRow(),1,1,4).getValues()[0];
    if (r[0]) {
      updatedAt = Utilities.formatDate(
        new Date(r[0]),
        Session.getScriptTimeZone(),
        'yyyy/MM/dd HH:mm'
      );
    }
    if (r[2]) basis = String(r[2]);
  }

  return {
    count: cached.length,
    basis: basis,
    updatedAt: updatedAt,
    source: '東京都オープンデータ／東京消防庁',
    target: '南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200m',
    cacheSeconds: WATER_CACHE_TTL
  };
}

function cachePutWater_(data) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(data);

  cache.remove(WATER_CACHE_KEY + '_COUNT');

  if (json.length <= 90000) {
    cache.put(WATER_CACHE_KEY, json, WATER_CACHE_TTL);
    return;
  }

  cache.remove(WATER_CACHE_KEY);

  const chunkSize = 80000;
  const count = Math.ceil(json.length / chunkSize);

  for (let i = 0; i < count; i++) {
    cache.put(
      WATER_CACHE_KEY + '_' + i,
      json.slice(i * chunkSize, (i + 1) * chunkSize),
      WATER_CACHE_TTL
    );
  }

  cache.put(
    WATER_CACHE_KEY + '_COUNT',
    String(count),
    WATER_CACHE_TTL
  );
}

function getCachedWater_() {
  const cache = CacheService.getScriptCache();

  const hit = cache.get(WATER_CACHE_KEY);
  if (hit) {
    try {
      return JSON.parse(hit);
    } catch (e) {}
  }

  const cnt = Number(cache.get(WATER_CACHE_KEY + '_COUNT') || 0);

  if (cnt > 0) {
    let json = '';

    for (let i = 0; i < cnt; i++) {
      const part = cache.get(WATER_CACHE_KEY + '_' + i);
      if (!part) {
        json = '';
        break;
      }
      json += part;
    }

    if (json) {
      try {
        return JSON.parse(json);
      } catch (e) {}
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const again = cache.get(WATER_CACHE_KEY);

    if (again) {
      try {
        return JSON.parse(again);
      } catch (e) {}
    }

    const data = buildTargetWaterCache_();
    cachePutWater_(data);
    return data;

  } finally {
    lock.releaseLock();
  }
}

function buildTargetWaterCache_() {
  const out = [];

  getImportSheets_().forEach(sh => {
    const values = sh.getDataRange().getDisplayValues();
    if (!values.length) return;

    const meta = detectSheetMeta_(values);
    const type = String(sh.getName()).indexOf('防火') >= 0
      ? '防火水槽等'
      : '消火栓';

    const start = meta.headerRow >= 0
      ? meta.headerRow + 1
      : 0;

    for (let i = start; i < values.length; i++) {
      const row = values[i];
      const coord = extractCoordinateFromRow_(row, meta);

      if (!coord || !isInTargetArea_(coord.lat, coord.lng)) continue;

      const id = buildWaterId_(
        sh.getName(),
        i + 1,
        row,
        meta
      );

      out.push({
        id: id,
        type: type,
        lat: coord.lat,
        lng: coord.lng,
        capacity: meta.capacityCol >= 0
          ? row[meta.capacityCol]
          : '',
        address: meta.addressCol >= 0
          ? row[meta.addressCol]
          : '',
        source: '東京都オープンデータ／東京消防庁',
        sourceDate: SOURCE_DATE
      });
    }
  });

  return out;
}

function refreshWaterCache() {
  const data = buildTargetWaterCache_();
  cachePutWater_(data);

  return {
    count: data.length,
    target: '南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200m'
  };
}

/* ---------- 水利検索 ---------- */

function getNearbyWater_(lat, lng, radiusM, optStates, optJurisdiction) {
  const radius = Math.min(
    Number(radiusM || SEARCH_RADIUS_M),
    SEARCH_RADIUS_M
  );

  const lat0 = Number(lat);
  const lng0 = Number(lng);

  if (!isFinite(lat0) || !isFinite(lng0)) return [];

  const latPad = radius / 111320;
  const lngPad = radius /
    (111320 * Math.max(0.5, Math.cos(lat0 * Math.PI / 180)));

  const states = optStates || getIncidentWaterStates_();
  const jurisdiction = optJurisdiction || getJurisdiction_();
  const base = getCachedWater_();

  const out = [];

  base.forEach(w => {
    if (Math.abs(w.lat - lat0) > latPad) return;
    if (Math.abs(w.lng - lng0) > lngPad) return;

    const d = distanceM_(
      lat0,
      lng0,
      w.lat,
      w.lng
    );

    if (d > radius) return;

    out.push(Object.assign({}, w, {
      sourceDate: SOURCE_DATE,
      updatedAt: '',
      zone: classifyJurisdictionFast_(
        w.lat,
        w.lng,
        jurisdiction
      ),
      distance: d,
      status: states[w.id] || '未確認'
    }));
  });

  out.sort((a,b) => a.distance - b.distance);

  return out.slice(0, 100);
}

/* ---------- シート列判定 ---------- */

function detectSheetMeta_(values) {
  const maxScan = Math.min(values.length, 25);

  let headerRow = -1;
  let latCol = -1;
  let lngCol = -1;
  let capacityCol = -1;
  let addressCol = -1;
  let idCol = -1;

  const latTokens = ['緯度','latitude','lat','y'];
  const lngTokens = ['経度','longitude','lng','lon','x'];
  const capTokens = ['容量','貯水量','水量','capacity'];
  const addrTokens = ['住所','所在地','場所'];
  const idTokens = ['水利番号','水利id','id','番号','施設番号'];

  for (let r = 0; r < maxScan; r++) {
    const row = values[r].map(v =>
      String(v).trim().toLowerCase()
    );

    for (let c = 0; c < row.length; c++) {
      const s = row[c];

      if (
        latCol < 0 &&
        latTokens.some(t => s === t || s.indexOf(t) >= 0) &&
        s.indexOf('経') < 0
      ) latCol = c;

      if (
        lngCol < 0 &&
        lngTokens.some(t => s === t || s.indexOf(t) >= 0)
      ) lngCol = c;

      if (
        capacityCol < 0 &&
        capTokens.some(t => s.indexOf(t) >= 0)
      ) capacityCol = c;

      if (
        addressCol < 0 &&
        addrTokens.some(t => s.indexOf(t) >= 0)
      ) addressCol = c;

      if (
        idCol < 0 &&
        idTokens.some(t => s === t || s.indexOf(t) >= 0)
      ) idCol = c;
    }

    if (latCol >= 0 && lngCol >= 0) {
      headerRow = r;
      break;
    }
  }

  return {
    headerRow,
    latCol,
    lngCol,
    capacityCol,
    addressCol,
    idCol
  };
}

function extractCoordinateFromRow_(row, meta) {
  if (meta.latCol >= 0 && meta.lngCol >= 0) {
    const lat = toNumber_(row[meta.latCol]);
    const lng = toNumber_(row[meta.lngCol]);

    if (isValidCoord_(lat, lng)) {
      return {lat:lat, lng:lng};
    }
  }

  const nums = [];

  for (let c = 0; c < row.length; c++) {
    const n = toNumber_(row[c]);
    if (isFinite(n)) nums.push(n);
  }

  for (let i = 0; i < nums.length; i++) {
    for (let j = 0; j < nums.length; j++) {
      if (i === j) continue;

      if (isValidCoord_(nums[i], nums[j])) {
        return {
          lat: nums[i],
          lng: nums[j]
        };
      }
    }
  }

  return null;
}

function toNumber_(v) {
  if (v === null || v === undefined || v === '') return NaN;

  return Number(
    String(v)
      .replace(/,/g,'')
      .replace(/[^\d.\-+]/g,'')
  );
}

function isValidCoord_(lat, lng) {
  return isFinite(lat) &&
    isFinite(lng) &&
    lat >= 20 &&
    lat <= 46 &&
    lng >= 122 &&
    lng <= 154;
}

function isInTargetArea_(lat, lng) {
  return Number(lat) >= TARGET_AREA.south &&
    Number(lat) <= TARGET_AREA.north &&
    Number(lng) >= TARGET_AREA.west &&
    Number(lng) <= TARGET_AREA.east;
}

function buildWaterId_(sheetName, rowNo, row, meta) {
  const raw = meta.idCol >= 0
    ? String(row[meta.idCol] || '').trim()
    : '';

  return raw
    ? sheetName + ':' + raw
    : sheetName + ':row' + rowNo;
}

function distanceM_(lat1, lng1, lat2, lng2) {
  const p = Math.PI / 180;

  const a =
    0.5 -
    Math.cos((lat2 - lat1) * p) / 2 +
    Math.cos(lat1 * p) *
    Math.cos(lat2 * p) *
    (1 - Math.cos((lng2 - lng1) * p)) / 2;

  return Math.round(
    12742000 * Math.asin(Math.sqrt(a))
  );
}

/* ---------- Jurisdiction キャッシュ ---------- */

function getJurisdiction_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(JUR_CACHE_KEY);

  if (hit) {
    try {
      return JSON.parse(hit);
    } catch (e) {}
  }

  const sh = ss_().getSheetByName(SHEETS.JURISDICTION);

  if (!sh || sh.getLastRow() < 2) return [];

  const rows = sh.getDataRange()
    .getValues()
    .slice(1)
    .map(r => ({
      name: r[0],
      type: r[1],
      south: Number(r[2]),
      west: Number(r[3]),
      north: Number(r[4]),
      east: Number(r[5]),
      bufferM: Number(r[6] || 200)
    }));

  cache.put(
    JUR_CACHE_KEY,
    JSON.stringify(rows),
    JUR_CACHE_TTL
  );

  return rows;
}

function classifyJurisdictionFast_(lat, lng, rows) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    if (
      isFinite(r.south) &&
      isFinite(r.north) &&
      isFinite(r.west) &&
      isFinite(r.east) &&
      lat >= r.south &&
      lat <= r.north &&
      lng >= r.west &&
      lng <= r.east
    ) {
      return '管轄内';
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    if (
      isFinite(r.south) &&
      isFinite(r.north) &&
      isFinite(r.west) &&
      isFinite(r.east) &&
      distanceToBoxM_(
        lat,
        lng,
        r.south,
        r.west,
        r.north,
        r.east
      ) <= Number(r.bufferM || 200)
    ) {
      return '周辺';
    }
  }

  return '未設定';
}

function classifyJurisdiction_(lat, lng) {
  return classifyJurisdictionFast_(
    lat,
    lng,
    getJurisdiction_()
  );
}

function distanceToBoxM_(lat, lng, s, w, n, e) {
  const cl = Math.max(s, Math.min(n, lat));
  const cn = Math.max(w, Math.min(e, lng));

  return distanceM_(
    lat,
    lng,
    cl,
    cn
  );
}

/* ---------- ジオコード ---------- */

function geocodeAddress_(address) {
  const text = String(address || '').trim();

  if (!text) {
    throw new Error('住所を入力してください。');
  }

  const r = Maps.newGeocoder()
    .setLanguage('ja')
    .setRegion('jp')
    .geocode(text);

  if (
    !r ||
    r.status !== 'OK' ||
    !r.results ||
    !r.results.length
  ) {
    throw new Error('住所を検索できませんでした。');
  }

  const g = r.results[0];

  return {
    address: g.formatted_address || text,
    lat: g.geometry.location.lat,
    lng: g.geometry.location.lng
  };
}

/* ---------- Google Routes API ---------- */

function getGoogleMapsServerKey_() {
  const key = PropertiesService.getScriptProperties()
    .getProperty('GOOGLE_MAPS_SERVER_KEY');

  if (!key) {
    throw new Error(
      'Google MapsのサーバーAPIキーが設定されていません。' +
      'Apps Scriptの「プロジェクトの設定」→「スクリプト プロパティ」に ' +
      'GOOGLE_MAPS_SERVER_KEY を登録してください。'
    );
  }

  return key.trim();
}

function computeRoute_(originLat, originLng, destLat, destLng, mode) {
  const olat = Number(originLat);
  const olng = Number(originLng);
  const dlat = Number(destLat);
  const dlng = Number(destLng);

  if (![olat, olng, dlat, dlng].every(isFinite)) {
    throw new Error('ルート計算用の座標が不正です。');
  }

  // ホース敷設距離は徒歩ネットワークを基準にする。
  // 車両用ルートの一方通行による遠回りを避ける目的。
  const travelMode = String(mode || 'WALK').toUpperCase() === 'WALK'
    ? 'WALK'
    : 'WALK';

  const url = 'https://routes.googleapis.com/directions/v2:computeRoutes';
  const body = {
    origin: {
      location: {
        latLng: { latitude: olat, longitude: olng }
      }
    },
    destination: {
      location: {
        latLng: { latitude: dlat, longitude: dlng }
      }
    },
    travelMode: travelMode,
    computeAlternativeRoutes: false,
    languageCode: 'ja',
    units: 'METRIC'
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Api-Key': getGoogleMapsServerKey_(),
      'X-Goog-FieldMask': 'routes.distanceMeters,routes.duration,routes.polyline.encodedPolyline'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    let msg = 'Google Routes APIエラー (' + code + ')';
    try {
      const err = JSON.parse(text);
      if (err.error && err.error.message) msg += ': ' + err.error.message;
    } catch (ignore) {}
    throw new Error(msg);
  }

  const data = JSON.parse(text);
  const route = data.routes && data.routes[0];

  if (!route || !route.distanceMeters) {
    throw new Error('Google Mapsでルートを取得できませんでした。');
  }

  return {
    distanceMeters: Number(route.distanceMeters),
    duration: route.duration || '',
    encodedPolyline: route.polyline
      ? String(route.polyline.encodedPolyline || '')
      : ''
  };
}

/* ---------- 現場 ---------- */

function getInitialData() {
  const incident = getActiveIncident_();

  return {
    incident: incident,
    markers: getMarkers_(incident),
    jurisdiction: getJurisdiction_()
  };
}

function startIncident_(address, lat, lng, radiusM) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    clearActiveIncident_();

    const id = 'I' + Utilities.getUuid().slice(0,8);

    ss_()
      .getSheetByName(SHEETS.INCIDENT)
      .appendRow([
        id,
        address,
        Number(lat),
        Number(lng),
        Number(radiusM || SEARCH_RADIUS_M),
        new Date(),
        true
      ]);

    const incident = getActiveIncident_();
    const jurisdiction = getJurisdiction_();
    const states = getIncidentWaterStates_(incident);

    return {
      incident: incident,
      markers: getMarkers_(incident),
      jurisdiction: jurisdiction,
      water: getNearbyWater_(
        incident.lat,
        incident.lng,
        incident.radiusM,
        states,
        jurisdiction
      )
    };

  } finally {
    lock.releaseLock();
  }
}

function getActiveIncident_() {
  const sh = ss_().getSheetByName(SHEETS.INCIDENT);

  if (!sh || sh.getLastRow() < 2) return null;

  const v = sh.getDataRange().getValues();

  for (let i = v.length - 1; i >= 1; i--) {
    if (
      v[i][6] === true ||
      String(v[i][6]).toUpperCase() === 'TRUE'
    ) {
      return {
        incidentId: v[i][0],
        address: v[i][1],
        lat: Number(v[i][2]),
        lng: Number(v[i][3]),
        radiusM: Number(v[i][4]),
        createdAt: v[i][5],
        active: true
      };
    }
  }

  return null;
}

function getMarkers_(incident) {
  const inc = incident || getActiveIncident_();
  const sh = ss_().getSheetByName(SHEETS.MARKERS);

  if (!sh || !inc || sh.getLastRow() < 2) return [];

  return sh.getDataRange()
    .getValues()
    .slice(1)
    .filter(r => r[1] === inc.incidentId)
    .map(r => ({
      markerId: r[0],
      type: r[2],
      name: r[3],
      lat: Number(r[4]),
      lng: Number(r[5]),
      status: r[6],
      geometry: r[7]
    }));
}

function getIncidentWaterStates_(incident) {
  const inc = incident || getActiveIncident_();
  const sh = ss_().getSheetByName(SHEETS.INCIDENT_WATER);
  const out = {};

  if (!sh || !inc || sh.getLastRow() < 2) return out;

  sh.getDataRange()
    .getValues()
    .slice(1)
    .forEach(r => {
      if (r[0] === inc.incidentId) {
        out[String(r[1])] = r[2];
      }
    });

  return out;
}

function getIncidentSnapshot_(lat, lng, radiusM) {
  const incident = getActiveIncident_();

  if (!incident) {
    return {
      incident: null,
      markers: [],
      jurisdiction: getJurisdiction_(),
      water: []
    };
  }

  const jurisdiction = getJurisdiction_();
  const states = getIncidentWaterStates_(incident);

  return {
    incident: incident,
    markers: getMarkers_(incident),
    jurisdiction: jurisdiction,
    water: getNearbyWater_(
      lat !== undefined && lat !== ''
        ? lat
        : incident.lat,
      lng !== undefined && lng !== ''
        ? lng
        : incident.lng,
      radiusM !== undefined && radiusM !== ''
        ? radiusM
        : incident.radiusM,
      states,
      jurisdiction
    )
  };
}

/* ---------- 水利状態 ---------- */

function saveWaterStatus_(waterId, status) {
  const allowed = [
    '未確認',
    '使用可能',
    '消防署使用中',
    '使用困難'
  ];

  if (allowed.indexOf(String(status)) < 0) {
    throw new Error('水利状態が不正です。');
  }

  const id = String(waterId || '').trim();
  if (!id) throw new Error('水利IDがありません。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const inc = getActiveIncident_();

    if (!inc) {
      throw new Error('現場がありません。');
    }

    const sh = ss_().getSheetByName(SHEETS.INCIDENT_WATER);
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][0]) === String(inc.incidentId) &&
        String(data[i][1]) === id
      ) {
        sh.getRange(i + 1, 3, 1, 2)
          .setValues([[String(status), new Date()]]);
        return {
          ok: true,
          waterId: id,
          status: String(status)
        };
      }
    }

    sh.appendRow([
      inc.incidentId,
      id,
      String(status),
      new Date()
    ]);

    return {
      ok: true,
      waterId: id,
      status: String(status)
    };

  } finally {
    lock.releaseLock();
  }
}

/* ---------- マーカー ---------- */

function addMarker_(type, name, lat, lng, status) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const inc = getActiveIncident_();

    if (!inc) {
      throw new Error('火災現場がありません。');
    }

    const sh = ss_().getSheetByName(SHEETS.MARKERS);
    const id = 'M' + Utilities.getUuid().slice(0,8);

    sh.appendRow([
      id,
      inc.incidentId,
      type,
      name || '',
      Number(lat),
      Number(lng),
      status || '',
      type === 'road' ? (status || '') : '',
      new Date()
    ]);

    return {
      markerId: id,
      type: type,
      name: name || '',
      lat: Number(lat),
      lng: Number(lng),
      status: status || ''
    };

  } finally {
    lock.releaseLock();
  }
}

function deleteMarker_(markerId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = ss_().getSheetByName(SHEETS.MARKERS);
    const v = sh.getDataRange().getValues();

    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) === String(markerId)) {
        sh.deleteRow(i + 1);
        return true;
      }
    }

    return false;

  } finally {
    lock.releaseLock();
  }
}

/* ---------- 現場終了 ---------- */

function clearActiveIncident_() {
  const sh = ss_().getSheetByName(SHEETS.INCIDENT);

  if (!sh || sh.getLastRow() <= 1) return;

  const v = sh.getDataRange().getValues();

  for (let i = v.length - 1; i >= 1; i--) {
    if (
      v[i][6] === true ||
      String(v[i][6]).toUpperCase() === 'TRUE'
    ) {
      sh.getRange(i + 1, 7).setValue(false);
    }
  }
}

function endIncident_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    clearActiveIncident_();
    return true;
  } finally {
    lock.releaseLock();
  }
}
