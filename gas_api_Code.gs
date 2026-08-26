/**
 * 🚒 消防団 現場支援マップ - GAS API
 *
 * GitHub Pages上のフロントエンドからJSONPで呼び出すAPI。
 * Supabaseは使用しません。
 *
 * 使い方：
 * 1. このCode.gsを「水利テスト用Googleスプレッドシート」に紐づくApps Scriptへ貼る
 * 2. 必要なら SPREADSHEET_ID を入力
 * 3. setupAll()を1回実行
 * 4. ウェブアプリとしてデプロイ（自分として実行／アクセスできるユーザーをテスト対象に合わせる）
 * 5. WebアプリURLをGitHub版index.htmlのAPI_URLへ設定
 *
 * ※GitHub PagesからGASを呼ぶため、JSONPを使用しています。
 *   テスト運用を前提にし、団員の個人情報などはこのAPIに入れないでください。
 */
const SPREADSHEET_ID = ''; // 例：1AbC...。空欄なら紐づいたスプレッドシートを使用
const SEARCH_RADIUS_M = 200;
const SOURCE_DATE = '令和7年4月1日';
const IMPORT_SHEETS = ['IMPORT_消火栓','IMPORT_防火水槽','消火栓','防火水槽等'];

// v12: 対象範囲を固定。東京全域は扱わない。
// 南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200mをカバーするための外接範囲。
// 正確な町丁目ポリゴンを後から入れられるよう、範囲は定数として分離。
const TARGET_AREA = {south:35.5450,north:35.5665,west:139.7150,east:139.7555,bufferM:200};
const WATER_CACHE_KEY = 'TARGET_WATER_V12';
const WATER_CACHE_TTL = 21600; // 6時間。水利基本データはほぼ読み取り専用。
const POLL_SECONDS = 20;

function isInTargetArea_(lat,lng){
  return Number(lat)>=TARGET_AREA.south && Number(lat)<=TARGET_AREA.north && Number(lng)>=TARGET_AREA.west && Number(lng)<=TARGET_AREA.east;
}

const SHEETS = {
  WATER_LOG:'Water_Update_Log',
  MEMBERS:'Members',
  INCIDENT:'Incident',
  INCIDENT_WATER:'IncidentWater',
  MARKERS:'Markers',
  JURISDICTION:'Jurisdiction'
};

function ss_(){
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function doGet(e){
  const p=e&&e.parameter||{};
  const action=p.action||'health';
  const callback=p.callback;
  let out;
  try{
    out=dispatch_(action,p);
    return respond_(callback,{ok:true,data:out});
  }catch(err){
    return respond_(callback,{ok:false,error:String(err&&err.message||err)});
  }
}
function respond_(callback,obj){
  const json=JSON.stringify(obj).replace(/</g,'\\u003c');
  if(callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)){
    return ContentService.createTextOutput(callback+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function dispatch_(action,p){
  switch(action){
    case 'health': return {ok:true,time:new Date().toISOString()};
    case 'setupAll': return setupAll();
    case 'getWaterDataInfo': return getWaterDataInfo();
    case 'getInitialData': return getInitialData();
    case 'getNearbyWater': return getNearbyWater_(p.lat,p.lng,p.radiusM);
    case 'getIncidentSnapshot': return getIncidentSnapshot_(p.lat,p.lng,p.radiusM);
    case 'refreshWaterCache': return refreshWaterCache();
    case 'geocodeAddress': return geocodeAddress_(p.address);
    case 'startIncident': return startIncident_(p.address,p.lat,p.lng,p.radiusM);
    case 'saveWaterStatus': return saveWaterStatus_(p.waterId,p.status);
    case 'addMarker': return addMarker_(p.type,p.name,p.lat,p.lng,p.status);
    case 'deleteMarker': return deleteMarker_(p.markerId);
    case 'endIncident': return endIncident_();
    default: throw new Error('Unknown action: '+action);
  }
}
function setupAll(){
  const ss=ss_();
  const defs={
    Water_Update_Log:['更新日時','件数','基準日','結果'],
    Members:['団員名'],
    Incident:['incidentId','address','lat','lng','radiusM','createdAt','active'],
    IncidentWater:['incidentId','waterId','status','updatedAt'],
    Markers:['markerId','incidentId','type','name','lat','lng','status','geometry','createdAt'],
    Jurisdiction:['名称','種別','南緯度','西経度','北緯度','東経度','周辺拡張m','備考']
  };
  Object.keys(defs).forEach(name=>{
    const sh=ss.getSheetByName(name)||ss.insertSheet(name);
    if(sh.getLastRow()===0)sh.getRange(1,1,1,defs[name].length).setValues([defs[name]]);
    sh.setFrozenRows(1);
  });
  const jur=ss.getSheetByName('Jurisdiction');
  if(jur.getLastRow()<2){
    const rows=[
      ['南蒲田3丁目','管轄内','','','','',200,''],
      ['西糀谷1丁目','管轄内','','','','',200,''],
      ['西糀谷2丁目','管轄内','','','','',200,''],
      ['西糀谷3丁目','管轄内','','','','',200,''],
      ['西糀谷4丁目','管轄内','','','','',200,''],
      ['北糀谷','管轄内','','','','',200,'']
    ];
    jur.getRange(2,1,rows.length,8).setValues(rows);
  }
  return '初期設定完了';
}
function getImportSheets_(){
  return ss_().getSheets().filter(sh=>{
    const n=sh.getName();
    return IMPORT_SHEETS.indexOf(n)>=0 || /^IMPORT_/.test(n);
  });
}
function getWaterDataInfo(){
  const cached=getCachedWater_();
  const count=cached.length;
  const log=ss_().getSheetByName(SHEETS.WATER_LOG);
  let updatedAt='',basis=SOURCE_DATE;
  if(log&&log.getLastRow()>=2){
    const r=log.getRange(log.getLastRow(),1,1,4).getValues()[0];
    if(r[0])updatedAt=Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),'yyyy/MM/dd HH:mm');
    if(r[2])basis=String(r[2]);
  }
  return {count,basis,updatedAt,source:'東京都オープンデータ／東京消防庁',target:'南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200m',cacheSeconds:WATER_CACHE_TTL};
}

// 水利の基本データはCacheServiceへ。40人が同時閲覧しても、毎回シート全体を読むのを避ける。
function cachePutWater_(data){
  const cache=CacheService.getScriptCache(), json=JSON.stringify(data);
  // CacheServiceの1値上限を避けるため、必要なら分割。
  cache.remove(WATER_CACHE_KEY+'_COUNT');
  if(json.length<=90000){ cache.put(WATER_CACHE_KEY,json,WATER_CACHE_TTL); return; }
  cache.remove(WATER_CACHE_KEY);
  const chunkSize=80000, count=Math.ceil(json.length/chunkSize);
  for(let i=0;i<count;i++) cache.put(WATER_CACHE_KEY+'_'+i,json.slice(i*chunkSize,(i+1)*chunkSize),WATER_CACHE_TTL);
  cache.put(WATER_CACHE_KEY+'_COUNT',String(count),WATER_CACHE_TTL);
}
function getCachedWater_(){
  const cache=CacheService.getScriptCache();
  const hit=cache.get(WATER_CACHE_KEY);
  if(hit){try{return JSON.parse(hit)}catch(e){}}
  const cnt=Number(cache.get(WATER_CACHE_KEY+'_COUNT')||0);
  if(cnt>0){
    let json=''; for(let i=0;i<cnt;i++){const part=cache.get(WATER_CACHE_KEY+'_'+i); if(!part){json='';break;} json+=part;}
    if(json){try{return JSON.parse(json)}catch(e){}}
  }
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    // 他の端末が先にキャッシュを作った場合は、それをそのまま使う。
    const again=cache.get(WATER_CACHE_KEY);
    if(again){try{return JSON.parse(again)}catch(e){}}
    const data=buildTargetWaterCache_(); cachePutWater_(data); return data;
  } finally { lock.releaseLock(); }
}
function buildTargetWaterCache_(){
  const out=[];
  getImportSheets_().forEach(sh=>{
    const values=sh.getDataRange().getDisplayValues();
    if(!values.length)return;
    const meta=detectSheetMeta_(values);
    const type=String(sh.getName()).indexOf('防火')>=0?'防火水槽等':'消火栓';
    const start=meta.headerRow>=0?meta.headerRow+1:0;
    for(let i=start;i<values.length;i++){
      const row=values[i],coord=extractCoordinateFromRow_(row,meta);
      if(!coord || !isInTargetArea_(coord.lat,coord.lng))continue;
      const id=buildWaterId_(sh.getName(),i+1,row,meta);
      out.push({id,type,lat:coord.lat,lng:coord.lng,capacity:meta.capacityCol>=0?row[meta.capacityCol]:'',address:meta.addressCol>=0?row[meta.addressCol]:'',source:'東京都オープンデータ／東京消防庁',sourceDate:SOURCE_DATE});
    }
  });
  return out;
}
function refreshWaterCache(){
  const data=buildTargetWaterCache_();
  cachePutWater_(data);
  return {count:data.length,target:'南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200m'};
}
function getNearbyWater_(lat,lng,radiusM){
  const radius=Math.min(Number(radiusM||SEARCH_RADIUS_M),SEARCH_RADIUS_M),lat0=Number(lat),lng0=Number(lng);
  if(!isFinite(lat0)||!isFinite(lng0))return[];
  const latPad=radius/111320,lngPad=radius/(111320*Math.max(.5,Math.cos(lat0*Math.PI/180)));
  const states=getIncidentWaterStates_(),base=getCachedWater_(),out=[];
  base.forEach(w=>{
    if(Math.abs(w.lat-lat0)>latPad||Math.abs(w.lng-lng0)>lngPad)return;
    const d=distanceM_(lat0,lng0,w.lat,w.lng); if(d>radius)return;
    out.push(Object.assign({},w,{sourceDate:SOURCE_DATE,updatedAt:'',zone:classifyJurisdiction_(w.lat,w.lng),distance:d,status:states[w.id]||'未確認'}));
  });
  out.sort((a,b)=>a.distance-b.distance); return out.slice(0,100);
}
function detectSheetMeta_(values){
  const maxScan=Math.min(values.length,25);
  let headerRow=-1,latCol=-1,lngCol=-1,capacityCol=-1,addressCol=-1,idCol=-1;
  const latTokens=['緯度','latitude','lat','y'],lngTokens=['経度','longitude','lng','lon','x'],capTokens=['容量','貯水量','水量','capacity'],addrTokens=['住所','所在地','場所'],idTokens=['水利番号','水利id','id','番号','施設番号'];
  for(let r=0;r<maxScan;r++){
    const row=values[r].map(v=>String(v).trim().toLowerCase());
    for(let c=0;c<row.length;c++){
      const s=row[c];
      if(latCol<0&&latTokens.some(t=>s===t||s.indexOf(t)>=0)&&s.indexOf('経')<0)latCol=c;
      if(lngCol<0&&lngTokens.some(t=>s===t||s.indexOf(t)>=0))lngCol=c;
      if(capacityCol<0&&capTokens.some(t=>s.indexOf(t)>=0))capacityCol=c;
      if(addressCol<0&&addrTokens.some(t=>s.indexOf(t)>=0))addressCol=c;
      if(idCol<0&&idTokens.some(t=>s===t||s.indexOf(t)>=0))idCol=c;
    }
    if(latCol>=0&&lngCol>=0){headerRow=r;break;}
  }
  return {headerRow,latCol,lngCol,capacityCol,addressCol,idCol};
}
function extractCoordinateFromRow_(row,meta){
  if(meta.latCol>=0&&meta.lngCol>=0){
    const lat=toNumber_(row[meta.latCol]),lng=toNumber_(row[meta.lngCol]);
    if(isValidCoord_(lat,lng))return{lat,lng};
  }
  const nums=[];
  for(let c=0;c<row.length;c++){const n=toNumber_(row[c]);if(isFinite(n))nums.push(n)}
  for(let i=0;i<nums.length;i++)for(let j=0;j<nums.length;j++)if(i!==j&&isValidCoord_(nums[i],nums[j]))return{lat:nums[i],lng:nums[j]};
  return null;
}
function toNumber_(v){if(v===null||v===undefined||v==='')return NaN;return Number(String(v).replace(/,/g,'').replace(/[^\d.\-+]/g,''))}
function isValidCoord_(lat,lng){return isFinite(lat)&&isFinite(lng)&&lat>=20&&lat<=46&&lng>=122&&lng<=154}
function buildWaterId_(sheetName,rowNo,row,meta){const raw=meta.idCol>=0?String(row[meta.idCol]||'').trim():'';return raw?sheetName+':'+raw:sheetName+':row'+rowNo}
function distanceM_(lat1,lng1,lat2,lng2){const p=Math.PI/180,a=.5-Math.cos((lat2-lat1)*p)/2+Math.cos(lat1*p)*Math.cos(lat2*p)*(1-Math.cos((lng2-lng1)*p))/2;return Math.round(12742000*Math.asin(Math.sqrt(a)))}
function geocodeAddress_(address){
  const r=Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(String(address||'').trim());
  if(!r||r.status!=='OK'||!r.results||!r.results.length)throw new Error('住所を検索できませんでした。');
  const g=r.results[0];
  return {address:g.formatted_address||address,lat:g.geometry.location.lat,lng:g.geometry.location.lng};
}
function startIncident_(address,lat,lng,radiusM){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{ clearActiveIncident_();
  const id='I'+Utilities.getUuid().slice(0,8);
    ss_().getSheetByName(SHEETS.INCIDENT).appendRow([id,address,Number(lat),Number(lng),Number(radiusM||SEARCH_RADIUS_M),new Date(),true]);
    return getInitialData();
  } finally { lock.releaseLock(); }
}
function getInitialData(){
  return {members:getMembers_(),incident:getActiveIncident_(),markers:getMarkers_(),jurisdiction:getJurisdiction_()};
}
function getMembers_(){
  const sh=ss_().getSheetByName(SHEETS.MEMBERS);
  return sh&&sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().filter(Boolean):[];
}
function getActiveIncident_(){
  const sh=ss_().getSheetByName(SHEETS.INCIDENT);
  if(!sh||sh.getLastRow()<2)return null;
  const v=sh.getDataRange().getValues();
  for(let i=v.length-1;i>=1;i--)if(v[i][6]===true||String(v[i][6]).toUpperCase()==='TRUE')return{incidentId:v[i][0],address:v[i][1],lat:Number(v[i][2]),lng:Number(v[i][3]),radiusM:Number(v[i][4]),createdAt:v[i][5],active:true};
  return null;
}
function getMarkers_(){
  const sh=ss_().getSheetByName(SHEETS.MARKERS),inc=getActiveIncident_();
  if(!sh||!inc||sh.getLastRow()<2)return[];
  return sh.getDataRange().getValues().slice(1).filter(r=>r[1]===inc.incidentId).map(r=>({markerId:r[0],type:r[2],name:r[3],lat:Number(r[4]),lng:Number(r[5]),status:r[6],geometry:r[7]}));
}
function getIncidentWaterStates_(){
  const sh=ss_().getSheetByName(SHEETS.INCIDENT_WATER),inc=getActiveIncident_(),o={};
  if(!sh||!inc||sh.getLastRow()<2)return o;
  sh.getDataRange().getValues().slice(1).filter(r=>r[0]===inc.incidentId).forEach(r=>o[r[1]]=r[2]);
  return o;
}
function getJurisdiction_(){
  const sh=ss_().getSheetByName(SHEETS.JURISDICTION);
  if(!sh||sh.getLastRow()<2)return[];
  return sh.getDataRange().getValues().slice(1).map(r=>({name:r[0],type:r[1],south:r[2],west:r[3],north:r[4],east:r[5],bufferM:r[6]||200}));
}
function classifyJurisdiction_(lat,lng){
  const rows=getJurisdiction_();
  for(const r of rows)if(isFinite(r.south)&&isFinite(r.north)&&isFinite(r.west)&&isFinite(r.east)&&lat>=r.south&&lat<=r.north&&lng>=r.west&&lng<=r.east)return'管轄内';
  for(const r of rows)if(isFinite(r.south)&&isFinite(r.north)&&isFinite(r.west)&&isFinite(r.east)&&distanceToBoxM_(lat,lng,r.south,r.west,r.north,r.east)<=Number(r.bufferM||200))return'周辺';
  return'未設定';
}
function distanceToBoxM_(lat,lng,s,w,n,e){const cl=Math.max(s,Math.min(n,lat)),cn=Math.max(w,Math.min(e,lng));return distanceM_(lat,lng,cl,cn)}
function saveWaterStatus_(waterId,status){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const sh=ss_().getSheetByName(SHEETS.INCIDENT_WATER),inc=getActiveIncident_();
    if(!inc)throw new Error('現場がありません。');
    const data=sh.getDataRange().getValues();
    for(let i=1;i<data.length;i++){
      if(data[i][0]===inc.incidentId&&data[i][1]===waterId){
        sh.getRange(i+1,3,1,2).setValues([[status,new Date()]]); return true;
      }
    }
    sh.appendRow([inc.incidentId,waterId,status,new Date()]); return true;
  } finally { lock.releaseLock(); }
}
function addMarker_(type,name,lat,lng,status){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const sh=ss_().getSheetByName(SHEETS.MARKERS),inc=getActiveIncident_();
    if(!inc)throw new Error('火災現場がありません。');
    const id='M'+Utilities.getUuid().slice(0,8);
    // road: status contains the encoded polyline points; store it in geometry too.
    sh.appendRow([id,inc.incidentId,type,name,Number(lat),Number(lng),status||'',type==='road'?(status||''):'',new Date()]);
    return id;
  } finally { lock.releaseLock(); }
}
function deleteMarker_(markerId){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const sh=ss_().getSheetByName(SHEETS.MARKERS),v=sh.getDataRange().getValues();
    for(let i=1;i<v.length;i++)if(v[i][0]===markerId){sh.deleteRow(i+1);return true;}
    return false;
  } finally { lock.releaseLock(); }
}
function clearActiveIncident_(){
  const ss=ss_(),inc=ss.getSheetByName(SHEETS.INCIDENT);
  if(inc&&inc.getLastRow()>1){
    const v=inc.getDataRange().getValues();
    for(let i=v.length-1;i>=1;i--)if(v[i][6]===true||String(v[i][6]).toUpperCase()==='TRUE')inc.getRange(i+1,7).setValue(false);
  }
  // Incident / IncidentWater / Markers は履歴として残す。現在の現場だけactive=falseにする。
}
function endIncident_(){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{ clearActiveIncident_(); return true; }
  finally { lock.releaseLock(); }
}
function getTargetAreaInfo(){
  return {
    names:['南蒲田3丁目','西糀谷1丁目','西糀谷2丁目','西糀谷3丁目','西糀谷4丁目','北糀谷'],
    south:TARGET_AREA.south,north:TARGET_AREA.north,west:TARGET_AREA.west,east:TARGET_AREA.east,bufferM:TARGET_AREA.bufferM,
    note:'南蒲田3丁目・西糀谷1〜4丁目・北糀谷＋周囲200mを最大対象範囲とします。'
  };
}
function getIncidentSnapshot_(lat,lng,radiusM){
  return {
    incident:getActiveIncident_(),
    markers:getMarkers_(),
    jurisdiction:getJurisdiction_(),
    water:getNearbyWater_(lat,lng,radiusM)
  };
}

