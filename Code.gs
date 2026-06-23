/**********************************************************************
 * FUNDEDNEXT PARTNERS DASHBOARD  —  Code.gs   (single combined app)
 * ------------------------------------------------------------------
 * One web app, two tabs inside Index.html:
 *   • Dashboard       — activity status, RM response desk, content
 *                       cards, searchable partner database.
 *   • Content Tracker — opens each partner's linked tracker sheet
 *                       LIVE and lists every post with its date.
 *
 * Reads the "Futures" and "RM_AVG" tabs of your Partners sheet.
 *
 * HOW TO PUBLISH
 *   1. Open your Google Sheet → Extensions → Apps Script.
 *   2. Paste this file into Code.gs and Index.html into a file named
 *      exactly "Index" (HTML).
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as:  Me
 *        Who has access:  Anyone (or "Anyone with link")
 *   4. Open the web-app URL. The first run asks permission to read
 *      spreadsheets (needed to follow each tracker link) — allow it.
 *
 * If the script is NOT bound to the sheet, paste the sheet ID below.
 *********************************************************************/

var SPREADSHEET_ID = '';          // leave '' if bound to the sheet
var FUTURES_SHEET  = 'Sheet1';    // main partner tab (auto-detected if renamed)
var RM_AVG_SHEET   = 'RM_AVG';
var RM_ART_SHEET   = 'RM_ART';

/* Map the e-mails used in the "Assigned RM" column to a display name.
   The name is also used to look up the average response time in RM_AVG.
   Add/adjust rows here if RMs change.                                   */
var RM_DIRECTORY = {
  'sadia.akter@nextventures.io' : 'Sadia Akter',
  'shirajul@nextventures.io'    : 'Shiraj Rifat',
  'sanjida.ahmed@nextventures.io': 'Sanjida Ahmed Orna',
  'naimul.islam@nextventures.io': 'Naimul Islam Durjoy',
  'zaion.abrar@nextventures.io' : 'Mohammad Zaion Abrar Akhand'
};

/* ------------------------------------------------------------------ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FundedNext Partners Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBook_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/* normalise a header so newline/whitespace variants still match */
function norm_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/* ================================================================== */
/*  MAIN ENDPOINT                                                     */
/* ================================================================== */
function getDashboardData() {
  var book = getBook_();
  var rmAvg = readRmAverages_(book);
  var partners = readFutures_(book, rmAvg);

  /* ---- aggregate the headline numbers ---- */
  var stats = {
    total: partners.length,
    superActive: 0, lessActive: 0, dormant: 0, notInRespond: 0, unset: 0,
    highValue: 0, midValue: 0, lowValue: 0,
    retainer: 0, freeOnboarding: 0,
    withVideo: 0, withContract: 0
  };

  partners.forEach(function (p) {
    if      (p.statusKey === 'super')   stats.superActive++;
    else if (p.statusKey === 'less')    stats.lessActive++;
    else if (p.statusKey === 'dormant') stats.dormant++;
    else if (p.statusKey === 'notin')   stats.notInRespond++;
    else                                stats.unset++;

    if (/high/i.test(p.value)) stats.highValue++;
    else if (/mid/i.test(p.value))  stats.midValue++;
    else if (/low/i.test(p.value))  stats.lowValue++;

    if (/retainer/i.test(p.dealType)) stats.retainer++;
    else if (/free/i.test(p.dealType)) stats.freeOnboarding++;

    if (p.featuredVideo && p.featuredVideo.id) stats.withVideo++;
    if (p.contract && p.contract.url)          stats.withContract++;
  });

  /* videos for the moving strip (featured videos only) */
  var videos = partners.filter(function (p) {
    return p.featuredVideo && p.featuredVideo.id;
  }).map(function (p) {
    return {
      name: p.name, username: p.username, country: p.country,
      status: p.status, statusKey: p.statusKey,
      id: p.featuredVideo.id, thumb: p.featuredVideo.thumb,
      embed: p.featuredVideo.embed, url: p.featuredVideo.url
    };
  });

  /* RM roster with counts + average response time */
  var rmCounts = {};
  partners.forEach(function (p) {
    if (!p.rmName) return;
    rmCounts[p.rmName] = (rmCounts[p.rmName] || 0) + 1;
  });
  var rmRoster = Object.keys(rmCounts).map(function (name) {
    return { name: name, partners: rmCounts[name], art: rmAvg.byName[name] || '—' };
  }).sort(function (a, b) { return b.partners - a.partners; });

  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy, h:mm a'),
    stats: stats,
    partners: partners,
    videos: videos,
    rmRoster: rmRoster,
    teamArt: {
      futures: rmAvg.byName['Futures RM'] || '—',
      cfd: rmAvg.byName['CFD RM'] || '—',
      onboarding: rmAvg.byName['Onboarding Team'] || '—',
      overall: rmAvg.overall || '—'
    }
  };
}

/* ================================================================== */
/*  READ  RM_AVG                                                      */
/* ================================================================== */
function readRmAverages_(book) {
  var out = { byName: {}, overall: '—' };
  var sh = book.getSheetByName(RM_AVG_SHEET);
  if (!sh) return out;
  var values = sh.getDataRange().getValues();
  values.forEach(function (row, i) {
    var key = String(row[0] == null ? '' : row[0]).trim();
    var val = String(row[1] == null ? '' : row[1]).trim();
    if (!key) return;
    if (i === 0 && /^=/.test(key)) {          // header formula row -> overall avg
      if (val) out.overall = val;
      return;
    }
    if (/^=/.test(key)) return;               // skip stray formulas
    out.byName[key] = val;
  });
  return out;
}

/* ================================================================== */
/*  READ  MAIN PARTNER SHEET  (Sheet1 / Futures — auto-detected)      */
/* ================================================================== */
function mainSheet_(book){
  var sh = book.getSheetByName(FUTURES_SHEET);
  if (sh) return sh;
  var all = book.getSheets();
  for (var i=0;i<all.length;i++){
    var hdr = all[i].getRange(1,1,1,Math.min(40, all[i].getLastColumn()||1)).getValues()[0].map(norm_);
    if (hdr.indexOf('name')>-1 && hdr.some(function(h){return h.indexOf('assigned rm')>-1;})) return all[i];
  }
  return book.getSheets()[0];
}

/* ================================================================== */
/*  SMART-CHIP / HYPERLINK READER                                     */
/*  Returns { "row,col": url } (0-indexed, aligned with getDataRange) */
/*  Reads file/link smart chips, plain hyperlinks and run links via   */
/*  the Sheets advanced service. Requires the "Google Sheets API"     */
/*  advanced service to be enabled (Extensions ▸ Apps Script ▸        */
/*  Services ▸ + ▸ Google Sheets API). Fails silently if unavailable. */
/* ================================================================== */
function chipLinkMap_(sheetName){
  var rows = chipRowData_(sheetName);          // [{values:[{...}]}] from whichever route works
  var map = {};
  for (var r = 0; r < rows.length; r++){
    var cells = (rows[r] && rows[r].values) || [];
    for (var c = 0; c < cells.length; c++){
      var cell = cells[c]; if (!cell) continue;
      var url = '';
      if (cell.hyperlink) url = cell.hyperlink;
      if (!url && cell.chipRuns){
        for (var i = 0; i < cell.chipRuns.length; i++){
          var ch = cell.chipRuns[i];
          if (ch && ch.chip && ch.chip.richLinkProperties && ch.chip.richLinkProperties.uri){
            url = ch.chip.richLinkProperties.uri; break;
          }
        }
      }
      if (!url && cell.textFormatRuns){
        for (var j = 0; j < cell.textFormatRuns.length; j++){
          var fr = cell.textFormatRuns[j];
          if (fr && fr.format && fr.format.link && fr.format.link.uri){ url = fr.format.link.uri; break; }
        }
      }
      if (url) map[r + ',' + c] = url;
    }
  }
  return map;
}

/* Get rowData (with chip/hyperlink info) via two routes:
   1) the Sheets advanced service (if added in the editor), then
   2) a direct Sheets REST call signed with the script's own OAuth token.
   Whichever returns first wins; both fail silently → RichText fallback.    */
function chipRowData_(sheetName){
  var fields = 'sheets.data.rowData.values(hyperlink,chipRuns.chip.richLinkProperties.uri,textFormatRuns.format.link.uri)';
  /* route 1 — advanced service */
  try {
    if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets){
      var resp = Sheets.Spreadsheets.get(getBook_().getId(), {
        ranges: [sheetName], includeGridData: true, fields: fields
      });
      var s = (resp.sheets || [])[0];
      if (s && s.data && s.data[0] && s.data[0].rowData) return s.data[0].rowData;
    }
  } catch (e){ /* fall through to REST */ }
  /* route 2 — REST with the script token (needs the Sheets API reachable) */
  try {
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + getBook_().getId()
            + '?ranges=' + encodeURIComponent(sheetName)
            + '&includeGridData=true&fields=' + encodeURIComponent(fields);
    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200){
      var json = JSON.parse(res.getContentText());
      var sh2 = (json.sheets || [])[0];
      if (sh2 && sh2.data && sh2.data[0] && sh2.data[0].rowData) return sh2.data[0].rowData;
    }
  } catch (e2){ /* give up → RichText fallback */ }
  return [];
}

function readFutures_(book, rmAvg) {
  var sh = mainSheet_(book);
  if (!sh) return [];
  var rng = sh.getDataRange();
  var values = rng.getValues();
  var rich   = rng.getRichTextValues();
  if (values.length < 2) return [];

  var headers = values[0].map(norm_);
  /* find first column whose header contains ANY of the given phrases */
  function pick(){
    for (var a=0;a<arguments.length;a++){
      var kw = norm_(arguments[a]);
      for (var j=0;j<headers.length;j++){ if (headers[j] && headers[j].indexOf(kw)>-1) return j; }
    }
    return -1;
  }
  var idx = {
    name:     pick('name'),
    username: pick('platform username','username'),
    plinks:   pick('social media links','platform links','social media'),
    status:   pick('status'),
    country:  pick('country'),
    email:    pick('partner email','email'),
    code:     pick('coupon code','code'),
    phone:    pick('phone'),
    platform: pick('type of creator','platform','creator'),
    value:    pick('value of the partner','value'),
    dealType: pick('deal type'),
    retainer: pick('compensation','retainer'),
    rm:       pick('assigned rm'),
    tracker:  pick('content tracker'),
    drive:    pick('drive link'),
    remarks:  pick('remarks'),
    meetings: pick('meetings'),
    discord:  pick('discord'),
    video:    pick('featured video'),
    contract: pick('docusign contract','contract link'),
    cStart:   pick('contract start'),
    cEnd:     pick('contract end'),
    deliver:  pick('deliverables')
  };

  /* Read chip/hyperlink/run data ONCE via the Sheets API (file & link smart
     chips, plain hyperlinks, and every per-run link). Used to resolve single
     link cells (tracker/drive) AND to harvest every social link in column C. */
  var rowData = chipRowData_(sh.getName());
  var chipMap = {};                                    // first url per cell
  for (var rr = 0; rr < rowData.length; rr++){
    var cs = (rowData[rr] && rowData[rr].values) || [];
    for (var cc = 0; cc < cs.length; cc++){
      var cd = cs[cc]; if (!cd) continue;
      var u1 = cd.hyperlink || '';
      if (!u1 && cd.chipRuns){ for (var ii=0; ii<cd.chipRuns.length; ii++){ var chh=cd.chipRuns[ii]; if(chh&&chh.chip&&chh.chip.richLinkProperties&&chh.chip.richLinkProperties.uri){ u1=chh.chip.richLinkProperties.uri; break; } } }
      if (!u1 && cd.textFormatRuns){ for (var jj=0; jj<cd.textFormatRuns.length; jj++){ var frr=cd.textFormatRuns[jj]; if(frr&&frr.format&&frr.format.link&&frr.format.link.uri){ u1=frr.format.link.uri; break; } } }
      if (u1) chipMap[rr+','+cc] = u1;
    }
  }
  /* every link in a cell (all chip uris + all run uris) from the Sheets API */
  function allCellUrls_(r, ci){
    var urls = [];
    var cell = rowData[r] && rowData[r].values && rowData[r].values[ci];
    if (cell){
      if (cell.hyperlink) urls.push(cell.hyperlink);
      if (cell.chipRuns) cell.chipRuns.forEach(function(ch){ var u=ch&&ch.chip&&ch.chip.richLinkProperties&&ch.chip.richLinkProperties.uri; if(u) urls.push(u); });
      if (cell.textFormatRuns) cell.textFormatRuns.forEach(function(fr){ var u=fr&&fr.format&&fr.format.link&&fr.format.link.uri; if(u) urls.push(u); });
    }
    return urls;
  }

  /* embedded URL for a cell: smart chip → hyperlink → run link → text URL */
  function linkAt(r, ci){
    if (ci < 0) return '';
    var key = r + ',' + ci;
    if (chipMap[key]) return chipMap[key];
    var url = '';
    if (rich[r] && rich[r][ci]){
      var rt = rich[r][ci];
      url = rt.getLinkUrl() || '';
      if (!url){ rt.getRuns().forEach(function(run){ if(!url) url = run.getLinkUrl()||''; }); }
    }
    if (!url) url = firstUrl_(values[r][ci]);
    return url;
  }
  function fieldAt(r, ci, icon){
    if (ci < 0) return null;
    var text = cell_(values[r], ci);
    var url  = linkAt(r, ci);
    if (!text && !url) return null;
    return { text: text || url, url: url, isUrl: !!url, icon: icon };
  }
  /* social/platform links from column C: plain-text URLs + bare domains +
     EVERY embedded hyperlink/chip (handles multiple links in one cell) */
  function socialLinks_(r, ci){
    if (ci < 0) return [];
    var links = parseLinks_(cell_(values[r], ci));
    var seen = {};
    links.forEach(function(l){ seen[l.url.toLowerCase().replace(/\/+$/,'')] = 1; });
    function push_(u, label){
      if (!u) return;
      var mk = makeLink_(label || '', u);
      if (!mk || !mk.url) return;
      var k = mk.url.toLowerCase().replace(/\/+$/,'');
      if (!seen[k]){ seen[k] = 1; links.push(mk); }
    }
    /* all embedded links via the Sheets API (multiple per cell) */
    allCellUrls_(r, ci).forEach(function(u){ push_(u, ''); });
    /* RichText fallback when the Sheets API isn't reachable */
    if (rich[r] && rich[r][ci]){
      var rt = rich[r][ci];
      (rt.getRuns ? rt.getRuns() : []).forEach(function(run){
        push_(run.getLinkUrl && run.getLinkUrl(), (run.getText && run.getText()) || '');
      });
      push_(rt.getLinkUrl && rt.getLinkUrl(), '');
    }
    return links;
  }

  var tz = Session.getScriptTimeZone();
  var partners = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var name = cell_(row, idx.name);
    if (!name) continue;

    var statusRaw = cell_(row, idx.status);
    var rmEmail   = cell_(row, idx.rm);
    var rmName    = rmNameFromEmail_(rmEmail);
    var trackerUrl = linkAt(r, idx.tracker);

    partners.push({
      name:        name,
      username:    cell_(row, idx.username),
      status:      statusRaw || 'Unset',
      statusKey:   statusKey_(statusRaw),
      country:     cell_(row, idx.country),
      email:       cell_(row, idx.email),
      code:        cell_(row, idx.code),
      phone:       cell_(row, idx.phone),
      platform:    cell_(row, idx.platform),
      value:       cell_(row, idx.value),
      dealType:    cell_(row, idx.dealType),
      retainer:    cell_(row, idx.retainer),
      onboardedBy: '',
      pmName:      '',
      rmEmail:     rmEmail,
      rmName:      rmName,
      rmArt:       rmName ? (rmAvg.byName[rmName] || '—') : '—',
      assigned:    !!statusRaw,
      remarks:     cell_(row, idx.remarks),
      deliverables: cell_(row, idx.deliver),
      platformLinks: socialLinks_(r, idx.plinks),
      contentTracker: fieldAt(r, idx.tracker, '📊'),
      trackerLabel: cell_(row, idx.tracker),
      trackerUrl:   trackerUrl,
      hasTracker:   !!trackerUrl,
      drive:        fieldAt(r, idx.drive, '📁'),
      discord:      fieldAt(r, idx.discord, '💬'),
      meetings:     fieldAt(r, idx.meetings, '🎥'),
      contract: {
        url:   linkAt(r, idx.contract),
        raw:   cell_(row, idx.contract),
        start: fmtDate_(row[idx.cStart], tz),
        end:   fmtDate_(row[idx.cEnd], tz)
      },
      featuredVideo: youTube_(linkAt(r, idx.video) || cell_(row, idx.video))
    });
  }
  return partners;
}

/* ----------------------- small helpers ---------------------------- */
function cell_(row, i) {
  if (i < 0 || i >= row.length) return '';
  var v = row[i];
  return v == null ? '' : String(v).trim();
}

function statusKey_(s) {
  s = norm_(s);
  if (!s) return 'unset';
  if (s.indexOf('super') > -1) return 'super';
  if (s.indexOf('less')  > -1) return 'less';
  if (s.indexOf('dormant') > -1) return 'dormant';
  if (s.indexOf('not in') > -1 || s.indexOf('respond') > -1) return 'notin';
  return 'unset';
}

function rmNameFromEmail_(email) {
  if (!email) return '';
  var key = String(email).trim().toLowerCase();
  if (RM_DIRECTORY[key]) return RM_DIRECTORY[key];
  // fall back: derive a name from the local part
  var local = key.split('@')[0].replace(/[._]+/g, ' ').trim();
  return local.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

function fmtDate_(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, 'd MMM yyyy');
  }
  var s = String(v == null ? '' : v).trim();
  return s;
}

function firstUrl_(text) {
  if (!text) return '';
  var m = String(text).match(/https?:\/\/[^\s)]+/i);
  return m ? m[0] : '';
}

/* a single labelled field that may be a URL or just descriptive text */
function linkField_(text, icon) {
  text = String(text == null ? '' : text).trim();
  if (!text) return null;
  var url = firstUrl_(text);
  return { text: text, url: url, isUrl: !!url, icon: icon };
}

/* parse the multi-line "Platform links" cell into labelled links.
   Pulls EVERY url/bare-domain out of each line (so "youtube..., discord.gg/x"
   yields both), classifies each by domain, and de-duplicates.            */
function parseLinks_(text) {
  if (!text) return [];
  var lines = String(text).split(/\r?\n/);
  var out = [], seen = {}, currentLabel = '';

  function add(label, raw){
    var mk = makeLink_(label || '', raw);
    if (mk && mk.url){
      var k = mk.url.toLowerCase().replace(/\/+$/, '');
      if (!seen[k]){ seen[k] = 1; out.push(mk); }
    }
  }
  /* harvest all links from one chunk of text */
  function harvest(label, chunk){
    var urls = chunk.match(/https?:\/\/[^\s,;]+/gi) || [];
    urls.forEach(function(u){ add(label, u.replace(/[)\].,;]+$/, '')); });
    var bares = chunk.match(/(?:^|[\s,;(])((?:[\w-]+\.)+[a-z]{2,}(?:\/[^\s,;]*)?)/gi) || [];
    bares.forEach(function(b){
      var d = b.replace(/^[\s,;(]+/, '').replace(/[)\].,;]+$/, '');
      if (!d || /^https?:\/\//i.test(d)) return;            // already captured above
      add(label, d);
    });
  }

  lines.forEach(function (ln) {
    ln = ln.trim();
    if (!ln) return;
    var labelOnly = ln.match(/^([A-Za-z][A-Za-z .\/]{1,24}):\s*$/);
    if (labelOnly) { currentLabel = labelOnly[1].trim(); return; }

    var inline = ln.match(/^([A-Za-z][A-Za-z .\/]{1,24}):\s*(\S.*)$/);
    if (inline) { harvest(inline[1].trim(), inline[2].trim()); currentLabel=''; return; }

    harvest(currentLabel, ln);
  });
  return out;
}

function makeLink_(label, raw) {
  raw = raw.trim();
  if (!raw) return null;
  var url = raw;
  if (!/^https?:\/\//i.test(url)) {
    if (/^[\w.-]+\.[a-z]{2,}(\/|$)/i.test(url)) url = 'https://' + url; else return null;
  }
  var key = (label + ' ' + url).toLowerCase();
  var icon = '🌐', kind = label || 'Link';
  if (/youtu/.test(key))            { icon = '📺'; kind = 'YouTube'; }
  else if (/twitter|x\.com/.test(key)) { icon = '🐦'; kind = 'X / Twitter'; }
  else if (/instagram/.test(key))   { icon = '📸'; kind = 'Instagram'; }
  else if (/tiktok/.test(key))      { icon = '🎵'; kind = 'TikTok'; }
  else if (/discord/.test(key))     { icon = '💬'; kind = 'Discord'; }
  else if (/telegram|t\.me/.test(key)){ icon = '✈️'; kind = 'Telegram'; }
  else if (/facebook|fb\.com/.test(key)){ icon='👍'; kind='Facebook'; }
  else if (/linkedin/.test(key))    { icon = '💼'; kind = 'LinkedIn'; }
  else if (/website|\.com|\.co|\.io/.test(key)) { icon = '🌐'; kind = label || 'Website'; }
  return { label: kind, url: url, icon: icon };
}

/* extract a YouTube id from any watch / youtu.be / embed url */
function youTube_(text) {
  if (!text) return null;
  var s = String(text).trim();
  var m = s.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  if (!m) return null;
  var id = m[1];
  return {
    id: id,
    url: s,
    embed: 'https://www.youtube.com/embed/' + id,
    thumb: 'https://i.ytimg.com/vi/' + id + '/hqdefault.jpg'
  };
}

/* ================================================================== */
/*  CONTENT TRACKER  —  read one partner's linked tracker sheet LIVE  */
/*  (powers the "Content Tracker" tab in Index.html)                  */
/* ================================================================== */
function getPartnerContent(trackerUrl) {
  var res = readTrackerItems_(trackerUrl);
  if (!res.ok) return res;
  var items = res.items, byPlatform = {}, posted = 0;
  items.forEach(function(it){
    if (it.platform){ byPlatform[it.platform] = (byPlatform[it.platform]||0)+1; }
    if (/post|publish|live|done|ok/i.test(it.status)) posted++;
  });
  return {
    ok:true, sheetName: res.sheetName, total: items.length, posted: posted,
    byPlatform: byPlatform, items: items,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy, h:mm a')
  };
}

/* shared low-level reader: returns every dated/linked row, newest first */
function readTrackerItems_(trackerUrl) {
  if (!trackerUrl) return { ok:false, error:'This partner has no tracker URL in the sheet.' };
  var ss;
  try { ss = SpreadsheetApp.openByUrl(trackerUrl); }
  catch (e) {
    return { ok:false, error:'Cannot open this tracker. Make sure the running account has access. ('+e.message+')' };
  }
  var sh = pickContentSheet_(ss);
  if (!sh) return { ok:false, error:'No readable content sheet found in that tracker.' };

  var rng  = sh.getDataRange();
  var vals = rng.getValues();
  var rich = rng.getRichTextValues();

  var hr = -1, H = [];
  for (var r=0; r<Math.min(vals.length, 20); r++){
    var rowNorm = vals[r].map(norm_);
    var hasDate = rowNorm.some(function(x){ return x==='date' || x.indexOf('date')===0; });
    var hasKey  = rowNorm.some(function(x){
      return x.indexOf('status')>-1 || x.indexOf('platform')>-1 ||
             x.indexOf('content type')>-1 || x.indexOf('location')>-1 || x.indexOf('url')>-1; });
    if (hasDate && hasKey){ hr=r; H=rowNorm; break; }
  }
  if (hr === -1){ hr = 0; H = vals[0].map(norm_); }

  function find(){
    for (var a=0;a<arguments.length;a++){
      var kw=arguments[a];
      for (var j=0;j<H.length;j++){ if (H[j] && H[j].indexOf(kw)>-1) return j; }
    } return -1;
  }
  var col = {
    date:    find('date'),
    type:    find('content type','type'),
    platform:find('platform'),
    topic:   find('topic','title','description'),
    status:  find('status'),
    url:     find('location','url','screenshot','link'),
    qc:      find('qc','quality'),
    remarks: find('remark','pm remark','manager')
  };

  var tz = Session.getScriptTimeZone();
  var items = [];
  for (var r2=hr+1; r2<vals.length; r2++){
    var row = vals[r2];
    var url = '';
    if (col.url>-1 && rich[r2] && rich[r2][col.url]){
      var rt = rich[r2][col.url];
      url = rt.getLinkUrl() || '';
      if (!url){ rt.getRuns().forEach(function(run){ if(!url) url = run.getLinkUrl()||''; }); }
    }
    if (!url && col.url>-1) url = firstUrl_(row[col.url]);

    var d = parseDate_(col.date>-1 ? row[col.date] : '');
    if (!d && !url) continue;

    var topic  = col.topic>-1 ? String(row[col.topic]||'').trim() : '';
    var type   = col.type>-1  ? String(row[col.type]||'').trim()  : '';
    var plat   = col.platform>-1 ? String(row[col.platform]||'').trim() : '';
    var status = col.status>-1 ? String(row[col.status]||'').trim() : '';
    if (!url && !type && !plat && !status && !topic) continue;

    items.push({
      dateMs:  d ? d.getTime() : 0,
      date:    d ? Utilities.formatDate(d, tz, 'd MMM yyyy') : '',
      month:   d ? Utilities.formatDate(d, tz, 'MMMM yyyy') : 'Undated',
      type:    type, platform: plat, topic: topic, status: status,
      qc:      col.qc>-1 ? String(row[col.qc]||'').trim() : '',
      remarks: col.remarks>-1 ? String(row[col.remarks]||'').trim() : '',
      url:     url
    });
  }
  items.sort(function(a,b){ return b.dateMs - a.dateMs; });   /* newest first */
  return { ok:true, sheetName: sh.getName(), items: items };
}

/* the most-recent YouTube video in a tracker (items already newest-first) */
function latestVideo_(trackerUrl){
  var res = readTrackerItems_(trackerUrl);
  if (!res.ok) return null;
  for (var i=0;i<res.items.length;i++){
    var yt = youTube_(res.items[i].url);
    if (yt) return { id: yt.id, url: res.items[i].url, date: res.items[i].date, topic: res.items[i].topic, platform: res.items[i].platform };
  }
  return null;
}

/* ================================================================== */
/*  SPOTLIGHT  —  latest video per partner, pulled from their tracker */
/*  Cached 30 min so the dashboard stays fast. Pass true to force.    */
/* ================================================================== */
function getSpotlightVideos(force){
  var cache = CacheService.getScriptCache();
  var KEY = 'spotlight_v3';
  if (!force){ try { var c = cache.get(KEY); if (c) return JSON.parse(c); } catch(e){} }

  var book = getBook_();
  var partners = readFutures_(book, readRmAverages_(book));
  var vids = [];
  for (var i=0;i<partners.length;i++){
    var p = partners[i];
    if (!p.trackerUrl) continue;
    var v = latestVideo_(p.trackerUrl);
    if (!v) continue;
    vids.push({
      name:p.name, username:p.username, country:p.country,
      status:p.status, statusKey:p.statusKey,
      id:v.id,
      thumb:'https://i.ytimg.com/vi/'+v.id+'/hqdefault.jpg',
      embed:'https://www.youtube.com/embed/'+v.id,
      url:v.url, date:v.date, topic:v.topic
    });
  }
  try { cache.put(KEY, JSON.stringify(vids), 1800); } catch(e){}
  return vids;
}

function pickContentSheet_(ss){
  var sheets = ss.getSheets();
  for (var s=0; s<sheets.length; s++){
    var sh = sheets[s];
    var top = sh.getRange(1,1,Math.min(20,sh.getLastRow()||1),Math.min(20,sh.getLastColumn()||1)).getValues();
    for (var r=0;r<top.length;r++){
      var rn = top[r].map(norm_);
      var hasDate = rn.some(function(x){return x.indexOf('date')>-1;});
      var hasKey  = rn.some(function(x){return x.indexOf('platform')>-1||x.indexOf('status')>-1||x.indexOf('content type')>-1||x.indexOf('location')>-1;});
      if (hasDate && hasKey) return sh;
    }
  }
  return sheets[0] || null;
}

function parseDate_(v){
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v==null?'':v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m){
    var mo=+m[1], da=+m[2], yr=+m[3]; if (yr<100) yr+=2000;
    if (mo>=1&&mo<=12&&da>=1&&da<=31){ var d=new Date(yr,mo-1,da); if(!isNaN(d.getTime())) return d; }
  }
  var t = Date.parse(s);
  return isNaN(t) ? null : new Date(t);
}

/* ================================================================== */
/*  RM ART TIMELINE  —  daily response times for FUTURES RMs only     */
/*  Reads the RM_ART tab (date columns across the top, RMs in rows    */
/*  grouped by team in column A). Returns the per-day series so the   */
/*  front-end can average over Today / Last 7 / Last 30 / All.        */
/* ================================================================== */
function getRmArtTimeline(){
  var sh = getBook_().getSheetByName(RM_ART_SHEET);
  if (!sh) return { ok:false, error:'RM_ART tab not found.' };

  var lastRow = Math.min(sh.getLastRow(), 40);
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var tz = Session.getScriptTimeZone();

  /* row 0 = "Teams | Date | <date> <date> ..."  → collect date columns */
  var dateRow = vals[0] || [];
  var dates = [], dateCols = [];
  for (var c = 2; c < dateRow.length; c++){
    var d = dateRow[c];
    if (d instanceof Date && !isNaN(d.getTime())){
      dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
      dateCols.push(c);
    }
  }

  /* walk rows; column A carries the team label for the group that follows.
     Collect individual RM rows whose group is the Futures team.          */
  var rms = [], team = '';
  for (var r = 2; r < vals.length; r++){
    var a = String(vals[r][0] == null ? '' : vals[r][0]).trim();
    var b = String(vals[r][1] == null ? '' : vals[r][1]).trim();
    if (a) team = a;
    if (/^date$/i.test(b)) break;          /* reached the team-summary block → stop */
    if (!b) continue;
    if (!/futures/i.test(team)) continue;  /* Futures RMs only — skip CFD / OS / Onboarding */

    var series = dateCols.map(function(ci){ return artToSeconds_(vals[r][ci]); });
    rms.push({ name: b, seconds: series });
  }

  /* trim trailing dates where NO Futures RM has any activity (drops the
     empty future months so the timeline reflects real data only)        */
  var lastActive = -1;
  for (var d2 = 0; d2 < dates.length; d2++){
    var any = rms.some(function(rm){ return rm.seconds[d2] != null; });
    if (any) lastActive = d2;
  }
  if (lastActive >= 0 && lastActive < dates.length - 1){
    dates = dates.slice(0, lastActive + 1);
    rms.forEach(function(rm){ rm.seconds = rm.seconds.slice(0, lastActive + 1); });
  }

  return {
    ok: true,
    team: 'Futures RM',
    dates: dates,
    rms: rms,
    teamAvg: (function(){ var x=readRmAverages_(getBook_()); return x.byName['Futures RM'] || '—'; })(),
    generatedAt: Utilities.formatDate(new Date(), tz, 'd MMM yyyy, h:mm a')
  };
}

/* "2h 17m" / "7m 2s" / "25s" / "0s" -> seconds.  "—" or blank -> null */
function artToSeconds_(v){
  var s = String(v == null ? '' : v).trim();
  if (!s || s === '—' || s === '-') return null;
  var sec = 0, m, matched = false;
  m = s.match(/(\d+)\s*h/);      if (m){ sec += (+m[1])*3600; matched = true; }
  m = s.match(/(\d+)\s*m(?!s)/); if (m){ sec += (+m[1])*60;   matched = true; }
  m = s.match(/(\d+)\s*s/);      if (m){ sec += (+m[1]);      matched = true; }
  return matched ? sec : null;
}