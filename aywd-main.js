(function () {
    'use strict';

    const REQUIRED_HASH = '#AYWD';
    if (String(location.hash || '').toUpperCase() !== REQUIRED_HASH) {
        console.log('[AYWD] not started. Hash:', location.hash);
        return;
    }

    const APP_ID = 'aywd_overlay';
    const STORAGE_KEY = 'AYWD_V11_12_STATE';
    const PRODUCTIVITY_STORAGE_KEY = 'AYWD_PRODUCTIVITY_V11_PICK_ONLY';
    const SHORTS_STORAGE_KEY = 'AYWD_SHORTS_WORKERS_V11';
    const SHAREPOINT_ORIGIN = 'https://cevalogisticsoffice365.sharepoint.com';
    const SHAREPOINT_FILES_API = SHAREPOINT_ORIGIN + "/sites/Europe-AboutYou-WRO3AYWhiteCollars/_api/web/GetFolderByServerRelativeUrl('/sites/Europe-AboutYou-WRO3AYWhiteCollars/Shared%20Documents/General/Back%20D')/Files?$select=Name,TimeLastModified,ServerRelativeUrl";
    const PRODUCTIVITY_FILE_KEYWORDS = ['productivity', 'bonus', 'weekly', 'monthly'];
    const GRAFANA_QUERY_URL = '/api/ds/query';
    const DATASOURCE_UID = 'mFpJIAhVk';
    const DATASOURCE_ID = 108;
    const AUTO_REFRESH_MS = 30000;
    const LOW_UPH_LIMIT = 90;

    // UI order stays unchanged. Assignment order is controlled separately.
    const ASSIGNMENT_PRIORITY = ['ICON', 'SCON', 'BCON', 'RCON', 'TCON', 'PCON', 'CCON', 'THAN'];
    const FLOOR_SINGLE_LIST_GUARD_PREFIXES = ['SCON', 'BCON', 'RCON'];
    const SHORT_PREFIXES = ['PCON', 'CCON'];
    const SHORTS_TO_MULTI_MIN_ACTIVE = 3; // more than 2 SHORTS already issued on PCON/CCON
    const SHORTS_MULTI_PREFIX = 'TCON';
    const SHORTS_MULTI_OSR_LIMIT = 2;
    const HIGH_RISK_ZONE_RAWS = ['50P5L1A'];

    let autoRefreshTimer = null;
    let isWmsLoading = false;

    let STREAMS = [
        { p: 'ICON', s: 'Item cross-dock', n: 20 },
        { p: 'SCON', s: 'Single', n: 6 },
        { p: 'TCON', s: 'Multies', n: 40 },
        { p: 'BCON', s: 'B2B', n: 0 },
        { p: 'PCON', s: 'Short list Multies', n: 0 },
        { p: 'CCON', s: 'Short list ICD', n: 0 },
        { p: 'RCON', s: 'Pre. Relo.', n: 0 },
        { p: 'THAN', s: 'Oversizes', n: 0 }
    ];

    const PREFIX_COLORS = {
        TCON: '#B10252', THAN: '#46745d', BCON: '#0ea9bb', SCON: 'green',
        PCON: '#E033FF', CCON: '#E033FF', RCON: '#5F5D9C', ICON: '#4169e1'
    };

    let state = {
        raw: [],
        picklists: [],
        zones: {},
        assignments: [],
        productivity: {},
        productivityLoadedAt: null,
        productivitySourceName: '',
        productivitySourceModified: null,
        shortWorkers: [],
        showUPH: false,
        login: '',
        start: null,
        end: null,
        loadedAt: null,
        autoRefreshEnabled: false
    };

    init();

    function init() {
        loadPersistedState();
        loadProductivityData();
        loadShortWorkers();
        addStyles();
        buildOverlay();
        bindEvents();
        renderStreams();
        renderAssignments(false);
        renderSummary();
        renderProductivityStatus();
        renderShortsStatus();
        updateStats();
        setTimeout(() => loadProductivityFromSharePoint({ auto: true }), 800);
        if (state.autoRefreshEnabled) startAutoRefresh(false);
        console.log('[AYWD] V11.13 remote main initialized');
    }

    // ============================================================
    // Utility
    // ============================================================
    function prefixColor(prefix) { return PREFIX_COLORS[prefix] || '#38bdf8'; }
    function normalizeLogin(value) { return String(value || '').trim().toUpperCase(); }
    function isSinglePrefix(prefix) { return String(prefix || '').toUpperCase() === 'SCON'; }
    function isMultiPrefix(prefix) { const p = String(prefix || '').toUpperCase(); return p === 'TCON' || p === 'PCON'; }
    function isLowUPH(worker) { return !worker || worker.uph === null || worker.uph === undefined || Number(worker.uph) < LOW_UPH_LIMIT; }
    function escapeHtml(value) { return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;'); }
    function makeId() { return 'T' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7).toUpperCase(); }
    function prefixDone(prefix) { return state.assignments.filter(x => x.prefix === prefix).length; }

    function getWorkerPerf(login) {
        const id = normalizeLogin(login);
        return state.productivity && state.productivity[id]
            ? state.productivity[id]
            : { login: id, uph: null, singleUPH: null, multiUPH: null, icdUPH: null, productivity: null, source: 'NO_PICK_ROWS' };
    }

    function isShortWorker(login) {
        return Array.isArray(state.shortWorkers) && state.shortWorkers.includes(normalizeLogin(login));
    }

    function osrOrder(value) {
        const match = String(value || '').match(/OSR(\d+)/i);
        return match ? Number(match[1]) : 999999999;
    }

    function getOsrLabel(value) {
        const raw = String(value || '').trim();
        if (!raw) return 'NO OSR';
        const match = raw.match(/OSR\s*0*([0-9]+)/i);
        return match ? 'OSR' + match[1] : raw;
    }

    // ============================================================
    // OSR Summary
    // ============================================================
    function buildIssuedCountByPrefixOsr() {
        const listToOsr = {};
        state.picklists.forEach(pl => { if (pl && pl.list) listToOsr[String(pl.list)] = getOsrLabel(pl.shipmentGroup); });
        const issued = {};
        state.assignments.forEach(a => {
            const key = String(a.prefix || '') + '|' + (listToOsr[String(a.list || '')] || 'NO OSR');
            issued[key] = (issued[key] || 0) + 1;
        });
        return issued;
    }

    function buildOsrSummaryForPrefix(prefix) {
        const groups = {};
        const issued = buildIssuedCountByPrefixOsr();
        state.picklists.filter(pl => pl && pl.prefix === prefix).forEach(pl => {
            const osr = getOsrLabel(pl.shipmentGroup);
            if (!groups[osr]) {
                groups[osr] = { osr, lists: 0, releasedLists: 0, inProgressLists: 0, units: 0, prio: 0, issued: 0, osrSort: osrOrder(pl.shipmentGroup) };
            }
            const g = groups[osr];
            const status = String(pl.status || '').toLowerCase();
            g.lists += 1;
            g.units += Number(pl.units || 0);
            g.prio = Math.max(g.prio, Number(pl.prio || 0));
            g.osrSort = Math.min(g.osrSort, osrOrder(pl.shipmentGroup));
            if (status === 'released') g.releasedLists += 1;
            if (status === 'in progress') g.inProgressLists += 1;
        });
        Object.keys(groups).forEach(osr => { groups[osr].issued = issued[prefix + '|' + osr] || 0; });
        return Object.values(groups).sort((a, b) =>
            (Number(b.prio || 0) - Number(a.prio || 0)) ||
            (Number(a.osrSort || 999999999) - Number(b.osrSort || 999999999)) ||
            String(a.osr).localeCompare(String(b.osr))
        );
    }

    // ============================================================
    // Assignment list helpers
    // ============================================================
    function assignmentKey(prefix, zoneRaw, list) { return String(prefix || '') + '|' + String(zoneRaw || '') + '|' + String(list || ''); }
    function hasIssuedSameZoneList(prefix, zoneRaw, list) { const key = assignmentKey(prefix, zoneRaw, list); return state.assignments.some(a => assignmentKey(a.prefix, a.zoneRaw, a.list) === key); }
    function isListAlreadyIssued(listId) { return state.assignments.some(a => String(a.list || '') === String(listId || '')); }

    function getUnissuedReleasedLists(zone) {
        if (!zone || !Array.isArray(zone.lists)) return [];
        return zone.lists.filter(pl => pl && pl.list && String(pl.status || '').toLowerCase() === 'released' && !isListAlreadyIssued(pl.list));
    }

    function zoneHasAnyUnissuedReleasedList(zone) { return getUnissuedReleasedLists(zone).length > 0; }

    function getAllAssignableLists(prefix) {
        const output = [];
        Object.values(state.zones[prefix] || {}).forEach(zone => getUnissuedReleasedLists(zone).forEach(list => output.push({ zone, list })));
        return output;
    }

    function sortGlobalListCandidates(items) {
        const floorCounts = {};
        state.assignments.forEach(a => { floorCounts[a.floor] = (floorCounts[a.floor] || 0) + 1; });
        items.sort((a, b) => {
            const floorA = floorCounts[a.zone.floor] || 0;
            const floorB = floorCounts[b.zone.floor] || 0;
            return (Number(b.list.prio || 0) - Number(a.list.prio || 0)) ||
                (osrOrder(a.list.shipmentGroup) - osrOrder(b.list.shipmentGroup)) ||
                (a.zone.count - b.zone.count) ||
                (floorA - floorB) ||
                (Number(b.list.units || 0) - Number(a.list.units || 0)) ||
                String(a.list.list).localeCompare(String(b.list.list));
        });
        return items;
    }

    function zoneHasImportantOpened(zone, importantPrio) {
        if (!zone || !Array.isArray(zone.lists)) return false;
        return zone.lists.some(pl => pl && pl.list && Number(pl.prio || 0) === Number(importantPrio || 0) && (String(pl.status || '').toLowerCase() === 'in progress' || hasIssuedSameZoneList(pl.prefix, pl.zoneRaw, pl.list)));
    }

    function getImportantPrioForPrefix(prefix) {
        const all = getAllAssignableLists(prefix);
        return all.length ? Math.max(...all.map(x => Number(x.list.prio || 0))) : null;
    }

    function allRemainingMultiesListsAreImportant(prefix, importantPrio) {
        const all = getAllAssignableLists(prefix);
        return all.length > 0 && all.every(x => Number(x.list.prio || 0) === Number(importantPrio || 0));
    }

    function remainingHCForPrefix(prefix) {
        const stream = STREAMS.find(x => x.p === prefix);
        if (!stream) return 0;
        return Math.max(0, Number(stream.n || 0) - prefixDone(prefix));
    }

    function openZoneCount(prefix) {
        return Object.values(state.zones[prefix] || {}).filter(z => zoneHasAnyUnissuedReleasedList(z)).length;
    }

    function shouldLowUPHBeAllowedToImportant(prefix, importantPrio) {
        if (allRemainingMultiesListsAreImportant(prefix, importantPrio)) return true;
        return remainingHCForPrefix(prefix) >= openZoneCount(prefix);
    }

    // ============================================================
    // New guards and High Risk SHORTS
    // ============================================================
    function isHighRiskZone(zone) {
        if (!zone) return false;
        return HIGH_RISK_ZONE_RAWS.includes(String(zone.zoneRaw || '').toUpperCase());
    }

    function activeShortsOnShortListsCount() {
        const shortSet = new Set((state.shortWorkers || []).map(normalizeLogin));
        return state.assignments.filter(a =>
            a &&
            shortSet.has(normalizeLogin(a.login)) &&
            (a.prefix === 'PCON' || a.prefix === 'CCON') &&
            String(a.wmsStatus || '').toUpperCase() !== 'NOT_FOUND'
        ).length;
    }

    function topReleasedOsrLabelsForPrefix(prefix, limit) {
        const groups = {};
        getAllAssignableLists(prefix).forEach(x => {
            const osr = getOsrLabel(x.list.shipmentGroup);
            if (!groups[osr]) groups[osr] = { osr, prio: 0, osrSort: osrOrder(x.list.shipmentGroup), lists: 0 };
            groups[osr].prio = Math.max(groups[osr].prio, Number(x.list.prio || 0));
            groups[osr].osrSort = Math.min(groups[osr].osrSort, osrOrder(x.list.shipmentGroup));
            groups[osr].lists += 1;
        });
        return Object.values(groups)
            .sort((a, b) =>
                (Number(b.prio || 0) - Number(a.prio || 0)) ||
                (Number(a.osrSort || 999999999) - Number(b.osrSort || 999999999)) ||
                String(a.osr).localeCompare(String(b.osr))
            )
            .slice(0, limit)
            .map(x => x.osr);
    }

    function getHighRiskShortsMultiCandidates() {
        const allowedOsrs = new Set(topReleasedOsrLabelsForPrefix(SHORTS_MULTI_PREFIX, SHORTS_MULTI_OSR_LIMIT));
        return getAllAssignableLists(SHORTS_MULTI_PREFIX).filter(x => isHighRiskZone(x.zone) && allowedOsrs.has(getOsrLabel(x.list.shipmentGroup)));
    }

    function isShortsHighRiskMode(worker, prefix) {
        return worker && isShortWorker(worker.login) && prefix === SHORTS_MULTI_PREFIX && activeShortsOnShortListsCount() >= SHORTS_TO_MULTI_MIN_ACTIVE;
    }

    function getShortsPreferredPrefix(worker) {
        if (!worker || !isShortWorker(worker.login)) return null;
        for (const prefix of SHORT_PREFIXES) if (getAllAssignableLists(prefix).length > 0) return prefix;
        return null;
    }

    function applySingleListFloorGuard(prefix, candidates) {
        if (!FLOOR_SINGLE_LIST_GUARD_PREFIXES.includes(prefix)) return candidates;
        if (!Array.isArray(candidates) || candidates.length <= 1) return candidates;

        const floorCounts = {};
        candidates.forEach(x => { floorCounts[String(x.zone.floor || '')] = (floorCounts[String(x.zone.floor || '')] || 0) + 1; });
        const hasMultiListFloor = Object.values(floorCounts).some(count => count > 1);
        if (!hasMultiListFloor) return candidates;

        const filtered = candidates.filter(x => {
            const floor = String(x.zone.floor || '');
            if ((floorCounts[floor] || 0) > 1) return true;
            const other = candidates.filter(y => String(y.zone.floor || '') !== floor);
            const maxOtherPrio = other.length ? Math.max(...other.map(y => Number(y.list.prio || 0))) : -1;
            return Number(x.list.prio || 0) > maxOtherPrio;
        });
        return filtered.length ? filtered : candidates;
    }

    function getBestGlobalAssignment(prefix, worker) {
        let all = getAllAssignableLists(prefix);
        if (!all.length) return null;

        if (isShortsHighRiskMode(worker, prefix)) {
            const highRisk = getHighRiskShortsMultiCandidates();
            if (highRisk.length) return sortGlobalListCandidates(highRisk)[0] || null;
        }

        all = applySingleListFloorGuard(prefix, all);
        const low = isLowUPH(worker);
        const importantPrio = Math.max(...all.map(x => Number(x.list.prio || 0)));

        if (!low || !isMultiPrefix(prefix)) return sortGlobalListCandidates(all)[0] || null;

        const safe = all.filter(x => Number(x.list.prio || 0) < Number(importantPrio || 0) || zoneHasImportantOpened(x.zone, importantPrio));
        if (safe.length) return sortGlobalListCandidates(safe)[0] || null;

        if (shouldLowUPHBeAllowedToImportant(prefix, importantPrio)) {
            const important = all.filter(x => Number(x.list.prio || 0) === Number(importantPrio || 0));
            important.sort((a, b) =>
                (Number(a.list.units || 0) - Number(b.list.units || 0)) ||
                (osrOrder(a.list.shipmentGroup) - osrOrder(b.list.shipmentGroup)) ||
                String(a.list.list).localeCompare(String(b.list.list))
            );
            return important[0] || null;
        }
        return sortGlobalListCandidates(all)[0] || null;
    }

    // ============================================================
    // UI
    // ============================================================
    function addStyles() {
        const old = document.getElementById(APP_ID + '_style');
        if (old) old.remove();
        const style = document.createElement('style');
        style.id = APP_ID + '_style';
        style.textContent = `
#${APP_ID}{position:fixed;inset:0;z-index:2147483647;background:radial-gradient(circle at top left,rgba(37,99,235,.25),transparent 35%),linear-gradient(135deg,#07111f 0%,#0b1020 55%,#111827 100%);color:#e5e7eb;font-family:Arial,sans-serif;display:flex;box-sizing:border-box}
#${APP_ID} *{box-sizing:border-box}
#${APP_ID} .ay-left{width:520px;padding:14px;background:rgba(15,23,42,.94);border-right:1px solid rgba(148,163,184,.25);overflow-y:auto}
#${APP_ID} .ay-title{font-size:23px;font-weight:900;color:white;margin-bottom:2px;cursor:help}
#${APP_ID} .ay-email{font-size:12px;color:#38bdf8;margin-bottom:6px;font-weight:800}
#${APP_ID} .ay-sub{color:#94a3b8;font-size:12px;margin-bottom:10px}
#${APP_ID} .ay-buttons{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}
#${APP_ID} button{border:0;border-radius:8px;padding:8px 12px;font-weight:900;cursor:pointer;color:white;background:#334155;box-shadow:0 2px 0 rgba(0,0,0,.4)}
#${APP_ID} button:hover{filter:brightness(1.12)} #${APP_ID} button:disabled{opacity:.55;cursor:wait}
#${APP_ID} .btn-load{background:#2563eb} #${APP_ID} .btn-copy{background:#16a34a} #${APP_ID} .btn-undo{background:#f59e0b} #${APP_ID} .btn-clear{background:#dc2626} #${APP_ID} .btn-close{background:#475569;margin-left:auto} #${APP_ID} .btn-prod{background:#7c3aed} #${APP_ID} .btn-uph{background:#0ea5e9} #${APP_ID} .btn-shorts{background:#059669}
#${APP_ID} .section{margin:14px 0 7px;font-size:17px;font-weight:900;color:#f8fafc;border-left:4px solid #38bdf8;padding-left:8px}
#${APP_ID} table{width:100%;border-collapse:collapse;font-size:13px;border-radius:10px;overflow:hidden}
#${APP_ID} th{background:#1f4e79;color:white;padding:7px;text-align:left;border:1px solid #334155}
#${APP_ID} td{padding:6px;border:1px solid #334155;background:rgba(15,23,42,.9)} #${APP_ID} tr:nth-child(even) td{background:rgba(30,41,59,.9)}
#${APP_ID} .need-input{width:60px;background:#020617;color:white;border:1px solid #475569;border-radius:6px;padding:5px;font-weight:900;text-align:center}
#${APP_ID} .scan-input{width:100%;height:44px;background:#020617;color:white;border:1px solid #475569;border-radius:10px;font-size:19px;padding:0 12px;margin-bottom:8px;outline:none;text-transform:uppercase}
#${APP_ID} .stats,#${APP_ID} .summary{margin-top:8px;padding:10px;border-radius:10px;background:rgba(2,6,23,.65);border:1px solid rgba(148,163,184,.25);color:#cbd5e1;font-size:13px;line-height:1.55}
#${APP_ID} .ay-grid{flex:1;overflow:auto;padding:12px;display:flex;flex-direction:column;flex-wrap:wrap;gap:8px;align-content:flex-start;max-height:100vh;scroll-behavior:smooth}
#${APP_ID} .card{width:265px;min-height:86px;border-radius:12px;padding:9px 10px;background:rgba(15,23,42,.94);border:1px solid rgba(148,163,184,.25);border-left:7px solid #38bdf8;box-shadow:0 6px 18px rgba(0,0,0,.28);cursor:pointer;user-select:none}
#${APP_ID} .card:hover{filter:brightness(1.12)} #${APP_ID} .card.duplicate{background:rgba(127,29,29,.75);border-color:#ef4444}
#${APP_ID} .card.in-progress-gray{background:rgba(51,65,85,.72);border-color:#64748b;color:#cbd5e1;filter:grayscale(.55)}
#${APP_ID} .login{font-size:22px;font-weight:900;color:white;margin-bottom:4px} #${APP_ID} .zone{font-size:18px;font-weight:900;color:#facc15}
#${APP_ID} .list{display:inline-block;margin-top:5px;font-size:14px;font-weight:900;letter-spacing:.2px}.list-first{font-size:22px;line-height:12px;font-weight:1000;margin-right:1px}
#${APP_ID} .status-badge{display:inline-block;margin-top:5px;margin-right:4px;padding:2px 6px;border-radius:999px;font-size:10px;font-weight:900;background:#475569;color:#e5e7eb}
#${APP_ID} .status-badge.in-progress{background:#64748b} #${APP_ID} .status-badge.not-found{background:#dc2626} #${APP_ID} .status-badge.low-uph{background:#f97316} #${APP_ID} .status-badge.good-uph{background:#0ea5e9}
#${APP_ID} .uph-hidden .uph-badge{display:none!important}
#${APP_ID} .pill{display:inline-block;padding:2px 7px;border-radius:999px;background:#334155;color:#e5e7eb;font-size:11px;font-weight:900}
#${APP_ID} .ok{color:#22c55e;font-weight:900} #${APP_ID} .warn{color:#facc15;font-weight:900} #${APP_ID} .full{color:#ef4444;font-weight:900} #${APP_ID} .small{font-size:12px;color:#94a3b8}
#${APP_ID} .file-row{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:6px}.hidden-file{display:none}
#${APP_ID}_toast{position:fixed;top:18px;right:18px;z-index:2147483647;background:#16a34a;color:white;padding:10px 14px;border-radius:10px;font-weight:900;display:none;max-width:540px;white-space:pre-line}`;
        document.head.appendChild(style);
    }

    function buildOverlay() {
        const old = document.getElementById(APP_ID); if (old) old.remove();
        const oldToast = document.getElementById(APP_ID + '_toast'); if (oldToast) oldToast.remove();
        const root = document.createElement('div');
        root.id = APP_ID;
        root.innerHTML = `
<div class="ay-left">
 <div class="ay-title" title="Designed by Kateryna Androshchuk">AY Work Distributor</div>
 <div class="ay-email">yevhenii.karpenko@cevalogistics.com</div>
 <div class="ay-sub">V11.13 remote · SharePoint auto productivity · High Risk 50P5L1A</div>
 <div class="ay-buttons"><button class="btn-load" id="aywdLoad">LOAD WMS</button><button class="btn-copy" id="aywdCopy">COPY</button><button class="btn-undo" id="aywdUndo">UNDO</button><button class="btn-clear" id="aywdClear">CLEAR</button><button class="btn-close" id="aywdClose">CLOSE</button></div>
 <div class="small">Last load: <span id="aywdLoadedAt">-</span></div>
 <div class="small">Auto refresh: <span id="aywdAutoStatus">OFF</span></div>
 <div class="small">WMS refresh state: <span id="aywdRefreshState">idle</span></div>
 <div class="section">Streams</div><table id="aywdStreams"></table>
 <div class="section">Scan</div><input class="scan-input" id="aywdLogin" placeholder="LOGIN" autocomplete="off"><input class="scan-input" id="aywdScanner" placeholder="SCANNER" autocomplete="off">
 <div class="stats">Count: <span class="ok" id="aywdCount">0</span><br>Current stream: <span id="aywdCurrentStream">-</span><br>Start: <span id="aywdStart">-</span><br>End: <span id="aywdEnd">-</span><br>Duration: <span id="aywdDuration">-</span></div>
 <div class="section">WMS Summary</div><div class="summary" id="aywdSummary">No data loaded.</div>
 <div class="section">Productivity file</div>
 <div class="file-row"><button class="btn-prod" id="aywdProductivityBtn">AUTO PRODUCTIVITY</button><button class="btn-uph" id="aywdToggleUPH">SHOW UPH</button><button class="btn-shorts" id="aywdShortsBtn">SHORTS</button><input class="hidden-file" id="aywdProductivityFile" type="file" accept=".xlsx,.xlsm,.xls"></div>
 <div class="small">Productivity: <span id="aywdProductivityStatus">not loaded</span></div>
 <div class="small">SHORTS: <span id="aywdShortsStatus">0 workers</span></div>
 <div class="small">Rule: productivity auto-loads from SharePoint. Pick-only UPH. SHORTS overflow to TCON High Risk after 2 active PCON/CCON tokens.</div>
</div><div class="ay-grid" id="aywdGrid"></div>`;
        document.body.appendChild(root);
        const toast = document.createElement('div'); toast.id = APP_ID + '_toast'; document.body.appendChild(toast);
        applyUPHVisibility();
    }

    function bindEvents() {
        document.getElementById('aywdLoad').addEventListener('click', () => loadWms({ manual: true }));
        document.getElementById('aywdCopy').addEventListener('click', copyData);
        document.getElementById('aywdUndo').addEventListener('click', undo);
        document.getElementById('aywdClear').addEventListener('click', clearAll);
        document.getElementById('aywdClose').addEventListener('click', closeOverlay);
        document.getElementById('aywdToggleUPH').addEventListener('click', toggleUPHVisibility);
        document.getElementById('aywdProductivityBtn').addEventListener('click', () => loadProductivityFromSharePoint({ auto: false }));
        document.getElementById('aywdShortsBtn').addEventListener('click', editShortWorkers);
        document.getElementById('aywdProductivityFile').addEventListener('change', e => handleProductivityUpload(e.target.files && e.target.files[0]));

        const loginInput = document.getElementById('aywdLogin');
        const scannerInput = document.getElementById('aywdScanner');
        loginInput.addEventListener('keydown', e => {
            if (e.key === 'Enter' || e.key === 'Tab') {
                e.preventDefault();
                const value = normalizeLogin(loginInput.value);
                if (!value) return;
                state.login = value;
                if (!state.start) state.start = new Date().toISOString();
                loginInput.value = '';
                scannerInput.focus();
                updateStats();
                savePersistedState();
            }
        });
        scannerInput.addEventListener('keydown', e => {
            if (e.key === 'Enter' || e.key === 'Tab') {
                e.preventDefault();
                const scanner = scannerInput.value.trim().toUpperCase();
                if (state.login && scanner) assign(state.login, scanner);
                state.login = '';
                scannerInput.value = '';
                loginInput.focus();
            }
        });
        loginInput.focus();
    }

    function toggleUPHVisibility() { state.showUPH = !state.showUPH; savePersistedState(); applyUPHVisibility(); renderAssignments(false); }
    function applyUPHVisibility() { const root = document.getElementById(APP_ID); if (root) root.classList.toggle('uph-hidden', !state.showUPH); const btn = document.getElementById('aywdToggleUPH'); if (btn) btn.textContent = state.showUPH ? 'HIDE UPH' : 'SHOW UPH'; }

    // ============================================================
    // Productivity Pick-only UPH
    // ============================================================
    async function handleProductivityUpload(file) {
        if (!file) return;
        try {
            const buffer = await file.arrayBuffer();
            await loadProductivityFromArrayBuffer(buffer, file.name || 'manual file', null);
        } catch (err) {
            console.error('[AYWD] manual productivity upload failed', err);
            alert('PRODUCTIVITY LOAD ERROR:\n' + err.message);
        }
    }

    function pickBestSharePointProductivityFile(files) {
        const candidates = (files || []).filter(file => {
            const name = String(file.Name || '').toLowerCase();
            const hasKeyword = PRODUCTIVITY_FILE_KEYWORDS.some(keyword => name.includes(keyword));
            const isExcel = /\.(xlsx|xlsm|xls)$/i.test(String(file.Name || ''));
            return hasKeyword && isExcel;
        });

        candidates.sort((a, b) => new Date(b.TimeLastModified || 0) - new Date(a.TimeLastModified || 0));
        return candidates[0] || null;
    }

    function gmRequest(options) {
        return new Promise((resolve, reject) => {
            if (typeof GM_xmlhttpRequest === 'function') {
                GM_xmlhttpRequest({
                    method: options.method || 'GET',
                    url: options.url,
                    headers: options.headers || {},
                    responseType: options.responseType || 'text',
                    withCredentials: true,
                    timeout: options.timeout || 30000,
                    onload: response => resolve(response),
                    onerror: error => reject(new Error('GM request failed')),
                    ontimeout: () => reject(new Error('SharePoint request timeout'))
                });
                return;
            }

            fetch(options.url, {
                method: options.method || 'GET',
                headers: options.headers || {},
                credentials: 'include'
            }).then(async response => {
                if (!response.ok) throw new Error('HTTP ' + response.status);
                if (options.responseType === 'arraybuffer') {
                    resolve({ status: response.status, response: await response.arrayBuffer() });
                } else {
                    resolve({ status: response.status, responseText: await response.text() });
                }
            }).catch(reject);
        });
    }

    function parseSharePointFilesResponse(response) {
        const payload = typeof response.response === 'object' && response.response !== null
            ? response.response
            : JSON.parse(response.responseText || '{}');

        if (payload && payload.d && Array.isArray(payload.d.results)) return payload.d.results;
        if (payload && Array.isArray(payload.value)) return payload.value;
        return [];
    }

    async function loadProductivityFromSharePoint(options = {}) {
        const isAuto = !!options.auto;
        try {
            showToast((isAuto ? 'Auto loading' : 'Loading') + ' productivity from SharePoint...');

            const listResponse = await gmRequest({
                url: SHAREPOINT_FILES_API,
                responseType: 'json',
                headers: {
                    'accept': 'application/json;odata=verbose'
                }
            });

            const files = parseSharePointFilesResponse(listResponse);
            const best = pickBestSharePointProductivityFile(files);

            if (!best) {
                showToast('No matching SharePoint productivity file found. Need name with productivity / bonus / weekly / monthly.', true);
                return;
            }

            const serverRelativeUrl = String(best.ServerRelativeUrl || '');
            if (!serverRelativeUrl) {
                showToast('Selected SharePoint file has no ServerRelativeUrl: ' + best.Name, true);
                return;
            }

            const downloadUrl = SHAREPOINT_ORIGIN + encodeURI(serverRelativeUrl).replace(/#/g, '%23');
            showToast('Selected productivity: ' + best.Name);

            const fileResponse = await gmRequest({
                url: downloadUrl,
                responseType: 'arraybuffer',
                headers: {
                    'accept': 'application/octet-stream'
                },
                timeout: 60000
            });

            await loadProductivityFromArrayBuffer(fileResponse.response, best.Name, best.TimeLastModified);
        } catch (err) {
            console.error('[AYWD] SharePoint productivity load failed', err);
            showToast('SharePoint productivity load error: ' + err.message, true);
        }
    }

    async function loadProductivityFromArrayBuffer(buffer, sourceName, sourceModified) {
        if (typeof XLSX === 'undefined') {
            alert('XLSX library not loaded.');
            return;
        }

        try {
            const wb = XLSX.read(buffer, { type: 'array' });
            const sheetName = wb.SheetNames.find(n => String(n).trim().toLowerCase() === 'productivity by activity') ||
                wb.SheetNames.find(n => String(n).toLowerCase().includes('productivity')) ||
                wb.SheetNames[0];

            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
            state.productivity = buildProductivityMap(rows);
            state.productivityLoadedAt = new Date().toISOString();
            state.productivitySourceName = sourceName || '';
            state.productivitySourceModified = sourceModified || null;
            saveProductivityData();
            renderProductivityStatus();
            renderAssignments(false);
            showToast('Productivity loaded: ' + Object.keys(state.productivity).length + ' Pick workers\nFile: ' + (sourceName || 'unknown'));
        } catch (err) {
            console.error('[AYWD] Excel productivity parse failed', err);
            showToast('Excel productivity parse error: ' + err.message, true);
        }
    }

    function findColumn(row, names) {
        const keys = Object.keys(row);
        for (const wanted of names) {
            const found = keys.find(k => String(k).trim().toLowerCase() === String(wanted).trim().toLowerCase());
            if (found) return found;
        }
        const lower = names.map(x => String(x).toLowerCase());
        return keys.find(k => lower.some(n => String(k).toLowerCase().includes(n))) || null;
    }

    function parseNumberSafe(value) { if (value === null || value === undefined || value === '') return null; const n = Number(String(value).replace(',', '.')); return Number.isFinite(n) ? n : null; }
    function avg(arr) { return arr && arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null; }

    function isPickProductivityRow(row) {
        const codeCol = findColumn(row, ['Code', 'CODE', 'Activity Code', 'ACTIVITY_CODE']);
        const activityCol = findColumn(row, ['Activity_Description', 'Activity Description', 'ACTIVITY_DESCRIPTION', 'Description']);
        const code = codeCol ? String(row[codeCol] || '').trim().toLowerCase() : '';
        const activity = activityCol ? String(row[activityCol] || '').trim().toLowerCase() : '';
        if (codeCol) return code.startsWith('pick');
        return activity.startsWith('pick');
    }

    function productivityActivityType(row) {
        const activityCol = findColumn(row, ['Activity_Description', 'Activity Description', 'ACTIVITY_DESCRIPTION', 'Description']);
        const codeCol = findColumn(row, ['Code', 'CODE', 'Activity Code', 'ACTIVITY_CODE']);
        const joined = ((activityCol ? String(row[activityCol] || '') : '') + ' ' + (codeCol ? String(row[codeCol] || '') : '')).toLowerCase();
        if (joined.includes('single')) return 'single';
        if (joined.includes('mult')) return 'multi';
        if (joined.includes('cross-dock') || joined.includes('cross dock') || joined.includes('icd')) return 'icd';
        return 'pick';
    }

    function buildProductivityMap(rows) {
        const map = {};
        rows.forEach(row => {
            if (!isPickProductivityRow(row)) return;
            const userCol = findColumn(row, ['USER_ID', 'User ID', 'Login', 'Employee', 'Operator']);
            if (!userCol) return;
            const login = normalizeLogin(row[userCol]);
            if (!login) return;
            const uphCol = findColumn(row, ['UpH', 'UPH', 'uph']);
            const prodCol = findColumn(row, ['Productivity', 'PRODUCTIVITY']);
            const itemsCol = findColumn(row, ['Items', 'ITEMS']);
            const type = productivityActivityType(row);
            const uph = parseNumberSafe(uphCol ? row[uphCol] : null);
            const prod = parseNumberSafe(prodCol ? row[prodCol] : null);
            const items = parseNumberSafe(itemsCol ? row[itemsCol] : null) || 0;
            if (!map[login]) map[login] = { login, pickValues: [], singleValues: [], multiValues: [], icdValues: [], productivityValues: [], items: 0, source: 'EXCEL_PICK_ONLY' };
            const rec = map[login];
            rec.items += items;
            if (uph !== null) {
                rec.pickValues.push(uph);
                if (type === 'single') rec.singleValues.push(uph);
                if (type === 'multi') rec.multiValues.push(uph);
                if (type === 'icd') rec.icdValues.push(uph);
            }
            if (prod !== null) rec.productivityValues.push(prod);
        });
        Object.keys(map).forEach(login => {
            const rec = map[login];
            rec.uph = avg(rec.pickValues); rec.singleUPH = avg(rec.singleValues); rec.multiUPH = avg(rec.multiValues); rec.icdUPH = avg(rec.icdValues); rec.productivity = avg(rec.productivityValues);
            delete rec.pickValues; delete rec.singleValues; delete rec.multiValues; delete rec.icdValues; delete rec.productivityValues;
        });
        return map;
    }

    function saveProductivityData() { localStorage.setItem(PRODUCTIVITY_STORAGE_KEY, JSON.stringify({ productivity: state.productivity || {}, productivityLoadedAt: state.productivityLoadedAt || null, productivitySourceName: state.productivitySourceName || '', productivitySourceModified: state.productivitySourceModified || null })); }
    function loadProductivityData() { try { const raw = localStorage.getItem(PRODUCTIVITY_STORAGE_KEY); if (!raw) return; const data = JSON.parse(raw); state.productivity = data.productivity || {}; state.productivityLoadedAt = data.productivityLoadedAt || null; state.productivitySourceName = data.productivitySourceName || ''; state.productivitySourceModified = data.productivitySourceModified || null; } catch (e) { console.warn('[AYWD] productivity load failed', e); } }
    function renderProductivityStatus() { const el = document.getElementById('aywdProductivityStatus'); if (!el) return; const count = state.productivity ? Object.keys(state.productivity).length : 0; const time = state.productivityLoadedAt ? new Date(state.productivityLoadedAt).toLocaleTimeString('pl-PL') : '-'; const file = state.productivitySourceName ? ' / ' + state.productivitySourceName : ''; el.textContent = count ? count + ' Pick workers / ' + time + file : 'not loaded'; el.className = count ? 'ok' : 'warn'; }

    // ============================================================
    // SHORTS workers
    // ============================================================
    function loadShortWorkers() { try { const raw = localStorage.getItem(SHORTS_STORAGE_KEY); state.shortWorkers = raw ? JSON.parse(raw).map(normalizeLogin).filter(Boolean) : []; } catch (e) { state.shortWorkers = []; console.warn('[AYWD] shorts workers load failed', e); } }
    function saveShortWorkers() { state.shortWorkers = Array.from(new Set((state.shortWorkers || []).map(normalizeLogin).filter(Boolean))).sort(); localStorage.setItem(SHORTS_STORAGE_KEY, JSON.stringify(state.shortWorkers)); renderShortsStatus(); }
    function renderShortsStatus() { const el = document.getElementById('aywdShortsStatus'); if (!el) return; const count = state.shortWorkers ? state.shortWorkers.length : 0; el.textContent = count ? count + ' workers' : '0 workers'; el.className = count ? 'ok' : 'warn'; }
    function editShortWorkers() { const input = prompt('SHORTS workers for PCON / CCON\nPaste logins separated by comma, space, semicolon or new line:', (state.shortWorkers || []).join('\n')); if (input === null) return; state.shortWorkers = Array.from(new Set(String(input).split(/[\s,;]+/).map(normalizeLogin).filter(Boolean))).sort(); saveShortWorkers(); showToast('SHORTS workers saved: ' + state.shortWorkers.length); }

    // ============================================================
    // WMS loading
    // ============================================================
    async function loadWms(options = {}) {
        const manual = !!options.manual;
        const silent = !!options.silent;
        if (isWmsLoading) { if (manual) showToast('WMS refresh already running. You can continue scanning.'); return; }
        try {
            isWmsLoading = true;
            setLoading(true, manual ? 'manual' : 'auto');
            const prefixes = STREAMS.map(x => x.p).join("','");
            const sql = 'SELECT ' +
                'mt."PRIORITY" AS "WMS_PRIO", mt."TASK_ID" AS "ORDER_ID", mt."WORK_ZONE", SUBSTR(mt."FROM_LOC_ID",1,2) AS "floor", mt."LIST_ID", mt."STATUS", mt."USER_ID", mt."DSTAMP", mt."WORK_GROUP", mt."SHIPMENT_GROUP", oh."ORDER_TYPE", oh."CARRIER_ID", oh."SHIP_BY_DATE", mt."FROM_LOC_ID", mt."TAG_ID", mt."SKU_ID" ' +
                'FROM "GODLX83P"."MOVE_TASK" mt INNER JOIN "GODLX83P"."ORDER_HEADER" oh ON mt."TASK_ID" = oh."ORDER_ID" ' +
                "WHERE SUBSTR(mt.\"LIST_ID\",1,4) in ('" + prefixes + "') " +
                'AND mt."STATUS" in (\'Released\',\'In Progress\') AND mt."TASK_TYPE" = \'O\' AND SUBSTR(mt."WORK_ZONE",5,1) = \'L\' ORDER BY mt."LIST_ID"';
            const response = await grafanaQuery(sql);
            const rows = parseGrafanaResponse(response);
            state.raw = rows;
            preparePicklists(rows);
            buildZones();
            rebuildZoneCountsFromAssignments();
            const changed = updateAssignmentWmsStatuses();
            state.loadedAt = new Date().toISOString();
            document.getElementById('aywdLoadedAt').textContent = new Date(state.loadedAt).toLocaleTimeString('pl-PL');
            renderStreams(); renderSummary(); updateStats(); savePersistedState();
            if (manual || changed) renderAssignments(false);
            if (!silent) showToast((manual ? 'Loaded' : 'Auto refreshed') + ' picklists: ' + state.picklists.length);
            if (manual) startAutoRefresh(true);
        } catch (err) {
            console.error('[AYWD] LOAD ERROR:', err);
            if (manual) alert('LOAD ERROR:\n' + err.message); else showToast('Auto refresh error:\n' + err.message, true);
            if (typeof GM_notification === 'function' && manual) GM_notification({ title: 'AYWD Load Error', text: err.message, timeout: 5000 });
        } finally {
            isWmsLoading = false; setLoading(false);
            const loginInput = document.getElementById('aywdLogin');
            const scannerInput = document.getElementById('aywdScanner');
            if (document.activeElement !== scannerInput) loginInput.focus();
        }
    }

    function startAutoRefresh(showMessage) { if (autoRefreshTimer) clearInterval(autoRefreshTimer); state.autoRefreshEnabled = true; savePersistedState(); autoRefreshTimer = setInterval(() => loadWms({ manual: false, silent: true }), AUTO_REFRESH_MS); updateAutoStatus(); if (showMessage) showToast('Auto WMS refresh ON: every 30 sec'); }
    function updateAutoStatus() { const el = document.getElementById('aywdAutoStatus'); if (!el) return; el.textContent = state.autoRefreshEnabled ? 'ON / 30 sec' : 'OFF'; el.className = state.autoRefreshEnabled ? 'ok' : 'warn'; }

    async function grafanaQuery(sql) {
        const now = new Date();
        const body = { queries: [{ refId: 'A', datasource: { uid: DATASOURCE_UID, type: 'postgres' }, rawSql: sql, format: 'table', datasourceId: DATASOURCE_ID, intervalMs: 60000, maxDataPoints: 1447 }], range: { from: now.toISOString(), to: now.toISOString(), raw: { from: now.toISOString(), to: now.toISOString() } }, from: String(now.getTime()), to: String(now.getTime()) };
        const res = await fetch(GRAFANA_QUERY_URL, { method: 'POST', headers: { 'accept': 'application/json, text/plain, */*', 'content-type': 'application/json', 'x-grafana-org-id': '1' }, body: JSON.stringify(body), credentials: 'include' });
        if (!res.ok) { const text = await res.text(); throw new Error('Grafana query failed: ' + res.status + ' | ' + text.slice(0, 300)); }
        return await res.json();
    }

    function parseGrafanaResponse(raw) {
        const frame = raw && raw.results && raw.results.A && raw.results.A.frames && raw.results.A.frames[0];
        if (!frame) return [];
        const values = frame.data.values;
        const fields = frame.schema.fields.map(f => f.name);
        const output = [];
        for (let r = 0; r < values[0].length; r++) { const row = {}; for (let c = 0; c < fields.length; c++) row[fields[c]] = values[c][r]; output.push(row); }
        return output;
    }

    function preparePicklists(rows) {
        const map = new Map();
        rows.forEach(mt => {
            const listId = String(mt.LIST_ID || '');
            const prefix = listId.substring(0, 4);
            const zone = mt.WORK_ZONE;
            if (!listId || !prefix || !zone) return;
            if (!map.has(listId)) map.set(listId, { list: listId, prefix, zoneRaw: zone, zone: formatZone(zone), floor: getFloor(zone), prio: Number(mt.WMS_PRIO || 0), units: 0, status: mt.STATUS, user: mt.USER_ID, shipmentGroup: String(mt.SHIPMENT_GROUP || '') });
            const pl = map.get(listId);
            pl.units++;
            pl.prio = Math.max(pl.prio, Number(mt.WMS_PRIO || 0));
            if (mt.STATUS === 'In Progress') { pl.status = mt.STATUS; pl.user = mt.USER_ID; }
        });
        state.picklists = Array.from(map.values()).sort((a, b) => (b.prio - a.prio) || (b.units - a.units) || a.list.localeCompare(b.list));
    }

    function buildZones() {
        state.zones = {};
        state.picklists.forEach(pl => {
            const status = String(pl.status || '').toLowerCase();
            if (status !== 'released' && status !== 'in progress') return;
            if (!state.zones[pl.prefix]) state.zones[pl.prefix] = {};
            if (!state.zones[pl.prefix][pl.zoneRaw]) state.zones[pl.prefix][pl.zoneRaw] = { prefix: pl.prefix, zoneRaw: pl.zoneRaw, zone: pl.zone, floor: pl.floor, count: 0, prio: 0, releasedPrio: 0, units: 0, releasedUnits: 0, lists: [], osrOrder: 999999999, hasInProgress: false };
            const z = state.zones[pl.prefix][pl.zoneRaw];
            z.prio = Math.max(z.prio, Number(pl.prio || 0));
            z.units += Number(pl.units || 0);
            z.lists.push(pl);
            z.osrOrder = Math.min(z.osrOrder, osrOrder(pl.shipmentGroup));
            if (status === 'released') { z.releasedPrio = Math.max(z.releasedPrio || 0, Number(pl.prio || 0)); z.releasedUnits += Number(pl.units || 0); }
            if (status === 'in progress') z.hasInProgress = true;
        });
    }

    function rebuildZoneCountsFromAssignments() { Object.keys(state.zones).forEach(prefix => Object.keys(state.zones[prefix]).forEach(zoneRaw => state.zones[prefix][zoneRaw].count = 0)); state.assignments.forEach(a => { if (state.zones[a.prefix] && state.zones[a.prefix][a.zoneRaw]) state.zones[a.prefix][a.zoneRaw].count++; }); }
    function updateAssignmentWmsStatuses() { const statusByList = {}; state.picklists.forEach(pl => { statusByList[pl.list] = pl.status; }); let changed = false; state.assignments.forEach(a => { const next = statusByList[a.list] || 'NOT_FOUND'; if (a.wmsStatus !== next) { a.wmsStatus = next; changed = true; } }); return changed; }

    // ============================================================
    // Assignment core
    // ============================================================
    function assign(login, scanner) {
        updateStreamValuesFromTable();
        if (state.picklists.length === 0) return alert('First LOAD WMS');
        const worker = getWorkerPerf(login);
        const prefix = nextStreamForWorker(worker);
        if (!prefix) return alert('All streams full or no WMS lists for remaining streams.');
        const selected = getBestGlobalAssignment(prefix, worker);
        if (!selected || !selected.zone || !selected.list) return alert('NO RELEASED LIST AVAILABLE FOR ' + prefix);
        const zone = selected.zone;
        const listId = selected.list.list;
        if (isListAlreadyIssued(listId)) return alert('LIST ALREADY ISSUED: ' + listId + '\nAYWD will not issue the same list twice. Scan again or LOAD WMS.');
        zone.count++;
        state.end = new Date().toISOString();
        state.assignments.push({ id: makeId(), login, scanner, prefix, zoneRaw: zone.zoneRaw, zone: zone.zone, floor: zone.floor, list: listId, time: state.end, wmsStatus: getPicklistStatus(listId), workerUPH: worker.uph, workerProductivity: worker.productivity, workerSource: worker.source });
        renderStreams(); renderAssignments(true); renderSummary(); updateStats(); savePersistedState();
    }

    function getPicklistStatus(listId) { const pl = state.picklists.find(x => x.list === listId); return pl ? pl.status : 'NOT_FOUND'; }

    function nextStreamForWorker(worker) {
        const low = isLowUPH(worker);
        if (worker && isShortWorker(worker.login)) {
            if (activeShortsOnShortListsCount() >= SHORTS_TO_MULTI_MIN_ACTIVE && getHighRiskShortsMultiCandidates().length > 0) return SHORTS_MULTI_PREFIX;
            const shortsPrefix = getShortsPreferredPrefix(worker);
            if (shortsPrefix) return shortsPrefix;
        }
        for (const prefix of ASSIGNMENT_PRIORITY) {
            const stream = STREAMS.find(x => x.p === prefix);
            if (!stream) continue;
            const needed = Number(stream.n || 0);
            const used = prefixDone(stream.p);
            const hasAssignable = getAllAssignableLists(stream.p).length > 0;
            if (!hasAssignable || needed <= 0 || used >= needed) continue;
            if (low && isSinglePrefix(stream.p)) {
                const alternative = ASSIGNMENT_PRIORITY.some(altPrefix => {
                    if (altPrefix === stream.p) return false;
                    const alt = STREAMS.find(x => x.p === altPrefix);
                    if (!alt) return false;
                    return Number(alt.n || 0) > prefixDone(alt.p) && getAllAssignableLists(alt.p).length > 0;
                });
                if (alternative) continue;
            }
            return stream.p;
        }
        return null;
    }

    function nextStream() { return nextStreamForWorker({ uph: null }); }

    function removeAssignmentByIndex(index) {
        if (index < 0 || index >= state.assignments.length) return;
        const removed = state.assignments.splice(index, 1)[0];
        if (removed && state.zones[removed.prefix] && state.zones[removed.prefix][removed.zoneRaw]) state.zones[removed.prefix][removed.zoneRaw].count = Math.max(0, state.zones[removed.prefix][removed.zoneRaw].count - 1);
        if (state.assignments.length === 0) { state.start = null; state.end = null; } else state.end = state.assignments[state.assignments.length - 1].time || new Date().toISOString();
        renderStreams(); renderAssignments(false); renderSummary(); updateStats(); savePersistedState();
    }

    // ============================================================
    // Render
    // ============================================================
    function showDuplicateReason(index) {
        const a = state.assignments[index]; if (!a) return;
        const sameScanner = state.assignments.map((x, i) => ({ ...x, index: i })).filter(x => x.scanner === a.scanner && x.index !== index);
        const sameLogin = state.assignments.map((x, i) => ({ ...x, index: i })).filter(x => x.login === a.login && x.index !== index);
        let msg = 'TOKEN DETAILS\nLogin: ' + a.login + '\nScanner: ' + a.scanner + '\nStrefa: ' + a.zone + '\nLista: ' + a.list + '\nWMS status: ' + (a.wmsStatus || '-') + '\nUPH: ' + (a.workerUPH !== null && a.workerUPH !== undefined ? Number(a.workerUPH).toFixed(1) : 'NO UPH / LOW');
        if (sameScanner.length || sameLogin.length) {
            msg += '\n\nDUPLICATE REASON:';
            if (sameScanner.length) { msg += '\nScanner ' + a.scanner + ' already used by:'; sameScanner.forEach(x => msg += '\n- ' + x.login + ' / ' + x.zone + ' / ' + x.list); }
            if (sameLogin.length) { msg += '\nLogin ' + a.login + ' already exists on:'; sameLogin.forEach(x => msg += '\n- ' + x.scanner + ' / ' + x.zone + ' / ' + x.list); }
        } else msg += '\n\nNo duplicate found.';
        showToast(msg, sameScanner.length || sameLogin.length);
    }

    function renderStreams() {
        const table = document.getElementById('aywdStreams');
        table.innerHTML = '<tr><th>Prefix</th><th>Stream</th><th>Need</th><th>Done</th><th>Lists</th></tr>';
        STREAMS.forEach((stream, i) => {
            const assigned = prefixDone(stream.p);
            const listCount = state.picklists.filter(x => x.prefix === stream.p).length;
            const statusClass = assigned >= Number(stream.n || 0) && Number(stream.n || 0) > 0 ? 'ok' : 'warn';
            table.innerHTML += '<tr><td><span class="pill" style="background:' + prefixColor(stream.p) + '">' + stream.p + '</span></td><td>' + stream.s + '</td><td><input class="need-input" value="' + stream.n + '" data-index="' + i + '" type="number" min="0"></td><td class="' + statusClass + '">' + assigned + '</td><td>' + listCount + '</td></tr>';
        });
        document.querySelectorAll('.need-input').forEach(input => input.addEventListener('input', e => { const idx = Number(e.target.dataset.index); STREAMS[idx].n = Number(e.target.value || 0); savePersistedState(); updateStats(); }));
    }

    function scrollTokensToLatest() { const grid = document.getElementById('aywdGrid'); if (!grid) return; setTimeout(() => { grid.scrollLeft = grid.scrollWidth; }, 50); }

    function renderAssignments(scrollToLatest = false) {
        const grid = document.getElementById('aywdGrid');
        const scannerCounter = {}, loginCounter = {};
        state.assignments.forEach(a => { scannerCounter[a.scanner] = (scannerCounter[a.scanner] || 0) + 1; loginCounter[a.login] = (loginCounter[a.login] || 0) + 1; });
        const frag = document.createDocumentFragment();
        state.assignments.forEach((a, index) => {
            const duplicate = scannerCounter[a.scanner] > 1 || loginCounter[a.login] > 1;
            const isInProgress = String(a.wmsStatus || '').toLowerCase() === 'in progress';
            const isNotFound = String(a.wmsStatus || '').toUpperCase() === 'NOT_FOUND';
            const color = isInProgress ? '#64748b' : prefixColor(a.prefix);
            const uphLow = a.workerUPH === null || a.workerUPH === undefined || Number(a.workerUPH) < LOW_UPH_LIMIT;
            const uphClass = uphLow ? 'low-uph' : 'good-uph';
            const uphText = a.workerUPH !== null && a.workerUPH !== undefined ? 'UPH ' + Number(a.workerUPH).toFixed(1) : 'NO UPH';
            const card = document.createElement('div');
            card.className = 'card ' + (duplicate ? 'duplicate ' : '') + (isInProgress ? 'in-progress-gray' : '');
            card.dataset.index = String(index);
            card.style.borderLeftColor = color;
            let badges = '';
            if (isInProgress) badges += '<span class="status-badge in-progress">In Progress</span>';
            if (isNotFound) badges += '<span class="status-badge not-found">NOT_FOUND</span>';
            if (state.showUPH) badges += '<span class="status-badge uph-badge ' + uphClass + '">' + escapeHtml(uphText) + '</span>';
            card.innerHTML = '<div class="login">' + escapeHtml(a.login) + '</div><div class="zone">' + escapeHtml(a.zone) + '</div><div class="list" style="color:' + color + '">' + formatListForDisplay(a.list) + '</div>' + (badges ? '<div>' + badges + '</div>' : '');
            frag.appendChild(card);
        });
        grid.replaceChildren(frag);
        grid.querySelectorAll('.card').forEach(card => {
            let clickTimer = null;
            card.addEventListener('click', () => { const idx = Number(card.dataset.index); clearTimeout(clickTimer); clickTimer = setTimeout(() => showDuplicateReason(idx), 220); });
            card.addEventListener('dblclick', e => { e.preventDefault(); clearTimeout(clickTimer); removeAssignmentByIndex(Number(card.dataset.index)); showToast('Token removed. List can be selected again by priority.'); });
        });
        applyUPHVisibility(); if (scrollToLatest) scrollTokensToLatest();
    }

    function renderSummary() {
        const summary = document.getElementById('aywdSummary');
        if (state.picklists.length === 0) { summary.innerHTML = 'No data loaded.'; return; }
        let html = '';
        STREAMS.forEach(stream => {
            const lists = state.picklists.filter(x => x.prefix === stream.p);
            const zones = Object.values(state.zones[stream.p] || {});
            const units = lists.reduce((sum, x) => sum + Number(x.units || 0), 0);
            const maxPrio = getImportantPrioForPrefix(stream.p) || 0;
            const done = prefixDone(stream.p);
            html += '<div><b style="color:' + prefixColor(stream.p) + '">' + stream.p + '</b>: ' + lists.length + ' lists / ' + units + ' units / ' + zones.length + ' zones / prio ' + maxPrio + ' / issued ' + done + '</div>';
            if (stream.p === 'TCON') {
                buildOsrSummaryForPrefix(stream.p).forEach(g => {
                    html += '<div style="margin-left:14px;color:#cbd5e1;font-size:12px;line-height:1.35">↳ <b>' + escapeHtml(g.osr) + '</b>: ' + g.lists + ' lists / rel ' + g.releasedLists + ' / IP ' + g.inProgressLists + ' / ' + g.units + ' units / prio ' + g.prio + ' / issued <span class="ok">' + g.issued + '</span></div>';
                });
            }
        });
        summary.innerHTML = html;
    }

    function updateStats() {
        document.getElementById('aywdCount').textContent = state.assignments.length;
        const next = nextStream();
        const el = document.getElementById('aywdCurrentStream');
        el.textContent = next ? next : 'FULL';
        el.className = next ? 'ok' : 'full';
        document.getElementById('aywdStart').textContent = state.start ? new Date(state.start).toLocaleTimeString('pl-PL') : '-';
        document.getElementById('aywdEnd').textContent = state.end ? new Date(state.end).toLocaleTimeString('pl-PL') : '-';
        if (state.start && state.end) { const sec = Math.round((new Date(state.end) - new Date(state.start)) / 1000); document.getElementById('aywdDuration').textContent = Math.floor(sec / 60) + 'm ' + (sec % 60) + 's'; }
        else document.getElementById('aywdDuration').textContent = '-';
    }

    // ============================================================
    // Actions and persistence
    // ============================================================
    function updateStreamValuesFromTable() { document.querySelectorAll('.need-input').forEach(input => { const idx = Number(input.dataset.index); STREAMS[idx].n = Number(input.value || 0); }); }
    function copyData() { const text = state.assignments.map(a => a.login + '\t' + a.scanner).join('\n'); navigator.clipboard.writeText(text); showToast('Copied login + scanner'); }
    function undo() { removeAssignmentByIndex(state.assignments.length - 1); document.getElementById('aywdLogin').focus(); }
    function clearAll() { state.assignments = []; state.login = ''; state.start = null; state.end = null; Object.keys(state.zones).forEach(prefix => Object.keys(state.zones[prefix]).forEach(zone => { state.zones[prefix][zone].count = 0; })); savePersistedState(); renderStreams(); renderAssignments(false); renderSummary(); updateStats(); document.getElementById('aywdLogin').focus(); showToast('All tokens cleared'); }
    function closeOverlay() { if (autoRefreshTimer) { clearInterval(autoRefreshTimer); autoRefreshTimer = null; } const root = document.getElementById(APP_ID); if (root) root.remove(); const toast = document.getElementById(APP_ID + '_toast'); if (toast) toast.remove(); }

    function formatZone(zone) {
        if (zone === '50P4L0B') return 'L1B ICD (żółta kartka)';
        if (zone === '50P1L1B') return 'L1B SHOE';
        const m = String(zone || '').match(/L(\d)([A-Z])$/);
        return m ? 'L' + m[1] + m[2] : zone;
    }
    function getFloor(zone) { if (zone === '50P4L0B') return 'L1'; const m = String(zone || '').match(/L(\d)/); return m ? 'L' + m[1] : 'L?'; }
    function formatListForDisplay(list) { const safe = escapeHtml(String(list || '')); return safe ? '<span class="list-first">' + safe[0] + '</span>' + safe.slice(1) : ''; }
    function setLoading(isLoading, mode) { const btn = document.getElementById('aywdLoad'), st = document.getElementById('aywdRefreshState'); btn.disabled = isLoading && mode === 'manual'; btn.textContent = isLoading ? (mode === 'auto' ? 'AUTO...' : 'LOADING...') : 'LOAD WMS'; if (st) st.textContent = isLoading ? (mode === 'auto' ? 'auto refreshing...' : 'manual loading...') : 'idle'; }
    function showToast(text, error = false) { const toast = document.getElementById(APP_ID + '_toast'); toast.textContent = text; toast.style.background = error ? '#dc2626' : '#16a34a'; toast.style.display = 'block'; clearTimeout(showToast.timer); showToast.timer = setTimeout(() => { toast.style.display = 'none'; }, 4500); }

    function savePersistedState() {
        try {
            const payload = { assignments: state.assignments, start: state.start, end: state.end, STREAMS, autoRefreshEnabled: state.autoRefreshEnabled, showUPH: state.showUPH };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
        } catch (e) { console.warn('[AYWD] save failed', e); }
    }

    function loadPersistedState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return;
            const payload = JSON.parse(raw);
            if (Array.isArray(payload.assignments)) state.assignments = payload.assignments;
            if (payload.start) state.start = payload.start;
            if (payload.end) state.end = payload.end;
            if (Array.isArray(payload.STREAMS)) STREAMS = payload.STREAMS;
            if (typeof payload.autoRefreshEnabled === 'boolean') state.autoRefreshEnabled = payload.autoRefreshEnabled;
            if (typeof payload.showUPH === 'boolean') state.showUPH = payload.showUPH;
        } catch (e) { console.warn('[AYWD] load failed', e); }
    }
})();
