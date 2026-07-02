// ==UserScript==
// @name         AYWD - Work Distributor LOADER
// @namespace    https://github.com/yevh-k/AYWorkDistributor
// @version      1.0.1
// @description  Loader for AYWD remote main script from GitHub; keeps SharePoint productivity auto-load
// @author       Yevhenii Karpenko
// @match        https://gdcgrafana-eu.logistics.corp/d/dOHkIABY83AGEIHMDTLOPSPRODSTD/ageing-heatmap-detail-prod*
// @grant        GM_notification
// @grant        GM_xmlhttpRequest
// @connect      raw.githubusercontent.com
// @connect      github.com
// @connect      cdn.jsdelivr.net
// @connect      cevalogisticsoffice365.sharepoint.com
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  var ALWAYS_FRESH = true;

  // If your branch is master, change main to master in these URLs.
  var AYWD_URLS = [
    'https://raw.githubusercontent.com/yevh-k/AYWorkDistributor/main/aywd-main.js',
    'https://github.com/yevh-k/AYWorkDistributor/raw/refs/heads/main/aywd-main.js',
    'https://cdn.jsdelivr.net/gh/yevh-k/AYWorkDistributor@main/aywd-main.js'
  ];

  function log() {
    try { console.log.apply(console, ['[AYWD LOADER]'].concat([].slice.call(arguments))); } catch (e) {}
  }

  function showError(message) {
    try {
      console.error('[AYWD LOADER]', message);
      var old = document.getElementById('aywd-loader-error');
      if (old) old.remove();
      var box = document.createElement('div');
      box.id = 'aywd-loader-error';
      box.style.position = 'fixed';
      box.style.top = '12px';
      box.style.right = '12px';
      box.style.zIndex = '2147483647';
      box.style.maxWidth = '620px';
      box.style.padding = '12px 14px';
      box.style.background = '#fee2e2';
      box.style.color = '#7f1d1d';
      box.style.border = '1px solid #ef4444';
      box.style.borderRadius = '8px';
      box.style.font = '12px Arial, Segoe UI, sans-serif';
      box.style.boxShadow = '0 8px 24px rgba(0,0,0,.18)';
      box.style.whiteSpace = 'pre-wrap';
      box.textContent = 'AYWD loader error:' + String.fromCharCode(10) + String(message || 'Unknown error');
      (document.body || document.documentElement).appendChild(box);
    } catch (e) {
      alert('AYWD loader error: ' + String(message || e));
    }
  }

  function looksLikeHtml(text) {
    text = String(text || '').trim().slice(0, 300).toLowerCase();
    return text.indexOf('<!doctype html') === 0 || text.indexOf('<html') === 0 || text.indexOf('<head') === 0;
  }

  function looksLikeAywdJs(text) {
    text = String(text || '');
    return text.indexOf('AYWD') >= 0 &&
           (text.indexOf('function init') >= 0 || text.indexOf('const REQUIRED_HASH') >= 0 || text.indexOf("'#AYWD'") >= 0);
  }

  function runCode(code, sourceUrl) {
    try {
      code = String(code || '').replace(/^\uFEFF/, '');
      if (looksLikeHtml(code)) {
        throw new Error('Downloaded HTML instead of JavaScript from: ' + sourceUrl + '\nFirst chars:\n' + code.slice(0, 300));
      }
      if (!looksLikeAywdJs(code)) {
        throw new Error('Downloaded file does not look like AYWD JavaScript from: ' + sourceUrl + '\nFirst chars:\n' + code.slice(0, 300));
      }

      // Direct eval keeps Tampermonkey scope. AYWD SharePoint needs GM_xmlhttpRequest and GM_notification.
      eval(code + '\n//# sourceURL=aywd-main.remote.js');
      log('Remote main loaded from', sourceUrl);
    } catch (err) {
      showError(err && err.stack ? err.stack : err);
    }
  }

  function addCacheBuster(url) {
    if (!ALWAYS_FRESH) return url;
    return url + (url.indexOf('?') === -1 ? '?' : '&') + 'v=' + Date.now();
  }

  function loadFrom(index, history) {
    history = history || [];
    if (index >= AYWD_URLS.length) {
      showError('All AYWD URLs failed.' + String.fromCharCode(10) + history.join(String.fromCharCode(10) + String.fromCharCode(10)));
      return;
    }

    var baseUrl = AYWD_URLS[index];
    var url = addCacheBuster(baseUrl);
    log('Loading', url);

    GM_xmlhttpRequest({
      method: 'GET',
      url: url,
      headers: { 'Accept': 'text/javascript,text/plain,*/*' },
      onload: function (r) {
        var status = r ? r.status : 'NO_RESPONSE';
        var text = String((r && r.responseText) || '');
        var ct = '';
        try { ct = r.responseHeaders || ''; } catch (e) {}

        log('Response', status, baseUrl, 'first chars:', text.slice(0, 80));

        if (!r || status < 200 || status >= 300 || looksLikeHtml(text)) {
          history.push('FAILED: ' + baseUrl + '\nHTTP: ' + status + '\nHeaders: ' + ct.slice(0, 250) + '\nFirst chars: ' + text.slice(0, 300));
          loadFrom(index + 1, history);
          return;
        }

        runCode(text, baseUrl);
      },
      onerror: function (e) {
        history.push('NETWORK ERROR: ' + baseUrl + '\n' + JSON.stringify(e));
        loadFrom(index + 1, history);
      },
      ontimeout: function () {
        history.push('TIMEOUT: ' + baseUrl);
        loadFrom(index + 1, history);
      }
    });
  }

  loadFrom(0, []);
})();
