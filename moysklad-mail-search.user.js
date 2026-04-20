// ==UserScript==
// @name         MoySklad - Поиск писем по заказу поставщику
// @namespace    https://tampermonkey.net/
// @version      0.1.17
// @description  Ищет письма по заказу поставщику через Google Apps Script
// @author       Codex + Spiralwave
// @match        https://online.moysklad.ru/app/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @connect      www.dbschenker.com
// @connect      www.ups.com
// @connect      www.dhl.com
// @updateURL    https://raw.githubusercontent.com/spiralvawe/po-work-tm-userscript/main/moysklad-mail-search.user.js
// @downloadURL  https://raw.githubusercontent.com/spiralvawe/po-work-tm-userscript/main/moysklad-mail-search.user.js
// @supportURL   https://github.com/spiralvawe/po-work-tm/issues
// ==/UserScript==

(function () {
  'use strict';

  const APP_CONFIG = {
    GAS_URL: 'https://script.google.com/macros/s/AKfycbwrpec9vzI3mFAtDnysT69N7qwkmNU7cQzxXS_jZSYgBfbaAHayIZL1RvCKizdkXMd5iw/exec',
    DEFAULT_USER_TOKEN: '',
    SEARCH_BUTTON_ID: 'tm-ms-mail-search-button',
    PLACEMENT_BUTTON_ID: 'tm-ms-placement-button',
    TRACKING_BUTTON_ID: 'tm-ms-tracking-button',
    PANEL_ID: 'tm-ms-mail-search-panel',
    PANEL_TITLE_CLASS: 'tm-ms-panel-title',
    STATUS_ID: 'tm-ms-mail-search-status',
    DEFAULT_GMAIL_ACCOUNT_INDEX: 1,
    TRACKING_BUTTON_TOP: '166px',
    PANEL_TOP: '216px',
    STORAGE_KEYS: {
      userToken: 'ms_mail_search_user_token',
      gmailAccountIndex: 'ms_mail_search_gmail_account_index'
    },
    POLL_INTERVAL_MS: 500,
    SECONDARY_PREFETCH_DELAY_MS: 1200
  };

  const state = {
    currentOrderId: null,
    searchPrefetchPromise: null,
    searchPrefetchResult: null,
    searchPrefetchError: null,
    searchPrefetchConsumed: false,
    searchPrefetchStartedAt: null,
    placementMetaPromise: null,
    placementMetaResult: null,
    placementMetaError: null,
    lastPlacementDownloadUrl: '',
    placementDraftId: '',
    placementEmailSent: false,
    trackingFieldValue: '',
    trackingSourceType: '',
    trackingSourceLabel: '',
    trackingSourceId: '',
    trackingSourceHref: '',
    trackingEntries: [],
    trackingLastResult: null,
    isPanelOpen: false,
    activePanelMode: 'search'
  };

  function normalizeGmailAccountIndex(value) {
    if (value == null || value === '') {
      return APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX;
    }

    var parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed < 0) {
      return APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX;
    }

    return Math.floor(parsed);
  }

  function getUserSettings() {
    var storedUserToken = String(
      GM_getValue(APP_CONFIG.STORAGE_KEYS.userToken, '') || ''
    ).trim();
    var storedGmailAccountIndex = GM_getValue(APP_CONFIG.STORAGE_KEYS.gmailAccountIndex, null);
    var shouldPersistDefaults =
      !storedUserToken &&
      String(APP_CONFIG.DEFAULT_USER_TOKEN || '').trim();

    if (shouldPersistDefaults) {
      saveUserSettings({
        userToken: APP_CONFIG.DEFAULT_USER_TOKEN,
        gmailAccountIndex:
          storedGmailAccountIndex == null
            ? APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX
            : normalizeGmailAccountIndex(storedGmailAccountIndex)
      });
      storedUserToken = String(
        GM_getValue(APP_CONFIG.STORAGE_KEYS.userToken, APP_CONFIG.DEFAULT_USER_TOKEN) || ''
      ).trim();
      storedGmailAccountIndex = GM_getValue(
        APP_CONFIG.STORAGE_KEYS.gmailAccountIndex,
        APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX
      );
    }

    return {
      userToken: storedUserToken,
      gmailAccountIndex: normalizeGmailAccountIndex(
        storedGmailAccountIndex == null
          ? APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX
          : storedGmailAccountIndex
      )
    };
  }

  function hasRequiredUserSettings(settings) {
    return Boolean(settings && settings.userToken);
  }

  function saveUserSettings(settings) {
    GM_setValue(APP_CONFIG.STORAGE_KEYS.userToken, String(settings.userToken || '').trim());
    GM_setValue(
      APP_CONFIG.STORAGE_KEYS.gmailAccountIndex,
      normalizeGmailAccountIndex(settings.gmailAccountIndex)
    );
  }

  function promptForUserSettings(existingSettings) {
    var current = existingSettings || getUserSettings();
    var userToken = window.prompt(
      'Вставьте персональный токен сотрудника для доступа к Apps Script',
      current.userToken || ''
    );
    if (userToken == null) {
      return null;
    }

    var gmailAccountIndex = window.prompt(
      'Введите индекс Gmail-аккаунта (обычно 0 или 1)',
      String(current.gmailAccountIndex)
    );
    if (gmailAccountIndex == null) {
      return null;
    }

    return {
      userToken: userToken.trim(),
      gmailAccountIndex: normalizeGmailAccountIndex(gmailAccountIndex)
    };
  }

  function ensureUserSettings(options) {
    var settings = getUserSettings();

    if (hasRequiredUserSettings(settings)) {
      return settings;
    }

    if (options && options.silent) {
      return null;
    }

    var promptedSettings = promptForUserSettings(settings);
    if (!promptedSettings || !hasRequiredUserSettings(promptedSettings)) {
      return null;
    }

    saveUserSettings(promptedSettings);
    return getUserSettings();
  }

  function openSettingsFlow() {
    var promptedSettings = promptForUserSettings(getUserSettings());
    if (!promptedSettings || !hasRequiredUserSettings(promptedSettings)) {
      return false;
    }

    saveUserSettings(promptedSettings);
    updateVisibilityAndPrefetch();
    return true;
  }

  function resetUserSettings() {
    saveUserSettings({
      userToken: '',
      gmailAccountIndex: APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX
    });

    resetOrderState(null);
    setStatus('Настройки очищены', 'neutral');
    hidePanel();
    setPlacementButtonVisible(false);
    setTrackingButtonVisible(false);
  }

  function registerMenuCommands() {
    GM_registerMenuCommand('MoySklad: настроить доступ', function () {
      openSettingsFlow();
    });

    GM_registerMenuCommand('MoySklad: очистить токен и настройки', function () {
      resetUserSettings();
    });
  }

  function forceGmailAccount(link, settings) {
    if (!link) {
      return '#';
    }

    var accountIndex = settings ? settings.gmailAccountIndex : APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX;
    return String(link).replace(/\/mail\/u\/\d+\//, '/mail/u/' + accountIndex + '/');
  }

  function extractEmails(value) {
    var matches = String(value || '').match(/[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}/gi);
    return matches || [];
  }

  function normalizeRecipientList(value) {
    var seen = new Set();

    return extractEmails(value)
      .map(function (email) {
        return String(email || '').trim().toLowerCase();
      })
      .filter(function (email) {
        if (!email || seen.has(email)) {
          return false;
        }

        seen.add(email);
        return true;
      });
  }

  function buildGmailComposeLink(message, settings) {
    var accountIndex = settings ? settings.gmailAccountIndex : APP_CONFIG.DEFAULT_GMAIL_ACCOUNT_INDEX;
    var url = new URL('https://mail.google.com/mail/u/' + accountIndex + '/');
    url.searchParams.set('view', 'cm');
    url.searchParams.set('fs', '1');
    url.searchParams.set('to', String(message.to || ''));
    url.searchParams.set('su', String(message.subject || ''));
    url.searchParams.set('body', String(message.body || ''));
    return url.toString();
  }

  function openExternalUrl(url) {
    if (!url) {
      return;
    }

    var link = document.createElement('a');
    link.href = url;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function normalizeWhitespace(value) {
    return String(value || '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function formatDate(value) {
    if (!value) {
      return '';
    }

    try {
      return new Date(value).toLocaleString('ru-RU');
    } catch (error) {
      return String(value);
    }
  }

  function buildPlacementEmailSubject(data) {
    return 'Purchase Order ' + String((data && data.orderNumber) || '').trim();
  }

  function buildPlacementEmailBody(data) {
    var orderNumber = String((data && data.orderNumber) || '').trim();
    var supplierName = String((data && data.supplierName) || '').trim();

    return [
      'Hello' + (supplierName ? ' ' + supplierName : '') + ',',
      '',
      'Please find attached purchase order ' + orderNumber + '.',
      '',
      'Best regards'
    ].join('\n');
  }

  function formatTrackingEventDate(value) {
    if (!value) {
      return '';
    }

    try {
      return new Date(value).toLocaleString('ru-RU');
    } catch (error) {
      return String(value);
    }
  }

  function isPurchaseOrderPage() {
    var hash = window.location.hash || '';
    return hash.startsWith('#purchaseorder/edit');
  }

  function getOrderIdFromUrl() {
    var hash = window.location.hash || '';
    var hashWithoutSharp = hash.startsWith('#') ? hash.slice(1) : hash;
    var parts = hashWithoutSharp.split('?');

    if (parts.length < 2) {
      return null;
    }

    var params = new URLSearchParams(parts[1]);
    return params.get('id');
  }

  function cleanTrackingToken(token) {
    return String(token || '')
      .replace(/^[^A-Za-z0-9]+/, '')
      .replace(/[^A-Za-z0-9-]+$/, '')
      .trim();
  }

  function isSchenkerTrackingNumber(value) {
    return /^\d{6}-\d{6}$/.test(String(value || '').trim());
  }

  function isUpsTrackingNumber(value) {
    return /^1Z[0-9A-Z]{16}$/i.test(String(value || '').trim());
  }

  function isDhlTrackingNumber(value) {
    return /^(?:\d{10}|\d{20})$/.test(String(value || '').trim());
  }

  function detectCarrierFromTrackingToken(token, carrierHint) {
    var normalizedToken = String(token || '').trim().toUpperCase();
    var normalizedHint = String(carrierHint || '').trim().toLowerCase();

    if (!normalizedToken) {
      return '';
    }

    if (isSchenkerTrackingNumber(normalizedToken)) {
      return 'schenker';
    }

    if (isUpsTrackingNumber(normalizedToken)) {
      return 'ups';
    }

    if (isDhlTrackingNumber(normalizedToken)) {
      return 'dhl';
    }

    if (normalizedHint === 'dhl' && /^[A-Z0-9-]{8,25}$/.test(normalizedToken)) {
      return 'dhl';
    }

    if (normalizedHint === 'ups' && /^[A-Z0-9-]{8,25}$/.test(normalizedToken)) {
      return 'ups';
    }

    if (normalizedHint === 'schenker' && /^[A-Z0-9-]{8,25}$/.test(normalizedToken)) {
      return 'schenker';
    }

    return '';
  }

  function getCarrierLabel(carrier) {
    if (carrier === 'schenker') {
      return 'Schenker';
    }

    if (carrier === 'ups') {
      return 'UPS';
    }

    if (carrier === 'dhl') {
      return 'DHL';
    }

    return 'Неизвестно';
  }

  function buildTrackingUrl(entry) {
    var carrier = entry && entry.carrier;
    var trackingNumber = entry && entry.trackingNumber ? encodeURIComponent(entry.trackingNumber) : '';

    if (!trackingNumber) {
      return '';
    }

    if (carrier === 'schenker') {
      return 'https://www.dbschenker.com/app/tracking-public/?refNumber=' + trackingNumber;
    }

    if (carrier === 'ups') {
      return 'https://www.ups.com/track?tracknum=' + trackingNumber;
    }

    if (carrier === 'dhl') {
      return 'https://www.dhl.com/global-en/home/tracking.html?tracking-id=' + trackingNumber + '&submit=1';
    }

    return '';
  }

  function buildTrackingEntry(trackingNumber, carrier, rawToken) {
    var normalizedTrackingNumber = String(trackingNumber || '').trim();
    var resolvedCarrier = carrier || '';
    var entry = {
      rawToken: String(rawToken || normalizedTrackingNumber).trim(),
      trackingNumber: normalizedTrackingNumber,
      carrier: resolvedCarrier,
      carrierLabel: getCarrierLabel(resolvedCarrier)
    };

    entry.url = buildTrackingUrl(entry);
    return entry;
  }

  function extractTrackingEntriesFromText(value) {
    var normalizedValue = normalizeWhitespace(value);
    var parts;
    var entries = [];
    var seen = {};
    var carrierHint = '';

    if (!normalizedValue) {
      return [];
    }

    parts = normalizedValue.split(/[\s,;|/]+/);

    parts.forEach(function (part) {
      var cleanedToken = cleanTrackingToken(part);
      var normalizedToken = cleanedToken.toUpperCase();
      var carrier;

      if (!cleanedToken) {
        return;
      }

      if (/^(schenker|dbschenker)$/i.test(cleanedToken)) {
        carrierHint = 'schenker';
        return;
      }

      if (/^ups$/i.test(cleanedToken)) {
        carrierHint = 'ups';
        return;
      }

      if (/^dhl$/i.test(cleanedToken)) {
        carrierHint = 'dhl';
        return;
      }

      carrier = detectCarrierFromTrackingToken(normalizedToken, carrierHint);

      if (!carrier) {
        return;
      }

      if (seen[normalizedToken]) {
        return;
      }

      seen[normalizedToken] = true;
      entries.push(buildTrackingEntry(normalizedToken, carrier, cleanedToken));
    });

    return entries;
  }

  function syncTrackingStateFromPlacementData(data) {
    var fieldValue = normalizeWhitespace(data && data.trackingRawValue);

    state.trackingFieldValue = fieldValue;
    state.trackingSourceType = String(data && data.trackingSourceType || '').trim();
    state.trackingSourceLabel = String(data && data.trackingSourceLabel || '').trim();
    state.trackingSourceId = String(data && data.trackingSourceId || '').trim();
    state.trackingSourceHref = String(data && data.trackingSourceHref || '').trim();
    state.trackingEntries = extractTrackingEntriesFromText(fieldValue);

    return state.trackingEntries.slice();
  }

  function getCurrentTrackingEntries() {
    if (state.placementMetaResult && state.placementMetaResult.ok && state.placementMetaResult.data) {
      return syncTrackingStateFromPlacementData(state.placementMetaResult.data);
    }

    return state.trackingEntries.slice();
  }

  function gmRequest(options) {
    return new Promise(function (resolve, reject) {
      GM_xmlhttpRequest({
        method: options.method || 'GET',
        url: options.url,
        headers: options.headers || {},
        data: options.data,
        timeout: options.timeout || 15000,
        onload: function (response) {
          resolve(response);
        },
        onerror: function (error) {
          reject(new Error(error && error.error ? error.error : 'GM request failed'));
        },
        ontimeout: function () {
          reject(new Error('GM request timed out'));
        }
      });
    });
  }

  function tryParseJson(text) {
    var trimmed = String(text || '').trim();

    if (!trimmed || (trimmed.charAt(0) !== '{' && trimmed.charAt(0) !== '[')) {
      return null;
    }

    try {
      return JSON.parse(trimmed);
    } catch (error) {
      return null;
    }
  }

  function scoreTrackingEventArray(items) {
    var score = 0;

    (items || []).forEach(function (item) {
      var keys;

      if (!item || typeof item !== 'object' || Array.isArray(item)) {
        return;
      }

      keys = Object.keys(item);

      if (keys.some(function (key) { return /date|time|timestamp/i.test(key); })) {
        score += 2;
      }

      if (keys.some(function (key) { return /event|status|description|milestone/i.test(key); })) {
        score += 2;
      }

      if (keys.some(function (key) { return /location|city|country|place/i.test(key); })) {
        score += 1;
      }
    });

    return score;
  }

  function findBestTrackingEventArray(value, bestMatch) {
    var currentBest = bestMatch || {
      score: 0,
      value: null
    };

    if (!value || typeof value !== 'object') {
      return currentBest;
    }

    if (Array.isArray(value)) {
      if (value.length && scoreTrackingEventArray(value) > currentBest.score) {
        currentBest = {
          score: scoreTrackingEventArray(value),
          value: value
        };
      }

      value.forEach(function (item) {
        currentBest = findBestTrackingEventArray(item, currentBest);
      });

      return currentBest;
    }

    Object.keys(value).forEach(function (key) {
      currentBest = findBestTrackingEventArray(value[key], currentBest);
    });

    return currentBest;
  }

  function findTrackingSummaryValue(value, patterns) {
    var matchedValue = '';

    if (!value || typeof value !== 'object') {
      return '';
    }

    Object.keys(value).some(function (key) {
      var item = value[key];

      if (patterns.test(key) && typeof item === 'string' && normalizeWhitespace(item)) {
        matchedValue = normalizeWhitespace(item);
        return true;
      }

      if (item && typeof item === 'object') {
        matchedValue = findTrackingSummaryValue(item, patterns);
        return Boolean(matchedValue);
      }

      return false;
    });

    return matchedValue;
  }

  function normalizeTrackingEvent(item) {
    var dateValue = '';
    var title = '';
    var description = '';
    var location = '';
    var source = item || {};

    Object.keys(source).forEach(function (key) {
      var value = source[key];

      if (value == null || value === '') {
        return;
      }

      if (!dateValue && /date|time|timestamp/i.test(key)) {
        dateValue = String(value);
        return;
      }

      if (!title && /event|status|milestone/i.test(key)) {
        title = normalizeWhitespace(value);
        return;
      }

      if (!description && /description|comment|reason|details?/i.test(key)) {
        description = normalizeWhitespace(value);
        return;
      }

      if (!location && /location|city|country|place/i.test(key)) {
        location = normalizeWhitespace(value);
      }
    });

    if (!title && description) {
      title = description;
      description = '';
    }

    return {
      date: dateValue,
      title: title,
      description: description,
      location: location
    };
  }

  function parseSchenkerTrackingResponse(json, entry, sourceUrl) {
    var eventMatch = findBestTrackingEventArray(json);
    var events = eventMatch.value ? eventMatch.value.map(normalizeTrackingEvent).filter(function (item) {
      return item.title || item.description || item.location || item.date;
    }) : [];
    var latestEvent = events.length ? events[0] : null;
    var currentStatus = findTrackingSummaryValue(
      json,
      /current.?status|status.?description|shipment.?status|milestone|status$/i
    ) || (latestEvent ? latestEvent.title : '');

    return {
      success: Boolean(currentStatus || events.length),
      carrier: entry.carrier,
      trackingNumber: entry.trackingNumber,
      currentStatus: currentStatus || 'Статус найден, но не удалось красиво распознать поле.',
      history: events,
      officialUrl: entry.url,
      sourceUrl: sourceUrl
    };
  }

  async function tryFetchSchenkerTracking(entry) {
    var candidateUrls = [
      'https://www.dbschenker.com/nges-portal/api/public/tracking-public?refNumber=' + encodeURIComponent(entry.trackingNumber),
      'https://www.dbschenker.com/nges-portal/api/public/tracking-public?reference=' + encodeURIComponent(entry.trackingNumber),
      'https://www.dbschenker.com/nges-portal/api/public/tracking-public?query=' + encodeURIComponent(entry.trackingNumber)
    ];
    var index;
    var response;
    var json;

    for (index = 0; index < candidateUrls.length; index += 1) {
      try {
        response = await gmRequest({
          method: 'GET',
          url: candidateUrls[index],
          headers: {
            Accept: 'application/json,text/plain,*/*'
          }
        });
        json = tryParseJson(response.responseText);

        if (response.status >= 200 && response.status < 300 && json) {
          return parseSchenkerTrackingResponse(json, entry, candidateUrls[index]);
        }
      } catch (error) {
        // Try the next public probe URL.
      }
    }

    return {
      success: false,
      carrier: entry.carrier,
      trackingNumber: entry.trackingNumber,
      currentStatus: '',
      history: [],
      officialUrl: entry.url,
      sourceUrl: entry.url,
      error: 'Не удалось получить структурированный ответ от публичного сервиса Schenker. Открой официальный трекинг по ссылке.'
    };
  }

  async function fetchTrackingDetails(entry) {
    if (entry.carrier === 'schenker') {
      return tryFetchSchenkerTracking(entry);
    }

    if (entry.carrier === 'dhl') {
      return {
        success: false,
        carrier: entry.carrier,
        trackingNumber: entry.trackingNumber,
        currentStatus: '',
        history: [],
        officialUrl: entry.url,
        sourceUrl: 'https://developer.dhl.com/api-reference/shipment-tracking',
        error: 'У DHL официальное получение истории в API требует subscription key. Без ключа даю прямую ссылку на официальный трекинг.'
      };
    }

    if (entry.carrier === 'ups') {
      return {
        success: false,
        carrier: entry.carrier,
        trackingNumber: entry.trackingNumber,
        currentStatus: '',
        history: [],
        officialUrl: entry.url,
        sourceUrl: 'https://developer.ups.com/us/en/business-solutions/expand-your-online-business/upgrade-digital-technology/developer-resource-center',
        error: 'Для UPS без отдельной developer-интеграции надёжнее открывать официальный трекинг. В панели оставляю прямую ссылку.'
      };
    }

    return {
      success: false,
      carrier: entry.carrier,
      trackingNumber: entry.trackingNumber,
      currentStatus: '',
      history: [],
      officialUrl: entry.url,
      sourceUrl: '',
      error: 'Не удалось определить службу доставки для этого номера.'
    };
  }

  function buildRequestUrl(action, orderId, settings, options) {
    var requestOptions = options || {};
    var url = new URL(APP_CONFIG.GAS_URL);

    url.searchParams.set('action', action);
    url.searchParams.set('id', orderId);
    url.searchParams.set('token', settings.userToken);
    url.searchParams.set('_ts', String(Date.now()));

    if (requestOptions.skipLog) {
      url.searchParams.set('skipLog', '1');
    }

    if (requestOptions.prefetch) {
      url.searchParams.set('prefetch', '1');
    }

    return url.toString();
  }

  async function fetchAppsScriptData(action, orderId, settings, options) {
    var requestUrl = buildRequestUrl(action, orderId, settings, options || {});
    var response = await fetch(requestUrl, {
      method: 'GET',
      redirect: 'follow',
      credentials: 'omit',
      headers: {
        Accept: 'application/json'
      }
    });

    var text = await response.text();
    var trimmed = text.trim();
    var isJson = trimmed.startsWith('{') || trimmed.startsWith('[');

    return {
      ok: isJson,
      status: response.status,
      requestUrl: requestUrl,
      text: text,
      data: isJson ? JSON.parse(trimmed) : null
    };
  }

  async function fetchSearchData(orderId, settings, options) {
    return fetchAppsScriptData('search', orderId, settings, options);
  }

  async function fetchPlacementMeta(orderId, settings) {
    return fetchAppsScriptData('placeMeta', orderId, settings);
  }

  async function fetchPlacementExport(orderId, settings) {
    return fetchAppsScriptData('placeExport', orderId, settings);
  }

  async function fetchPlacementSetState(orderId, settings) {
    return fetchAppsScriptData('placeSetState', orderId, settings);
  }

  async function sendPlacementEmail(orderId, settings, payload) {
    var response = await fetch(APP_CONFIG.GAS_URL, {
      method: 'POST',
      redirect: 'follow',
      credentials: 'omit',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify({
        action: 'placeSend',
        token: settings.userToken,
        id: orderId,
        to: payload.to,
        subject: payload.subject,
        body: payload.body
      })
    });
    var text = await response.text();
    var trimmed = text.trim();
    var isJson = trimmed.startsWith('{') || trimmed.startsWith('[');

    return {
      ok: isJson,
      status: response.status,
      requestUrl: APP_CONFIG.GAS_URL,
      text: text,
      data: isJson ? JSON.parse(trimmed) : null
    };
  }

  async function savePlacementDraft(orderId, settings, payload) {
    var response = await fetch(APP_CONFIG.GAS_URL, {
      method: 'POST',
      redirect: 'follow',
      credentials: 'omit',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify({
        action: 'placeDraft',
        token: settings.userToken,
        id: orderId,
        draftId: payload.draftId,
        to: payload.to,
        subject: payload.subject,
        body: payload.body
      })
    });
    var text = await response.text();
    var trimmed = text.trim();
    var isJson = trimmed.startsWith('{') || trimmed.startsWith('[');

    return {
      ok: isJson,
      status: response.status,
      requestUrl: APP_CONFIG.GAS_URL,
      text: text,
      data: isJson ? JSON.parse(trimmed) : null
    };
  }

  async function saveSupplierEmailsToMoySklad(orderId, settings, payload) {
    var response = await fetch(APP_CONFIG.GAS_URL, {
      method: 'POST',
      redirect: 'follow',
      credentials: 'omit',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify({
        action: 'placeSaveCounterpartyEmails',
        token: settings.userToken,
        id: orderId,
        emails: payload.emails
      })
    });
    var text = await response.text();
    var trimmed = text.trim();
    var isJson = trimmed.startsWith('{') || trimmed.startsWith('[');

    return {
      ok: isJson,
      status: response.status,
      requestUrl: APP_CONFIG.GAS_URL,
      text: text,
      data: isJson ? JSON.parse(trimmed) : null
    };
  }

  async function savePlacementCounterpartyEmails(orderId, settings, payload) {
    return saveSupplierEmailsToMoySklad(orderId, settings, payload);
  }

  async function saveSearchEmailToList(orderId, settings, payload) {
    var response = await fetch(APP_CONFIG.GAS_URL, {
      method: 'POST',
      redirect: 'follow',
      credentials: 'omit',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify({
        action: 'searchSaveEmailList',
        token: settings.userToken,
        id: orderId,
        listType: payload.listType,
        emails: payload.emails
      })
    });
    var text = await response.text();
    var trimmed = text.trim();
    var isJson = trimmed.startsWith('{') || trimmed.startsWith('[');

    return {
      ok: isJson,
      status: response.status,
      requestUrl: APP_CONFIG.GAS_URL,
      text: text,
      data: isJson ? JSON.parse(trimmed) : null
    };
  }

  var SEARCH_STATUS_BORDER_COLORS = {
    loading: '#eab308',
    ok: '#16a34a',
    error: '#dc2626'
  };
  var PLACEMENT_CONFIRMED_BORDER_COLOR = '#9664bf';
  var TRACKING_BUTTON_BORDER_COLOR = '#0f766e';

  function applyFloatingButtonStyles(button, options) {
    var styleOptions = options || {};

    button.type = 'button';
    button.style.position = 'fixed';
    button.style.top = styleOptions.top || '70px';
    button.style.right = styleOptions.right || '20px';
    button.style.zIndex = '999999';
    button.style.width = styleOptions.width || '168px';
    button.style.height = styleOptions.height || '40px';
    button.style.padding = '0 16px';
    button.style.background = styleOptions.background || '#1976d2';
    button.style.color = '#fff';
    button.style.border = '2px solid transparent';
    button.style.borderRadius = '10px';
    button.style.cursor = 'pointer';
    button.style.fontSize = '14px';
    button.style.fontFamily = 'Arial, sans-serif';
    button.style.fontWeight = '600';
    button.style.display = 'flex';
    button.style.alignItems = 'center';
    button.style.justifyContent = 'center';
    button.style.boxShadow = '0 4px 12px rgba(0,0,0,0.25)';
    button.style.boxSizing = 'border-box';
  }

  function createSearchButton() {
    var button = document.getElementById(APP_CONFIG.SEARCH_BUTTON_ID);
    if (button) {
      return button;
    }

    button = document.createElement('button');
    button.id = APP_CONFIG.SEARCH_BUTTON_ID;
    button.textContent = 'Поискать письма';
    applyFloatingButtonStyles(button, {
      right: '150px',
      top: '70px',
      background: '#1976d2'
    });
    button.addEventListener('click', onSearchButtonClick);
    document.body.appendChild(button);

    return button;
  }

  function createPlacementButton() {
    var button = document.getElementById(APP_CONFIG.PLACEMENT_BUTTON_ID);
    if (button) {
      return button;
    }

    button = document.createElement('button');
    button.id = APP_CONFIG.PLACEMENT_BUTTON_ID;
    button.textContent = 'Размесить заказ';
    button.style.display = 'none';
    applyFloatingButtonStyles(button, {
      right: '150px',
      top: '118px',
      background: '#1976d2'
    });
    button.addEventListener('click', onPlacementButtonClick);
    document.body.appendChild(button);

    return button;
  }

  function createTrackingButton() {
    var button = document.getElementById(APP_CONFIG.TRACKING_BUTTON_ID);
    if (button) {
      return button;
    }

    button = document.createElement('button');
    button.id = APP_CONFIG.TRACKING_BUTTON_ID;
    button.textContent = 'Проверить трекинг';
    button.style.display = 'none';
    applyFloatingButtonStyles(button, {
      right: '150px',
      top: APP_CONFIG.TRACKING_BUTTON_TOP,
      background: '#0f766e'
    });
    button.addEventListener('click', onTrackingButtonClick);
    document.body.appendChild(button);

    return button;
  }

  function getPlacementButtonBorderColor() {
    return createPlacementButton().style.display === 'none'
      ? 'transparent'
      : PLACEMENT_CONFIRMED_BORDER_COLOR;
  }

  function resetActionButtonBorders() {
    createSearchButton().style.borderColor = 'transparent';
    createPlacementButton().style.borderColor = getPlacementButtonBorderColor();
    createTrackingButton().style.borderColor =
      createTrackingButton().style.display === 'none' ? 'transparent' : TRACKING_BUTTON_BORDER_COLOR;
  }

  function setStatus(text, mode, buttonKey, options) {
    var badge = document.getElementById(APP_CONFIG.STATUS_ID);
    var targetButton;
    var statusMode = mode || 'neutral';

    if (buttonKey === 'placement') {
      targetButton = createPlacementButton();
    } else if (buttonKey === 'tracking') {
      targetButton = createTrackingButton();
    } else {
      targetButton = createSearchButton();
    }

    resetActionButtonBorders();
    if (badge) {
      badge.style.display = 'none';
      badge.textContent = '';
    }

    if (SEARCH_STATUS_BORDER_COLORS[statusMode]) {
      targetButton.style.borderColor = SEARCH_STATUS_BORDER_COLORS[statusMode];
    }
  }

  function createPanel() {
    var panel = document.getElementById(APP_CONFIG.PANEL_ID);
    if (panel) {
      return panel;
    }

    panel = document.createElement('div');
    panel.id = APP_CONFIG.PANEL_ID;
    panel.style.position = 'fixed';
    panel.style.top = APP_CONFIG.PANEL_TOP;
    panel.style.right = '20px';
    panel.style.width = '540px';
    panel.style.maxWidth = 'calc(100vw - 40px)';
    panel.style.maxHeight = '78vh';
    panel.style.overflowY = 'auto';
    panel.style.zIndex = '999999';
    panel.style.background = '#fff';
    panel.style.color = '#222';
    panel.style.border = '1px solid #d0d7de';
    panel.style.borderRadius = '12px';
    panel.style.boxShadow = '0 10px 28px rgba(0,0,0,0.25)';
    panel.style.fontSize = '14px';
    panel.style.fontFamily = 'Arial, sans-serif';
    panel.style.display = 'none';

    var header = document.createElement('div');
    header.style.display = 'flex';
    header.style.justifyContent = 'space-between';
    header.style.alignItems = 'center';
    header.style.padding = '12px 14px';
    header.style.borderBottom = '1px solid #e5e7eb';
    header.style.fontWeight = 'bold';
    header.style.fontSize = '15px';

    var title = document.createElement('div');
    title.className = APP_CONFIG.PANEL_TITLE_CLASS;
    title.textContent = 'Письма по заказу';

    var closeBtn = document.createElement('button');
    closeBtn.textContent = '×';
    closeBtn.type = 'button';
    closeBtn.style.border = 'none';
    closeBtn.style.background = 'transparent';
    closeBtn.style.fontSize = '22px';
    closeBtn.style.cursor = 'pointer';
    closeBtn.style.color = '#666';
    closeBtn.addEventListener('click', function () {
      panel.style.display = 'none';
      state.isPanelOpen = false;
    });

    var body = document.createElement('div');
    body.className = 'tm-ms-mail-search-body';
    body.style.padding = '12px 14px';

    header.appendChild(title);
    header.appendChild(closeBtn);
    panel.appendChild(header);
    panel.appendChild(body);
    document.body.appendChild(panel);

    return panel;
  }

  function setPanelTitle(title) {
    var panelTitle = createPanel().querySelector('.' + APP_CONFIG.PANEL_TITLE_CLASS);
    panelTitle.textContent = title || 'Письма по заказу';
  }

  function getPanelBody() {
    return createPanel().querySelector('.tm-ms-mail-search-body');
  }

  function setPanelHtml(html, title, panelMode) {
    var body = getPanelBody();

    if (title) {
      setPanelTitle(title);
    }

    if (panelMode) {
      state.activePanelMode = panelMode;
    }

    body.innerHTML = html;
    createPanel().style.display = 'block';
    state.isPanelOpen = true;
  }

  function hidePanel() {
    var panel = document.getElementById(APP_CONFIG.PANEL_ID);
    if (panel) {
      panel.style.display = 'none';
    }
    state.isPanelOpen = false;
  }

  function setPlacementButtonVisible(isVisible) {
    var button = createPlacementButton();

    button.style.display = isVisible ? 'block' : 'none';
    button.style.borderColor = isVisible ? PLACEMENT_CONFIRMED_BORDER_COLOR : 'transparent';
  }

  function setTrackingButtonVisible(isVisible) {
    var button = createTrackingButton();

    button.style.display = isVisible ? 'block' : 'none';
    button.style.borderColor = isVisible ? TRACKING_BUTTON_BORDER_COLOR : 'transparent';
  }

  function renderError(title, detailsHtml, panelTitle) {
    setPanelHtml(
      '<div style="color:#b00020;font-weight:bold;font-size:15px;margin-bottom:10px;">' +
        escapeHtml(title) +
        '</div>' +
        detailsHtml,
      panelTitle || 'Ошибка'
    );
  }

  function renderApiFailure(data, panelTitle) {
    renderError(
      'Ошибка',
      '<div style="margin-bottom:8px;">' +
        escapeHtml(data && data.error ? data.error : 'Неизвестная ошибка') +
        '</div>' +
        (data && data.stack
          ? '<pre style="white-space:pre-wrap;font-size:12px;color:#444;background:#f6f8fa;border:1px solid #e5e7eb;padding:10px;border-radius:8px;">' +
            escapeHtml(data.stack) +
            '</pre>'
          : ''),
      panelTitle
    );
  }

  function renderNonJsonResponse(text, requestUrl, status, panelTitle) {
    renderError(
      'Не удалось разобрать ответ',
      '<div style="margin-bottom:8px;"><b>URL:</b><br><span style="font-size:12px;color:#555;word-break:break-all;">' +
        escapeHtml(requestUrl) +
        '</span></div>' +
        '<div style="margin-bottom:8px;"><b>HTTP status:</b> ' +
        escapeHtml(String(status || '')) +
        '</div>' +
        '<div style="margin-bottom:8px;"><b>Первые 1000 символов ответа:</b></div>' +
        '<pre style="white-space:pre-wrap;font-size:12px;color:#444;background:#f6f8fa;border:1px solid #e5e7eb;padding:10px;border-radius:8px;">' +
        escapeHtml((text || '').slice(0, 1000)) +
        '</pre>',
      panelTitle
    );
  }

  function renderSettingsRequired(actionLabel) {
    var normalizedAction = actionLabel || 'работы';

    renderError(
      'Нужно заполнить доступ',
      '<div style="margin-bottom:10px;">Для ' +
        escapeHtml(normalizedAction) +
        ' нужен персональный токен сотрудника. Он хранится отдельно у каждого пользователя и не лежит в общем userscript.</div>' +
        '<div style="margin-bottom:10px;">Открой меню Tampermonkey и выбери <b>MoySklad: настроить доступ</b>, либо повтори действие и вставь свой токен.</div>',
      'Нужен доступ'
    );
  }

  function isInternalSearchHeaderEmail(email) {
    var normalizedEmail = String(email || '').trim().toLowerCase();
    var domain = normalizedEmail.split('@')[1] || '';

    return [
      'sparrowssons.com',
      'wiredtunes.pl',
      'united-music.by',
      'united-music.ru'
    ].indexOf(domain) !== -1;
  }

  function isLikelyTransportSearchHeaderEmail(email) {
    var normalizedEmail = String(email || '').trim().toLowerCase();
    var domain = normalizedEmail.split('@')[1] || '';

    return [
      'dbschenker.com',
      'schenker',
      'dsv.com',
      'dhl.com',
      'ups.com',
      'fedex.com',
      'tnt.com',
      'gls-',
      'baz-log.com',
      'cargoline',
      'cargo'
    ].some(function (hint) {
      return domain.indexOf(hint) !== -1;
    });
  }

  function isSavableSearchHeaderEmail(email) {
    var normalizedEmail = String(email || '').trim().toLowerCase();
    var localPart = normalizedEmail.split('@')[0] || '';

    if (!normalizedEmail || normalizedEmail.indexOf('@') === -1) {
      return false;
    }

    if (isInternalSearchHeaderEmail(normalizedEmail) || isLikelyTransportSearchHeaderEmail(normalizedEmail)) {
      return false;
    }

    if (/@(?:e\.)?moysklad\.ru$/i.test(normalizedEmail)) {
      return false;
    }

    if (/(^|[\W_])(no-?reply|noreply|postmaster|mailer-daemon|daemon|bounce|notification|notifications|robot|bot|do-not-reply)([\W_]|$)/i.test(localPart)) {
      return false;
    }

    return true;
  }

  function collectSearchSaveCandidateEmails(data) {
    var emails = [];
    var seen = {};

    if (data && Array.isArray(data.emails)) {
      data.emails.forEach(function (thread) {
        if (!thread || thread.threadCategory === 'transport') {
          return;
        }

        emails = emails
          .concat(normalizeRecipientList(thread.from || ''))
          .concat(normalizeRecipientList(thread.to || ''))
          .concat(normalizeRecipientList(thread.cc || ''));
      });
    }

    return emails
      .map(function (email) {
        return String(email || '').trim().toLowerCase();
      })
      .filter(function (email) {
        if (!email || seen[email] || !isSavableSearchHeaderEmail(email)) {
          return false;
        }

        seen[email] = true;
        return true;
      });
  }

  function renderSearchHeaderLine(label, rawValue, addableEmails) {
    var normalizedValue = String(rawValue || '').trim();
    var headerEmails = normalizeRecipientList(normalizedValue);
    var actionableEmails = headerEmails.filter(function (email) {
      return addableEmails.indexOf(email) !== -1;
    });
    var html = '<div style="font-size:12px;color:#666;margin-bottom:6px;"><b>' + escapeHtml(label) + '</b> ' + escapeHtml(normalizedValue) + '</div>';

    if (!actionableEmails.length) {
      return html;
    }

    html += '<div style="display:flex;flex-wrap:wrap;gap:6px 10px;margin:-2px 0 8px 38px;">';

    actionableEmails.forEach(function (email) {
      html += '<span style="display:inline-flex;align-items:center;gap:6px;">';
      html += '<span style="display:inline-flex;align-items:center;padding:2px 8px;border-radius:999px;background:#fff7ed;color:#9a3412;border:1px solid #fdba74;font-size:12px;font-weight:bold;">' + escapeHtml(email) + '</span>';
      html += '<button class="tm-ms-search-email-action-btn" data-email="' + escapeHtml(email) + '" data-action="ms" type="button" style="border:none;background:none;padding:0;color:#1976d2;cursor:pointer;font-size:12px;font-weight:bold;text-decoration:underline;">Add email to MS</button>';
      html += '<button class="tm-ms-search-email-action-btn" data-email="' + escapeHtml(email) + '" data-action="transport" type="button" style="border:none;background:none;padding:0;color:#1976d2;cursor:pointer;font-size:12px;font-weight:bold;text-decoration:underline;">to transport list</button>';
      html += '<button class="tm-ms-search-email-action-btn" data-email="' + escapeHtml(email) + '" data-action="ignore" type="button" style="border:none;background:none;padding:0;color:#1976d2;cursor:pointer;font-size:12px;font-weight:bold;text-decoration:underline;">to ignore list</button>';
      html += '</span>';
    });

    html += '</div>';
    return html;
  }

  function renderSearchEmailCard(email, settings, addableEmails) {
    var fixedLink = forceGmailAccount(email.link, settings);
    var html = '';

    html += '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px;background:#fff;">';
    html += '<div style="font-weight:bold;font-size:14px;margin-bottom:8px;">' + escapeHtml(email.subject || '(без темы)') + '</div>';
    html += renderSearchHeaderLine('От:', email.from || '', addableEmails);

    if (email.to) {
      html += renderSearchHeaderLine('Кому:', email.to, addableEmails);
    }

    if (email.cc) {
      html += renderSearchHeaderLine('Копия:', email.cc, addableEmails);
    }

    html += '<div style="font-size:12px;color:#666;margin-bottom:10px;"><b>Дата:</b> ' + escapeHtml(formatDate(email.date)) + '</div>';
    html += '<div style="font-size:13px;line-height:1.45;margin-bottom:10px;color:#222;">' + escapeHtml(email.snippet || '') + '</div>';
    html += '<div style="font-size:12px;color:#666;margin-bottom:10px;"><b>Сообщений в треде:</b> ' + escapeHtml(email.messageCount || '') + '</div>';
    html += '<div><a href="' + escapeHtml(fixedLink) + '" target="_blank" style="color:#1976d2;text-decoration:none;font-weight:bold;">Открыть письмо</a></div>';
    html += '</div>';

    return html;
  }

  function renderEmails(data, sourceLabel, settings) {
    var html = '';
    var suggestions = data && data.supplierEmailSuggestions;
    var suggestionEmails = suggestions && Array.isArray(suggestions.suggestedEmails)
      ? suggestions.suggestedEmails
      : [];
    var hasSupplierEmail = Boolean(String(data && data.supplierEmail || '').trim());
    var addableEmails = collectSearchSaveCandidateEmails(data);
    var primaryEmails = data && Array.isArray(data.emailsPrimary)
      ? data.emailsPrimary
      : ((data && Array.isArray(data.emails)) ? data.emails.filter(function (email) {
          return !email || email.threadCategory !== 'transport';
        }) : []);
    var transportEmails = data && Array.isArray(data.emailsTransport)
      ? data.emailsTransport
      : ((data && Array.isArray(data.emails)) ? data.emails.filter(function (email) {
          return email && email.threadCategory === 'transport';
        }) : []);
    var primaryCount = data && data.countsByCategory && typeof data.countsByCategory.primary === 'number'
      ? data.countsByCategory.primary
      : primaryEmails.length;
    var transportCount = data && data.countsByCategory && typeof data.countsByCategory.transport === 'number'
      ? data.countsByCategory.transport
      : transportEmails.length;
    var defaultFilter = primaryCount > 0 ? 'primary' : (transportCount > 0 ? 'transport' : 'primary');

    if (!data || !data.success) {
      renderApiFailure(data, 'Письма по заказу');
      return;
    }

    html += '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:14px;background:#f9fafb;">';
    html += '<div style="margin-bottom:6px;"><b>Номер заказа:</b> ' + escapeHtml(data.orderNumber || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Поставщик:</b> ' + escapeHtml(data.supplierName || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Email поставщика:</b> ' + escapeHtml(data.supplierEmail || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Домен:</b> ' + escapeHtml(data.supplierDomain || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Режим поиска:</b> ' + escapeHtml(data.searchMode || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Из кэша заказа:</b> ' + escapeHtml(data.fromCache ? 'yes' : 'no') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Сотрудник:</b> ' + escapeHtml(data.userName || '') + ' (' + escapeHtml(data.userCode || '') + ')</div>';
    html += '<div style="margin-bottom:6px;"><b>Режим доступа:</b> ' + escapeHtml(data.authMode || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Источник ответа:</b> ' + escapeHtml(sourceLabel || 'manual') + '</div>';

    if (Array.isArray(data.searchAttempts) && data.searchAttempts.length) {
      html += '<div style="margin-top:10px;"><b>Попытки поиска:</b></div>';
      html += '<ul style="margin:6px 0 0 18px;padding:0;">';

      data.searchAttempts.forEach(function (attempt) {
        html += '<li style="margin-bottom:4px;color:#555;">' + escapeHtml(attempt) + '</li>';
      });

      html += '</ul>';
    }

    html += '</div>';
    html += '<div id="tm-ms-search-save-note" style="margin:-4px 0 12px 0;font-size:12px;color:#555;"></div>';

    if (!hasSupplierEmail && suggestionEmails.length) {
      html += '<div style="border:1px solid #fdba74;border-radius:10px;padding:12px;margin-bottom:14px;background:#fff7ed;">';
      html += '<div style="font-weight:bold;font-size:14px;color:#9a3412;margin-bottom:8px;">Вероятные email поставщика из переписки</div>';
      html += '<div style="font-size:12px;color:#7c2d12;margin-bottom:10px;">Если в карточке контрагента пусто, эти адреса можно использовать как черновые получатели для размещения.</div>';
      html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:' + (suggestions.candidates && suggestions.candidates.length ? '10px' : '0') + ';">';

      suggestionEmails.forEach(function (email) {
        html += '<span style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;background:#ffedd5;color:#9a3412;font-size:12px;font-weight:bold;">' + escapeHtml(email) + '</span>';
      });

      html += '</div>';

      if (Array.isArray(suggestions.candidates) && suggestions.candidates.length) {
        html += '<div style="font-size:12px;color:#7c2d12;">';

        suggestions.candidates.forEach(function (candidate) {
          html += '<div style="margin-bottom:4px;"><b>' + escapeHtml(candidate.email) + '</b>: ' + escapeHtml(candidate.reasonSummary || '') + '</div>';
        });

        html += '</div>';
      }

      html += '</div>';
    }

    if (!primaryCount && !transportCount) {
      html += '<div style="padding:10px 0;">Письма не найдены</div>';
      setPanelHtml(html, 'Письма по заказу', 'search');
      return;
    }

    html += '<div style="margin-bottom:12px;color:#555;">';
    html += '<button class="tm-ms-search-filter-btn" data-filter="primary" type="button" style="border:none;background:none;padding:0;color:' + (defaultFilter === 'primary' ? '#1d4ed8' : '#1976d2') + ';cursor:pointer;font-size:14px;font-weight:' + (defaultFilter === 'primary' ? 'bold' : 'normal') + ';text-decoration:underline;">Писем по поставщику: ' + escapeHtml(primaryCount) + '</button>';
    html += ', ';
    html += '<button class="tm-ms-search-filter-btn" data-filter="transport" type="button" style="border:none;background:none;padding:0;color:' + (defaultFilter === 'transport' ? '#1d4ed8' : '#1976d2') + ';cursor:pointer;font-size:14px;font-weight:' + (defaultFilter === 'transport' ? 'bold' : 'normal') + ';text-decoration:underline;">писем по транспорту: ' + escapeHtml(transportCount) + '</button>';
    html += '</div>';

    html += '<div class="tm-ms-search-email-group" data-email-category="primary" style="display:' + (defaultFilter === 'primary' ? 'block' : 'none') + ';">';
    if (primaryEmails.length) {
      primaryEmails.forEach(function (email) {
        html += renderSearchEmailCard(email, settings, addableEmails);
      });
    } else {
      html += '<div style="padding:10px 0;color:#666;">Писем по поставщику не найдено.</div>';
    }
    html += '</div>';

    html += '<div class="tm-ms-search-email-group" data-email-category="transport" style="display:' + (defaultFilter === 'transport' ? 'block' : 'none') + ';">';
    if (transportEmails.length) {
      transportEmails.forEach(function (email) {
        html += renderSearchEmailCard(email, settings, addableEmails);
      });
    } else {
      html += '<div style="padding:10px 0;color:#666;">Писем по транспорту не найдено.</div>';
    }
    html += '</div>';

    setPanelHtml(html, 'Письма по заказу', 'search');
  }

  function renderPlacementPanel(data) {
    var recipients = Array.isArray(data && data.prefillEmails)
      ? data.prefillEmails.join(', ')
      : (Array.isArray(data && data.emails) ? data.emails.join(', ') : '');
    var normalizedRecipients = normalizeRecipientList(recipients);
    var selectedSuggestedEmail = normalizedRecipients.length ? normalizedRecipients[0] : '';
    var subject = buildPlacementEmailSubject(data);
    var body = buildPlacementEmailBody(data);
    var attachmentFileName = String((data && data.attachmentFileName) || 'PO.xls');
    var statusButtonDisabled = !state.placementEmailSent;
    var sendButtonDisabled = state.placementEmailSent;
    var draftButtonDisabled = state.placementEmailSent;
    var draftButtonLabel = state.placementDraftId ? 'Обновить черновик' : 'Сохранить черновик';
    var useGmailSuggestions = Boolean(data && data.useGmailSuggestions);
    var gmailSuggestedEmails = data && Array.isArray(data.gmailSuggestedEmails)
      ? data.gmailSuggestedEmails
      : [];
    var gmailSuggestionCandidates = data && Array.isArray(data.gmailSuggestionCandidates)
      ? data.gmailSuggestionCandidates
      : [];
    var html = '';

    if (!data || !data.success) {
      renderApiFailure(data, 'Размещение PO');
      return;
    }

    html += '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:14px;background:#f9fafb;">';
    html += '<div style="margin-bottom:6px;"><b>Номер заказа:</b> ' + escapeHtml(data.orderNumber || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Поставщик:</b> ' + escapeHtml(data.supplierName || '') + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Текущий статус:</b> <span id="tm-ms-placement-state-value">' + escapeHtml(data.currentStateName || '') + '</span></div>';
    html += '<div style="margin-bottom:0;"><b>Шаблон:</b> ' + escapeHtml(data.templateName || '') + ' (' + escapeHtml('xls') + ')</div>';
    html += '</div>';

    if (!data.canPlace) {
      html += '<div style="padding:12px;border:1px solid #fecaca;border-radius:10px;background:#fef2f2;color:#991b1b;">Размещение доступно только для заказов в статусе <b>Подтвержден</b>.</div>';
      setPanelHtml(html, 'Размещение PO', 'placement');
      return;
    }

    html += '<div style="display:flex;flex-direction:column;gap:14px;">';

    html += '<section style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fff;">';
    html += '<div style="font-weight:bold;margin-bottom:8px;">1. Вложение</div>';
    html += '<div style="margin-bottom:6px;"><b>Файл:</b> ' + escapeHtml(attachmentFileName) + '</div>';
    html += '<div style="font-size:12px;color:#666;margin-bottom:10px;">При отправке Apps Script сам сформирует XLS и приложит его к письму. При желании можно скачать копию для проверки.</div>';
    html += '<button id="tm-ms-placement-download-btn" type="button" style="padding:9px 12px;border:none;border-radius:8px;background:#1976d2;color:#fff;cursor:pointer;font-weight:bold;">Скачать вложение</button>';
    html += '<div id="tm-ms-placement-download-note" style="margin-top:10px;font-size:12px;color:#555;"></div>';
    html += '</section>';

    html += '<section style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fff;">';
    html += '<div style="font-weight:bold;margin-bottom:8px;">2. Подготовить письмо</div>';

    if (useGmailSuggestions) {
      html += '<div style="border:1px solid #fdba74;border-radius:10px;padding:12px;margin-bottom:12px;background:#fff7ed;">';
      html += '<div style="font-weight:bold;color:#9a3412;margin-bottom:6px;">Suggestions из переписки: в карточке контрагента нет email</div>';
      html += '<div style="font-size:13px;color:#7c2d12;margin-bottom:10px;">Ниже адреса, найденные по истории переписки. Клик по адресу делает его основным email и подставляет первым в поле <b>Кому</b>.</div>';
      html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:' + (gmailSuggestionCandidates.length ? '10px' : '12px') + ';">';

      gmailSuggestedEmails.forEach(function (email) {
        var isSelected = selectedSuggestedEmail && selectedSuggestedEmail === String(email || '').trim().toLowerCase();
        html += '<button class="tm-ms-placement-suggested-email-btn" data-email="' + escapeHtml(email) + '" type="button" style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;border:1px solid ' + (isSelected ? '#9a3412' : '#fdba74') + ';background:' + (isSelected ? '#c2410c' : '#ffedd5') + ';color:' + (isSelected ? '#fff' : '#9a3412') + ';font-size:12px;font-weight:bold;cursor:pointer;">' + escapeHtml(isSelected ? 'Основной: ' + email : email) + '</button>';
      });

      html += '</div>';

      if (gmailSuggestionCandidates.length) {
        html += '<div style="font-size:12px;color:#7c2d12;margin-bottom:10px;">';

        gmailSuggestionCandidates.forEach(function (candidate) {
          html += '<div style="margin-bottom:4px;"><b>' + escapeHtml(candidate.email) + '</b>: ' + escapeHtml(candidate.reasonSummary || '') + '</div>';
        });

        html += '</div>';
      }

      html += '<div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">';
      html += '<button id="tm-ms-placement-save-counterparty-btn" type="button" style="padding:9px 12px;border:1px solid #c2410c;border-radius:8px;background:#fff;color:#9a3412;cursor:pointer;font-weight:bold;">Добавить email в МойСклад</button>';

      if (data.counterpartyLink) {
        html += '<a href="' + escapeHtml(data.counterpartyLink) + '" target="_blank" rel="noopener noreferrer" style="font-size:12px;color:#9a3412;text-decoration:none;font-weight:bold;">Открыть контрагента</a>';
      }

      html += '</div>';
      html += '<div id="tm-ms-placement-suggestion-note" style="margin-top:10px;font-size:12px;color:#7c2d12;">Если поле email контрагента пустое, адреса попадут туда. Если уже занято, адреса сохранятся в контактное лицо.</div>';
      html += '</div>';
    }

    html += '<label style="display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;">Кому</label>';
    html += '<textarea id="tm-ms-placement-to" rows="3" style="width:100%;box-sizing:border-box;border:1px solid #d1d5db;border-radius:8px;padding:8px 10px;font:inherit;resize:vertical;margin-bottom:10px;">' + escapeHtml(recipients) + '</textarea>';

    if (useGmailSuggestions && gmailSuggestedEmails.length) {
      html += '<div style="margin:-2px 0 10px 0;padding:10px 12px;border:1px dashed #fdba74;border-radius:8px;background:#fffaf5;">';
      html += '<div style="font-size:12px;font-weight:bold;color:#9a3412;margin-bottom:6px;">Найдено в переписке:</div>';
      html += '<div style="display:flex;flex-wrap:wrap;gap:6px;">';

      gmailSuggestedEmails.forEach(function (email) {
        var isSelectedInline = selectedSuggestedEmail && selectedSuggestedEmail === String(email || '').trim().toLowerCase();
        html += '<button class="tm-ms-placement-suggested-email-btn" data-email="' + escapeHtml(email) + '" type="button" style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;border:1px solid ' + (isSelectedInline ? '#9a3412' : '#fdba74') + ';background:' + (isSelectedInline ? '#c2410c' : '#ffedd5') + ';color:' + (isSelectedInline ? '#fff' : '#9a3412') + ';font-size:12px;font-weight:bold;cursor:pointer;">' + escapeHtml(isSelectedInline ? 'Основной: ' + email : email) + '</button>';
      });

      html += '</div>';
      html += '<div style="font-size:11px;color:#7c2d12;margin-top:6px;">Клик по адресу делает его основным получателем для размещения.</div>';
      html += '</div>';
    }

    html += '<div style="font-size:12px;color:#666;margin:-4px 0 10px 0;">Можно удалить лишние адреса перед отправкой или сохранением черновика.</div>';
    html += '<label style="display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;">Тема</label>';
    html += '<input id="tm-ms-placement-subject" type="text" value="' + escapeHtml(subject) + '" style="width:100%;box-sizing:border-box;border:1px solid #d1d5db;border-radius:8px;padding:8px 10px;font:inherit;margin-bottom:10px;" />';
    html += '<label style="display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;">Текст</label>';
    html += '<textarea id="tm-ms-placement-body" rows="7" style="width:100%;box-sizing:border-box;border:1px solid #d1d5db;border-radius:8px;padding:8px 10px;font:inherit;resize:vertical;">' + escapeHtml(body) + '</textarea>';
    html += '</section>';

    html += '<section style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fff;">';
    html += '<div style="font-weight:bold;margin-bottom:8px;">3. Отправить или сохранить черновик</div>';
    html += '<div style="font-size:12px;color:#666;margin-bottom:10px;">Письмо уйдет через Apps Script от имени того Gmail-аккаунта, под которым сейчас работает этот скрипт. Вместо отправки можно сначала сохранить полноценный черновик в Gmail.</div>';
    html += '<div style="display:flex;flex-wrap:wrap;gap:8px;">';
    html += '<button id="tm-ms-placement-send-btn" type="button" ' + (sendButtonDisabled ? 'disabled ' : '') + 'style="padding:9px 12px;border:none;border-radius:8px;background:#15803d;color:#fff;cursor:' + (sendButtonDisabled ? 'default' : 'pointer') + ';font-weight:bold;opacity:' + (sendButtonDisabled ? '0.7' : '1') + ';">' + (sendButtonDisabled ? 'Письмо отправлено' : 'Отправить письмо') + '</button>';
    html += '<button id="tm-ms-placement-draft-btn" type="button" ' + (draftButtonDisabled ? 'disabled ' : '') + 'style="padding:9px 12px;border:1px solid #1d4ed8;border-radius:8px;background:#eff6ff;color:#1d4ed8;cursor:' + (draftButtonDisabled ? 'default' : 'pointer') + ';font-weight:bold;opacity:' + (draftButtonDisabled ? '0.7' : '1') + ';">' + escapeHtml(draftButtonLabel) + '</button>';
    html += '</div>';
    html += '<div id="tm-ms-placement-send-note" style="margin-top:10px;font-size:12px;color:#555;">' + escapeHtml(sendButtonDisabled ? 'Письмо уже отправлено в этой сессии. Можно подтверждать статус.' : 'Проверь получателей, тему, текст и вложение перед отправкой или сохранением черновика.') + '</div>';
    html += '</section>';

    html += '<section style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fff;">';
    html += '<div style="font-weight:bold;margin-bottom:8px;">4. Подтвердить размещение</div>';
    html += '<div style="font-size:12px;color:#666;margin-bottom:10px;">Статус меняется отдельным кликом уже после успешной отправки письма.</div>';
    html += '<button id="tm-ms-placement-state-btn" type="button" ' + (statusButtonDisabled ? 'disabled ' : '') + 'style="padding:9px 12px;border:none;border-radius:8px;background:#b45309;color:#fff;cursor:' + (statusButtonDisabled ? 'default' : 'pointer') + ';font-weight:bold;opacity:' + (statusButtonDisabled ? '0.7' : '1') + ';">Поставить статус "Размещен"</button>';
    html += '<div id="tm-ms-placement-state-note" style="margin-top:10px;font-size:12px;color:#555;">' + escapeHtml(statusButtonDisabled ? 'Кнопка станет активной после успешной отправки письма.' : 'Письмо отправлено, можно ставить статус.') + '</div>';
    html += '</section>';

    html += '</div>';

    setPanelHtml(html, 'Размещение PO', 'placement');

    getPanelBody().querySelector('#tm-ms-placement-download-btn').addEventListener('click', onPlacementDownloadClick);
    getPanelBody().querySelector('#tm-ms-placement-send-btn').addEventListener('click', onPlacementSendEmailClick);
    getPanelBody().querySelector('#tm-ms-placement-draft-btn').addEventListener('click', onPlacementSaveDraftClick);
    getPanelBody().querySelector('#tm-ms-placement-state-btn').addEventListener('click', onPlacementSetStateClick);

    if (useGmailSuggestions) {
      getPanelBody().querySelectorAll('.tm-ms-placement-suggested-email-btn').forEach(function (button) {
        button.addEventListener('click', onPlacementSuggestedEmailClick);
      });
      getPanelBody().querySelector('#tm-ms-placement-to').addEventListener('input', syncPlacementSuggestedEmailSelectionFromTextarea);
      getPanelBody().querySelector('#tm-ms-placement-save-counterparty-btn').addEventListener('click', onPlacementSaveCounterpartyEmailsClick);
    }
  }

  function renderTrackingPanel(data) {
    var entries = data && Array.isArray(data.entries) ? data.entries : [];
    var sourceLabel = normalizeWhitespace(data && data.sourceLabel) || 'Трекинг номер';
    var html = '';

    html += '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:14px;background:#f9fafb;">';
    html += '<div style="margin-bottom:6px;"><b>Поле заказа:</b> ' + escapeHtml(sourceLabel) + '</div>';
    html += '<div style="margin-bottom:6px;"><b>Значение:</b> ' + escapeHtml(data && data.rawFieldValue ? data.rawFieldValue : '') + '</div>';
    html += '<div style="margin-bottom:0;"><b>Распознано треков:</b> ' + escapeHtml(String(entries.length)) + '</div>';
    html += '</div>';

    if (!entries.length) {
      html += '<div style="padding:12px;border:1px solid #fecaca;border-radius:10px;background:#fef2f2;color:#991b1b;">Не удалось распознать трек-номер из поля заказа <b>' + escapeHtml(sourceLabel) + '</b>.</div>';
      setPanelHtml(html, 'Трекинг поставки', 'tracking');
      return;
    }

    entries.forEach(function (entry, index) {
      var history = Array.isArray(entry.history) ? entry.history : [];
      var statusColor = entry.success ? '#166534' : '#92400e';

      html += '<section style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px;background:#fff;">';
      html += '<div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:8px;">';
      html += '<div>';
      html += '<div style="font-weight:bold;font-size:15px;margin-bottom:4px;">' + escapeHtml(entry.trackingNumber || '') + '</div>';
      html += '<div style="font-size:12px;color:#666;">' + escapeHtml(entry.carrierLabel || '') + '</div>';
      html += '</div>';
      html += '<div style="font-size:12px;color:' + statusColor + ';font-weight:bold;text-align:right;">' + escapeHtml(entry.success ? 'Статус загружен' : 'Официальный fallback') + '</div>';
      html += '</div>';

      if (entry.currentStatus) {
        html += '<div style="margin-bottom:8px;"><b>Текущий статус:</b> <span style="color:' + statusColor + ';">' + escapeHtml(entry.currentStatus) + '</span></div>';
      }

      if (entry.error) {
        html += '<div style="margin-bottom:8px;font-size:12px;color:#7c2d12;">' + escapeHtml(entry.error) + '</div>';
      }

      if (history.length) {
        html += '<div style="font-weight:bold;margin:10px 0 8px 0;">История</div>';
        html += '<div style="display:flex;flex-direction:column;gap:8px;">';

        history.forEach(function (event) {
          html += '<div style="border-left:3px solid #99f6e4;padding-left:10px;">';
          html += '<div style="font-size:12px;color:#666;margin-bottom:2px;">' + escapeHtml(formatTrackingEventDate(event.date)) + '</div>';
          html += '<div style="font-size:13px;font-weight:bold;color:#134e4a;margin-bottom:' + (event.description || event.location ? '2px' : '0') + ';">' + escapeHtml(event.title || 'Событие') + '</div>';

          if (event.description) {
            html += '<div style="font-size:12px;color:#444;margin-bottom:' + (event.location ? '2px' : '0') + ';">' + escapeHtml(event.description) + '</div>';
          }

          if (event.location) {
            html += '<div style="font-size:12px;color:#666;">' + escapeHtml(event.location) + '</div>';
          }

          html += '</div>';
        });

        html += '</div>';
      }

      html += '<div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px;">';

      if (entry.officialUrl) {
        html += '<a href="' + escapeHtml(entry.officialUrl) + '" target="_blank" rel="noopener noreferrer" style="display:inline-flex;align-items:center;padding:8px 12px;border-radius:8px;background:#0f766e;color:#fff;text-decoration:none;font-weight:bold;">Открыть официальный трекинг</a>';
      }

      if (entry.sourceUrl && entry.sourceUrl !== entry.officialUrl) {
        html += '<a href="' + escapeHtml(entry.sourceUrl) + '" target="_blank" rel="noopener noreferrer" style="font-size:12px;color:#0f766e;text-decoration:none;font-weight:bold;">Источник интеграции</a>';
      }

      html += '</div>';
      html += '</section>';
    });

    if (data && data.showSupportNote) {
      html += '<div style="padding:12px;border:1px solid #bfdbfe;border-radius:10px;background:#eff6ff;color:#1d4ed8;font-size:12px;">UPS и DHL в этой версии работают через официальный fallback, потому что их стабильное API-получение статуса обычно требует отдельную developer-подписку или ключ.</div>';
    }

    setPanelHtml(html, 'Трекинг поставки', 'tracking');
  }

  function resetOrderState(orderId) {
    state.currentOrderId = orderId;
    state.searchPrefetchPromise = null;
    state.searchPrefetchResult = null;
    state.searchPrefetchError = null;
    state.searchPrefetchConsumed = false;
    state.searchPrefetchStartedAt = null;
    state.placementMetaPromise = null;
    state.placementMetaResult = null;
    state.placementMetaError = null;
    state.lastPlacementDownloadUrl = '';
    state.placementDraftId = '';
    state.placementEmailSent = false;
    state.trackingFieldValue = '';
    state.trackingSourceType = '';
    state.trackingSourceLabel = '';
    state.trackingSourceId = '';
    state.trackingSourceHref = '';
    state.trackingEntries = [];
    state.trackingLastResult = null;
  }

  function setPlacementMessage(selector, html, color) {
    var element = getPanelBody().querySelector(selector);

    if (!element) {
      return;
    }

    element.innerHTML = html || '';

    if (color) {
      element.style.color = color;
    }
  }

  function setSearchSaveMessage(html, color) {
    var element = getPanelBody().querySelector('#tm-ms-search-save-note');

    if (!element) {
      return;
    }

    element.innerHTML = html || '';

    if (color) {
      element.style.color = color;
    }
  }

  function setSearchEmailFilter(activeFilter) {
    var normalizedFilter = activeFilter === 'transport' ? 'transport' : 'primary';

    Array.prototype.slice.call(
      getPanelBody().querySelectorAll('.tm-ms-search-email-group')
    ).forEach(function (group) {
      var filter = String(group.getAttribute('data-email-category') || '').trim().toLowerCase();
      group.style.display = filter === normalizedFilter ? 'block' : 'none';
    });

    Array.prototype.slice.call(
      getPanelBody().querySelectorAll('.tm-ms-search-filter-btn')
    ).forEach(function (button) {
      var filter = String(button.getAttribute('data-filter') || '').trim().toLowerCase();
      var isActive = filter === normalizedFilter;
      button.style.color = isActive ? '#1d4ed8' : '#1976d2';
      button.style.fontWeight = isActive ? 'bold' : 'normal';
    });
  }

  function setPlacementStateButtonState(disabled, label) {
    var button = getPanelBody().querySelector('#tm-ms-placement-state-btn');

    if (!button) {
      return;
    }

    button.disabled = Boolean(disabled);
    button.style.opacity = disabled ? '0.7' : '1';
    button.style.cursor = disabled ? 'default' : 'pointer';

    if (label) {
      button.textContent = label;
    }
  }

  function setPlacementSendButtonState(disabled, label) {
    var button = getPanelBody().querySelector('#tm-ms-placement-send-btn');

    if (!button) {
      return;
    }

    button.disabled = Boolean(disabled);
    button.style.opacity = disabled ? '0.7' : '1';
    button.style.cursor = disabled ? 'default' : 'pointer';

    if (label) {
      button.textContent = label;
    }
  }

  function setPlacementDraftButtonState(disabled, label) {
    var button = getPanelBody().querySelector('#tm-ms-placement-draft-btn');

    if (!button) {
      return;
    }

    button.disabled = Boolean(disabled);
    button.style.opacity = disabled ? '0.7' : '1';
    button.style.cursor = disabled ? 'default' : 'pointer';

    if (label) {
      button.textContent = label;
    }
  }

  function setPlacementCounterpartySaveButtonState(disabled, label) {
    var button = getPanelBody().querySelector('#tm-ms-placement-save-counterparty-btn');

    if (!button) {
      return;
    }

    button.disabled = Boolean(disabled);
    button.style.opacity = disabled ? '0.7' : '1';
    button.style.cursor = disabled ? 'default' : 'pointer';

    if (label) {
      button.textContent = label;
    }
  }

  function getSearchEmailActionButtons(targetEmail, actionName) {
    var normalizedTargetEmail = String(targetEmail || '').trim().toLowerCase();
    var normalizedActionName = String(actionName || '').trim().toLowerCase();

    return Array.prototype.slice.call(
      getPanelBody().querySelectorAll('.tm-ms-search-email-action-btn')
    ).filter(function (button) {
      var buttonAction = String(button.getAttribute('data-action') || '').trim().toLowerCase();
      return String(button.getAttribute('data-email') || '').trim().toLowerCase() === normalizedTargetEmail &&
        buttonAction === normalizedActionName;
    });
  }

  function setSearchEmailActionButtonState(targetEmail, actionName, disabled, label) {
    getSearchEmailActionButtons(targetEmail, actionName).forEach(function (button) {
      button.disabled = Boolean(disabled);
      button.style.opacity = disabled ? '0.7' : '1';
      button.style.cursor = disabled ? 'default' : 'pointer';

      if (label) {
        button.textContent = label;
      }
    });
  }

  function getSearchEmailActionDefaultLabel(actionName) {
    var normalizedActionName = String(actionName || '').trim().toLowerCase();

    if (normalizedActionName === 'transport') {
      return 'to transport list';
    }

    if (normalizedActionName === 'ignore') {
      return 'to ignore list';
    }

    return 'Add email to MS';
  }

  function getSearchEmailActionLoadingLabel(actionName) {
    var normalizedActionName = String(actionName || '').trim().toLowerCase();

    if (normalizedActionName === 'transport') {
      return 'adding...';
    }

    if (normalizedActionName === 'ignore') {
      return 'adding...';
    }

    return 'saving...';
  }

  function getPlacementDraftButtonLabel() {
    return state.placementDraftId ? 'Обновить черновик' : 'Сохранить черновик';
  }

  function getPlacementMoySkladSaveButtonLabel() {
    return 'Добавить email в МойСклад';
  }

  function updatePlacementStateValue(text) {
    var valueNode = getPanelBody().querySelector('#tm-ms-placement-state-value');

    if (valueNode) {
      valueNode.textContent = text || '';
    }
  }

  function syncPlacementButtonWithResult(orderId, result) {
    if (state.currentOrderId !== orderId) {
      return;
    }

    setPlacementButtonVisible(Boolean(result && result.ok && result.data && result.data.success && result.data.canPlace));
  }

  function syncTrackingButtonVisibility() {
    if (!isPurchaseOrderPage()) {
      setTrackingButtonVisible(false);
      return [];
    }

    setTrackingButtonVisible(Boolean(getCurrentTrackingEntries().length));
    return state.trackingEntries.slice();
  }

  function syncSavedSupplierEmailsState(resultData) {
    var savedEmails = Array.isArray(resultData && resultData.emails) ? resultData.emails.slice() : [];

    if (state.placementMetaResult && state.placementMetaResult.data) {
      state.placementMetaResult.data.emails = savedEmails.slice();
      state.placementMetaResult.data.counterpartyEmail = resultData.counterpartyEmail || '';
      state.placementMetaResult.data.counterpartyFieldEmails = Array.isArray(resultData.counterpartyFieldEmails)
        ? resultData.counterpartyFieldEmails.slice()
        : [];
      state.placementMetaResult.data.contactPersonEmails = Array.isArray(resultData.contactPersonEmails)
        ? resultData.contactPersonEmails.slice()
        : [];
      state.placementMetaResult.data.prefillEmails = savedEmails.slice();
      state.placementMetaResult.data.useGmailSuggestions = false;
      state.placementMetaResult.data.gmailSuggestionCandidates = [];
      state.placementMetaResult.data.gmailSuggestedEmails = [];
    }
  }

  function requestPlacementMeta(orderId, settings) {
    state.placementMetaPromise = fetchPlacementMeta(orderId, settings)
      .then(function (result) {
        if (state.currentOrderId !== orderId) {
          return result;
        }

        state.placementMetaResult = result;
        state.placementMetaError = null;
        syncPlacementButtonWithResult(orderId, result);
        syncTrackingButtonVisibility();
        return result;
      })
      .catch(function (error) {
        if (state.currentOrderId === orderId) {
          state.placementMetaError = error;
          state.placementMetaResult = null;
          setPlacementButtonVisible(false);
          setTrackingButtonVisible(false);
        }

        throw error;
      });

    return state.placementMetaPromise;
  }

  function startPlacementMetaPrefetch(orderId) {
    var settings = ensureUserSettings({ silent: true });

    if (!orderId || !settings) {
      setPlacementButtonVisible(false);
      return null;
    }

    if (state.currentOrderId !== orderId) {
      resetOrderState(orderId);
    }

    if (state.placementMetaPromise || state.placementMetaResult) {
      return state.placementMetaPromise || Promise.resolve(state.placementMetaResult);
    }

    return requestPlacementMeta(orderId, settings);
  }

  async function loadPlacementMeta(orderId, settings, options) {
    var requestOptions = options || {};

    if (state.currentOrderId !== orderId) {
      resetOrderState(orderId);
    }

    if (requestOptions.forceRefresh) {
      state.placementMetaPromise = null;
      state.placementMetaResult = null;
      state.placementMetaError = null;
    }

    if (state.placementMetaResult && !requestOptions.forceRefresh) {
      return state.placementMetaResult;
    }

    if (state.placementMetaPromise && !requestOptions.forceRefresh) {
      return state.placementMetaPromise;
    }

    return requestPlacementMeta(orderId, settings);
  }

  async function loadSearchDataForPlacement(orderId, settings) {
    if (state.currentOrderId === orderId) {
      if (state.searchPrefetchResult) {
        return state.searchPrefetchResult;
      }

      if (state.searchPrefetchPromise) {
        return state.searchPrefetchPromise;
      }
    }

    return fetchSearchData(orderId, settings, {
      skipLog: true,
      prefetch: false
    });
  }

  async function buildPlacementPanelData(orderId, settings, placementData) {
    var panelData = Object.assign({}, placementData || {});
    var baseEmails = Array.isArray(panelData.emails) ? panelData.emails.slice() : [];
    var searchResult;
    var suggestions;
    var suggestedEmails;

    panelData.prefillEmails = baseEmails.slice();
    panelData.useGmailSuggestions = false;
    panelData.gmailSuggestionCandidates = [];
    panelData.gmailSuggestedEmails = [];

    if (baseEmails.length) {
      return panelData;
    }

    try {
      searchResult = await loadSearchDataForPlacement(orderId, settings);
    } catch (error) {
      return panelData;
    }

    if (!searchResult || !searchResult.ok || !searchResult.data || !searchResult.data.success) {
      return panelData;
    }

    suggestions = searchResult.data.supplierEmailSuggestions || {};
    suggestedEmails = Array.isArray(suggestions.suggestedEmails)
      ? suggestions.suggestedEmails.slice()
      : [];

    panelData.gmailSuggestionCandidates = Array.isArray(suggestions.candidates)
      ? suggestions.candidates
      : [];
    panelData.gmailSuggestedEmails = suggestedEmails.slice();

    if (suggestedEmails.length) {
      panelData.prefillEmails = [suggestedEmails[0]];
      panelData.useGmailSuggestions = true;
    }

    return panelData;
  }

  function getPlacementSuggestedEmailButtons() {
    return Array.prototype.slice.call(
      getPanelBody().querySelectorAll('.tm-ms-placement-suggested-email-btn')
    );
  }

  function getPlacementSuggestedEmails() {
    return getPlacementSuggestedEmailButtons()
      .map(function (button) {
        return String(button && button.getAttribute('data-email') || '').trim().toLowerCase();
      })
      .filter(Boolean);
  }

  function setPlacementSuggestedEmailButtonStates(selectedEmail) {
    var normalizedSelectedEmail = String(selectedEmail || '').trim().toLowerCase();

    getPlacementSuggestedEmailButtons().forEach(function (button) {
      var email = String(button.getAttribute('data-email') || '').trim().toLowerCase();
      var isSelected = Boolean(email && normalizedSelectedEmail && email === normalizedSelectedEmail);

      button.style.background = isSelected ? '#c2410c' : '#ffedd5';
      button.style.color = isSelected ? '#fff' : '#9a3412';
      button.style.borderColor = isSelected ? '#9a3412' : '#fdba74';
      button.textContent = isSelected ? 'Основной: ' + email : email;
    });
  }

  function syncPlacementSuggestedEmailSelectionFromTextarea() {
    var fields = getPlacementDraftFields();
    var recipients = normalizeRecipientList(fields.to && fields.to.value);
    setPlacementSuggestedEmailButtonStates(recipients.length ? recipients[0] : '');
  }

  function applyPlacementSuggestedEmail(email) {
    var normalizedEmail = String(email || '').trim().toLowerCase();
    var fields = getPlacementDraftFields();
    var recipients;
    var suggestedEmails;
    var remainingRecipients;
    var nextRecipients;

    if (!normalizedEmail || !fields.to) {
      return;
    }

    recipients = normalizeRecipientList(fields.to.value);
    suggestedEmails = getPlacementSuggestedEmails();
    remainingRecipients = recipients.filter(function (recipient) {
      return recipient !== normalizedEmail && suggestedEmails.indexOf(recipient) === -1;
    });
    nextRecipients = [normalizedEmail].concat(remainingRecipients);

    fields.to.value = nextRecipients.join(', ');
    setPlacementSuggestedEmailButtonStates(normalizedEmail);
    setPlacementMessage(
      '#tm-ms-placement-suggestion-note',
      'Выбран основной email для размещения: <b>' + escapeHtml(normalizedEmail) + '</b>.',
      '#7c2d12'
    );
    fields.to.focus();
  }

  function onPlacementSuggestedEmailClick(event) {
    var button = event && event.currentTarget;
    var email = button ? button.getAttribute('data-email') : '';

    applyPlacementSuggestedEmail(email);
  }

  function attachSearchEmailSaveHandlers() {
    getPanelBody().querySelectorAll('.tm-ms-search-email-action-btn').forEach(function (button) {
      button.addEventListener('click', onSearchEmailActionClick);
    });
  }

  function attachSearchFilterHandlers() {
    getPanelBody().querySelectorAll('.tm-ms-search-filter-btn').forEach(function (button) {
      button.addEventListener('click', function () {
        setSearchEmailFilter(button.getAttribute('data-filter'));
      });
    });
  }

  async function refreshSearchResultsAfterEmailAction(orderId, settings, preferredFilter) {
    var result = await fetchSearchData(orderId, settings, {
      skipLog: true,
      prefetch: false
    });

    if (!result.ok) {
      renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Письма по заказу');
      throw new Error('Не удалось обновить выдачу поиска после изменения списка.');
    }

    if (!result.data || !result.data.success) {
      throw new Error((result.data && result.data.error) || 'Не удалось обновить выдачу поиска.');
    }

    renderEmails(result.data, 'manual', settings);
    attachSearchEmailSaveHandlers();
    attachSearchFilterHandlers();

    if (preferredFilter) {
      setSearchEmailFilter(preferredFilter);
    }

    if (state.currentOrderId === orderId) {
      state.searchPrefetchResult = result;
      state.searchPrefetchPromise = null;
      state.searchPrefetchConsumed = true;
    }

    return result;
  }

  async function onSearchSaveEmailClick(normalizedEmail, orderId, settings) {
    var result;
    var linkHtml = '';
    var locationLabel = 'МойСклад';
    var message = '';

    setSearchEmailActionButtonState(normalizedEmail, 'ms', true, getSearchEmailActionLoadingLabel('ms'));
    setSearchSaveMessage(
      'Добавляю <b>' + escapeHtml(normalizedEmail) + '</b> в МойСклад...',
      '#92400e'
    );
    setStatus('Сохраняю email...', 'loading', 'search');

    try {
      result = await saveSupplierEmailsToMoySklad(orderId, settings, {
        emails: [normalizedEmail]
      });

      if (!result.ok) {
        setSearchEmailActionButtonState(normalizedEmail, 'ms', false, getSearchEmailActionDefaultLabel('ms'));
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Письма по заказу');
        setStatus('Сохранение email: ошибка', 'error', 'search');
        return;
      }

      if (!result.data || !result.data.success) {
        setSearchEmailActionButtonState(normalizedEmail, 'ms', false, getSearchEmailActionDefaultLabel('ms'));
        setSearchSaveMessage(
          escapeHtml((result.data && result.data.error) || 'Не удалось сохранить email в МойСклад'),
          '#991b1b'
        );
        setStatus('Сохранение email: ошибка', 'error', 'search');
        return;
      }

      syncSavedSupplierEmailsState(result.data);

      if (result.data.counterpartyLink) {
        linkHtml =
          ' <a href="' +
          escapeHtml(result.data.counterpartyLink) +
          '" target="_blank" rel="noopener noreferrer" style="color:inherit;font-weight:bold;">Открыть карточку</a>.';
      }

      if (result.data.storageTarget === 'contactPerson' && result.data.contactPersonName) {
        locationLabel = 'контактное лицо <b>' + escapeHtml(result.data.contactPersonName) + '</b>';
      } else if (result.data.storageTarget === 'counterparty') {
        locationLabel = 'поле email контрагента';
      } else if (result.data.storageTargetLabel) {
        locationLabel = escapeHtml(result.data.storageTargetLabel);
      }

      message = result.data.addedEmails && result.data.addedEmails.length
        ? 'Email <b>' + escapeHtml(normalizedEmail) + '</b> сохранен в ' + locationLabel + '.' + linkHtml
        : 'Email <b>' + escapeHtml(normalizedEmail) + '</b> уже был сохранен в МойСклад.' + linkHtml;

      setSearchEmailActionButtonState(
        normalizedEmail,
        'ms',
        true,
        result.data.addedEmails && result.data.addedEmails.length ? 'added to MS' : 'already in MS'
      );
      setSearchSaveMessage(message, '#166534');
      setStatus('Email сохранен', 'ok', 'search');
    } catch (error) {
      setSearchEmailActionButtonState(normalizedEmail, 'ms', false, getSearchEmailActionDefaultLabel('ms'));
      setSearchSaveMessage(
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Сохранение email: ошибка', 'error', 'search');
    }
  }

  async function onSearchSaveEmailListClick(normalizedEmail, listType, orderId, settings) {
    var result;
    var listLabel = listType === 'transport' ? 'transport list' : 'ignore list';
    var successLabel = listType === 'transport' ? 'in transport list' : 'in ignore list';

    setSearchEmailActionButtonState(normalizedEmail, listType, true, getSearchEmailActionLoadingLabel(listType));
    setSearchSaveMessage(
      'Добавляю <b>' + escapeHtml(normalizedEmail) + '</b> в ' + escapeHtml(listLabel) + '...',
      '#92400e'
    );
    setStatus('Обновляю список поиска...', 'loading', 'search');

    try {
      result = await saveSearchEmailToList(orderId, settings, {
        listType: listType,
        emails: [normalizedEmail]
      });

      if (!result.ok) {
        setSearchEmailActionButtonState(normalizedEmail, listType, false, getSearchEmailActionDefaultLabel(listType));
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Письма по заказу');
        setStatus('Обновление списка: ошибка', 'error', 'search');
        return;
      }

      if (!result.data || !result.data.success) {
        setSearchEmailActionButtonState(normalizedEmail, listType, false, getSearchEmailActionDefaultLabel(listType));
        setSearchSaveMessage(
          escapeHtml((result.data && result.data.error) || 'Не удалось обновить список поиска'),
          '#991b1b'
        );
        setStatus('Обновление списка: ошибка', 'error', 'search');
        return;
      }

      setSearchEmailActionButtonState(
        normalizedEmail,
        listType,
        true,
        result.data.addedEmails && result.data.addedEmails.length ? successLabel : 'already there'
      );

      await refreshSearchResultsAfterEmailAction(
        orderId,
        settings,
        listType === 'transport' ? 'transport' : ''
      );

      setSearchSaveMessage(
        result.data.addedEmails && result.data.addedEmails.length
          ? 'Email <b>' + escapeHtml(normalizedEmail) + '</b> added to ' + escapeHtml(listLabel) + '.'
          : 'Email <b>' + escapeHtml(normalizedEmail) + '</b> уже был в ' + escapeHtml(listLabel) + '.',
        '#166534'
      );
      setStatus('Список обновлен', 'ok', 'search');
    } catch (error) {
      setSearchEmailActionButtonState(normalizedEmail, listType, false, getSearchEmailActionDefaultLabel(listType));
      setSearchSaveMessage(
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Обновление списка: ошибка', 'error', 'search');
    }
  }

  async function onSearchEmailActionClick(event) {
    var button = event && event.currentTarget;
    var email = button ? button.getAttribute('data-email') : '';
    var actionName = button ? button.getAttribute('data-action') : '';
    var normalizedEmail = String(email || '').trim().toLowerCase();
    var normalizedActionName = String(actionName || '').trim().toLowerCase();
    var orderId = getOrderIdFromUrl();
    var settings = ensureUserSettings();

    if (!orderId || !settings) {
      renderSettingsRequired('работы с email в поиске');
      setStatus('Нужно заполнить настройки', 'error', 'search');
      return;
    }

    if (!normalizedEmail) {
      setSearchSaveMessage('Не удалось определить email для действия.', '#991b1b');
      setStatus('Действие с email: ошибка', 'error', 'search');
      return;
    }

    if (normalizedActionName === 'transport' || normalizedActionName === 'ignore') {
      await onSearchSaveEmailListClick(normalizedEmail, normalizedActionName, orderId, settings);
      return;
    }

    await onSearchSaveEmailClick(normalizedEmail, orderId, settings);
  }

  function startBackgroundPrefetch(orderId) {
    var settings = ensureUserSettings({ silent: true });

    if (!orderId || !settings) {
      return;
    }

    if (state.currentOrderId !== orderId) {
      resetOrderState(orderId);
    }

    if (state.searchPrefetchPromise || state.searchPrefetchResult) {
      return;
    }

    state.searchPrefetchStartedAt = Date.now();
    setStatus('Фоновый поиск писем...', 'loading', 'search');

    state.searchPrefetchPromise = fetchSearchData(orderId, settings, {
      skipLog: true,
      prefetch: true
    })
      .then(function (result) {
        if (state.currentOrderId !== orderId) {
          return result;
        }

        state.searchPrefetchResult = result;
        state.searchPrefetchError = null;

        if (result.ok && result.data && result.data.success) {
          setStatus('Фоновый поиск готов', 'ok', 'search');
        } else if (result.ok && result.data && !result.data.success) {
          setStatus('Фоновый поиск: ошибка', 'error', 'search');
        } else {
          setStatus('Фоновый поиск: неверный ответ', 'error', 'search');
        }

        return result;
      })
      .catch(function (error) {
        if (state.currentOrderId === orderId) {
          state.searchPrefetchError = error;
          state.searchPrefetchResult = null;
          setStatus('Фоновый поиск: ошибка', 'error', 'search');
        }

        throw error;
      });
  }

  async function onSearchButtonClick() {
    var orderId = getOrderIdFromUrl();
    var settings;
    var canUsePrefetch;

    if (!orderId) {
      renderError('Ошибка', '<div>Не удалось определить ID заказа из URL</div>', 'Письма по заказу');
      return;
    }

    settings = ensureUserSettings();
    if (!settings) {
      renderSettingsRequired('поиска писем');
      setStatus('Нужно заполнить настройки', 'error', 'search');
      return;
    }

    canUsePrefetch =
      state.currentOrderId === orderId &&
      state.searchPrefetchPromise &&
      !state.searchPrefetchConsumed;

    if (canUsePrefetch) {
      setPanelHtml('<div>Дожидаюсь фонового поиска...</div>', 'Письма по заказу', 'search');

      try {
        var prefetchResult = await state.searchPrefetchPromise;
        state.searchPrefetchConsumed = true;

        if (!prefetchResult.ok) {
          renderNonJsonResponse(prefetchResult.text, prefetchResult.requestUrl, prefetchResult.status, 'Письма по заказу');
          return;
        }

        renderEmails(prefetchResult.data, 'prefetch', settings);
        attachSearchEmailSaveHandlers();
        attachSearchFilterHandlers();
        return;
      } catch (error) {
        renderError(
          'Ошибка фонового запроса',
          '<pre style="white-space:pre-wrap;font-size:12px;color:#444;background:#f6f8fa;border:1px solid #e5e7eb;padding:10px;border-radius:8px;">' +
            escapeHtml(error && error.stack ? error.stack : String(error)) +
            '</pre>',
          'Письма по заказу'
        );
        return;
      }
    }

    setPanelHtml('<div>Ищу письма...</div>', 'Письма по заказу', 'search');
    setStatus('Ручной поиск писем...', 'loading', 'search');

    try {
      var result = await fetchSearchData(orderId, settings, {
        skipLog: false,
        prefetch: false
      });

      if (!result.ok) {
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Письма по заказу');
        setStatus('Ручной поиск: ошибка', 'error', 'search');
        return;
      }

      renderEmails(result.data, 'manual', settings);
      attachSearchEmailSaveHandlers();
      attachSearchFilterHandlers();
      setStatus('Ручной поиск завершен', 'ok', 'search');

      if (state.currentOrderId === orderId) {
        state.searchPrefetchConsumed = true;
      }
    } catch (error) {
      renderError(
        'Ошибка запроса к Apps Script',
        '<pre style="white-space:pre-wrap;font-size:12px;color:#444;background:#f6f8fa;border:1px solid #e5e7eb;padding:10px;border-radius:8px;">' +
          escapeHtml(error && error.stack ? error.stack : String(error)) +
          '</pre>',
        'Письма по заказу'
      );
      setStatus('Ручной поиск: ошибка', 'error', 'search');
    }
  }

  async function onPlacementButtonClick() {
    var orderId = getOrderIdFromUrl();
    var settings;
    var result;
    var hasPlacementPrefetch;

    if (!orderId) {
      renderError('Ошибка', '<div>Не удалось определить ID заказа из URL</div>', 'Размещение PO');
      return;
    }

    settings = ensureUserSettings();
    if (!settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    hasPlacementPrefetch =
      state.currentOrderId === orderId &&
      Boolean(state.placementMetaResult || state.placementMetaPromise);

    if (!hasPlacementPrefetch) {
      setPanelHtml('<div>Загружаю данные размещения...</div>', 'Размещение PO', 'placement');
      setStatus('Загружаю размещение...', 'loading', 'placement');
    }

    try {
      result = await loadPlacementMeta(orderId, settings);

      if (!result.ok) {
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Размещение: ошибка', 'error', 'placement');
        return;
      }

      if (!Array.isArray(result.data && result.data.emails) || !result.data.emails.length) {
        setPanelHtml('<div>Ищу вероятные email поставщика в уже найденных письмах...</div>', 'Размещение PO', 'placement');
      }

      var panelData = await buildPlacementPanelData(orderId, settings, result.data);

      if (state.currentOrderId === orderId && state.placementMetaResult && state.placementMetaResult.data) {
        state.placementMetaResult.data = panelData;
      }

      renderPlacementPanel(panelData);
      setStatus('', result.data && result.data.success && result.data.canPlace ? 'ok' : 'error', 'placement');
    } catch (error) {
      renderError(
        'Ошибка запроса к Apps Script',
        '<pre style="white-space:pre-wrap;font-size:12px;color:#444;background:#f6f8fa;border:1px solid #e5e7eb;padding:10px;border-radius:8px;">' +
          escapeHtml(error && error.stack ? error.stack : String(error)) +
          '</pre>',
        'Размещение PO'
      );
      setStatus('Размещение: ошибка', 'error', 'placement');
    }
  }

  async function onTrackingButtonClick() {
    var entries = syncTrackingButtonVisibility();
    var results = [];
    var index;
    var showSupportNote = false;

    if (!entries.length) {
      renderTrackingPanel({
        rawFieldValue: state.trackingFieldValue,
        sourceLabel: state.trackingSourceLabel,
        entries: []
      });
      setStatus('Трекинг не найден', 'error', 'tracking');
      return;
    }

    setPanelHtml('<div>Проверяю трекинг по официальным сервисам...</div>', 'Трекинг поставки', 'tracking');
    setStatus('Проверяю трекинг...', 'loading', 'tracking');

    for (index = 0; index < entries.length; index += 1) {
      results.push(Object.assign({}, entries[index], await fetchTrackingDetails(entries[index])));

      if (entries[index].carrier === 'ups' || entries[index].carrier === 'dhl') {
        showSupportNote = true;
      }
    }

    state.trackingLastResult = results.slice();
    renderTrackingPanel({
      rawFieldValue: state.trackingFieldValue,
      sourceLabel: state.trackingSourceLabel,
      entries: results,
      showSupportNote: showSupportNote
    });
    setStatus(
      results.some(function (item) { return item && item.success; }) ? 'Трекинг загружен' : 'Открыт fallback по ссылкам',
      results.some(function (item) { return item && item.success; }) ? 'ok' : 'error',
      'tracking'
    );
  }

  async function onPlacementDownloadClick() {
    var orderId = getOrderIdFromUrl();
    var settings = ensureUserSettings();
    var result;

    if (!orderId || !settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    setPlacementMessage('#tm-ms-placement-download-note', 'Готовлю временную ссылку на XLS...', '#92400e');
    setStatus('Готовлю XLS...', 'loading', 'placement');

    try {
      result = await fetchPlacementExport(orderId, settings);

      if (!result.ok) {
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Скачивание XLS: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success || !result.data.downloadUrl) {
        setPlacementMessage(
          '#tm-ms-placement-download-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось получить ссылку на XLS'),
          '#991b1b'
        );
        setStatus('Скачивание XLS: ошибка', 'error', 'placement');
        return;
      }

      state.lastPlacementDownloadUrl = result.data.downloadUrl;
      openExternalUrl(state.lastPlacementDownloadUrl);
      setPlacementMessage(
        '#tm-ms-placement-download-note',
        'XLS готов. Если браузер не начал скачивание, <a href="' +
          escapeHtml(state.lastPlacementDownloadUrl) +
          '" target="_blank" rel="noopener noreferrer">открой ссылку вручную</a>.',
        '#166534'
      );
      setStatus('XLS готов', 'ok', 'placement');
    } catch (error) {
      setPlacementMessage(
        '#tm-ms-placement-download-note',
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Скачивание XLS: ошибка', 'error', 'placement');
    }
  }

  function getPlacementDraftFields() {
    var panelBody = getPanelBody();

    return {
      to: panelBody.querySelector('#tm-ms-placement-to'),
      subject: panelBody.querySelector('#tm-ms-placement-subject'),
      body: panelBody.querySelector('#tm-ms-placement-body')
    };
  }

  async function onPlacementSendEmailClick() {
    var settings = ensureUserSettings();
    var fields;
    var recipients;
    var orderId = getOrderIdFromUrl();
    var result;

    if (!orderId || !settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    fields = getPlacementDraftFields();
    recipients = normalizeRecipientList(fields.to && fields.to.value);

    if (!recipients.length) {
      setPlacementMessage('#tm-ms-placement-send-note', 'Добавь хотя бы один email получателя.', '#991b1b');
      setStatus('Отправка: нет получателей', 'error', 'placement');
      if (fields.to) {
        fields.to.focus();
      }
      return;
    }

    if (fields.to) {
      fields.to.value = recipients.join(', ');
    }

    setPlacementSendButtonState(true, 'Отправляю...');
    setPlacementDraftButtonState(true, getPlacementDraftButtonLabel());
    setPlacementMessage('#tm-ms-placement-send-note', 'Отправляю письмо через Apps Script...', '#92400e');
    setStatus('Отправляю письмо...', 'loading', 'placement');

    try {
      result = await sendPlacementEmail(orderId, settings, {
        to: recipients,
        subject: fields.subject ? fields.subject.value : '',
        body: fields.body ? fields.body.value : ''
      });

      if (!result.ok) {
        setPlacementSendButtonState(false, 'Отправить письмо');
        setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Отправка письма: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success) {
        setPlacementSendButtonState(false, 'Отправить письмо');
        setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
        setPlacementMessage(
          '#tm-ms-placement-send-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось отправить письмо'),
          '#991b1b'
        );
        setStatus('Отправка письма: ошибка', 'error', 'placement');
        return;
      }

      state.placementEmailSent = true;
      setPlacementSendButtonState(true, 'Письмо отправлено');
      setPlacementDraftButtonState(true, 'Черновик не нужен');
      setPlacementStateButtonState(false, 'Поставить статус "Размещен"');
      setPlacementMessage(
        '#tm-ms-placement-send-note',
        'Письмо отправлено. Вложение: ' + escapeHtml(result.data.attachmentFileName || ''),
        '#166534'
      );
      setPlacementMessage(
        '#tm-ms-placement-state-note',
        'Теперь можно подтвердить размещение и поставить статус "Размещен".',
        '#166534'
      );
      setStatus('Письмо отправлено', 'ok', 'placement');
    } catch (error) {
      setPlacementSendButtonState(false, 'Отправить письмо');
      setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
      setPlacementMessage(
        '#tm-ms-placement-send-note',
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Отправка письма: ошибка', 'error', 'placement');
    }
  }

  async function onPlacementSaveDraftClick() {
    var settings = ensureUserSettings();
    var fields;
    var recipients;
    var orderId = getOrderIdFromUrl();
    var result;

    if (!orderId || !settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    fields = getPlacementDraftFields();
    recipients = normalizeRecipientList(fields.to && fields.to.value);

    if (!recipients.length) {
      setPlacementMessage('#tm-ms-placement-send-note', 'Добавь хотя бы один email получателя.', '#991b1b');
      setStatus('Черновик: нет получателей', 'error', 'placement');
      if (fields.to) {
        fields.to.focus();
      }
      return;
    }

    if (fields.to) {
      fields.to.value = recipients.join(', ');
    }

    setPlacementDraftButtonState(true, 'Сохраняю...');
    setPlacementSendButtonState(true, 'Отправить письмо');
    setPlacementMessage('#tm-ms-placement-send-note', 'Сохраняю черновик в Gmail через Apps Script...', '#92400e');
    setStatus('Сохраняю черновик...', 'loading', 'placement');

    try {
      result = await savePlacementDraft(orderId, settings, {
        draftId: state.placementDraftId,
        to: recipients,
        subject: fields.subject ? fields.subject.value : '',
        body: fields.body ? fields.body.value : ''
      });

      if (!result.ok) {
        setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
        setPlacementSendButtonState(false, 'Отправить письмо');
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Сохранение черновика: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success) {
        setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
        setPlacementSendButtonState(false, 'Отправить письмо');
        setPlacementMessage(
          '#tm-ms-placement-send-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось сохранить черновик'),
          '#991b1b'
        );
        setStatus('Сохранение черновика: ошибка', 'error', 'placement');
        return;
      }

      state.placementDraftId = result.data.draftId || state.placementDraftId;
      setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
      setPlacementSendButtonState(false, 'Отправить письмо');
      setPlacementMessage(
        '#tm-ms-placement-send-note',
        result.data.updatedExisting
          ? 'Черновик обновлен в Gmail. Статус пока не менялся; после проверки можно отправить письмо отсюда или из папки Черновики.'
          : 'Черновик сохранен в Gmail. Статус пока не менялся; после проверки можно отправить письмо отсюда или из папки Черновики.',
        '#166534'
      );
      setPlacementMessage(
        '#tm-ms-placement-state-note',
        'Сохранение черновика не меняет статус. Для статуса нужно именно отправить письмо.',
        '#92400e'
      );
      setStatus('Черновик сохранен', 'ok', 'placement');
    } catch (error) {
      setPlacementDraftButtonState(false, getPlacementDraftButtonLabel());
      setPlacementSendButtonState(false, 'Отправить письмо');
      setPlacementMessage(
        '#tm-ms-placement-send-note',
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Сохранение черновика: ошибка', 'error', 'placement');
    }
  }

  async function onPlacementSaveCounterpartyEmailsClick() {
    var orderId = getOrderIdFromUrl();
    var settings = ensureUserSettings();
    var fields;
    var recipients;
    var result;
    var linkHtml = '';

    if (!orderId || !settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    fields = getPlacementDraftFields();
    recipients = normalizeRecipientList(fields.to && fields.to.value);

    if (!recipients.length) {
      setPlacementMessage('#tm-ms-placement-suggestion-note', 'Нечего сохранять в МойСклад: список получателей пуст.', '#991b1b');
      setStatus('Нет email для сохранения', 'error', 'placement');
      return;
    }

    if (fields.to) {
      fields.to.value = recipients.join(', ');
    }

    setPlacementCounterpartySaveButtonState(true, 'Сохраняю...');
    setPlacementMessage('#tm-ms-placement-suggestion-note', 'Добавляю email в МойСклад...', '#92400e');
    setStatus('Сохраняю email в МойСклад...', 'loading', 'placement');

    try {
      result = await savePlacementCounterpartyEmails(orderId, settings, {
        emails: recipients
      });

      if (!result.ok) {
        setPlacementCounterpartySaveButtonState(false, getPlacementMoySkladSaveButtonLabel());
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Сохранение email: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success) {
        setPlacementCounterpartySaveButtonState(false, getPlacementMoySkladSaveButtonLabel());
        setPlacementMessage(
          '#tm-ms-placement-suggestion-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось сохранить email в МойСклад'),
          '#991b1b'
        );
        setStatus('Сохранение email: ошибка', 'error', 'placement');
        return;
      }

      syncSavedSupplierEmailsState(result.data);

      if (result.data.counterpartyLink) {
        linkHtml =
          ' <a href="' +
          escapeHtml(result.data.counterpartyLink) +
          '" target="_blank" rel="noopener noreferrer" style="color:inherit;font-weight:bold;">Открыть карточку</a>.';
      }

      setPlacementCounterpartySaveButtonState(
        true,
        result.data.addedEmails && result.data.addedEmails.length ? 'Добавлено в МойСклад' : 'Уже есть в МойСклад'
      );
      setPlacementMessage(
        '#tm-ms-placement-suggestion-note',
        result.data.addedEmails && result.data.addedEmails.length
          ? (
              result.data.storageTarget === 'contactPerson' && result.data.contactPersonName
                ? 'Email сохранены в контактное лицо <b>' + escapeHtml(result.data.contactPersonName) + '</b>.' + linkHtml
                : 'Email сохранены в поле email контрагента.' + linkHtml
            )
          : 'Эти email уже были сохранены в МойСклад.' + linkHtml,
        '#166534'
      );
      setStatus('Email сохранены в МойСклад', 'ok', 'placement');
    } catch (error) {
      setPlacementCounterpartySaveButtonState(false, getPlacementMoySkladSaveButtonLabel());
      setPlacementMessage(
        '#tm-ms-placement-suggestion-note',
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Сохранение email: ошибка', 'error', 'placement');
    }
  }

  async function onPlacementSetStateClick() {
    var orderId = getOrderIdFromUrl();
    var settings = ensureUserSettings();
    var result;
    var cachedData = state.placementMetaResult && state.placementMetaResult.data ? state.placementMetaResult.data : null;

    if (!orderId || !settings) {
      renderSettingsRequired('размещения PO');
      setStatus('Нужно заполнить настройки', 'error', 'placement');
      return;
    }

    if (!state.placementEmailSent) {
      setPlacementMessage(
        '#tm-ms-placement-state-note',
        'Сначала отправь письмо из этой панели, затем подтверждай статус.',
        '#991b1b'
      );
      setStatus('Сначала отправь письмо', 'error', 'placement');
      return;
    }

    if (!window.confirm('Поставить статус "Размещен"? Проверь, что письмо уже отправлено из этой панели.')) {
      return;
    }

    setPlacementStateButtonState(true, 'Меняю статус...');
    setPlacementMessage('#tm-ms-placement-state-note', 'Обновляю статус заказа в MoySklad...', '#92400e');
    setStatus('Меняю статус...', 'loading', 'placement');

    try {
      result = await fetchPlacementSetState(orderId, settings);

      if (!result.ok) {
        setPlacementStateButtonState(false, 'Поставить статус "Размещен"');
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Смена статуса: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success) {
        setPlacementStateButtonState(false, 'Поставить статус "Размещен"');
        setPlacementMessage(
          '#tm-ms-placement-state-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось изменить статус'),
          '#991b1b'
        );
        setStatus('Смена статуса: ошибка', 'error', 'placement');
        return;
      }

      state.placementMetaResult = {
        ok: true,
        status: result.status,
        requestUrl: result.requestUrl,
        text: result.text,
        data: Object.assign({}, cachedData || {}, {
          success: true,
          currentStateName: result.data.stateName,
          currentStateHref: result.data.stateHref,
          canPlace: false
        })
      };

      setPlacementButtonVisible(false);
      updatePlacementStateValue(result.data.stateName || 'Размещен');
      setPlacementStateButtonState(true, result.data.alreadyPlaced ? 'Статус уже установлен' : 'Статус установлен');
      setPlacementMessage(
        '#tm-ms-placement-state-note',
        result.data.alreadyPlaced
          ? 'Заказ уже был в статусе "Размещен".'
          : 'Статус обновлен. Кнопка "Размесить заказ" для этого заказа скрыта до следующего изменения состояния.',
        '#166534'
      );
      setStatus('Статус обновлен', 'ok', 'placement');
    } catch (error) {
      setPlacementStateButtonState(false, 'Поставить статус "Размещен"');
      setPlacementMessage(
        '#tm-ms-placement-state-note',
        escapeHtml(error && error.message ? error.message : String(error)),
        '#991b1b'
      );
      setStatus('Смена статуса: ошибка', 'error', 'placement');
    }
  }

  function updateVisibilityAndPrefetch() {
    var searchButton = createSearchButton();
    var orderId;
    var settings;

    createPlacementButton();
    createTrackingButton();

    if (isPurchaseOrderPage()) {
      searchButton.style.display = 'block';
      orderId = getOrderIdFromUrl();

      if (state.currentOrderId !== orderId) {
        resetOrderState(orderId);
      }

      settings = ensureUserSettings({ silent: true });
      if (!settings) {
        setPlacementButtonVisible(false);
        syncTrackingButtonVisibility();
        setStatus('Нужно заполнить настройки', 'error', 'search');
        return;
      }

      if (orderId) {
        startBackgroundPrefetch(orderId);
        startPlacementMetaPrefetch(orderId);
        syncTrackingButtonVisibility();
      } else {
        setPlacementButtonVisible(false);
        setTrackingButtonVisible(false);
      }

      return;
    }

    searchButton.style.display = 'none';
    setPlacementButtonVisible(false);
    setTrackingButtonVisible(false);
    hidePanel();
    setStatus('');
    resetOrderState(null);
  }

  function init() {
    createSearchButton();
    createPlacementButton();
    createTrackingButton();
    createPanel();
    registerMenuCommands();
    updateVisibilityAndPrefetch();

    var lastHash = window.location.hash;

    window.setInterval(function () {
      if (window.location.hash !== lastHash) {
        lastHash = window.location.hash;
        updateVisibilityAndPrefetch();
        return;
      }

      if (isPurchaseOrderPage()) {
        syncTrackingButtonVisibility();
      }
    }, APP_CONFIG.POLL_INTERVAL_MS);

    window.setTimeout(function () {
      if (isPurchaseOrderPage()) {
        var orderId = getOrderIdFromUrl();

        if (orderId) {
          startBackgroundPrefetch(orderId);
          startPlacementMetaPrefetch(orderId);
        }
      }
    }, APP_CONFIG.SECONDARY_PREFETCH_DELAY_MS);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
