// ==UserScript==
// @name         MoySklad - Поиск писем по заказу поставщику
// @namespace    https://tampermonkey.net/
// @version      0.1.7
// @description  Ищет письма по заказу поставщику через Google Apps Script
// @author       Codex + Spiralwave
// @match        https://online.moysklad.ru/app/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
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
    PANEL_ID: 'tm-ms-mail-search-panel',
    PANEL_TITLE_CLASS: 'tm-ms-panel-title',
    STATUS_ID: 'tm-ms-mail-search-status',
    DEFAULT_GMAIL_ACCOUNT_INDEX: 1,
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

  async function savePlacementCounterpartyEmails(orderId, settings, payload) {
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

  var SEARCH_STATUS_BORDER_COLORS = {
    loading: '#eab308',
    ok: '#16a34a',
    error: '#dc2626'
  };
  var PLACEMENT_CONFIRMED_BORDER_COLOR = '#9664bf';

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

  function getPlacementButtonBorderColor() {
    return createPlacementButton().style.display === 'none'
      ? 'transparent'
      : PLACEMENT_CONFIRMED_BORDER_COLOR;
  }

  function resetActionButtonBorders() {
    createSearchButton().style.borderColor = 'transparent';
    createPlacementButton().style.borderColor = getPlacementButtonBorderColor();
  }

  function setStatus(text, mode, buttonKey, options) {
    var badge = document.getElementById(APP_CONFIG.STATUS_ID);
    var targetButton = buttonKey === 'placement' ? createPlacementButton() : createSearchButton();
    var statusMode = mode || 'neutral';

    resetActionButtonBorders();
    if (badge) {
      badge.style.display = 'none';
      badge.textContent = '';
    }

    if (buttonKey !== 'search') {
      return;
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
    panel.style.top = '168px';
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

  function renderEmails(data, sourceLabel, settings) {
    var html = '';
    var suggestions = data && data.supplierEmailSuggestions;
    var suggestionEmails = suggestions && Array.isArray(suggestions.suggestedEmails)
      ? suggestions.suggestedEmails
      : [];

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

    if (suggestionEmails.length) {
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

    if (!data.emails || !data.emails.length) {
      html += '<div style="padding:10px 0;">Письма не найдены</div>';
      setPanelHtml(html, 'Письма по заказу', 'search');
      return;
    }

    html += '<div style="margin-bottom:12px;color:#555;">Найдено писем: <b>' + escapeHtml(data.count) + '</b></div>';

    data.emails.forEach(function (email) {
      var fixedLink = forceGmailAccount(email.link, settings);

      html += '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px;background:#fff;">';
      html += '<div style="font-weight:bold;font-size:14px;margin-bottom:8px;">' + escapeHtml(email.subject || '(без темы)') + '</div>';
      html += '<div style="font-size:12px;color:#666;margin-bottom:6px;"><b>От:</b> ' + escapeHtml(email.from || '') + '</div>';

      if (email.to) {
        html += '<div style="font-size:12px;color:#666;margin-bottom:6px;"><b>Кому:</b> ' + escapeHtml(email.to) + '</div>';
      }

      if (email.cc) {
        html += '<div style="font-size:12px;color:#666;margin-bottom:6px;"><b>Копия:</b> ' + escapeHtml(email.cc) + '</div>';
      }

      html += '<div style="font-size:12px;color:#666;margin-bottom:10px;"><b>Дата:</b> ' + escapeHtml(formatDate(email.date)) + '</div>';
      html += '<div style="font-size:13px;line-height:1.45;margin-bottom:10px;color:#222;">' + escapeHtml(email.snippet || '') + '</div>';
      html += '<div style="font-size:12px;color:#666;margin-bottom:10px;"><b>Сообщений в треде:</b> ' + escapeHtml(email.messageCount || '') + '</div>';
      html += '<div><a href="' + escapeHtml(fixedLink) + '" target="_blank" style="color:#1976d2;text-decoration:none;font-weight:bold;">Открыть письмо</a></div>';
      html += '</div>';
    });

    setPanelHtml(html, 'Письма по заказу', 'search');
  }

  function renderPlacementPanel(data) {
    var recipients = Array.isArray(data && data.prefillEmails)
      ? data.prefillEmails.join(', ')
      : (Array.isArray(data && data.emails) ? data.emails.join(', ') : '');
    var subject = buildPlacementEmailSubject(data);
    var body = buildPlacementEmailBody(data);
    var attachmentFileName = String((data && data.attachmentFileName) || 'PO.xls');
    var statusButtonDisabled = !state.placementEmailSent;
    var sendButtonDisabled = state.placementEmailSent;
    var draftButtonDisabled = state.placementEmailSent;
    var draftButtonLabel = state.placementDraftId ? 'Обновить черновик' : 'Сохранить черновик';
    var useGmailSuggestions = Boolean(data && data.useGmailSuggestions);
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
      html += '<div style="font-weight:bold;color:#9a3412;margin-bottom:6px;">В карточке контрагента нет email</div>';
      html += '<div style="font-size:13px;color:#7c2d12;margin-bottom:10px;">Предлагаю использовать адреса, найденные в почте. Они уже подставлены в получателей, но их все равно можно отредактировать перед отправкой.</div>';
      html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:' + (gmailSuggestionCandidates.length ? '10px' : '12px') + ';">';

      (Array.isArray(data.prefillEmails) ? data.prefillEmails : []).forEach(function (email) {
        html += '<span style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;background:#ffedd5;color:#9a3412;font-size:12px;font-weight:bold;">' + escapeHtml(email) + '</span>';
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
      html += '<button id="tm-ms-placement-save-counterparty-btn" type="button" style="padding:9px 12px;border:1px solid #c2410c;border-radius:8px;background:#fff;color:#9a3412;cursor:pointer;font-weight:bold;">Добавить в карточку контрагента</button>';

      if (data.counterpartyLink) {
        html += '<a href="' + escapeHtml(data.counterpartyLink) + '" target="_blank" rel="noopener noreferrer" style="font-size:12px;color:#9a3412;text-decoration:none;font-weight:bold;">Открыть контрагента</a>';
      }

      html += '</div>';
      html += '<div id="tm-ms-placement-suggestion-note" style="margin-top:10px;font-size:12px;color:#7c2d12;">Один клик добавит текущий список получателей в поле email контрагента.</div>';
      html += '</div>';
    }

    html += '<label style="display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:6px;">Кому</label>';
    html += '<textarea id="tm-ms-placement-to" rows="3" style="width:100%;box-sizing:border-box;border:1px solid #d1d5db;border-radius:8px;padding:8px 10px;font:inherit;resize:vertical;margin-bottom:10px;">' + escapeHtml(recipients) + '</textarea>';
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
      getPanelBody().querySelector('#tm-ms-placement-save-counterparty-btn').addEventListener('click', onPlacementSaveCounterpartyEmailsClick);
    }
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

  function getPlacementDraftButtonLabel() {
    return state.placementDraftId ? 'Обновить черновик' : 'Сохранить черновик';
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

  function requestPlacementMeta(orderId, settings) {
    state.placementMetaPromise = fetchPlacementMeta(orderId, settings)
      .then(function (result) {
        if (state.currentOrderId !== orderId) {
          return result;
        }

        state.placementMetaResult = result;
        state.placementMetaError = null;
        syncPlacementButtonWithResult(orderId, result);
        return result;
      })
      .catch(function (error) {
        if (state.currentOrderId === orderId) {
          state.placementMetaError = error;
          state.placementMetaResult = null;
          setPlacementButtonVisible(false);
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

    panelData.prefillEmails = baseEmails.slice();
    panelData.useGmailSuggestions = false;
    panelData.gmailSuggestionCandidates = [];

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
    panelData.gmailSuggestionCandidates = Array.isArray(suggestions.candidates)
      ? suggestions.candidates
      : [];

    if (Array.isArray(suggestions.suggestedEmails) && suggestions.suggestedEmails.length) {
      panelData.prefillEmails = suggestions.suggestedEmails.slice();
      panelData.useGmailSuggestions = true;
    }

    return panelData;
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
      setPlacementMessage('#tm-ms-placement-suggestion-note', 'Нечего сохранять в карточку: список получателей пуст.', '#991b1b');
      setStatus('Нет email для карточки', 'error', 'placement');
      return;
    }

    if (fields.to) {
      fields.to.value = recipients.join(', ');
    }

    setPlacementCounterpartySaveButtonState(true, 'Сохраняю...');
    setPlacementMessage('#tm-ms-placement-suggestion-note', 'Добавляю email в карточку контрагента...', '#92400e');
    setStatus('Сохраняю email в карточку...', 'loading', 'placement');

    try {
      result = await savePlacementCounterpartyEmails(orderId, settings, {
        emails: recipients
      });

      if (!result.ok) {
        setPlacementCounterpartySaveButtonState(false, 'Добавить в карточку контрагента');
        renderNonJsonResponse(result.text, result.requestUrl, result.status, 'Размещение PO');
        setStatus('Сохранение email: ошибка', 'error', 'placement');
        return;
      }

      if (!result.data || !result.data.success) {
        setPlacementCounterpartySaveButtonState(false, 'Добавить в карточку контрагента');
        setPlacementMessage(
          '#tm-ms-placement-suggestion-note',
          escapeHtml((result.data && result.data.error) || 'Не удалось сохранить email в карточку'),
          '#991b1b'
        );
        setStatus('Сохранение email: ошибка', 'error', 'placement');
        return;
      }

      if (state.placementMetaResult && state.placementMetaResult.data) {
        state.placementMetaResult.data.emails = Array.isArray(result.data.emails) ? result.data.emails.slice() : recipients.slice();
        state.placementMetaResult.data.counterpartyEmail = result.data.counterpartyEmail || '';
        state.placementMetaResult.data.prefillEmails = Array.isArray(result.data.emails) ? result.data.emails.slice() : recipients.slice();
        state.placementMetaResult.data.useGmailSuggestions = false;
        state.placementMetaResult.data.gmailSuggestionCandidates = [];
      }

      if (result.data.counterpartyLink) {
        linkHtml =
          ' <a href="' +
          escapeHtml(result.data.counterpartyLink) +
          '" target="_blank" rel="noopener noreferrer" style="color:inherit;font-weight:bold;">Открыть карточку</a>.';
      }

      setPlacementCounterpartySaveButtonState(
        true,
        result.data.addedEmails && result.data.addedEmails.length ? 'Добавлено в карточку' : 'Уже есть в карточке'
      );
      setPlacementMessage(
        '#tm-ms-placement-suggestion-note',
        result.data.addedEmails && result.data.addedEmails.length
          ? 'Email сохранены в карточке контрагента.' + linkHtml
          : 'Эти email уже были в карточке контрагента.' + linkHtml,
        '#166534'
      );
      setStatus('Email сохранены в карточку', 'ok', 'placement');
    } catch (error) {
      setPlacementCounterpartySaveButtonState(false, 'Добавить в карточку контрагента');
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

    if (isPurchaseOrderPage()) {
      searchButton.style.display = 'block';
      orderId = getOrderIdFromUrl();

      if (state.currentOrderId !== orderId) {
        resetOrderState(orderId);
      }

      settings = ensureUserSettings({ silent: true });
      if (!settings) {
        setPlacementButtonVisible(false);
        setStatus('Нужно заполнить настройки', 'error', 'search');
        return;
      }

      if (orderId) {
        startBackgroundPrefetch(orderId);
        startPlacementMetaPrefetch(orderId);
      } else {
        setPlacementButtonVisible(false);
      }

      return;
    }

    searchButton.style.display = 'none';
    setPlacementButtonVisible(false);
    hidePanel();
    setStatus('');
    resetOrderState(null);
  }

  function init() {
    createSearchButton();
    createPlacementButton();
    createPanel();
    registerMenuCommands();
    updateVisibilityAndPrefetch();

    var lastHash = window.location.hash;

    window.setInterval(function () {
      if (window.location.hash !== lastHash) {
        lastHash = window.location.hash;
        updateVisibilityAndPrefetch();
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
