/**
 * Excel × No VBA Lab のサイト共通スクリプト
 */
(function () {
  'use strict';

  var INDEX_STATE_KEY = 'excel_no_vba_lab_index_state_v1';

  function getRareTechConfig() {
    return (window.ExcelLab && window.ExcelLab.rareTech) || window.RARETECH_CONFIG || null;
  }

  function syncLayoutVars() {
    var root = document.documentElement;
    var header = document.querySelector('.site-header');
    var tabNav = document.querySelector('.tab-nav');

    if (header) {
      root.style.setProperty('--header-height', header.offsetHeight + 'px');
    }
    if (tabNav) {
      root.style.setProperty('--tab-nav-height', tabNav.offsetHeight + 'px');
    } else {
      root.style.setProperty('--tab-nav-height', '0px');
    }
  }

  function initLayoutVars() {
    syncLayoutVars();
    window.addEventListener('resize', syncLayoutVars);
    window.addEventListener('orientationchange', syncLayoutVars);
  }

  function initResponsiveTables() {
    var tables = document.querySelectorAll('.article-body table');
    tables.forEach(function (table) {
      if (table.closest('.table-wrapper, .excel-wrap, .formula-copy-list, .helper-display-table, .a4-paper, .a4-grid')) return;

      var wrapper = document.createElement('div');
      wrapper.className = 'table-wrapper';
      wrapper.tabIndex = 0;
      wrapper.setAttribute('aria-label', '表を横スクロールできます');

      table.parentNode.insertBefore(wrapper, table);
      wrapper.appendChild(table);
    });
  }

  function getHeaderOffset() {
    var header = document.querySelector('.site-header');
    var tabNav = document.querySelector('.tab-nav');
    var headerHeight = header ? header.offsetHeight : 0;
    var tabNavHeight = tabNav ? tabNav.offsetHeight : 0;
    return headerHeight + tabNavHeight + 12;
  }

  function scrollToMainContentTop() {
    var main = document.getElementById('main-content');
    if (!main) return;
    var targetY = main.getBoundingClientRect().top + window.pageYOffset - getHeaderOffset();
    window.scrollTo({ top: Math.max(targetY, 0), behavior: 'auto' });
  }

  function getActiveTabId() {
    var activePanel = document.querySelector('.tab-panel.active');
    return activePanel ? activePanel.id : 'home';
  }

  function saveIndexState() {
    try {
      var state = {
        tabId: getActiveTabId(),
        scrollY: window.pageYOffset || document.documentElement.scrollTop || 0,
        savedAt: Date.now()
      };
      sessionStorage.setItem(INDEX_STATE_KEY, JSON.stringify(state));
    } catch (e) {
      // ストレージ例外は無視する。
    }
  }

  function readIndexState() {
    try {
      var raw = sessionStorage.getItem(INDEX_STATE_KEY);
      if (!raw) return null;
      var state = JSON.parse(raw);
      if (!state || typeof state !== 'object') return null;
      return state;
    } catch (e) {
      return null;
    }
  }

  function shouldRestoreFromState() {
    var navEntry = (performance.getEntriesByType && performance.getEntriesByType('navigation')[0]) || null;
    var isBackForward = !!(navEntry && navEntry.type === 'back_forward');
    var sameOriginReferrer = document.referrer && document.referrer.indexOf(location.origin) === 0;
    return isBackForward || sameOriginReferrer;
  }

  function loadDeferredPanelMedia(panel) {
    if (!panel) return;
    panel.querySelectorAll('source[data-srcset]').forEach(function (source) {
      if (!source.getAttribute('srcset')) {
        source.setAttribute('srcset', source.getAttribute('data-srcset'));
      }
    });
    panel.querySelectorAll('img[data-src]').forEach(function (img) {
      if (!img.getAttribute('data-loaded-src')) {
        img.setAttribute('src', img.getAttribute('data-src'));
        img.setAttribute('data-loaded-src', 'true');
      }
    });
    panel.querySelectorAll('video[data-poster]').forEach(function (video) {
      if (!video.getAttribute('poster')) {
        video.setAttribute('poster', video.getAttribute('data-poster'));
      }
    });
  }

  function activateTab(targetId, options) {
    options = options || {};
    var shouldScroll = options.scroll !== false;
    var updateHash = options.updateHash === true;
    var focusPanel = options.focusPanel === true;

    var tabButtons = document.querySelectorAll('.tab-btn');
    var tabPanels = document.querySelectorAll('.tab-panel');
    var targetPanel = document.getElementById(targetId);
    var targetButton = document.querySelector('.tab-btn[data-tab="' + targetId + '"]');

    if (!targetPanel || !targetButton) return;

    tabButtons.forEach(function (btn) {
      btn.classList.remove('active');
      btn.setAttribute('aria-selected', 'false');
      btn.setAttribute('tabindex', '-1');
    });

    tabPanels.forEach(function (panel) {
      panel.classList.remove('active');
    });

    targetButton.classList.add('active');
    targetButton.setAttribute('aria-selected', 'true');
    targetButton.setAttribute('tabindex', '0');
    targetPanel.classList.add('active');
    loadDeferredPanelMedia(targetPanel);

    if (updateHash && history.pushState) {
      history.pushState(null, '', '#' + targetId);
    }

    if (shouldScroll) {
      requestAnimationFrame(function () {
        scrollToMainContentTop();
        if (focusPanel) targetPanel.focus();
      });
    } else if (focusPanel) {
      targetPanel.focus();
    }
  }

  window.activateTab = activateTab;

  function initTabs() {
    var tabNav = document.querySelector('.tab-nav-inner[role="tablist"]');
    var tabButtons = document.querySelectorAll('.tab-btn');
    if (!tabButtons.length) return;

    tabButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        activateTab(btn.getAttribute('data-tab'), {
          scroll: true,
          updateHash: true,
          focusPanel: false
        });
        saveIndexState();
      });
    });

    if (tabNav) {
      tabNav.addEventListener('keydown', function (e) {
        var buttons = Array.prototype.slice.call(tabButtons);
        var currentIndex = buttons.indexOf(document.activeElement);
        if (currentIndex < 0) return;

        var target = null;
        if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
          target = buttons[currentIndex + 1] || buttons[0];
        } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
          target = buttons[currentIndex - 1] || buttons[buttons.length - 1];
        } else if (e.key === 'Home') {
          target = buttons[0];
        } else if (e.key === 'End') {
          target = buttons[buttons.length - 1];
        }

        if (target) {
          e.preventDefault();
          target.focus();
          activateTab(target.getAttribute('data-tab'), {
            scroll: false,
            updateHash: true,
            focusPanel: false
          });
          saveIndexState();
        }
      });
    }

    document.addEventListener('click', function (e) {
      var rankingTrigger = e.target.closest('[data-ranking-trigger]');
      if (rankingTrigger) {
        if (e.defaultPrevented || e.button !== 0) return;
        if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
        e.preventDefault();
        window.switchRankingTab(rankingTrigger, rankingTrigger.dataset.rankingTrigger);
        saveIndexState();
        return;
      }

      var trigger = e.target.closest('[data-tab-trigger]');
      if (trigger) {
        if (e.defaultPrevented || e.button !== 0) return;
        if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
        e.preventDefault();
        activateTab(trigger.dataset.tabTrigger, {
          scroll: trigger.dataset.tabScroll === 'true',
          focusPanel: trigger.dataset.tabFocus === 'true',
          updateHash: trigger.dataset.tabUpdateHash === 'true'
        });
        saveIndexState();
        return;
      }

      var link = e.target.closest('a[href]');
      if (!link) return;
      if (e.defaultPrevented || e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
      if (link.target === '_blank') return;

      var href = link.getAttribute('href') || '';
      if (!href || href.charAt(0) === '#') return;
      if (/^javascript:/i.test(href)) return;
      if (/\.html([?#].*)?$/i.test(href)) saveIndexState();
    });

    window.addEventListener('pagehide', saveIndexState);
    window.addEventListener('beforeunload', saveIndexState);
    window.addEventListener('popstate', function () {
      var hash = window.location.hash ? window.location.hash.substring(1) : 'home';
      var panel = document.getElementById(hash);
      if (panel && panel.classList.contains('tab-panel')) {
        activateTab(hash, { scroll: false, updateHash: false, focusPanel: false });
      } else if (!hash || hash === 'home') {
        activateTab('home', { scroll: false, updateHash: false, focusPanel: false });
      }
    });

    var initialHash = window.location.hash ? window.location.hash.substring(1) : '';
    var validPanel = initialHash && document.getElementById(initialHash);

    if (validPanel && validPanel.classList.contains('tab-panel')) {
      activateTab(initialHash, { scroll: false, updateHash: false });
      return;
    }

    if (shouldRestoreFromState()) {
      var state = readIndexState();
      if (state && state.tabId && document.getElementById(state.tabId)) {
        activateTab(state.tabId, { scroll: false, updateHash: false });
        requestAnimationFrame(function () {
          window.scrollTo(0, Math.max(state.scrollY || 0, 0));
        });
      }
    }
  }

  function updateRarityBadges() {
    var cfg = getRareTechConfig();
    if (!cfg || typeof cfg !== 'object') return;

    var badges = document.querySelectorAll('[data-rare-id]');
    badges.forEach(function (badge) {
      var id = badge.getAttribute('data-rare-id');
      var rarity = cfg[id];

      if (!rarity || !rarity.rarity) {
        badge.style.display = 'none';
        return;
      }

      var text = rarity.rarity + '人に1人';
      if (badge.classList.contains('tag-rare')) {
        text = '💎 レアテク ' + text;
      } else if (badge.classList.contains('rank-impact')) {
        text = '🔥 ' + text;
      } else if (badge.classList.contains('rank-combined')) {
        text = (badge.getAttribute('data-rare-prefix') || '🔥💎') + ' ' + text;
      } else {
        text = '💎 レアテク ' + text;
      }

      badge.textContent = text;
      badge.style.display = 'inline-block';
    });
  }

  window.switchRankingTab = function (btn, tabId) {
    var btns = btn.parentElement.querySelectorAll('.ranking-tab-btn');
    btns.forEach(function (b) { b.classList.remove('active'); });
    btn.classList.add('active');

    var card = btn.closest('.ranking-card');
    if (!card) return;

    var panels = card.querySelectorAll('.ranking-panel');
    panels.forEach(function (p) { p.classList.remove('active'); });

    var target = card.querySelector('#ranking-' + tabId);
    if (target) target.classList.add('active');
  };

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }

    return new Promise(function (resolve, reject) {
      try {
        var ta = document.createElement('textarea');
        ta.value = text;
        ta.setAttribute('readonly', '');
        ta.style.cssText = 'position:fixed;top:-9999px;opacity:0;';
        document.body.appendChild(ta);
        ta.select();
        ta.setSelectionRange(0, text.length);
        var ok = document.execCommand('copy');
        document.body.removeChild(ta);
        ok ? resolve() : reject(new Error('execCommand failed'));
      } catch (err) {
        reject(err);
      }
    });
  }

  function ensureCopyAnnouncer() {
    var announcer = document.getElementById('copy-announcer');
    if (announcer) return announcer;

    announcer = document.createElement('div');
    announcer.id = 'copy-announcer';
    announcer.className = 'sr-only';
    announcer.setAttribute('role', 'status');
    announcer.setAttribute('aria-live', 'polite');
    document.body.appendChild(announcer);
    return announcer;
  }

  function announceCopy(message) {
    var announcer = ensureCopyAnnouncer();
    announcer.textContent = '';
    window.setTimeout(function () {
      announcer.textContent = message;
    }, 0);
  }

  function initCopyButtons() {
    var blocks = document.querySelectorAll('.code-block, .code-readonly');
    blocks.forEach(function (block) {
      if (block.querySelector('.copy-btn')) return;

      var btn = document.createElement('button');
      btn.className = 'copy-btn';
      btn.type = 'button';
      btn.textContent = 'コピー';
      btn.setAttribute('aria-label', 'コードをコピーする');
      btn.setAttribute('aria-live', 'polite');

      btn.addEventListener('click', function () {
        var clone = block.cloneNode(true);
        var cloneBtn = clone.querySelector('.copy-btn');
        if (cloneBtn && cloneBtn.parentNode) cloneBtn.parentNode.removeChild(cloneBtn);
        var textToCopy = clone.innerText.trim();

        copyText(textToCopy).then(function () {
          btn.textContent = 'コピー済み';
          btn.classList.add('copied');
          announceCopy('コードをコピーしました');
          setTimeout(function () {
            btn.textContent = 'コピー';
            btn.classList.remove('copied');
          }, 1600);
        }).catch(function () {
          btn.textContent = '失敗';
          announceCopy('コピーに失敗しました');
          setTimeout(function () {
            btn.textContent = 'コピー';
          }, 1200);
        });
      });

      block.appendChild(btn);
    });
  }

  window.copyFormula = function (btn) {
    var code = btn && btn.parentElement ? btn.parentElement.querySelector('code') : null;
    var text = code ? code.textContent.trim() : '';

    if (!text && btn && btn.parentElement) {
      text = btn.parentElement.textContent.replace('コードをコピー', '').replace('コピー', '').trim();
    }

    copyText(text).then(function () {
      var original = btn.textContent;
      btn.textContent = 'コピー済み';
      btn.classList.add('copied');
      btn.setAttribute('aria-live', 'polite');
      announceCopy('コードをコピーしました');
      setTimeout(function () {
        btn.textContent = original;
        btn.classList.remove('copied');
      }, 1600);
    }).catch(function () {
      btn.textContent = '失敗';
      btn.setAttribute('aria-live', 'polite');
      announceCopy('コピーに失敗しました');
      setTimeout(function () { btn.textContent = 'コピー'; }, 1200);
    });
  };

  function initReadingProgress() {
    var progressBar = document.querySelector('.reading-progress');
    var articleBody = document.querySelector('.article-body');
    if (!progressBar || !articleBody) return;

    function updateProgress() {
      var scrollTop = window.pageYOffset || document.documentElement.scrollTop || 0;
      var articleTop = articleBody.getBoundingClientRect().top + scrollTop;
      var articleHeight = articleBody.scrollHeight;
      var viewport = window.innerHeight || document.documentElement.clientHeight || 1;
      var effectiveHeight = Math.max(articleHeight - viewport * 0.65, 1);
      var ratio = (scrollTop - articleTop) / effectiveHeight;

      if (ratio < 0) ratio = 0;
      if (ratio > 1) ratio = 1;

      progressBar.style.width = (ratio * 100).toFixed(2) + '%';
    }

    updateProgress();
    window.addEventListener('scroll', updateProgress, { passive: true });
    window.addEventListener('resize', updateProgress);
    window.addEventListener('orientationchange', updateProgress);
  }

  if ('scrollRestoration' in history) {
    history.scrollRestoration = 'auto';
  }

  document.addEventListener('DOMContentLoaded', function () {
    initLayoutVars();
    initResponsiveTables();
    initTabs();
    updateRarityBadges();
    initCopyButtons();
    initReadingProgress();
  });
})();
