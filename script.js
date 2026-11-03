/**
 * Excel × No VBA Lab のサイト共通スクリプト
 */
(function () {
  'use strict';

  var INDEX_STATE_KEY = 'excel_no_vba_lab_index_state_v1';
  // script.js is served from the site root, so its URL gives the root on every page.
  var SITE_ROOT = (function () {
    var current = document.currentScript;
    return current && current.src ? new URL('./', current.src).href : new URL('./', document.baseURI).href;
  })();
  var catalogPromise;
  function loadCatalog() {
    if (!catalogPromise) {
      catalogPromise = fetch(SITE_ROOT + 'articles.json').then(function (response) {
        if (!response.ok) throw new Error('記事一覧を取得できません');
        return response.json();
      }).catch(function (error) { catalogPromise = null; throw error; });
    }
    return catalogPromise;
  }
  function getRareTechConfig() {
    return (window.ExcelLab && window.ExcelLab.rareTech) || window.RARETECH_CONFIG || null;
  }

  function syncLayoutVars() {
    var root = document.documentElement;
    var header = document.querySelector('.site-header');
    var tabNav = document.querySelector('.tab-nav');

    if (header) {
      // The header height itself uses --header-height on desktop.
      // Reset the inline value before measuring to avoid feeding a stale
      // measured height back into the next responsive resize.
      root.style.removeProperty('--header-height');
      root.style.setProperty(
        '--header-height',
        window.matchMedia('(max-width: 768px)').matches ? '0px' : header.offsetHeight + 'px'
      );
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

    var tabScroller = targetButton.closest('.tab-nav');
    if (tabScroller && window.matchMedia('(max-width: 768px)').matches) {
      requestAnimationFrame(function () {
        var targetLeft = targetButton.offsetLeft - ((tabScroller.clientWidth - targetButton.offsetWidth) / 2);
        tabScroller.scrollLeft = Math.max(targetLeft, 0);
      });
    }

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

  function showSectionInTab(section) {
    var hostPanel = section.closest('.tab-panel');
    if (!hostPanel) return false;
    activateTab(hostPanel.id, { scroll: false, updateHash: false });
    var collapsed = section.querySelector('details.philosophy-details');
    if (collapsed) collapsed.open = true;
    requestAnimationFrame(function () { section.scrollIntoView(); });
    return true;
  }

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
      } else if (panel) {
        showSectionInTab(panel);
      }
    });

    var initialHash = window.location.hash ? window.location.hash.substring(1) : '';
    var validPanel = initialHash && document.getElementById(initialHash);

    if (validPanel && validPanel.classList.contains('tab-panel')) {
      activateTab(initialHash, { scroll: false, updateHash: false });
      return;
    }

    // A section anchor such as #article-finder or #philosophy lives inside a tab:
    // show that tab and go to the section instead of restoring the previous state.
    if (validPanel && showSectionInTab(validPanel)) return;
    if (new URLSearchParams(window.location.search).get('q')) return;

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
    document.querySelectorAll('[data-rare-id]').forEach(function (badge) { badge.remove(); });
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

  function initFormulaCopyButtons() {
    var buttons = document.querySelectorAll('.formula-copy-btn');
    buttons.forEach(function (btn) {
      if (btn.dataset.formulaCopyReady === 'true') return;
      btn.dataset.formulaCopyReady = 'true';
      btn.setAttribute('aria-label', '数式をコピーする');
      btn.setAttribute('aria-live', 'polite');

      btn.addEventListener('click', function () {
        var row = btn.closest('.formula-copy-row');
        var code = row ? row.querySelector('.formula-copy-code') : null;
        var textToCopy = code ? code.textContent.trim() : '';

        if (!textToCopy) {
          btn.textContent = '失敗';
          announceCopy('コピーする数式が見つかりません');
          setTimeout(function () { btn.textContent = 'コピー'; }, 1200);
          return;
        }

        copyText(textToCopy).then(function () {
          btn.textContent = 'コピー済み';
          btn.classList.add('copied');
          announceCopy('数式をコピーしました');
          setTimeout(function () {
            btn.textContent = 'コピー';
            btn.classList.remove('copied');
          }, 1600);
        }).catch(function () {
          btn.textContent = '失敗';
          announceCopy('コピーに失敗しました');
          setTimeout(function () { btn.textContent = 'コピー'; }, 1200);
        });
      });
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

  function getArticleMeta() {
    var articleBody = document.querySelector('.article-body');
    var titleNode = document.querySelector('.article-header .article-title, .article-title');
    if (!articleBody || !titleNode) return null;

    var path = location.pathname || '';
    var fileName = path.split('/').pop() || '';
    if (!fileName || fileName === 'index.html') return null;

    var normalizedPath = path.replace(/^\/Excel-No-VBA-Lab\//, '').replace(/^\//, '');
    var section = normalizedPath.split('/')[0] || 'article';

    return {
      body: articleBody,
      title: titleNode.textContent.trim(),
      path: normalizedPath,
      section: section
    };
  }

  function sendArticleEvent(eventName, params) {
    if (location.protocol === 'file:') return;
    if (typeof window.gtag !== 'function') return;

    window.gtag('event', eventName, params);
  }

  function safeStorageGet(key) {
    try {
      return localStorage.getItem(key);
    } catch (e) {
      return null;
    }
  }

  function safeStorageSet(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch (e) {
      // ストレージが使えない環境では、ボタン表示だけ継続する。
    }
  }

  var ARTICLE_FEEDBACK_FORM = {
    viewUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSdvSi9yzHNaPpu1HCSbSeP4qJwEfY-g633dnh_SD5Xs9kMmdw/viewform',
    nameEntry: 'entry.1002031830',
    emailEntry: 'entry.356247462',
    categoryEntry: 'entry.1026400745',
    messageEntry: 'entry.113328341'
  };

  // 記事末尾の5択（P4-2：2026-11 変更）。category は既存Googleフォームの選択肢名に合わせる
  var ARTICLE_FEEDBACK_REACTIONS = {
    solved: {
      label: '解決した',
      category: '情報交換・感想',
      placeholder: '例：Excel 2021で手順どおりにできました。'
    },
    stuck_step: {
      label: '手順◯で止まった',
      category: '記事への質問・誤り報告',
      placeholder: '例：手順3で止まりました。ボタンが見つかりません。'
    },
    error: {
      label: 'エラーが出た',
      category: '記事への質問・誤り報告',
      placeholder: '例：#VALUE! が出ます。（業務データは書かないでください）'
    },
    env_diff: {
      label: '環境が違う',
      category: '記事に関するご質問',
      placeholder: '例：Excel 2016／Mac版／Web版を使っています。'
    },
    more: {
      label: 'もっと詳しく',
      category: '記事に関するご質問',
      placeholder: '例：複数条件の場合の例がほしいです。'
    }
  };

  function buildArticleCommentUrl(meta, reactionId, comment, allowPublish) {
    var reaction = ARTICLE_FEEDBACK_REACTIONS[reactionId] || ARTICLE_FEEDBACK_REACTIONS.more;

    var message = [
      '【記事末尾の匿名質問・コメント】',
      '記事名: ' + meta.title,
      '記事URL: ' + location.href.split('#')[0],
      '反応: ' + reaction.label,
      '匿名での掲載許可: ' + (allowPublish ? 'はい' : 'いいえ'),
      '',
      comment
    ].join('\n');

    var url = new URL(ARTICLE_FEEDBACK_FORM.viewUrl);
    url.searchParams.set('usp', 'pp_url');
    url.searchParams.set(ARTICLE_FEEDBACK_FORM.nameEntry, '匿名フィードバック');
    url.searchParams.set(ARTICLE_FEEDBACK_FORM.emailEntry, 'anonymous-feedback@example.invalid');
    url.searchParams.set(ARTICLE_FEEDBACK_FORM.categoryEntry, reaction.category);
    url.searchParams.set(ARTICLE_FEEDBACK_FORM.messageEntry, message);
    return url.href;
  }

  function initArticleFeedback() {
    var meta = getArticleMeta();
    if (!meta) return;
    if (meta.body.querySelector('.article-feedback')) return;

    var reactionStorageKey = 'excel_no_vba_lab_feedback_reaction_' + meta.path;
    var legacyStorageKey = 'excel_no_vba_lab_helpful_' + meta.path;
    var selectedReaction = safeStorageGet(reactionStorageKey);

    if (!ARTICLE_FEEDBACK_REACTIONS[selectedReaction]) {
      selectedReaction = safeStorageGet(legacyStorageKey) === '1' ? 'solved' : '';
    }

    var feedback = document.createElement('aside');
    feedback.className = 'article-feedback';
    feedback.setAttribute('aria-label', '記事へのフィードバック');

    var textWrap = document.createElement('div');
    textWrap.className = 'article-feedback-text';

    var heading = document.createElement('h2');
    heading.textContent = 'この記事はどうでしたか？';

    var note = document.createElement('p');
    note.textContent = '当てはまるものを1つ押してください。押した方にだけ1行の入力欄が出ます。名前・メールアドレスは不要です。業務データ・個人情報は書かないでください。';

    textWrap.appendChild(heading);
    textWrap.appendChild(note);

    var optionWrap = document.createElement('div');
    optionWrap.className = 'article-feedback-options';

    var status = document.createElement('p');
    status.className = 'article-feedback-status';
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    status.textContent = selectedReaction ? '種類を選択済みです。下の入力欄から送れます。' : '';

    var commentPanel = document.createElement('div');
    commentPanel.className = 'article-feedback-comment';
    commentPanel.hidden = !selectedReaction;

    var commentLabel = document.createElement('label');
    commentLabel.className = 'article-feedback-comment-label';
    commentLabel.textContent = 'ひとこと（任意・200文字以内）';

    var textarea = document.createElement('input');
    textarea.type = 'text';
    textarea.className = 'article-feedback-textarea article-feedback-input';
    textarea.maxLength = 200;
    textarea.placeholder = selectedReaction && ARTICLE_FEEDBACK_REACTIONS[selectedReaction] ? ARTICLE_FEEDBACK_REACTIONS[selectedReaction].placeholder : '';
    commentLabel.appendChild(textarea);

    var counter = document.createElement('span');
    counter.className = 'article-feedback-counter';
    counter.textContent = '0 / 200';

    textarea.addEventListener('input', function () {
      counter.textContent = textarea.value.length + ' / 200';
    });

    var publishLabel = document.createElement('label');
    publishLabel.className = 'article-feedback-publish';
    var publishCheckbox = document.createElement('input');
    publishCheckbox.type = 'checkbox';
    publishLabel.appendChild(publishCheckbox);
    publishLabel.appendChild(document.createTextNode(' 内容を匿名の「読者の声」として掲載してよい'));

    var honeyLabel = document.createElement('label');
    honeyLabel.className = 'article-feedback-honeypot';
    honeyLabel.setAttribute('aria-hidden', 'true');
    honeyLabel.textContent = '会社名';
    var honeyInput = document.createElement('input');
    honeyInput.type = 'text';
    honeyInput.tabIndex = -1;
    honeyInput.autocomplete = 'off';
    honeyLabel.appendChild(honeyInput);

    var commentActions = document.createElement('div');
    commentActions.className = 'article-feedback-comment-actions';

    var submitButton = document.createElement('button');
    submitButton.type = 'button';
    submitButton.className = 'article-feedback-submit';
    submitButton.textContent = '確認画面へ進む（Googleフォーム）';

    var confirmationLink = document.createElement('a');
    confirmationLink.className = 'article-feedback-detail-link';
    confirmationLink.target = '_blank';
    confirmationLink.rel = 'noopener';
    confirmationLink.textContent = '確認フォームを開き直す';
    confirmationLink.hidden = true;

    var detailLink = document.createElement('a');
    detailLink.className = 'article-feedback-detail-link';
    detailLink.href = 'https://forms.gle/6SnCHNWK6EcRy1XG8';
    detailLink.target = '_blank';
    detailLink.rel = 'noopener';
    detailLink.textContent = '個別の返信が必要な方 → お問い合わせフォーム';

    submitButton.addEventListener('click', function () {
      var comment = textarea.value.trim() || '（ひとことなし）';

      if (honeyInput.value) {
        status.textContent = '入力内容を確認してください。';
        return;
      }

      if (!navigator.onLine && location.protocol !== 'file:') {
        status.textContent = '通信できませんでした。接続を確認してもう一度お試しください。';
        return;
      }

      if (location.protocol === 'file:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1') {
        status.textContent = 'ローカル確認中のためフォームを開いていません。公開ページでは確認画面へ進みます。';
        return;
      }

      var confirmationUrl = buildArticleCommentUrl(meta, selectedReaction || 'more', comment, publishCheckbox.checked);
      confirmationLink.href = confirmationUrl;
      confirmationLink.hidden = false;
      window.open(confirmationUrl, '_blank', 'noopener');
      status.textContent = '確認画面へ進みます。Googleフォーム内の［送信］を押して完了してください。画面が開かない場合は「確認フォームを開き直す」を使えます。入力内容はこのページに残しています。';

      sendArticleEvent('article_comment_form_open', {
        event_category: 'article_feedback',
        event_label: meta.path,
        article_title: meta.title,
        article_path: meta.path,
        article_section: meta.section,
        feedback_reaction: selectedReaction || 'more',
        publish_allowed: publishCheckbox.checked ? 'yes' : 'no'
      });
    });

    commentActions.appendChild(submitButton);
    commentActions.appendChild(confirmationLink);
    commentActions.appendChild(detailLink);
    commentPanel.appendChild(commentLabel);
    commentPanel.appendChild(counter);
    commentPanel.appendChild(publishLabel);
    commentPanel.appendChild(honeyLabel);
    commentPanel.appendChild(commentActions);

    Object.keys(ARTICLE_FEEDBACK_REACTIONS).forEach(function (reactionId) {
      var reaction = ARTICLE_FEEDBACK_REACTIONS[reactionId];
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'article-feedback-option';
      button.textContent = reaction.label;
      button.setAttribute('data-reaction', reactionId);
      button.setAttribute('aria-pressed', selectedReaction === reactionId ? 'true' : 'false');
      if (selectedReaction === reactionId) button.classList.add('selected');

      button.addEventListener('click', function () {
        var previousReaction = selectedReaction;
        selectedReaction = reactionId;
        safeStorageSet(reactionStorageKey, reactionId);
        if (reactionId === 'solved') safeStorageSet(legacyStorageKey, '1');

        optionWrap.querySelectorAll('.article-feedback-option').forEach(function (option) {
          var selected = option.getAttribute('data-reaction') === reactionId;
          option.classList.toggle('selected', selected);
          option.setAttribute('aria-pressed', selected ? 'true' : 'false');
        });

        status.textContent = '「' + ARTICLE_FEEDBACK_REACTIONS[reactionId].label + '」を選びました。よければひとこと添えて送信してください。';
        commentPanel.hidden = false;
        textarea.placeholder = ARTICLE_FEEDBACK_REACTIONS[reactionId].placeholder;

        if (previousReaction !== reactionId) {
          sendArticleEvent('article_feedback_reaction', {
            event_category: 'article_feedback',
            event_label: meta.path,
            article_title: meta.title,
            article_path: meta.path,
            article_section: meta.section,
            feedback_reaction: reactionId
          });

          if (reactionId === 'solved') {
            sendArticleEvent('article_helpful_click', {
              event_category: 'article_feedback',
              event_label: meta.path,
              article_title: meta.title,
              article_path: meta.path,
              article_section: meta.section
            });
          }
        }
      });

      optionWrap.appendChild(button);
    });

    feedback.appendChild(textWrap);
    feedback.appendChild(optionWrap);
    feedback.appendChild(status);
    feedback.appendChild(commentPanel);

    var nav = meta.body.querySelector('.stage-nav');
    if (nav && nav.parentNode === meta.body) {
      meta.body.insertBefore(feedback, nav);
    } else {
      meta.body.appendChild(feedback);
    }
  }

  function initArticleReadDepth() {
    var meta = getArticleMeta();
    if (!meta) return;

    var sent = false;
    var threshold = 0.75;

    function checkDepth() {
      if (sent) return;

      var scrollTop = window.pageYOffset || document.documentElement.scrollTop || 0;
      var articleTop = meta.body.getBoundingClientRect().top + scrollTop;
      var articleHeight = meta.body.scrollHeight;
      var viewport = window.innerHeight || document.documentElement.clientHeight || 1;
      var effectiveHeight = Math.max(articleHeight - viewport * 0.65, 1);
      var ratio = (scrollTop - articleTop) / effectiveHeight;

      if (ratio >= threshold) {
        sent = true;
        window.removeEventListener('scroll', checkDepth);
        sendArticleEvent('article_read_75', {
          event_category: 'article_engagement',
          event_label: meta.path,
          article_title: meta.title,
          article_path: meta.path,
          article_section: meta.section,
          read_depth: 75
        });
      }
    }

    checkDepth();
    window.addEventListener('scroll', checkDepth, { passive: true });
    window.addEventListener('resize', checkDepth);
    window.addEventListener('orientationchange', checkDepth);
  }

  if ('scrollRestoration' in history) {
    history.scrollRestoration = 'auto';
  }

  // ===== GA4 読者行動イベント（P4-1：2026-11 追加） =====
  function getArticleSlug() {
    var p = (location.pathname || '').replace(/^\/Excel-No-VBA-Lab\//, '').replace(/^\//, '');
    p = p.replace(/index\.html$/, '').replace(/\.html$/, '').replace(/\/$/, '');
    return p || 'home';
  }

  function sendSiteEvent(eventName, params) {
    if (location.protocol === 'file:') return;
    if (typeof window.gtag !== 'function') return;
    var payload = { article_slug: getArticleSlug() };
    Object.keys(params || {}).forEach(function (k) { payload[k] = params[k]; });
    window.gtag('event', eventName, payload);
  }

  function initAnalyticsEvents() {
    var COPY_SELECTOR = '.copy-btn, .formula-copy-btn, [onclick*="copyFormula"]';
    document.addEventListener('click', function (event) {
      var target = event.target;
      if (!target || !target.closest) return;

      // code_copy：コードのコピーボタン（code_index はページ内で何番目のボタンか）
      var copyBtn = target.closest(COPY_SELECTOR);
      if (copyBtn) {
        var buttons = Array.prototype.slice.call(document.querySelectorAll(COPY_SELECTOR));
        sendSiteEvent('code_copy', { code_index: buttons.indexOf(copyBtn) + 1 });
        return;
      }

      var link = target.closest('a[href]');

      // sample_download：サンプルファイルのダウンロード
      if (link && (link.hasAttribute('download') || /\/samples\/[^\/?#]+\.(csv|txt|xlsx|zip)(\?|#|$)/i.test(link.href))) {
        sendSiteEvent('sample_download', { file_name: (link.getAttribute('href') || '').split('/').pop() });
        return;
      }

      // tool_generate：ブラウザツールの生成ボタン（data-ga-tool-generate、または /tools/ の #gen）
      var genBtn = target.closest('[data-ga-tool-generate]');
      if (!genBtn && /\/tools\//.test(location.pathname)) genBtn = target.closest('#gen');
      if (genBtn) {
        sendSiteEvent('tool_generate', { tool_name: genBtn.getAttribute('data-ga-tool-generate') || getArticleSlug() });
        return;
      }

      // next_article_click：シリーズ・関連記事・次の記事へのリンク
      if (link && link.closest('.stage-nav, .series-nav, .related-articles, .related-articles-section, .related, .next-article, [data-ga-next]')) {
        sendSiteEvent('next_article_click', {
          link_url: link.getAttribute('href') || '',
          link_text: (link.textContent || '').trim().slice(0, 80)
        });
      }
    }, true);

    // scroll_90：記事本文の90%まで読んだ
    var body = document.querySelector('.article-body');
    if (!body) return;
    var sent = false;
    function check() {
      if (sent) return;
      var scrollTop = window.pageYOffset || document.documentElement.scrollTop || 0;
      var top = body.getBoundingClientRect().top + scrollTop;
      var viewport = window.innerHeight || document.documentElement.clientHeight || 1;
      var effective = Math.max(body.scrollHeight - viewport * 0.65, 1);
      if ((scrollTop - top) / effective >= 0.9) {
        sent = true;
        window.removeEventListener('scroll', check);
        sendSiteEvent('scroll_90', { read_depth: 90 });
      }
    }
    window.addEventListener('scroll', check, { passive: true });
  }


  function initArticleSearch() {
    var form = document.getElementById('article-search-form');
    if (!form) return;
    var input = document.getElementById('article-search-input');
    var status = document.getElementById('article-search-status');
    var list = document.getElementById('article-search-results');
    var categorySelect = document.getElementById('article-search-category');
    var versionSelect = document.getElementById('article-search-version');
    var finder = document.getElementById('article-finder');
    var syncsUrl = !!document.getElementById('home');
    var requestId = 0;
    var indexed;
    var debounce;
    var PAGE_SIZE = 20;
    function plainText(html) { var box = document.createElement('div'); box.innerHTML = html || ''; return box.textContent || ''; }
    function getIndex(data) {
      if (!indexed) {
        indexed = data.articles.filter(function (article) { return !article.is_prep; }).map(function (article) {
          var title = plainText(article.title);
          var summary = plainText(article.summary);
          return {
            article: article, title: title, summary: summary,
            titleKey: normalizeSearchText(title), summaryKey: normalizeSearchText(summary),
            extraKey: normalizeSearchText([SEARCH_CATEGORIES[article.category], article.category, SEARCH_SUBCATEGORIES[article.subcategory], article.version].concat(article.tags || []).join(' '))
          };
        });
      }
      return indexed;
    }
    function findMatches(entries, query, category, version) {
      return entries.map(function (entry, order) {
        if (category && entry.article.category !== category) return null;
        if (version && !(entry.article.min_excel && entry.article.min_excel <= version)) return null;
        var hay = entry.titleKey + ' ' + entry.summaryKey + ' ' + entry.extraKey;
        if (query.version && !(entry.article.min_excel && entry.article.min_excel <= query.version) && hay.indexOf(String(query.version)) < 0) return null;
        var score = 0;
        for (var i = 0; i < query.terms.length; i++) {
          var term = query.terms[i];
          if (hay.indexOf(term) < 0) return null;
          if (entry.titleKey.indexOf(term) === 0) score += 6;
          if (entry.titleKey.indexOf(term) >= 0) score += 10;
          if (entry.summaryKey.indexOf(term) >= 0) score += 2;
          if (entry.article.category === term) score += 4;
        }
        if (entry.article.is_highlight) score += 1;
        if (entry.article.kind === 'hub' && query.terms.length) score += 8;
        if (entry.article.category === 'thinking') score -= 2;
        if (entry.article.category === 'unlocked') score -= 1;
        return { entry: entry, score: score, order: order };
      }).filter(Boolean).sort(function (a, b) { return (b.score - a.score) || (a.order - b.order); });
    }
    function renderItem(entry) {
      var article = entry.article;
      var item = document.createElement('li'); item.className = 'article-item search-result';
      var link = document.createElement('a'); link.className = 'article-link'; link.href = article.url; link.textContent = entry.title;
      var summary = document.createElement('p'); summary.className = 'article-summary'; summary.textContent = entry.summary;
      var meta = document.createElement('p'); meta.className = 'search-result-meta';
      function addTag(text, className) {
        if (!text) return;
        var tag = document.createElement('span'); tag.textContent = text;
        if (className) tag.className = className;
        meta.appendChild(tag);
      }
      addTag(SEARCH_CATEGORIES[article.category]);
      addTag(article.version);
      if (article.level) addTag('Lv.' + article.level);
      if (article.subcategory === 'vba') addTag('VBA利用環境向け', 'search-tag-env');
      if (article.subcategory === 'ai') addTag('AI利用環境向け', 'search-tag-env');
      item.append(link, meta, summary);
      return item;
    }
    function renderResults(matches, shown) {
      list.replaceChildren();
      matches.slice(0, shown).forEach(function (match) { list.appendChild(renderItem(match.entry)); });
      if (matches.length > shown) {
        var row = document.createElement('li'); row.className = 'search-more';
        var more = document.createElement('button'); more.type = 'button';
        more.textContent = 'さらに表示（残り' + (matches.length - shown) + '件）';
        more.addEventListener('click', function () {
          renderResults(matches, shown + PAGE_SIZE);
          var next = list.children[shown];
          var nextLink = next && next.querySelector('a, button');
          if (nextLink) nextLink.focus();
        });
        row.appendChild(more); list.appendChild(row);
      }
    }
    function describeFilter(query, category, version) {
      var parts = [];
      if (category && SEARCH_CATEGORIES[category]) parts.push(SEARCH_CATEGORIES[category]);
      var excel = version || query.version;
      if (excel) parts.push('Excel ' + excel + 'で使える記事（各記事の対応表示による）');
      return parts.length ? '［' + parts.join('・') + '］' : '';
    }
    function offerWithoutVersion(count) {
      var row = document.createElement('li'); row.className = 'search-more';
      var clear = document.createElement('button'); clear.type = 'button';
      clear.textContent = '対応Excelの指定を外して表示（' + count + '件）';
      clear.addEventListener('click', function () {
        if (versionSelect) versionSelect.value = '';
        input.value = input.value.replace(/(excel\s*)?(2016|2019|2021|2024)/ig, ' ').replace(/\s+/g, ' ').trim();
        run(true);
      });
      row.appendChild(clear); list.appendChild(row);
    }
    function syncUrl(query, category, version) {
      if (!syncsUrl || !history.replaceState) return;
      var params = new URLSearchParams();
      if (query) params.set('q', query);
      if (category) params.set('cat', category);
      if (version) params.set('ver', version);
      var search = params.toString();
      history.replaceState(null, '', window.location.pathname + (search ? '?' + search : '') + window.location.hash);
    }
    async function run(updateUrl) {
      var currentRequest = ++requestId;
      var raw = input.value.trim();
      var query = parseSearchQuery(raw);
      var category = categorySelect ? categorySelect.value : '';
      var version = versionSelect ? +versionSelect.value || 0 : 0;
      if (updateUrl) syncUrl(raw, category, version);
      if (!query.terms.length && !query.version && !category && !version) {
        list.replaceChildren();
        status.textContent = 'キーワードを入力してください。';
        return;
      }
      status.textContent = '記事を検索しています…';
      try {
        var data = await loadCatalog();
        if (currentRequest !== requestId) return;
        var entries = getIndex(data);
        var matches = findMatches(entries, query, category, version);
        renderResults(matches, PAGE_SIZE);
        status.textContent = matches.length
          ? matches.length + '件の記事が見つかりました。' + describeFilter(query, category, version)
          : '一致する記事がありません。短い言葉に変えるか、目的別の入口から探してください。' + describeFilter(query, category, version);
        if (!matches.length && (version || query.version)) {
          var relaxed = findMatches(entries, { terms: query.terms, version: 0 }, category, 0);
          if (relaxed.length) offerWithoutVersion(relaxed.length);
        }
      } catch (error) { if (currentRequest === requestId) status.textContent = '記事一覧を読み込めませんでした。再検索するか、目的別の入口から探してください。'; }
    }
    form.addEventListener('submit', function (event) {
      event.preventDefault();
      window.clearTimeout(debounce);
      run(true);
    });
    input.addEventListener('input', function (event) {
      if (event.isComposing) return;
      window.clearTimeout(debounce);
      if (input.value.trim().length < 2) return;
      debounce = window.setTimeout(function () { run(true); }, 300);
    });
    input.addEventListener('compositionend', function () {
      window.clearTimeout(debounce);
      debounce = window.setTimeout(function () { run(true); }, 300);
    });
    [categorySelect, versionSelect].forEach(function (select) {
      if (select) select.addEventListener('change', function () { run(true); });
    });
    if (finder) {
      finder.querySelectorAll('[data-search-example]').forEach(function (button) {
        button.addEventListener('click', function () {
          input.value = button.getAttribute('data-search-example');
          run(true);
          input.focus();
        });
      });
    }
    if (syncsUrl) {
      var params = new URLSearchParams(window.location.search);
      if (params.get('q') || params.get('cat') || params.get('ver')) {
        input.value = params.get('q') || '';
        if (categorySelect) categorySelect.value = params.get('cat') || '';
        if (versionSelect) versionSelect.value = params.get('ver') || '';
        run(false).then(function () {
          var navEntry = performance.getEntriesByType && performance.getEntriesByType('navigation')[0];
          var state = readIndexState();
          if (navEntry && navEntry.type === 'back_forward' && state) window.scrollTo(0, Math.max(state.scrollY || 0, 0));
        });
      }
    }
  }

  // 記事を探すときの表記ゆれ（全角半角・カタカナ表記・送り仮名）をそろえる。
  var SEARCH_CATEGORIES = { functions: '関数', powerquery: 'Power Query', cobol: 'COBOL帳票', qr: 'QR入力', others: '現場実務', thinking: '技術の流れ', unlocked: 'VBA・AI', resources: '資料' };
  var SEARCH_SUBCATEGORIES = { word: 'Word', print: '印刷', speed: '速度 高速化', pdf: 'PDF', ai: 'AI', vba: 'VBA', data: 'データ整理', office: '録画 手順共有', ui: '操作', migration: '移行 マクロ' };
  var SEARCH_HUBS = { functions: 'functions/', powerquery: 'powerquery/', cobol: 'cobol/', qr: 'qr/', others: 'others/', thinking: 'thinking/', unlocked: 'unlocked/' };
  function normalizeSearchText(value) {
    return String(value || '').normalize('NFKC').toLocaleLowerCase()
      .replace(/power\s*query|パワークエリ|パワクエ/g, 'powerquery')
      .replace(/コボル/g, 'cobol')
      .replace(/ワード/g, 'word')
      .replace(/キューアール/g, 'qr')
      .replace(/差し込み|差込み|差込/g, '差込');
  }
  function parseSearchQuery(raw) {
    var result = { terms: [], version: 0 };
    normalizeSearchText(raw).split(/[\s+×,、・\/]+/).filter(Boolean).forEach(function (token) {
      var match = token.match(/^(?:excel)?(2016|2019|2021|2024)$/);
      if (match) {
        if (result.terms[result.terms.length - 1] === 'excel') result.terms.pop();
        result.version = +match[1];
        return;
      }
      result.terms.push(token === 'pq' ? 'powerquery' : token);
    });
    return result;
  }

  // 公開済みの長い記事は _dev/add_static_toc.py で静的な目次を持つ。これは未処理の新規記事向けの補完。
  function initAutoToc() {
    var body = document.querySelector('.article-body');
    if (!body || body.innerHTML.indexOf('目次') >= 0) return;
    var headings = Array.prototype.filter.call(body.querySelectorAll('h2'), function (heading) {
      return !heading.closest('.related-articles-section, .related-articles') && heading.textContent.indexOf('関連記事') < 0;
    });
    if (headings.length < 4) return;
    var details = document.createElement('details');
    details.className = 'auto-toc';
    details.open = !window.matchMedia('(max-width: 768px)').matches;
    var summary = document.createElement('summary');
    summary.textContent = '目次（' + headings.length + '項目）';
    var nav = document.createElement('nav');
    nav.setAttribute('aria-label', '記事内の目次');
    var list = document.createElement('ol');
    headings.forEach(function (heading, index) {
      if (!heading.id) heading.id = 'sec-' + (index + 1);
      var item = document.createElement('li');
      var link = document.createElement('a');
      link.href = '#' + heading.id;
      link.textContent = heading.textContent.replace(/\s+/g, ' ').trim();
      item.appendChild(link); list.appendChild(item);
    });
    nav.appendChild(list);
    details.append(summary, nav);
    var header = body.querySelector('.article-header');
    if (header && header.parentNode === body) body.insertBefore(details, header.nextSibling);
    else body.insertBefore(details, body.firstChild);
  }

  // シリーズ案内がない記事に、同じ分類の前後の記事とカテゴリ一覧への導線を付ける。
  function initArticleNeighbors() {
    var body = document.querySelector('.article-body');
    if (!body || !document.querySelector('meta[property="og:type"][content="article"]')) return;
    if (body.querySelector('.series-nav, .stage-nav, .nav-series, .next-link, .next-page-link, .prev-page-link, .nav-link-box')) return;
    var current = decodeURI(window.location.pathname).slice(decodeURI(new URL(SITE_ROOT).pathname).length);
    function build() {
      loadCatalog().then(function (data) {
        var list = data.articles.filter(function (article) { return !article.is_prep && !article.kind && article.url.indexOf('?') < 0; });
        var self = list.filter(function (article) { return article.url === current; })[0];
        if (!self || !SEARCH_HUBS[self.category]) return;
        var siblings = list.filter(function (article) { return article.category === self.category && article.subcategory === self.subcategory; });
        var position = siblings.indexOf(self);
        var nav = document.createElement('nav');
        nav.className = 'article-neighbors';
        nav.setAttribute('aria-label', '同じ分類の前後の記事');
        function addLink(article, label, className) {
          if (!article) return;
          var link = document.createElement('a');
          link.className = className;
          link.href = SITE_ROOT + article.url;
          var caption = document.createElement('span'); caption.textContent = label;
          var title = document.createElement('strong');
          var box = document.createElement('div'); box.innerHTML = article.title; title.textContent = box.textContent;
          link.append(caption, title); nav.appendChild(link);
        }
        addLink(siblings[position - 1], '← 前の記事', 'article-neighbor-prev');
        addLink(siblings[position + 1], '次の記事 →', 'article-neighbor-next');
        var hub = document.createElement('a');
        hub.className = 'article-neighbor-hub';
        var isWordGuide = self.subcategory === 'word';
        hub.href = SITE_ROOT + (isWordGuide ? 'others/word/' : SEARCH_HUBS[self.category]);
        hub.textContent = isWordGuide ? 'Word差し込み印刷ガイドへ' : SEARCH_CATEGORIES[self.category] + 'の記事一覧へ';
        nav.appendChild(hub);
        var feedback = body.querySelector('.article-feedback');
        body.insertBefore(nav, feedback && feedback.parentNode === body ? feedback : null);
      }).catch(function () { /* 導線の補助なので、読み込めなければ何も足さない。 */ });
    }
    var footer = document.querySelector('.site-footer');
    if (!('IntersectionObserver' in window) || !footer) { build(); return; }
    var observer = new IntersectionObserver(function (entries) {
      if (!entries.some(function (entry) { return entry.isIntersecting; })) return;
      observer.disconnect();
      build();
    }, { rootMargin: '1200px 0px' });
    observer.observe(footer);
  }

  function initCuratedCounts() {
    document.querySelectorAll('.curated-tabs [data-ranking-trigger]').forEach(function (button) {
      var panel = document.getElementById('ranking-' + button.getAttribute('data-ranking-trigger'));
      var count = button.querySelector('.curated-tab-count');
      if (panel && count) count.textContent = panel.querySelectorAll('.curated-card').length;
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    initArticleSearch();
    initCuratedCounts();
    initLayoutVars();
    initResponsiveTables();
    initAutoToc();
    initArticleNeighbors();
    initTabs();
    updateRarityBadges();
    initCopyButtons();
    initFormulaCopyButtons();
    initReadingProgress();
    initArticleFeedback();
    initArticleReadDepth();
    initAnalyticsEvents();
  });
})();
