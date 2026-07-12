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
    action: 'https://docs.google.com/forms/d/e/1FAIpQLSdvSi9yzHNaPpu1HCSbSeP4qJwEfY-g633dnh_SD5Xs9kMmdw/formResponse',
    nameEntry: 'entry.1002031830',
    emailEntry: 'entry.356247462',
    categoryEntry: 'entry.1026400745',
    messageEntry: 'entry.113328341'
  };

  var ARTICLE_FEEDBACK_REACTIONS = {
    helpful: {
      label: '👍 役に立った',
      category: '情報交換・感想'
    },
    correction: {
      label: '✍ 間違い・分かりにくい',
      category: '記事への質問・誤り報告'
    },
    request: {
      label: '💡 続編・別例がほしい',
      category: '記事に関するご質問'
    }
  };

  function appendFeedbackField(form, name, value) {
    var input = document.createElement('input');
    input.type = 'hidden';
    input.name = name;
    input.value = value;
    form.appendChild(input);
  }

  function submitArticleComment(meta, reactionId, comment, allowPublish) {
    var reaction = ARTICLE_FEEDBACK_REACTIONS[reactionId] || ARTICLE_FEEDBACK_REACTIONS.helpful;
    var frameName = 'article-feedback-submit-frame';
    var frame = document.querySelector('iframe[name="' + frameName + '"]');

    if (!frame) {
      frame = document.createElement('iframe');
      frame.name = frameName;
      frame.title = '記事フィードバック送信先';
      frame.hidden = true;
      document.body.appendChild(frame);
    }

    var message = [
      '【記事末尾の匿名フィードバック】',
      '記事名: ' + meta.title,
      '記事URL: ' + location.href.split('#')[0],
      '反応: ' + reaction.label.replace(/^[^\s]+\s*/, ''),
      '匿名での掲載許可: ' + (allowPublish ? 'はい' : 'いいえ'),
      '',
      comment
    ].join('\n');

    var form = document.createElement('form');
    form.method = 'POST';
    form.action = ARTICLE_FEEDBACK_FORM.action;
    form.target = frameName;
    form.hidden = true;

    appendFeedbackField(form, ARTICLE_FEEDBACK_FORM.nameEntry, '匿名フィードバック');
    appendFeedbackField(form, ARTICLE_FEEDBACK_FORM.emailEntry, 'anonymous-feedback@example.invalid');
    appendFeedbackField(form, ARTICLE_FEEDBACK_FORM.categoryEntry, reaction.category);
    appendFeedbackField(form, ARTICLE_FEEDBACK_FORM.messageEntry, message);
    appendFeedbackField(form, 'fvv', '1');
    appendFeedbackField(form, 'pageHistory', '0');

    document.body.appendChild(form);
    form.submit();
    setTimeout(function () { form.remove(); }, 1200);
  }

  function initArticleFeedback() {
    var meta = getArticleMeta();
    if (!meta) return;
    if (meta.body.querySelector('.article-feedback')) return;

    var reactionStorageKey = 'excel_no_vba_lab_feedback_reaction_' + meta.path;
    var legacyStorageKey = 'excel_no_vba_lab_helpful_' + meta.path;
    var commentStorageKey = 'excel_no_vba_lab_feedback_comment_' + meta.path;
    var selectedReaction = safeStorageGet(reactionStorageKey);
    var commentAlreadySent = safeStorageGet(commentStorageKey) === '1';

    if (!ARTICLE_FEEDBACK_REACTIONS[selectedReaction]) {
      selectedReaction = safeStorageGet(legacyStorageKey) === '1' ? 'helpful' : '';
    }

    var feedback = document.createElement('aside');
    feedback.className = 'article-feedback';
    feedback.setAttribute('aria-label', '記事へのフィードバック');

    var textWrap = document.createElement('div');
    textWrap.className = 'article-feedback-text';

    var heading = document.createElement('h2');
    heading.textContent = 'この記事はどうでしたか？';

    var note = document.createElement('p');
    note.textContent = '選ぶだけでも送信されます。文章は任意・匿名・非公開です。';

    textWrap.appendChild(heading);
    textWrap.appendChild(note);

    var optionWrap = document.createElement('div');
    optionWrap.className = 'article-feedback-options';

    var status = document.createElement('p');
    status.className = 'article-feedback-status';
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    status.textContent = selectedReaction ? '反応ありがとうございます。よければ一言もお寄せください。' : '';

    var commentPanel = document.createElement('div');
    commentPanel.className = 'article-feedback-comment';
    commentPanel.hidden = !selectedReaction;

    var commentLabel = document.createElement('label');
    commentLabel.className = 'article-feedback-comment-label';
    commentLabel.textContent = '間違いの指摘、分かりにくかった点、感想など（800文字以内）';

    var textarea = document.createElement('textarea');
    textarea.className = 'article-feedback-textarea';
    textarea.maxLength = 800;
    textarea.rows = 6;
    textarea.placeholder = '例：○○の説明は△△の場合には当てはまらないと思います。／この手順で解決しました。／別の条件の例も読みたいです。';
    textarea.disabled = commentAlreadySent;
    commentLabel.appendChild(textarea);

    var counter = document.createElement('span');
    counter.className = 'article-feedback-counter';
    counter.textContent = '0 / 800';

    textarea.addEventListener('input', function () {
      counter.textContent = textarea.value.length + ' / 800';
    });

    var publishLabel = document.createElement('label');
    publishLabel.className = 'article-feedback-publish';
    var publishCheckbox = document.createElement('input');
    publishCheckbox.type = 'checkbox';
    publishCheckbox.disabled = commentAlreadySent;
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
    submitButton.textContent = commentAlreadySent ? '送信済み' : '匿名で送る';
    submitButton.disabled = commentAlreadySent;

    var detailLink = document.createElement('a');
    detailLink.className = 'article-feedback-detail-link';
    detailLink.href = 'https://forms.gle/6SnCHNWK6EcRy1XG8';
    detailLink.target = '_blank';
    detailLink.rel = 'noopener';
    detailLink.textContent = '返信が必要な相談はこちら';

    submitButton.addEventListener('click', function () {
      var comment = textarea.value.trim();
      if (!comment) {
        status.textContent = '一言入力してから送信してください。反応だけなら、すでに届いています。';
        textarea.focus();
        return;
      }

      if (honeyInput.value) {
        textarea.value = '';
        counter.textContent = '0 / 800';
        status.textContent = '送信ありがとうございます。';
        return;
      }

      if (!navigator.onLine && location.protocol !== 'file:') {
        status.textContent = '通信できませんでした。接続を確認してもう一度お試しください。';
        return;
      }

      if (location.protocol === 'file:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1') {
        status.textContent = 'ローカル確認中のため送信していません。公開ページでは匿名で送信されます。';
        return;
      }

      submitArticleComment(meta, selectedReaction || 'helpful', comment, publishCheckbox.checked);
      safeStorageSet(commentStorageKey, '1');
      textarea.disabled = true;
      publishCheckbox.disabled = true;
      submitButton.disabled = true;
      submitButton.textContent = '送信済み';
      status.textContent = '送信ありがとうございます。内容は公開されず、運営者だけが確認します。';

      sendArticleEvent('article_comment_submit', {
        event_category: 'article_feedback',
        event_label: meta.path,
        article_title: meta.title,
        article_path: meta.path,
        article_section: meta.section,
        feedback_reaction: selectedReaction || 'helpful',
        publish_allowed: publishCheckbox.checked ? 'yes' : 'no'
      });
    });

    commentActions.appendChild(submitButton);
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
        if (reactionId === 'helpful') safeStorageSet(legacyStorageKey, '1');

        optionWrap.querySelectorAll('.article-feedback-option').forEach(function (option) {
          var selected = option.getAttribute('data-reaction') === reactionId;
          option.classList.toggle('selected', selected);
          option.setAttribute('aria-pressed', selected ? 'true' : 'false');
        });

        commentPanel.hidden = false;
        status.textContent = '反応ありがとうございます。よければ一言もお寄せください。';

        if (previousReaction !== reactionId) {
          sendArticleEvent('article_feedback_reaction', {
            event_category: 'article_feedback',
            event_label: meta.path,
            article_title: meta.title,
            article_path: meta.path,
            article_section: meta.section,
            feedback_reaction: reactionId
          });

          if (reactionId === 'helpful') {
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

  document.addEventListener('DOMContentLoaded', function () {
    initLayoutVars();
    initResponsiveTables();
    initTabs();
    updateRarityBadges();
    initCopyButtons();
    initFormulaCopyButtons();
    initReadingProgress();
    initArticleFeedback();
    initArticleReadDepth();
  });
})();
