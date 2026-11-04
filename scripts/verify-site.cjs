// Run with: node scripts/verify-site.cjs
// articles.json is the authority for article levels and public search entries.
// warnings（失敗扱いにしない検査。issues とは別に出力する）:
// W-1 速度の言い切りや記事の口調から浮く強い表現が、本文に残るのを防ぐ。
// W-2 「関連記事」の一覧に、そのページ自身へのリンクが混ざるのを防ぐ。
// W-3 表示の更新日・JSON-LD の dateModified・sitemap.xml の lastmod の食い違いや欠落を防ぐ。
// W-4 「差し込み予定」「準備中」などの作りかけの文言が、公開ページに残るのを防ぐ。
// W-5 articles.json に Lv がある記事のカードで、一覧の Lv バッジが抜けるのを防ぐ。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.resolve(__dirname, '..');
const base = 'https://excel-lab-coder.github.io/Excel-No-VBA-Lab/';
const issues = [];
function walk(dir) {
  return fs.readdirSync(dir, {withFileTypes: true}).flatMap(item =>
    ['.git', '_dev', 'tmp', 'node_modules', '記事候補_2026-07-15'].includes(item.name) ? [] :
    item.isDirectory() ? walk(path.join(dir, item.name)) : [path.relative(root, path.join(dir, item.name)).replaceAll('\\', '/')]);
}
function attributes(tag) {
  return Object.fromEntries([...tag.matchAll(/(?:^|\s)([\w:-]+)\s*=\s*(['"])([\s\S]*?)\2/g)].map(m => [m[1], m[3].replaceAll('&amp;', '&')]));
}
// Inspect file signatures rather than trusting the extension: a JPEG named
// .png previously produced billion-pixel HTML dimensions in the Word guides.
function imageDimensions(filename) {
  const bytes = fs.readFileSync(filename);
  if (bytes.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex'))) {
    return [bytes.readUInt32BE(16), bytes.readUInt32BE(20), 'png'];
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 8 < bytes.length) {
      if (bytes[offset++] !== 0xff) continue;
      while (bytes[offset] === 0xff) offset++;
      const marker = bytes[offset++];
      if (marker === 0xd9 || marker === 0xda) break;
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
      const length = bytes.readUInt16BE(offset);
      if (length < 2) break;
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        return [bytes.readUInt16BE(offset + 5), bytes.readUInt16BE(offset + 3), 'jpg'];
      }
      offset += length;
    }
  }
  throw new Error('Unsupported or invalid image header');
}
function resolve(file, href) {
  if (/^(data:|mailto:|tel:|javascript:)/i.test(href)) return null;
  let url;
  try { url = new URL(href, new URL(file, base)); } catch (_) { return null; }
  if (!url.href.startsWith(base)) return null;
  let target = decodeURIComponent(url.pathname.slice(new URL(base).pathname.length));
  if (!target || target.endsWith('/')) target += 'index.html';
  return {target, key: target + url.search, fragment: decodeURIComponent(url.hash.slice(1))};
}
const files = walk(root);
const pages = new Map(files.filter(f => f.endsWith('.html')).map(f => [f, fs.readFileSync(path.join(root, f), 'utf8')]));
const ids = new Map();
for (const [file, raw] of pages) {
  const html = raw.replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, '');
  if (file.startsWith('others/word/')) {
    for (const match of html.matchAll(/<img\b[^>]*>/g)) {
      const attr = attributes(match[0]);
      const dest = attr.src && resolve(file, attr.src);
      if (!dest || !fs.existsSync(path.join(root, dest.target))) continue;
      if (!/\.(png|jpe?g)$/i.test(dest.target)) continue;
      try {
        const [width, height, format] = imageDimensions(path.join(root, dest.target));
        if (+attr.width !== width || +attr.height !== height) issues.push(`${file}: incorrect image dimensions ${attr.src} (${attr.width}x${attr.height}, actual ${width}x${height})`);
        if (!attr.alt?.trim()) issues.push(`${file}: missing image description ${attr.src}`);
        if ((format === 'png') !== /\.png$/i.test(dest.target)) issues.push(`${file}: image extension does not match contents ${attr.src}`);
      } catch (error) { issues.push(`${file}: ${attr.src}: ${error.message}`); }
    }
  }
  const set = new Set();
  for (const match of html.matchAll(/<(?!\/)[^>]+>/g)) {
    const id = attributes(match[0]).id;
    if (!id) continue;
    if (set.has(id)) issues.push(`${file}: duplicate id ${id}`);
    set.add(id);
  }
  ids.set(file, set);
}
const catalog = JSON.parse(fs.readFileSync(path.join(root, 'articles.json'), 'utf8')).articles;
const byUrl = new Map(catalog.map(entry => [entry.url, entry]));
for (const entry of catalog) {
  if (!fs.existsSync(path.join(root, entry.url.split('?')[0]))) issues.push(`Missing catalog file: ${entry.url}`);
  if (entry.url.includes('11_qr_vba_generate')) issues.push('Unpublished article in search');
}
for (const [file, raw] of pages) {
  for (const match of raw.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)) {
    const attr = attributes(match[1]);
    try {
      if (attr.type === 'application/ld+json') JSON.parse(match[2]);
      else if (!attr.src && (!attr.type || attr.type === 'text/javascript')) new vm.Script(match[2]);
    } catch (error) { issues.push(`${file}: invalid script: ${error.message}`); }
  }
  const html = raw.replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, '');
  const robotsTag = [...html.matchAll(/<meta\b[^>]*>/g)].map(m => attributes(m[0])).find(a => a.name === 'robots');
  const redirect = /http-equiv=["']refresh/i.test(html);
  if (/class=["'][^"']*article-body/.test(html) && !redirect && !robotsTag?.content?.includes('noindex') &&
      !['about.html', 'privacy.html', 'disclaimer.html'].includes(file) && !byUrl.has(file)) issues.push(`${file}: missing from search catalog`);
  const difficulty = [...html.matchAll(/<meta\b[^>]*>/g)].map(m => attributes(m[0])).find(a => a.name === 'difficulty');
  const entry = byUrl.get(file);
  if (difficulty && entry?.level != null && +difficulty.content !== entry.level) issues.push(`${file}: difficulty mismatch`);
  for (const match of html.matchAll(/<(?:a|link|script|img|source|video)\b[^>]*>/g)) {
    const attr = attributes(match[0]);
    const references = [attr.href, attr.src, attr.poster, ...(attr.srcset ? attr.srcset.split(',').map(s => s.trim().split(/\s+/)[0]) : [])].filter(Boolean);
    for (const href of references) {
      const dest = resolve(file, href);
      if (!dest) continue;
      if (!fs.existsSync(path.join(root, dest.target))) issues.push(`${file}: missing link ${href}`);
      else if (dest.fragment && ids.has(dest.target) && !ids.get(dest.target).has(dest.fragment)) issues.push(`${file}: missing anchor ${href}`);
    }
  }
  for (const match of html.matchAll(/<li\b[^>]*class=["']article-item[^"']*["'][^>]*>[\s\S]*?<\/li>/g)) {
    const link = [...match[0].matchAll(/<a\b[^>]*>/g)].map(m => attributes(m[0])).find(a => a.class === 'article-link');
    const dest = link && resolve(file, link.href);
    const expected = dest && byUrl.get(dest.key)?.level;
    const level = match[0].match(/class=["']level-label["'][^>]*>Lv\.(\d)/)?.[1];
    if (expected != null && level && +level !== expected) issues.push(`${file}: list level mismatch for ${dest.target}`);
  }
  if (/人に1人|ネット上にほぼ情報なし|専門書にも載っていない|情報希少だが効果が高い/.test(html)) issues.push(`${file}: unsupported rarity wording`);
  const ogImage = [...html.matchAll(/<meta\b[^>]*>/g)].map(m => attributes(m[0])).find(a => a.property === 'og:image')?.content;
  const ogTarget = ogImage && resolve(file, ogImage);
  if (ogTarget && !fs.existsSync(path.join(root, ogTarget.target))) issues.push(`${file}: missing og:image ${ogImage}`);
  if (!redirect && !robotsTag?.content?.includes('noindex') && byUrl.has(file) && !ogImage) issues.push(`${file}: missing og:image`);
}
const index = pages.get('index.html');
for (const kind of ['combined', 'rare', 'impact']) {
  const cards = (index.match(new RegExp('<article class="curated-card curated-card-' + kind + '"', 'g')) || []).length;
  const label = index.match(new RegExp('data-ranking-trigger="curated-' + kind + '"[^>]*>[\\s\\S]*?curated-tab-count">(\\d+)'))?.[1];
  if (+label !== cards) issues.push(`Curated ${kind}: ${label} labels / ${cards} cards`);
}
if (/TAKE\(/.test(pages.get('functions/latest/article_let_iferror.html'))) issues.push('TAKE in the Excel 2021 LET article');
const shared = fs.readFileSync(path.join(root, 'script.js'), 'utf8');
new vm.Script(shared);
new vm.Script(fs.readFileSync(path.join(root, 'others/word/word-images.js'), 'utf8'));
if (/form\.submit\(\)|article_comment_submit|commentAlreadySent/.test(shared)) issues.push('Unverified hidden comment submission');

// ---- warnings: exitCode には影響しない ----
const warnings = [];
// W-1 の例外。例: {file: 'functions/latest/article_filter.html', word: '劇的'}
const STRONG_WORD_ALLOWLIST = [];
const STRONG_WORDS = ['定数時間', '一瞬で', '1万行で一瞬', '歴然', '劇的', '絶対に壊れない', '完全に解消', '完全に代替', '狂気', '呪文', '駆逐', '最凶', '侵食', '激重', 'フリーズ級'];
const UNFINISHED_WORDS = ['差し込み予定', 'スクリーンショット予定', '準備中', 'TODO', '後で書く'];
const stripNonContent = raw => raw
  .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
  .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
  .replace(/<!--[\s\S]*?-->/g, '');
const bodyOf = html => html.match(/<body\b[^>]*>([\s\S]*)<\/body>/i)?.[1] ?? html;
const toText = html => html.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ');
function findWords(file, text, words, check) {
  for (const word of words) {
    let at = text.indexOf(word);
    while (at !== -1) {
      const context = text.slice(Math.max(0, at - 20), at + word.length + 20);
      warnings.push(`${check} ${file}: 「${word}」 … ${context}`);
      at = text.indexOf(word, at + word.length);
    }
  }
}
const lastmods = new Map([...fs.readFileSync(path.join(root, 'sitemap.xml'), 'utf8').matchAll(/<loc>([^<]+)<\/loc>\s*<lastmod>([^<]+)<\/lastmod>/g)].map(m => {
  const relative = decodeURI(m[1].trim()).slice(base.length);
  return [!relative || relative.endsWith('/') ? relative + 'index.html' : relative, m[2].trim().slice(0, 10)];
}));
const pad = value => value.padStart(2, '0');
for (const [file, raw] of pages) {
  const metas = [...raw.replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, '').matchAll(/<meta\b[^>]*>/g)].map(m => attributes(m[0]));
  if (metas.find(a => a.name === 'robots')?.content?.includes('noindex') || /http-equiv=["']refresh/i.test(raw)) continue;
  const clean = stripNonContent(raw);
  const body = bodyOf(clean);
  const text = toText(body);

  findWords(file, text, STRONG_WORDS.filter(word => !STRONG_WORD_ALLOWLIST.some(a => a.file === file && a.word === word)), 'W-1');

  for (const heading of body.matchAll(/<h([1-6])\b[^>]*>((?:(?!<\/h[1-6]>)[\s\S])*?関連記事(?:(?!<\/h[1-6]>)[\s\S])*?)<\/h\1>/g)) {
    const rest = body.slice(heading.index + heading[0].length);
    const section = rest.slice(0, rest.search(/<h[1-6]\b|<\/section>/) === -1 ? rest.length : rest.search(/<h[1-6]\b|<\/section>/));
    for (const link of section.matchAll(/<a\b[^>]*>/g)) {
      const href = attributes(link[0]).href;
      const dest = href && resolve(file, href);
      if (dest?.target === file) warnings.push(`W-2 ${file}: 関連記事に自己リンク ${href}`);
    }
  }

  const isListPage = /(^|\/)index\.html$/.test(file);
  if (!isListPage && !['about.html', 'privacy.html', 'disclaimer.html', '404.html'].includes(file) && (byUrl.has(file) || /class=["'][^"']*article-body/.test(body))) {
    const shown = text.match(/(?:最終)?更新日?\s*[:：]\s*(\d{4})[-\/年.](\d{1,2})[-\/月.](\d{1,2})/);
    const displayed = shown ? `${shown[1]}-${pad(shown[2])}-${pad(shown[3])}` : null;
    const modified = raw.match(/"dateModified"\s*:\s*"([^"]+)"/)?.[1].slice(0, 10) ?? null;
    const lastmod = lastmods.get(file) ?? null;
    const missing = [displayed ? null : '表示の更新日', modified ? null : 'dateModified', lastmod ? null : 'sitemap lastmod'].filter(Boolean);
    if (missing.length) warnings.push(`W-3 ${file}: 欠落 ${missing.join('・')}（表示 ${displayed ?? '-'} / dateModified ${modified ?? '-'} / lastmod ${lastmod ?? '-'}）`);
    if (new Set([displayed, modified, lastmod].filter(Boolean)).size > 1) warnings.push(`W-3 ${file}: 日付の不一致（表示 ${displayed ?? '-'} / dateModified ${modified ?? '-'} / lastmod ${lastmod ?? '-'}）`);
  }

  // カテゴリページの「準備中」バッジ（badge-prep）は明示された予告なので対象外。
  findWords(file, toText(body.replace(/<span\b[^>]*class=["'][^"']*badge-prep[^"']*["'][^>]*>[\s\S]*?<\/span>/g, '')), UNFINISHED_WORDS, 'W-4');

  // Lv バッジを使っていない一覧（troubleshoot、samples など）は対象外。
  // 同じ記事がページ内の別カードで Lv 付きなら、案内用のリンクとみなして出さない。
  if (isListPage && /class=["'](?:level-label|rank-badge)\b/.test(body)) {
    const cards = new Map();
    for (const card of body.matchAll(/<li\b[^>]*class=["']article-item[^"']*["'][^>]*>[\s\S]*?<\/li>|<article\b[^>]*class=["']curated-card\b[^"']*["'][^>]*>[\s\S]*?<\/article>/g)) {
      const link = [...card[0].matchAll(/<a\b[^>]*>/g)].map(m => attributes(m[0])).find(a => a.href && (a.class === 'article-link' || card[0].startsWith('<article')));
      const dest = link && resolve(file, link.href);
      const expected = dest && byUrl.get(dest.key)?.level;
      if (expected == null) continue;
      const label = card[0].match(/class=["'](?:level-label|rank-badge)[^"']*["'][^>]*>([^<]*)/)?.[1].trim() ?? '';
      const state = cards.get(dest.key) ?? {expected, hasBadge: false, labels: new Set()};
      // Lv.7・Lv.8 の記事は「仕組みを解説」「数式例・解説あり」と表示する方針なので、Lv バッジと同じ扱いにする。
      if (/Lv\.\d/.test(label) || (expected >= 7 && ['仕組みを解説', '数式例・解説あり'].includes(label))) state.hasBadge = true;
      state.labels.add(label || 'なし');
      cards.set(dest.key, state);
    }
    for (const [key, state] of cards) {
      if (!state.hasBadge) warnings.push(`W-5 ${file}: Lv バッジなし ${key}（articles.json: Lv.${state.expected} / 表示: ${[...state.labels].join('、')}）`);
    }
  }
}
console.log(JSON.stringify({pages: pages.size,entries: catalog.length,searchable: catalog.filter(a => !a.is_prep).length,issues,warnings}, null, 2));
process.exitCode = issues.length ? 1 : 0;
