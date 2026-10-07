// Run with: node scripts/verify-site.cjs
// articles.json is the authority for article levels and public search entries.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.resolve(__dirname, '..');
const base = 'https://excel-lab-coder.github.io/Excel-No-VBA-Lab/';
const issues = [];
function walk(dir) {
  return fs.readdirSync(dir, {withFileTypes: true}).flatMap(item =>
    ['.git', 'tmp', 'node_modules', '記事候補_2026-07-15'].includes(item.name) ? [] :
    item.isDirectory() ? walk(path.join(dir, item.name)) : [path.relative(root, path.join(dir, item.name)).replaceAll('\\', '/')]);
}
function attributes(tag) {
  return Object.fromEntries([...tag.matchAll(/(?:^|\s)([\w:-]+)\s*=\s*(['"])([\s\S]*?)\2/g)].map(m => [m[1], m[3].replaceAll('&amp;', '&')]));
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
if (/form\.submit\(\)|article_comment_submit|commentAlreadySent/.test(shared)) issues.push('Unverified hidden comment submission');
console.log(JSON.stringify({pages: pages.size,entries: catalog.length,searchable: catalog.filter(a => !a.is_prep).length,issues}, null, 2));
process.exitCode = issues.length ? 1 : 0;
