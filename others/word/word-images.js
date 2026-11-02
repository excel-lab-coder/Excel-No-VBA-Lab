// Keep the image link usable when JavaScript or <dialog> is unavailable.
(() => {
  const links = document.querySelectorAll('.word-guide .article-screenshot a');
  if (!links.length || !window.HTMLDialogElement) return;

  const viewer = document.createElement('dialog');
  viewer.className = 'word-image-viewer';
  viewer.setAttribute('aria-label', '操作画像の拡大表示');
  viewer.innerHTML = '<div class="word-image-toolbar"><button type="button" data-fit>画面に合わせる</button><button type="button" data-original aria-pressed="false">原寸で読む</button><a target="_blank" rel="noopener">画像を別タブで開く</a><button type="button" data-close autofocus>閉じる</button></div><p class="word-image-description"></p><div class="word-image-scroll" tabindex="0" aria-label="操作画像。原寸表示では縦横にスクロールできます"><img alt=""></div>';
  document.body.append(viewer);
  const photo = viewer.querySelector('img');
  const scroll = viewer.querySelector('.word-image-scroll');
  const caption = viewer.querySelector('.word-image-description');
  const original = viewer.querySelector('[data-original]');
  const fit = viewer.querySelector('[data-fit]');
  let opener;
  let previousOverflow;

  function mode(nativeSize) {
    viewer.classList.toggle('is-original', nativeSize);
    original.setAttribute('aria-pressed', String(nativeSize));
    fit.setAttribute('aria-pressed', String(!nativeSize));
    scroll.scrollTo(0, 0);
  }
  links.forEach(link => link.addEventListener('click', event => {
    if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    const source = link.querySelector('img');
    if (!source) return;
    event.preventDefault();
    opener = link;
    photo.src = link.href;
    photo.alt = source.alt;
    caption.textContent = source.alt;
    viewer.querySelector('a').href = link.href;
    mode(false);
    previousOverflow = document.documentElement.style.overflow;
    viewer.showModal();
    document.documentElement.style.overflow = 'hidden';
  }));
  fit.addEventListener('click', () => mode(false));
  original.addEventListener('click', () => mode(true));
  viewer.querySelector('[data-close]').addEventListener('click', () => viewer.close());
  viewer.addEventListener('click', event => { if (event.target === viewer) viewer.close(); });
  viewer.addEventListener('close', () => {
    document.documentElement.style.overflow = previousOverflow;
    opener?.focus({preventScroll: true});
  });
})();
