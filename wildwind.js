// Wildwind · gemensamt skript: mobilmeny, YouTube vid klick, trasiga bilder
(function () {
  // Mobilmeny
  var btn = document.querySelector('.nav-toggle');
  var links = document.querySelector('.nav-links');
  if (btn && links) {
    btn.addEventListener('click', function () {
      var open = links.classList.toggle('open');
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      btn.textContent = open ? '✕' : '☰';
    });
    links.querySelectorAll('a').forEach(function (a) {
      a.addEventListener('click', function () {
        links.classList.remove('open');
        btn.setAttribute('aria-expanded', 'false');
        btn.textContent = '☰';
      });
    });
  }

  // YouTube: visa bara miniatyr tills någon klickar (snabbare sida, inga spårkakor i förväg)
  document.querySelectorAll('.yt[data-id]').forEach(function (el) {
    var id = el.getAttribute('data-id');
    el.style.backgroundImage = 'url(https://i.ytimg.com/vi/' + id + '/hqdefault.jpg)';
    el.addEventListener('click', function () {
      var f = document.createElement('iframe');
      f.src = 'https://www.youtube-nocookie.com/embed/' + id + '?autoplay=1&rel=0';
      f.allow = 'accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture';
      f.allowFullscreen = true;
      f.title = el.getAttribute('aria-label') || 'Video';
      el.innerHTML = '';
      el.appendChild(f);
      el.style.cursor = 'default';
    }, { once: true });
  });

  // Bilder som inte laddar döljs (ytan bakom är sandfärgad)
  document.querySelectorAll('img').forEach(function (img) {
    img.addEventListener('error', function () { img.classList.add('broken'); });
    if (img.complete && img.naturalWidth === 0 && img.src) img.classList.add('broken');
  });

  // Bildspel: byter bild var 5:e sekund, pilar, prickar och svep på mobil
  document.querySelectorAll('.slides').forEach(function (box) {
    var imgs = box.querySelectorAll('img'); if (imgs.length < 2) return;
    var i = 0, timer;
    var dots = document.createElement('div'); dots.className = 'sl-dots';
    imgs.forEach(function (_, n) {
      var b = document.createElement('button'); b.type = 'button'; b.setAttribute('aria-label', 'Bild ' + (n + 1));
      b.addEventListener('click', function () { show(n); restart(); }); dots.appendChild(b);
    });
    function btn(cls, txt, d) { var b = document.createElement('button'); b.type = 'button'; b.className = 'sl-btn ' + cls; b.textContent = txt; b.setAttribute('aria-label', d > 0 ? 'Nästa bild' : 'Föregående bild'); b.addEventListener('click', function () { show(i + d); restart(); }); box.appendChild(b); }
    btn('sl-prev', '‹', -1); btn('sl-next', '›', 1); box.appendChild(dots);
    function show(n) {
      i = (n + imgs.length) % imgs.length;
      imgs.forEach(function (im, k) { im.classList.toggle('on', k === i); });
      dots.querySelectorAll('button').forEach(function (d, k) { d.classList.toggle('on', k === i); });
    }
    function restart() { clearInterval(timer); timer = setInterval(function () { show(i + 1); }, 5000); }
    box.addEventListener('mouseenter', function () { clearInterval(timer); });
    box.addEventListener('mouseleave', restart);
    var x0 = null;
    box.addEventListener('touchstart', function (e) { x0 = e.touches[0].clientX; }, { passive: true });
    box.addEventListener('touchend', function (e) { if (x0 === null) return; var dx = e.changedTouches[0].clientX - x0; if (Math.abs(dx) > 40) { show(i + (dx < 0 ? 1 : -1)); restart(); } x0 = null; });
    show(0); restart();
  });
})();
