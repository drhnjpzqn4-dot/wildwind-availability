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
})();
