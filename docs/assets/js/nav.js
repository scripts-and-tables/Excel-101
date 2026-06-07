// Responsive site menu: hamburger toggle + Modules dropdown.
(function () {
  var header = document.querySelector('.site-header');
  var toggle = document.querySelector('.nav-toggle');
  var dds = Array.prototype.slice.call(document.querySelectorAll('.nav-dd'));

  if (toggle && header) {
    toggle.addEventListener('click', function (e) {
      e.stopPropagation();
      var open = header.classList.toggle('nav-open');
      toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  }

  dds.forEach(function (dd) {
    var btn = dd.querySelector('.nav-dd__btn');
    if (!btn) return;
    btn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      var open = dd.classList.toggle('is-open');
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  });

  // Close dropdowns / mobile menu when clicking outside.
  document.addEventListener('click', function (e) {
    dds.forEach(function (dd) {
      if (!dd.contains(e.target)) {
        dd.classList.remove('is-open');
        var b = dd.querySelector('.nav-dd__btn');
        if (b) b.setAttribute('aria-expanded', 'false');
      }
    });
    if (header && !header.contains(e.target)) {
      header.classList.remove('nav-open');
      if (toggle) toggle.setAttribute('aria-expanded', 'false');
    }
  });

  // Escape closes everything.
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      dds.forEach(function (dd) { dd.classList.remove('is-open'); });
      if (header) header.classList.remove('nav-open');
    }
  });

  // Reset the mobile menu when resizing up to desktop.
  window.addEventListener('resize', function () {
    if (window.innerWidth > 760 && header) {
      header.classList.remove('nav-open');
      if (toggle) toggle.setAttribute('aria-expanded', 'false');
    }
  });
})();
