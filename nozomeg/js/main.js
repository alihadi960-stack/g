/* Nozom ERP — landing page interactions */
(function () {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  /* Navbar shadow on scroll */
  const navbar = $('#navbar');
  const onScroll = () => navbar.classList.toggle('scrolled', window.scrollY > 8);
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });

  /* Mobile menu */
  const toggle = $('#navToggle');
  const menu = $('#navMenu');
  toggle.addEventListener('click', () => {
    const open = menu.classList.toggle('show');
    toggle.setAttribute('aria-expanded', String(open));
    document.body.style.overflow = open ? 'hidden' : '';
  });
  $$('#navMenu a').forEach(a => a.addEventListener('click', () => {
    menu.classList.remove('show');
    toggle.setAttribute('aria-expanded', 'false');
    document.body.style.overflow = '';
  }));

  /* Mega menu (click to open; closes on outside click / Esc) */
  $$('.has-mega').forEach(li => {
    const btn = $('button', li);
    btn.addEventListener('click', e => {
      e.stopPropagation();
      li.classList.toggle('open');
    });
  });
  document.addEventListener('click', e => {
    if (!e.target.closest('.has-mega')) $$('.has-mega').forEach(li => li.classList.remove('open'));
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      $$('.has-mega').forEach(li => li.classList.remove('open'));
      menu.classList.remove('show');
      document.body.style.overflow = '';
    }
  });

  /* Industries tabs */
  $$('#chips .chip').forEach(chip => chip.addEventListener('click', () => {
    $$('#chips .chip').forEach(c => c.classList.remove('active'));
    chip.classList.add('active');
    const key = chip.dataset.ind;
    $$('.ind-panel').forEach(p => p.classList.toggle('active', p.dataset.panel === key));
  }));

  /* Testimonials slider */
  const slides = $('#slides');
  const cards = $$('.tcard', slides);
  const dots = $('#dots');
  let index = 0;
  const perView = () => (window.innerWidth <= 600 ? 1 : window.innerWidth <= 900 ? 2 : 3);
  const pages = () => Math.ceil(cards.length / perView());
  function render() {
    const n = perView();
    const total = pages();
    if (index >= total) index = 0;
    if (index < 0) index = total - 1;
    cards.forEach((c, i) => {
      const visible = i >= index * n && i < index * n + n;
      c.style.display = visible ? '' : 'none';
    });
    dots.innerHTML = '';
    for (let i = 0; i < total; i++) {
      const d = document.createElement('i');
      if (i === index) d.classList.add('active');
      d.addEventListener('click', () => { index = i; render(); });
      dots.appendChild(d);
    }
  }
  $('#nextSlide').addEventListener('click', () => { index++; render(); });
  $('#prevSlide').addEventListener('click', () => { index--; render(); });
  window.addEventListener('resize', render);
  render();
  let auto = setInterval(() => { index++; render(); }, 6000);
  slides.addEventListener('mouseenter', () => clearInterval(auto));
  slides.addEventListener('mouseleave', () => { auto = setInterval(() => { index++; render(); }, 6000); });

  /* Pricing toggle */
  const fmt = n => Number(n).toLocaleString('en-US');
  $$('#billingToggle button').forEach(b => b.addEventListener('click', () => {
    $$('#billingToggle button').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    const p = b.dataset.period;
    $$('.price b[data-monthly]').forEach(el => { el.textContent = fmt(el.dataset[p]); });
  }));
  $$('.price b[data-monthly]').forEach(el => { el.textContent = fmt(el.dataset.monthly); });

  /* FAQ accordion */
  $$('#faqList .faq-item button').forEach(btn => btn.addEventListener('click', () => {
    const item = btn.parentElement;
    const wasOpen = item.classList.contains('open');
    $$('#faqList .faq-item').forEach(i => i.classList.remove('open'));
    if (!wasOpen) item.classList.add('open');
  }));

  /* Reveal on scroll + counters */
  const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const countUp = el => {
    const to = +el.dataset.to;
    if (reduce) { el.textContent = fmt(to); return; }
    const start = performance.now();
    const dur = 1600;
    const tick = now => {
      const t = Math.min(1, (now - start) / dur);
      const eased = 1 - Math.pow(1 - t, 3);
      el.textContent = fmt(Math.round(to * eased));
      if (t < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  };
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver(entries => {
      entries.forEach(en => {
        if (!en.isIntersecting) return;
        en.target.classList.add('in');
        $$('.count', en.target).forEach(countUp);
        io.unobserve(en.target);
      });
    }, { threshold: 0.05, rootMargin: '0px 0px -5% 0px' });
    $$('.reveal').forEach(el => io.observe(el));
  } else {
    $$('.reveal').forEach(el => el.classList.add('in'));
    $$('.count').forEach(countUp);
  }
})();
