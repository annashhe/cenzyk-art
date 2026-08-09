(function () {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;

  document.documentElement.classList.add('motion-on');

  var hero = document.querySelector('.hero, .dossier');
  if (hero) {
    hero.style.opacity = '0';
    hero.style.transform = 'translateY(12px)';
    requestAnimationFrame(function () {
      hero.style.transition = 'opacity .7s ease, transform .7s ease';
      hero.style.opacity = '1';
      hero.style.transform = 'none';
    });
  }

  var revealEls = document.querySelectorAll('.service, .proof article, .steps li, .meeting-card');
  if (!('IntersectionObserver' in window) || !revealEls.length) return;

  revealEls.forEach(function (el, i) {
    el.style.opacity = '0';
    el.style.transform = 'translateY(16px)';
    el.style.transition = 'opacity .55s ease, transform .55s ease';
    el.style.transitionDelay = (i % 4) * 60 + 'ms';
  });

  var io = new IntersectionObserver(
    function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        entry.target.style.opacity = '1';
        entry.target.style.transform = 'none';
        io.unobserve(entry.target);
      });
    },
    { threshold: 0.16 }
  );

  revealEls.forEach(function (el) {
    io.observe(el);
  });
})();
