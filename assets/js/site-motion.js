(function () {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;

  // Brutal / staccato: hard steps, not soft luxury fades
  var hero = document.querySelector('.hero, .dossier, .chooser-box');
  if (hero) {
    hero.style.opacity = '0';
    hero.style.transform = 'translateY(18px)';
    requestAnimationFrame(function () {
      hero.style.transition = 'opacity .35s steps(3, end), transform .35s steps(3, end)';
      hero.style.opacity = '1';
      hero.style.transform = 'none';
    });
  }

  var revealEls = document.querySelectorAll('.service, .proof article, .steps li, .meeting-card, .price-card');
  if (!('IntersectionObserver' in window) || !revealEls.length) return;

  revealEls.forEach(function (el) {
    el.style.opacity = '0';
    el.style.transform = 'translateY(10px)';
    el.style.transition = 'opacity .28s steps(2, end), transform .28s steps(2, end)';
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
    { threshold: 0.12 }
  );

  revealEls.forEach(function (el) {
    io.observe(el);
  });
})();
