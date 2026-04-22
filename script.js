(function () {
  "use strict";

  const SLIDE_INTERVAL = 5000;

  document.addEventListener("DOMContentLoaded", function () {
    initSlideshow();
    initHeaderScroll();
  });

  function initSlideshow() {
    const slides = document.querySelectorAll(".slide");
    const dots = document.querySelectorAll(".dot");
    if (slides.length === 0) return;

    let current = 0;
    let timer = null;

    function show(index) {
      slides[current].classList.remove("active");
      dots[current] && dots[current].classList.remove("active");
      current = (index + slides.length) % slides.length;
      slides[current].classList.add("active");
      dots[current] && dots[current].classList.add("active");
    }

    function next() {
      show(current + 1);
    }

    function start() {
      stop();
      timer = setInterval(next, SLIDE_INTERVAL);
    }

    function stop() {
      if (timer) {
        clearInterval(timer);
        timer = null;
      }
    }

    dots.forEach(function (dot, i) {
      dot.addEventListener("click", function () {
        show(i);
        start();
      });
    });

    document.addEventListener("visibilitychange", function () {
      if (document.hidden) stop();
      else start();
    });

    start();
  }

  function initHeaderScroll() {
    const header = document.getElementById("siteHeader");
    if (!header) return;

    function update() {
      if (window.scrollY > 40) header.classList.add("scrolled");
      else header.classList.remove("scrolled");
    }

    update();
    window.addEventListener("scroll", update, { passive: true });
  }
})();
