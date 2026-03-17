/* ===== Header scroll behaviour ===== */
(function () {
    const header = document.querySelector('.header');
    if (!header) return;

    let lastY = window.scrollY;

    function onScroll() {
        const y = window.scrollY;
        if (y > 60) {
            header.classList.add('scrolled');
        } else {
            header.classList.remove('scrolled');
        }
        lastY = y;
    }

    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();                 // run once on load
})();

/* ===== Burger menu toggle ===== */
(function () {
    const burger = document.getElementById('burger');
    if (!burger) return;

    const nav = document.querySelector('.nav');
    if (!nav) return;

    burger.addEventListener('click', function () {
        const isOpen = nav.classList.toggle('open');
        burger.classList.toggle('active', isOpen);
        burger.setAttribute('aria-expanded', isOpen);
        document.body.style.overflow = isOpen ? 'hidden' : '';
    });

    // Close menu when a nav link is clicked
    nav.querySelectorAll('a, button').forEach(function (el) {
        el.addEventListener('click', function () {
            nav.classList.remove('open');
            burger.classList.remove('active');
            burger.setAttribute('aria-expanded', 'false');
            document.body.style.overflow = '';
        });
    });
})();
