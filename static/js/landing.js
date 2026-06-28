// landing.js — Typewriter, counter animations, scroll navbar
document.addEventListener('DOMContentLoaded', () => {

    // Typewriter effect
    const words = ['handwriting', 'scanned notes', 'PDF pages', 'rough drafts'];
    let wi = 0, ci = 0, deleting = false;
    const el = document.getElementById('typewriter');
    function type() {
        if (!el) return;
        const word = words[wi];
        if (!deleting) {
            el.textContent = word.slice(0, ++ci);
            if (ci === word.length) { deleting = true; setTimeout(type, 1800); return; }
        } else {
            el.textContent = word.slice(0, --ci);
            if (ci === 0) { deleting = false; wi = (wi + 1) % words.length; }
        }
        setTimeout(type, deleting ? 60 : 110);
    }
    type();

    // Counter animation
    document.querySelectorAll('.stat-num').forEach(el => {
        const target = +el.dataset.target;
        let cur = 0;
        const step = Math.ceil(target / 60);
        const timer = setInterval(() => {
            cur = Math.min(cur + step, target);
            el.textContent = cur.toLocaleString();
            if (cur >= target) clearInterval(timer);
        }, 25);
    });
});
