// ── Cyberpunk Effects & Shared Logic ──

document.addEventListener('DOMContentLoaded', () => {
    const hamburger = document.getElementById('nav-hamburger');
    const mobileMenu = document.getElementById('nav-mobile');
    const navbar = document.getElementById('navbar');

    // Mobile hamburger
    if (hamburger && mobileMenu) {
        hamburger.addEventListener('click', () => {
            mobileMenu.classList.toggle('open');
        });
    }

    // Scroll shrink navbar
    if (navbar) {
        window.addEventListener('scroll', () => {
            navbar.classList.toggle('scrolled', window.scrollY > 20);
        });
    }

    // Update login button text if logged in
    const user = localStorage.getItem('docify_user');
    const loginLink = document.getElementById('nav-login-link');
    if (user && loginLink) {
        loginLink.innerHTML = `[ USER: ${user} ]`;
        loginLink.href = '#';
        loginLink.title = 'Logged in as ' + user;
        loginLink.addEventListener('click', (e) => {
            e.preventDefault();
            if (confirm('TERMINATE SESSION?')) {
                localStorage.removeItem('docify_user');
                location.reload();
            }
        });
    }

    // Add Glitch Effect to Elements on Hover
    const glitchBtns = document.querySelectorAll('.cyber-glitch');
    glitchBtns.forEach(btn => {
        btn.addEventListener('mouseenter', () => {
            btn.style.animation = 'glitch 0.3s cubic-bezier(.25, .46, .45, .94) both infinite';
        });
        btn.addEventListener('mouseleave', () => {
            btn.style.animation = 'none';
        });
    });
});
