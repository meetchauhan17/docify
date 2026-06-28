// login.js — Demo login with admin/123
document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('login-form');
    const errorBox = document.getElementById('login-error');
    const errorMsg = document.getElementById('login-error-msg');
    const loginBtn = document.getElementById('login-btn');
    const togglePass = document.getElementById('toggle-pass');
    const passInput = document.getElementById('login-pass');
    const eyeIcon = document.getElementById('eye-icon');

    // Show/hide password
    if (togglePass) {
        togglePass.addEventListener('click', () => {
            const isPass = passInput.type === 'password';
            passInput.type = isPass ? 'text' : 'password';
            eyeIcon.setAttribute('data-lucide', isPass ? 'eye-off' : 'eye');
            lucide.createIcons();
        });
    }

    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const id = document.getElementById('login-id').value.trim();
            const pass = document.getElementById('login-pass').value;

            // Show loading state
            loginBtn.disabled = true;
            loginBtn.querySelector('.btn-text').textContent = 'AUTHENTICATING...';

            setTimeout(() => {
                if (id === 'admin' && pass === '123') {
                    localStorage.setItem('docify_user', 'admin');
                    // Redirect to convert page
                    window.location.href = '/convert';
                } else {
                    errorMsg.textContent = 'ACCESS DENIED.';
                    errorBox.classList.remove('hidden');
                    loginBtn.disabled = false;
                    loginBtn.querySelector('.btn-text').textContent = 'AUTHENTICATE';
                    
                    // Clear inputs
                    document.getElementById('login-id').value = '';
                    document.getElementById('login-pass').value = '';

                    // Shake animation
                    form.classList.add('shake');
                    setTimeout(() => form.classList.remove('shake'), 500);
                }
            }, 800);
        });
    }

    // Hide error on input
    ['login-id', 'login-pass'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', () => errorBox.classList.add('hidden'));
    });

    // Forgot password mock
    document.querySelectorAll('.lf-forgot').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            alert('Demo app — no real password recovery.');
        });
    });
});
