// history.js — Reads conversion history from localStorage
document.addEventListener('DOMContentLoaded', () => {
    const grid = document.getElementById('history-grid');
    const emptyState = document.getElementById('empty-state');
    const clearBtn = document.getElementById('clear-history-btn');
    const searchInput = document.getElementById('history-search');
    const countText = document.getElementById('history-count-text');
    const filterPills = document.querySelectorAll('.filter-pill');
    const filterBar = document.getElementById('filter-bar');

    let activeFilter = 'all';
    let searchQuery = '';

    function getHistory() {
        return JSON.parse(localStorage.getItem('docify_history') || '[]');
    }

    function formatDate(iso) {
        const d = new Date(iso);
        const now = new Date();
        const diff = (now - d) / 1000;
        if (diff < 60) return 'Just now';
        if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
        if (diff < 86400) return Math.floor(diff / 3600) + 'h ago';
        return d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    }

    function getFileExt(name) {
        return (name || '').split('.').pop().toUpperCase() || 'IMG';
    }

    function renderHistory() {
        let items = getHistory();
        if (activeFilter !== 'all') items = items.filter(i => i.format === activeFilter);
        if (searchQuery) items = items.filter(i => i.name.toLowerCase().includes(searchQuery));

        const count = items.length;
        countText.textContent = count === 0 ? 'No conversions' : `${count} conversion${count !== 1 ? 's' : ''}`;

        grid.innerHTML = '';
        if (count === 0) {
            emptyState.style.display = 'flex';
            filterBar.style.display = getHistory().length === 0 ? 'none' : 'flex';
            return;
        }
        emptyState.style.display = 'none';
        filterBar.style.display = 'flex';

        items.forEach((item, idx) => {
            const card = document.createElement('div');
            card.className = 'history-card';
            card.style.animationDelay = (idx * 0.05) + 's';
            const ext = getFileExt(item.name);
            const fmtBadge = item.format === 'pdf'
                ? '<span class="hist-badge hist-badge-pdf">PDF</span>'
                : '<span class="hist-badge hist-badge-docx">DOCX</span>';
            card.innerHTML = `
                <div class="hist-card-icon ${item.format === 'pdf' ? 'hist-icon-pdf' : 'hist-icon-docx'}">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    <span class="hist-ext-label">${ext}</span>
                </div>
                <div class="hist-card-info">
                    <p class="hist-filename" title="${item.name}">${item.name}</p>
                    <p class="hist-meta">
                        ${fmtBadge}
                        <span class="hist-date"><svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>${formatDate(item.date)}</span>
                    </p>
                </div>
                <div class="hist-card-actions">
                    <a href="/convert" class="hist-action-btn" title="New conversion">
                        <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12l7-7 7 7"/></svg>
                        Re-Convert
                    </a>
                    <button class="hist-delete-btn" data-idx="${idx}" title="Remove">
                        <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>
                    </button>
                </div>`;
            grid.appendChild(card);
        });

        // Delete individual items
        grid.querySelectorAll('.hist-delete-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const all = getHistory();
                // map filtered index back to full array
                let filtered = all;
                if (activeFilter !== 'all') filtered = all.filter(i => i.format === activeFilter);
                if (searchQuery) filtered = filtered.filter(i => i.name.toLowerCase().includes(searchQuery));
                const itemToRemove = filtered[+btn.dataset.idx];
                const newAll = all.filter(i => i !== itemToRemove);
                localStorage.setItem('docify_history', JSON.stringify(newAll));
                renderHistory();
                lucide.createIcons();
            });
        });
    }

    // Filter pills
    filterPills.forEach(pill => {
        pill.addEventListener('click', () => {
            filterPills.forEach(p => p.classList.remove('active'));
            pill.classList.add('active');
            activeFilter = pill.dataset.filter;
            renderHistory();
        });
    });

    // Search
    if (searchInput) {
        searchInput.addEventListener('input', () => {
            searchQuery = searchInput.value.trim().toLowerCase();
            renderHistory();
        });
    }

    // Clear all
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            if (confirm('Clear all conversion history? This cannot be undone.')) {
                localStorage.removeItem('docify_history');
                renderHistory();
            }
        });
    }

    renderHistory();
    lucide.createIcons();
});
