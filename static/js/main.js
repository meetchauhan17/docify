document.addEventListener('DOMContentLoaded', () => {
    const form             = document.getElementById('upload-form');
    const fileInput        = document.getElementById('file-input');
    const dropzone         = document.getElementById('dropzone');
    const filePreview      = document.getElementById('file-preview');
    const fileNameSpan     = document.getElementById('file-name');
    const removeBtn        = document.getElementById('remove-file');
    const submitBtn        = document.getElementById('submit-btn');
    const loadingState     = document.getElementById('loading-state');
    const loadingMsg       = document.getElementById('loading-msg');
    const dropzoneTexts    = dropzone.querySelectorAll('h3, p');

    // Auto-format toggle
    const autoToggle       = document.getElementById('auto-format-toggle');
    const autoBadges       = document.getElementById('auto-badges');
    const manualPanel      = document.getElementById('manual-panel');
    const autoFormatCard   = document.getElementById('auto-format-card');

    // Manual preview
    const swatchText       = document.getElementById('swatch-text');
    const fontFamilySelect = document.getElementById('font_family');
    const fontSizeSlider   = document.getElementById('font_size');
    const lineSpacingSlider= document.getElementById('line_spacing');
    const boldCheck        = document.getElementById('text-bold');
    const italicCheck      = document.getElementById('text-italic');
    const underlineCheck   = document.getElementById('text-underline');
    const alignRadios      = document.querySelectorAll('input[name="text_align"]');

    // ── Toggle: Auto ↔ Manual ──
    function applyToggleState() {
        const isAuto = autoToggle.checked;

        if (isAuto) {
            autoFormatCard.classList.add('is-active');
            autoBadges.style.display = 'flex';
            manualPanel.classList.add('hidden');
            loadingMsg.innerHTML = 'Analysing handwriting formatting with AI...<br><span class="loading-subtext">Font sizes, spacing &amp; alignment are being detected.</span>';
        } else {
            autoFormatCard.classList.remove('is-active');
            autoBadges.style.display = 'none';
            manualPanel.classList.remove('hidden');
            loadingMsg.innerHTML = 'Converting document with AI OCR...<br><span class="loading-subtext">This may take a moment depending on the length.</span>';
        }
    }

    autoToggle.addEventListener('change', applyToggleState);
    applyToggleState(); // initial state

    // ── Live Preview (manual mode) ──
    function updatePreview() {
        if (!swatchText) return;
        const fontFamily  = fontFamilySelect.value;
        const fontSize    = fontSizeSlider.value + 'px';
        const lineSpacing = lineSpacingSlider.value;
        const isBold      = boldCheck.checked;
        const isItalic    = italicCheck.checked;
        const isUnderline = underlineCheck.checked;
        const alignEl     = document.querySelector('input[name="text_align"]:checked');
        const textAlign   = alignEl ? alignEl.value : 'left';

        swatchText.style.fontFamily    = `"${fontFamily}", serif`;
        swatchText.style.fontSize      = fontSize;
        swatchText.style.lineHeight    = lineSpacing;
        swatchText.style.fontWeight    = isBold      ? '700'       : '400';
        swatchText.style.fontStyle     = isItalic    ? 'italic'    : 'normal';
        swatchText.style.textDecoration= isUnderline ? 'underline' : 'none';
        swatchText.style.textAlign     = textAlign;
    }

    if (fontFamilySelect)  fontFamilySelect.addEventListener('change', updatePreview);
    if (fontSizeSlider)    fontSizeSlider.addEventListener('input', updatePreview);
    if (lineSpacingSlider) lineSpacingSlider.addEventListener('input', updatePreview);
    if (boldCheck)         boldCheck.addEventListener('change', updatePreview);
    if (italicCheck)       italicCheck.addEventListener('change', updatePreview);
    if (underlineCheck)    underlineCheck.addEventListener('change', updatePreview);
    alignRadios.forEach(r => r.addEventListener('change', updatePreview));
    updatePreview();

    // ── Drag & Drop ──
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(e => {
        dropzone.addEventListener(e, ev => { ev.preventDefault(); ev.stopPropagation(); }, false);
    });
    ['dragenter', 'dragover'].forEach(e => dropzone.addEventListener(e, () => dropzone.classList.add('dragover'), false));
    ['dragleave', 'drop'].forEach(e => dropzone.addEventListener(e, () => dropzone.classList.remove('dragover'), false));

    fileInput.addEventListener('change', handleFileSelect);
    dropzone.addEventListener('drop', e => {
        const files = e.dataTransfer.files;
        if (files.length > 0) { fileInput.files = files; handleFileSelect(); }
    });

    function handleFileSelect() {
        if (fileInput.files && fileInput.files.length > 0) {
            showFilePreview(fileInput.files[0].name);
            submitBtn.disabled = false;
        } else {
            hideFilePreview();
        }
    }

    function showFilePreview(name) {
        fileNameSpan.textContent = name;
        filePreview.classList.remove('hidden');
        dropzoneTexts.forEach(el => el.style.display = 'none');
        dropzone.classList.add('has-file');
    }

    function hideFilePreview() {
        fileInput.value = '';
        filePreview.classList.add('hidden');
        dropzoneTexts.forEach(el => el.style.display = 'block');
        dropzone.classList.remove('has-file');
        submitBtn.disabled = true;
    }

    removeBtn.addEventListener('click', e => { e.preventDefault(); hideFilePreview(); });

    // ── Form Submit ──
    form.addEventListener('submit', async e => {
        e.preventDefault();
        if (!fileInput.files || fileInput.files.length === 0) return;

        submitBtn.style.display = 'none';
        loadingState.classList.remove('hidden');

        const formData = new FormData(form);

        // Ensure unchecked checkboxes send their false value
        if (!formData.has('text_bold'))      formData.set('text_bold', 'false');
        if (!formData.has('text_italic'))    formData.set('text_italic', 'false');
        if (!formData.has('text_underline')) formData.set('text_underline', 'false');
        // Auto-format toggle: checked = "true" already; unchecked needs explicit "false"
        if (!formData.has('auto_format'))    formData.set('auto_format', 'false');

        try {
            const response = await fetch('/convert', { method: 'POST', body: formData });

            if (!response.ok) {
                const err = await response.json().catch(() => ({}));
                throw new Error(err.detail || `Server error ${response.status}`);
            }

            // Derive a friendly filename from the uploaded file, fallback to output
            const originalName = (fileInput.files && fileInput.files[0]) ? fileInput.files[0].name : 'output';
            const dotIdx = originalName.lastIndexOf('.');
            const baseName = dotIdx !== -1 ? originalName.substring(0, dotIdx) : originalName;
            const ext = formData.get('outtype') || 'pdf';
            const filename = `${baseName}.${ext}`;

            const blob = await response.blob();
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href = url;
            a.setAttribute('download', filename);
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            
            // Delay revoking the Object URL to allow browser to start the download with the filename intact
            setTimeout(() => {
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
            }, 250);

            // Save to history
            try {
                const hist = JSON.parse(localStorage.getItem('docify_history') || '[]');
                const origName = fileInput.files[0]?.name || 'document';
                hist.unshift({ name: origName, format: formData.get('outtype'), date: new Date().toISOString() });
                localStorage.setItem('docify_history', JSON.stringify(hist.slice(0, 30)));
            } catch(_) {}

        } catch (err) {
            console.error('Conversion failed:', err);
            alert('Conversion failed: ' + err.message);
        } finally {
            submitBtn.style.display = 'flex';
            loadingState.classList.add('hidden');
        }
    });
});
