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

    // ── Global State for Editor ──
    let globalParsedData = null;

    // ── Form Submit (Parse Only) ──
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
        if (!formData.has('auto_format'))    formData.set('auto_format', 'false');

        try {
            const response = await fetch('/convert', { method: 'POST', body: formData });

            if (!response.ok) {
                const err = await response.json().catch(() => ({}));
                throw new Error(err.detail || `Server error ${response.status}`);
            }

            const jsonResponse = await response.json();
            if (jsonResponse.success) {
                globalParsedData = jsonResponse;
                showEditor();
            } else {
                throw new Error("Parsing failed on server.");
            }
        } catch (err) {
            console.error('Conversion failed:', err);
            alert('Conversion failed: ' + err.message);
        } finally {
            submitBtn.style.display = 'flex';
            loadingState.classList.add('hidden');
        }
    });

    // ── Split Screen Editor App ──
    const mainApp = document.getElementById('main-app');
    const editorApp = document.getElementById('editor-app');
    const backBtn = document.getElementById('back-btn');
    const downloadFinalBtn = document.getElementById('download-final-btn');
    const sourceImagesContainer = document.getElementById('source-images-container');
    const documentEditor = document.getElementById('document-editor');
    const exportFormatOpts = document.getElementById('export-format');

    function showEditor() {
        mainApp.classList.add('hidden');
        editorApp.classList.remove('hidden');
        
        // Sync the select box default with the initially requested format
        const initFormat = document.getElementById('outtype').value;
        if(initFormat) exportFormatOpts.value = initFormat;

        buildSourceImages();
        buildEditableDocument();
    }

    backBtn.addEventListener('click', () => {
        editorApp.classList.add('hidden');
        mainApp.classList.remove('hidden');
    });

    // We can't actually display local disk paths via typical absolute paths easily due to browser security,
    // but the backend stored the file. Usually we would return a blob URL or base64. 
    // For now, we will create a local object URL of the ORIGINAL uploaded file for reference
    function buildSourceImages() {
        sourceImagesContainer.innerHTML = '';
        if (fileInput.files.length > 0) {
            const file = fileInput.files[0];
            const url = URL.createObjectURL(file);
            // If it's an image, show it. If PDF, maybe iframe or generic icon.
            if (file.type.startsWith('image/')) {
                const img = document.createElement('img');
                img.src = url;
                sourceImagesContainer.appendChild(img);
            } else {
                sourceImagesContainer.innerHTML = `<p style="color:var(--text-secondary); text-align:center; padding: 2rem;"><i>PDF preview not fully supported. Refer to original document.</i></p>`;
            }
        }
    }

    function buildEditableDocument() {
        documentEditor.innerHTML = '';
        if (!globalParsedData || !globalParsedData.data) return;

        let elId = 0;
        globalParsedData.data.forEach((pageData, pageIndex) => {
            if (pageIndex > 0) {
                const hr = document.createElement('hr');
                hr.style.margin = '2rem 0';
                documentEditor.appendChild(hr);
            }

            const elements = pageData[0];
            elements.forEach((el, indexInPage) => {
                const wrapper = document.createElement('div');
                
                if (el.type === 'blank_line') {
                    wrapper.style.height = '12pt';
                } else if (el.type === 'text') {
                    const div = document.createElement('div');
                    div.className = 'doc-el';
                    div.contentEditable = 'true';
                    div.dataset.page = pageIndex;
                    div.dataset.index = indexInPage;
                    div.textContent = el.text;
                    applyStylesToPreview(div, el);
                    wrapper.appendChild(div);
                } else if (el.type === 'checklist') {
                    const div = document.createElement('div');
                    div.className = 'doc-el';
                    div.style.display = 'flex';
                    div.style.alignItems = 'center';
                    div.style.gap = '8px';
                    
                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.checked = !!el.checked;
                    cb.dataset.page = pageIndex;
                    cb.dataset.index = indexInPage;
                    cb.dataset.isCheck = "true";
                    
                    const span = document.createElement('span');
                    span.contentEditable = 'true';
                    span.dataset.page = pageIndex;
                    span.dataset.index = indexInPage;
                    span.textContent = el.text;
                    applyStylesToPreview(span, el);

                    div.appendChild(cb);
                    div.appendChild(span);
                    wrapper.appendChild(div);
                } else if (el.type === 'table') {
                    const table = document.createElement('table');
                    table.className = 'doc-el-table';
                    el.rows.forEach((row, ri) => {
                        const tr = document.createElement('tr');
                        row.forEach((cellData, ci) => {
                            const td = document.createElement('td');
                            td.contentEditable = 'true';
                            td.dataset.page = pageIndex;
                            td.dataset.index = indexInPage;
                            td.dataset.row = ri;
                            td.dataset.col = ci;
                            td.textContent = cellData;
                            tr.appendChild(td);
                        });
                        table.appendChild(tr);
                    });
                    wrapper.appendChild(table);
                } else if (el.type === 'drawing') {
                    const block = document.createElement('div');
                    block.className = 'doc-el-drawing';
                    block.textContent = "🖼️ Image / Drawing Blob";
                    block.style.padding = "2rem";
                    block.style.textAlign = "center";
                    block.style.background = "#f1f5f9";
                    wrapper.appendChild(block);
                } else if (el.type === 'bracket') {
                    const block = document.createElement('div');
                    block.style.fontSize = '36pt';
                    block.style.fontWeight = 'bold';
                    block.textContent = el.bracket_char;
                    block.style.textAlign = el.bracket_side === 'right' ? 'right' : 'left';
                    wrapper.appendChild(block);
                } else if (el.type === 'arrow') {
                    const block = document.createElement('div');
                    block.style.fontSize = '24pt';
                    block.style.textAlign = 'center';
                    block.textContent = el.arrow_char;
                    wrapper.appendChild(block);
                }

                documentEditor.appendChild(wrapper);
            });
        });
    }

    function applyStylesToPreview(node, el) {
        if (el.bold) node.style.fontWeight = 'bold';
        if (el.italic) node.style.fontStyle = 'italic';
        if (el.underline) node.style.textDecoration = 'underline';
        if (el.font_size_pt) node.style.fontSize = el.font_size_pt + 'pt';
        if (el.alignment) node.style.textAlign = el.alignment;
        if (el.left_indent_cm) node.style.marginLeft = (parseFloat(el.left_indent_cm) * 38) + 'px'; // approx cm to px
    }

    // Capture edits and send to render endpoint
    downloadFinalBtn.addEventListener('click', async () => {
        if (!globalParsedData) return;

        // Save back DOM changes into JSON payload
        const editables = documentEditor.querySelectorAll('[contenteditable="true"]');
        editables.forEach(node => {
            const pageId = parseInt(node.dataset.page);
            const idx = parseInt(node.dataset.index);
            const el = globalParsedData.data[pageId][0][idx];

            if (el.type === 'text' || el.type === 'checklist') {
                el.text = node.textContent;
            } else if (el.type === 'table') {
                const r = parseInt(node.dataset.row);
                const c = parseInt(node.dataset.col);
                el.rows[r][c] = node.textContent;
            }
        });
        
        // Save back checkboxes status
        const checkboxes = documentEditor.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(cb => {
            const pageId = parseInt(cb.dataset.page);
            const idx = parseInt(cb.dataset.index);
            globalParsedData.data[pageId][0][idx].checked = cb.checked;
        });

        // Execute Render
        const btnText = downloadFinalBtn.querySelector('.btn-text');
        btnText.textContent = "Rendering...";
        downloadFinalBtn.disabled = true;

        try {
            const reqBody = { ...globalParsedData };
            reqBody.outtype = exportFormatOpts.value;

            const response = await fetch('/render', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(reqBody)
            });

            if (!response.ok) throw new Error(`Server error ${response.status}`);

            let filename = 'document.' + reqBody.outtype;
            const cd = response.headers.get('Content-Disposition');
            if (cd && cd.includes('filename=')) filename = cd.split('filename=')[1].replace(/['"]/g, '');

            const blob = await response.blob();
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            URL.revokeObjectURL(url);
            document.body.removeChild(a);

        } catch (err) {
            console.error('Render failed:', err);
            alert('Render failed: ' + err.message);
        } finally {
            btnText.textContent = "Download Document";
            downloadFinalBtn.disabled = false;
        }
    });

});
