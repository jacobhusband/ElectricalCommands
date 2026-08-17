document.addEventListener('DOMContentLoaded', () => {
    // ==========================================
    // 1. FULLSCREEN IMAGE MODAL (with Zoom & Pan)
    // ==========================================
    const modal = document.getElementById('fullscreen-modal');
    const modalImage = document.getElementById('modal-image');
    const modalClose = document.querySelector('.modal-close');
    const modalBackdrop = document.querySelector('.modal-backdrop');

    // Zoom and Pan State
    let scale = 1;
    let panX = 0;
    let panY = 0;
    let isDragging = false;
    let dragStartX = 0;
    let dragStartY = 0;

    const updateImageTransform = () => {
        if (modalImage) {
            modalImage.style.transform = `translate(${panX}px, ${panY}px) scale(${scale})`;
        }
    };

    const resetZoomPan = () => {
        scale = 1;
        panX = 0;
        panY = 0;
        updateImageTransform();
    };

    // Open modal when clicking any image in an img-container
    document.addEventListener('click', (e) => {
        if (e.target.matches('.img-container img') && e.target.src) {
            modalImage.src = e.target.src;
            resetZoomPan();
            modal.classList.remove('hidden');
            document.body.style.overflow = 'hidden';
        }
    });

    // Close modal functions
    const closeModal = () => {
        modal.classList.add('hidden');
        document.body.style.overflow = '';
        resetZoomPan();
    };

    if (modalClose) modalClose.addEventListener('click', closeModal);
    if (modalBackdrop) modalBackdrop.addEventListener('click', (e) => {
        if (!isDragging) closeModal();
    });

    document.addEventListener('keydown', (e) => {
        if (modal.classList.contains('hidden')) return;
        if (e.key === 'Escape') closeModal();
        if (e.key === 'r' || e.key === 'R') resetZoomPan();
    });

    if (modalImage) {
        modalImage.addEventListener('wheel', (e) => {
            e.preventDefault();
            const zoomSpeed = 0.1;
            const delta = e.deltaY > 0 ? -zoomSpeed : zoomSpeed;
            const newScale = Math.min(Math.max(0.5, scale + delta), 10);

            const rect = modalImage.getBoundingClientRect();
            const mouseX = e.clientX - rect.left - rect.width / 2;
            const mouseY = e.clientY - rect.top - rect.height / 2;

            const scaleDiff = newScale - scale;
            panX -= mouseX * scaleDiff / scale;
            panY -= mouseY * scaleDiff / scale;

            scale = newScale;
            updateImageTransform();
        });

        modalImage.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            isDragging = true;
            dragStartX = e.clientX - panX;
            dragStartY = e.clientY - panY;
            modalImage.style.cursor = 'grabbing';
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            panX = e.clientX - dragStartX;
            panY = e.clientY - dragStartY;
            updateImageTransform();
        });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                if (modalImage) modalImage.style.cursor = 'grab';
            }
        });
    }

    // ==========================================
    // 2. PANEL SCHEDULE GENERATOR LOGIC
    // ==========================================
    const btnAddPanel = document.getElementById('btn-add-panel');
    const tabsList = document.getElementById('panel-tabs-list');
    const contentContainer = document.getElementById('panel-content-container');
    const template = document.getElementById('panel-template');

    let panelCounter = 0;

    if (btnAddPanel) {
        btnAddPanel.addEventListener('click', createNewPanel);
        // Create first panel by default
        createNewPanel();
    }

    function createNewPanel() {
        panelCounter++;
        const panelId = `panel-${panelCounter}`;
        const panelName = `Panel ${panelCounter}`;

        // 1. Create Tab Button
        const tabBtn = document.createElement('button');
        tabBtn.className = 'panel-tab-btn';
        tabBtn.textContent = panelName;
        tabBtn.dataset.target = panelId;
        tabBtn.addEventListener('click', () => switchPanel(panelId));

        // 2. Create Panel Content from Template
        const clone = template.content.cloneNode(true);
        const pane = clone.querySelector('.panel-pane');
        pane.id = panelId;

        // --- Logic Init for this Panel Instance ---
        initPanelLogic(pane, tabBtn, panelId);

        // 3. Append to DOM
        tabsList.appendChild(tabBtn);
        contentContainer.appendChild(pane);

        // 4. Switch to it
        switchPanel(panelId);
    }

    function switchPanel(panelId) {
        document.querySelectorAll('.panel-tab-btn').forEach(b => {
            b.classList.toggle('active', b.dataset.target === panelId);
        });
        document.querySelectorAll('.panel-pane').forEach(p => {
            p.classList.toggle('active', p.id === panelId);
            p.classList.toggle('hidden', p.id !== panelId);
        });
    }

    function initPanelLogic(pane, tabBtn, panelId) {
        const nameInput = pane.querySelector('.panel-name-input');
        const btnDelete = pane.querySelector('.btn-delete-panel');
        const btnDone = pane.querySelector('.btn-done');

        // File Handling State
        let breakerFiles = [];
        let directoryFiles = [];

        // Input Mode Selection
        let inputMode = 'field_photos';
        const modeOptions = pane.querySelectorAll('.mode-option');
        const directoryTitle = pane.querySelector('.directory-section-title');
        const directoryDropText = pane.querySelector('.directory-drop-text');

        modeOptions.forEach(opt => {
            opt.addEventListener('click', () => {
                modeOptions.forEach(o => o.classList.remove('active'));
                opt.classList.add('active');
                inputMode = opt.dataset.value;

                if (inputMode === 'existing_directory') {
                    pane.classList.add('mode-existing-directory');
                    if (directoryTitle) directoryTitle.textContent = 'Existing Directory Document';
                    if (directoryDropText) directoryDropText.textContent = 'Drag & Drop Directory Document';
                } else {
                    pane.classList.remove('mode-existing-directory');
                    if (directoryTitle) directoryTitle.textContent = '2. Circuit Directory Picture';
                    if (directoryDropText) directoryDropText.textContent = 'Drag & Drop Directory';
                }
            });
        });

        // Upload Helper
        const setupUpload = (btnClass, inputClass, dropClass, listClass, isBreaker) => {
            const btn = pane.querySelector(btnClass);
            const input = pane.querySelector(inputClass);
            const drop = pane.querySelector(dropClass);
            const list = pane.querySelector(listClass);

            const addFiles = (newFiles) => {
                const targetArr = isBreaker ? breakerFiles : directoryFiles;
                [...newFiles].forEach(f => targetArr.push(f));
                renderList();
            };

            const renderList = () => {
                list.innerHTML = '';
                const targetArr = isBreaker ? breakerFiles : directoryFiles;
                targetArr.forEach((file, idx) => {
                    const li = document.createElement('li');
                    li.innerHTML = `<span>${file.name}</span> <span class="remove-x">&times;</span>`;
                    li.querySelector('.remove-x').addEventListener('click', (e) => {
                        e.stopPropagation();
                        targetArr.splice(idx, 1);
                        renderList();
                    });
                    list.appendChild(li);
                });
            };

            btn.addEventListener('click', () => input.click());
            input.addEventListener('change', (e) => addFiles(e.target.files));

            // Drag and Drop
            drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('dragover'); });
            drop.addEventListener('dragleave', (e) => { e.preventDefault(); drop.classList.remove('dragover'); });
            drop.addEventListener('drop', (e) => {
                e.preventDefault();
                drop.classList.remove('dragover');
                addFiles(e.dataTransfer.files);
            });
        };

        // Setup both sections
        setupUpload('.btn-breaker-upload', '.input-breaker', '.breaker-drop', '.list-breaker', true);
        setupUpload('.btn-directory-upload', '.input-directory', '.directory-drop', '.list-directory', false);

        // Name Update
        nameInput.addEventListener('input', (e) => {
            if (e.target.value.trim()) {
                tabBtn.textContent = e.target.value.trim();
            } else {
                tabBtn.textContent = "Untitled Panel";
            }
        });

        // Delete Panel
        btnDelete.addEventListener('click', () => {
            if (confirm('Are you sure you want to delete this panel tab?')) {
                pane.remove();
                tabBtn.remove();
                const remaining = document.querySelectorAll('.panel-tab-btn');
                if (remaining.length > 0) {
                    switchPanel(remaining[remaining.length - 1].dataset.target);
                }
            }
        });

        // Done / Generate
        btnDone.addEventListener('click', async () => {
            const finalName = nameInput.value.trim();
            if (!finalName) return alert("Please enter a Panel Name.");
            
            if (inputMode === 'existing_directory') {
                if (directoryFiles.length === 0) {
                    return alert("Please upload at least one directory document photo.");
                }
            } else {
                if (breakerFiles.length === 0 || directoryFiles.length === 0) {
                    return alert("Please upload at least one breaker photo and at least one directory photo.");
                }
            }

            // UI State: Loading
            const statusDiv = pane.querySelector('.panel-status');
            const resultDiv = pane.querySelector('.panel-result');
            const statusText = pane.querySelector('.status-text');

            pane.querySelector('.card').classList.add('faded');
            statusDiv.classList.remove('hidden');
            resultDiv.classList.add('hidden');
            statusText.textContent = "Uploading images...";

            const formData = new FormData();
            formData.append('panel_name', finalName);
            formData.append('input_mode', inputMode);
            if (inputMode !== 'existing_directory') {
                breakerFiles.forEach(f => formData.append('breaker_images', f));
            }
            directoryFiles.forEach(f => formData.append('directory_images', f));

            try {
                // Use streaming endpoint
                const res = await fetch('/api/analyze-panel-stream', {
                    method: 'POST',
                    body: formData
                });

                if (!res.ok) throw new Error("Request failed");

                const reader = res.body.getReader();
                const decoder = new TextDecoder("utf-8");
                let buffer = "";

                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;

                    buffer += decoder.decode(value, { stream: true });
                    const parts = buffer.split("\n\n");
                    buffer = parts.pop();

                    for (const part of parts) {
                        if (part.startsWith('data: ')) {
                            try {
                                const eventData = JSON.parse(part.substring(6));

                                if (eventData.event === 'gemini_started') {
                                    statusText.textContent = "Analyzing panel with Gemini AI...";
                                } else if (eventData.event === 'complete') {
                                    statusDiv.classList.add('hidden');
                                    pane.querySelector('.card').classList.remove('faded');
                                    resultDiv.classList.remove('hidden');
                                    // Show the save button now that a panel has been processed
                                    const sessionActions = document.getElementById('session-actions');
                                    if (sessionActions) sessionActions.style.display = '';

                                } else if (eventData.event === 'error') {
                                    alert("Error: " + eventData.detail);
                                    statusDiv.classList.add('hidden');
                                    pane.querySelector('.card').classList.remove('faded');
                                }
                            } catch (e) {
                                console.error('Error parsing SSE:', e);
                            }
                        }
                    }
                }

            } catch (err) {
                console.error(err);
                alert("Network Error");
                statusDiv.classList.add('hidden');
                pane.querySelector('.card').classList.remove('faded');
            }
        });
    }

    // ==========================================
    // 3. SAVE SCHEDULE LOGIC
    // ==========================================
    const btnSaveSchedule = document.getElementById('btn-save-schedule');

    if (btnSaveSchedule) {
        btnSaveSchedule.addEventListener('click', async () => {
            const panelNames = collectPanelNames();
            const suggestedName = buildFilename(panelNames);
            await downloadSchedule(suggestedName);
        });
    }

    function collectPanelNames() {
        const tabButtons = document.querySelectorAll('.panel-tab-btn');
        const names = [];
        tabButtons.forEach(btn => {
            const name = btn.textContent.trim();
            if (name && name !== 'Untitled Panel') {
                const safe = name.replace(/[^a-zA-Z0-9_-]/g, '');
                if (safe) names.push(safe);
            }
        });
        return names;
    }

    function buildFilename(panelNames) {
        if (panelNames.length === 0) {
            return 'Panel_Schedule.xlsx';
        }
        return 'Panels_' + panelNames.join('_') + '.xlsx';
    }

    async function requestHostSaveSchedule(suggestedName, timeoutMs = 15000) {
        if (!window.parent || window.parent === window || !window.parent.postMessage) {
            return null;
        }
        return new Promise((resolve) => {
            const requestId = `cb-save-${Date.now()}-${Math.random().toString(36).slice(2)}`;

            const handleMessage = (event) => {
                const data = event.data || {};
                if (data.type !== 'circuit-breaker-save-schedule-response' || data.requestId !== requestId) {
                    return;
                }
                window.removeEventListener('message', handleMessage);
                resolve(data.result || null);
            };

            window.addEventListener('message', handleMessage);
            window.parent.postMessage({
                type: 'circuit-breaker-save-schedule',
                requestId,
                suggestedName
            }, '*');

            setTimeout(() => {
                window.removeEventListener('message', handleMessage);
                resolve(null);
            }, timeoutMs);
        });
    }

    async function saveScheduleViaHost(suggestedName) {
        if (window.pywebview?.api?.save_circuit_breaker_schedule) {
            try {
                return await window.pywebview.api.save_circuit_breaker_schedule(suggestedName);
            } catch (e) {
                return { status: 'error', message: e?.message || String(e) };
            }
        }
        return await requestHostSaveSchedule(suggestedName);
    }

    async function downloadSchedule(suggestedName) {
        try {
            const hostResult = await saveScheduleViaHost(suggestedName);
            if (hostResult) {
                if (hostResult.status === 'success') {
                    alert(`Saved schedule to:\n${hostResult.path}`);
                    return;
                }
                if (hostResult.status === 'cancelled' || hostResult.status === 'canceled') {
                    return;
                }
                alert(hostResult.message || 'Failed to save schedule.');
                return;
            }

            const response = await fetch('/api/download-schedule');
            if (!response.ok) {
                alert('No schedule file found. Please generate at least one panel first.');
                return;
            }

            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = suggestedName;
            document.body.appendChild(a);
            a.click();

            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 100);
        } catch (err) {
            console.error('Download failed:', err);
            alert('Failed to download file. ' + err.message);
        }
    }

});
