/**
 * ACIES Engineering - Panel Schedules UI Controller
 * Manages panel schedule viewing, circuit rendering, phase balance gauges, and Excel/CAD sync.
 */

(function () {
    let currentProjectId = null;
    let currentProjectTitle = "";
    let currentPanels = [];
    let selectedPanelId = null;

    window.openPanelSchedules = async function (projectId, projectTitle) {
        currentProjectId = projectId;
        currentProjectTitle = projectTitle || projectId;

        const dlg = document.getElementById("panelSchedulesDlg");
        const titleEl = document.getElementById("psModalProjectTitle");
        const badgeEl = document.getElementById("psModalProjectBadge");

        if (titleEl) titleEl.textContent = `Panel Schedules: ${currentProjectTitle}`;
        if (badgeEl) badgeEl.textContent = currentProjectId;

        if (dlg && typeof dlg.showModal === "function") {
            dlg.showModal();
        }

        await loadProjectPanels(currentProjectId);
    };

    async function loadProjectPanels(projectId) {
        const container = document.getElementById("psPanelListContainer");
        const countBadge = document.getElementById("psPanelCountBadge");
        if (container) {
            container.innerHTML = '<div class="ps-loading">Loading panel schedules...</div>';
        }

        try {
            let res;
            if (window.pywebview && window.pywebview.api && window.pywebview.api.get_panel_schedules) {
                res = await window.pywebview.api.get_panel_schedules(projectId);
            } else {
                // Mock / fallback for testing UI without active backend
                res = {
                    status: "success",
                    panels: [
                        {
                            id: "mock-1",
                            panel_name: "LP-1",
                            voltage: "120/208V",
                            phase: 3,
                            wire: 4,
                            main_bus_amps: 225,
                            main_type: "MCB",
                            validation_status: "VALID",
                            diagnostics: [],
                            loadSummary: {
                                phaseAConnectedVA: 8200,
                                phaseBConnectedVA: 8400,
                                phaseCConnectedVA: 8100,
                                totalConnectedVA: 24700,
                                unbalancePercentage: 3.6
                            }
                        }
                    ]
                };
            }

            if (res && res.status === "success") {
                currentPanels = res.panels || [];
                if (countBadge) countBadge.textContent = String(currentPanels.length);
                renderPanelList(currentPanels);

                if (currentPanels.length > 0) {
                    selectPanel(currentPanels[0].id);
                } else {
                    showEmptyState();
                }
            } else {
                if (container) {
                    container.innerHTML = `<div class="ps-error">${res ? res.message : "Failed to load panels"}</div>`;
                }
                showEmptyState();
            }
        } catch (err) {
            console.error("Error loading panel schedules:", err);
            if (container) {
                container.innerHTML = `<div class="ps-error">${err.message}</div>`;
            }
            showEmptyState();
        }
    }

    function renderPanelList(panels) {
        const container = document.getElementById("psPanelListContainer");
        if (!container) return;

        if (!panels || panels.length === 0) {
            container.innerHTML = '<div class="ps-empty-note">No panels found. Click "Sync" to import from Excel.</div>';
            return;
        }

        container.innerHTML = panels.map(p => {
            const isSelected = p.id === selectedPanelId;
            const statusClass = (p.validation_status || "VALID").toLowerCase();
            const unbalance = p.loadSummary && typeof p.loadSummary.unbalancePercentage === "number"
                ? `${p.loadSummary.unbalancePercentage.toFixed(1)}% unbalance`
                : "";

            return `
                <div class="ps-panel-item ${isSelected ? 'is-active' : ''}" data-panel-id="${p.id}" onclick="selectPanel('${p.id}')">
                    <div class="ps-panel-item-header">
                        <span class="ps-panel-name">⚡ ${escapeHtml(p.panel_name)}</span>
                        <span class="ps-badge ${statusClass}">${escapeHtml(p.validation_status || 'VALID')}</span>
                    </div>
                    <div class="ps-panel-item-meta">
                        <span>${escapeHtml(p.voltage || '120/208V')} • ${p.main_bus_amps || 225}A</span>
                        ${unbalance ? `<span class="ps-unbalance-tag">${unbalance}</span>` : ''}
                    </div>
                </div>
            `;
        }).join("");
    }

    window.selectPanel = async function (panelId) {
        selectedPanelId = panelId;
        renderPanelList(currentPanels);

        const emptyEl = document.getElementById("psEmptyState");
        const detailEl = document.getElementById("psDetailContainer");

        try {
            let res;
            if (window.pywebview && window.pywebview.api && window.pywebview.api.get_panel_schedule_detail) {
                res = await window.pywebview.api.get_panel_schedule_detail(panelId);
            } else {
                const found = currentPanels.find(p => p.id === panelId);
                res = {
                    status: "success",
                    panel: {
                        ...found,
                        circuits: generateMockCircuits()
                    }
                };
            }

            if (res && res.status === "success" && res.panel) {
                if (emptyEl) emptyEl.hidden = true;
                if (detailEl) detailEl.hidden = false;
                renderPanelDetail(res.panel);
            }
        } catch (err) {
            console.error("Error selecting panel:", err);
        }
    };

    function renderPanelDetail(panel) {
        // Summary cards
        const voltEl = document.getElementById("psStatVoltage");
        const busEl = document.getElementById("psStatBus");
        const unbalEl = document.getElementById("psStatUnbalance");
        const unbalStatEl = document.getElementById("psStatUnbalanceStatus");
        const validEl = document.getElementById("psStatValidation");

        if (voltEl) voltEl.textContent = `${panel.voltage || '120/208V'}, ${panel.phase || 3}PH ${panel.wire || 4}W`;
        if (busEl) busEl.textContent = `${panel.main_bus_amps || 225}A ${panel.main_type || 'MCB'}`;
        
        const loadSum = panel.loadSummary || {};
        const unbalanceVal = loadSum.unbalancePercentage || 0;
        if (unbalEl) unbalEl.textContent = `${unbalanceVal.toFixed(1)}%`;
        if (unbalStatEl) {
            if (unbalanceVal <= 5.0) {
                unbalStatEl.textContent = "Compliant (≤ 5.0%)";
                unbalStatEl.className = "ps-stat-sub text-success";
            } else if (unbalanceVal <= 10.0) {
                unbalStatEl.textContent = "Moderate Unbalance";
                unbalStatEl.className = "ps-stat-sub text-warning";
            } else {
                unbalStatEl.textContent = "High Unbalance (> 10%)";
                unbalStatEl.className = "ps-stat-sub text-danger";
            }
        }

        if (validEl) {
            const status = panel.validation_status || "VALID";
            validEl.textContent = status;
            validEl.className = `ps-stat-value ps-status-${status.toLowerCase()}`;
        }

        // Diagnostics
        const diagBanner = document.getElementById("psDiagnosticsBanner");
        const diagList = document.getElementById("psDiagnosticsList");
        const diags = panel.diagnostics || [];
        if (diagBanner && diagList) {
            if (diags.length > 0) {
                diagBanner.hidden = false;
                diagList.innerHTML = diags.map(d => `<li>${escapeHtml(d)}</li>`).join("");
            } else {
                diagBanner.hidden = true;
                diagList.innerHTML = "";
            }
        }

        // Circuit Grid (Rows 1 to 21 for 42 circuits)
        const tbody = document.getElementById("psCircuitsTbody");
        if (!tbody) return;

        const circuits = panel.circuits || [];
        const cktMap = {};
        circuits.forEach(c => {
            cktMap[c.circuit_number] = c;
        });

        const rowsHtml = [];
        for (let row = 0; row < 21; row++) {
            const oddNum = row * 2 + 1;
            const evenNum = row * 2 + 2;

            const oddCkt = cktMap[oddNum] || { circuit_number: oddNum, phase_pole: getPhaseFor(row, panel.phase), load_description: "SPARE", load_type: "SPARE", breaker_amps: 20, poles: 1, connected_va: 0 };
            const evenCkt = cktMap[evenNum] || { circuit_number: evenNum, phase_pole: getPhaseFor(row, panel.phase), load_description: "SPARE", load_type: "SPARE", breaker_amps: 20, poles: 1, connected_va: 0 };

            const isOddSpare = !oddCkt.load_description || oddCkt.load_description.toUpperCase() === "SPARE" || oddCkt.load_description.toUpperCase() === "SPACE";
            const isEvenSpare = !evenCkt.load_description || evenCkt.load_description.toUpperCase() === "SPARE" || evenCkt.load_description.toUpperCase() === "SPACE";

            rowsHtml.push(`
                <tr class="${row % 2 === 0 ? 'even-row' : 'odd-row'}">
                    <!-- Left Side (Odds) -->
                    <td class="ps-cell-cnum"><strong>${oddNum}</strong></td>
                    <td class="ps-cell-phase phase-${oddCkt.phase_pole}">${oddCkt.phase_pole}</td>
                    <td class="ps-cell-desc ${isOddSpare ? 'muted' : ''}">${escapeHtml(oddCkt.load_description || 'SPARE')}</td>
                    <td class="ps-cell-type">${escapeHtml(shortType(oddCkt.load_type))}</td>
                    <td class="ps-cell-trip">${isOddSpare ? '-' : oddCkt.breaker_amps}</td>
                    <td class="ps-cell-pole">${isOddSpare ? '-' : oddCkt.poles}</td>
                    <td class="ps-cell-va ${oddCkt.connected_va > 0 ? 'highlight-va' : ''}">${oddCkt.connected_va > 0 ? Math.round(oddCkt.connected_va) : '-'}</td>
                    
                    <!-- Center Bus Pole Indicator -->
                    <td class="ps-bus-divider phase-${oddCkt.phase_pole}">|</td>
                    
                    <!-- Right Side (Evens) -->
                    <td class="ps-cell-va ${evenCkt.connected_va > 0 ? 'highlight-va' : ''}">${evenCkt.connected_va > 0 ? Math.round(evenCkt.connected_va) : '-'}</td>
                    <td class="ps-cell-pole">${isEvenSpare ? '-' : evenCkt.poles}</td>
                    <td class="ps-cell-trip">${isEvenSpare ? '-' : evenCkt.breaker_amps}</td>
                    <td class="ps-cell-type">${escapeHtml(shortType(evenCkt.load_type))}</td>
                    <td class="ps-cell-desc ${isEvenSpare ? 'muted' : ''}">${escapeHtml(evenCkt.load_description || 'SPARE')}</td>
                    <td class="ps-cell-phase phase-${evenCkt.phase_pole}">${evenCkt.phase_pole}</td>
                    <td class="ps-cell-cnum"><strong>${evenNum}</strong></td>
                </tr>
            `);
        }
        tbody.innerHTML = rowsHtml.join("");

        // Totals bar
        const totA = document.getElementById("psTotalA");
        const totB = document.getElementById("psTotalB");
        const totC = document.getElementById("psTotalC");
        const totAll = document.getElementById("psTotalAll");

        if (totA) totA.textContent = `${((loadSum.phaseAConnectedVA || 0) / 1000).toFixed(2)} kVA`;
        if (totB) totB.textContent = `${((loadSum.phaseBConnectedVA || 0) / 1000).toFixed(2)} kVA`;
        if (totC) totC.textContent = `${((loadSum.phaseCConnectedVA || 0) / 1000).toFixed(2)} kVA`;
        if (totAll) totAll.textContent = `${((loadSum.totalConnectedVA || 0) / 1000).toFixed(2)} kVA`;
    }

    function getPhaseFor(rowIdx, phaseCount) {
        if (phaseCount === 1) {
            return rowIdx % 2 === 0 ? "A" : "B";
        }
        const phases = ["A", "B", "C"];
        return phases[rowIdx % 3];
    }

    function shortType(type) {
        if (!type) return "";
        const map = {
            "LIGHTING_CONTINUOUS": "L",
            "RECEPTACLE_NON_CONTINUOUS": "R",
            "MOTOR": "M",
            "HVAC_CONTINUOUS": "AC",
            "KITCHEN_EQUIPMENT": "K",
            "ELECTRIC_HEATING": "H",
            "SPARE": "",
            "SPACE": ""
        };
        return map[type] || type.substring(0, 2);
    }

    function showEmptyState() {
        const emptyEl = document.getElementById("psEmptyState");
        const detailEl = document.getElementById("psDetailContainer");
        if (emptyEl) emptyEl.hidden = false;
        if (detailEl) detailEl.hidden = true;
    }

    function escapeHtml(str) {
        if (!str) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function generateMockCircuits() {
        return [
            { circuit_number: 1, phase_pole: "A", load_description: "Lighting Office 101", load_type: "LIGHTING_CONTINUOUS", breaker_amps: 20, poles: 1, connected_va: 1600 },
            { circuit_number: 2, phase_pole: "A", load_description: "Receptacles Office 101", load_type: "RECEPTACLE_NON_CONTINUOUS", breaker_amps: 20, poles: 1, connected_va: 1800 },
            { circuit_number: 3, phase_pole: "B", load_description: "Lighting Corridor", load_type: "LIGHTING_CONTINUOUS", breaker_amps: 20, poles: 1, connected_va: 1400 },
            { circuit_number: 4, phase_pole: "B", load_description: "Breakroom Outlets", load_type: "RECEPTACLE_NON_CONTINUOUS", breaker_amps: 20, poles: 1, connected_va: 1500 },
            { circuit_number: 5, phase_pole: "C", load_description: "Exhaust Fan EF-1", load_type: "MOTOR", breaker_amps: 20, poles: 1, connected_va: 1200 },
            { circuit_number: 6, phase_pole: "C", load_description: "Copy Machine", load_type: "RECEPTACLE_NON_CONTINUOUS", breaker_amps: 20, poles: 1, connected_va: 1600 }
        ];
    }

    // Attach Event Listeners on DOM Ready
    document.addEventListener("DOMContentLoaded", () => {
        const openToolbarBtn = document.getElementById("openPanelSchedulesBtn");
        if (openToolbarBtn) {
            openToolbarBtn.addEventListener("click", () => {
                // Find selected project in UI or fallback to first available
                const selectedProjEl = document.querySelector(".project-row.is-selected, .project-card.is-selected, .project-row, .project-card");
                const projId = selectedProjEl ? (selectedProjEl.dataset.projectId || selectedProjEl.dataset.id || "Current Project") : "Current Project";
                const projName = selectedProjEl ? (selectedProjEl.querySelector(".project-title, .title")?.textContent || projId) : "Active Project";
                window.openPanelSchedules(projId, projName);
            });
        }

        const syncBtn = document.getElementById("btnSyncProjectPanels");
        if (syncBtn) {
            syncBtn.addEventListener("click", async () => {
                if (!currentProjectId) return;
                syncBtn.disabled = true;
                syncBtn.innerHTML = "⏳ Syncing...";
                try {
                    if (window.pywebview && window.pywebview.api && window.pywebview.api.sync_project_panel_schedules) {
                        const res = await window.pywebview.api.sync_project_panel_schedules(currentProjectId);
                        if (res && res.status === "success") {
                            await loadProjectPanels(currentProjectId);
                        } else {
                            alert(res ? res.message : "Sync failed");
                        }
                    }
                } catch (e) {
                    alert(`Sync error: ${e.message}`);
                } finally {
                    syncBtn.disabled = false;
                    syncBtn.innerHTML = `
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <polyline points="23 4 23 10 17 10"></polyline>
                            <polyline points="1 20 1 14 7 14"></polyline>
                            <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
                        </svg>
                        Sync from Excel / CAD
                    `;
                }
            });
        }
    });
})();
