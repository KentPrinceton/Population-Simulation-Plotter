    // Global State
    let globalTrialsData = null;
    let globalObsData = [];
    let isLogScale = false;
    let currentTab = 'view-profile';
    let globalBEData = []; // Store BE table results for export
    let simRefFileCount = 0;
    let simTestFileCount = 0;
    let obsFileCount = 0;
    let beReferenceIndex = '';
    let beTestIndex = '';
    let hasReviewedResults = false;
    let beNeedsRerun = false;
    let ingestDiagnostics = [];
    let hasAutoOpenedAxes = false;
    let boxResultsOnly = false;
    let boxPlotRenderToken = 0;
    let boxPlotFrameId = null;
    let boxPlotLoadingTimer = null;
    let boxPlotLoadingSeq = 0;
    let targetConcUnit = null; // canonical unit all trials are normalised to; set on first ingest
    const columnMappingDefaults = {
        simTimeContains: 'simtime',
        simSubjectPrefix: 'S-',
        obsTimeContains: 'time',
        obsConcContains: 'cp|conc|obs'
    };

    // ── Concentration Unit Normalisation ──────────────────────────────────────
    // Factors: how many ug/mL equals 1 of this unit (i.e. multiply by this to get ug/mL)
    const CONC_UNIT_FACTORS = {
        'pg/ml': 1e-6,
        'ng/ml': 1e-3,
        'ug/ml': 1,
        'mcg/ml': 1,
        'mg/l': 1,
        'mg/ml': 1e3,
        'g/l': 1e3,
        'g/ml': 1e6,
        'ug/dl': 0.1,
        'ng/dl': 1e-4,
        'mg/dl': 1e2,
    };

    function normConcUnit(u) {
        return String(u || '')
            .toLowerCase()
            .replace(/[µμ]g/g, 'ug')
            .replace(/mcg/g, 'ug')
            .replace(/\s+/g, '');
    }

    // Detect concentration unit from GastroPlus sheet name or column header strings.
    // Returns normalised string like 'ng/ml' or null if not found.
    function detectConcUnit(sheetName, fieldNames) {
        const pat = /\b(pg|ng|[uµμ]g|mcg|mg|g)\/(m[lL]|d[lL]|[lL])\b/i;
        for (const s of [sheetName, ...(fieldNames || [])].filter(Boolean)) {
            const m = String(s).match(pat);
            if (m) return normConcUnit(m[0]);
        }
        return null;
    }

    // Returns factor to multiply values by to convert FROM fromUnit TO toUnit.
    function concConvFactor(fromUnit, toUnit) {
        if (!fromUnit || !toUnit) return 1;
        const f = CONC_UNIT_FACTORS[normConcUnit(fromUnit)];
        const t = CONC_UNIT_FACTORS[normConcUnit(toUnit)];
        if (!f || !t) return 1;
        return f / t;
    }

    // Returns true for PK parameter names whose values scale with concentration unit
    // (Cmax, AUC, Css, etc.). Excludes time-only params (Tmax, t1/2, MRT, kel, etc.).
    function isConcDepParam(name) {
        const n = String(name || '').toLowerCase().replace(/\s+/g, '');
        return /cmax|cmin|css|cpeak|ctrough|auc|aumc|\bcp[^a-z]|\bconc/.test(n);
    }

    // Match numeric tokens including grouped thousands and scientific notation.
    const NUMERIC_TOKEN_REGEX = /[-+]?(?:\d{1,3}(?:[.,\s\u00A0]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)(?:e[+-]?\d+)?/i;

    function scaleNumericTokenString(text, factor) {
        if (factor === 1) return text;
        return String(text).replace(new RegExp(NUMERIC_TOKEN_REGEX.source, 'gi'), (token) => {
            const num = parseNumericCell(token);
            if (!Number.isFinite(num)) return token;
            const scaled = num * factor;
            return Number.isFinite(scaled) ? String(scaled) : token;
        });
    }

    function scaleStatsCellValueByKey(key, value, factor) {
        if (factor === 1) return value;
        const nk = normalizeKey(key);

        // Unitless/log-space fields must not be scaled.
        if (nk.includes('cv') || nk.includes('%') || nk.includes('ln')) return value;

        // Scale arithmetic CI ranges by scaling each numeric token in-place.
        if (nk.includes('ci')) return scaleNumericTokenString(value, factor);

        // Scale only concentration-valued summary columns.
        const shouldScaleNumeric = (
            nk === 'mean' ||
            nk === 'min' ||
            nk === 'max' ||
            (nk.includes('geom') && !nk.includes('cv')) ||
            nk.includes('median')
        );
        if (!shouldScaleNumeric) return value;

        const parsed = parseNumericCell(value);
        return Number.isFinite(parsed) ? parsed * factor : value;
    }
    // ──────────────────────────────────────────────────────────────────────────
    const SESSION_KEY = 'gastroplus_dashboard_session_v4';
    const THEME_KEY = 'gastroplus_dashboard_theme_v1';

    // Collapsible sidebar sections
    function toggleSection(bodyId, btnEl) {
        const body = document.getElementById(bodyId);
        const chevron = btnEl.querySelector('.collapse-chevron');
        body.classList.toggle('collapsed');
        if (chevron) chevron.classList.toggle('collapsed');
    }

    function getTrialLabel(trial) {
        if (!trial) return '';
        const custom = typeof trial.displayName === 'string' ? trial.displayName.trim() : '';
        if (custom) return custom;
        const prefix = trial.formulationType === 'reference' ? 'Ref' : 'Test';
        const num = String(trial.trialNumber).padStart(2, '0');
        return `${prefix} Trial ${num}`;
    }

    function escapeHtml(value) {
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function getMappingConfig() {
        return {
            simTimeContains: (mapSimTimeContains && mapSimTimeContains.value ? mapSimTimeContains.value : columnMappingDefaults.simTimeContains).trim(),
            simSubjectPrefix: (mapSimSubjectPrefix && mapSimSubjectPrefix.value ? mapSimSubjectPrefix.value : columnMappingDefaults.simSubjectPrefix).trim(),
            obsTimeContains: (mapObsTimeContains && mapObsTimeContains.value ? mapObsTimeContains.value : columnMappingDefaults.obsTimeContains).trim(),
            obsConcContains: (mapObsConcContains && mapObsConcContains.value ? mapObsConcContains.value : columnMappingDefaults.obsConcContains).trim()
        };
    }

    function splitKeywords(rawValue, fallback) {
        const src = (rawValue || fallback || '').toLowerCase();
        const parts = src.split(/[|,]/).map(s => s.trim()).filter(Boolean);
        return parts.length ? parts : [String(fallback || '').toLowerCase()];
    }

    function updateFileBadges() {
        const simRefBadge = document.getElementById('simRefFileBadge');
        const simTestBadge = document.getElementById('simTestFileBadge');
        const obsBadge = document.getElementById('obsFileBadge');
        if (simRefFileCount > 0) {
            simRefBadge.textContent = simRefFileCount + ' ref';
            simRefBadge.classList.remove('hidden');
        } else {
            simRefBadge.classList.add('hidden');
        }
        if (simTestFileCount > 0) {
            simTestBadge.textContent = simTestFileCount + ' test';
            simTestBadge.classList.remove('hidden');
        } else {
            simTestBadge.classList.add('hidden');
        }
        if (obsFileCount > 0) {
            obsBadge.textContent = obsFileCount + ' obs';
            obsBadge.classList.remove('hidden');
        } else {
            obsBadge.classList.add('hidden');
        }
    }

    function updateObsFooter() {
        const el = document.getElementById('obsCountFooter');
        const count = document.getElementById('statObsCount');
        if (globalObsData.length > 0) {
            el.classList.remove('hidden');
            count.textContent = globalObsData.length;
        } else {
            el.classList.add('hidden');
        }
        updateRunStatBar();
    }

    // Distinct Colors for Multiple Trials
    const chartColors = [
        { rgb: '37, 99, 235', hex: '#2563eb' },   // Blue
        { rgb: '220, 38, 38', hex: '#dc2626' },   // Red
        { rgb: '22, 163, 74', hex: '#16a34a' },   // Green
        { rgb: '147, 51, 234', hex: '#9333ea' },  // Purple
        { rgb: '234, 88, 12', hex: '#ea580c' },   // Orange
        { rgb: '13, 148, 136', hex: '#0d9488' },  // Teal
        { rgb: '219, 39, 119', hex: '#db2777' },  // Pink
        { rgb: '202, 138, 4', hex: '#ca8a04' }    // Yellow
    ];

    // DOM Elements
    const fileInputRef = document.getElementById('fileInputRef');
    const fileInputTest = document.getElementById('fileInputTest');
    const obsFileInput = document.getElementById('obsFileInput');
    
    // Aesthetic Overlays
    const showContour100 = document.getElementById('showContour100');
    const showContour95 = document.getElementById('showContour95');
    const showContour90 = document.getElementById('showContour90');
    const showContoursCheck = document.getElementById('showContours'); // 50%
    const showMean = document.getElementById('showMean');
    const showMedian = document.getElementById('showMedian');
    const showCIMean = document.getElementById('showCIMean');
    const showIndividualsCheck = document.getElementById('showIndividuals');
    
    // Aesthetic Toggles
    const btnScaleLinear = document.getElementById('btnScaleLinear');
    const btnScaleLog = document.getElementById('btnScaleLog');
    const boxShowOutliers = document.getElementById('boxShowOutliers');
    const boxShowMean = document.getElementById('boxShowMean');
    const btnBoxResultsOnly = document.getElementById('btnBoxResultsOnly');
    
    // Trial Panel
    const trialSelectionPanel = document.getElementById('trialSelectionPanel');
    const trialList = document.getElementById('trialList');
    const btnSelectAll = document.getElementById('btnSelectAll');
    const btnSelectNone = document.getElementById('btnSelectNone');
    
    // Axes & Selectors
    const inputXTitle = document.getElementById('inputXTitle');
    const inputYTitle = document.getElementById('inputYTitle');
    const inputXMax = document.getElementById('inputXMax');
    const inputYMax = document.getElementById('inputYMax');
    const targetConcUnitSelect = document.getElementById('targetConcUnitSelect');
    const mapSimTimeContains = document.getElementById('mapSimTimeContains');
    const mapSimSubjectPrefix = document.getElementById('mapSimSubjectPrefix');
    const mapObsTimeContains = document.getElementById('mapObsTimeContains');
    const mapObsConcContains = document.getElementById('mapObsConcContains');
    const beRefTrialSelect = document.getElementById('beRefTrial');
    const beTestTrialSelect = document.getElementById('beTestTrial');
    const beMethodSelect = document.getElementById('beMethodSelect');
    const btnRunBE = document.getElementById('btnRunBE');
    const beMethodHint = document.getElementById('beMethodHint');
    const bePairingPanel = document.getElementById('bePairingPanel');
    const bePairingSummary = document.getElementById('bePairingSummary');
    const bePairingDetail = document.getElementById('bePairingDetail');

    // Display / Export
    const emptyState = document.getElementById('emptyState');
    const loadingOverlay = document.getElementById('loadingOverlay');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const statSubjects = document.getElementById('statSubjects');
    const statTrials = document.getElementById('statTrials');
    const statActiveTrials = document.getElementById('statActiveTrials');
    const btnExportPNG = document.getElementById('btnExportPNG');
    const btnExportCSV = document.getElementById('btnExportCSV');
    const btnExportBundle = document.getElementById('btnExportBundle');
    const txtExportBundle = document.getElementById('txtExportBundle');
    const btnThemeToggle = document.getElementById('btnThemeToggle');
    const txtThemeToggle = document.getElementById('txtThemeToggle');
    const txtExportCSV = document.getElementById('txtExportCSV');
    const btnReset = document.getElementById('btnReset');
    const btnSidebarIngestion = document.getElementById('btnSidebarIngestion');
    const btnSidebarPopulation = document.getElementById('btnSidebarPopulation');
    const btnSidebarAnalyze = document.getElementById('btnSidebarAnalyze');
    const sidebarScrollArea = document.getElementById('sidebarScrollArea');
    const dataIngestionPanel = document.getElementById('dataIngestionPanel');
    const flowStatusText = document.getElementById('flowStatusText');
    const flowActionHint = document.getElementById('flowActionHint');
    const flowStepIngest = document.getElementById('flowStepIngest');
    const flowStepConfig = document.getElementById('flowStepConfig');
    const flowStepAnalyze = document.getElementById('flowStepAnalyze');
    const qualityWarning = document.getElementById('qualityWarning');
    const ingestDiagnosticsPanel = document.getElementById('ingestDiagnostics');
    const ingestDiagnosticsList = document.getElementById('ingestDiagnosticsList');
    const boxPlotLoadingMsg = document.getElementById('boxPlotLoadingMsg');
    const boxPlotProgressBarFill = document.getElementById('boxPlotProgressBarFill');
    const boxPlotProgressText = document.getElementById('boxPlotProgressText');
    const boxPlotsGridEl = document.getElementById('boxPlotsGrid');
    const viewParamsPanel = document.getElementById('view-params');
    const runStatBar = document.getElementById('runStatBar');
    const runStatSubjects = document.getElementById('runStatSubjects');
    const runStatTrials = document.getElementById('runStatTrials');
    const runStatObs = document.getElementById('runStatObs');
    const runStatCmax = document.getElementById('runStatCmax');
    const beRefSummary = document.getElementById('beRefSummary');
    const ingestIssueBadge = document.getElementById('ingestIssueBadge');
    const profileAxesBody = document.getElementById('profileAxesBody');
    const btnProfileAxesToggle = document.getElementById('btnProfileAxesToggle');
    const statusToast = document.getElementById('statusToast');
    let statusToastTimer = null;

    function showStatusToast(message, tone = 'info') {
        if (!statusToast || !message) return;
        statusToast.textContent = message;
        statusToast.classList.remove('error', 'warn', 'show');
        if (tone === 'error') statusToast.classList.add('error');
        if (tone === 'warn') statusToast.classList.add('warn');
        statusToast.classList.add('show');
        if (statusToastTimer) window.clearTimeout(statusToastTimer);
        statusToastTimer = window.setTimeout(() => {
            statusToast.classList.remove('show', 'error', 'warn');
        }, tone === 'error' ? 5200 : 3600);
    }

    function syncBoxResultsOnlyButton() {
        if (!btnBoxResultsOnly) return;
        btnBoxResultsOnly.textContent = boxResultsOnly ? 'Showing Results Only' : 'Results Only';
        if (boxResultsOnly) {
            btnBoxResultsOnly.style.background = 'var(--accent-soft)';
            btnBoxResultsOnly.style.color = 'var(--accent)';
            btnBoxResultsOnly.style.borderColor = 'rgba(212,149,106,0.3)';
        } else {
            btnBoxResultsOnly.style.background = 'var(--bg-base)';
            btnBoxResultsOnly.style.color = 'var(--text-secondary)';
            btnBoxResultsOnly.style.borderColor = 'var(--border-default)';
        }
    }

    function addListener(el, eventName, handler) {
        if (el) el.addEventListener(eventName, handler);
    }

    function getRenderablePlotElement(plotId) {
        const el = document.getElementById(plotId);
        return el && el.data ? el : null;
    }

    const VIEW_EXPORT_CONFIG = {
        'view-profile': {
            csvLabel: 'Export Profile Data',
            canExportPNG: () => !!globalTrialsData || globalObsData.length > 0,
            canExportCSV: () => !!globalTrialsData
        },
        'view-params': {
            csvLabel: 'Export Param Data',
            canExportPNG: () => !!globalTrialsData,
            canExportCSV: () => !!globalTrialsData
        },
        'view-be': {
            csvLabel: 'Export BE Data',
            canExportPNG: () => !!globalTrialsData,
            canExportCSV: () => !!globalTrialsData
        },
        'view-stats': {
            csvLabel: 'Export Stats Data',
            canExportPNG: () => false,
            canExportCSV: () => !!globalTrialsData
        }
    };

    function applyExportControlsForView(viewId) {
        const cfg = VIEW_EXPORT_CONFIG[viewId] || VIEW_EXPORT_CONFIG['view-profile'];
        if (btnExportPNG) btnExportPNG.disabled = !cfg.canExportPNG();
        if (txtExportCSV) txtExportCSV.innerText = cfg.csvLabel;
        if (btnExportCSV) btnExportCSV.disabled = !cfg.canExportCSV();
    }

    // Tab Logic
    const tabBtns = document.querySelectorAll('.tab-btn');
    const viewContainers = document.querySelectorAll('.view-container');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            // Update Tab styles
            tabBtns.forEach(b => {
                b.classList.remove('tab-active');
                b.classList.add('tab-inactive');
            });
            btn.classList.add('tab-active');
            btn.classList.remove('tab-inactive');

            // Show corresponding view
            const targetId = btn.getAttribute('data-target');
            currentTab = targetId;
            if (targetId !== 'view-profile' && (globalTrialsData || globalObsData.length > 0) && !beNeedsRerun) {
                hasReviewedResults = true;
            }
            viewContainers.forEach(vc => {
                if(vc.id === targetId) vc.classList.remove('view-hidden');
                else vc.classList.add('view-hidden');
            });

            // Adjust exports based on view
            applyExportControlsForView(targetId);

            // Force redraw to fix Plotly resizing issues when div is unhidden
            updateAllViews();
            updateFlowSetupState();
            saveSessionState();
        });
    });

    // -- Event Listeners --
    addListener(fileInputRef, 'change', (e) => handleSimUpload(e, 'reference'));
    addListener(fileInputTest, 'change', (e) => handleSimUpload(e, 'test'));
    addListener(obsFileInput, 'change', handleObsUpload);
    addListener(beRefTrialSelect, 'change', () => {
        if (beRefTrialSelect.value !== '') {
            setReferenceTrial(beRefTrialSelect.value);
            return;
        } else {
            beReferenceIndex = '';
            if (beRefSummary) {
                beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
            }
        }
        renderTrialList();
        updateBERefSelect();
        updateBEView();
        saveSessionState();
    });
    addListener(beTestTrialSelect, 'change', () => {
        beTestIndex = beTestTrialSelect.value;
        updateBEView();
        saveSessionState();
    });

    function runBEAnalysis() {
        if (btnRunBE && btnRunBE.disabled) {
            showStatusToast('Run BE is unavailable. Ensure at least one valid Reference and Test trial is available.', 'warn');
            return;
        }

        // Sync selector options/state in case formulation activity changed before rerun.
        updateBERefSelect();

        if (beRefTrialSelect) {
            beReferenceIndex = beRefTrialSelect.value;
        }
        if (beTestTrialSelect) {
            beTestIndex = beTestTrialSelect.value;
        }

        globalBEData = []; // Reset global results for Export CSV/Bundle when "Run BE" is manually clicked
        hasReviewedResults = true;
        beNeedsRerun = false;
        updateFlowSetupState();

        try {
            updateBEView(true);
            
            if (globalBEData && globalBEData.length > 0) {
                showStatusToast(`BE analysis complete: ${globalBEData.length} result row${globalBEData.length === 1 ? '' : 's'}.`);
            } else {
                showStatusToast('No BE results generated. Check active trials and parameter availability.', 'warn');
            }
            saveSessionState();
        } catch (err) {
            console.error('Run BE failed:', err);
            showStatusToast('Run BE failed. Check browser console for details.', 'error');
        }
    }

    function handleBEMethodChange() {
        // Method switches can invalidate previous test filters; force a safe re-sync.
        if (beTestTrialSelect) {
            beTestIndex = beTestTrialSelect.value;
        }
        if (beRefTrialSelect) {
            beReferenceIndex = beRefTrialSelect.value;
        }
        updateBERefSelect();

        // Mark BE outputs stale and require explicit rerun.
        hasReviewedResults = false;
        beNeedsRerun = true;

        const beWarningMsg = document.getElementById('beWarningMsg');
        if (beWarningMsg) {
            beWarningMsg.textContent = 'BE method changed. Click "Run BE" to regenerate charts and table.';
            beWarningMsg.classList.remove('hidden');
        }

        updateBEView();
        updateFlowSetupState();
        saveSessionState();

        // Notify user to rerun since method has changed
        showStatusToast('BE Method changed. Click "Run BE" to update results and charts.', 'warn');
    }

    if (beMethodSelect) {
        beMethodSelect.addEventListener('change', handleBEMethodChange);
    }
    addListener(btnRunBE, 'click', runBEAnalysis);

    if (targetConcUnitSelect) {
        targetConcUnitSelect.addEventListener('change', () => {
            const v = targetConcUnitSelect.value;
            // Set targetConcUnit; null means auto-detect on next upload
            targetConcUnit = v ? normConcUnit(v) : null;
            if (targetConcUnit && inputYTitle && inputYTitle.value.startsWith('Concentration')) {
                inputYTitle.value = `Concentration (${targetConcUnit})`;
            }
            if (globalTrialsData) {
                showStatusToast('Unit changed. Re-upload simulation files to apply conversion.', 'warn');
            }
            saveSessionState();
        });
    }

    if (btnSidebarIngestion) {
        btnSidebarIngestion.addEventListener('click', () => {
            if (!globalTrialsData && globalObsData.length === 0 && fileInputRef) {
                fileInputRef.click();
            }
            focusSidebarSection(dataIngestionPanel, btnSidebarIngestion);
        });
    }

    if (btnSidebarPopulation) {
        btnSidebarPopulation.addEventListener('click', () => {
            if (trialSelectionPanel.classList.contains('hidden')) {
                focusSidebarSection(dataIngestionPanel, btnSidebarIngestion);
                return;
            }
            focusSidebarSection(trialSelectionPanel, btnSidebarPopulation);
        });
    }

    function activateTab(targetId) {
        const tabBtn = document.querySelector(`.tab-btn[data-target="${targetId}"]`);
        if (tabBtn) tabBtn.click();
    }

    if (btnSidebarAnalyze) {
        btnSidebarAnalyze.addEventListener('click', () => {
            const hasAnyData = !!globalTrialsData || globalObsData.length > 0;
            if (!hasAnyData) {
                focusSidebarSection(dataIngestionPanel, btnSidebarIngestion);
                if (fileInputRef) fileInputRef.click();
                return;
            }
            hasReviewedResults = true;
            setRailActive(btnSidebarAnalyze);
            activateTab('view-profile');
            const tabNav = document.getElementById('tabNav');
            if (tabNav) tabNav.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            updateFlowSetupState();
            saveSessionState();
        });
    }

    initTheme();
    if (btnThemeToggle) {
        btnThemeToggle.addEventListener('click', () => {
            const current = document.documentElement.getAttribute('data-theme') || 'dark';
            applyTheme(current === 'dark' ? 'light' : 'dark');
        });
    }

    loadSessionState();
    syncBoxResultsOnlyButton();
    updateFlowSetupState();
    applyExportControlsForView(currentTab);

    function computeActiveCmaxMean() {
        if (!globalTrialsData) return null;
        const values = [];
        globalTrialsData.trials.forEach(trial => {
            if (!trial.active || !trial.rawParams || trial.rawParams.length === 0) return;
            const arr = extractParamData(trial, 'Cmax');
            arr.forEach(v => {
                if (Number.isFinite(v)) values.push(v);
            });
        });
        if (values.length === 0) return null;
        return values.reduce((a, b) => a + b, 0) / values.length;
    }

    function updateRunStatBar() {
        const hasData = !!globalTrialsData || globalObsData.length > 0;
        runStatBar.classList.toggle('hidden', !hasData);
        if (!hasData) return;

        const activeTrials = globalTrialsData ? globalTrialsData.trials.filter(t => t.active).length : 0;
        const subjects = globalTrialsData ? globalTrialsData.totalSubjects : 0;
        const cmaxMean = computeActiveCmaxMean();

        runStatSubjects.textContent = subjects;
        runStatTrials.textContent = activeTrials;
        runStatObs.textContent = globalObsData.length;
        runStatCmax.textContent = Number.isFinite(cmaxMean) ? cmaxMean.toFixed(3) : '-';
    }

    function setRailActive(activeBtn) {
        [btnSidebarIngestion, btnSidebarPopulation, btnSidebarAnalyze].forEach(btn => {
            if (!btn) return;
            btn.classList.toggle('active', btn === activeBtn);
        });
    }

    function renderIngestDiagnostics() {
        if (!ingestDiagnosticsPanel || !ingestDiagnosticsList) return;
        if (ingestIssueBadge) {
            const errorCount = ingestDiagnostics.filter(d => d.status === 'error').length;
            const warnCount = ingestDiagnostics.filter(d => d.status === 'warn').length;
            if (errorCount > 0) {
                ingestIssueBadge.classList.remove('hidden');
                ingestIssueBadge.style.background = 'rgba(248,113,113,0.08)';
                ingestIssueBadge.style.borderColor = 'rgba(248,113,113,0.25)';
                ingestIssueBadge.style.color = '#f87171';
                ingestIssueBadge.textContent = `Ingestion issues: ${errorCount} error${errorCount > 1 ? 's' : ''}${warnCount ? `, ${warnCount} warning${warnCount > 1 ? 's' : ''}` : ''}`;
            } else if (warnCount > 0) {
                ingestIssueBadge.classList.remove('hidden');
                ingestIssueBadge.style.background = 'rgba(251,191,36,0.08)';
                ingestIssueBadge.style.borderColor = 'rgba(251,191,36,0.25)';
                ingestIssueBadge.style.color = '#fbbf24';
                ingestIssueBadge.textContent = `Ingestion warnings: ${warnCount}`;
            } else {
                ingestIssueBadge.classList.add('hidden');
                ingestIssueBadge.textContent = '';
            }
        }
        if (!ingestDiagnostics.length) {
            ingestDiagnosticsPanel.classList.add('hidden');
            ingestDiagnosticsList.innerHTML = '';
            return;
        }
        ingestDiagnosticsPanel.classList.remove('hidden');
        ingestDiagnosticsList.innerHTML = ingestDiagnostics.slice(-10).map(d => {
            const cls = d.status === 'error' ? 'diag-err' : (d.status === 'warn' ? 'diag-warn' : 'diag-ok');
            return `<div class="diag-list-item text-[10px]">
                <div class="font-semibold ${cls}">${d.scope.toUpperCase()} - ${d.fileName}</div>
                <div style="color: var(--text-secondary);">${d.message}</div>
            </div>`;
        }).join('');
    }

    function cssVar(name) {
        return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    }

    function getPlotTheme() {
        return {
            bg: cssVar('--plot-bg'),
            paper: cssVar('--plot-paper'),
            grid: cssVar('--plot-grid'),
            zeroline: cssVar('--plot-zeroline'),
            axisline: cssVar('--plot-axisline'),
            tick: cssVar('--plot-tick'),
            font: cssVar('--plot-font'),
            title: cssVar('--plot-title'),
            subtitle: cssVar('--plot-subtitle'),
            legendBg: cssVar('--plot-legend-bg'),
            hoverBg: cssVar('--plot-hover-bg'),
            hoverBorder: cssVar('--plot-hover-border')
        };
    }

    function applyTheme(theme) {
        const normalized = theme === 'light' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', normalized);
        const themeMeta = document.querySelector('meta[name="theme-color"]');
        if (themeMeta) {
            themeMeta.setAttribute('content', normalized === 'dark' ? '#18181b' : '#f3f4f6');
        }
        if (txtThemeToggle) txtThemeToggle.textContent = normalized === 'dark' ? 'Light' : 'Dark';
        try { localStorage.setItem(THEME_KEY, normalized); } catch (e) {}
        if (globalTrialsData || globalObsData.length > 0) {
            updateAllViews();
        }
    }

    function initTheme() {
        let theme = 'dark';
        try {
            const stored = localStorage.getItem(THEME_KEY);
            if (stored === 'dark' || stored === 'light') {
                theme = stored;
            }
        } catch (e) {}
        applyTheme(theme);
    }

    function pushDiagnostic(scope, fileName, status, message) {
        ingestDiagnostics.push({ scope, fileName, status, message });
        renderIngestDiagnostics();
        if (status === 'error') {
            showStatusToast(`${scope.toUpperCase()} upload failed: ${fileName}`, 'error');
        } else if (status === 'warn') {
            showStatusToast(`${scope.toUpperCase()} upload warning: ${fileName}`, 'warn');
        }
    }

    function evaluateQualityWarnings() {
        if (!qualityWarning) return;
        const warnings = [];
        const ingestErrors = ingestDiagnostics.filter(d => d.status === 'error').length;
        const ingestWarns = ingestDiagnostics.filter(d => d.status === 'warn').length;
        if (ingestErrors > 0) warnings.push(`${ingestErrors} file(s) failed ingestion parsing.`);
        else if (ingestWarns > 0) warnings.push(`${ingestWarns} file(s) ingested with warnings.`);
        if (globalTrialsData && globalTrialsData.trials.length > 0) {
            const activeTrials = globalTrialsData.trials.filter(t => t.active);
            if (activeTrials.length < 1) warnings.push('No active trials selected.');

            let nonMonotonic = 0;
            let nonPositiveConc = 0;
            activeTrials.forEach(t => {
                for (let i = 1; i < t.times.length; i++) {
                    if (Number(t.times[i]) < Number(t.times[i - 1])) { nonMonotonic += 1; break; }
                }
                t.concsAtTime.forEach(arr => {
                    arr.forEach(v => { if (Number.isFinite(v) && v <= 0) nonPositiveConc += 1; });
                });
            });

            if (nonMonotonic > 0) warnings.push(`${nonMonotonic} trial(s) contain non-monotonic time rows.`);
            if (nonPositiveConc > 0) warnings.push(`${nonPositiveConc} non-positive concentration values detected in active trials.`);
        }
        if (globalObsData.length > 0) {
            let nonPositiveObs = 0;
            globalObsData.forEach(o => o.y.forEach(v => { if (Number.isFinite(v) && v <= 0) nonPositiveObs += 1; }));
            if (nonPositiveObs > 0) warnings.push(`${nonPositiveObs} non-positive observed concentrations detected.`);
        }

        if (!warnings.length) {
            qualityWarning.classList.add('hidden');
            qualityWarning.textContent = '';
            return;
        }
        qualityWarning.classList.remove('hidden');
        qualityWarning.textContent = warnings.join(' ');
    }

    function saveSessionState() {
        try {
            const payload = {
                currentTab,
                isLogScale,
                hasReviewedResults,
                beNeedsRerun,
                hasAutoOpenedAxes,
                beReferenceIndex,
                beTestIndex,
                beMethod: beMethodSelect ? beMethodSelect.value : 'simple',
                targetConcUnit: targetConcUnit || null,
                toggles: {
                    showContour100: !!(showContour100 && showContour100.checked),
                    showContour95: !!(showContour95 && showContour95.checked),
                    showContour90: !!(showContour90 && showContour90.checked),
                    showContours: !!(showContoursCheck && showContoursCheck.checked),
                    showMean: !!(showMean && showMean.checked),
                    showMedian: !!(showMedian && showMedian.checked),
                    showCIMean: !!(showCIMean && showCIMean.checked),
                    showIndividuals: !!(showIndividualsCheck && showIndividualsCheck.checked),
                    boxShowOutliers: !!(boxShowOutliers && boxShowOutliers.checked),
                    boxShowMean: !!(boxShowMean && boxShowMean.checked),
                    boxResultsOnly
                },
                axes: {
                    xTitle: inputXTitle ? inputXTitle.value : 'Time (h)',
                    yTitle: inputYTitle ? inputYTitle.value : 'Concentration (ug/mL)',
                    xMax: inputXMax ? inputXMax.value : '',
                    yMax: inputYMax ? inputYMax.value : ''
                },
                mapping: {
                    simTimeContains: mapSimTimeContains ? mapSimTimeContains.value : columnMappingDefaults.simTimeContains,
                    simSubjectPrefix: mapSimSubjectPrefix ? mapSimSubjectPrefix.value : columnMappingDefaults.simSubjectPrefix,
                    obsTimeContains: mapObsTimeContains ? mapObsTimeContains.value : columnMappingDefaults.obsTimeContains,
                    obsConcContains: mapObsConcContains ? mapObsConcContains.value : columnMappingDefaults.obsConcContains
                }
            };
            localStorage.setItem(SESSION_KEY, JSON.stringify(payload));
        } catch (e) {}
    }

    function loadSessionState() {
        try {
            const raw = localStorage.getItem(SESSION_KEY);
            if (!raw) return;
            const s = JSON.parse(raw);
            if (s.axes) {
                if (inputXTitle) inputXTitle.value = s.axes.xTitle || inputXTitle.value;
                if (inputYTitle) inputYTitle.value = s.axes.yTitle || inputYTitle.value;
                if (inputXMax) inputXMax.value = s.axes.xMax || '';
                if (inputYMax) inputYMax.value = s.axes.yMax || '';
            }
            if (s.mapping) {
                if (mapSimTimeContains) mapSimTimeContains.value = s.mapping.simTimeContains || columnMappingDefaults.simTimeContains;
                if (mapSimSubjectPrefix) mapSimSubjectPrefix.value = s.mapping.simSubjectPrefix || columnMappingDefaults.simSubjectPrefix;
                if (mapObsTimeContains) mapObsTimeContains.value = s.mapping.obsTimeContains || columnMappingDefaults.obsTimeContains;
                if (mapObsConcContains) mapObsConcContains.value = s.mapping.obsConcContains || columnMappingDefaults.obsConcContains;
            }
            if (s.toggles) {
                if (showContour100) showContour100.checked = !!s.toggles.showContour100;
                if (showContour95) showContour95.checked = !!s.toggles.showContour95;
                if (showContour90) showContour90.checked = !!s.toggles.showContour90;
                if (showContoursCheck) showContoursCheck.checked = !!s.toggles.showContours;
                if (showMean) showMean.checked = !!s.toggles.showMean;
                if (showMedian) showMedian.checked = !!s.toggles.showMedian;
                if (showCIMean) showCIMean.checked = !!s.toggles.showCIMean;
                if (showIndividualsCheck) showIndividualsCheck.checked = !!s.toggles.showIndividuals;
                if (boxShowOutliers) boxShowOutliers.checked = !!s.toggles.boxShowOutliers;
                if (boxShowMean) boxShowMean.checked = !!s.toggles.boxShowMean;
                boxResultsOnly = !!s.toggles.boxResultsOnly;
            }
            isLogScale = !!s.isLogScale;
            setScaleButtonStyles(isLogScale ? btnScaleLog : btnScaleLinear, isLogScale ? btnScaleLinear : btnScaleLog);
            if (beMethodSelect && s.beMethod) {
                const allowedMethods = ['simple', 'paired-trial-number'];
                beMethodSelect.value = allowedMethods.includes(s.beMethod) ? s.beMethod : 'simple';
            }
            if (typeof s.hasReviewedResults === 'boolean') hasReviewedResults = s.hasReviewedResults;
            if (typeof s.beNeedsRerun === 'boolean') beNeedsRerun = s.beNeedsRerun;
            if (typeof s.hasAutoOpenedAxes === 'boolean') hasAutoOpenedAxes = s.hasAutoOpenedAxes;
            if (typeof s.beReferenceIndex === 'string') beReferenceIndex = s.beReferenceIndex;
            if (typeof s.beTestIndex === 'string') beTestIndex = s.beTestIndex;
            if (s.targetConcUnit) {
                targetConcUnit = normConcUnit(s.targetConcUnit);
                if (targetConcUnitSelect) {
                    // Try to match select option; fallback to '' (auto)
                    const opt = Array.from(targetConcUnitSelect.options).find(o => o.value && normConcUnit(o.value) === targetConcUnit);
                    targetConcUnitSelect.value = opt ? opt.value : '';
                }
            }
            if (s.currentTab) {
                const tabBtn = document.querySelector(`.tab-btn[data-target="${s.currentTab}"]`);
                if (tabBtn) tabBtn.click();
            }
        } catch (e) {}
    }

    function updateFlowSetupState() {
        if (!btnSidebarPopulation) return;
        const hasPopulation = !!globalTrialsData && Array.isArray(globalTrialsData.trials) && globalTrialsData.trials.length > 0;
        const hasAnyData = !!globalTrialsData || globalObsData.length > 0;
        btnSidebarPopulation.classList.toggle('flow-disabled', !hasPopulation);
        btnSidebarPopulation.disabled = !hasPopulation;
        btnSidebarPopulation.title = hasPopulation
            ? 'Population Configuration'
            : 'Population Configuration (available after simulation upload)';
        btnSidebarPopulation.setAttribute('aria-disabled', String(!hasPopulation));

        if (btnSidebarAnalyze) {
            btnSidebarAnalyze.classList.toggle('flow-disabled', !hasAnyData);
            btnSidebarAnalyze.disabled = !hasAnyData;
            btnSidebarAnalyze.title = hasAnyData
                ? 'Analyze & Review'
                : 'Analyze & Review (available after data upload)';
            btnSidebarAnalyze.setAttribute('aria-disabled', String(!hasAnyData));
        }

        if (btnExportBundle) btnExportBundle.disabled = !hasAnyData;

        if (flowStepIngest) {
            flowStepIngest.classList.toggle('done', hasAnyData);
            flowStepIngest.classList.toggle('active', !hasAnyData);
        }
        if (flowStepConfig) {
            flowStepConfig.classList.toggle('done', hasPopulation);
            flowStepConfig.classList.toggle('active', hasAnyData && !hasPopulation);
        }
        if (flowStepAnalyze) {
            flowStepAnalyze.classList.toggle('done', hasReviewedResults && hasAnyData);
            flowStepAnalyze.classList.toggle('active', hasAnyData && hasPopulation && !hasReviewedResults);
        }

        if (flowStatusText && flowActionHint) {
            if (!hasAnyData) {
                flowStatusText.textContent = 'Step 1: Upload Reference and Test simulations, plus observed data.';
                flowActionHint.textContent = 'Use separate upload cards for Reference Sim and Test Sim. Optional: map custom columns in Advanced Settings.';
            } else if (!hasPopulation) {
                flowStatusText.textContent = 'Step 1 complete: data loaded.';
                flowActionHint.textContent = 'Set each trial formulation (Test/Reference) and choose BE comparison trials.';
            } else if (!hasReviewedResults) {
                flowStatusText.textContent = 'Step 2 complete: population is configured.';
                flowActionHint.textContent = 'Assign Test/Reference formulation in Population and open analysis tabs to complete Step 3.';
            } else {
                flowStatusText.textContent = 'Workflow complete: analysis reviewed.';
                flowActionHint.textContent = 'Use Export Bundle for all tables and available plot images.';
            }
        }

        evaluateQualityWarnings();
    }

    function focusSidebarSection(sectionEl, activeBtn) {
        if (!sectionEl) return;
        setRailActive(activeBtn);
        sectionEl.classList.add('section-focus-glow');
        window.setTimeout(() => sectionEl.classList.remove('section-focus-glow'), 900);

        if (!sidebarScrollArea) {
            sectionEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return;
        }

        const containerRect = sidebarScrollArea.getBoundingClientRect();
        const targetRect = sectionEl.getBoundingClientRect();
        const offsetTop = targetRect.top - containerRect.top + sidebarScrollArea.scrollTop - 10;
        sidebarScrollArea.scrollTo({ top: Math.max(0, offsetTop), behavior: 'smooth' });
    }

    function setReferenceTrial(indexStr) {
        if (!globalTrialsData) return;
        const idx = String(indexStr);
        if (!globalTrialsData.trials[idx]) return;
        beReferenceIndex = idx;
        beRefTrialSelect.value = idx;
        if (beRefSummary) {
            beRefSummary.textContent = `Reference: ${getTrialLabel(globalTrialsData.trials[idx])}`;
        }
        renderTrialList();
        updateBEView();
        saveSessionState();
    }

    function computeTrialProfileStats(trial) {
        if (!trial || !Array.isArray(trial.concsAtTime)) return;
        trial.stats = { means: [], medians: [], p00: [], p100: [], p025: [], p975: [], p05: [], p95: [], p25: [], p75: [], lowerCI: [], upperCI: [] };

        trial.concsAtTime.forEach(concs => {
            if (concs.length > 0) {
                const mean = concs.reduce((a, b) => a + b, 0) / concs.length;
                trial.stats.means.push(mean);
                trial.stats.medians.push(calculatePercentile(concs, 0.50));

                trial.stats.p00.push(Math.min(...concs));
                trial.stats.p100.push(Math.max(...concs));
                trial.stats.p025.push(calculatePercentile(concs, 0.025));
                trial.stats.p975.push(calculatePercentile(concs, 0.975));
                trial.stats.p05.push(calculatePercentile(concs, 0.05));
                trial.stats.p95.push(calculatePercentile(concs, 0.95));
                trial.stats.p25.push(calculatePercentile(concs, 0.25));
                trial.stats.p75.push(calculatePercentile(concs, 0.75));

                const ci = calculateMeanAndCI(concs);
                trial.stats.lowerCI.push(ci.lowerCI);
                trial.stats.upperCI.push(ci.upperCI);
            } else {
                Object.keys(trial.stats).forEach(k => trial.stats[k].push(null));
            }
        });
    }
    
    // Graceful reset without reloading the page
    addListener(btnReset, 'click', () => {
        const hasLoadedData = !!globalTrialsData || globalObsData.length > 0;
        if (hasLoadedData) {
            const confirmed = window.confirm('Reset will clear all uploaded data, selected trials, and current analysis state. Continue?');
            if (!confirmed) return;
        }

        // 1. Reset Global State
        globalTrialsData = null;
        globalObsData = [];
        globalBEData = [];
        beReferenceIndex = '';
        beTestIndex = '';
        isLogScale = false;
        hasReviewedResults = false;
        beNeedsRerun = false;
        hasAutoOpenedAxes = false;
        targetConcUnit = null;
        ingestDiagnostics = [];
        renderIngestDiagnostics();
        try { localStorage.removeItem(SESSION_KEY); } catch (e) {}

        // 2. Clear File Inputs
        fileInputRef.value = '';
        fileInputTest.value = '';
        obsFileInput.value = '';
        simRefFileCount = 0;
        simTestFileCount = 0;
        obsFileCount = 0;
        updateFileBadges();
        updateObsFooter();

        // 3. Reset Sidebar & Top UI
        trialSelectionPanel.classList.add('hidden');
        updateFlowSetupState();
        trialList.innerHTML = '';
        statSubjects.innerText = '0';
        statTrials.innerText = '0';
        statActiveTrials.innerText = '0';
        beRefTrialSelect.innerHTML = '<option value="">-- Select Reference Trial --</option>';
        if (beTestTrialSelect) beTestTrialSelect.innerHTML = '<option value="">All Test Trials</option>';
        if (beRefSummary) beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
        btnExportPNG.disabled = true;
        btnExportCSV.disabled = true;
        if (beMethodSelect) beMethodSelect.value = 'simple';

        // 3b. Reset view controls to defaults
        if (showContour100) showContour100.checked = false;
        if (showContour95) showContour95.checked = false;
        if (showContour90) showContour90.checked = true;
        if (showContoursCheck) showContoursCheck.checked = false;
        if (showMean) showMean.checked = true;
        if (showMedian) showMedian.checked = false;
        if (showCIMean) showCIMean.checked = true;
        if (showIndividualsCheck) showIndividualsCheck.checked = false;
        if (boxShowOutliers) boxShowOutliers.checked = true;
        if (boxShowMean) boxShowMean.checked = true;
        boxResultsOnly = false;
        if (inputXTitle) inputXTitle.value = 'Time (h)';
        if (inputYTitle) inputYTitle.value = 'Concentration (ug/mL)';
        if (targetConcUnitSelect) targetConcUnitSelect.value = 'ug/ml';
        if (inputXMax) inputXMax.value = '';
        if (inputYMax) inputYMax.value = '';
        setScaleButtonStyles(btnScaleLinear, btnScaleLog);

        // 4. Purge Plotly Memory
        try {
            Plotly.purge('plotlyChart');
            ['bePlotCmax', 'bePlotAUCinf', 'bePlotAUCt', 'bePlotConclusion'].forEach(id => Plotly.purge(id));
            Array.from(document.querySelectorAll('[id^="plotlyBox_"]')).forEach(el => Plotly.purge(el.id));
        } catch(e) {}

        // 5. Clear Grids & Tables
        document.getElementById('statsTableBody').innerHTML = '';
        document.getElementById('statsCompareBody').innerHTML = '';
        document.getElementById('statsComparePanel').classList.add('hidden');
        document.getElementById('statsTable').parentElement.classList.add('hidden');
        document.getElementById('boxPlotsGrid').style.display = 'none';
        document.getElementById('beContent').classList.add('hidden');
        setBoxPlotLoadingState(false);

        // 6. Restore Empty States
        emptyState.classList.remove('hidden');
        document.getElementById('statsEmptyMsg').classList.remove('hidden');
        document.getElementById('boxEmptyMsg').classList.remove('hidden');
        document.getElementById('beEmptyMsg').classList.remove('hidden');
        const beNoParamsText = document.querySelector('#beNoParamsMsg p');
        if (beNoParamsText) beNoParamsText.textContent = 'No valid individual PK parameters found in the selected trials to compute BE.';

        // 7. Restore default tab
        currentTab = 'view-profile';
        setRailActive(btnSidebarIngestion);
        tabBtns.forEach(b => {
            b.classList.remove('tab-active');
            b.classList.add('tab-inactive');
        });
        const profileTabBtn = Array.from(tabBtns).find(b => b.getAttribute('data-target') === 'view-profile');
        if (profileTabBtn) {
            profileTabBtn.classList.add('tab-active');
            profileTabBtn.classList.remove('tab-inactive');
        }
        viewContainers.forEach(vc => {
            if (vc.id === 'view-profile') vc.classList.remove('view-hidden');
            else vc.classList.add('view-hidden');
        });
        txtExportCSV.innerText = 'Export Profile Data';
        updateRunStatBar();
    });
    
    const triggerVisualUpdate = () => { 
        if(currentTab === 'view-profile' && (globalTrialsData || globalObsData.length > 0)) updatePlot(); 
        if(currentTab === 'view-params' && globalTrialsData) updateBoxPlots();
        if(currentTab === 'view-be' && globalTrialsData) updateBEView();
        saveSessionState();
    };

    addListener(showContour100, 'change', triggerVisualUpdate);
    addListener(showContour95, 'change', triggerVisualUpdate);
    addListener(showContour90, 'change', triggerVisualUpdate);
    addListener(showContoursCheck, 'change', triggerVisualUpdate);
    addListener(showMean, 'change', triggerVisualUpdate);
    addListener(showMedian, 'change', triggerVisualUpdate);
    addListener(showCIMean, 'change', triggerVisualUpdate);
    addListener(showIndividualsCheck, 'change', triggerVisualUpdate);
    addListener(inputXTitle, 'input', triggerVisualUpdate);
    addListener(inputYTitle, 'input', triggerVisualUpdate);
    addListener(inputXMax, 'input', triggerVisualUpdate);
    addListener(inputYMax, 'input', triggerVisualUpdate);
    addListener(mapSimTimeContains, 'input', saveSessionState);
    addListener(mapSimSubjectPrefix, 'input', saveSessionState);
    addListener(mapObsTimeContains, 'input', saveSessionState);
    addListener(mapObsConcContains, 'input', saveSessionState);
    addListener(boxShowOutliers, 'change', triggerVisualUpdate);
    addListener(boxShowMean, 'change', triggerVisualUpdate);
    if (btnBoxResultsOnly) {
        btnBoxResultsOnly.addEventListener('click', () => {
            boxResultsOnly = !boxResultsOnly;
            syncBoxResultsOnlyButton();
            triggerVisualUpdate();
        });
    }
    function setScaleButtonStyles(activeBtn, inactiveBtn) {
        if (!activeBtn || !inactiveBtn) return;
        activeBtn.style.cssText = 'background: var(--bg-elevated); color: var(--text-primary); border: 1px solid var(--border-strong);';
        inactiveBtn.style.cssText = 'color: var(--text-secondary); background: transparent; border: 1px solid transparent;';
    }

    const bePlotTargets = [
        { id: 'bePlotCmax', name: 'Bioequivalence_Cmax' },
        { id: 'bePlotAUCinf', name: 'Bioequivalence_AUCinf' },
        { id: 'bePlotAUCt', name: 'Bioequivalence_AUCt' },
        { id: 'bePlotConclusion', name: 'Bioequivalence_Conclusion' }
    ];

    function getRenderablePlotTargets(targets) {
        return targets.filter(plot => {
            const el = document.getElementById(plot.id);
            return !!(el && el.data && el.data.length);
        });
    }

    addListener(btnScaleLinear, 'click', () => {
        isLogScale = false;
        setScaleButtonStyles(btnScaleLinear, btnScaleLog);
        triggerVisualUpdate();
    });
    
    addListener(btnScaleLog, 'click', () => {
        isLogScale = true;
        setScaleButtonStyles(btnScaleLog, btnScaleLinear);
        triggerVisualUpdate();
    });

    addListener(btnSelectAll, 'click', () => {
        if (!globalTrialsData) return;
        globalTrialsData.trials.forEach(t => t.active = true);
        renderTrialList();
        updateAllViews();
        updateFlowSetupState();
        saveSessionState();
    });

    addListener(btnSelectNone, 'click', () => {
        if (!globalTrialsData) return;
        globalTrialsData.trials.forEach(t => t.active = false);
        renderTrialList();
        updateAllViews();
        updateFlowSetupState();
        saveSessionState();
    });

    addListener(btnExportPNG, 'click', () => {
        const txtExportPNG = document.getElementById('txtExportPNG');
        if (!txtExportPNG) return;
        const originalText = txtExportPNG.innerText;

        if (currentTab === 'view-profile') {
            const profilePlot = getRenderablePlotElement('plotlyChart');
            if (profilePlot) Plotly.downloadImage('plotlyChart', {format: 'png', width: 1200, height: 700, filename: `PopPK_Profile`});
        } else if (currentTab === 'view-be') {
            const activeBEPlots = getRenderablePlotTargets(bePlotTargets);
            if (activeBEPlots.length === 0) return;

            btnExportPNG.disabled = true;
            activeBEPlots.forEach((plot, idx) => {
                setTimeout(() => {
                    txtExportPNG.innerText = `Downloading ${idx + 1}/${activeBEPlots.length}...`;
                    Plotly.downloadImage(plot.id, {
                        format: 'png',
                        width: 1100,
                        height: 700,
                        filename: `PopPK_${plot.name}`
                    });

                    if (idx === activeBEPlots.length - 1) {
                        setTimeout(() => {
                            txtExportPNG.innerText = originalText;
                            btnExportPNG.disabled = false;
                        }, 800);
                    }
                }, idx * 900);
            });
        } else if (currentTab === 'view-params') {
            const params = Array.from(document.querySelectorAll('[id^="plotlyBox_"]'))
                .filter(el => el && el.data)
                .map(el => ({
                    id: el.id,
                    name: (el.dataset.paramLabel || el.id.replace(/^plotlyBox_/, '')).replace(/\s+/g, '_')
                }));

            if (params.length === 0) return;

            btnExportPNG.disabled = true;
            params.forEach((p, idx) => {
                setTimeout(() => {
                    txtExportPNG.innerText = `Downloading ${idx + 1}/${params.length}...`;
                    Plotly.downloadImage(p.id, {format: 'png', width: 500, height: 600, filename: `PopPK_Boxplot_${p.name}`});
                    
                    if (idx === params.length - 1) {
                        setTimeout(() => { 
                            txtExportPNG.innerText = originalText; 
                            btnExportPNG.disabled = false; 
                        }, 800);
                    }
                }, idx * 1000); 
            });
        }
    });

    addListener(btnExportCSV, 'click', () => {
        hasReviewedResults = true;
        updateFlowSetupState();
        if (currentTab === 'view-stats') {
            exportStatsToCSV();
        } else if (currentTab === 'view-params') {
            exportParamsToCSV();
        } else if (currentTab === 'view-be') {
            exportBEDataToCSV();
        } else {
            exportDataToCSV();
        }
        saveSessionState();
    });

    if (btnExportBundle) {
        btnExportBundle.addEventListener('click', async () => {
            if (!globalTrialsData && globalObsData.length === 0) return;
            hasReviewedResults = true;
            updateFlowSetupState();
            const originalBundleText = txtExportBundle ? txtExportBundle.textContent : 'Export Bundle';

            const zipCtor = window.JSZip;
            if (!zipCtor) {
                if (txtExportBundle) txtExportBundle.textContent = 'Fallback Export...';
                if (globalTrialsData) exportDataToCSV('GastroPlus_Aggregated_Profiles_Bundle.csv');
                if (globalTrialsData) exportStatsToCSV('GastroPlus_Summary_Stats_Bundle.csv');
                if (globalTrialsData) exportParamsToCSV('GastroPlus_PK_Parameters_Bundle.csv');
                if (globalBEData.length > 0) exportBEDataToCSV('GastroPlus_BE_Results_Bundle.csv');
                showStatusToast('Bundle library unavailable. Exported separate CSV files instead.', 'warn');
                if (txtExportBundle) txtExportBundle.textContent = originalBundleText;
                return;
            }

            try {
                btnExportBundle.disabled = true;
                if (txtExportBundle) txtExportBundle.textContent = 'Preparing...';
                const zip = new zipCtor();

                if (globalTrialsData) {
                    const profileCSV = getProfileCSVContent();
                    const statsCSV = getStatsCSVContent();
                    const paramsCSV = getParamsCSVContent();
                    if (profileCSV) zip.file('GastroPlus_Aggregated_Profiles.csv', profileCSV);
                    if (statsCSV) zip.file('GastroPlus_Summary_Stats.csv', statsCSV);
                    if (paramsCSV) zip.file('GastroPlus_PK_Parameters.csv', paramsCSV);
                }
                if (globalBEData.length > 0) {
                    const beCSV = getBECSVContent();
                    if (beCSV) zip.file('GastroPlus_BE_Results.csv', beCSV);
                }

                const dynamicBoxPlots = Array.from(document.querySelectorAll('[id^="plotlyBox_"]')).map(el => {
                    const label = (el.dataset.paramLabel || el.id.replace(/^plotlyBox_/, '')).replace(/\s+/g, '_');
                    return { id: el.id, name: `PK_Boxplot_${label}` };
                });

                const plotTargets = [
                    { id: 'plotlyChart', name: 'Profile_Conc_Time' },
                    ...bePlotTargets,
                    ...dynamicBoxPlots
                ];

                const activePlotTargets = getRenderablePlotTargets(plotTargets);

                for (let idx = 0; idx < activePlotTargets.length; idx++) {
                    const plot = activePlotTargets[idx];
                    if (txtExportBundle) txtExportBundle.textContent = `Rendering ${idx + 1}/${activePlotTargets.length}`;
                    const el = document.getElementById(plot.id);
                    const dataUrl = await Plotly.toImage(el, { format: 'png', width: 1100, height: 700 });
                    const base64 = dataUrl.split(',')[1];
                    if (base64) zip.file(`${plot.name}.png`, base64, { base64: true });
                }

                if (txtExportBundle) txtExportBundle.textContent = 'Compressing...';
                const blob = await zip.generateAsync({ type: 'blob' });
                const bundleName = `GastroPlus_Bundle_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.zip`;
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = bundleName;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                showStatusToast('Bundle downloaded successfully.');
            } catch (e) {
                console.error('Bundle export failed:', e);
                showStatusToast('Bundle export failed. Please retry.', 'error');
            } finally {
                btnExportBundle.disabled = false;
                if (txtExportBundle) txtExportBundle.textContent = originalBundleText;
            }
            saveSessionState();
        });
    }

    // -- Drag and Drop Setup --
    function setupDragAndDrop(zoneId, inputId, activeClass) {
        const zone = document.getElementById(zoneId);
        const input = document.getElementById(inputId);
        if(!zone || !input) return;
        let dragDepth = 0;

        const setActive = (active) => {
            zone.classList.toggle(activeClass, active);
        };

        zone.addEventListener('dragenter', (e) => {
            e.preventDefault();
            dragDepth += 1;
            setActive(true);
        });

        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            setActive(true);
        });

        zone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            dragDepth = Math.max(0, dragDepth - 1);
            if (dragDepth === 0) setActive(false);
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            dragDepth = 0;
            setActive(false);
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                input.files = e.dataTransfer.files;
                input.dispatchEvent(new Event('change'));
            }
        });
    }

    setupDragAndDrop('simRefDropZone', 'fileInputRef', 'dropzone-active-sim');
    setupDragAndDrop('simTestDropZone', 'fileInputTest', 'dropzone-active-sim');
    setupDragAndDrop('obsDropZone', 'obsFileInput', 'dropzone-active-obs');


    function updateAllViews() {
        if(currentTab === 'view-profile') updatePlot();
        if(currentTab === 'view-params') updateBoxPlots();
        if(currentTab === 'view-be') updateBEView();
        updateStatsTable();
        updateRunStatBar();
    }

    // -- File Handling Logic (Simulated) --
    async function handleSimUpload(event, formulationType) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        activateTab('view-profile');

        showLoading(true);
        emptyState.classList.add('hidden');
        if (formulationType === 'reference') simRefFileCount += files.length;
        else simTestFileCount += files.length;
        updateFileBadges();
        
        try {
            const settled = await Promise.allSettled(files.map(file => processSimFile(file, formulationType)));
            const parsedFilesData = [];
            settled.forEach((s, idx) => {
                const fileName = files[idx].name;
                if (s.status === 'fulfilled') {
                    parsedFilesData.push(s.value);
                    const d = s.value.diagnostics || {};
                    pushDiagnostic('sim', fileName, d.warn ? 'warn' : 'ok', d.message || 'Parsed successfully.');
                } else {
                    const errText = s.reason && s.reason.message ? s.reason.message : 'Parse failed';
                    pushDiagnostic('sim', fileName, 'error', errText);
                }
            });

            if (parsedFilesData.length > 0) {
                processTrialsData(parsedFilesData);
            } else if (!globalTrialsData && globalObsData.length === 0) {
                emptyState.classList.remove('hidden');
            }
        } finally {
            showLoading(false);
            updateFlowSetupState();
            saveSessionState();
        }
    }

    function processSimFile(file, formulationType) {
        return new Promise((resolve, reject) => {
            const ext = file.name.split('.').pop().toLowerCase();
            const reader = new FileReader();
            const mapping = getMappingConfig();
            const simTimeNeedle = (mapping.simTimeContains || columnMappingDefaults.simTimeContains).toLowerCase();
            const simSubjectPrefix = (mapping.simSubjectPrefix || columnMappingDefaults.simSubjectPrefix).toLowerCase();

            reader.onload = function(e) {
                try {
                    let cpText = "", statsText = null, paramsText = null;
                    let detectedSheetName = file.name; // fallback for CSV: use filename
                    
                    if (ext === 'xlsx' || ext === 'xls') {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: 'array'});
                        
                        let cpSheetName = workbook.SheetNames.find(n => n.toLowerCase().startsWith('cp-') || n.toLowerCase().includes('cp'));
                        let statsSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('summary stats') || n.toLowerCase().includes('statistics'));
                        let paramsSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('subj params') || n.toLowerCase().includes('subject params'));

                        if (!cpSheetName) cpSheetName = workbook.SheetNames[0];
                        detectedSheetName = cpSheetName || file.name;
                        
                        cpText = XLSX.utils.sheet_to_csv(workbook.Sheets[cpSheetName]);
                        if (statsSheetName) statsText = XLSX.utils.sheet_to_csv(workbook.Sheets[statsSheetName]);
                        if (paramsSheetName) paramsText = XLSX.utils.sheet_to_csv(workbook.Sheets[paramsSheetName]);
                        
                    } else if (ext === 'csv') {
                        cpText = e.target.result;
                    } else {
                        return reject(new Error("Unsupported format"));
                    }
                    
                    const result = { fileName: file.name, formulationType, profile: {}, stats: [], paramSamples: [], diagnostics: {}, concUnit: null };

                    // 1. Parse Cp Profile
                    const cpLines = cpText.split(/\r?\n/);
                    let headerLineIndex = cpLines.findIndex(l => l.toLowerCase().includes(simTimeNeedle));
                    if (headerLineIndex === -1) return reject(new Error(`No column matching '${mapping.simTimeContains || columnMappingDefaults.simTimeContains}' found in profile data.`));
                    
                    const cleanCpText = cpLines.slice(headerLineIndex).join('\n');
                    Papa.parse(cleanCpText, {
                        header: true, dynamicTyping: true, skipEmptyLines: 'greedy',
                        complete: function(parsed) {
                            const fieldNames = parsed.meta && Array.isArray(parsed.meta.fields) ? parsed.meta.fields : [];
                            // Detect concentration unit from sheet name and headers
                            result.concUnit = detectConcUnit(detectedSheetName, fieldNames);
                            const timeField = fieldNames.find(f => String(f).toLowerCase().includes(simTimeNeedle));
                            let subjectFields = fieldNames.filter(f => String(f).toLowerCase().startsWith(simSubjectPrefix));
                            if (subjectFields.length === 0 && simSubjectPrefix !== 's-') {
                                subjectFields = fieldNames.filter(f => String(f).match(/^S-\d+/i));
                            }
                            if (!timeField) { reject(new Error('SimTime column missing.')); return; }
                            if (subjectFields.length === 0) { reject(new Error(`No subject columns found with prefix '${mapping.simSubjectPrefix || columnMappingDefaults.simSubjectPrefix}'.`)); return; }
                            
                            const times = [];
                            const subjects = {};
                            let nonPositiveCount = 0;
                            subjectFields.forEach(sf => subjects[sf] = []);

                            parsed.data.forEach(row => {
                                const t = parseNumericCell(row[timeField]);
                                if (!Number.isFinite(t)) return;
                                times.push(t);
                                subjectFields.forEach(sf => {
                                    const rawVal = parseNumericCell(row[sf]);
                                    if (Number.isFinite(rawVal) && rawVal <= 0) nonPositiveCount += 1;
                                    subjects[sf].push(Number.isFinite(rawVal) ? rawVal : null);
                                });
                            });
                            result.profile = { times, subjects };
                            result.diagnostics.nonPositiveCount = nonPositiveCount;
                        }
                    });

                    // 2. Parse Summary Stats
                    if (statsText) {
                        const statsLines = statsText.split(/\r?\n/);
                        let sHeaderIdx = statsLines.findIndex(l => l.toLowerCase().startsWith('endpoint'));
                        if (sHeaderIdx !== -1) {
                            const cleanStatsText = statsLines.slice(sHeaderIdx).join('\n');
                            Papa.parse(cleanStatsText, {
                                header: true, dynamicTyping: true, skipEmptyLines: 'greedy',
                                complete: function(parsed) { result.stats = parsed.data; }
                            });
                        }
                    }

                    // 3. Parse Subject Params
                    if (paramsText) {
                        const paramsLines = paramsText.split(/\r?\n/);
                        let pHeaderIdx = paramsLines.findIndex(l => {
                            const s = l.toLowerCase();
                            return s.includes('cmax') || s.includes('auc') || s.includes('tmax') || s.includes('fa') || s.includes('fdp') || /\bf\b/.test(s);
                        });
                        if (pHeaderIdx !== -1) {
                            const cleanParamsText = paramsLines.slice(pHeaderIdx).join('\n');
                            Papa.parse(cleanParamsText, {
                                header: true, dynamicTyping: true, skipEmptyLines: 'greedy',
                                complete: function(parsed) { result.paramSamples = parsed.data; }
                            });
                        }
                    }

                    const subjectCount = Object.keys(result.profile.subjects || {}).length;
                    const warn = (result.diagnostics.nonPositiveCount || 0) > 0;
                    result.diagnostics = {
                        warn,
                        message: `Rows: ${(result.profile.times || []).length}, Subjects: ${subjectCount}, Stats rows: ${result.stats.length}, Param rows: ${result.paramSamples.length}${warn ? `, Non-positive conc: ${result.diagnostics.nonPositiveCount}` : ''}`
                    };

                    resolve(result);

                } catch (err) { reject(err); }
            };
            reader.onerror = reject;
            ext === 'csv' ? reader.readAsText(file) : reader.readAsArrayBuffer(file);
        });
    }

    // -- File Handling Logic (Observed Data) --
    async function handleObsUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        activateTab('view-profile');

        showLoading(true);
        emptyState.classList.add('hidden');
        try {
            const settled = await Promise.allSettled(files.map(file => processObsFile(file)));
            const parsedObsData = [];
            settled.forEach((s, idx) => {
                const fileName = files[idx].name;
                if (s.status === 'fulfilled') {
                    parsedObsData.push(s.value);
                    const warn = s.value.nonPositiveCount > 0;
                    pushDiagnostic('obs', fileName, warn ? 'warn' : 'ok', `Points: ${s.value.x.length}${warn ? `, Non-positive conc: ${s.value.nonPositiveCount}` : ''}`);
                } else {
                    const errText = s.reason && s.reason.message ? s.reason.message : 'Parse failed';
                    pushDiagnostic('obs', fileName, 'error', errText);
                }
            });

            if (parsedObsData.length > 0) {
                globalObsData = globalObsData.concat(parsedObsData.map(({ fileName, x, y }) => ({ fileName, x, y })));
                obsFileCount += parsedObsData.length;
                updateFileBadges();
                updateObsFooter();
                applyExportControlsForView(currentTab);
                if(currentTab==='view-profile') updatePlot();
            } else if (!globalTrialsData && globalObsData.length === 0) {
                emptyState.classList.remove('hidden');
            }
        } finally {
            showLoading(false);
            updateFlowSetupState();
            saveSessionState();
        }
    }

    function processObsFile(file) {
        return new Promise((resolve, reject) => {
            const ext = file.name.split('.').pop().toLowerCase();
            const reader = new FileReader();
            const mapping = getMappingConfig();
            const obsTimeNeedles = splitKeywords(mapping.obsTimeContains, columnMappingDefaults.obsTimeContains);
            const obsConcNeedles = splitKeywords(mapping.obsConcContains, columnMappingDefaults.obsConcContains);
            reader.onload = function(e) {
                try {
                    let text;
                    if (ext === 'csv') {
                        text = e.target.result;
                    } else {
                        const wb = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
                        text = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
                    }
                    Papa.parse(text, {
                        header: true, dynamicTyping: true, skipEmptyLines: 'greedy',
                        complete: function(results) {
                            const fieldNames = results.meta.fields;
                            if (!fieldNames || fieldNames.length < 2) return reject(new Error('Observed file has insufficient columns.'));
                            let timeField = fieldNames.find(f => f && obsTimeNeedles.some(k => String(f).toLowerCase().includes(k))) || fieldNames[0];
                            let concField = fieldNames.find(f => f && obsConcNeedles.some(k => String(f).toLowerCase().includes(k))) || fieldNames[1];
                            const x = [], y = [];
                            let nonPositiveCount = 0;
                            results.data.forEach(row => {
                                if (row[timeField] != null && row[concField] != null && row[timeField] !== '' && row[concField] !== '') {
                                    const t = parseNumericCell(row[timeField]);
                                    const c = parseNumericCell(row[concField]);
                                    if (Number.isFinite(t) && Number.isFinite(c)) {
                                        if (c <= 0) nonPositiveCount += 1;
                                        x.push(t);
                                        y.push(c);
                                    }
                                }
                            });
                            if (x.length === 0) return reject(new Error('No valid numeric rows found in observed file.'));
                            resolve({ fileName: file.name, x, y, nonPositiveCount });
                        }, error: reject
                    });
                } catch (err) { reject(err); }
            };
            reader.onerror = reject;
            ext === 'csv' ? reader.readAsText(file) : reader.readAsArrayBuffer(file);
        });
    }

    // -- Data Assembly --
    function processTrialsData(allFilesData) {
        let totalSubjects = 0;
        let existingTrials = globalTrialsData ? globalTrialsData.trials : [];

        // ── Determine canonical concentration unit for this ingest batch ──────────
        // Priority: existing session's targetConcUnit → first reference trial's unit
        //           → first any detected unit → 'ug/ml' fallback
        if (!targetConcUnit) {
            const refWithUnit = allFilesData.find(d => d.formulationType === 'reference' && d.concUnit);
            const anyWithUnit = allFilesData.find(d => d.concUnit);
            targetConcUnit = (refWithUnit || anyWithUnit)
                ? normConcUnit((refWithUnit || anyWithUnit).concUnit)
                : 'ug/ml';
        }

        // Warn only for true cross-file mismatches (not merely different from target unit).
        // Files can be consistently in one unit and still be safely normalised to targetConcUnit.
        const detectedUnits = [...new Set(
            allFilesData
                .map(d => normConcUnit(d.concUnit))
                .filter(Boolean)
        )];
        if (detectedUnits.length > 1) {
            showStatusToast(
                `Unit mismatch detected across files (${detectedUnits.join(', ')}). Converting all values to ${targetConcUnit}.`,
                'warn'
            );
        }

        // Auto-update Y-axis title to reflect the canonical unit
        if (inputYTitle && (inputYTitle.value === '' || inputYTitle.value.startsWith('Concentration'))) {
            inputYTitle.value = `Concentration (${targetConcUnit})`;
        }

        const newTrials = allFilesData.map((fileData, idx) => {
            const subjects = Object.keys(fileData.profile.subjects);
            totalSubjects += subjects.length;

            const times = fileData.profile.times;
            // Compute conversion factor: multiply raw values by this to reach targetConcUnit
            const factor = concConvFactor(fileData.concUnit || targetConcUnit, targetConcUnit);

            const concsAtTime = times.map((t, idx2) => {
                const concs = [];
                subjects.forEach(sf => {
                    const val = fileData.profile.subjects[sf][idx2];
                    if (Number.isFinite(val)) concs.push(val * factor);
                });
                return concs;
            });

            const individualLines = subjects.map(sf => ({
                x: times,
                y: fileData.profile.subjects[sf].map(v => Number.isFinite(v) ? v * factor : null)
            }));

            // Scale concentration-dependent columns in rawParams (Cmax, AUC, etc.)
            let rawParams = fileData.paramSamples;
            if (factor !== 1 && rawParams && rawParams.length > 0) {
                const paramKeys = Object.keys(rawParams[0]);
                const concKeys = paramKeys.filter(k => isConcDepParam(k));
                if (concKeys.length > 0) {
                    rawParams = rawParams.map(row => {
                        const r = Object.assign({}, row);
                        concKeys.forEach(k => {
                            const parsed = parseNumericCell(r[k]);
                            if (Number.isFinite(parsed)) r[k] = parsed * factor;
                        });
                        return r;
                    });
                }
            }

            // Scale concentration-dependent rows in rawStats (Cmax, AUC endpoint rows)
            let rawStats = fileData.stats;
            if (factor !== 1 && rawStats && rawStats.length > 0) {
                const epKey = Object.keys(rawStats[0]).find(k => k && (normalizeKey(k).includes('endpoint') || normalizeKey(k).includes('parameter')));
                if (epKey) {
                    rawStats = rawStats.map(row => {
                        if (!isConcDepParam(String(row[epKey] || ''))) return row;
                        const r = Object.assign({}, row);
                        Object.keys(r).forEach(k => {
                            if (k === epKey) return;
                            r[k] = scaleStatsCellValueByKey(k, r[k], factor);
                        });
                        return r;
                    });
                }
            }

            return {
                fileName: fileData.fileName,
                trialNumber: getTrialNumber(fileData.fileName, existingTrials.length + idx + 1),
                displayName: '',
                formulationType: fileData.formulationType === 'reference' ? 'reference' : 'test',
                concUnit: fileData.concUnit || targetConcUnit,
                concFactor: factor,
                times: times,
                concsAtTime: concsAtTime,
                individuals: individualLines,
                subjectCount: subjects.length,
                rawStats: rawStats,
                rawParams: rawParams,
                active: true,
                stats: {}
            };
        });

        const combinedTrials = existingTrials.concat(newTrials);
        totalSubjects = combinedTrials.reduce((sum, t) => sum + t.subjectCount, 0);

        globalTrialsData = { trials: combinedTrials, totalSubjects, totalTrials: combinedTrials.length };

        const hasStoredRef = beReferenceIndex !== '' && combinedTrials[beReferenceIndex] && combinedTrials[beReferenceIndex].formulationType === 'reference';
        const refIdxExisting = combinedTrials.findIndex(t => t.formulationType === 'reference');
        if (!hasStoredRef) {
            if (refIdxExisting >= 0) beReferenceIndex = String(refIdxExisting);
            else beReferenceIndex = '';
        }

        statSubjects.innerText = totalSubjects;
        statTrials.innerText = combinedTrials.length;
        statActiveTrials.innerText = combinedTrials.filter(t => t.active).length;
        trialSelectionPanel.classList.remove('hidden');
        updateFlowSetupState();
        setRailActive(btnSidebarPopulation);
        applyExportControlsForView(currentTab);

        renderTrialList();
        updateBERefSelect();
        updateAllViews();

        if (!hasAutoOpenedAxes && profileAxesBody && btnProfileAxesToggle && profileAxesBody.classList.contains('collapsed')) {
            toggleSection('profileAxesBody', btnProfileAxesToggle);
            hasAutoOpenedAxes = true;
            showStatusToast('Axes controls expanded. You can set custom titles and ranges here.', 'warn');
        }

        saveSessionState();
    }

    function renderTrialList() {
        if (!globalTrialsData) return;
        trialList.innerHTML = '';
        globalTrialsData.trials.forEach((trial, index) => {
            const colorObj = chartColors[index % chartColors.length];
            const div = document.createElement('div');
            const isRef = String(index) === String(beReferenceIndex);
            div.className = `trial-row ${isRef ? 'reference' : ''} ${trial.active ? 'is-active' : ''}`;
            const displayName = getTrialLabel(trial);
            const safeDisplayName = escapeHtml(displayName);
            const isRefType = trial.formulationType === 'reference';
            const safeFileName = escapeHtml(trial.fileName);
            // Shorten filename for display: strip extension, cap at 28 chars
            const shortFileName = trial.fileName.replace(/\.[^/.]+$/, '').substring(0, 32);
            const safeShortFileName = escapeHtml(shortFileName);

            div.innerHTML = `
                <div class="trial-row-top">
                    <input type="checkbox" id="chk_${index}" class="cursor-pointer flex-shrink-0"
                        style="accent-color: ${colorObj.hex};" ${trial.active ? 'checked' : ''}>
                    <span class="trial-color-dot" style="background:${colorObj.hex};"></span>
                    <label for="chk_${index}" class="trial-label-text cursor-pointer" style="color: var(--text-primary);" title="${safeFileName}">${safeDisplayName}</label>
                    <button type="button" data-rename-index="${index}" class="trial-rename-btn w-5 h-5 inline-flex items-center justify-center rounded flex-shrink-0" style="color: var(--text-tertiary); border: 1px solid var(--border-subtle); background: var(--bg-base);" title="Rename trial label" aria-label="Rename trial label">
                        <svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
                    </button>
                </div>
                <div class="trial-row-sub">
                    <span class="trial-filename-sub" title="${safeFileName}">${safeShortFileName}</span>
                    <button type="button" data-formulation-index="${index}" class="trial-formulation-pill ${isRefType ? 'ref' : 'test'}" title="Click to toggle formulation type">${isRefType ? 'REF' : 'TEST'}</button>
                </div>
            `;
            
            div.querySelector('input').addEventListener('change', (e) => {
                trial.active = e.target.checked;
                statActiveTrials.innerText = globalTrialsData.trials.filter(t => t.active).length;
                updateAllViews();
            });

            div.querySelector('[data-formulation-index]').addEventListener('click', (e) => {
                // Toggle between reference and test
                const val = trial.formulationType === 'reference' ? 'test' : 'reference';
                trial.formulationType = val;

                if (trial.formulationType === 'reference' && beReferenceIndex === '') {
                    beReferenceIndex = String(index);
                }

                if (trial.formulationType === 'test' && String(index) === String(beReferenceIndex)) {
                    const nextRefIdx = globalTrialsData.trials.findIndex(t => t.formulationType === 'reference');
                    beReferenceIndex = nextRefIdx >= 0 ? String(nextRefIdx) : '';
                }

                if (trial.formulationType === 'reference' && String(index) === String(beTestIndex)) {
                    beTestIndex = '';
                }

                renderTrialList();
                updateBERefSelect();
                updateAllViews();
                saveSessionState();
            });

            div.querySelector('[data-rename-index]').addEventListener('click', () => {
                const current = getTrialLabel(trial);
                const next = window.prompt('Set trial display name:', current);
                if (next === null) return;
                const cleaned = next.trim();
                trial.displayName = cleaned;
                renderTrialList();
                updateBERefSelect();
                updateAllViews();
                saveSessionState();
            });

            trialList.appendChild(div);
        });
        statActiveTrials.innerText = globalTrialsData.trials.filter(t => t.active).length;
    }

    function updateBERefSelect() {
        beRefTrialSelect.innerHTML = '<option value="">-- Select Reference Trial --</option>';
        if (beTestTrialSelect) {
            beTestTrialSelect.innerHTML = '<option value="">All Test Trials</option>';
        }

        if (!globalTrialsData) {
            if (btnRunBE) {
                btnRunBE.disabled = true;
                btnRunBE.style.opacity = '0.55';
                btnRunBE.style.cursor = 'not-allowed';
                btnRunBE.title = 'Upload simulation data before running BE analysis.';
            }
            return;
        }

        const { refActiveIndices, testActiveIndices } = getActiveBETrialIndices();
        const refIndices = refActiveIndices;
        const testIndices = testActiveIndices;

        refIndices.forEach(i => {
            const trial = globalTrialsData.trials[i];
            beRefTrialSelect.innerHTML += `<option value="${i}">${escapeHtml(getTrialLabel(trial))}</option>`;
        });

        if (beTestTrialSelect) {
            testIndices.forEach(i => {
                const trial = globalTrialsData.trials[i];
                beTestTrialSelect.innerHTML += `<option value="${i}">${escapeHtml(getTrialLabel(trial))}</option>`;
            });
        }

        if (beReferenceIndex === '' && refIndices.length > 0) {
            beReferenceIndex = String(refIndices[0]);
        }
        if (beReferenceIndex !== '' && !refIndices.includes(Number(beReferenceIndex))) {
            beReferenceIndex = refIndices.length ? String(refIndices[0]) : '';
        }
        if (beTestIndex !== '' && !testIndices.includes(Number(beTestIndex))) {
            beTestIndex = '';
        }

        if (beReferenceIndex !== '' && globalTrialsData.trials[beReferenceIndex]) {
            beRefTrialSelect.value = String(beReferenceIndex);
            if (beRefSummary) beRefSummary.textContent = `Reference: ${getTrialLabel(globalTrialsData.trials[beReferenceIndex])}`;
        } else if (beRefSummary) {
            beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
        }

        if (beTestTrialSelect) {
            beTestTrialSelect.value = beTestIndex;
        }

        if (btnRunBE) {
            const method = beMethodSelect ? beMethodSelect.value : 'simple';
            const needsPairs = method === 'paired-trial-number';
            const hasValidRef = refIndices.length > 0;
            const hasValidTest = testIndices.length > 0;
            const pairedEligibility = needsPairs
                ? getPairedTrialNumberEligibility(refIndices, testIndices)
                : { matchedUniqueCount: 0 };
            const canRun = needsPairs
                ? pairedEligibility.matchedUniqueCount > 0
                : (hasValidRef && hasValidTest);
            btnRunBE.disabled = !canRun;
            btnRunBE.style.opacity = canRun ? '1' : '0.55';
            btnRunBE.style.cursor = canRun ? 'pointer' : 'not-allowed';
            btnRunBE.title = canRun
                ? 'Run Bioequivalence analysis'
                : (needsPairs
                    ? 'Need at least one unique matched active trial number between Reference and Test trials.'
                    : 'Need at least one active Reference and one active Test trial with PK parameter rows.');
        }
    }

    // -- Math Utilities --
    const getTrialNumber = (fileName, fallback) => {
        const baseName = fileName.replace(/\.[^/.]+$/, "");
        const endPatterns = [ /(?:trial|run|sim(?:ulation)?)[\s_-]*(\d+)$/i, /[_-](?:t|s)[\s_-]*(\d+)$/i ];
        for (let p of endPatterns) { 
            const match = baseName.match(p); 
            if (match) return parseInt(match[1], 10); 
        }
        const generalPatterns = [ /(?:trial|run|sim(?:ulation)?)[\s_-]*(\d+)/i, /^(\d+)[\s_-]/, /\b(?:t|s)[\s_-]*(\d+)\b/i ];
        for (let p of generalPatterns) { 
            const match = baseName.match(p); 
            if (match) return parseInt(match[1], 10); 
        }
        return fallback; 
    };

    const calculatePercentile = (arr, p) => {
        if (arr.length === 0) return null;
        const sorted = arr.slice().sort((a, b) => a - b);
        const index = (sorted.length - 1) * p;
        const lower = Math.floor(index);
        const upper = lower + 1;
        const weight = index % 1;
        return upper >= sorted.length ? sorted[lower] : sorted[lower] * (1 - weight) + sorted[upper] * weight;
    };

    const calculateMeanAndCI = (arr) => {
        if (arr.length === 0) return { mean: null, lowerCI: null, upperCI: null };
        const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
        const sumSqDiff = arr.reduce((acc, val) => acc + (val - mean) ** 2, 0);
        const stdErr = Math.sqrt(sumSqDiff / (arr.length > 1 ? arr.length - 1 : 1)) / Math.sqrt(arr.length);
        const tVal = getTValue(arr.length - 1);
        const marginOfError = tVal * stdErr;
        return { mean, lowerCI: mean - marginOfError, upperCI: mean + marginOfError };
    };

    function normalizeKey(k) {
        return String(k || '').toLowerCase().replace(/\s+/g, ' ').trim();
    }

    function getActiveBETrialIndices() {
        const refActiveIndices = [];
        const testActiveIndices = [];
        if (!globalTrialsData || !Array.isArray(globalTrialsData.trials)) {
            return { refActiveIndices, testActiveIndices };
        }

        globalTrialsData.trials.forEach((t, i) => {
            const hasParams = !!(t && t.rawParams && t.rawParams.length > 0);
            if (!hasParams || !t.active) return;
            if (t.formulationType === 'reference') refActiveIndices.push(i);
            else if (t.formulationType === 'test') testActiveIndices.push(i);
        });

        return { refActiveIndices, testActiveIndices };
    }

    function getPairedTrialNumberEligibility(refActiveIndices, testActiveIndices) {
        const refCounts = new Map();
        const testCounts = new Map();

        refActiveIndices.forEach(i => {
            const n = Number(globalTrialsData.trials[i].trialNumber);
            if (!Number.isFinite(n)) return;
            refCounts.set(n, (refCounts.get(n) || 0) + 1);
        });
        testActiveIndices.forEach(i => {
            const n = Number(globalTrialsData.trials[i].trialNumber);
            if (!Number.isFinite(n)) return;
            testCounts.set(n, (testCounts.get(n) || 0) + 1);
        });

        const matchedUniqueNumbers = [];
        refCounts.forEach((refCount, n) => {
            const testCount = testCounts.get(n);
            if (refCount === 1 && testCount === 1) matchedUniqueNumbers.push(n);
        });

        return { matchedUniqueCount: matchedUniqueNumbers.length };
    }

    function normalizeParamToken(k) {
        return String(k || '')
            .toLowerCase()
            .replace(/\[[^\]]*\]/g, '')
            .replace(/\([^\)]*\)/g, '')
            .replace(/[^a-z0-9]+/g, '')
            .trim();
    }

    function prettifyParamLabel(k) {
        const raw = String(k || '').trim();
        if (!raw) return 'Parameter';
        const clean = raw.replace(/\s+/g, ' ');
        return clean.charAt(0).toUpperCase() + clean.slice(1);
    }

    function shortenAxisLabel(label, maxLen = 24) {
        const s = String(label || '').trim();
        if (s.length <= maxLen) return s;
        return `${s.slice(0, Math.max(8, maxLen - 1)).trim()}…`;
    }

    function parseNumericCell(val) {
        if (typeof val === 'number') return val;
        if (val === null || val === undefined) return null;
        const s = String(val).trim().replace(/\u00A0/g, '').replace(/\s+/g, '');
        if (!s) return null;
        const direct = Number(s);
        if (!isNaN(direct)) return direct;

        // Locale-aware normalization for values like "1,234.56" or "1.234,56".
        let normalized = s;
        const hasComma = s.includes(',');
        const hasDot = s.includes('.');
        if (hasComma && hasDot) {
            normalized = s.lastIndexOf(',') > s.lastIndexOf('.')
                ? s.replace(/\./g, '').replace(/,/g, '.')
                : s.replace(/,/g, '');
        } else if (hasComma) {
            normalized = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(/,/g, '.');
        }

        const normalizedNum = Number(normalized);
        if (!isNaN(normalizedNum)) return normalizedNum;

        const tokenMatch = String(val).match(NUMERIC_TOKEN_REGEX);
        return tokenMatch ? parseNumericCell(tokenMatch[0]) : null;
    }

    function canonicalParamNameFromKey(key) {
        const n = normalizeParamToken(key);
        if (!n) return null;
        if (n === 'cmax' || n.startsWith('cmax') || n.includes('maxconcentration') || n.includes('peakconcentration')) return 'Cmax';
        if (n.includes('auc0t') || n === 'auct' || n.startsWith('auct') || n.includes('auclast') || n.includes('auc0last') || n.includes('auctlast')) return 'AUCt';
        if (n.includes('aucinf') || n.includes('auc0inf') || n.includes('aucinfinity')) return 'AUCinf';
        if (n === 'tmax' || n.startsWith('tmax') || n.includes('timetomax') || n.includes('timetopeak')) return 'Tmax';
        if (n === 'fdp' || n.startsWith('fdp')) return 'Fdp';
        if (n === 'fa' || n.startsWith('fa')) return 'Fa';
        if (n === 'f' || n.includes('bioavailability')) return 'F';
        return null;
    }

    function isLikelyMetadataParam(key) {
        const n = normalizeParamToken(key);
        return /^(subject|subj|id|period|sequence|seq|treatment|trt|arm|cohort|group|formulation|file|trial)$/.test(n);
    }

    function paramHasFiniteValues(trial, key) {
        if (!trial || !trial.rawParams || trial.rawParams.length === 0) return false;
        return trial.rawParams.some(row => Number.isFinite(parseNumericCell(row[key])));
    }

    function findParamKey(keys, paramType) {
        const normalized = keys.map(k => ({ raw: k, n: normalizeKey(k) }));
        const tokenized = keys.map(k => ({ raw: k, t: normalizeParamToken(k) }));

        if (paramType === 'Cmax') {
            const cmaxAliases = ['cmax', 'max concentration', 'peak concentration'];
            const found = normalized.find(x => cmaxAliases.some(a => x.n.includes(a)));
            return found ? found.raw : null;
        }

        if (paramType === 'AUCt') {
            const aucTAliases = ['auc(0-t)', 'auc 0-t', 'auc0-t', 'auct', 'auc(last)', 'auclast', 'auc last', 'auc0-last'];
            const found = normalized.find(x => aucTAliases.some(a => x.n.includes(a)));
            return found ? found.raw : null;
        }

        if (paramType === 'AUCinf') {
            const aucInfAliases = ['aucinf', 'auc inf', 'auc(inf)', 'auc0-inf', 'auc(0-inf)', 'auc infinity'];
            const found = normalized.find(x => aucInfAliases.some(a => x.n.includes(a)));
            if (found) return found.raw;

            const fallback = normalized.find(x => x.n.includes('auc') && !x.n.includes('0-t') && !x.n.includes('last') && !x.n.includes('auct'));
            return fallback ? fallback.raw : null;
        }

        if (paramType === 'Tmax') {
            const found = normalized.find(x => x.n.includes('tmax') || x.n.includes('time to max') || x.n.includes('time to peak'));
            return found ? found.raw : null;
        }

        if (paramType === 'Fa') {
            const found = tokenized.find(x => x.t === 'fa' || x.t.startsWith('fa'));
            return found ? found.raw : null;
        }

        if (paramType === 'Fdp') {
            const found = tokenized.find(x => x.t === 'fdp' || x.t.startsWith('fdp'));
            return found ? found.raw : null;
        }

        if (paramType === 'F') {
            const found = tokenized.find(x => x.t === 'f' || x.t.includes('bioavailability'));
            return found ? found.raw : null;
        }

        return null;
    }

    const extractStatsRow = (row) => {
        let stats = { epName: '-', mean: '-', cv: '-', min: '-', max: '-', geom: '-', ci90: '-', ci90ln: '-' };
        
        for (let key of Object.keys(row)) {
            if (!key) continue;
            let k = key.trim().toLowerCase();
            let val = row[key] !== null && row[key] !== undefined ? row[key] : '-';
            
            if (k.includes('endpoint') || k.includes('parameter')) stats.epName = String(val);
            else if (k === 'mean') stats.mean = val;
            else if (k.includes('cv') && !k.includes('geom')) stats.cv = val;
            else if (k === 'min') stats.min = val;
            else if (k === 'max') stats.max = val;
            else if (k.includes('geom') && !k.includes('cv')) stats.geom = val;
            else if (k === '90% ci') stats.ci90 = val;
            else if (k.includes('90% ci') && k.includes('ln')) stats.ci90ln = val; 
        }
        return stats;
    };

    // Statistical lookup for t-distribution (90% CI means two-tailed alpha=0.10)
    function getTValue(df) {
        if (df <= 0) return Infinity;
        const tTable = {
            1:6.314, 2:2.920, 3:2.353, 4:2.132, 5:2.015, 6:1.943, 7:1.895, 8:1.860, 9:1.833, 10:1.812, 
            11:1.796, 12:1.782, 13:1.771, 14:1.761, 15:1.753, 16:1.746, 17:1.740, 18:1.734, 19:1.729, 20:1.725, 
            21:1.721, 22:1.717, 23:1.714, 24:1.711, 25:1.708, 26:1.706, 27:1.703, 28:1.701, 29:1.699, 30:1.697, 
            40:1.684, 50:1.676, 60:1.671, 80:1.664, 100:1.660, 120:1.658
        };
        if(tTable[df]) return tTable[df];
        if (df > 120) return 1.645; // Approx for infinity
        
        const keys = Object.keys(tTable).map(Number).sort((a,b)=>a-b);
        for(let i=0; i<keys.length-1; i++){
            if(df > keys[i] && df < keys[i+1]){
                const frac = (df - keys[i]) / (keys[i+1] - keys[i]);
                return tTable[keys[i]] - frac * (tTable[keys[i]] - tTable[keys[i+1]]);
            }
        }
        return 1.645;
    }

    function calculateBE(refLogs, testLogs) {
        const nR = refLogs.length;
        const nT = testLogs.length;
        if(nR < 2 || nT < 2) return null;
        
        const muR = refLogs.reduce((a,b)=>a+b,0)/nR;
        const muT = testLogs.reduce((a,b)=>a+b,0)/nT;
        
        const varR = refLogs.reduce((a,b)=>a + Math.pow(b-muR, 2),0)/(nR-1);
        const varT = testLogs.reduce((a,b)=>a + Math.pow(b-muT, 2),0)/(nT-1);
        
        const sp = Math.sqrt(((nT-1)*varT + (nR-1)*varR)/(nT+nR-2));
        const se = sp * Math.sqrt(1/nT + 1/nR);
        
        const df = nT + nR - 2;
        const tVal = getTValue(df);
        
        const pe = Math.exp(muT - muR) * 100;
        const lower = Math.exp((muT - muR) - tVal*se) * 100;
        const upper = Math.exp((muT - muR) + tVal*se) * 100;
        
        return { pe, lower, upper };
    }

    function extractParamDataWithDiagnostics(trial, paramType) {
        if (!trial || !trial.rawParams || trial.rawParams.length === 0) return { values: [], excluded: 0, keyFound: false };
        const keys = Object.keys(trial.rawParams[0]);
        const k = findParamKey(keys, paramType);

        if (!k) return { values: [], excluded: 0, keyFound: false };

        const values = [];
        let excluded = 0;
        trial.rawParams.forEach(row => {
            const v = parseNumericCell(row[k]);
            if (!Number.isFinite(v)) return;
            if (v > 0) values.push(v);
            else excluded += 1;
        });

        return { values, excluded, keyFound: true };
    }

    function extractParamData(trial, paramType) {
        return extractParamDataWithDiagnostics(trial, paramType).values;
    }

    function getBoxplotParams() {
        if (!globalTrialsData || !Array.isArray(globalTrialsData.trials)) return [];
        const preferred = ['Cmax', 'AUCt', 'AUCinf', 'Tmax', 'Fa', 'Fdp', 'F'];
        const paramsByKey = new Map();
        let genericIdx = 0;

        preferred.forEach(param => {
            const foundInTrials = globalTrialsData.trials.some(trial => {
                if (!trial.active || !trial.rawParams || trial.rawParams.length === 0) return false;
                return !!findParamKey(Object.keys(trial.rawParams[0]), param);
            });
            if (foundInTrials) {
                paramsByKey.set(param.toLowerCase(), {
                    id: param.toLowerCase(),
                    label: param,
                    type: 'canonical',
                    order: preferred.indexOf(param)
                });
            }
        });

        globalTrialsData.trials.forEach(trial => {
            if (!trial.active || !trial.rawParams || trial.rawParams.length === 0) return;
            const keys = Object.keys(trial.rawParams[0]);
            keys.forEach(key => {
                if (!key || isLikelyMetadataParam(key) || !paramHasFiniteValues(trial, key)) return;

                const canonical = canonicalParamNameFromKey(key);
                if (canonical) {
                    const id = canonical.toLowerCase();
                    if (!paramsByKey.has(id)) {
                        paramsByKey.set(id, {
                            id,
                            label: canonical,
                            type: 'canonical',
                            order: preferred.indexOf(canonical)
                        });
                    }
                    return;
                }

                const token = normalizeParamToken(key);
                if (!token || paramsByKey.has(token)) return;

                genericIdx += 1;
                paramsByKey.set(token, {
                    id: `param${genericIdx}_${token}`,
                    label: prettifyParamLabel(key),
                    type: 'exact',
                    token,
                    order: 100 + genericIdx
                });
            });
        });

        return Array.from(paramsByKey.values()).sort((a, b) => {
            if (a.order !== b.order) return a.order - b.order;
            return a.label.localeCompare(b.label);
        });
    }

    function getVisibleBoxplotParams() {
        const params = getBoxplotParams();
        if (!boxResultsOnly) return params;
        const resultTokens = new Set(['cmax', 'auct', 'aucinf', 'fa', 'f', 'fdp']);
        return params.filter(p => {
            const canonicalId = String(p.id || '').toLowerCase();
            if (resultTokens.has(canonicalId)) return true;
            const token = String(p.token || '').toLowerCase();
            return resultTokens.has(token);
        });
    }

    function findKeyForBoxParam(keys, paramDef) {
        if (!Array.isArray(keys) || !paramDef) return null;
        if (paramDef.type === 'canonical') return findParamKey(keys, paramDef.label);
        const token = String(paramDef.token || '');
        if (!token) return null;
        return keys.find(k => normalizeParamToken(k) === token) || null;
    }

    function cancelBoxPlotRender() {
        boxPlotRenderToken += 1;
        if (boxPlotFrameId !== null) {
            window.cancelAnimationFrame(boxPlotFrameId);
            boxPlotFrameId = null;
        }
        setBoxPlotLoadingState(false);
    }

    function setBoxPlotLoadingState(isVisible, delayMs = 0) {
        if (!boxPlotLoadingMsg) return;

        boxPlotLoadingSeq += 1;
        const seq = boxPlotLoadingSeq;

        if (boxPlotLoadingTimer !== null) {
            window.clearTimeout(boxPlotLoadingTimer);
            boxPlotLoadingTimer = null;
        }

        const applyLoadingShell = (loading) => {
            if (viewParamsPanel) viewParamsPanel.classList.toggle('boxplot-rendering', loading);
            if (boxPlotsGridEl) boxPlotsGridEl.setAttribute('aria-busy', loading ? 'true' : 'false');
        };

        if (!isVisible) {
            applyLoadingShell(false);
            if (boxPlotProgressBarFill) boxPlotProgressBarFill.style.width = '0%';
            if (boxPlotProgressText) boxPlotProgressText.textContent = '0%';
            boxPlotLoadingMsg.classList.add('hidden');
            boxPlotLoadingMsg.classList.remove('flex');
            return;
        }

        // Hide stale plots immediately and lock scroll while rendering.
        applyLoadingShell(true);

        const showOverlayNow = () => {
            if (seq !== boxPlotLoadingSeq) return;
            if (boxPlotProgressBarFill) boxPlotProgressBarFill.style.width = '0%';
            if (boxPlotProgressText) boxPlotProgressText.textContent = '0%';
            boxPlotLoadingMsg.classList.remove('hidden');
            boxPlotLoadingMsg.classList.add('flex');
        };

        if (delayMs > 0) {
            boxPlotLoadingTimer = window.setTimeout(showOverlayNow, delayMs);
        } else {
            showOverlayNow();
        }
    }

    function getBoxPlotCardSignature(paramDefs) {
        return paramDefs.map(p => p.id).join('|');
    }

    function getBoxPlotRenderKey(paramDefs) {
        if (!globalTrialsData || !Array.isArray(globalTrialsData.trials)) return 'no-data';
        const activeTrials = globalTrialsData.trials
            .map((trial, index) => {
                if (!trial.active) return null;
                return [
                    index,
                    trial.fileName,
                    trial.rawParams ? trial.rawParams.length : 0,
                    getTrialLabel(trial)
                ].join(':');
            })
            .filter(Boolean);

        return JSON.stringify({
            params: paramDefs.map(p => p.id),
            trials: activeTrials,
            log: isLogScale,
            outliers: !!(boxShowOutliers && boxShowOutliers.checked),
            mean: !!(boxShowMean && boxShowMean.checked)
        });
    }

    function resizeRenderedBoxPlots() {
        Array.from(document.querySelectorAll('[id^="plotlyBox_"]')).forEach(el => {
            if (!el || !el.data) return;
            const plotHeight = getPlotPanelCanvasHeight(el);
            el.style.height = `${plotHeight}px`;
            Plotly.relayout(el, { height: plotHeight }).then(() => {
                Plotly.Plots.resize(el);
            });
        });
    }

    function getPlotPanelCanvasHeight(targetEl) {
        if (!targetEl) return 320;

        const panelBody = targetEl.closest('.plot-panel-body');
        const card = targetEl.closest('.plot-panel-card');
        const titleBlock = card ? card.querySelector('.shrink-0') : null;

        let height = panelBody ? panelBody.clientHeight : 0;

        if (!height && card) {
            const cardHeight = card.clientHeight;
            const titleHeight = titleBlock ? titleBlock.getBoundingClientRect().height : 0;
            height = cardHeight - titleHeight - 20;
        }

        return Math.max(260, Math.floor(height || 320));
    }

    function renderBoxPlotCards(paramDefs) {
        const grid = document.getElementById('boxPlotsGrid');
        if (!grid) return;
        const signature = getBoxPlotCardSignature(paramDefs);
        if (grid.dataset.paramSignature === signature) return;
        grid.innerHTML = paramDefs.map(p => `
            <div class="plot-panel-card rounded-lg p-2" style="background: var(--bg-surface); border: 1px solid var(--border-subtle); min-height: 320px;">
                <div class="shrink-0 px-1 pb-2">
                    <div class="text-[12px] font-semibold leading-tight" style="color: var(--text-primary);">${escapeHtml(p.label)}</div>
                    <div class="text-[10px] mt-0.5" style="color: var(--text-tertiary);">Distribution across active trials</div>
                </div>
                <div class="plot-panel-body rounded-md" style="border: 1px solid var(--border-subtle); background: var(--bg-base);">
                    <div id="plotlyBox_${p.id}" data-param-label="${escapeHtml(p.label)}" class="plot-panel-canvas text-[11px]" style="color: var(--text-tertiary);">Rendering…</div>
                </div>
            </div>
        `).join('');
        grid.dataset.paramSignature = signature;
        grid.dataset.renderKey = '';
    }

    // -- Plotting Logic: Concentration-Time --
    function updatePlot() {
        if (!globalTrialsData && globalObsData.length === 0) return;

        const traces = [];
        const overlayFlags = [];
        let sampledIndividuals = false;

        if (globalTrialsData) {
            globalTrialsData.trials.forEach((trial, index) => {
                if (!trial.active) return;
                const { times, concsAtTime, individuals } = trial;
                const hex = chartColors[index % chartColors.length].hex;
                const rgb = chartColors[index % chartColors.length].rgb;
                
                let displayName = getTrialLabel(trial);
                computeTrialProfileStats(trial);

                const legendGroup = `group_${index}`;

                traces.push({
                    x: [null], y: [null], type: 'scatter', mode: 'lines',
                    line: { color: hex, width: 3 },
                    legendgroup: legendGroup,
                    name: displayName,
                    showlegend: true,
                    hoverinfo: 'skip'
                });

                if (showContour100.checked) {
                    traces.push({ x: times, y: trial.stats.p100, type: 'scatter', mode: 'lines', line: { width: 1, color: `rgba(${rgb}, 0.5)`, dash: 'dot' }, legendgroup: legendGroup, name: `${displayName} (Max)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                    traces.push({ x: times, y: trial.stats.p00, type: 'scatter', mode: 'lines', fill: 'tonexty', fillcolor: `rgba(${rgb}, 0.05)`, line: { width: 1, color: `rgba(${rgb}, 0.5)`, dash: 'dot' }, legendgroup: legendGroup, name: `${displayName} (Min)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                }
                
                if (showContour95.checked) {
                    traces.push({ x: times, y: trial.stats.p975, type: 'scatter', mode: 'lines', line: { width: 1, color: `rgba(${rgb}, 0.6)`, dash: 'dash' }, legendgroup: legendGroup, name: `${displayName} (97.5th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                    traces.push({ x: times, y: trial.stats.p025, type: 'scatter', mode: 'lines', fill: 'tonexty', fillcolor: `rgba(${rgb}, 0.1)`, line: { width: 1, color: `rgba(${rgb}, 0.6)`, dash: 'dash' }, legendgroup: legendGroup, name: `${displayName} (2.5th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                }

                if (showContour90.checked) {
                    traces.push({ x: times, y: trial.stats.p95, type: 'scatter', mode: 'lines', line: { width: 1.5, color: `rgba(${rgb}, 0.6)`, dash: 'dash' }, legendgroup: legendGroup, name: `${displayName} (95th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                    traces.push({ x: times, y: trial.stats.p05, type: 'scatter', mode: 'lines', fill: 'tonexty', fillcolor: `rgba(${rgb}, 0.15)`, line: { width: 1.5, color: `rgba(${rgb}, 0.6)`, dash: 'dash' }, legendgroup: legendGroup, name: `${displayName} (5th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                }

                if (showContoursCheck.checked) {
                    traces.push({ x: times, y: trial.stats.p75, type: 'scatter', mode: 'lines', line: { width: 0, color: 'transparent' }, legendgroup: legendGroup, name: `${displayName} (75th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                    traces.push({ x: times, y: trial.stats.p25, type: 'scatter', mode: 'lines', fill: 'tonexty', fillcolor: `rgba(${rgb}, 0.25)`, line: { width: 0, color: 'transparent' }, legendgroup: legendGroup, name: `${displayName} (25th Perc)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                }

                if (showCIMean.checked) {
                    traces.push({ x: times, y: trial.stats.upperCI, type: 'scatter', mode: 'lines', line: { width: 0, color: 'transparent' }, legendgroup: legendGroup, name: `${displayName} (Upper 90% CI)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                    traces.push({ x: times, y: trial.stats.lowerCI, type: 'scatter', mode: 'lines', fill: 'tonexty', fillcolor: `rgba(${rgb}, 0.25)`, line: { width: 0, color: 'transparent' }, legendgroup: legendGroup, name: `${displayName} (Lower 90% CI)`, showlegend: false, hoverinfo: 'skip', connectgaps: true });
                }

                if (showIndividualsCheck.checked) {
                    const opacity = Math.max(0.1, Math.min(0.4, 4 / trial.subjectCount)).toFixed(3);
                    const maxIndividualTraces = 250;
                    const step = trial.subjectCount > maxIndividualTraces ? Math.ceil(trial.subjectCount / maxIndividualTraces) : 1;
                    if (step > 1) sampledIndividuals = true;
                    for (let i = 0; i < individuals.length; i += step) {
                        const ind = individuals[i];
                        traces.push({
                            x: ind.x, y: ind.y, type: 'scatter', mode: 'lines',
                            line: { color: `rgba(${rgb}, ${opacity})`, width: 1.5 },
                            legendgroup: legendGroup, showlegend: false, hoverinfo: 'skip'
                        });
                    }
                }

                if (showMedian.checked) {
                    traces.push({
                        x: times, y: trial.stats.medians, type: 'scatter', mode: 'lines',
                        line: { color: hex, width: 2, dash: 'dash', shape: 'spline', smoothing: 0.4 },
                        legendgroup: legendGroup, name: `${displayName} (Median)`, showlegend: false,
                        connectgaps: true,
                        hovertemplate: `${displayName} Median<br>Time: %{x:.2f} h<br>Conc: %{y:.4g}<extra></extra>`
                    });
                }

                if (showMean.checked) {
                    traces.push({
                        x: times, y: trial.stats.means, type: 'scatter', mode: 'lines',
                        line: { color: hex, width: 3, shape: 'spline', smoothing: 0.45 },
                        legendgroup: legendGroup,
                        name: `${displayName} (Mean)`, showlegend: false, connectgaps: true,
                        hovertemplate: `${displayName} Mean<br>Time: %{x:.2f} h<br>Conc: %{y:.4g}<extra></extra>`
                    });
                }

                if (showContour100.checked) overlayFlags.push('100%');
                if (showContour95.checked) overlayFlags.push('95%');
                if (showContour90.checked) overlayFlags.push('90%');
                if (showContoursCheck.checked) overlayFlags.push('50%');
                if (showCIMean.checked) overlayFlags.push('CI');
                if (showIndividualsCheck.checked) overlayFlags.push('Individuals');
            });
        }

        if (globalObsData.length > 0) {
            globalObsData.forEach((obs, idx) => {
                traces.push({
                    x: obs.x, y: obs.y, type: 'scatter', mode: 'markers',
                    marker: { symbol: 'diamond-open', size: 8, color: cssVar('--text-primary'), line: { color: cssVar('--bg-surface'), width: 1.3 } },
                    name: `Observed Data ${idx + 1}`,
                    hovertemplate: `Observed ${idx + 1}<br>Time: %{x:.2f} h<br>Conc: %{y:.4g}<extra></extra>`
                });
            });
        }

        const xMax = parseFloat(inputXMax.value);
        const yMax = parseFloat(inputYMax.value);
        
        let minPositiveY = Infinity;
        traces.forEach(trace => {
            if (!trace || !Array.isArray(trace.y)) return;
            trace.y.forEach(v => {
                if (Number.isFinite(v) && v > 0) minPositiveY = Math.min(minPositiveY, v);
            });
        });

        let yRange = undefined;
        if (!isNaN(yMax) && (!isLogScale || yMax > 0)) {
            const logFloor = Number.isFinite(minPositiveY) ? Math.max(minPositiveY * 0.5, 1e-6) : 1e-6;
            if (isLogScale) {
                const lower = Math.log10(logFloor);
                const upper = Math.log10(Math.max(yMax, logFloor * 1.01));
                yRange = upper > lower ? [lower, upper] : undefined;
            } else {
                yRange = [0, yMax];
            }
        }

        const activeTrials = globalTrialsData ? globalTrialsData.trials.filter(t => t.active).length : 0;
        const subtitleParts = [];
        if (activeTrials > 0) subtitleParts.push(`${activeTrials} active trial${activeTrials > 1 ? 's' : ''}`);
        if (globalObsData.length > 0) subtitleParts.push(`${globalObsData.length} observed set${globalObsData.length > 1 ? 's' : ''}`);
        if (overlayFlags.length > 0) subtitleParts.push(`Overlays: ${Array.from(new Set(overlayFlags)).join(', ')}`);
        if (sampledIndividuals) subtitleParts.push('Individuals sampled for performance');
        subtitleParts.push(isLogScale ? 'log scale' : 'linear scale');
        const profileSubtitle = subtitleParts.join(' • ');
        const plotTheme = getPlotTheme();

        const layout = {
            title: {
                text: `<b>Population Concentration-Time Profile</b><br><span style="font-size:11px;color:${plotTheme.subtitle};">${profileSubtitle}</span>`,
                x: 0,
                xanchor: 'left'
            },
            margin: { t: 78, r: 28, b: 94, l: 72 }, 
            hovermode: 'x unified',
            hoverdistance: 30,
            spikedistance: 1000,
            dragmode: 'pan',
            plot_bgcolor: plotTheme.bg,
            paper_bgcolor: plotTheme.paper,
            font: { family: 'Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif', color: plotTheme.font },
            hoverlabel: { font: { family: 'inherit', size: 13 }, bgcolor: plotTheme.hoverBg, bordercolor: plotTheme.hoverBorder, namelength: -1 },
            xaxis: { 
                title: { text: inputXTitle.value || 'Time (h)', font: { size: 13, color: plotTheme.title } }, 
                gridcolor: plotTheme.grid, zerolinecolor: plotTheme.zeroline,
                showline: true, linecolor: plotTheme.axisline, linewidth: 1, mirror: true, ticks: 'outside', tickcolor: plotTheme.tick, tickfont: { color: plotTheme.font },
                showspikes: true, spikecolor: plotTheme.tick, spikethickness: 1, spikemode: 'across', spikesnap: 'cursor',
                range: !isNaN(xMax) ? [0, xMax] : undefined, autorange: isNaN(xMax) 
            },
            yaxis: { 
                title: { text: inputYTitle.value || 'Concentration (ug/mL)', font: { size: 13, color: plotTheme.title } }, 
                gridcolor: plotTheme.grid, zerolinecolor: plotTheme.zeroline,
                showline: true, linecolor: plotTheme.axisline, linewidth: 1, mirror: true, ticks: 'outside', tickcolor: plotTheme.tick, tickfont: { color: plotTheme.font },
                showspikes: true, spikecolor: plotTheme.tick, spikethickness: 1, spikemode: 'across', spikesnap: 'cursor',
                type: isLogScale ? 'log' : 'linear', range: yRange, autorange: yRange === undefined 
            },
            legend: { 
                orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', 
                itemclick: 'toggle', itemdoubleclick: 'toggleothers',
                bgcolor: plotTheme.legendBg, bordercolor: plotTheme.axisline, borderwidth: 1, borderpad: 6, font: { color: plotTheme.font }
            }
        };

        Plotly.react('plotlyChart', traces, layout, {
            responsive: true,
            displaylogo: false,
            scrollZoom: true,
            modeBarButtonsToRemove: ['lasso2d', 'select2d']
        });
    }

    // -- Plotting Logic: Boxplots (Grid) --
    function renderBoxPlot(targetDiv, paramType) {
        if (!globalTrialsData) return;
        const traces = [];
        let dataFound = false;
        const paramLabel = typeof paramType === 'string' ? paramType : (paramType && paramType.label ? paramType.label : 'Parameter');
        const axisLabel = shortenAxisLabel(paramLabel, 22);

        globalTrialsData.trials.forEach((trial, index) => {
            if (!trial.active || !trial.rawParams || trial.rawParams.length === 0) return;
            
            const paramKeys = Object.keys(trial.rawParams[0]);
            const targetKey = typeof paramType === 'string'
                ? findParamKey(paramKeys, paramType)
                : findKeyForBoxParam(paramKeys, paramType);
            
            if (!targetKey) return;
            dataFound = true;

            const paramValues = [];
            trial.rawParams.forEach(row => {
                const v = parseNumericCell(row[targetKey]);
                if (!Number.isFinite(v)) return;
                if (isLogScale && v <= 0) return;
                paramValues.push(v);
            });

            if (paramValues.length === 0) return;

            const hex = chartColors[index % chartColors.length].hex;
            const rgb = chartColors[index % chartColors.length].rgb;
            let displayName = getTrialLabel(trial);
            dataFound = true;

            traces.push({
                y: paramValues,
                type: 'box',
                name: displayName,
                boxpoints: boxShowOutliers && boxShowOutliers.checked ? 'outliers' : false,
                boxmean: boxShowMean && boxShowMean.checked ? 'sd' : false, 
                marker: { color: hex, size: 4, outliercolor: 'rgba(0,0,0,0.3)', line: {outliercolor: 'rgba(0,0,0,0.3)', outlierwidth: 1} },
                line: { color: hex, width: 2 },
                fillcolor: `rgba(${rgb}, 0.15)`,
                hoverinfo: 'y+name'
            });
        });

        const targetEl = document.getElementById(targetDiv);
        if (!targetEl) return;
        if (!dataFound) {
            Plotly.purge(targetDiv);
            const emptyReason = isLogScale
                ? 'No positive numeric values available for log-scale boxplots.'
                : 'No numeric values found for this parameter.';
            targetEl.innerHTML = `<div class="flex h-full items-center justify-center text-gray-400 text-sm text-center px-3">${emptyReason}</div>`;
            return;
        }

        targetEl.innerHTML = '';
        const plotHeight = getPlotPanelCanvasHeight(targetEl);
        targetEl.style.height = `${plotHeight}px`;

        const plotTheme = getPlotTheme();
        const layout = {
            height: plotHeight,
            margin: { t: 16, r: 20, b: 64, l: 64 },
            plot_bgcolor: plotTheme.bg, paper_bgcolor: plotTheme.paper,
            font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif', color: plotTheme.font },
            xaxis: {
                showline: true,
                linecolor: plotTheme.axisline,
                linewidth: 1,
                mirror: true,
                ticks: 'outside',
                tickcolor: plotTheme.tick,
                tickfont: { color: plotTheme.font },
                automargin: true
            },
            yaxis: { 
                title: { text: axisLabel, standoff: 8, font: { size: 12, color: plotTheme.title } },
                gridcolor: plotTheme.grid, zerolinecolor: plotTheme.zeroline,
                showline: true, linecolor: plotTheme.axisline, linewidth: 1, mirror: true, ticks: 'outside', tickcolor: plotTheme.tick, tickfont: { color: plotTheme.font },
                automargin: true,
                type: isLogScale ? 'log' : 'linear' 
            },
            showlegend: false
        };

        Plotly.react(targetDiv, traces, layout, { responsive: true, displaylogo: false });
    }

    function updateBoxPlots() {
        const msgDiv = document.getElementById('boxEmptyMsg');
        const grid = document.getElementById('boxPlotsGrid');
        if (!msgDiv || !grid) return;
        const hasActiveTrials = !!(globalTrialsData && globalTrialsData.trials.some(t => t.active));
        if (!hasActiveTrials) {
            setBoxPlotLoadingState(false);
            msgDiv.classList.remove('hidden');
            const msgText = msgDiv.querySelector('p');
            if (msgText) msgText.textContent = 'No active trials selected. Enable trials in Population to render PK boxplots.';
            grid.style.display = 'none';
            grid.dataset.renderKey = '';
            return;
        }
        const hasParams = globalTrialsData && globalTrialsData.trials.some(t => t.active && t.rawParams && t.rawParams.length > 0);
        cancelBoxPlotRender();
        if (hasParams) {
            const params = getVisibleBoxplotParams();
            if (!params.length) {
                setBoxPlotLoadingState(false);
                msgDiv.classList.remove('hidden');
                const msgText = msgDiv.querySelector('p');
                if (msgText) msgText.textContent = boxResultsOnly
                    ? 'None of the selected result parameters (Cmax, AUCt, AUCinf, Fa, F, Fdp) were found in uploaded files.'
                    : "No 'Subj Params and Results' sheet found in uploaded files.";
                grid.style.display = 'none';
                grid.dataset.renderKey = '';
                return;
            }

            const renderKey = getBoxPlotRenderKey(params);
            msgDiv.classList.add('hidden');
            renderBoxPlotCards(params);
            grid.style.display = 'grid';

            if (grid.dataset.renderKey === renderKey) {
                resizeRenderedBoxPlots();
                setBoxPlotLoadingState(false);
                return;
            }

            const loadingDelay = params.length > 5 ? 120 : 220;
            setBoxPlotLoadingState(true, loadingDelay);
            const renderToken = boxPlotRenderToken;
            let paramIndex = 0;

            const renderBatch = () => {
                if (renderToken !== boxPlotRenderToken) {
                    setBoxPlotLoadingState(false);
                    return;
                }

                const batchStart = window.performance && typeof window.performance.now === 'function'
                    ? window.performance.now()
                    : Date.now();

                while (paramIndex < params.length) {
                    renderBoxPlot(`plotlyBox_${params[paramIndex].id}`, params[paramIndex]);
                    paramIndex += 1;
                    const percentDone = Math.max(0, Math.min(100, Math.round((paramIndex / params.length) * 100)));
                    if (boxPlotProgressBarFill) boxPlotProgressBarFill.style.width = `${percentDone}%`;
                    if (boxPlotProgressText) boxPlotProgressText.textContent = `${percentDone}%`;

                    const now = window.performance && typeof window.performance.now === 'function'
                        ? window.performance.now()
                        : Date.now();
                    if (now - batchStart > 24) break;
                }

                if (paramIndex < params.length) {
                    boxPlotFrameId = window.requestAnimationFrame(renderBatch);
                    return;
                }

                grid.dataset.renderKey = renderKey;
                boxPlotFrameId = null;
                resizeRenderedBoxPlots();
                setBoxPlotLoadingState(false);
            };

            boxPlotFrameId = window.requestAnimationFrame(() => {
                boxPlotFrameId = window.requestAnimationFrame(renderBatch);
            });
        } else {
            setBoxPlotLoadingState(false);
            msgDiv.classList.remove('hidden');
            const msgText = msgDiv.querySelector('p');
            if (msgText) msgText.textContent = "No 'Subj Params and Results' sheet found in uploaded files.";
            grid.style.display = 'none';
            grid.dataset.renderKey = '';
        }
    }

    // -- Plotting Logic: Bioequivalence (BE) --
    function updateBEView(forceRun = false) {
        const beContent = document.getElementById('beContent');
        const beEmptyMsg = document.getElementById('beEmptyMsg');
        const beEmptyText = beEmptyMsg ? beEmptyMsg.querySelector('p') : null;
        const BE_LOWER = 80;
        const BE_UPPER = 125;
        const beNoParamsMsg = document.getElementById('beNoParamsMsg');
        const beNoParamsText = beNoParamsMsg ? beNoParamsMsg.querySelector('p') : null;
        const beWarningMsg = document.getElementById('beWarningMsg');
        const tbody = document.getElementById('beTableBody');
        const beCompareHeader = document.getElementById('beCompareHeader');
        const bePlotCmaxEl = document.getElementById('bePlotCmax');
        const bePlotAUCinfEl = document.getElementById('bePlotAUCinf');
        const bePlotAUCtEl = document.getElementById('bePlotAUCt');
        const bePlotConclusionEl = document.getElementById('bePlotConclusion');
        if (tbody) tbody.innerHTML = '';
        const liveRefIdx = beRefTrialSelect ? String(beRefTrialSelect.value || '') : '';
        const liveTestIdx = beTestTrialSelect ? String(beTestTrialSelect.value || '') : '';
        const refIdx = liveRefIdx !== '' ? liveRefIdx : (beReferenceIndex !== '' ? String(beReferenceIndex) : '');
        let testIdx = liveTestIdx !== '' ? liveTestIdx : (beTestIndex !== '' ? String(beTestIndex) : '');
        const beMethod = beMethodSelect ? beMethodSelect.value : 'simple';
        const isPairedTrialNumberMode = beMethod === 'paired-trial-number';

        beReferenceIndex = refIdx;
        if (isPairedTrialNumberMode) {
            beTestIndex = '';
            testIdx = '';
            if (beTestTrialSelect) beTestTrialSelect.value = '';
        } else {
            beTestIndex = testIdx;
        }

        if (beRefTrialSelect) beRefTrialSelect.disabled = isPairedTrialNumberMode;
        if (beTestTrialSelect) beTestTrialSelect.disabled = isPairedTrialNumberMode;
        if (beCompareHeader) beCompareHeader.textContent = isPairedTrialNumberMode ? 'Ref/Test Pair' : 'Test Trial';
        if (beMethodHint) {
            beMethodHint.textContent = isPairedTrialNumberMode
                ? 'Automatically compares active Reference and Test trials with the same trial number.'
                : 'All active non-reference trials are compared against the reference.';
        }
        if (bePairingPanel) bePairingPanel.classList.add('hidden');
        if (bePairingSummary) bePairingSummary.textContent = '';
        if (bePairingDetail) bePairingDetail.textContent = '';
        
        // globalBEData = []; // Clear previous exports - MOVED to runBEAnalysis to preserve results when switching views or methods
        const isRerunPromptState = beNeedsRerun && !hasReviewedResults && !forceRun;
        if (isRerunPromptState) {
            beWarningMsg.textContent = 'BE method changed. Click "Run BE" to regenerate charts and table.';
            beWarningMsg.classList.remove('hidden');
        } else {
            beWarningMsg.classList.add('hidden');
            beWarningMsg.textContent = '';
        }

        if (!hasReviewedResults && !forceRun) {
            beContent.classList.add('hidden');
            beEmptyMsg.classList.remove('hidden');
            beNoParamsMsg.classList.add('hidden');
            if (beNeedsRerun && beEmptyText) {
                beEmptyText.textContent = 'Method changed. Click "Run BE" to regenerate results.';
            }
            if (beRefSummary && globalTrialsData && globalTrialsData.trials) {
                const selRef = beRefTrialSelect ? String(beRefTrialSelect.value || '') : String(beReferenceIndex || '');
                if (selRef !== '' && globalTrialsData.trials[selRef]) {
                    beRefSummary.textContent = `Reference: ${getTrialLabel(globalTrialsData.trials[selRef])}`;
                } else {
                    beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
                }
            } else if (beRefSummary) {
                beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
            }
            return;
        }

        if (forceRun) {
            hasReviewedResults = true;
            beNeedsRerun = false;
        }

        if (!globalTrialsData) {
            beContent.classList.add('hidden');
            beEmptyMsg.classList.remove('hidden');
            beNoParamsMsg.classList.add('hidden');
            if (beRefSummary) beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
            return;
        }

        if (!isPairedTrialNumberMode && (refIdx === "" || !globalTrialsData.trials[refIdx])) {
            const fallbackRefIdx = globalTrialsData.trials.findIndex(t => t.active && t.formulationType === 'reference' && t.rawParams && t.rawParams.length > 0);
            if (fallbackRefIdx >= 0) {
                beReferenceIndex = String(fallbackRefIdx);
                if (beRefTrialSelect) beRefTrialSelect.value = beReferenceIndex;
            }
        }

        const resolvedRefIdx = beReferenceIndex !== '' ? String(beReferenceIndex) : '';
        if (!isPairedTrialNumberMode && (resolvedRefIdx === "" || !globalTrialsData.trials[resolvedRefIdx])) {
            beContent.classList.add('hidden');
            beEmptyMsg.classList.remove('hidden');
            beNoParamsMsg.classList.add('hidden');
            if (beRefSummary) beRefSummary.textContent = 'Select a Reference formulation trial for BE.';
            return;
        }

        if (!isPairedTrialNumberMode && testIdx !== '') {
            const selectedTestTrial = globalTrialsData.trials[Number(testIdx)];
            const selectedTestStillValid = !!selectedTestTrial
                && selectedTestTrial.formulationType === 'test'
                && !!selectedTestTrial.active
                && selectedTestTrial.rawParams
                && selectedTestTrial.rawParams.length > 0;
            if (!selectedTestStillValid) {
                beTestIndex = '';
                if (beTestTrialSelect) beTestTrialSelect.value = '';
                testIdx = '';
            }
        }
        
        const refTrial = isPairedTrialNumberMode ? null : globalTrialsData.trials[resolvedRefIdx];
        let comparisonPairs = [];
        let pairingDetailText = '';

        if (isPairedTrialNumberMode) {
            const activeRefTrials = globalTrialsData.trials.filter(t => t.active && t.formulationType === 'reference' && t.rawParams && t.rawParams.length > 0);
            const activeTestTrials = globalTrialsData.trials.filter(t => t.active && t.formulationType === 'test' && t.rawParams && t.rawParams.length > 0);
            const refByNumber = new Map();
            const testByNumber = new Map();

            activeRefTrials.forEach(t => {
                const n = Number(t.trialNumber);
                if (!Number.isFinite(n)) return;
                if (!refByNumber.has(n)) refByNumber.set(n, []);
                refByNumber.get(n).push(t);
            });
            activeTestTrials.forEach(t => {
                const n = Number(t.trialNumber);
                if (!Number.isFinite(n)) return;
                if (!testByNumber.has(n)) testByNumber.set(n, []);
                testByNumber.get(n).push(t);
            });

            const duplicateRef = Array.from(refByNumber.entries()).filter(([, rows]) => rows.length > 1).map(([n]) => n).sort((a, b) => a - b);
            const duplicateTest = Array.from(testByNumber.entries()).filter(([, rows]) => rows.length > 1).map(([n]) => n).sort((a, b) => a - b);
            const ambiguous = new Set([...duplicateRef, ...duplicateTest]);

            const matchedNumbers = Array.from(refByNumber.keys())
                .filter(n => testByNumber.has(n) && !ambiguous.has(n))
                .sort((a, b) => a - b);

            comparisonPairs = matchedNumbers.map(n => ({
                trialNumber: n,
                refTrial: refByNumber.get(n)[0],
                testTrial: testByNumber.get(n)[0],
                label: `Trial ${n}`
            }));

            const unmatchedRef = Array.from(refByNumber.keys()).filter(n => !testByNumber.has(n) && !ambiguous.has(n)).sort((a, b) => a - b);
            const unmatchedTest = Array.from(testByNumber.keys()).filter(n => !refByNumber.has(n) && !ambiguous.has(n)).sort((a, b) => a - b);
            const detailParts = [];
            if (unmatchedRef.length) detailParts.push(`Reference-only trial numbers: ${unmatchedRef.join(', ')}`);
            if (unmatchedTest.length) detailParts.push(`Test-only trial numbers: ${unmatchedTest.join(', ')}`);
            if (duplicateRef.length) detailParts.push(`Ambiguous Reference duplicates (excluded): ${duplicateRef.join(', ')}`);
            if (duplicateTest.length) detailParts.push(`Ambiguous Test duplicates (excluded): ${duplicateTest.join(', ')}`);
            pairingDetailText = detailParts.length ? detailParts.join(' | ') : 'All active Reference and Test trial numbers are matched.';

            if ((duplicateRef.length || duplicateTest.length) && beWarningMsg) {
                const warningParts = [];
                if (duplicateRef.length) warningParts.push(`duplicate Reference trial numbers: ${duplicateRef.join(', ')}`);
                if (duplicateTest.length) warningParts.push(`duplicate Test trial numbers: ${duplicateTest.join(', ')}`);
                beWarningMsg.textContent = `Paired mode excluded ambiguous matches due to ${warningParts.join(' and ')}.`;
                beWarningMsg.classList.remove('hidden');
            }

            if (beRefSummary) {
                beRefSummary.textContent = comparisonPairs.length
                    ? `Matched pairs: ${comparisonPairs.length} (Ref n vs Test n)`
                    : 'No matched Reference/Test trial numbers among active trials.';
            }
        } else {
            if (beRefSummary && refTrial) {
                beRefSummary.textContent = `Reference: ${getTrialLabel(refTrial)}`;
            }
            const testTrials = globalTrialsData.trials.filter((t, i) => {
                const isTestFormulation = t.formulationType === 'test';
                const isActive = !!t.active;
                const selectedTestMatch = testIdx === '' ? true : String(i) === testIdx;
                return isTestFormulation && isActive && selectedTestMatch;
            });
            comparisonPairs = testTrials.map(t => ({ refTrial, testTrial: t, label: getTrialLabel(t) }));
        }

        if (comparisonPairs.length === 0) {
            beContent.classList.add('hidden');
            beEmptyMsg.classList.add('hidden');
            beNoParamsMsg.classList.remove('hidden');
            if (beNoParamsText) {
                beNoParamsText.textContent = isPairedTrialNumberMode
                    ? 'No matched active trial numbers found between Reference and Test trials. Keep both formulations active for the same trial numbers.'
                    : 'No active Test formulation trials found with valid PK parameters to compute BE.';
            }
            return;
        }

        if (!isPairedTrialNumberMode && (!refTrial.rawParams || refTrial.rawParams.length === 0)) {
            beContent.classList.add('hidden');
            beEmptyMsg.classList.add('hidden');
            beNoParamsMsg.classList.remove('hidden');
            if (beNoParamsText) beNoParamsText.textContent = 'Selected Reference trial has no valid PK parameter rows for BE computation.';
            return;
        }
        
        const parameters = ['Cmax', 'AUCt', 'AUCinf'];
        let hasValidData = false;
        let excludedCount = 0;
        const beExportRows = [];
        
        const bePlotRows = [];
        const beTableRows = [];
        const beXValues = [];
        
        comparisonPairs.forEach((pair, pairIdx) => {
            const refTrialInPair = pair.refTrial;
            const testTrial = pair.testTrial;
            const hex = chartColors[globalTrialsData.trials.indexOf(testTrial) % chartColors.length].hex;
            
            const xs = [];
            const ys = [];
            const errorMinus = [];
            const errorPlus = [];
            const hoverTexts = [];
            const markerColors = [];
            const markerSymbols = [];
            
            parameters.slice().reverse().forEach(param => {
                const refDiag = extractParamDataWithDiagnostics(refTrialInPair, param);
                const testDiag = extractParamDataWithDiagnostics(testTrial, param);
                const refVals = refDiag.values;
                const testVals = testDiag.values;
                excludedCount += refDiag.excluded + testDiag.excluded;
                
                if (refVals.length > 1 && testVals.length > 1) {
                    const refLogs = refVals.map(v => Math.log(v));
                    const testLogs = testVals.map(v => Math.log(v));
                    
                    const beStats = calculateBE(refLogs, testLogs);
                    if(beStats) {
                        hasValidData = true;
                        const isBE = beStats.lower >= BE_LOWER && beStats.upper <= BE_UPPER;
                        
                        xs.push(beStats.pe);
                        ys.push(`${param}`);
                        errorMinus.push(beStats.pe - beStats.lower);
                        errorPlus.push(beStats.upper - beStats.pe);
                        beXValues.push(beStats.lower, beStats.pe, beStats.upper);
                        hoverTexts.push(`PE: ${beStats.pe.toFixed(2)}%<br>90% CI: [${beStats.lower.toFixed(2)}%, ${beStats.upper.toFixed(2)}%]<br>Status: ${isBE ? 'Pass' : 'Fail'}`);
                        markerColors.push(isBE ? hex : '#dc2626');
                        markerSymbols.push(isBE ? 'square' : 'x-thin-open');

                        const rowLabel = isPairedTrialNumberMode
                            ? `${param} - Trial ${pair.trialNumber}`
                            : `${param} - ${getTrialLabel(testTrial)}`;
                        const trialNumRaw = isPairedTrialNumberMode ? pair.trialNumber : testTrial.trialNumber;
                        const trialNumber = Number.isFinite(Number(trialNumRaw)) ? Number(trialNumRaw) : (pairIdx + 1);
                        bePlotRows.push({
                            param,
                            rowLabel,
                            trialNumber,
                            pe: beStats.pe,
                            lower: beStats.lower,
                            upper: beStats.upper,
                            isBE,
                            trialLabel: getTrialLabel(testTrial),
                            refLabel: getTrialLabel(refTrialInPair),
                            pairLabel: isPairedTrialNumberMode ? `Trial ${pair.trialNumber}` : getTrialLabel(testTrial)
                        });
                        
                        // Push to export data
                        beExportRows.push({
                            param: param,
                            testTrial: getTrialLabel(testTrial),
                            refTrial: getTrialLabel(refTrialInPair),
                            pair: isPairedTrialNumberMode ? `Trial ${pair.trialNumber}` : '',
                            pe: beStats.pe.toFixed(2),
                            lower: beStats.lower.toFixed(2),
                            upper: beStats.upper.toFixed(2),
                            status: isBE ? 'Pass' : 'Fail'
                        });

                        const trialCellLabel = isPairedTrialNumberMode
                            ? `${escapeHtml(getTrialLabel(testTrial))} vs ${escapeHtml(getTrialLabel(refTrialInPair))}`
                            : `${escapeHtml(getTrialLabel(testTrial))}`;

                        beTableRows.push({
                            param,
                            trialNumber,
                            trialCellLabel,
                            peText: beStats.pe.toFixed(2),
                            lowerText: beStats.lower.toFixed(2),
                            upperText: beStats.upper.toFixed(2),
                            isLowerFail: beStats.lower < BE_LOWER,
                            isUpperFail: beStats.upper > BE_UPPER,
                            isBE,
                            hex
                        });
                    }
                }
            });
            
            if(xs.length > 0) {
                // Keep stats vectors populated for table/export side effects in this loop.
            }
        });

        if(!hasValidData) {
            globalBEData = [];
            beContent.classList.add('hidden');
            beEmptyMsg.classList.add('hidden');
            beNoParamsMsg.classList.remove('hidden');
            if (beNoParamsText) beNoParamsText.textContent = 'No valid individual PK parameters found in the selected trials to compute BE.';
            return;
        }

        globalBEData = beExportRows;

        beContent.classList.remove('hidden');
        beEmptyMsg.classList.add('hidden');
        beNoParamsMsg.classList.add('hidden');

        const paramOrder = { Cmax: 0, AUCt: 1, AUCinf: 2 };
        const orderedTableRows = beTableRows.slice().sort((a, b) => {
            const byParam = (paramOrder[a.param] ?? 99) - (paramOrder[b.param] ?? 99);
            if (byParam !== 0) return byParam;
            if (a.trialNumber !== b.trialNumber) return a.trialNumber - b.trialNumber;
            return a.trialCellLabel.localeCompare(b.trialCellLabel);
        });

        orderedTableRows.forEach(row => {
            const tr = document.createElement('tr');
            tr.className = `transition-colors ${row.isBE ? '' : 'be-fail-highlight'}`;
            tr.style.borderBottom = '1px solid rgba(255,255,255,0.06)';
            tr.innerHTML = `
                <td class="px-5 py-2.5 whitespace-nowrap text-sm font-semibold" style="color:var(--text-primary)">${row.param}</td>
                <td class="px-5 py-2.5 whitespace-nowrap text-sm" style="color:var(--text-secondary)">
                    <span class="inline-block w-2 h-2 rounded-full mr-1.5" style="background-color: ${row.hex};"></span>
                    ${row.trialCellLabel}
                </td>
                <td class="px-5 py-2.5 whitespace-nowrap text-sm text-right font-bold" style="color:var(--text-primary)">${row.peText}</td>
                <td class="px-5 py-2.5 whitespace-nowrap text-sm text-right" style="color:${row.isLowerFail ? '#f87171' : 'var(--text-secondary)'}; font-weight:${row.isLowerFail ? '700' : '400'}">${row.lowerText}</td>
                <td class="px-5 py-2.5 whitespace-nowrap text-sm text-right" style="color:${row.isUpperFail ? '#f87171' : 'var(--text-secondary)'}; font-weight:${row.isUpperFail ? '700' : '400'}">${row.upperText}</td>
                <td class="px-5 py-2.5 whitespace-nowrap text-center">
                    ${row.isBE
                        ? '<span style="background:rgba(74,222,128,0.1);color:#4ade80;border:1px solid rgba(74,222,128,0.25);padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;">Pass</span>'
                        : '<span style="background:rgba(248,113,113,0.1);color:#f87171;border:1px solid rgba(248,113,113,0.25);padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;">Fail</span>'}
                </td>
            `;
            tbody.appendChild(tr);
        });

        if (isPairedTrialNumberMode && bePairingPanel) {
            bePairingPanel.classList.remove('hidden');
            if (bePairingSummary) bePairingSummary.textContent = `Matched Trial Number Analysis: ${comparisonPairs.length} pair(s)`;
            if (bePairingDetail) bePairingDetail.textContent = pairingDetailText;
        }

        if (excludedCount > 0) {
            beWarningMsg.textContent = `${excludedCount} non-positive PK values were excluded from log-transformed BE calculations.`;
            beWarningMsg.classList.remove('hidden');
        }

        const plotTheme = getPlotTheme();
        const paramColorMap = { Cmax: '#60a5fa', AUCinf: '#f59e0b', AUCt: '#34d399' };
        let numericRows = bePlotRows.filter(r => Number.isFinite(r.trialNumber));
        if (!numericRows.length && bePlotRows.length) {
            // Fallback numbering keeps charts renderable even when trial numbers are missing or malformed.
            numericRows = bePlotRows.map((r, i) => ({ ...r, trialNumber: i + 1 }));
        }
        const uniqueTrials = Array.from(new Set(numericRows.map(r => r.trialNumber))).sort((a, b) => a - b);

        const renderParamChart = (targetId, param, title) => {
            const target = document.getElementById(targetId);
            const rows = numericRows
                .filter(r => r.param === param)
                .sort((a, b) => a.trialNumber - b.trialNumber);
            if (!target) return;
            if (!rows.length) {
                Plotly.purge(targetId);
                target.innerHTML = '<div class="h-full flex items-center justify-center text-sm" style="color: var(--text-tertiary);">No data</div>';
                return;
            }

            target.innerHTML = '';
            const minTrial = rows[0].trialNumber;
            const maxTrial = rows[rows.length - 1].trialNumber;
            const trace = {
                x: rows.map(r => r.trialNumber),
                y: rows.map(r => r.pe),
                type: 'scatter',
                mode: 'markers+lines',
                marker: {
                    size: 10,
                    color: rows.map(r => r.isBE ? paramColorMap[param] : '#ef4444'),
                    symbol: rows.map(r => r.isBE ? 'circle' : 'x-thin-open'),
                    line: { color: 'white', width: 1 }
                },
                line: { color: paramColorMap[param], width: 1.5, shape: 'linear' },
                error_y: {
                    type: 'data',
                    symmetric: false,
                    array: rows.map(r => r.upper - r.pe),
                    arrayminus: rows.map(r => r.pe - r.lower),
                    thickness: 1.8,
                    width: 6,
                    color: paramColorMap[param]
                },
                text: rows.map(r => `${r.pairLabel}<br>PE: ${r.pe.toFixed(2)}%<br>90% CI: [${r.lower.toFixed(2)}%, ${r.upper.toFixed(2)}%]<br>Status: ${r.isBE ? 'Pass' : 'Fail'}`),
                hovertemplate: `Trial %{x}<br>%{text}<extra>${title}</extra>`
            };

            const layout = {
                margin: { t: 8, r: 12, b: 38, l: 50 },
                plot_bgcolor: plotTheme.bg,
                paper_bgcolor: plotTheme.paper,
                font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif', color: plotTheme.font },
                xaxis: {
                    title: { text: 'Trial Number', font: { size: 11, color: plotTheme.subtitle } },
                    showline: true,
                    linecolor: plotTheme.axisline,
                    tickmode: 'array',
                    tickvals: uniqueTrials,
                    ticktext: uniqueTrials.map(v => String(v)),
                    range: [minTrial - 0.5, maxTrial + 0.5],
                    gridcolor: plotTheme.grid,
                    zerolinecolor: plotTheme.zeroline,
                    tickfont: { size: 10, color: plotTheme.font }
                },
                yaxis: {
                    title: { text: 'BE Range (%)', font: { size: 11, color: plotTheme.subtitle } },
                    range: [40, 160],
                    gridcolor: plotTheme.grid,
                    zerolinecolor: plotTheme.zeroline,
                    tickfont: { size: 10, color: plotTheme.font }
                },
                shapes: [
                    { type: 'rect', xref: 'x', yref: 'y', x0: minTrial - 0.5, x1: maxTrial + 0.5, y0: BE_LOWER, y1: BE_UPPER, fillcolor: 'rgba(212,149,106,0.08)', line: { width: 0 }, layer: 'below' },
                    { type: 'line', xref: 'x', yref: 'y', x0: minTrial - 0.5, x1: maxTrial + 0.5, y0: BE_LOWER, y1: BE_LOWER, line: { color: '#d4956a', width: 1.2, dash: 'dot' }, layer: 'below' },
                    { type: 'line', xref: 'x', yref: 'y', x0: minTrial - 0.5, x1: maxTrial + 0.5, y0: BE_UPPER, y1: BE_UPPER, line: { color: '#d4956a', width: 1.2, dash: 'dot' }, layer: 'below' },
                    { type: 'line', xref: 'x', yref: 'y', x0: minTrial - 0.5, x1: maxTrial + 0.5, y0: 100, y1: 100, line: { color: plotTheme.tick, width: 1, dash: 'dash' }, layer: 'below' }
                ],
                annotations: [
                    {
                        x: maxTrial,
                        y: BE_UPPER + 4,
                        xref: 'x',
                        yref: 'y',
                        text: `${BE_LOWER}-${BE_UPPER}% window`,
                        showarrow: false,
                        font: { size: 9, color: '#d4956a' }
                    }
                ],
                showlegend: false
            };

            Plotly.react(targetId, [trace], layout, { responsive: true, displaylogo: false });
        };

        const renderConclusionChart = () => {
            const targetId = 'bePlotConclusion';
            const target = document.getElementById(targetId);
            if (!target) return;

            const trialMap = new Map();
            numericRows.forEach(r => {
                if (!trialMap.has(r.trialNumber)) {
                    trialMap.set(r.trialNumber, { pass: 0, fail: 0 });
                }
                const bucket = trialMap.get(r.trialNumber);
                if (r.isBE) bucket.pass += 1;
                else bucket.fail += 1;
            });

            const rows = Array.from(trialMap.entries())
                .map(([trialNumber, v]) => ({ trialNumber, pass: v.pass, fail: v.fail, overallPass: v.fail === 0 && v.pass > 0 }))
                .sort((a, b) => a.trialNumber - b.trialNumber);

            if (!rows.length) {
                Plotly.purge(targetId);
                target.innerHTML = '<div class="h-full flex items-center justify-center text-sm" style="color: var(--text-tertiary);">No data</div>';
                return;
            }

            target.innerHTML = '';
            const minTrial = rows[0].trialNumber;
            const maxTrial = rows[rows.length - 1].trialNumber;
            const trace = {
                x: rows.map(r => r.trialNumber),
                y: rows.map(r => r.overallPass ? 1 : 0),
                type: 'bar',
                marker: {
                    color: rows.map(r => r.overallPass ? '#22c55e' : '#ef4444')
                },
                text: rows.map(r => `${r.overallPass ? 'Pass' : 'Fail'} (${r.pass}/${r.pass + r.fail} params)`),
                textposition: 'outside',
                hovertemplate: 'Trial %{x}<br>%{text}<extra>Conclusion</extra>'
            };

            const layout = {
                margin: { t: 8, r: 12, b: 38, l: 50 },
                plot_bgcolor: plotTheme.bg,
                paper_bgcolor: plotTheme.paper,
                font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif', color: plotTheme.font },
                xaxis: {
                    title: { text: 'Trial Number', font: { size: 11, color: plotTheme.subtitle } },
                    tickmode: 'array',
                    tickvals: rows.map(r => r.trialNumber),
                    ticktext: rows.map(r => String(r.trialNumber)),
                    range: [minTrial - 0.5, maxTrial + 0.5],
                    gridcolor: plotTheme.grid,
                    zerolinecolor: plotTheme.zeroline,
                    tickfont: { size: 10, color: plotTheme.font }
                },
                yaxis: {
                    title: { text: 'Overall BE', font: { size: 11, color: plotTheme.subtitle } },
                    tickmode: 'array',
                    tickvals: [0, 1],
                    ticktext: ['Fail', 'Pass'],
                    range: [0, 1.2],
                    gridcolor: plotTheme.grid,
                    zerolinecolor: plotTheme.zeroline,
                    tickfont: { size: 10, color: plotTheme.font }
                },
                showlegend: false
            };

            Plotly.react(targetId, [trace], layout, { responsive: true, displaylogo: false });
        };

        try {
            renderParamChart('bePlotCmax', 'Cmax', 'Cmax');
            renderParamChart('bePlotAUCinf', 'AUCinf', 'AUCinf');
            renderParamChart('bePlotAUCt', 'AUCt', 'AUCt');
            renderConclusionChart();
        } catch (plotErr) {
            console.error('BE chart render failed:', plotErr);
            const fallback = (el) => {
                if (!el) return;
                el.innerHTML = '<div class="h-full flex items-center justify-center text-sm" style="color: var(--text-tertiary);">Chart render error</div>';
            };
            fallback(bePlotCmaxEl);
            fallback(bePlotAUCinfEl);
            fallback(bePlotAUCtEl);
            fallback(bePlotConclusionEl);
            showStatusToast('BE chart rendering failed for this selection; table results are still available.', 'warn');
        }
    }

    // -- Rendering Logic: Summary Stats Table --
    function updateStatsTable() {
        if (!globalTrialsData) return;
        
        const tbody = document.getElementById('statsTableBody');
        const compareBody = document.getElementById('statsCompareBody');
        const comparePanel = document.getElementById('statsComparePanel');
        const msgDiv = document.getElementById('statsEmptyMsg');
        const table = document.getElementById('statsTable');
        
        tbody.innerHTML = '';
        compareBody.innerHTML = '';
        let dataFound = false;

        const endpointMap = new Map();

        globalTrialsData.trials.forEach((trial, index) => {
            if (!trial.active || !trial.rawStats || trial.rawStats.length === 0) return;
            dataFound = true;
            const hex = chartColors[index % chartColors.length].hex;

            const trialHeader = document.createElement('tr');
            trialHeader.style.cssText = 'background: var(--bg-elevated); border-top: 1px solid rgba(255,255,255,0.1); border-bottom: 1px solid rgba(255,255,255,0.06);';
            trialHeader.innerHTML = `
                <td colspan="8" class="px-5 py-2.5 whitespace-nowrap text-xs font-bold" style="color:var(--text-primary)">
                    <span class="inline-block w-2.5 h-2.5 rounded-full mr-2" style="background-color: ${hex}; box-shadow: 0 0 6px ${hex}55;"></span>
                    ${escapeHtml(getTrialLabel(trial))}
                    <span class="ml-2 px-2 py-0.5 rounded-full text-[10px] font-semibold" style="background:var(--bg-base);color:var(--text-tertiary);border:1px solid rgba(255,255,255,0.08)">${trial.subjectCount} subjects</span>
                </td>
            `;
            tbody.appendChild(trialHeader);

            trial.rawStats.forEach(row => {
                const ext = extractStatsRow(row);
                
                if (ext.epName === '-' || (ext.mean === '-' && ext.cv === '-')) return;
                if (typeof ext.epName === 'string' && ext.epName.trim() === '') return;

                const meanNum = Number(ext.mean);
                if (Number.isFinite(meanNum)) {
                    const key = String(ext.epName).toLowerCase().replace(/\s+/g, ' ').trim();
                    if (!endpointMap.has(key)) endpointMap.set(key, { label: String(ext.epName), values: [] });
                    endpointMap.get(key).values.push({ trialNumber: trial.trialNumber, mean: meanNum });
                }

                const formatNum = (val, dec) => typeof val === 'number' ? val.toFixed(dec) : val;

                const tr = document.createElement('tr');
                tr.style.borderBottom = '1px solid rgba(255,255,255,0.05)';
                tr.innerHTML = `
                    <td class="px-5 py-2 whitespace-nowrap text-sm font-medium pl-9" style="color:var(--text-primary)">${ext.epName}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-primary);font-variant-numeric:tabular-nums">${formatNum(ext.mean, 3)}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-secondary);font-variant-numeric:tabular-nums">${formatNum(ext.cv, 1)}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-tertiary);font-variant-numeric:tabular-nums">${formatNum(ext.min, 3)}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-tertiary);font-variant-numeric:tabular-nums">${formatNum(ext.max, 3)}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right font-semibold" style="color:var(--accent);font-variant-numeric:tabular-nums">${formatNum(ext.geom, 3)}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-secondary);font-variant-numeric:tabular-nums">${ext.ci90}</td>
                    <td class="px-5 py-2 whitespace-nowrap text-sm text-right" style="color:var(--text-secondary);font-variant-numeric:tabular-nums">${ext.ci90ln}</td>
                `;
                tbody.appendChild(tr);
            });
        });

        const activeTrialsWithStats = globalTrialsData.trials.filter(t => t.active && t.rawStats && t.rawStats.length > 0);
        if (activeTrialsWithStats.length > 1) {
            const rows = [];
            endpointMap.forEach((entry) => {
                if (entry.values.length < 2) return;
                const sorted = entry.values.slice().sort((a, b) => a.mean - b.mean);
                const min = sorted[0];
                const max = sorted[sorted.length - 1];
                const spreadPct = min.mean > 0 ? ((max.mean - min.mean) / min.mean) * 100 : null;
                rows.push({
                    label: entry.label,
                    trialCount: sorted.length,
                    min,
                    max,
                    spreadPct
                });
            });

            rows.sort((a, b) => (b.spreadPct || -1) - (a.spreadPct || -1));

            rows.slice(0, 12).forEach(r => {
                const tr = document.createElement('tr');
                tr.style.borderBottom = '1px solid rgba(255,255,255,0.05)';
                const spreadColor = r.spreadPct != null && r.spreadPct > 20 ? '#fbbf24' : 'var(--text-secondary)';
                const spreadWeight = r.spreadPct != null && r.spreadPct > 20 ? '700' : '400';
                tr.innerHTML = `
                    <td class="px-4 py-2 text-sm font-semibold" style="color:var(--text-primary)">${r.label}</td>
                    <td class="px-4 py-2 text-sm text-right" style="color:var(--text-secondary)">${r.trialCount}</td>
                    <td class="px-4 py-2 text-sm text-right" style="color:var(--text-secondary)">${r.min.mean.toFixed(3)}</td>
                    <td class="px-4 py-2 text-sm text-right font-semibold" style="color:var(--text-primary)">${r.max.mean.toFixed(3)}</td>
                    <td class="px-4 py-2 text-sm text-right" style="color:${spreadColor};font-weight:${spreadWeight}">${r.spreadPct != null ? r.spreadPct.toFixed(1) + '%' : '-'}</td>
                    <td class="px-4 py-2 text-sm" style="color:var(--text-secondary)">Trial ${r.max.trialNumber} / Trial ${r.min.trialNumber}</td>
                `;
                compareBody.appendChild(tr);
            });

            if (rows.length > 0) comparePanel.classList.remove('hidden');
            else comparePanel.classList.add('hidden');
        } else {
            comparePanel.classList.add('hidden');
        }

        const hasActiveTrials = globalTrialsData.trials.some(t => t.active);
        if (!dataFound && hasActiveTrials) {
            msgDiv.classList.remove('hidden');
            const msgText = msgDiv.querySelector('p');
            if (msgText) msgText.textContent = "No 'Summary Stats' sheet found in uploaded active trials.";
            table.parentElement.classList.add('hidden');
        } else if (!hasActiveTrials) {
            msgDiv.classList.remove('hidden');
            const msgText = msgDiv.querySelector('p');
            if (msgText) msgText.textContent = 'No active trials selected. Enable at least one trial in Population to view summary statistics.';
            table.parentElement.classList.add('hidden');
        } else {
            msgDiv.classList.add('hidden');
            table.parentElement.classList.remove('hidden');
        }
    }

    // -- CSV Data Export --
    function buildExportMetadata(scopeLabel) {
        const now = new Date().toISOString();
        const activeTrials = globalTrialsData ? globalTrialsData.trials.filter(t => t.active).map(t => t.trialNumber).join('|') : '';
        return [
            `# Scope,${scopeLabel}`,
            `# ExportedAt,${now}`,
            `# ActiveTab,${currentTab}`,
            `# LogScale,${isLogScale}`,
            `# ActiveTrials,${activeTrials}`,
            `# ObsFileCount,${obsFileCount}`,
            ''
        ].join('\n');
    }

    function escapeCSVValue(value) {
        const str = value === null || value === undefined ? '' : String(value);
        if (!/[",\n\r]/.test(str)) return str;
        return `"${str.replace(/"/g, '""')}"`;
    }

    function toCSVRow(values) {
        return values.map(escapeCSVValue).join(',');
    }

    function getStatsCSVContent() {
        if (!globalTrialsData) return null;
        const header = ['Trial_Number', 'Endpoint', 'Mean', 'CV_Percent', 'Min', 'Max', 'Geom_Mean', '90_CI_Arith', '90_CI_Ln'];
        const rows = [];
        
        globalTrialsData.trials.forEach(trial => {
            if (!trial.active || !trial.rawStats || trial.rawStats.length === 0) return;
            
            trial.rawStats.forEach(row => {
                const ext = extractStatsRow(row);
                
                if (ext.epName === '-' || (ext.mean === '-' && ext.cv === '-')) return;
                if (typeof ext.epName === 'string' && ext.epName.trim() === '') return;

                rows.push([
                    trial.trialNumber,
                    ext.epName,
                    ext.mean,
                    ext.cv,
                    ext.min,
                    ext.max,
                    ext.geom,
                    ext.ci90,
                    ext.ci90ln
                ]);
            });
        });

        return `${buildExportMetadata('Summary Stats')}\n${toCSVRow(header)}\n${rows.map(toCSVRow).join('\n')}${rows.length ? '\n' : ''}`;
    }

    function exportStatsToCSV(filename = "GastroPlus_Summary_Stats.csv") {
        const csvContent = getStatsCSVContent();
        if (!csvContent) return;
        downloadCSV(csvContent, filename);
    }
    
    function getParamsCSVContent() {
        if (!globalTrialsData) return null;
        const header = ['Trial_Number', 'Subject', 'Parameter', 'Value'];
        const rows = [];
        
        globalTrialsData.trials.forEach(trial => {
            if (!trial.active || !trial.rawParams || trial.rawParams.length === 0) return;
            
            const paramKeys = Object.keys(trial.rawParams[0]);
            const targets = ['cmax', 'auc'];
            const validKeys = paramKeys.filter(k => k && targets.some(t => k.toLowerCase().includes(t)));

            trial.rawParams.forEach((row, i) => {
                validKeys.forEach(k => {
                    const parsed = parseNumericCell(row[k]);
                    if (!Number.isFinite(parsed)) return;
                    rows.push([trial.trialNumber, `Subject_${i + 1}`, k, parsed]);
                });
            });
        });

        return `${buildExportMetadata('PK Parameters')}\n${toCSVRow(header)}\n${rows.map(toCSVRow).join('\n')}${rows.length ? '\n' : ''}`;
    }

    function exportParamsToCSV(filename = "GastroPlus_PK_Parameters.csv") {
        const csvContent = getParamsCSVContent();
        if (!csvContent) return;
        downloadCSV(csvContent, filename);
    }

    function getBECSVContent() {
        if (globalBEData.length === 0) return null;
        const header = ['Parameter', 'Test_Trial', 'Reference_Trial', 'Point_Estimate_Percent', 'Lower_90_CI', 'Upper_90_CI', 'Status'];
        const rows = globalBEData.map(row => [row.param, row.testTrial, row.refTrial, row.pe, row.lower, row.upper, row.status]);
        return `${buildExportMetadata('Bioequivalence')}\n${toCSVRow(header)}\n${rows.map(toCSVRow).join('\n')}\n`;
    }

    function exportBEDataToCSV(filename = "GastroPlus_BE_Results.csv") {
        const csvContent = getBECSVContent();
        if (!csvContent) return;
        downloadCSV(csvContent, filename);
    }

    function getProfileCSVContent() {
        if (!globalTrialsData) return null;

        globalTrialsData.trials.forEach(trial => {
            if (!trial.active) return;
            const hasStats = trial.stats && Array.isArray(trial.stats.means) && trial.stats.means.length === trial.times.length;
            if (!hasStats) computeTrialProfileStats(trial);
        });
        
        const header = ['Trial_Number', 'Time', 'Mean', 'Median', 'Min', '2.5th_Percentile', '5th_Percentile', '25th_Percentile', '75th_Percentile', '95th_Percentile', '97.5th_Percentile', 'Max', 'Lower_90CI_Mean', 'Upper_90CI_Mean'];
        const rows = [];

        globalTrialsData.trials.forEach(trial => {
            if (!trial.active || !trial.stats.means || !trial.stats.means.length) return;
            
            trial.times.forEach((t, i) => {
                rows.push([
                    trial.trialNumber, t,
                    trial.stats.means[i], trial.stats.medians[i],
                    trial.stats.p00[i], trial.stats.p025[i], trial.stats.p05[i], 
                    trial.stats.p25[i], trial.stats.p75[i], 
                    trial.stats.p95[i], trial.stats.p975[i], trial.stats.p100[i],
                    trial.stats.lowerCI[i], trial.stats.upperCI[i]
                ]);
            });
        });

        return `${buildExportMetadata('Aggregated Profile')}\n${toCSVRow(header)}\n${rows.map(toCSVRow).join('\n')}${rows.length ? '\n' : ''}`;
    }

    function exportDataToCSV(filename = "GastroPlus_Aggregated_Profiles.csv") {
        const csvContent = getProfileCSVContent();
        if (!csvContent) return;
        downloadCSV(csvContent, filename);
    }

    function downloadCSV(csvContent, filename) {
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    function showLoading(isVisible) {
        if (isVisible) {
            loadingSpinner.classList.remove('hidden');
            if (loadingOverlay) {
                loadingOverlay.classList.remove('hidden');
                loadingOverlay.classList.add('flex');
            }
            if (emptyState) emptyState.classList.add('hidden');
            document.body.classList.add('cursor-progress');
        } else {
            loadingSpinner.classList.add('hidden');
            if (loadingOverlay) {
                loadingOverlay.classList.add('hidden');
                loadingOverlay.classList.remove('flex');
            }
            document.body.classList.remove('cursor-progress');
        }
    }
