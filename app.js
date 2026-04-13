/* ═══════════════════════════════════════════════════════════════
   INSIGHT LEDGER — Document Automation & OCR Platform
   Application Logic
   ═══════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', () => {
    // ─── Init Modules ───
    DocumentStore.init();
    Navigation.init();
    Dashboard.init();
    ImportWizard.init();
    MasterLabels.init();
    OCRPreview.init();
    GeneratePage.init();
    ExportPage.init();
});

/* ═══════════════════════════════════════════════════════════════
   DOCUMENT STORE — Centralized data flow between modules
   ═══════════════════════════════════════════════════════════════ */
const DocumentStore = {
    // Uploaded templates (from Import Wizard)
    uploadedTemplates: [],
    // Generated documents (from Generate page)
    generatedDocuments: [],
    // OCR extracted data (from OCR Preview)
    ocrData: [],
    // Listeners for changes
    _listeners: {},

    init() {
        // Seed with demo templates so the app works out-of-the-box
        this.uploadedTemplates = [
            { id: 'tpl_001', name: 'Invoice_2024_001.pdf', size: '2.4 MB', ext: 'pdf', pages: 3, uploadedAt: new Date('2026-04-10') },
            { id: 'tpl_002', name: 'Receipt_ACME_Corp.pdf', size: '1.1 MB', ext: 'pdf', pages: 1, uploadedAt: new Date('2026-04-09') },
            { id: 'tpl_003', name: 'Contract_NDA_v2.pdf', size: '4.8 MB', ext: 'pdf', pages: 5, uploadedAt: new Date('2026-04-08') },
        ];
        // OCR data starts empty – populated by OCRPreview on document load
        this.ocrData = [];
    },

    addTemplate(file) {
        const id = 'tpl_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
        const template = {
            id,
            name: file.name,
            size: (file.size / 1024 / 1024).toFixed(2) + ' MB',
            ext: file.name.split('.').pop().toLowerCase(),
            pages: Math.ceil(Math.random() * 5) + 1,
            uploadedAt: new Date(),
            _file: file, // keep reference to the raw File object
        };
        this.uploadedTemplates.push(template);
        this._emit('templatesChanged');
        return template;
    },

    removeTemplate(id) {
        this.uploadedTemplates = this.uploadedTemplates.filter(t => t.id !== id);
        this._emit('templatesChanged');
    },

    addGeneratedDocuments(docs) {
        this.generatedDocuments.push(...docs);
        this._emit('documentsGenerated');
    },

    clearGeneratedDocuments() {
        this.generatedDocuments = [];
        this._emit('documentsGenerated');
    },

    getTemplateNames() {
        return this.uploadedTemplates.map(t => ({ id: t.id, name: t.name }));
    },

    on(event, callback) {
        if (!this._listeners[event]) this._listeners[event] = [];
        this._listeners[event].push(callback);
    },

    _emit(event) {
        (this._listeners[event] || []).forEach(cb => cb());
    },

    // Generate randomized data based on a field schema
    generateRandomValue(field) {
        const companies = ['ACME Corp', 'NextGen Solutions', 'TechVista Inc', 'Global Dynamics', 'Pinnacle Labs', 'BlueStar Holdings', 'RedPoint Systems', 'OmniTech Group'];
        const clients = ['Alpha Industries', 'Beta Solutions', 'Gamma Tech', 'Delta Corp', 'Epsilon LLC', 'Zeta Holdings', 'Eta Partners', 'Theta Consulting'];
        const rndInt = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
        const rndAmount = (min, max) => '$' + (Math.random() * (max - min) + min).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        const rndDate = () => { const d = new Date(2026, rndInt(0, 11), rndInt(1, 28)); return d.toISOString().split('T')[0]; };

        switch (field) {
            case 'company_name': return companies[rndInt(0, companies.length - 1)];
            case 'client_name': return clients[rndInt(0, clients.length - 1)];
            case 'invoice_number': return `INV-2026-${String(rndInt(1, 12)).padStart(2, '0')}-${String(rndInt(1, 9999)).padStart(4, '0')}`;
            case 'po_number': return `PO-${rndInt(10000, 99999)}`;
            case 'invoice_date': case 'due_date': return rndDate();
            case 'subtotal': return rndAmount(1000, 50000);
            case 'tax': return rndAmount(100, 5000);
            case 'total': return rndAmount(1100, 55000);
            default: return `Value_${rndInt(1000, 9999)}`;
        }
    }
};

/* ═══════════════════════════════════════════════════════════════
   NAVIGATION
   ═══════════════════════════════════════════════════════════════ */
const Navigation = {
    init() {
        const sidebar = document.getElementById('sidebar');
        const toggle = document.getElementById('sidebar-toggle');
        const navItems = document.querySelectorAll('.nav-item');

        toggle.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
        });

        navItems.forEach(item => {
            item.addEventListener('click', () => {
                const page = item.dataset.page;
                this.navigateTo(page);
            });
        });
    },

    navigateTo(pageId) {
        // Update nav
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        const activeNav = document.querySelector(`.nav-item[data-page="${pageId}"]`);
        if (activeNav) activeNav.classList.add('active');

        // Update pages
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        const activePage = document.getElementById(`page-${pageId}`);
        if (activePage) activePage.classList.add('active');

        // Update topbar
        const titles = {
            'dashboard': 'Dashboard',
            'import': 'Import & Configuration',
            'labels': 'Master Data Labels',
            'ocr': 'OCR Preview',
            'generate': 'Document Generation',
            'export': 'Batch Export',
        };
        const breadcrumbs = {
            'dashboard': 'Home / Dashboard',
            'import': 'Home / Import & Configuration',
            'labels': 'Home / Data / Master Labels',
            'ocr': 'Home / Data / OCR Preview',
            'generate': 'Home / Tools / Generate',
            'export': 'Home / Tools / Export',
        };
        document.getElementById('page-title').textContent = titles[pageId] || 'Dashboard';
        document.getElementById('page-breadcrumb').textContent = breadcrumbs[pageId] || '';

        // Re-animate KPIs if dashboard
        if (pageId === 'dashboard' || pageId === 'labels') {
            Dashboard.animateKPIs();
        }
    }
};

/* ═══════════════════════════════════════════════════════════════
   DASHBOARD
   ═══════════════════════════════════════════════════════════════ */
const Dashboard = {
    projects: [
        { name: 'Q1 Invoice Batch', date: 'Apr 10, 2026', docs: 342, accuracy: 98.2, status: 'Completed' },
        { name: 'Vendor Contracts Pack', date: 'Apr 09, 2026', docs: 56, accuracy: 95.4, status: 'Processing' },
        { name: 'Insurance Claims Set', date: 'Apr 08, 2026', docs: 128, accuracy: 92.1, status: 'Completed' },
        { name: 'HR Onboarding Forms', date: 'Apr 07, 2026', docs: 89, accuracy: 97.6, status: 'Completed' },
        { name: 'Tax Documents 2025', date: 'Apr 06, 2026', docs: 214, accuracy: 88.3, status: 'Queued' },
        { name: 'Receipt Digitization', date: 'Apr 05, 2026', docs: 567, accuracy: 96.8, status: 'Processing' },
    ],

    init() {
        this.renderProjects();
        this.animateKPIs();
        this.renderChart();
        this.initChartControls();
    },

    renderProjects() {
        const tbody = document.getElementById('projects-tbody');
        tbody.innerHTML = this.projects.map(p => {
            const statusClass = p.status === 'Completed' ? 'success' :
                p.status === 'Processing' ? 'info' : 'warning';
            const confidenceClass = p.accuracy >= 95 ? 'high' :
                p.accuracy >= 90 ? 'medium' : 'low';
            const badgeClass = p.accuracy >= 95 ? 'success' :
                p.accuracy >= 90 ? 'warning' : 'error';
            return `
                <tr>
                    <td><strong>${p.name}</strong></td>
                    <td>${p.date}</td>
                    <td>${p.docs}</td>
                    <td>
                        <div class="confidence-bar">
                            <div class="confidence-track">
                                <div class="confidence-fill confidence-fill--${confidenceClass}" style="width:${p.accuracy}%"></div>
                            </div>
                            <span class="badge badge--${badgeClass}">${p.accuracy}%</span>
                        </div>
                    </td>
                    <td><span class="badge badge--${statusClass}">${p.status}</span></td>
                    <td>
                        <button class="action-btn" aria-label="More actions for ${p.name}">
                            <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                                <circle cx="8" cy="3" r="1.5"/>
                                <circle cx="8" cy="8" r="1.5"/>
                                <circle cx="8" cy="13" r="1.5"/>
                            </svg>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    },

    animateKPIs() {
        document.querySelectorAll('.kpi-value').forEach(el => {
            const target = parseFloat(el.dataset.target);
            if (!target) return;
            const isPercent = el.classList.contains('kpi-value--percent');
            const duration = 1200;
            const start = performance.now();

            const animate = (now) => {
                const progress = Math.min((now - start) / duration, 1);
                const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
                const current = target * eased;

                if (isPercent) {
                    el.textContent = current.toFixed(1) + '%';
                } else {
                    el.textContent = Math.floor(current).toLocaleString();
                }

                if (progress < 1) requestAnimationFrame(animate);
            };
            requestAnimationFrame(animate);
        });
    },

    renderChart() {
        const canvas = document.getElementById('throughput-canvas');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;

        const container = canvas.parentElement;
        canvas.width = container.clientWidth * dpr;
        canvas.height = container.clientHeight * dpr;
        canvas.style.width = container.clientWidth + 'px';
        canvas.style.height = container.clientHeight + 'px';
        ctx.scale(dpr, dpr);

        const w = container.clientWidth;
        const h = container.clientHeight;
        const padding = { top: 20, right: 20, bottom: 30, left: 40 };
        const chartW = w - padding.left - padding.right;
        const chartH = h - padding.top - padding.bottom;

        // Data
        const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
        const processed = [180, 220, 195, 260, 240, 150, 210];
        const validated = [165, 198, 180, 245, 225, 138, 195];
        const errors = [12, 18, 8, 15, 10, 6, 14];

        const maxVal = Math.max(...processed) * 1.15;

        // Grid lines
        ctx.strokeStyle = '#E2E8F0';
        ctx.lineWidth = 0.5;
        for (let i = 0; i <= 4; i++) {
            const y = padding.top + (chartH / 4) * i;
            ctx.beginPath();
            ctx.moveTo(padding.left, y);
            ctx.lineTo(w - padding.right, y);
            ctx.stroke();

            // Y-axis labels
            ctx.fillStyle = '#8896A6';
            ctx.font = '10px Inter, sans-serif';
            ctx.textAlign = 'right';
            ctx.fillText(Math.round(maxVal - (maxVal / 4) * i), padding.left - 8, y + 3);
        }

        // X-axis labels
        ctx.textAlign = 'center';
        days.forEach((day, i) => {
            const x = padding.left + (chartW / (days.length - 1)) * i;
            ctx.fillStyle = '#8896A6';
            ctx.fillText(day, x, h - 8);
        });

        // Draw bars
        const barWidth = chartW / days.length * 0.6;
        const barGap = 3;

        processed.forEach((val, i) => {
            const x = padding.left + (chartW / days.length) * i + (chartW / days.length * 0.2);
            const barH = (val / maxVal) * chartH;
            const y = padding.top + chartH - barH;

            // Processed bar
            const grad = ctx.createLinearGradient(x, y, x, y + barH);
            grad.addColorStop(0, '#0453CD');
            grad.addColorStop(1, '#0453CD88');
            ctx.fillStyle = grad;
            ctx.beginPath();
            ctx.roundRect(x, y, barWidth / 3 - barGap, barH, [3, 3, 0, 0]);
            ctx.fill();

            // Validated bar
            const grad2 = ctx.createLinearGradient(x, y, x, y + barH);
            grad2.addColorStop(0, '#059669');
            grad2.addColorStop(1, '#05966988');
            const valH = (validated[i] / maxVal) * chartH;
            ctx.fillStyle = grad2;
            ctx.beginPath();
            ctx.roundRect(x + barWidth / 3, padding.top + chartH - valH, barWidth / 3 - barGap, valH, [3, 3, 0, 0]);
            ctx.fill();

            // Errors bar
            const errH = (errors[i] / maxVal) * chartH;
            const grad3 = ctx.createLinearGradient(x, padding.top + chartH - errH, x, padding.top + chartH);
            grad3.addColorStop(0, '#D97706');
            grad3.addColorStop(1, '#D9770688');
            ctx.fillStyle = grad3;
            ctx.beginPath();
            ctx.roundRect(x + (barWidth / 3) * 2, padding.top + chartH - errH, barWidth / 3 - barGap, errH, [3, 3, 0, 0]);
            ctx.fill();
        });
    },

    initChartControls() {
        document.querySelectorAll('.chip[data-range]').forEach(chip => {
            chip.addEventListener('click', () => {
                document.querySelectorAll('.chip[data-range]').forEach(c => c.classList.remove('active'));
                chip.classList.add('active');
                // Re-render chart with simulated different data
                this.renderChart();
            });
        });
    }
};

/* ═══════════════════════════════════════════════════════════════
   IMPORT WIZARD
   ═══════════════════════════════════════════════════════════════ */
const ImportWizard = {
    currentStep: 1,
    files: [],

    mappingData: [
        { source: 'Invoice_No', target: 'invoice_number', type: 'Alphanumeric', confidence: 98 },
        { source: 'Date_Issued', target: 'invoice_date', type: 'Date', confidence: 96 },
        { source: 'Due_Date', target: 'due_date', type: 'Date', confidence: 87 },
        { source: 'Client_Name', target: 'client_name', type: 'Alphanumeric', confidence: 94 },
        { source: 'Total_Amount', target: 'total', type: 'Currency', confidence: 99 },
        { source: 'Tax_Rate', target: 'tax_percentage', type: 'Percentage', confidence: 72 },
    ],

    init() {
        this.initUploadZone();
        this.initWizardNav();
        this.renderMappingTable();
        this.initThresholdSliders();
    },

    initUploadZone() {
        const zone = document.getElementById('upload-zone');
        const fileInput = document.getElementById('file-input');
        const browse = document.getElementById('upload-browse');

        browse.addEventListener('click', (e) => {
            e.stopPropagation();
            fileInput.click();
        });
        zone.addEventListener('click', () => fileInput.click());

        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            this.addFiles(e.dataTransfer.files);
        });

        fileInput.addEventListener('change', () => {
            this.addFiles(fileInput.files);
        });
    },

    addFiles(fileList) {
        const list = document.getElementById('upload-file-list');
        for (const file of fileList) {
            const ext = file.name.split('.').pop().toLowerCase();
            const iconClass = ext === 'pdf' ? 'pdf' : ext === 'jpg' || ext === 'jpeg' ? 'jpg' : 'png';
            const size = (file.size / 1024 / 1024).toFixed(2);

            // Register into central DocumentStore
            const template = DocumentStore.addTemplate(file);
            this.files.push({ id: template.id, name: file.name, size, ext });

            const item = document.createElement('div');
            item.className = 'upload-file-item';
            item.dataset.templateId = template.id;
            item.innerHTML = `
                <div class="file-icon file-icon--${iconClass}">${ext.toUpperCase()}</div>
                <div class="file-info">
                    <div class="file-name">${file.name}</div>
                    <div class="file-size">${size} MB</div>
                </div>
                <div class="file-progress">
                    <div class="file-progress-bar" style="width: 0%"></div>
                </div>
                <button class="file-remove" aria-label="Remove file">✕</button>
            `;
            list.appendChild(item);

            // Animate progress
            const bar = item.querySelector('.file-progress-bar');
            setTimeout(() => { bar.style.width = '100%'; }, 50);

            // Remove handler — also removes from DocumentStore
            const tplId = template.id;
            item.querySelector('.file-remove').addEventListener('click', () => {
                DocumentStore.removeTemplate(tplId);
                this.files = this.files.filter(f => f.id !== tplId);
                item.style.opacity = '0';
                item.style.transform = 'translateX(20px)';
                item.style.transition = 'all 0.3s ease';
                setTimeout(() => item.remove(), 300);
            });
        }
    },

    initWizardNav() {
        const backBtn = document.getElementById('wizard-back');
        const nextBtn = document.getElementById('wizard-next');

        nextBtn.addEventListener('click', () => {
            if (this.currentStep < 4) {
                this.goToStep(this.currentStep + 1);
            } else {
                // Final submit — navigate to OCR
                Navigation.navigateTo('ocr');
            }
        });

        backBtn.addEventListener('click', () => {
            if (this.currentStep > 1) {
                this.goToStep(this.currentStep - 1);
            }
        });
    },

    goToStep(step) {
        // Update step indicators
        document.querySelectorAll('.wizard-stepper .step').forEach((el, i) => {
            el.classList.remove('active', 'completed');
            if (i + 1 < step) el.classList.add('completed');
            if (i + 1 === step) el.classList.add('active');
        });
        document.querySelectorAll('.wizard-stepper .step-line').forEach((el, i) => {
            el.classList.remove('active', 'completed');
            if (i < step - 1) el.classList.add('completed');
            if (i === step - 1) el.classList.add('active');
        });

        // Update content
        document.querySelectorAll('.wizard-step').forEach(s => s.classList.remove('active'));
        document.getElementById(`wizard-step-${step}`).classList.add('active');

        // Update footer
        document.getElementById('wizard-back').disabled = step === 1;
        document.getElementById('wizard-next').textContent = step === 4 ? 'Start Processing' : 'Continue';
        document.getElementById('wizard-step-text').textContent = `Step ${step} of 4`;

        this.currentStep = step;

        // Render document preview on step 2
        if (step === 2) {
            this.renderStep2Preview();
        }

        // Update review summary on step 4
        if (step === 4) {
            this.updateReview();
            this.renderStep4Docs();
        }
    },

    /* ─── Step 2: PDF Preview ─── */
    _wizardPdf: null,
    _wizardPage: 1,
    _wizardTotalPages: 1,

    renderStep2Preview() {
        const templates = DocumentStore.templates.filter(t => t._file);
        const emptyEl = document.getElementById('wizard-preview-empty');
        const canvasArea = document.getElementById('wizard-preview-canvas-area');
        const select = document.getElementById('wizard-preview-select');

        if (templates.length === 0) {
            emptyEl.style.display = 'flex';
            canvasArea.style.display = 'none';
            select.innerHTML = '<option>— No documents uploaded —</option>';
            return;
        }

        // Populate dropdown
        select.innerHTML = templates.map((t, i) =>
            `<option value="${t.id}" ${i === 0 ? 'selected' : ''}>${t.name}</option>`
        ).join('');

        // Listen for dropdown change
        select.onchange = () => {
            const tpl = DocumentStore.templates.find(t => t.id === select.value);
            if (tpl && tpl._file) this._loadWizardPdf(tpl._file);
        };

        // Load first file
        this._loadWizardPdf(templates[0]._file);

        // Page navigation
        const prevBtn = document.getElementById('wizard-prev-page');
        const nextBtn = document.getElementById('wizard-next-page');
        prevBtn.onclick = () => { if (this._wizardPage > 1) { this._wizardPage--; this._renderWizardPage(); } };
        nextBtn.onclick = () => { if (this._wizardPage < this._wizardTotalPages) { this._wizardPage++; this._renderWizardPage(); } };
    },

    async _loadWizardPdf(file) {
        const emptyEl = document.getElementById('wizard-preview-empty');
        const canvasArea = document.getElementById('wizard-preview-canvas-area');

        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            this._wizardPdf = pdf;
            this._wizardPage = 1;
            this._wizardTotalPages = pdf.numPages;

            emptyEl.style.display = 'none';
            canvasArea.style.display = 'flex';

            this._renderWizardPage();
        } catch (err) {
            console.error('Wizard PDF load error:', err);
            emptyEl.querySelector('p').textContent = 'Unable to preview this file format';
            emptyEl.style.display = 'flex';
            canvasArea.style.display = 'none';
        }
    },

    async _renderWizardPage() {
        if (!this._wizardPdf) return;
        const page = await this._wizardPdf.getPage(this._wizardPage);
        const canvas = document.getElementById('wizard-preview-canvas');
        const ctx = canvas.getContext('2d');

        // Scale to fit container (max width ~480px)
        const viewport = page.getViewport({ scale: 1 });
        const containerWidth = document.getElementById('document-preview').clientWidth - 32;
        const scale = Math.min(containerWidth / viewport.width, 1.5);
        const scaledViewport = page.getViewport({ scale });

        canvas.width = scaledViewport.width;
        canvas.height = scaledViewport.height;

        await page.render({ canvasContext: ctx, viewport: scaledViewport }).promise;

        // Update page info
        document.getElementById('wizard-page-info').textContent =
            `${this._wizardPage} / ${this._wizardTotalPages}`;
    },

    /* ─── Step 4: Review ─── */
    updateReview() {
        const realFiles = DocumentStore.templates.filter(t => t._file);
        const fileCount = realFiles.length || this.files.length;
        document.getElementById('review-files').textContent = `${fileCount} document${fileCount !== 1 ? 's' : ''}`;

        const precision = document.querySelector('input[name="precision"]:checked');
        document.getElementById('review-precision').textContent = precision ? precision.value.charAt(0).toUpperCase() + precision.value.slice(1) : 'Balanced';
        document.getElementById('review-handwriting').textContent = document.getElementById('handwriting-toggle').checked ? 'Enabled' : 'Disabled';
        document.getElementById('review-autovalidate').textContent = document.getElementById('auto-validate-toggle').checked ? 'Enabled' : 'Disabled';
        document.getElementById('review-language').textContent = document.getElementById('language-select').value;
        document.getElementById('review-fields').textContent = `${this.mappingData.length} fields`;

        // Compute real estimates
        const totalPages = realFiles.reduce((sum, t) => sum + (t.pages || 1), 0) || 1;
        const precisionVal = precision ? precision.value : 'balanced';
        const secsPerPage = precisionVal === 'fast' ? 2 : precisionVal === 'ultra' ? 12 : 5;
        const estTime = totalPages * secsPerPage;
        const estCredits = totalPages * 2;

        document.getElementById('review-est-pages').textContent = totalPages;
        document.getElementById('review-est-time').textContent = estTime < 60 ? `~${estTime}s` : `~${Math.ceil(estTime / 60)}m`;
        document.getElementById('review-est-credits').textContent = estCredits;

        // Update progress ring (proportion of max 5 minutes = 300s)
        const circumference = 327;
        const ratio = Math.min(estTime / 300, 1);
        const offset = circumference * (1 - ratio);
        const progressCircle = document.querySelector('.estimate-progress');
        if (progressCircle) progressCircle.style.strokeDashoffset = offset;
    },

    async renderStep4Docs() {
        const grid = document.getElementById('review-docs-grid');
        const countBadge = document.getElementById('review-doc-count');
        const realFiles = DocumentStore.templates.filter(t => t._file);

        countBadge.textContent = `${realFiles.length} file${realFiles.length !== 1 ? 's' : ''}`;

        if (realFiles.length === 0) {
            grid.innerHTML = `
                <div class="review-docs-empty">
                    <svg width="40" height="40" viewBox="0 0 40 40" fill="none" stroke="var(--color-text-muted)" stroke-width="1.2">
                        <rect x="6" y="3" width="28" height="34" rx="3"/>
                        <path d="M12 12h16M12 18h12M12 24h8"/>
                    </svg>
                    <p>No documents uploaded yet. Go back to Step 1 to add files.</p>
                </div>`;
            return;
        }

        grid.innerHTML = '';

        for (const tpl of realFiles) {
            const card = document.createElement('div');
            card.className = 'review-doc-card';

            // Thumbnail canvas
            const thumbCanvas = document.createElement('canvas');
            thumbCanvas.className = 'review-doc-thumb';

            // Try to render PDF first page
            try {
                const arrayBuffer = await tpl._file.arrayBuffer();
                const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                const page = await pdf.getPage(1);
                const vp = page.getViewport({ scale: 0.4 });
                thumbCanvas.width = vp.width;
                thumbCanvas.height = vp.height;
                await page.render({ canvasContext: thumbCanvas.getContext('2d'), viewport: vp }).promise;
                tpl.pages = pdf.numPages; // update real page count
            } catch (e) {
                // Non-PDF fallback: draw a file icon placeholder
                thumbCanvas.width = 120;
                thumbCanvas.height = 160;
                const ctx = thumbCanvas.getContext('2d');
                ctx.fillStyle = '#F0F2F5';
                ctx.fillRect(0, 0, 120, 160);
                ctx.fillStyle = '#A0AEC0';
                ctx.font = '11px Inter, sans-serif';
                ctx.textAlign = 'center';
                ctx.fillText(tpl.ext?.toUpperCase() || 'FILE', 60, 85);
            }

            const size = tpl._file.size ? (tpl._file.size / 1024 / 1024).toFixed(2) + ' MB' : tpl.size;

            card.innerHTML = `
                <div class="review-doc-thumb-wrap"></div>
                <div class="review-doc-meta">
                    <span class="review-doc-name" title="${tpl.name}">${tpl.name}</span>
                    <span class="review-doc-info">${tpl.pages || 1} page${(tpl.pages || 1) > 1 ? 's' : ''} · ${size}</span>
                </div>
            `;
            card.querySelector('.review-doc-thumb-wrap').appendChild(thumbCanvas);
            grid.appendChild(card);
        }
    },

    renderMappingTable() {
        const tbody = document.getElementById('mapping-tbody');
        const targetFields = ['invoice_number', 'invoice_date', 'due_date', 'client_name', 'total', 'tax_percentage', 'po_number', 'subtotal', 'company_name'];

        tbody.innerHTML = this.mappingData.map(m => {
            const confClass = m.confidence >= 90 ? 'high' : m.confidence >= 80 ? 'medium' : 'low';
            const badgeClass = m.confidence >= 90 ? 'success' : m.confidence >= 80 ? 'warning' : 'error';
            return `
                <tr>
                    <td><code style="font-size:11px;padding:2px 6px;background:#F0F2F5;border-radius:3px;">${m.source}</code></td>
                    <td class="mapping-arrow-col">
                        <span class="mapping-arrow">
                            <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 7h10M9 4l3 3-3 3"/></svg>
                        </span>
                    </td>
                    <td>
                        <select class="mapping-select">
                            ${targetFields.map(f => `<option ${f === m.target ? 'selected' : ''}>${f}</option>`).join('')}
                        </select>
                    </td>
                    <td><span class="badge badge--neutral">${m.type}</span></td>
                    <td>
                        <div class="confidence-bar">
                            <div class="confidence-track">
                                <div class="confidence-fill confidence-fill--${confClass}" style="width:${m.confidence}%"></div>
                            </div>
                            <span class="badge badge--${badgeClass}">${m.confidence}%</span>
                        </div>
                    </td>
                    <td>
                        <button class="action-btn" aria-label="Remove mapping">
                            <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M2 2l10 10M12 2L2 12"/></svg>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    },

    initThresholdSliders() {
        const slider = document.getElementById('label-threshold-slider');
        const value = document.getElementById('label-threshold-value');
        if (slider && value) {
            slider.addEventListener('input', () => {
                value.textContent = slider.value + '%';
            });
        }
    }
};

/* ═══════════════════════════════════════════════════════════════
   MASTER DATA LABELS
   ═══════════════════════════════════════════════════════════════ */
const MasterLabels = {
    labels: [
        { name: 'Invoice Number', type: 'Alphanumeric', threshold: 95, tag: 'INV_NUM', status: 'Active' },
        { name: 'Invoice Date', type: 'Date', threshold: 92, tag: 'INV_DATE', status: 'Active' },
        { name: 'Due Date', type: 'Date', threshold: 88, tag: 'DUE_DATE', status: 'Active' },
        { name: 'Client Name', type: 'Alphanumeric', threshold: 90, tag: 'CLIENT', status: 'Active' },
        { name: 'PO Number', type: 'Alphanumeric', threshold: 75, tag: 'PO_NUM', status: 'Warning' },
        { name: 'Subtotal', type: 'Currency', threshold: 97, tag: 'SUBTOTAL', status: 'Active' },
        { name: 'Tax Amount', type: 'Currency', threshold: 93, tag: 'TAX_AMT', status: 'Active' },
        { name: 'Total Amount', type: 'Currency', threshold: 99, tag: 'TOTAL', status: 'Active' },
        { name: 'Company Name', type: 'Alphanumeric', threshold: 96, tag: 'COMPANY', status: 'Active' },
        { name: 'Company Address', type: 'Text Block', threshold: 84, tag: 'ADDR', status: 'Active' },
        { name: 'Tax Rate', type: 'Percentage', threshold: 68, tag: 'TAX_RATE', status: 'Low Confidence' },
        { name: 'Line Items', type: 'Text Block', threshold: 82, tag: 'ITEMS', status: 'Active' },
    ],

    init() {
        this.renderLabels();
        this.initModal();
    },

    renderLabels() {
        const tbody = document.getElementById('labels-tbody');
        tbody.innerHTML = this.labels.map((l, i) => {
            const confClass = l.threshold >= 90 ? 'high' : l.threshold >= 80 ? 'medium' : 'low';
            const statusBadge = l.status === 'Active' ? 'success' :
                l.status === 'Warning' ? 'warning' : 'error';
            return `
                <tr>
                    <td><input type="checkbox" aria-label="Select ${l.name}"></td>
                    <td><strong>${l.name}</strong></td>
                    <td><span class="badge badge--neutral">${l.type}</span></td>
                    <td>
                        <div class="confidence-bar">
                            <div class="confidence-track">
                                <div class="confidence-fill confidence-fill--${confClass}" style="width:${l.threshold}%"></div>
                            </div>
                            <span style="font-family:var(--font-data);font-size:11px;font-weight:600;color:var(--color-text-secondary)">${l.threshold}%</span>
                        </div>
                    </td>
                    <td><code style="font-size:10px;padding:2px 6px;background:#F0F2F5;border-radius:3px;font-family:var(--font-mono)">${l.tag}</code></td>
                    <td><span class="badge badge--${statusBadge}">${l.status}</span></td>
                    <td>
                        <button class="action-btn" aria-label="Edit ${l.name}">
                            <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5">
                                <path d="M8.5 2.5l3 3L4 13H1v-3L8.5 2.5z"/>
                            </svg>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    },

    initModal() {
        const overlay = document.getElementById('label-modal-overlay');
        const addBtn = document.getElementById('add-label-btn');
        const closeBtn = document.getElementById('label-modal-close');
        const cancelBtn = document.getElementById('label-modal-cancel');
        const saveBtn = document.getElementById('label-modal-save');

        addBtn.addEventListener('click', () => overlay.classList.add('visible'));
        closeBtn.addEventListener('click', () => overlay.classList.remove('visible'));
        cancelBtn.addEventListener('click', () => overlay.classList.remove('visible'));
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) overlay.classList.remove('visible');
        });

        saveBtn.addEventListener('click', () => {
            const name = document.getElementById('label-name-input').value;
            const type = document.getElementById('label-type-select').value;
            const tag = document.getElementById('label-tag-input').value;
            const threshold = parseInt(document.getElementById('label-threshold-slider').value);

            if (name && tag) {
                this.labels.push({ name, type, threshold, tag, status: threshold >= 80 ? 'Active' : 'Warning' });
                this.renderLabels();
                overlay.classList.remove('visible');

                // Clear form
                document.getElementById('label-name-input').value = '';
                document.getElementById('label-tag-input').value = '';
                document.getElementById('label-threshold-slider').value = 85;
                document.getElementById('label-threshold-value').textContent = '85%';
            }
        });
    }
};

/* ═══════════════════════════════════════════════════════════════
   OCR PREVIEW
   ═══════════════════════════════════════════════════════════════ */
const OCRPreview = {
    // State
    pdfDoc: null,
    currentPage: 1,
    totalPages: 0,
    zoomLevel: 100,
    entities: [],
    currentFileId: null,
    pageTexts: {}, // cache: page number → extracted text

    init() {
        this.initDocumentSelector();
        this.initPageNavigation();
        this.initZoomControls();
        this.initResizer();
        this.initSuggestionBanner();
        this.initDirectUpload();
        this.initSaveButton();

        // Listen for new uploads
        DocumentStore.on('templatesChanged', () => this.refreshDocumentSelector());
    },

    // ─── Document Selector (from DocumentStore) ─────────────────
    initDocumentSelector() {
        const select = document.getElementById('ocr-document-select');
        this.refreshDocumentSelector();

        select.addEventListener('change', () => {
            const tpl = DocumentStore.uploadedTemplates.find(t => t.id === select.value);
            if (tpl && tpl._file) {
                this.loadFile(tpl._file, tpl.id);
            }
        });
    },

    refreshDocumentSelector() {
        const select = document.getElementById('ocr-document-select');
        const currentVal = select.value;
        select.innerHTML = '';

        const templates = DocumentStore.uploadedTemplates;
        if (templates.length === 0) {
            select.innerHTML = '<option value="">No documents uploaded</option>';
            return;
        }

        // Add placeholder
        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent = '— Select a document —';
        select.appendChild(placeholder);

        templates.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t.id;
            opt.textContent = t.name;
            // Disable non-PDF for now (only PDF.js supported)
            if (t.ext && t.ext !== 'pdf') {
                opt.textContent += ' (image)';
            }
            select.appendChild(opt);
        });

        // Try to restore previous selection
        if (currentVal && Array.from(select.options).some(o => o.value === currentVal)) {
            select.value = currentVal;
        }
    },

    // ─── Direct Upload from OCR page ────────────────────────────
    initDirectUpload() {
        const input = document.getElementById('ocr-direct-upload');
        if (!input) return;
        input.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            // Register in DocumentStore
            const tpl = DocumentStore.addTemplate(file);
            // Select it in dropdown and load
            setTimeout(() => {
                const select = document.getElementById('ocr-document-select');
                select.value = tpl.id;
                this.loadFile(file, tpl.id);
            }, 100);
        });
    },

    // ─── Load & Render PDF ──────────────────────────────────────
    async loadFile(file, fileId) {
        this.currentFileId = fileId;
        this.pageTexts = {};
        this.entities = [];
        this.currentPage = 1;

        // Show loading
        document.getElementById('ocr-upload-prompt').style.display = 'none';
        document.getElementById('ocr-pdf-viewer').style.display = 'flex';
        document.getElementById('ocr-loading-overlay').style.display = 'flex';

        // Handle image files (non-PDF)
        const ext = file.name.split('.').pop().toLowerCase();
        if (['png', 'jpg', 'jpeg', 'webp'].includes(ext)) {
            await this.loadImage(file);
            return;
        }

        try {
            const arrayBuffer = await file.arrayBuffer();
            this.pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            this.totalPages = this.pdfDoc.numPages;

            // Update page info in DocumentStore
            const tpl = DocumentStore.uploadedTemplates.find(t => t.id === fileId);
            if (tpl) tpl.pages = this.totalPages;

            // Render first page
            await this.renderPage(1);

            // Generate thumbnails
            await this.generateThumbnails();

            // Extract text from ALL pages
            await this.extractAllPageText();

            // Run entity extraction
            this.detectEntities();

            document.getElementById('ocr-loading-overlay').style.display = 'none';
        } catch (err) {
            console.error('PDF load error:', err);
            document.getElementById('ocr-loading-overlay').style.display = 'none';
            this.showLoadError(err.message);
        }
    },

    async loadImage(file) {
        const canvas = document.getElementById('ocr-pdf-canvas');
        const ctx = canvas.getContext('2d');
        const img = new Image();
        const url = URL.createObjectURL(file);

        img.onload = () => {
            canvas.width = img.width;
            canvas.height = img.height;
            ctx.drawImage(img, 0, 0);
            URL.revokeObjectURL(url);
            this.totalPages = 1;
            this.updatePageInfo();
            document.getElementById('ocr-loading-overlay').style.display = 'none';
            document.getElementById('ocr-page-thumbs').innerHTML = '';

            // For images, we can't extract text via PDF.js
            this.entities = [
                { field: 'note', label: 'Note', value: 'Image files require OCR service for text extraction', confidence: 0, level: 'low' }
            ];
            this.renderEntities();
        };
        img.src = url;
    },

    async renderPage(pageNum) {
        const page = await this.pdfDoc.getPage(pageNum);
        const scale = (this.zoomLevel / 100) * 1.5; // Base scale for clarity
        const viewport = page.getViewport({ scale });

        const canvas = document.getElementById('ocr-pdf-canvas');
        const ctx = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        await page.render({ canvasContext: ctx, viewport }).promise;

        this.currentPage = pageNum;
        this.updatePageInfo();
        this.highlightThumb(pageNum);
    },

    async generateThumbnails() {
        const container = document.getElementById('ocr-page-thumbs');
        container.innerHTML = '';

        if (this.totalPages <= 1) {
            container.style.display = 'none';
            return;
        }
        container.style.display = 'flex';

        for (let i = 1; i <= this.totalPages; i++) {
            const page = await this.pdfDoc.getPage(i);
            const vp = page.getViewport({ scale: 0.3 });
            const thumbCanvas = document.createElement('canvas');
            thumbCanvas.width = vp.width;
            thumbCanvas.height = vp.height;
            thumbCanvas.className = 'ocr-thumb-canvas';
            thumbCanvas.dataset.page = i;
            thumbCanvas.title = `Page ${i}`;

            const ctx = thumbCanvas.getContext('2d');
            await page.render({ canvasContext: ctx, viewport: vp }).promise;

            const wrapper = document.createElement('div');
            wrapper.className = 'ocr-thumb-item' + (i === 1 ? ' active' : '');
            wrapper.dataset.page = i;
            wrapper.innerHTML = `<span class="ocr-thumb-label">${i}</span>`;
            wrapper.prepend(thumbCanvas);

            wrapper.addEventListener('click', () => {
                this.renderPage(i);
                // Re-extract entities for this page
                this.detectEntities();
            });

            container.appendChild(wrapper);
        }
    },

    highlightThumb(pageNum) {
        document.querySelectorAll('.ocr-thumb-item').forEach(t => t.classList.remove('active'));
        const active = document.querySelector(`.ocr-thumb-item[data-page="${pageNum}"]`);
        if (active) {
            active.classList.add('active');
            active.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    },

    // ─── Text Extraction ────────────────────────────────────────
    async extractAllPageText() {
        const allText = [];
        for (let i = 1; i <= this.totalPages; i++) {
            const page = await this.pdfDoc.getPage(i);
            const textContent = await page.getTextContent();
            const text = textContent.items.map(item => item.str).join(' ');
            this.pageTexts[i] = text;
            allText.push(text);
        }
        // Store full document text
        this.fullText = allText.join('\n\n--- PAGE BREAK ---\n\n');
    },

    // ─── Smart Entity Detection (regex pattern matching) ────────
    detectEntities() {
        const text = this.fullText || '';
        if (!text.trim()) {
            this.entities = [];
            this.renderEntities();
            return;
        }

        const entities = [];
        const docType = this.detectDocumentType(text);

        // ── Common patterns (work for all document types) ──
        // Contract / Document Number
        const numPatterns = [
            { regex: /(?:NO|No|Number|NUMBER|#)[.:;\s]*([A-Z0-9][\w\-\/]+)/i, label: 'Document Number' },
            { regex: /(?:Contract|Invoice|Receipt|PO|SO|Ref)[.\s#:]*([A-Z0-9][\w\-\/]+)/i, label: 'Reference Number' },
        ];
        for (const p of numPatterns) {
            const m = text.match(p.regex);
            if (m) {
                entities.push({ field: 'doc_number', label: p.label, value: m[1].trim(), confidence: 95.2, level: 'high' });
                break;
            }
        }

        // Dates
        const datePatterns = [
            /(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/g,
            /(\w+ \d{1,2},?\s*\d{4})/g,
        ];
        const foundDates = [];
        for (const dp of datePatterns) {
            let dm;
            while ((dm = dp.exec(text)) !== null) {
                const val = dm[1].trim();
                if (!foundDates.includes(val) && foundDates.length < 3) foundDates.push(val);
            }
        }
        if (foundDates.length > 0) {
            entities.push({ field: 'date_primary', label: 'Date', value: foundDates[0], confidence: 93.8, level: 'high' });
        }
        if (foundDates.length > 1) {
            entities.push({ field: 'date_secondary', label: 'Secondary Date', value: foundDates[1], confidence: 82.1, level: 'medium' });
        }

        // Money / Amount
        const moneyPatterns = [
            /(?:TOTAL|Total|Amount|AMOUNT)[:\s]*\$?([\d,]+\.?\d*)/gi,
            /([\d,]+\.\d{2})/g,
        ];
        const foundAmounts = [];
        for (const mp of moneyPatterns) {
            let mm;
            while ((mm = mp.exec(text)) !== null) {
                const val = mm[1].replace(/,/g, '');
                const num = parseFloat(val);
                if (num > 0 && !foundAmounts.some(a => a.raw === val)) {
                    foundAmounts.push({ raw: val, num, original: mm[0].trim() });
                }
            }
        }
        // Sort by value descending, take top amounts
        foundAmounts.sort((a, b) => b.num - a.num);
        if (foundAmounts.length > 0) {
            entities.push({
                field: 'total_amount', label: 'Total Amount',
                value: foundAmounts[0].num.toLocaleString('en-US', { minimumFractionDigits: 2 }),
                confidence: 97.1, level: 'high'
            });
        }
        if (foundAmounts.length > 1) {
            entities.push({
                field: 'unit_price', label: 'Unit Price',
                value: foundAmounts[1].num.toLocaleString('en-US', { minimumFractionDigits: 2 }),
                confidence: 88.5, level: 'medium'
            });
        }

        // ── Document-type specific patterns ──
        if (docType === 'Sales Contract' || docType === 'Contract') {
            // BUYER
            const buyerMatch = text.match(/BUYER[:\s]*(.+?)(?:(?:No|Tel|Address|SELLER)|$)/is);
            if (buyerMatch) {
                const buyerLine = buyerMatch[1].split(/\r?\n/)[0].trim().substring(0, 120);
                entities.push({ field: 'buyer', label: 'Buyer', value: buyerLine, confidence: 91.3, level: 'high' });
            }
            // SELLER
            const sellerMatch = text.match(/SELLER[:\s]*(.+?)(?:(?:Both|1\.|COMMODITY)|$)/is);
            if (sellerMatch) {
                const sellerLine = sellerMatch[1].split(/\r?\n/)[0].trim().substring(0, 120);
                entities.push({ field: 'seller', label: 'Seller', value: sellerLine, confidence: 90.7, level: 'high' });
            }
            // Goods
            const goodsMatch = text.match(/(?:Goods|GOODS|Description|DESCRIPTION)[:\s]*(.+?)(?:Quantity|QUANTITY|$)/is);
            if (goodsMatch) {
                entities.push({ field: 'goods', label: 'Goods Description', value: goodsMatch[1].trim().substring(0, 150), confidence: 85.4, level: 'medium' });
            }
            // Quantity
            const qtyMatch = text.match(/(?:Quantity|QTY|Qty)[:\s]*(\d+)/i);
            if (qtyMatch) {
                entities.push({ field: 'quantity', label: 'Quantity', value: qtyMatch[1], confidence: 94.2, level: 'high' });
            }
            // Shipment terms
            const shipMatch = text.match(/(?:Shipment|SHIPMENT)[:\s]*(.+?)(?:\d\.|$)/is);
            if (shipMatch) {
                entities.push({ field: 'shipment', label: 'Shipment Terms', value: shipMatch[1].trim().substring(0, 100), confidence: 78.6, level: 'medium' });
            }
            // Payment terms
            const payMatch = text.match(/(?:PAYMENT|Payment)\s*(?:TERMS|Terms)[:\s]*(.+?)(?:\d\.|$)/is);
            if (payMatch) {
                const payInfo = payMatch[1].trim().substring(0, 120);
                entities.push({ field: 'payment', label: 'Payment Terms', value: payInfo, confidence: 80.2, level: 'medium' });
            }
        }

        if (docType === 'Invoice') {
            const clientMatch = text.match(/(?:Bill\s*To|Client|Customer)[:\s]*(.+)/i);
            if (clientMatch) {
                entities.push({ field: 'client', label: 'Client', value: clientMatch[1].trim().substring(0, 80), confidence: 92.0, level: 'high' });
            }
            const poMatch = text.match(/(?:PO|Purchase Order)[:\s#]*([A-Z0-9\-]+)/i);
            if (poMatch) {
                entities.push({ field: 'po_number', label: 'PO Number', value: poMatch[1], confidence: 87.5, level: 'medium' });
            }
        }

        // ── Fallback: Company names (look for CO., LTD, Corp, Inc) ──
        const companyPattern = /([A-Z][A-Z\s&.,']+(?:CO\.?,?\s*LTD\.?|CORPORATION|CORP\.?|INC\.?|LLC|COMPANY|TRADING\s*CO))/gi;
        const companies = [];
        let cm;
        while ((cm = companyPattern.exec(text)) !== null) {
            const name = cm[1].trim();
            if (name.length > 5 && !companies.includes(name) && companies.length < 4) {
                companies.push(name);
            }
        }
        companies.forEach((c, i) => {
            if (!entities.some(e => e.value.includes(c.substring(0, 20)))) {
                entities.push({
                    field: `company_${i}`, label: i === 0 ? 'Company (Primary)' : `Company (${i + 1})`,
                    value: c, confidence: 89.0 - i * 5, level: i === 0 ? 'high' : 'medium'
                });
            }
        });

        // ── Add document type entity ──
        entities.unshift({
            field: 'doc_type', label: 'Document Type', value: docType, confidence: 96.5, level: 'high'
        });

        this.entities = entities;
        this.renderEntities();
        this.updateDocumentStoreOCR();
        this.showAISuggestion(docType);
    },

    detectDocumentType(text) {
        const upper = text.toUpperCase();
        if (upper.includes('SALES CONTRACT') || upper.includes('SALE CONTRACT')) return 'Sales Contract';
        if (upper.includes('PURCHASE ORDER') || upper.includes('P.O.')) return 'Purchase Order';
        if (upper.includes('INVOICE') && !upper.includes('COMMERCIAL INVOICE')) return 'Invoice';
        if (upper.includes('COMMERCIAL INVOICE')) return 'Commercial Invoice';
        if (upper.includes('PACKING LIST')) return 'Packing List';
        if (upper.includes('BILL OF LADING') || upper.includes('B/L')) return 'Bill of Lading';
        if (upper.includes('RECEIPT')) return 'Receipt';
        if (upper.includes('CONTRACT') || upper.includes('AGREEMENT')) return 'Contract';
        if (upper.includes('CERTIFICATE')) return 'Certificate';
        if (upper.includes('QUOTATION') || upper.includes('QUOTE')) return 'Quotation';
        return 'Document';
    },

    // ─── Render Entities ────────────────────────────────────────
    renderEntities() {
        const list = document.getElementById('entities-list');
        const emptyState = document.getElementById('ocr-empty-entities');

        if (this.entities.length === 0) {
            if (emptyState) emptyState.style.display = 'flex';
            list.querySelectorAll('.entity-item').forEach(e => e.remove());
            document.getElementById('entities-overall-confidence').textContent = 'Overall: —%';
            return;
        }

        if (emptyState) emptyState.style.display = 'none';

        // Calculate overall confidence
        const avg = (this.entities.reduce((sum, e) => sum + e.confidence, 0) / this.entities.length).toFixed(1);
        document.getElementById('entities-overall-confidence').textContent = `Overall: ${avg}%`;

        // Remove old items (keep empty state div)
        list.querySelectorAll('.entity-item').forEach(e => e.remove());

        this.entities.forEach(e => {
            const inputClass = e.level === 'low' ? 'entity-input entity-input--error' : 'entity-input';
            const div = document.createElement('div');
            div.className = 'entity-item';
            div.dataset.field = e.field;
            div.innerHTML = `
                <div class="entity-header">
                    <span class="entity-label">${e.label}</span>
                    <span class="entity-confidence entity-confidence--${e.level}">${e.confidence}%</span>
                </div>
                <input type="text" class="${inputClass}" value="${this._escapeHtml(e.value)}" aria-label="${e.label} value">
            `;
            list.appendChild(div);
        });
    },

    _escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    },

    // ─── Update DocumentStore with real OCR data ────────────────
    updateDocumentStoreOCR() {
        DocumentStore.ocrData = this.entities.map(e => ({
            field: e.field,
            label: e.label,
            value: e.value,
            type: this._guessFieldType(e),
        }));
    },

    _guessFieldType(entity) {
        if (/date/i.test(entity.label)) return 'Date';
        if (/amount|price|total|subtotal|tax/i.test(entity.label)) return 'Currency';
        if (/quantity|qty|count/i.test(entity.label)) return 'Integer';
        return 'Alphanumeric';
    },

    // ─── Page Navigation ────────────────────────────────────────
    initPageNavigation() {
        document.getElementById('ocr-page-prev').addEventListener('click', () => {
            if (this.currentPage > 1) {
                this.renderPage(this.currentPage - 1);
            }
        });
        document.getElementById('ocr-page-next').addEventListener('click', () => {
            if (this.currentPage < this.totalPages) {
                this.renderPage(this.currentPage + 1);
            }
        });
    },

    updatePageInfo() {
        document.getElementById('ocr-page-info').textContent = `Page ${this.currentPage} of ${this.totalPages}`;
    },

    // ─── Zoom Controls ──────────────────────────────────────────
    initZoomControls() {
        const zoomIn = document.getElementById('ocr-zoom-in');
        const zoomOut = document.getElementById('ocr-zoom-out');
        const zoomLevel = document.getElementById('ocr-zoom-level');

        zoomIn.addEventListener('click', () => {
            if (this.zoomLevel < 200) {
                this.zoomLevel += 25;
                zoomLevel.textContent = this.zoomLevel + '%';
                if (this.pdfDoc) this.renderPage(this.currentPage);
            }
        });

        zoomOut.addEventListener('click', () => {
            if (this.zoomLevel > 50) {
                this.zoomLevel -= 25;
                zoomLevel.textContent = this.zoomLevel + '%';
                if (this.pdfDoc) this.renderPage(this.currentPage);
            }
        });
    },

    // ─── Resizer ────────────────────────────────────────────────
    initResizer() {
        const resizer = document.getElementById('ocr-resizer');
        const leftPanel = document.getElementById('ocr-document-panel');
        const container = document.getElementById('ocr-split-view');
        let isResizing = false;

        resizer.addEventListener('mousedown', () => {
            isResizing = true;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
        });

        document.addEventListener('mousemove', (e) => {
            if (!isResizing) return;
            const rect = container.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const fraction = x / rect.width;
            if (fraction > 0.3 && fraction < 0.8) {
                leftPanel.style.flex = `0 0 ${fraction * 100}%`;
            }
        });

        document.addEventListener('mouseup', () => {
            isResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        });
    },

    // ─── AI Suggestion Banner ───────────────────────────────────
    showAISuggestion(docType) {
        const banner = document.getElementById('ai-suggestion-banner');
        const label = document.getElementById('ai-suggestion-label');
        if (docType && docType !== 'Document') {
            label.textContent = `Detected "${docType}" — Apply matching template for optimized extraction.`;
            banner.style.display = '';
            banner.style.opacity = '1';
        }
    },

    initSuggestionBanner() {
        const dismissBtn = document.getElementById('dismiss-template-btn');
        const banner = document.getElementById('ai-suggestion-banner');
        if (!dismissBtn || !banner) return;

        dismissBtn.addEventListener('click', () => {
            banner.style.opacity = '0';
            banner.style.transform = 'translateY(10px)';
            banner.style.transition = 'all 0.3s ease';
            setTimeout(() => banner.style.display = 'none', 300);
        });

        const applyBtn = document.getElementById('apply-template-btn');
        applyBtn.addEventListener('click', () => {
            banner.innerHTML = `
                <div class="ai-suggestion-icon" style="background:var(--color-success)">
                    <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M6 10l3 3 5-5"/>
                    </svg>
                </div>
                <div class="ai-suggestion-text">
                    <strong style="color:var(--color-success)">Template Applied!</strong> Document template has been applied successfully.
                </div>
            `;
            setTimeout(() => {
                banner.style.opacity = '0';
                setTimeout(() => banner.style.display = 'none', 300);
            }, 3000);
        });
    },

    // ─── Save Button ────────────────────────────────────────────
    initSaveButton() {
        const btn = document.getElementById('ocr-save-btn');
        btn.addEventListener('click', () => {
            // Collect edited values from inputs
            document.querySelectorAll('.entity-item').forEach(item => {
                const field = item.dataset.field;
                const input = item.querySelector('.entity-input');
                const entity = this.entities.find(e => e.field === field);
                if (entity && input) {
                    entity.value = input.value;
                }
            });
            this.updateDocumentStoreOCR();

            // Visual feedback
            const orig = btn.textContent;
            btn.textContent = '✓ Saved!';
            btn.style.background = 'var(--color-success)';
            btn.style.borderColor = 'var(--color-success)';
            setTimeout(() => {
                btn.textContent = orig;
                btn.style.background = '';
                btn.style.borderColor = '';
            }, 2000);
        });
    },

    showLoadError(message) {
        const list = document.getElementById('entities-list');
        const items = list.querySelectorAll('.entity-item');
        items.forEach(e => e.remove());

        const emptyState = document.getElementById('ocr-empty-entities');
        if (emptyState) {
            emptyState.innerHTML = `
                <svg width="32" height="32" viewBox="0 0 32 32" fill="none" stroke="var(--color-error)" stroke-width="1.2">
                    <circle cx="16" cy="16" r="12"/><path d="M12 12l8 8M20 12l-8 8"/>
                </svg>
                <p style="color:var(--color-error)">Error loading document: ${message}</p>
            `;
            emptyState.style.display = 'flex';
        }
    }
};

/* ═══════════════════════════════════════════════════════════════
   GENERATE PAGE
   ═══════════════════════════════════════════════════════════════ */
const GeneratePage = {
    isGenerating: false,

    init() {
        const generateBtn = document.getElementById('generate-btn');
        generateBtn.addEventListener('click', () => this.startGeneration());

        // Populate template dropdown from DocumentStore
        this.refreshTemplateDropdown();

        // Listen for template changes (new uploads / removals)
        DocumentStore.on('templatesChanged', () => this.refreshTemplateDropdown());

        // Show/hide CSV upload area based on data source selection
        const dataSourceSelect = document.getElementById('gen-data-source');
        dataSourceSelect.addEventListener('change', () => this.onDataSourceChange());

        // CSV file input handler
        const csvInput = document.getElementById('gen-csv-input');
        if (csvInput) {
            csvInput.addEventListener('change', (e) => this.onCsvSelected(e));
        }
    },

    refreshTemplateDropdown() {
        const select = document.getElementById('gen-source-select');
        const templates = DocumentStore.getTemplateNames();
        const currentValue = select.value;

        select.innerHTML = '';

        if (templates.length === 0) {
            select.innerHTML = '<option value="" disabled selected>No templates — upload via Import</option>';
            return;
        }

        templates.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t.id;
            opt.textContent = t.name;
            select.appendChild(opt);
        });

        // Restore previous selection if still valid
        if (Array.from(select.options).some(o => o.value === currentValue)) {
            select.value = currentValue;
        }

        // Update the template info card
        this.updateTemplateInfo();
        select.addEventListener('change', () => this.updateTemplateInfo());
    },

    updateTemplateInfo() {
        const select = document.getElementById('gen-source-select');
        const infoEl = document.getElementById('gen-template-info');
        if (!infoEl) return;

        const tpl = DocumentStore.uploadedTemplates.find(t => t.id === select.value);
        if (tpl) {
            infoEl.innerHTML = `
                <div class="template-info-row"><span>File:</span><strong>${tpl.name}</strong></div>
                <div class="template-info-row"><span>Size:</span><span>${tpl.size}</span></div>
                <div class="template-info-row"><span>Pages:</span><span>${tpl.pages}</span></div>
                <div class="template-info-row"><span>Uploaded:</span><span>${tpl.uploadedAt.toLocaleDateString()}</span></div>
                <div class="template-info-row"><span>Fields:</span><span>${DocumentStore.ocrData.length} OCR labels</span></div>
            `;
            infoEl.style.display = 'block';
        } else {
            infoEl.style.display = 'none';
        }
    },

    onDataSourceChange() {
        const source = document.getElementById('gen-data-source').value;
        const csvArea = document.getElementById('gen-csv-upload-area');
        if (csvArea) {
            csvArea.style.display = (source === 'CSV Upload' || source === 'Excel Upload') ? 'block' : 'none';
        }
    },

    csvData: null,

    onCsvSelected(e) {
        const file = e.target.files[0];
        if (!file) return;

        const ext = file.name.split('.').pop().toLowerCase();
        
        if (ext === 'xlsx' || ext === 'xls') {
            const reader = new FileReader();
            reader.onload = (ev) => {
                const data = new Uint8Array(ev.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                
                if (json.length > 0) {
                    const headers = Object.keys(json[0]);
                    this.csvData = { headers, rows: json };
                    const label = document.getElementById('gen-csv-label');
                    if (label) label.textContent = `✓ ${file.name} — ${json.length} rows loaded`;
                    this._syncOcrDataFromCsv();
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const reader = new FileReader();
            reader.onload = (ev) => {
                const text = ev.target.result;
                const lines = text.split('\n').filter(l => l.trim());
                if (lines.length < 2) return;
                const headers = lines[0].split(',').map(h => h.trim());
                const rows = lines.slice(1).map(line => {
                    const vals = line.split(',').map(v => v.trim());
                    const row = {};
                    headers.forEach((h, i) => row[h] = vals[i] || '');
                    return row;
                });
                this.csvData = { headers, rows };
                const label = document.getElementById('gen-csv-label');
                if (label) label.textContent = `✓ ${file.name} — ${rows.length} rows loaded`;
                this._syncOcrDataFromCsv();
            };
            reader.readAsText(file);
        }
    },

    _syncOcrDataFromCsv() {
        if (!this.csvData) return;
        let newFields = [];
        if (this.csvData.headers.includes('field_code') && this.csvData.headers.includes('value')) {
             const uniqueCodes = [...new Set(this.csvData.rows.map(r => r.field_code).filter(Boolean))];
             newFields = uniqueCodes.map(code => ({ field: code, label: code, type: 'Alphanumeric' }));
        } else {
             newFields = this.csvData.headers.map(h => ({ field: h, label: h, type: 'Alphanumeric' }));
        }
        DocumentStore.ocrData = newFields;
        this.updateTemplateInfo();
    },

    startGeneration() {
        if (this.isGenerating) return;

        const templateSelect = document.getElementById('gen-source-select');
        const templateId = templateSelect.value;
        const template = DocumentStore.uploadedTemplates.find(t => t.id === templateId);

        if (!template) {
            alert('Please select a source template. Upload documents via Import & Config first.');
            return;
        }

        const count = Math.min(parseInt(document.getElementById('gen-count-input').value) || 10, 500);
        const format = document.getElementById('gen-format-select').value;
        const dataSource = document.getElementById('gen-data-source').value;
        const preserveFormat = document.getElementById('gen-preserve-format').checked;

        const progress = document.getElementById('generate-progress');
        const fill = document.getElementById('gen-progress-fill');
        const countEl = document.getElementById('gen-progress-count');
        const log = document.getElementById('gen-progress-log');
        const generateBtn = document.getElementById('generate-btn');

        // Reset
        this.isGenerating = true;
        generateBtn.disabled = true;
        generateBtn.innerHTML = '<span class="spinner"></span> Generating...';
        progress.classList.remove('hidden');
        fill.style.width = '0%';
        log.innerHTML = '';
        DocumentStore.clearGeneratedDocuments();

        // Log: starting
        this._log(log, `Starting generation: ${count} documents from "${template.name}"`);
        this._log(log, `Data source: ${dataSource} | Format: ${format} | Preserve format: ${preserveFormat}`);

        const fields = DocumentStore.ocrData;
        const generatedDocs = [];
        let current = 0;

        // Normalize CSV/Excel data to a list of row objects for generation
        let normalizedData = [];
        if ((dataSource === 'CSV Upload' || dataSource === 'Excel Upload') && this.csvData) {
            if (this.csvData.headers.includes('field_code') && this.csvData.headers.includes('value')) {
                // Long format: group by case_id
                const grouped = {};
                this.csvData.rows.forEach(r => {
                    const id = r.case_id || r.file_name || 'doc_1';
                    if (!grouped[id]) grouped[id] = {};
                    grouped[id][r.field_code] = {
                        value: r.value,
                        x: parseFloat(r.x),
                        y: parseFloat(r.y),
                        width: r.width,
                        height: r.height,
                        page: parseInt(r.page) || 1,
                        rotation: r.page_rotation
                    };
                });
                normalizedData = Object.values(grouped);
            } else {
                // Wide format
                normalizedData = this.csvData.rows;
            }
        }

        // Adjust count slightly if we have explicit data
        const actualCount = Math.min(count, Math.max(normalizedData.length, 1));

        const interval = setInterval(() => {
            current++;
            const pct = (current / actualCount) * 100;
            fill.style.width = pct + '%';
            countEl.textContent = `${current} / ${actualCount}`;

            // Generate actual document data
            const docData = {};
            const docMeta = {};
            fields.forEach(f => {
                if (normalizedData.length > 0) {
                    const row = normalizedData[(current - 1) % normalizedData.length];
                    const cell = row[f.label] || row[f.field];
                    if (cell !== undefined && typeof cell === 'object') {
                        docData[f.field] = cell.value !== undefined ? String(cell.value) : '';
                        docMeta[f.field] = cell;
                    } else if (cell !== undefined) {
                        docData[f.field] = String(cell);
                    } else {
                        docData[f.field] = DocumentStore.generateRandomValue(f.field);
                    }
                } else {
                    docData[f.field] = DocumentStore.generateRandomValue(f.field);
                }
            });

            const docName = `${template.name.replace(/\.[^.]+$/, '')}_gen_${String(current).padStart(3, '0')}.${format === 'PDF' ? 'pdf' : format === 'PNG Image' ? 'png' : 'docx'}`;

            const doc = {
                id: `doc_${Date.now()}_${current}`,
                name: docName,
                templateId: template.id,
                templateName: template.name,
                data: docData,
                meta: docMeta,
                format: format,
                generatedAt: new Date(),
                preserveFormat: preserveFormat,
            };
            generatedDocs.push(doc);

            this._log(log, `Generated ${docName}`, 'success');

            if (current >= actualCount) {
                clearInterval(interval);

                // Store all generated documents
                DocumentStore.addGeneratedDocuments(generatedDocs);

                this._log(log, `✓ All ${actualCount} documents generated successfully!`, 'success');
                this._log(log, `→ ${actualCount} documents ready for export. Go to Export page to download.`, 'info');

                // Show completion actions
                this._showCompletionActions(actualCount);

                this.isGenerating = false;
                generateBtn.disabled = false;
                generateBtn.innerHTML = `
                    <svg width="18" height="18" viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 2v14M2 9h14"/></svg>
                    Generate Documents
                `;
            }
        }, 200);
    },

    _log(container, message, type = 'msg') {
        const entry = document.createElement('div');
        entry.className = 'progress-log-entry';
        const time = new Date().toLocaleTimeString();
        entry.innerHTML = `<span class="time">[${time}]</span> <span class="${type}">${message}</span>`;
        container.appendChild(entry);
        container.scrollTop = container.scrollHeight;
    },

    _showCompletionActions(count) {
        const existing = document.getElementById('gen-completion-actions');
        if (existing) existing.remove();

        const progress = document.getElementById('generate-progress');
        const div = document.createElement('div');
        div.id = 'gen-completion-actions';
        div.className = 'gen-completion-actions';
        div.innerHTML = `
            <div class="gen-completion-summary">
                <div class="gen-completion-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="var(--color-success)" stroke-width="2">
                        <circle cx="12" cy="12" r="10"/>
                        <path d="M8 12l3 3 5-5"/>
                    </svg>
                </div>
                <div>
                    <strong>${count} documents</strong> generated and stored in memory.
                    <p style="font-size:11px;color:var(--color-text-muted);margin-top:2px;">Documents are ready for download in the Export page.</p>
                </div>
            </div>
            <div class="gen-completion-btns">
                <button class="btn btn--sm btn--outline" id="gen-download-csv">Download CSV Data</button>
                <button class="btn btn--sm btn--outline" id="gen-download-json">Download JSON</button>
                <button class="btn btn--sm btn--primary" id="gen-go-export">Go to Export →</button>
            </div>
        `;
        progress.after(div);

        document.getElementById('gen-download-csv').addEventListener('click', () => ExportPage.downloadCSV());
        document.getElementById('gen-download-json').addEventListener('click', () => ExportPage.downloadJSON());
        document.getElementById('gen-go-export').addEventListener('click', () => Navigation.navigateTo('export'));
    }
};

/* ═══════════════════════════════════════════════════════════════
   EXPORT PAGE
   ═══════════════════════════════════════════════════════════════ */
const ExportPage = {
    init() {
        // Export option selection
        document.querySelectorAll('.export-option-card').forEach(card => {
            card.addEventListener('click', () => {
                card.classList.toggle('selected');
            });
        });

        // Export All Formats — sequential staggered downloads
        document.getElementById('export-all-btn').addEventListener('click', () => {
            const docs = DocumentStore.generatedDocuments;
            if (docs.length === 0) {
                this._showNoDataMessage();
                return;
            }
            this._exportSequential(['csv', 'json', 'summary'], 'export-all-btn');
        });

        // Export Selected — only selected formats
        document.getElementById('export-selected-btn').addEventListener('click', () => {
            const docs = DocumentStore.generatedDocuments;
            if (docs.length === 0) {
                this._showNoDataMessage();
                return;
            }
            const selected = document.querySelectorAll('.export-option-card.selected');
            if (selected.length === 0) {
                alert('Please select at least one export format by clicking on a card.');
                return;
            }
            const formats = [];
            selected.forEach(card => {
                if (card.id === 'export-csv-option') formats.push('csv');
                if (card.id === 'export-json-option') formats.push('json');
                if (card.id === 'export-pdf-option') formats.push('pdf');
            });
            this._exportSequential(formats, 'export-selected-btn');
        });

        // Listen for new generated documents to update counts
        DocumentStore.on('documentsGenerated', () => this.updateExportCounts());
    },

    // Show a styled message when no documents are available
    _showNoDataMessage() {
        const existing = document.getElementById('export-status-panel');
        if (existing) existing.remove();

        const exportLayout = document.querySelector('#page-export .export-layout');
        const panel = document.createElement('div');
        panel.id = 'export-status-panel';
        panel.className = 'export-status-panel export-status--warning';
        panel.innerHTML = `
            <div class="export-status-icon">
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="var(--color-warning)" stroke-width="1.5">
                    <circle cx="10" cy="10" r="8"/><path d="M10 6v5M10 13.5v.5"/>
                </svg>
            </div>
            <div class="export-status-content">
                <strong>No documents available for export</strong>
                <p>Please go to the <a href="#" id="export-goto-gen" style="color:var(--color-accent);text-decoration:underline;">Generate page</a> and generate documents first.</p>
            </div>
        `;
        exportLayout.appendChild(panel);
        document.getElementById('export-goto-gen').addEventListener('click', (e) => {
            e.preventDefault();
            Navigation.navigateTo('generate');
        });
        // Auto-remove after 6s
        setTimeout(() => { if (panel.parentNode) panel.remove(); }, 6000);
    },

    // Sequential download with staggered timing to avoid browser blocking
    _exportSequential(formats, triggerBtnId) {
        const btn = document.getElementById(triggerBtnId);
        const origHTML = btn.innerHTML;
        btn.disabled = true;
        btn.innerHTML = '<span class="spinner spinner--dark"></span> Exporting...';

        // Remove old status panel
        const existing = document.getElementById('export-status-panel');
        if (existing) existing.remove();

        const results = [];
        let index = 0;

        const downloadNext = async () => {
            if (index >= formats.length) {
                // All done — show summary
                btn.disabled = false;
                btn.innerHTML = '✓ Exported!';
                btn.style.background = 'var(--color-success)';
                btn.style.borderColor = 'var(--color-success)';
                btn.style.color = 'white';
                setTimeout(() => {
                    btn.innerHTML = origHTML;
                    btn.style.background = '';
                    btn.style.borderColor = '';
                    btn.style.color = '';
                }, 2500);
                this._showExportStatusPanel(results);
                return;
            }

            const format = formats[index];
            let success = false;
            let filename = '';

            try {
                switch (format) {
                    case 'csv':
                        filename = this._buildCSV();
                        break;
                    case 'json':
                        filename = this._buildJSON();
                        break;
                    case 'summary':
                        filename = this._buildSummary();
                        break;
                    case 'pdf':
                        filename = await this._buildPDFs();
                        break;
                }
                success = true;
            } catch (e) {
                console.error('Export error:', e);
                success = false;
            }

            results.push({ format, filename, success });
            index++;

            // Stagger next download by 500ms to avoid browser blocking
            setTimeout(downloadNext, 500);
        };

        downloadNext();
    },

    _buildCSV() {
        const docs = DocumentStore.generatedDocuments;
        const fields = Object.keys(docs[0].data);
        const headers = ['document_name', 'template', 'format', 'generated_at', ...fields];
        const rows = docs.map(doc => {
            const base = [
                `"${doc.name}"`,
                `"${doc.templateName}"`,
                `"${doc.format}"`,
                `"${doc.generatedAt.toISOString()}"`,
            ];
            const fieldValues = fields.map(f => `"${(doc.data[f] || '').replace(/"/g, '""')}"`);
            return [...base, ...fieldValues].join(',');
        });
        const csv = [headers.join(','), ...rows].join('\n');
        const filename = 'ocr_extracted_data.csv';
        this._downloadFile(csv, filename, 'text/csv');
        return filename;
    },

    _buildJSON() {
        const docs = DocumentStore.generatedDocuments;
        const exportData = {
            exportedAt: new Date().toISOString(),
            totalDocuments: docs.length,
            schema: DocumentStore.ocrData.map(f => ({ field: f.field, label: f.label, type: f.type })),
            documents: docs.map(doc => ({
                name: doc.name,
                template: doc.templateName,
                format: doc.format,
                generatedAt: doc.generatedAt.toISOString(),
                preserveFormat: doc.preserveFormat,
                extractedData: doc.data,
            })),
        };
        const json = JSON.stringify(exportData, null, 2);
        const filename = 'ocr_structured_data.json';
        this._downloadFile(json, filename, 'application/json');
        return filename;
    },

    _buildSummary() {
        const docs = DocumentStore.generatedDocuments;
        let report = '═══════════════════════════════════════════════════════════\n';
        report += '  INSIGHT LEDGER — Document Generation Report\n';
        report += '═══════════════════════════════════════════════════════════\n\n';
        report += `Generated: ${new Date().toLocaleString()}\n`;
        report += `Total Documents: ${docs.length}\n`;
        report += `Template: ${docs[0]?.templateName || 'N/A'}\n`;
        report += `Format: ${docs[0]?.format || 'N/A'}\n\n`;
        report += '───────────────────────────────────────────────────────────\n\n';
        docs.forEach((doc, i) => {
            report += `Document ${i + 1}: ${doc.name}\n`;
            report += `  Generated: ${doc.generatedAt.toLocaleString()}\n`;
            Object.entries(doc.data).forEach(([key, val]) => {
                report += `  ${key}: ${val}\n`;
            });
            report += '\n';
        });
        const filename = 'generation_report.txt';
        this._downloadFile(report, filename, 'text/plain');
        return filename;
    },

    async _buildPDFs() {
        try {
            const docs = DocumentStore.generatedDocuments;
            if (!docs || docs.length === 0) return null;
            
            if (!window.PDFLib) {
                alert('PDF-lib is not currently loaded.');
                return null;
            }

            const { PDFDocument, rgb, StandardFonts } = window.PDFLib;
            const outPdf = await PDFDocument.create();
            
            if (window.fontkit) {
                outPdf.registerFontkit(window.fontkit);
            }
            
            let font;
            try {
                const fontUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf';
                const fontBuffer = await fetch(fontUrl).then(res => res.arrayBuffer());
                font = await outPdf.embedFont(fontBuffer, { subset: true });
            } catch(e) {
                console.warn('Custom Roboto font fetch failed, using fallback Helvetica:', e);
                try {
                    font = await outPdf.embedFont(StandardFonts?.Helvetica || 'Helvetica');
                } catch(err) {
                    console.warn('Fallback font embed failed:', err);
                }
            }

            for (let i = 0; i < docs.length; i++) {
                const doc = docs[i];
                const template = DocumentStore.uploadedTemplates.find(t => t.id === doc.templateId);
                
                let templatePdfDoc;
                if (template && template._file) {
                     const arrayBuffer = await new Promise((resolve, reject) => {
                         const r = new FileReader();
                         r.onload = e => resolve(e.target.result);
                         r.onerror = e => reject(new Error('FileReader Error'));
                         r.readAsArrayBuffer(template._file);
                     });
                     templatePdfDoc = await PDFDocument.load(arrayBuffer);
                } else {
                     templatePdfDoc = await PDFDocument.create();
                     templatePdfDoc.addPage([595.28, 841.89]); // A4 fallback
                }
                
                const copiedPages = await outPdf.copyPages(templatePdfDoc, templatePdfDoc.getPageIndices());
                const startIdx = outPdf.getPageCount();
                copiedPages.forEach(p => outPdf.addPage(p));
                
                const allPages = outPdf.getPages();
                
                Object.keys(doc.data || {}).forEach(field => {
                     const meta = doc.meta && doc.meta[field];
                     const text = String(doc.data[field] || '');
                     
                     if (meta && typeof meta.x === 'number' && !isNaN(meta.x) && typeof meta.y === 'number' && !isNaN(meta.y)) {
                          const pageOffset = Math.max(0, (parseInt(meta.page) || 1) - 1);
                          const targetIdx = startIdx + pageOffset;
                          
                          if (targetIdx < allPages.length && font) {
                               const page = allPages[targetIdx];
                               const { height } = page.getSize();
                               
                               page.drawText(text, {
                                   x: meta.x,
                                   y: height - meta.y, 
                                   font: font,
                                   size: 11,
                                   color: rgb(0.8, 0, 0)
                               });
                          }
                     }
                });
            }

            const pdfBytes = await outPdf.save();
            const blob = new Blob([pdfBytes], { type: "application/pdf" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'generated_documents_batch_custom.pdf';
            a.click();
            
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            return 'generated_documents_batch_custom.pdf';
        } catch (e) {
            alert("Error generating PDF: " + e.message + "\n" + (e.stack ? e.stack.substring(0, 300) : ""));
            throw e;
        }
    },

    updateExportCounts() {
        const docs = DocumentStore.generatedDocuments;
        const pdfCount = document.querySelector('#export-pdf-option .export-option-count');
        const csvCount = document.querySelector('#export-csv-option .export-option-count');
        const jsonCount = document.querySelector('#export-json-option .export-option-count');

        if (pdfCount) pdfCount.textContent = `${docs.length} documents ready`;
        if (csvCount) csvCount.textContent = `${docs.length * DocumentStore.ocrData.length} records`;
        if (jsonCount) jsonCount.textContent = docs.length > 0 ? 'Full schema' : 'No data';
    },

    // Keep old API for Generate page direct download buttons
    downloadCSV() {
        const docs = DocumentStore.generatedDocuments;
        if (docs.length === 0) { alert('No documents generated yet.'); return; }
        this._buildCSV();
    },

    downloadJSON() {
        const docs = DocumentStore.generatedDocuments;
        if (docs.length === 0) { alert('No documents generated yet.'); return; }
        this._buildJSON();
    },

    downloadSummary() {
        const docs = DocumentStore.generatedDocuments;
        if (docs.length === 0) { alert('No documents generated yet.'); return; }
        this._buildSummary();
    },

    _downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        // Delay cleanup — give browser time to process the download
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 1000);
    },

    _showExportStatusPanel(results) {
        const existing = document.getElementById('export-status-panel');
        if (existing) existing.remove();

        const exportLayout = document.querySelector('#page-export .export-layout');
        const panel = document.createElement('div');
        panel.id = 'export-status-panel';
        panel.className = 'export-status-panel export-status--success';

        const formatLabels = { csv: 'CSV Data', json: 'Structured JSON', pdf: 'PDF Document', summary: 'Report (TXT)' };
        const formatIcons = { csv: '📊', json: '📋', pdf: '📄', summary: '📄' };

        const rows = results.map(r => `
            <div class="export-status-row">
                <span class="export-status-format">${formatIcons[r.format]} ${formatLabels[r.format]}</span>
                <span class="export-status-file">${r.filename}</span>
                <span class="export-status-badge ${r.success ? 'badge--success' : 'badge--error'}">${r.success ? '✓ Downloaded' : '✗ Failed'}</span>
                <button class="btn btn--sm btn--outline export-retry-btn" data-format="${r.format}">Re-download</button>
            </div>
        `).join('');

        panel.innerHTML = `
            <div class="export-status-header">
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="var(--color-success)" stroke-width="2">
                    <circle cx="10" cy="10" r="8"/><path d="M7 10l2 2 4-4"/>
                </svg>
                <strong>${results.length} file(s) exported successfully</strong>
                <button class="export-status-close" id="export-status-close">&times;</button>
            </div>
            <div class="export-status-list">${rows}</div>
            <p class="export-status-hint">Files are saved to your browser's default download folder.</p>
        `;
        exportLayout.appendChild(panel);

        // Close button
        document.getElementById('export-status-close').addEventListener('click', () => panel.remove());

        // Re-download buttons
        panel.querySelectorAll('.export-retry-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const fmt = btn.dataset.format;
                try {
                    switch (fmt) {
                        case 'csv': this._buildCSV(); break;
                        case 'json': this._buildJSON(); break;
                        case 'summary': this._buildSummary(); break;
                    }
                    btn.textContent = '✓ Done';
                    btn.disabled = true;
                    setTimeout(() => { btn.textContent = 'Re-download'; btn.disabled = false; }, 1500);
                } catch (e) { console.error(e); }
            });
        });
    },

    showExportFeedback() {
        const btn = document.getElementById('export-all-btn');
        const origHTML = btn.innerHTML;
        btn.innerHTML = '✓ Exported!';
        btn.style.background = 'var(--color-success)';
        btn.style.borderColor = 'var(--color-success)';
        setTimeout(() => {
            btn.innerHTML = origHTML;
            btn.style.background = '';
            btn.style.borderColor = '';
        }, 2500);
    }
};

/* ═══════════════════════════════════════════════════════════════
   RESIZE HANDLER
   ═══════════════════════════════════════════════════════════════ */
window.addEventListener('resize', () => {
    Dashboard.renderChart();
});
