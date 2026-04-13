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
        // Seed OCR data from preview entities
        this.ocrData = [
            { field: 'company_name', label: 'Company Name', value: 'ACME Corporation Ltd.', type: 'Alphanumeric' },
            { field: 'invoice_number', label: 'Invoice Number', value: 'INV-2026-04-0042', type: 'Alphanumeric' },
            { field: 'invoice_date', label: 'Invoice Date', value: '2026-04-10', type: 'Date' },
            { field: 'due_date', label: 'Due Date', value: '2026-05-10', type: 'Date' },
            { field: 'client_name', label: 'Client Name', value: 'NextGen Digital Solutions', type: 'Alphanumeric' },
            { field: 'po_number', label: 'PO Number', value: 'PO-78432', type: 'Alphanumeric' },
            { field: 'subtotal', label: 'Subtotal', value: '$12,450.00', type: 'Currency' },
            { field: 'tax', label: 'Tax Amount', value: '$1,120.50', type: 'Currency' },
            { field: 'total', label: 'Total Amount', value: '$13,570.50', type: 'Currency' },
        ];
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
        const rndDate = () => { const d = new Date(2026, rndInt(0,11), rndInt(1,28)); return d.toISOString().split('T')[0]; };

        switch (field) {
            case 'company_name': return companies[rndInt(0, companies.length - 1)];
            case 'client_name': return clients[rndInt(0, clients.length - 1)];
            case 'invoice_number': return `INV-2026-${String(rndInt(1,12)).padStart(2,'0')}-${String(rndInt(1,9999)).padStart(4,'0')}`;
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

        // Update review summary on step 4
        if (step === 4) {
            this.updateReview();
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

    updateReview() {
        document.getElementById('review-files').textContent = `${Math.max(this.files.length, 3)} documents`;
        const precision = document.querySelector('input[name="precision"]:checked');
        document.getElementById('review-precision').textContent = precision ? precision.value.charAt(0).toUpperCase() + precision.value.slice(1) : 'Balanced';
        document.getElementById('review-handwriting').textContent = document.getElementById('handwriting-toggle').checked ? 'Enabled' : 'Disabled';
        document.getElementById('review-autovalidate').textContent = document.getElementById('auto-validate-toggle').checked ? 'Enabled' : 'Disabled';
        document.getElementById('review-language').textContent = document.getElementById('language-select').value;
        document.getElementById('review-fields').textContent = `${this.mappingData.length} fields`;
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
    entities: [
        { field: 'company_name', label: 'Company Name', value: 'ACME Corporation Ltd.', confidence: 98.5, level: 'high' },
        { field: 'company_address', label: 'Company Address', value: '123 Business Ave, Suite 400, San Francisco, CA 94102', confidence: 96.2, level: 'high' },
        { field: 'invoice_title', label: 'Document Type', value: 'INVOICE', confidence: 99.8, level: 'high' },
        { field: 'invoice_number', label: 'Invoice Number', value: 'INV-2026-04-0042', confidence: 97.4, level: 'high' },
        { field: 'invoice_date', label: 'Invoice Date', value: '2026-04-10', confidence: 95.1, level: 'high' },
        { field: 'due_date', label: 'Due Date', value: '2026-05-10', confidence: 82.3, level: 'medium' },
        { field: 'client_name', label: 'Client Name', value: 'NextGen Digital Solutions', confidence: 93.7, level: 'high' },
        { field: 'po_number', label: 'PO Number', value: 'PO-78432', confidence: 65.8, level: 'low' },
        { field: 'subtotal', label: 'Subtotal', value: '$12,450.00', confidence: 98.1, level: 'high' },
        { field: 'tax', label: 'Tax Amount', value: '$1,120.50', confidence: 86.4, level: 'medium' },
        { field: 'total', label: 'Total Amount', value: '$13,570.50', confidence: 99.2, level: 'high' },
    ],

    zoomLevel: 100,

    init() {
        this.renderEntities();
        this.initBBoxInteraction();
        this.initZoomControls();
        this.initResizer();
        this.initSuggestionBanner();
    },

    renderEntities() {
        const list = document.getElementById('entities-list');
        list.innerHTML = this.entities.map(e => {
            const inputClass = e.level === 'low' ? 'entity-input entity-input--error' : 'entity-input';
            return `
                <div class="entity-item" data-field="${e.field}">
                    <div class="entity-header">
                        <span class="entity-label">${e.label}</span>
                        <span class="entity-confidence entity-confidence--${e.level}">${e.confidence}%</span>
                    </div>
                    <input type="text" class="${inputClass}" value="${e.value}" aria-label="${e.label} value">
                </div>
            `;
        }).join('');

        // Highlight bbox on entity hover
        list.querySelectorAll('.entity-item').forEach(item => {
            item.addEventListener('mouseenter', () => {
                const field = item.dataset.field;
                document.querySelectorAll('.ocr-bbox').forEach(b => b.classList.remove('active'));
                const bbox = document.querySelector(`.ocr-bbox[data-field="${field}"]`);
                if (bbox) bbox.classList.add('active');
            });
            item.addEventListener('mouseleave', () => {
                document.querySelectorAll('.ocr-bbox').forEach(b => b.classList.remove('active'));
            });
        });
    },

    initBBoxInteraction() {
        document.querySelectorAll('.ocr-bbox').forEach(bbox => {
            bbox.addEventListener('click', () => {
                const field = bbox.dataset.field;
                const entityItem = document.querySelector(`.entity-item[data-field="${field}"]`);
                if (entityItem) {
                    entityItem.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    const input = entityItem.querySelector('.entity-input');
                    if (input) input.focus();
                }
            });
        });
    },

    initZoomControls() {
        const zoomIn = document.getElementById('ocr-zoom-in');
        const zoomOut = document.getElementById('ocr-zoom-out');
        const zoomLevel = document.getElementById('ocr-zoom-level');
        const docSim = document.getElementById('ocr-document-sim');

        zoomIn.addEventListener('click', () => {
            if (this.zoomLevel < 200) {
                this.zoomLevel += 25;
                docSim.style.transform = `scale(${this.zoomLevel / 100})`;
                docSim.style.transformOrigin = 'top center';
                zoomLevel.textContent = this.zoomLevel + '%';
            }
        });

        zoomOut.addEventListener('click', () => {
            if (this.zoomLevel > 50) {
                this.zoomLevel -= 25;
                docSim.style.transform = `scale(${this.zoomLevel / 100})`;
                docSim.style.transformOrigin = 'top center';
                zoomLevel.textContent = this.zoomLevel + '%';
            }
        });
    },

    initResizer() {
        const resizer = document.getElementById('ocr-resizer');
        const leftPanel = document.getElementById('ocr-document-panel');
        const container = document.getElementById('ocr-split-view');
        let isResizing = false;

        resizer.addEventListener('mousedown', (e) => {
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

    initSuggestionBanner() {
        const dismissBtn = document.getElementById('dismiss-template-btn');
        const banner = document.getElementById('ai-suggestion-banner');
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
                    <strong style="color:var(--color-success)">Template Applied!</strong> "Standard Invoice" template has been applied successfully.
                </div>
            `;
            setTimeout(() => {
                banner.style.opacity = '0';
                setTimeout(() => banner.style.display = 'none', 300);
            }, 3000);
        });
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
            csvArea.style.display = source === 'CSV Upload' ? 'block' : 'none';
        }
    },

    csvData: null,

    onCsvSelected(e) {
        const file = e.target.files[0];
        if (!file) return;
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
        };
        reader.readAsText(file);
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

        const interval = setInterval(() => {
            current++;
            const pct = (current / count) * 100;
            fill.style.width = pct + '%';
            countEl.textContent = `${current} / ${count}`;

            // Generate actual document data
            const docData = {};
            fields.forEach(f => {
                if (dataSource === 'CSV Upload' && this.csvData) {
                    const row = this.csvData.rows[(current - 1) % this.csvData.rows.length];
                    docData[f.field] = row[f.label] || row[f.field] || DocumentStore.generateRandomValue(f.field);
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
                format: format,
                generatedAt: new Date(),
                preserveFormat: preserveFormat,
            };
            generatedDocs.push(doc);

            this._log(log, `Generated ${docName}`, 'success');

            if (current >= count) {
                clearInterval(interval);

                // Store all generated documents
                DocumentStore.addGeneratedDocuments(generatedDocs);

                this._log(log, `✓ All ${count} documents generated successfully!`, 'success');
                this._log(log, `→ ${count} documents ready for export. Go to Export page to download.`, 'info');

                // Show completion actions
                this._showCompletionActions(count);

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
                if (card.id === 'export-pdf-option') formats.push('summary');
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

        const downloadNext = () => {
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
        report +=    '  INSIGHT LEDGER — Document Generation Report\n';
        report +=    '═══════════════════════════════════════════════════════════\n\n';
        report +=    `Generated: ${new Date().toLocaleString()}\n`;
        report +=    `Total Documents: ${docs.length}\n`;
        report +=    `Template: ${docs[0]?.templateName || 'N/A'}\n`;
        report +=    `Format: ${docs[0]?.format || 'N/A'}\n\n`;
        report +=    '───────────────────────────────────────────────────────────\n\n';
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

        const formatLabels = { csv: 'CSV Data', json: 'Structured JSON', summary: 'Report (TXT)' };
        const formatIcons = { csv: '📊', json: '📋', summary: '📄' };

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
