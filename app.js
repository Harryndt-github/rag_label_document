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
    generateRandomValue(fieldCode) {
        const rndInt = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
        const rndAmount = (min, max) => (Math.random() * (max - min) + min).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        const rndDate = () => { const d = new Date(2026, rndInt(0, 11), rndInt(1, 28)); return d.toISOString().split('T')[0]; };
        const rndStr = (arr) => arr[rndInt(0, arr.length - 1)];

        // --- Data pools ---
        const companyNames = [
            'NINGBO MING YUE GLOBAL TRADING CO., LTD', 'THIEN LONG GROUP CORPORATION',
            'SHANGHAI BRIGHT IMPORT EXPORT CO., LTD', 'KOREA TRADING INTERNATIONAL',
            'SINGAPORE COMMODITIES PTE LTD', 'TOKYO ELECTRONICS JAPAN CO., LTD',
            'VIETNAM NATIONAL TEXTILE & GARMENT GROUP', 'SAIGON TRADING CORPORATION',
            'BINH DUONG INDUSTRIAL JOINT STOCK COMPANY', 'DONG NAI PACKAGING CO., LTD',
            'FORMOSA HA TINH STEEL CORPORATION', 'SAMSUNG ELECTRONICS VIETNAM',
            'FOXCONN INDUSTRIAL INTERNET CO., LTD', 'MITSUI & CO. VIETNAM LTD'
        ];
        const personNames = [
            'Nguyen Van Minh', 'Tran Thi Lan', 'Le Hoang Nam', 'Pham Quoc Tuan',
            'Hoang Duc Anh', 'Vo Thi Mai', 'Bui Thanh Long', 'Dang Van Hung',
            'Do Minh Hieu', 'Ngo Thi Huong', 'Ly Van Phong', 'Truong Quang Hai'
        ];
        const addresses = [
            'No.11-3-3, DongWei Building, No.399, Minghe Road, Ningbo, China',
            '200 Nguyen Hue, P.Ben Nghe, Q.1, TP.HCM',
            '45 Le Duan, P.Ben Nghe, Q.1, TP.HCM, Vietnam',
            'KCN Song Than, Phuong Di An, TP.Di An, Binh Duong',
            '15 Tran Hung Dao, P.Pham Ngu Lao, Q.1, TP.HCM',
            '789 Vo Van Kiet, P.1, Q.5, TP.HCM',
            'Lot A5, KCN Bien Hoa 2, Dong Nai, Vietnam',
            '26F, Shui On Centre, 6-8 Harbour Road, Wanchai, Hong Kong'
        ];
        const docNames = [
            'Sales Contract', 'Credit Note', 'Debit Note', 'Commercial Invoice',
            'Proforma Invoice', 'Packing List', 'Bill of Lading', 'Certificate of Origin',
            'Insurance Certificate', 'Inspection Certificate', 'Customs Declaration'
        ];
        const goodsDesc = [
            'Shoe materials: MSKB9506-L2, Brand: IMS, 100% new, origin from China',
            'Electronic components: PCB boards, IC chips, capacitors',
            'Textile fabric: 100% cotton, 60 inch width, white bleached',
            'Steel coils: HRC SS400, thickness 2.0mm, width 1219mm',
            'Plastic resin: PP H030SG, pellet form, prime grade',
            'Chemical: Sodium Hydroxide (NaOH) 99%, industrial grade',
            'Machinery parts: CNC spindle motor, servo drives, ball screws',
            'Packaging materials: corrugated carton boxes, kraft paper rolls'
        ];
        const banks = [
            'CHINA MINSHENG BANK, NINGBO BRANCH',
            'VIETCOMBANK, HO CHI MINH CITY BRANCH',
            'BIDV, BINH DUONG BRANCH',
            'HSBC VIETNAM, DISTRICT 1 BRANCH',
            'STANDARD CHARTERED BANK VIETNAM'
        ];
        const cities = ['HCM', 'Ha Noi', 'Da Nang', 'Can Tho', 'Hai Phong', 'Bien Hoa', 'Vung Tau', 'Qui Nhon'];
        const countries = ['Vietnam', 'China', 'Japan', 'Korea', 'USA', 'Singapore', 'Thailand', 'Germany'];
        const ports = ['Cat Lai, HCM', 'Hai Phong', 'Da Nang', 'Cai Mep', 'NINGBO, CHINA', 'BUSAN, KOREA', 'YOKOHAMA, JAPAN', 'SHANGHAI, CHINA'];
        const vessels = ['MV OCEAN STAR', 'MV PACIFIC GLORY', 'MV MEKONG TRADER', 'MV SAIGON EXPRESS', 'MV DRAGON PEARL', 'MV BLUE HORIZON'];
        const currencies = ['USD', 'VND', 'EUR', 'JPY', 'CNY', 'KRW'];
        const packingTypes = ['SETS', 'CTNS', 'PKGS', 'ROLLS', 'BAGS', 'PALLETS', 'DRUMS'];
        const paymentTerms = ['T/T IN ADVANCE', 'L/C AT SIGHT', 'D/P 60 DAYS', 'T/T 30 DAYS AFTER B/L DATE', 'CASH AGAINST DOCUMENTS'];
        const incoterms = ['FOB HO CHI MINH', 'CIF NINGBO', 'CNF BUSAN', 'CFR YOKOHAMA', 'EXW BINH DUONG', 'DDP HO CHI MINH'];
        const shipmentTerms = ['Not later than December 2025', 'Not later than March 2026', 'Within 30 days after L/C date', 'Prompt shipment', 'As per buyer schedule'];

        // --- Strip common document-type prefixes before matching ---
        // e.g., CRN_TEN_NGUOI_BAN → TEN_NGUOI_BAN, INV_SO_HOA_DON → SO_HOA_DON
        let fc = fieldCode.toLowerCase().replace(/[\s\-]/g, '_');
        fc = fc.replace(/^(crn|inv|bl|ci|pi|co|pl|pk|dn|cn|dc|ins|cert|cust|tkhq|sc|po)_/, '');

        // ═══════════════════════════════════════════════════
        //  Vietnamese field_code patterns (priority first)
        // ═══════════════════════════════════════════════════

        // TEN_CHUNG_TU / TEN_TAI_LIEU = Document name
        if (fc.includes('ten_chung_tu') || fc.includes('ten_tai_lieu') || fc.includes('loai_chung_tu'))
            return rndStr(docNames);
        // TEN_NGUOI_BAN / TEN_NCC / TEN_CONG_TY_BAN = Seller/Supplier name
        if (fc.includes('ten_nguoi_ban') || fc.includes('ten_ncc') || fc.includes('ten_cong_ty_ban') || fc.includes('nguoi_ban') || fc.includes('ben_ban'))
            return rndStr(companyNames);
        // TEN_NGUOI_MUA / TEN_CONG_TY_MUA = Buyer name
        if (fc.includes('ten_nguoi_mua') || fc.includes('nguoi_mua') || fc.includes('ben_mua') || fc.includes('ten_cong_ty_mua'))
            return rndStr(companyNames);
        // TEN = general name (person or company)
        if (fc.includes('ten_') && (fc.includes('nguoi') || fc.includes('dai_dien')))
            return rndStr(personNames);
        if (fc === 'ten' || fc.includes('ten_cong_ty') || fc.includes('ten_don_vi'))
            return rndStr(companyNames);

        // DIA_CHI = Address
        if (fc.includes('dia_chi'))
            return rndStr(addresses);
        // SO = Number (number-type fields)
        if (fc.includes('so_credit') || fc.includes('so_cn') || fc.includes('so_dn'))
            return `CN-${rndStr(['MY','TL','SG','BD'])}${rndInt(100000,999999)}`;
        if (fc.includes('so_hoa_don') || fc.includes('so_invoice'))
            return `INV-${rndStr(['MY','TL','SG','BD'])}${rndInt(100000,999999)}`;
        if (fc.includes('so_hop_dong') || fc.includes('so_contract') || fc.includes('so_sc'))
            return `SC-${rndStr(['MY','TL','SG','BD'])}${rndInt(100000,999999)}`;
        if (fc.includes('so_to_khai') || fc.includes('so_tkhq'))
            return `${rndInt(300,399)}${String(rndInt(1,99)).padStart(3,'0')}/${rndStr(['HCM','HAN','HPH'])}/${rndInt(2025,2026)}`;
        if (fc.includes('so_van_don') || fc.includes('so_bl') || fc.includes('so_bill'))
            return `COSU${rndInt(1000000,9999999)}`;
        if (fc.includes('so_container'))
            return `${rndStr(['MSCU','HLXU','CMAU','OOLU','TGHU'])}${rndInt(1000000,9999999)}`;
        if (fc.includes('so_seal') || fc.includes('seal'))
            return `SL${rndInt(100000,999999)}`;
        if (fc.includes('so_') || fc.includes('ma_'))
            return `${fieldCode.replace(/^(CRN|INV|BL|CI)_/i,'').substring(0,4).toUpperCase()}-${rndInt(100000,999999)}`;

        // NGAY = Date
        if (fc.includes('ngay'))
            return rndDate();
        // LOAI_TIEN / DONG_TIEN = Currency
        if (fc.includes('loai_tien') || fc.includes('dong_tien') || fc.includes('tien_te'))
            return rndStr(currencies);
        // DON_GIA = Unit price
        if (fc.includes('don_gia'))
            return rndAmount(10, 100000);
        // THANH_TIEN / TONG_TIEN / SO_TIEN = Amount
        if (fc.includes('thanh_tien') || fc.includes('tong_tien') || fc.includes('so_tien') || fc.includes('tong_gia_tri'))
            return rndAmount(1000, 500000);
        // SO_LUONG / SL = Quantity
        if (fc.includes('so_luong') || fc === 'sl')
            return String(rndInt(1, 10000));
        // DON_VI_TINH / DVT = Unit
        if (fc.includes('don_vi') || fc.includes('dvt'))
            return rndStr(packingTypes);
        // DIEU_KIEN_GIAO_HANG = Delivery/Incoterm
        if (fc.includes('dieu_kien') || fc.includes('incoterm'))
            return rndStr(incoterms);
        // PHUONG_THUC_THANH_TOAN = Payment terms
        if (fc.includes('thanh_toan') || fc.includes('phuong_thuc'))
            return rndStr(paymentTerms);
        // CANG = Port
        if (fc.includes('cang_xuat') || fc.includes('cang_di') || fc.includes('loading'))
            return rndStr(['NINGBO, CHINA', 'SHANGHAI, CHINA', 'BUSAN, KOREA', 'YOKOHAMA, JAPAN']);
        if (fc.includes('cang_nhap') || fc.includes('cang_den') || fc.includes('discharge'))
            return rndStr(['CAT LAI, HCM', 'HAI PHONG', 'DA NANG', 'CAI MEP']);
        if (fc.includes('cang'))
            return rndStr(ports);
        // QUOC_GIA / NUOC = Country
        if (fc.includes('quoc_gia') || fc.includes('nuoc') || fc.includes('xuat_xu'))
            return rndStr(countries);
        // MO_TA_HANG / TEN_HANG = Goods description
        if (fc.includes('mo_ta') || fc.includes('ten_hang') || fc.includes('hang_hoa') || fc.includes('goods'))
            return rndStr(goodsDesc);
        // TAU / PHUONG_TIEN = Vessel
        if (fc.includes('tau') || fc.includes('phuong_tien') || fc.includes('vessel'))
            return rndStr(vessels);
        if (fc.includes('chuyen') || fc.includes('voyage'))
            return `V.${rndInt(100,999)}${rndStr(['E','W','N','S'])}`;
        // TRONG_LUONG = Weight
        if (fc.includes('trong_luong') || fc.includes('weight') || fc.includes('gross') || fc.includes('net'))
            return `${rndAmount(100, 50000)} KG`;
        // NGAN_HANG / BANK
        if (fc.includes('ngan_hang') || fc.includes('bank'))
            return rndStr(banks);
        // TAI_KHOAN / ACCOUNT
        if (fc.includes('tai_khoan') || fc.includes('account'))
            return String(rndInt(100000000, 999999999));
        // SWIFT
        if (fc.includes('swift'))
            return rndStr(['MSBCCNBJ019', 'BFTVVNVX', 'BIDVVNVX', 'HSBCVNVX']);
        // THOI_HAN_GIAO = Shipment terms
        if (fc.includes('thoi_han') || fc.includes('giao_hang') || fc.includes('shipment'))
            return rndStr(shipmentTerms);
        // DONG_GOI / PACKING
        if (fc.includes('dong_goi') || fc.includes('packing'))
            return rndStr(['Export standard packing', 'Carton box on pallet', 'Wooden case', 'Bulk in container']);
        // GHI_CHU / NOTE
        if (fc.includes('ghi_chu') || fc.includes('note') || fc.includes('remark'))
            return rndStr(['As per contract terms', 'Subject to final confirmation', 'Original documents required', 'N/A']);
        // THUE / TAX / VAT
        if (fc.includes('thue') || fc.includes('tax') || fc.includes('vat'))
            return rndAmount(100, 50000);

        // ═══════════════════════════════════════════════════
        //  English field_code patterns (fallback)
        // ═══════════════════════════════════════════════════
        if (fc.includes('company') || fc.includes('shipper') || fc.includes('consignee') || fc.includes('exporter') || fc.includes('importer') || fc.includes('buyer') || fc.includes('seller'))
            return rndStr(companyNames);
        if (fc.includes('name') || fc.includes('contact') || fc.includes('person'))
            return rndStr(personNames);
        if (fc.includes('address'))
            return rndStr(addresses);
        if (fc.includes('invoice') || fc.includes('credit') || fc.includes('debit'))
            return `${fc.substring(0,3).toUpperCase()}-${rndStr(['MY','TL','SG'])}${rndInt(100000,999999)}`;
        if (fc.includes('date') || fc.includes('issued') || fc.includes('created') || fc.includes('etd') || fc.includes('eta'))
            return rndDate();
        if (fc.includes('amount') || fc.includes('total') || fc.includes('subtotal') || fc.includes('price') || fc.includes('value') || fc.includes('cost'))
            return rndAmount(1000, 500000);
        if (fc.includes('quantity') || fc.includes('qty'))
            return String(rndInt(1, 10000));
        if (fc.includes('currency'))
            return rndStr(currencies);
        if (fc.includes('country') || fc.includes('origin') || fc.includes('destination'))
            return rndStr(countries);
        if (fc.includes('port'))
            return rndStr(ports);
        if (fc.includes('product') || fc.includes('description') || fc.includes('commodity'))
            return rndStr(goodsDesc);
        if (fc.includes('payment'))
            return rndStr(paymentTerms);
        if (fc.includes('term'))
            return rndStr(incoterms);
        if (fc.includes('unit'))
            return rndStr(packingTypes);

        // Additional field codes from smart detector
        if (fc.includes('nguoi_thu_huong') || fc.includes('beneficiary'))
            return rndStr(companyNames);
        if (fc.includes('giao_hang_tung_phan') || fc.includes('partial'))
            return rndStr(['Allowed', 'Not Allowed']);
        if (fc.includes('chuyen_tai') || fc.includes('transshipment'))
            return rndStr(['Allowed', 'Not Allowed']);
        if (fc.includes('ky_ma_hieu') || fc.includes('marking'))
            return rndStr(['Export standard packing', 'As per contract', 'Neutral marking', 'Brand label']);
        if (fc.includes('bao_hiem') || fc.includes('insurance'))
            return rndStr(['By Buyer', 'By Seller', 'CIF terms', '110% invoice value']);
        if (fc.includes('ma_hang') || fc.includes('model'))
            return `${rndStr(['IMSKR','MFG','MDL','SKU'])}${rndInt(1000,9999)}-${rndStr(['L2','H1','A3','B5'])}`;
        if (fc.includes('thuong_hieu') || fc.includes('brand'))
            return rndStr(['IMS', 'Samsung', 'LG', 'Sony', 'Panasonic', 'Mitsubishi', 'Bosch']);
        if (fc.includes('dien_thoai') || fc.includes('phone') || fc.includes('tel'))
            return `0${rndInt(2,9)}${rndInt(0,9)}${rndInt(0,9)}-${rndInt(100,999)}-${rndInt(1000,9999)}`;
        if (fc.includes('fax'))
            return `0${rndInt(2,9)}${rndInt(0,9)}${rndInt(0,9)}-${rndInt(100,999)}-${rndInt(1000,9999)}`;
        if (fc.includes('so_chung_tu') || fc.includes('document_number') || fc.includes('doc_no'))
            return `${rndStr(['MY-TL','SC','HD'])}${rndInt(100000,999999)}`;

        // ═══════════════════════════════════════════════════
        //  Absolute fallback — use cleaned field name
        // ═══════════════════════════════════════════════════
        return `${fieldCode.replace(/^(CRN|INV|BL|CI|PI|CO|PL)_/i,'').substring(0,6).toUpperCase()}-${rndInt(1000, 9999)}`;
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
const defaultLabels = [
    { name: 'Số Document/ Invoice No', type: 'Alphanumeric', threshold: 95, tag: 'INV_SO_HOA_DON', status: 'Active' },
    { name: 'Số chứng từ/ PACKING LIST', type: 'Alphanumeric', threshold: 95, tag: 'PKL_SO_CHUNG_TU', status: 'Active' },
    { name: 'Ngày chứng từ/ Date', type: 'Date', threshold: 92, tag: 'INV_NGAY_LAP', status: 'Active' },
    { name: 'Ngày hợp đồng', type: 'Date', threshold: 92, tag: 'CON_NGAY_HOP_DONG', status: 'Active' },
    { name: 'Ngày phát hành/ Date of Issue', type: 'Date', threshold: 88, tag: 'LOC_NGAY_PHAT_HANH', status: 'Active' },
    { name: 'Người mua/ For account & risk of messrs', type: 'Alphanumeric', threshold: 90, tag: 'INV_B_NGUOI_MUA_TEN', status: 'Active' },
    { name: 'Người bán/ Seller', type: 'Alphanumeric', threshold: 90, tag: 'INV_B_NGUOI_BAN_TEN', status: 'Active' },
    { name: 'Thành tiền/ Amount', type: 'Currency', threshold: 97, tag: 'INV_C_THANH_TIEN', status: 'Active' },
    { name: 'Tên hàng hóa', type: 'Text Block', threshold: 82, tag: 'INV_C_TEN_HANG', status: 'Active' },
    { name: 'Số lượng/ Quantity', type: 'Numeric', threshold: 88, tag: 'INV_C_SO_LUONG', status: 'Active' },
    { name: 'Đơn giá', type: 'Currency', threshold: 93, tag: 'INV_C_DON_GIA', status: 'Warning' },
    { name: 'Địa Chỉ', type: 'Text Block', threshold: 84, tag: 'INV_DIA_CHI', status: 'Low Confidence' }
];

const MasterLabels = {
    labels: JSON.parse(localStorage.getItem('orc_master_labels')) || defaultLabels,

    saveToStorage() {
        localStorage.setItem('orc_master_labels', JSON.stringify(this.labels));
    },

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
                this.saveToStorage();
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
   AI ENGINE — LM Studio Integration (Qwen 3.5 9B)
   ═══════════════════════════════════════════════════════════════ */
const AIEngine = {
    endpoint: 'http://localhost:1235',  // CORS proxy → LM Studio at 1234
    model: 'qwen/qwen3.5-9b',
    isAvailable: false,
    _checking: false,

    // ── Low-level API call ───────────────────────────────────────
    // Qwen 3.5 uses reasoning_content for chain-of-thought, so we need
    // to allocate enough tokens AND extract from the right field.
    async _callLLM(messages, { temperature = 0.7, max_tokens = 8192, timeout = 120000, disableThinking = false } = {}) {
        // Prepend /no_think to first user message if thinking should be disabled
        // This is a Qwen 3.5 feature to skip extended reasoning for simple tasks
        const processedMessages = disableThinking
            ? messages.map((m, i) => {
                if (i === messages.length - 1 && m.role === 'user') {
                    return { ...m, content: '/no_think\n' + m.content };
                }
                return m;
            })
            : messages;

        const body = {
            model: this.model,
            messages: processedMessages,
            temperature,
            max_tokens,
            stream: false
        };

        // For Qwen 3.5: limit reasoning budget to leave room for actual content
        // budget_tokens tells the model how many tokens to use for thinking
        if (!disableThinking) {
            body.budget_tokens = Math.min(1024, Math.floor(max_tokens * 0.3));
        }

        const res = await fetch(`${this.endpoint}/v1/chat/completions`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
            signal: AbortSignal.timeout(timeout)
        });
        if (!res.ok) throw new Error(`LLM API returned ${res.status}`);
        const data = await res.json();

        const message = data.choices?.[0]?.message;
        let content = message?.content?.trim() || '';

        // Qwen 3.5: if content is empty but reasoning_content has data,
        // the model put everything in reasoning. Try to extract useful content from it.
        if (!content && message?.reasoning_content) {
            console.warn('[AIEngine] Content empty, extracting from reasoning_content');
            const rc = message.reasoning_content;
            // Look for JSON in reasoning content
            const jsonMatch = rc.match(/```json\s*([\s\S]*?)```/) || rc.match(/(\{[\s\S]*\})/);
            if (jsonMatch) {
                content = jsonMatch[1] || jsonMatch[0];
            } else {
                // Last resort: use the last line(s) of reasoning as content
                const lines = rc.split('\n').filter(l => l.trim());
                content = lines[lines.length - 1] || '';
            }
        }

        // Strip <think>...</think> blocks (Qwen reasoning markup)
        content = content.replace(/<think>[\s\S]*?<\/think>/gi, '').trim();
        return content;
    },

    // ── Extract JSON from LLM response (robust) ─────────────────
    _extractJSON(text) {
        if (!text) return null;
        // Strip markdown fences
        text = text.replace(/```json\s*/gi, '').replace(/```\s*/gi, '');
        // Try full parse first
        try { return JSON.parse(text); } catch (_) {}
        // Find first { ... } or [ ... ]
        const objMatch = text.match(/\{[\s\S]*\}/);
        if (objMatch) { try { return JSON.parse(objMatch[0]); } catch (_) {} }
        const arrMatch = text.match(/\[[\s\S]*\]/);
        if (arrMatch) { try { return JSON.parse(arrMatch[0]); } catch (_) {} }
        return null;
    },

    // ── Check if LM Studio is running ────────────────────────────
    async checkConnection() {
        const dot = document.getElementById('gen-ai-status-dot');
        const text = document.getElementById('gen-ai-status-text');
        if (dot) { dot.className = 'gen-ai-status-dot checking'; }
        if (text) { text.textContent = 'Checking AI Engine...'; }

        this._checking = true;
        try {
            const res = await fetch(`${this.endpoint}/v1/models`, {
                method: 'GET',
                signal: AbortSignal.timeout(5000)
            });
            if (res.ok) {
                const data = await res.json();
                this.isAvailable = true;
                if (dot) { dot.className = 'gen-ai-status-dot online'; }
                if (text) { text.textContent = `Online — ${data.data?.length || 1} model(s) loaded`; }
                // Auto-select qwen model if available, else use first model
                if (data.data && data.data.length > 0) {
                    const qwen = data.data.find(m => m.id.includes('qwen'));
                    this.model = qwen ? qwen.id : data.data[0].id;
                    const modelEl = document.getElementById('gen-ai-model');
                    if (modelEl) modelEl.textContent = this.model;
                }
                this._checking = false;
                return true;
            }
        } catch (e) {
            // Connection failed
        }
        this.isAvailable = false;
        if (dot) { dot.className = 'gen-ai-status-dot offline'; }
        if (text) { text.textContent = 'Offline — Start LM Studio to enable AI'; }
        this._checking = false;
        return false;
    },

    // ═══════════════════════════════════════════════════════════════
    //  STEP 1 AI: Analyze PDF template text to extract field list
    //  Returns: { documentType, fields: [ { code, label, sampleValue, page, x, y, width, height } ] }
    // ═══════════════════════════════════════════════════════════════
    async analyzeTemplateWithAI(templateStructure) {
        if (!this.isAvailable || !templateStructure) return null;

        // Build a condensed text representation for AI analysis
        let textForAI = '';
        for (const page of templateStructure.pages) {
            textForAI += `\n=== PAGE ${page.pageNumber} (${Math.round(page.width)}x${Math.round(page.height)}) ===\n`;
            for (const item of page.items) {
                textForAI += `[x:${Math.round(item.x)},y:${Math.round(item.y)},w:${Math.round(item.width || 50)}] "${item.text}"\n`;
            }
        }

        // Trim aggressively to reduce reasoning overhead on local models
        if (textForAI.length > 3000) {
            textForAI = textForAI.substring(0, 3000) + '\n... (truncated)';
        }

        const systemPrompt = `You are a document analysis expert. Given extracted text with coordinates from a PDF business document, identify:
1. The document type (Invoice, Sales Contract, Bill of Lading, Packing List, Certificate of Origin, Credit Note, Debit Note, etc.)
2. All fillable/variable fields - these are the label:value pairs where values change between documents.

For each field, provide:
- code: a snake_case identifier (Vietnamese style: TEN_NGUOI_BAN, SO_HOA_DON, NGAY, etc.)
- label: human-readable label
- sampleValue: the current value in the template
- page: page number (1-based)
- x, y: approximate coordinates of the VALUE (not the label)
- width, height: approximate size of the value area

IMPORTANT: Focus on identifying VARIABLE fields (names, dates, amounts, addresses, numbers) NOT static headers or titles.
Return ONLY valid JSON, no explanation.`;

        const userPrompt = `Analyze this PDF document and extract all variable fields:

${textForAI}

Return JSON in this exact format:
{
  "documentType": "Invoice",
  "fields": [
    {"code": "SO_HOA_DON", "label": "Invoice Number", "sampleValue": "INV-001", "page": 1, "x": 200, "y": 50, "width": 150, "height": 14},
    ...
  ]
}`;

        try {
            console.log('[AIEngine] Analyzing template with AI (this may take 2-5 minutes)...');
            const content = await this._callLLM([
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt }
            ], { temperature: 0.2, max_tokens: 16384, timeout: 600000, disableThinking: true });

            console.log('[AIEngine] AI response received, length:', content?.length || 0, 'preview:', content?.substring(0, 100));

            const result = this._extractJSON(content);
            if (result && result.fields && Array.isArray(result.fields) && result.fields.length > 0) {
                console.log(`[AIEngine] AI extracted ${result.fields.length} fields, type: ${result.documentType}`);
                return result;
            }
            console.warn('[AIEngine] AI returned invalid structure. Raw:', content?.substring(0, 500));
        } catch (e) {
            console.warn('[AIEngine] Template analysis failed:', e.message);
        }
        return null;
    },

    // ═══════════════════════════════════════════════════════════════
    //  STEP 2 AI: Generate a complete document worth of field values
    //  Uses template context + optional Excel samples as reference
    // ═══════════════════════════════════════════════════════════════
    async generateDocumentContent(fields, docContext, sampleValues, docIndex) {
        if (!this.isAvailable) return null;

        const fieldDesc = fields.map(f => {
            let desc = `- ${f.code} (${f.label})`;
            if (f.sampleValue) desc += ` [example: "${f.sampleValue}"]`;
            return desc;
        }).join('\n');

        const sampleSection = sampleValues && Object.keys(sampleValues).length > 0
            ? `\nReference data from Excel (create SIMILAR but DIFFERENT values):\n${Object.entries(sampleValues).map(([k,v]) => `  ${k}: "${v}"`).join('\n')}`
            : '';

        const systemPrompt = `You are a professional business document generator. Generate realistic data for Vietnamese/international trade documents.

STRICT RULES:
1. Return ONLY a valid JSON object, no explanations, no markdown
2. Keys must match the field codes exactly
3. Generate varied, realistic data — NOT copied from samples
4. Company names: use real-sounding Vietnamese/Chinese/international names
5. Dates: use YYYY-MM-DD format, dates should be recent (2025-2026)
6. Amounts: use numbers with 2 decimal places (e.g., "125,000.00")
7. Document numbers: use format like INV-XX123456, SC-YY789012
8. Addresses: use realistic Vietnamese/Chinese business addresses
9. This is document #${docIndex + 1}, so make data unique from other documents`;

        const userPrompt = `Generate realistic values for ALL fields below for a "${docContext}" document:

${fieldDesc}
${sampleSection}

Return JSON: {"field_code_1": "value1", "field_code_2": "value2", ...}`;

        try {
            const content = await this._callLLM([
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt }
            ], { temperature: 0.85, max_tokens: 4096, timeout: 120000, disableThinking: true });

            const result = this._extractJSON(content);
            if (result && typeof result === 'object' && !Array.isArray(result)) {
                return result;
            }
            console.warn('[AIEngine] Batch parse failed, content:', content?.substring(0, 200));
        } catch (e) {
            console.warn('[AIEngine] Batch generation failed:', e.message);
        }
        return null;
    },

    // ═══════════════════════════════════════════════════════════════
    //  Generate content for a single field (fallback)
    // ═══════════════════════════════════════════════════════════════
    async generateFieldContent(fieldCode, fieldLabel, docContext, existingValue) {
        if (!this.isAvailable) return existingValue || DocumentStore.generateRandomValue(fieldCode);

        const prompt = `Generate a single realistic value for a "${fieldLabel}" field (code: ${fieldCode}) in a ${docContext} document.
${existingValue ? `Reference: "${existingValue}" — generate something similar but different.` : ''}
Return ONLY the value, nothing else. No quotes, no explanation.`;

        try {
            const content = await this._callLLM([
                { role: 'user', content: prompt }
            ], { temperature: 0.8, max_tokens: 1024, timeout: 30000, disableThinking: true });

            if (content && content.length > 0 && content.length < 300) {
                // Remove surrounding quotes if present
                return content.replace(/^["']|["']$/g, '');
            }
        } catch (e) {
            console.warn(`[AIEngine] Field generation failed for ${fieldCode}:`, e.message);
        }
        return existingValue || DocumentStore.generateRandomValue(fieldCode);
    },

    // ═══════════════════════════════════════════════════════════════
    //  Enhance existing Excel data values with AI
    // ═══════════════════════════════════════════════════════════════
    async enhanceExcelContent(fieldCode, fieldLabel, originalValue, docContext) {
        if (!this.isAvailable || !originalValue) return originalValue;

        const prompt = `You are a document content improvement assistant.
Field: ${fieldLabel} (${fieldCode}) in a ${docContext} document.
Current value: "${originalValue}"

If the value looks correct and complete, return it unchanged.
If it can be improved (better formatting, more realistic, more complete), provide an improved version.
Return ONLY the value, no explanation.`;

        try {
            const content = await this._callLLM([
                { role: 'user', content: prompt }
            ], { temperature: 0.3, max_tokens: 1024, timeout: 30000, disableThinking: true });

            if (content && content.length > 0 && content.length < 300) {
                return content.replace(/^["']|["']$/g, '');
            }
        } catch (e) {
            console.warn(`[AIEngine] Enhancement failed for ${fieldCode}`);
        }
        return originalValue;
    }
};


/* ═══════════════════════════════════════════════════════════════
   GENERATE PAGE — 3-Step Document Generation Pipeline
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

        // Show/hide CSV upload area and AI panel based on data source selection
        const dataSourceSelect = document.getElementById('gen-data-source');
        dataSourceSelect.addEventListener('change', () => this.onDataSourceChange());

        // CSV file input handler
        const csvInput = document.getElementById('gen-csv-input');
        if (csvInput) {
            csvInput.addEventListener('change', (e) => this.onCsvSelected(e));
        }

        // Direct template upload on Generate page
        const templateInput = document.getElementById('gen-template-file-input');
        if (templateInput) {
            templateInput.addEventListener('change', (e) => this.onTemplateFileSelected(e));
        }

        // AI Engine controls
        const aiTestBtn = document.getElementById('gen-ai-test-btn');
        if (aiTestBtn) {
            aiTestBtn.addEventListener('click', () => AIEngine.checkConnection());
        }

        // AI Enhancement toggle → show/hide AI panel
        const aiEnhanceToggle = document.getElementById('gen-ai-enhance');
        if (aiEnhanceToggle) {
            aiEnhanceToggle.addEventListener('change', () => this._updateAIPanel());
        }

        // Template select change → extract template
        const sourceSelect = document.getElementById('gen-source-select');
        if (sourceSelect) {
            sourceSelect.addEventListener('change', () => {
                this.updateTemplateInfo();
                this._extractTemplateStructure();
            });
        }
    },

    // ─── AI Panel Visibility ────────────────────────────────────
    _updateAIPanel() {
        const aiPanel = document.getElementById('gen-ai-panel');
        const aiToggle = document.getElementById('gen-ai-enhance');
        const dataSource = document.getElementById('gen-data-source').value;

        const showAI = aiToggle?.checked || dataSource === 'AI Engine';

        if (aiPanel) {
            aiPanel.style.display = showAI ? 'block' : 'none';
            if (showAI && !AIEngine.isAvailable && !AIEngine._checking) {
                AIEngine.checkConnection();
            }
        }
    },

    onTemplateFileSelected(e) {
        const file = e.target.files[0];
        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.pdf')) {
            alert('Please upload a PDF file as template.');
            return;
        }

        // Register in DocumentStore
        const template = DocumentStore.addTemplate(file);

        // Auto-select the newly uploaded template
        const select = document.getElementById('gen-source-select');
        setTimeout(async () => {
            select.value = template.id;
            // Extract structure FIRST, then update info to show field count
            await this._extractTemplateStructure();
            this.updateTemplateInfo();
        }, 100);

        // Update UI label
        const label = document.getElementById('gen-template-upload-label');
        const dropzone = document.querySelector('.gen-template-dropzone');
        if (label) label.textContent = `✓ ${file.name} loaded`;
        if (dropzone) dropzone.classList.add('uploaded');
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
    },

    updateTemplateInfo() {
        const select = document.getElementById('gen-source-select');
        const infoEl = document.getElementById('gen-template-info');
        if (!infoEl) return;

        const tpl = DocumentStore.uploadedTemplates.find(t => t.id === select.value);
        if (tpl) {
            const masterSchema = this._masterFieldSchema;
            const allDocInstances = this._allDocInstances;
            const maxTemplatePage = this._maxTemplatePage || 1;
            const templateStruct = this._templateStructure;
            const templateFieldCount = templateStruct ? Object.keys(templateStruct.extractedFields).length : 0;
            const fieldCount = masterSchema ? Object.keys(masterSchema).length : (templateFieldCount || DocumentStore.ocrData.length);
            const docInstCount = allDocInstances ? Object.keys(allDocInstances).length : 0;

            // Count fields per page for multi-page display
            let pageBreakdown = '';
            if (masterSchema && maxTemplatePage > 1) {
                const pageFieldCounts = {};
                Object.values(masterSchema).forEach(f => {
                    const p = f.page || 1;
                    pageFieldCounts[p] = (pageFieldCounts[p] || 0) + 1;
                });
                pageBreakdown = Object.entries(pageFieldCounts)
                    .map(([p, c]) => `P${p}: ${c}`)
                    .join(' | ');
            }

            infoEl.innerHTML = `
                <div class="template-info-row"><span>File:</span><strong>${tpl.name}</strong></div>
                <div class="template-info-row"><span>Size:</span><span>${tpl.size}</span></div>
                <div class="template-info-row"><span>Template Pages:</span><span>${maxTemplatePage > 1 ? maxTemplatePage + ' pages' : (tpl.pages || 1) + ' page(s)'}</span></div>
                <div class="template-info-row"><span>Uploaded:</span><span>${tpl.uploadedAt ? tpl.uploadedAt.toLocaleDateString() : 'N/A'}</span></div>
                <div class="template-info-row"><span>Schema Fields:</span><span style="color:var(--color-accent);font-weight:600">${fieldCount} field(s)</span></div>
                ${templateStruct ? `<div class="template-info-row"><span>Template Structure:</span><span style="color:var(--color-success);font-weight:600">✓ Extracted</span></div>` : ''}
                ${pageBreakdown ? `<div class="template-info-row"><span>Page Layout:</span><span style="font-family:var(--font-mono);font-size:10px">${pageBreakdown}</span></div>` : ''}
                ${docInstCount > 0 ? `<div class="template-info-row"><span>Excel Instances:</span><span style="color:var(--color-success);font-weight:600">${docInstCount} document(s)</span></div>` : ''}
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
        this._updateAIPanel();
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
             // Build COMPLETE master schema by scanning ALL rows across ALL document instances
             const masterSchema = {};
             const allDocInstances = {};
             let maxPage = 1;

             for (const r of this.csvData.rows) {
                 const fieldCode = (r.field_code || '').trim();
                 if (!fieldCode) continue;

                 const caseId = (r.case_id || 'case_unknown').trim();
                 const docInst = (r.document_instance || r.file_name || 'doc_unknown').trim();
                 const page = parseInt(r.page) || 1;
                 const x = parseFloat(r.x) || 0;
                 const y = parseFloat(r.y) || 0;
                 const width = parseFloat(r.width) || 80;
                 const height = parseFloat(r.height) || 16;
                 const value = (r.value !== undefined && r.value !== null) ? String(r.value) : '';
                 const label = (r.label || r.field_label || fieldCode).trim();

                 if (page > maxPage) maxPage = page;

                 if (!masterSchema[fieldCode]) {
                     masterSchema[fieldCode] = {
                         x, y, width, height, page, label
                     };
                 }

                 const instKey = `${caseId}___${docInst}`;
                 if (!allDocInstances[instKey]) {
                     allDocInstances[instKey] = { caseId, docInstance: docInst, fields: {} };
                 }
                 allDocInstances[instKey].fields[fieldCode] = {
                     value, x, y, width, height, page, label
                 };
             }

             this._masterFieldSchema = masterSchema;
             this._allDocInstances = allDocInstances;
             this._maxTemplatePage = maxPage;

             newFields = Object.keys(masterSchema).map(code => ({
                 field: code, label: masterSchema[code].label || code, type: 'Alphanumeric'
             }));

             console.log(`[SyncCSV] Master schema: ${Object.keys(masterSchema).length} fields across ${maxPage} pages.`);
             console.log(`[SyncCSV] Document instances found: ${Object.keys(allDocInstances).length}`);
        } else {
             newFields = this.csvData.headers.map(h => ({ field: h, label: h, type: 'Alphanumeric' }));
             this._masterFieldSchema = null;
             this._allDocInstances = null;
             this._maxTemplatePage = 1;
        }
        DocumentStore.ocrData = newFields;
        this.updateTemplateInfo();

        // Auto-expand Master Labels
        if (window.MasterLabels && newFields.length > 0) {
            let changed = false;
            newFields.forEach(nf => {
                if (!MasterLabels.labels.find(l => l.tag === nf.field)) {
                    MasterLabels.labels.push({
                         name: nf.label || nf.field,
                         type: nf.type || 'Alphanumeric',
                         threshold: 95,
                         tag: nf.field,
                         status: 'Active'
                    });
                    changed = true;
                }
            });
            if (changed) {
                MasterLabels.saveToStorage();
                MasterLabels.renderLabels();
            }
        }
    },

    _masterFieldSchema: null,
    _allDocInstances: null,
    _maxTemplatePage: 1,
    _templateStructure: null, // Extracted template structure from Step 1
    _aiExtractedFields: null, // AI-analyzed fields from template

    // ═══════════════════════════════════════════════════════════════
    //  STEP 1: Extract Template Structure (PDF → Word-like Schema)
    //  Uses pdf.js for text extraction, then AI for intelligent field detection
    // ═══════════════════════════════════════════════════════════════
    async _extractTemplateStructure() {
        const select = document.getElementById('gen-source-select');
        const tpl = DocumentStore.uploadedTemplates.find(t => t.id === select.value);
        if (!tpl || !tpl._file) {
            this._templateStructure = null;
            this._aiExtractedFields = null;
            return;
        }

        try {
            const arrayBuffer = await tpl._file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            const totalPages = pdf.numPages;
            tpl.pages = totalPages;

            const structure = {
                pages: [],
                totalPages,
                documentType: 'Unknown',
                fullText: '',
                extractedFields: {}
            };

            for (let i = 1; i <= totalPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const viewport = page.getViewport({ scale: 1 });

                // Extract text items with their positions
                const pageItems = textContent.items.map(item => ({
                    text: item.str,
                    x: item.transform[4],
                    y: viewport.height - item.transform[5], // Convert to top-down Y
                    width: item.width,
                    height: item.height || 12,
                    fontName: item.fontName || 'default',
                    fontSize: Math.abs(item.transform[0]) || 12
                })).filter(item => item.text.trim().length > 0);

                const pageText = pageItems.map(item => item.text).join(' ');

                structure.pages.push({
                    pageNumber: i,
                    width: viewport.width,
                    height: viewport.height,
                    items: pageItems,
                    text: pageText
                });

                structure.fullText += (i > 1 ? '\n\n--- PAGE ' + i + ' ---\n\n' : '') + pageText;
            }

            // Detect document type from full text
            structure.documentType = this._detectDocTypeFromText(structure.fullText);

            // Auto-detect field locations from template text patterns (regex fallback)
            this._autoDetectFieldsFromTemplate(structure);

            // Try AI-powered field detection (much more comprehensive)
            if (AIEngine.isAvailable) {
                console.log('[Step1] Attempting AI-powered field extraction...');
                const aiResult = await AIEngine.analyzeTemplateWithAI(structure);
                if (aiResult && aiResult.fields && aiResult.fields.length > 0) {
                    this._aiExtractedFields = aiResult.fields;
                    // Override document type from AI if detected
                    if (aiResult.documentType) {
                        structure.documentType = aiResult.documentType;
                    }
                    // Merge AI fields into extractedFields (AI takes priority)
                    for (const f of aiResult.fields) {
                        structure.extractedFields[f.code] = {
                            label: f.label,
                            sampleValue: f.sampleValue || '',
                            x: f.x || 0,
                            y: f.y || 0,
                            width: f.width || 150,
                            height: f.height || 14,
                            page: f.page || 1,
                            fontName: 'default',
                            fontSize: 11
                        };
                    }
                    console.log(`[Step1] AI extracted ${aiResult.fields.length} fields (overrides regex: ${Object.keys(structure.extractedFields).length} total)`);
                } else {
                    this._aiExtractedFields = null;
                    console.log('[Step1] AI field extraction returned no results, using regex fallback');
                }
            } else {
                this._aiExtractedFields = null;
            }

            // ── FALLBACK: If still 0 fields, generate DEFAULT fields by doc type ──
            if (Object.keys(structure.extractedFields).length === 0) {
                console.log(`[Step1] No fields detected by regex or AI — generating defaults for: ${structure.documentType}`);
                const defaultFields = this._getDefaultFieldsForDocType(structure.documentType, structure.pages[0]);
                Object.assign(structure.extractedFields, defaultFields);
                console.log(`[Step1:Default] Added ${Object.keys(defaultFields).length} default fields`);
            }

            this._templateStructure = structure;
            this._maxTemplatePage = totalPages;

            console.log(`[Step1] Template extracted: ${totalPages} pages, ${Object.keys(structure.extractedFields).length} fields, type: ${structure.documentType}`);
            this.updateTemplateInfo();
        } catch (err) {
            console.error('[Step1] Template extraction error:', err);
            this._templateStructure = null;
            this._aiExtractedFields = null;
        }
    },

    _detectDocTypeFromText(text) {
        const upper = text.toUpperCase();
        if (upper.includes('SALES CONTRACT') || upper.includes('SALE CONTRACT')) return 'Sales Contract';
        if (upper.includes('PURCHASE ORDER') || upper.includes('P.O.')) return 'Purchase Order';
        if (upper.includes('COMMERCIAL INVOICE')) return 'Commercial Invoice';
        if (upper.includes('PROFORMA INVOICE')) return 'Proforma Invoice';
        if (upper.includes('INVOICE')) return 'Invoice';
        if (upper.includes('CREDIT NOTE')) return 'Credit Note';
        if (upper.includes('DEBIT NOTE')) return 'Debit Note';
        if (upper.includes('PACKING LIST')) return 'Packing List';
        if (upper.includes('BILL OF LADING') || upper.includes('B/L')) return 'Bill of Lading';
        if (upper.includes('CERTIFICATE OF ORIGIN')) return 'Certificate of Origin';
        if (upper.includes('RECEIPT')) return 'Receipt';
        if (upper.includes('CONTRACT') || upper.includes('AGREEMENT')) return 'Contract';
        return 'Document';
    },

    // Generate default field positions based on document type when detection fails
    _getDefaultFieldsForDocType(docType, firstPage) {
        const pw = firstPage?.width || 595;
        const ph = firstPage?.height || 842;
        const items = firstPage?.items || [];

        // Try to anchor fields to actual text positions in the PDF
        const findItemY = (keywords) => {
            for (const item of items) {
                const t = item.text.toLowerCase();
                for (const kw of keywords) {
                    if (t.includes(kw.toLowerCase())) {
                        return { x: item.x + (item.width || 50) + 5, y: item.y, w: Math.max(pw * 0.4, 200) };
                    }
                }
            }
            return null;
        };

        const makeField = (label, sampleValue, x, y, width) => ({
            label, sampleValue,
            x: x || pw * 0.45, y: y || 50, width: width || 200,
            height: 14, page: 1, fontName: 'default', fontSize: 10
        });

        const fields = {};

        // --- Common fields for ALL trade documents ---
        const docNoPos = findItemY(['NO', 'No.', 'Number']);
        fields['SO_CHUNG_TU'] = makeField('Document Number', 'SC-MY250001',
            docNoPos?.x || pw * 0.6, docNoPos?.y || 55, 180);

        const datePos = findItemY(['DATE', 'Date', 'Dated']);
        fields['NGAY'] = makeField('Date', '22/09/2025',
            datePos?.x || pw * 0.6, datePos?.y || 75, 120);

        const buyerPos = findItemY(['BUYER', 'Buyer']);
        fields['TEN_NGUOI_MUA'] = makeField('Buyer', 'THIEN LOC VIET NAM IMPORT EXPORT CO.,LTD',
            buyerPos?.x || 80, buyerPos?.y || 100, pw * 0.7);

        const sellerPos = findItemY(['SELLER', 'Seller']);
        fields['TEN_NGUOI_BAN'] = makeField('Seller', 'NINGBO MING YUE GLOBAL TRADING CO., LTD',
            sellerPos?.x || 80, sellerPos?.y || 140, pw * 0.7);

        // --- Type-specific fields ---
        if (['Sales Contract', 'Contract', 'Purchase Order'].includes(docType)) {
            const goodsPos = findItemY(['Goods', 'Description', 'Commodity']);
            fields['MO_TA_HANG'] = makeField('Goods', 'Shoe manufacturing machines, Model: IMSKR9506-L2',
                goodsPos?.x || 100, goodsPos?.y || 260, pw * 0.5);
            fields['SO_LUONG'] = makeField('Quantity', '2',
                (goodsPos?.x || pw * 0.6) + 200, goodsPos?.y || 260, 60);
            fields['DON_GIA'] = makeField('Unit Price', '75,000',
                (goodsPos?.x || pw * 0.7) + 260, goodsPos?.y || 260, 80);

            const totalPos = findItemY(['TOTAL', 'Total', 'Amount']);
            fields['TONG_TIEN'] = makeField('Total Amount', '150,000',
                totalPos?.x || pw * 0.7, totalPos?.y || 300, 100);

            const paymentPos = findItemY(['PAYMENT', 'Payment']);
            fields['PHUONG_THUC_THANH_TOAN'] = makeField('Payment', 'T/T IN ADVANCE',
                paymentPos?.x || 200, paymentPos?.y || 420, pw * 0.4);

            const shipPos = findItemY(['SHIPMENT', 'Shipment']);
            fields['THOI_HAN_GIAO'] = makeField('Shipment', 'Not later than December 2025',
                shipPos?.x || 200, shipPos?.y || 450, pw * 0.4);

            const portLoadPos = findItemY(['Loading port', 'LOADING']);
            fields['CANG_XUAT'] = makeField('Loading Port', 'NINGBO, CHINA',
                portLoadPos?.x || 250, portLoadPos?.y || 480, 180);

            const portDisPos = findItemY(['Discharge', 'DESTINATION']);
            fields['CANG_NHAP'] = makeField('Discharge Port', 'HO CHI MINH, VIETNAM',
                portDisPos?.x || 250, portDisPos?.y || 500, 180);
        }

        if (['Invoice', 'Commercial Invoice', 'Proforma Invoice'].includes(docType)) {
            fields['MO_TA_HANG'] = makeField('Goods', 'Electronic components', 100, 250, pw * 0.5);
            fields['SO_LUONG'] = makeField('Quantity', '500', pw * 0.6, 250, 60);
            fields['DON_GIA'] = makeField('Unit Price', '25.00', pw * 0.7, 250, 80);
            fields['TONG_TIEN'] = makeField('Total', '12,500.00', pw * 0.7, 290, 100);
        }

        if (docType === 'Bill of Lading') {
            fields['TAU'] = makeField('Vessel', 'MV OCEAN STAR', 200, 150, 200);
            fields['CANG_XUAT'] = makeField('Loading Port', 'NINGBO, CHINA', 200, 200, 200);
            fields['CANG_NHAP'] = makeField('Discharge Port', 'CAT LAI, HCM', 200, 230, 200);
            fields['SO_CONTAINER'] = makeField('Container No.', 'MSCU1234567', 200, 260, 150);
        }

        return fields;
    },

    // Auto-detect label:value pairs from template text positions (smart fallback)
    // This is the critical fallback when AI is offline
    _autoDetectFieldsFromTemplate(structure) {
        const docType = (structure.documentType || '').toUpperCase();
        let pfx = 'CRN_';
        if (docType.includes('INVOICE')) pfx = 'INV_';
        else if (docType.includes('CONTRACT') || docType.includes('AGREEMENT') || docType.includes('ORDER')) pfx = 'CON_';
        else if (docType.includes('PACKING') || docType.includes('LIST')) pfx = 'PKL_';
        else if (docType.includes('BILL') || docType.includes('LADING')) pfx = 'OBL_';
        else if (docType.includes('CREDIT') || docType.includes('L/C')) pfx = 'LOC_';

        const labelPatterns = [
            // Standard/Fallback English patterns
            { pattern: /^NO[.:\s#]*$/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : pfx === 'PKL_' ? 'Số chứng từ/ PACKING LIST' : pfx === 'LOC_' ? 'Số LC/ Documentary Credit Number' : 'Số chứng từ' },
            { pattern: /\bNO[:.\s]/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : 'Số chứng từ' },
            { pattern: /\bDATE[:\s]/i, code: pfx === 'CON_' ? `${pfx}NGAY_HOP_DONG` : pfx === 'LOC_' ? `${pfx}NGAY_PHAT_HANH` : `${pfx}NGAY_LAP`, label: pfx === 'LOC_' ? 'Ngày phát hành/ Date of Issue' : 'Ngày chứng từ/ Date' },
            { pattern: /\bDATED[:\s]/i, code: pfx === 'CON_' ? `${pfx}NGAY_HOP_DONG` : pfx === 'LOC_' ? `${pfx}NGAY_PHAT_HANH` : `${pfx}NGAY_LAP`, label: pfx === 'LOC_' ? 'Ngày phát hành/ Date of Issue' : 'Ngày chứng từ/ Date' },
            { pattern: /\bBUYER[:\s]/i, code: `${pfx}B_NGUOI_MUA_TEN`, label: pfx === 'INV_' ? 'Tên người mua/ For account & risk of messrs' : 'Người mua/ For account & risk of messrs' },
            { pattern: /\bSELLER[:\s]/i, code: `${pfx}B_NGUOI_BAN_TEN`, label: 'Tên người bán' },
            { pattern: /\bCOMMODITY/i, code: `${pfx}C_TEN_HANG`, label: 'Tên hàng hóa' },
            { pattern: /\bGOODS\s*DESCRIPTION/i, code: `${pfx}C_TEN_HANG`, label: 'Tên hàng hóa' },
            { pattern: /\bQUANTITY[:\s]/i, code: `${pfx}C_SO_LUONG`, label: 'Số lượng/ Quantity' },
            { pattern: /\bQTY[:\s]/i, code: `${pfx}C_SO_LUONG`, label: 'Số lượng/ Quantity' },
            { pattern: /\bUNIT\s*PRICE[:\s]/i, code: `${pfx}C_DON_GIA`, label: 'Đơn giá' },
            { pattern: /\bPRICE[:\s]/i, code: `${pfx}C_DON_GIA`, label: 'Đơn giá' },
            { pattern: /\bTOTAL[:\s]/i, code: `${pfx}C_THANH_TIEN`, label: 'Thành tiền/ Amount' },
            { pattern: /\bAMOUNT[:\s]/i, code: `${pfx}C_THANH_TIEN`, label: 'Thành tiền/ Amount' },
            
            // Transport / Notification
            { pattern: /\bPORT\s*OF\s*LOADING/i, code: `${pfx}CANG_XUAT`, label: 'Loading Port' },
            { pattern: /\bPORT\s*OF\s*(DISCHARGE|DESTINATION)/i, code: `${pfx}CANG_NHAP`, label: 'Port of Discharge' },
            { pattern: /\bNOTIFY\s*PARTY/i, code: pfx === 'OBL_' ? `${pfx}B_BEN_DUOC_THONG_BAO_TEN` : `${pfx}B_CONG_TY_NHAP_KHAU`, label: pfx === 'OBL_' ? 'Bên được thông báo/ Notify party' : 'Công ty nhập khẩu/ Notify party' },
            
            // Banking / LC
            { pattern: /\bAPPLICANT[:\s]/i, code: `${pfx}B_NGUOI_YEU_CAU_TEN`, label: 'Người yêu cầu/ Applicant' },
            { pattern: /\bBENEFICIARY[:\s]/i, code: `${pfx}B_NGUOI_THU_HUONG_TEN`, label: 'Người thụ hưởng/ Beneficiary' },
            { pattern: /\bISSUING\s*BANK/i, code: `${pfx}A_NGAN_HANG_PHAT_HANH_TEN`, label: 'Ngân hàng phát hành/ Issuing Bank' },
            { pattern: /\bADVISING\s*BANK/i, code: `${pfx}B_NGAN_HANG_THONG_BAO_TEN`, label: 'Ngân hàng thông báo/ Advising Bank' },
            
            // General Company
            { pattern: /\bADDRESS[:\s]/i, code: `${pfx}DIA_CHI`, label: 'Địa Chỉ' },
            { pattern: /\bTEL[:\s]/i, code: `${pfx}DIEN_THOAI`, label: 'Telephone' },
            { pattern: /\bPHONE[:\s]/i, code: `${pfx}DIEN_THOAI`, label: 'Phone' },
            { pattern: /\bFAX[:\s]/i, code: `${pfx}FAX`, label: 'Fax' },
            { pattern: /\bSWIFT\s*(?:CODE)?[:\s]/i, code: `${pfx}SWIFT_CODE`, label: 'SWIFT Code' },
            { pattern: /\bACCOUNT\s*(?:NUMBER|NO)?[:\s]/i, code: `${pfx}SO_TAI_KHOAN`, label: 'Account Number' },
            { pattern: /\bBANK\s*(?:NAME)?[:\s]/i, code: `${pfx}NGAN_HANG`, label: 'Bank' },

            // Vietnamese patterns
            { pattern: /S[ỐO]\s*(?:HÓA\s*ĐƠN|HOA\s*DON)[:\s]/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : 'Số chứng từ' },
            { pattern: /S[ỐO]\s*(?:CHỨNG\s*TỪ|CHUNG\s*TU)[:\s]/i, code: pfx === 'PKL_' ? `${pfx}SO_CHUNG_TU` : `${pfx}SO_HOA_DON`, label: pfx === 'PKL_' ? 'Số chứng từ/ PACKING LIST' : 'Số Document/ Invoice No' },
            { pattern: /S[ỐO][:.\s]/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : 'Số chứng từ' },
            { pattern: /NGÀY|NGAY[:\s]/i, code: pfx === 'CON_' ? `${pfx}NGAY_HOP_DONG` : `${pfx}NGAY_LAP`, label: 'Ngày chứng từ/ Date' },
            { pattern: /NG[ƯƯỜ]*I\s*MUA[:\s]/i, code: `${pfx}B_NGUOI_MUA_TEN`, label: pfx === 'INV_' ? 'Tên người mua/ For account & risk of messrs' : 'Người mua/ For account & risk of messrs' },
            { pattern: /NG[ƯƯỜ]*I\s*BÁN|NGUOI\s*BAN[:\s]/i, code: `${pfx}B_NGUOI_BAN_TEN`, label: 'Tên người bán' },
            { pattern: /Đ[ỊI]A\s*CH[ỈI][:\s]/i, code: `${pfx}DIA_CHI`, label: 'Địa Chỉ' },
            { pattern: /S[ỐO]\s*L[ƯƯỢ]*NG[:\s]/i, code: `${pfx}C_SO_LUONG`, label: 'Số lượng/ Quantity' },
            { pattern: /Đ[ƠO]N\s*GIÁ|DON\s*GIA[:\s]/i, code: `${pfx}C_DON_GIA`, label: 'Đơn giá' },
            { pattern: /TỔNG\s*(?:TIỀN|CONG)|TONG\s*(?:TIEN|CONG)[:\s]/i, code: `${pfx}C_THANH_TIEN`, label: 'Thành tiền/ Amount' },
            { pattern: /HÀNG\s*HÓA|HANG\s*HOA|MÔ\s*TẢ|MO\s*TA[:\s]/i, code: `${pfx}C_TEN_HANG`, label: 'Tên hàng hóa' }
        ];

        for (const page of structure.pages) {
            // ----- Step A: Reconstruct text LINES from fragmented PDF items -----
            // PDF text items are often fragmented. Group items on same Y-level into lines.
            const items = page.items.slice().sort((a, b) => {
                // If items are roughly on the same vertical line (within 10-15px), sort them horizontally.
                if (Math.abs(a.y - b.y) <= 12) {
                    return a.x - b.x;
                }
                return a.y - b.y;
            });
            const lines = [];
            let currentLine = null;

            for (const item of items) {
                if (!currentLine || Math.abs(item.y - currentLine.y) > 15) {
                    // New line
                    currentLine = {
                        y: item.y,
                        items: [item],
                        text: item.text,
                        x: item.x,
                        maxX: item.x + (item.width || 0),
                        height: item.height || 12
                    };
                    lines.push(currentLine);
                } else {
                    // Same line — append
                    currentLine.items.push(item);
                    currentLine.text += ' ' + item.text;
                    currentLine.maxX = Math.max(currentLine.maxX, item.x + (item.width || 0));
                    currentLine.height = Math.max(currentLine.height, item.height || 12);
                }
            }

            // ----- Step B: Match LABELS and extract VALUE from the same line -----
            for (const line of lines) {
                const lineText = line.text.trim();
                if (lineText.length < 2) continue;

                for (const lp of labelPatterns) {
                    if (structure.extractedFields[lp.code]) continue; // Already found
                    
                    const match = lineText.match(lp.pattern);
                    if (!match) continue;

                    // Extract the value part AFTER the label
                    const afterLabel = lineText.substring(match.index + match[0].length).trim();
                    // Clean up: remove leading colons, dashes, etc.
                    const cleanValue = afterLabel.replace(/^[:\s\-–—]+/, '').trim();

                    if (cleanValue.length > 0) {
                        // Find the physical position of the value text items
                        // Take items that are to the RIGHT of the label
                        const labelCharCount = match.index + match[0].length;
                        const labelEndX = line.items[0].x + (labelCharCount * 6); // Approx 6px per char
                        
                        // Filter items that start after the label (with a small fuzz factor)
                        const valueItems = line.items.filter(it => it.x >= labelEndX - 15);
                        const valueItem = valueItems.length > 0 ? valueItems[0] : line.items[line.items.length - 1];
                        const lastValueItem = valueItems.length > 0 ? valueItems[valueItems.length - 1] : valueItem;

                        structure.extractedFields[lp.code] = {
                            label: lp.label,
                            sampleValue: cleanValue.substring(0, 200), // Limit value length
                            x: valueItem.x,
                            y: line.y,
                            width: Math.max(lastValueItem.x + (lastValueItem.width || 80) - valueItem.x, Math.min(150, cleanValue.length * 7)),
                            height: line.height || 14,
                            page: page.pageNumber,
                            fontName: valueItem.fontName || 'default',
                            fontSize: valueItem.fontSize || 10
                        };
                        break; // Move to next line — one label per line
                    } else {
                        // Label found but value might be on the NEXT line
                        const lineIdx = lines.indexOf(line);
                        if (lineIdx < lines.length - 1) {
                            const nextLine = lines[lineIdx + 1];
                            const nextText = nextLine.text.trim();
                            // Only take next line if it looks like a value (not another label)
                            const isLabel = labelPatterns.some(p => p.pattern.test(nextText));
                            if (!isLabel && nextText.length > 1) {
                                const nextFirst = nextLine.items[0];
                                const nextLast = nextLine.items[nextLine.items.length - 1];
                                structure.extractedFields[lp.code] = {
                                    label: lp.label,
                                    sampleValue: nextText.substring(0, 200),
                                    x: nextFirst.x,
                                    y: nextLine.y,
                                    width: Math.max(nextLast.x + (nextLast.width || 80) - nextFirst.x, 120),
                                    height: nextLine.height || 14,
                                    page: page.pageNumber,
                                    fontName: nextFirst.fontName || 'default',
                                    fontSize: nextFirst.fontSize || 10
                                };
                            }
                        }
                        break;
                    }
                }
            }

            // ----- Step C: Also detect INLINE patterns like "NO: MY-TL250922" -----
            for (const line of lines) {
                const lineText = line.text.trim();
                // "NO: xxxx" or "NO.: xxxx"
                const inlinePatterns = [
                    { rx: /\bNO[.:]?\s*[:]\s*(.+)/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : pfx === 'PKL_' ? 'Số chứng từ/ PACKING LIST' : pfx === 'LOC_' ? 'Số LC/ Documentary Credit Number' : 'Số chứng từ' },
                    { rx: /\bDATE[.:]?\s*[:]\s*(.+)/i, code: pfx === 'CON_' ? `${pfx}NGAY_HOP_DONG` : pfx === 'LOC_' ? `${pfx}NGAY_PHAT_HANH` : `${pfx}NGAY_LAP`, label: pfx === 'LOC_' ? 'Ngày phát hành/ Date of Issue' : 'Ngày chứng từ/ Date' },
                    { rx: /\bBUYER[.:]?\s*[:]\s*(.+)/i, code: `${pfx}B_NGUOI_MUA_TEN`, label: pfx === 'INV_' ? 'Tên người mua/ For account & risk of messrs' : 'Người mua/ For account & risk of messrs' },
                    { rx: /\bSELLER[.:]?\s*[:]\s*(.+)/i, code: `${pfx}B_NGUOI_BAN_TEN`, label: 'Tên người bán' },
                    { rx: /\bTOTAL[.:]?\s*[:]\s*(.+)/i, code: `${pfx}C_THANH_TIEN`, label: 'Thành tiền/ Amount' },
                    { rx: /S[ỐO][.:]?\s*[:]\s*(.+)/i, code: pfx === 'INV_' ? `${pfx}SO_HOA_DON` : `${pfx}SO_CHUNG_TU`, label: pfx === 'INV_' ? 'Số Document/ Invoice No' : pfx === 'PKL_' ? 'Số chứng từ/ PACKING LIST' : 'Số chứng từ' },
                    { rx: /NGÀY|NGAY[.:]?\s*[:]\s*(.+)/i, code: pfx === 'CON_' ? `${pfx}NGAY_HOP_DONG` : `${pfx}NGAY_LAP`, label: 'Ngày chứng từ/ Date' },
                    { rx: /Đ[ƠO]N\s*GIÁ|DON\s*GIA[.:]?\s*[:]\s*(.+)/i, code: `${pfx}C_DON_GIA`, label: 'Đơn giá' }
                ];
                for (const ip of inlinePatterns) {
                    if (structure.extractedFields[ip.code]) continue;
                    const m = lineText.match(ip.rx);
                    if (m && m[1] && m[1].trim().length > 0) {
                        const val = m[1].trim();
                        const firstItem = line.items[0];
                        const lastItem = line.items[line.items.length - 1];
                        
                        const fullMatchText = m[0];
                        const labelText = fullMatchText.substring(0, fullMatchText.lastIndexOf(val));
                        const approxLabelWidth = labelText.length * 6;
                        
                        structure.extractedFields[ip.code] = {
                            label: ip.label,
                            sampleValue: val,
                            x: firstItem.x + approxLabelWidth,
                            y: line.y,
                            width: Math.max(line.maxX - (firstItem.x + approxLabelWidth), Math.min(150, val.length * 7)),
                            height: line.height || 14,
                            page: page.pageNumber,
                            fontName: firstItem.fontName || 'default',
                            fontSize: firstItem.fontSize || 10
                        };
                    }
                }
            }
        }

        console.log(`[Step1:Regex] Detected ${Object.keys(structure.extractedFields).length} fields from text patterns`);
    },

    // ═══════════════════════════════════════════════════════════════
    //  STEP 2: Populate Content (AI-powered or Excel or Random)
    //  Now uses AIEngine.generateDocumentContent for full-document generation
    // ═══════════════════════════════════════════════════════════════
    async _populateContentForDocument(fields, docContext, useAI, existingValues, log, docIndex) {
        const aiEnhance = document.getElementById('gen-ai-enhance')?.checked || false;
        const populatedFields = {};

        // ── Mode 1: Full AI generation (data source = AI Engine, or needsPopulation + AI) ──
        if (useAI && AIEngine.isAvailable) {
            const fieldList = Object.keys(fields).map(fc => ({
                code: fc,
                label: fields[fc].label || fc,
                sampleValue: fields[fc].sampleValue || ''
            }));

            this._log(log, `🤖 AI generating content for ${fieldList.length} fields (doc #${docIndex + 1})...`, 'info');

            // Try batch AI generation with the improved method
            const batchResult = await AIEngine.generateDocumentContent(fieldList, docContext, existingValues, docIndex);
            if (batchResult) {
                let matched = 0;
                Object.keys(fields).forEach(fc => {
                    if (batchResult[fc] !== undefined && batchResult[fc] !== '') {
                        populatedFields[fc] = String(batchResult[fc]);
                        matched++;
                    } else {
                        populatedFields[fc] = DocumentStore.generateRandomValue(fc);
                    }
                });
                this._log(log, `✓ AI generated ${matched}/${fieldList.length} field values`, 'success');
                return populatedFields;
            }

            // Fallback: individual field generation
            this._log(log, `⚠ Batch AI failed, generating fields individually...`, 'info');
            let generated = 0;
            for (const fc of Object.keys(fields)) {
                const val = await AIEngine.generateFieldContent(
                    fc, fields[fc].label || fc, docContext, existingValues?.[fc]
                );
                populatedFields[fc] = val;
                generated++;
            }
            this._log(log, `✓ Generated ${generated} fields individually`, 'success');
            return populatedFields;
        }

        // ── Mode 2: Enhance existing Excel values with AI ──
        if (aiEnhance && existingValues && Object.keys(existingValues).length > 0 && AIEngine.isAvailable) {
            this._log(log, `🤖 Enhancing Excel data with AI...`, 'info');
            let enhanced = 0;
            for (const fc of Object.keys(fields)) {
                const original = existingValues[fc] || '';
                if (original) {
                    populatedFields[fc] = await AIEngine.enhanceExcelContent(
                        fc, fields[fc].label || fc, original, docContext
                    );
                    if (populatedFields[fc] !== original) enhanced++;
                } else {
                    populatedFields[fc] = DocumentStore.generateRandomValue(fc);
                }
            }
            if (enhanced > 0) {
                this._log(log, `✓ AI enhanced ${enhanced} field values`, 'success');
            }
            return populatedFields;
        }

        // ── Mode 3: Use existing values or random generation ──
        Object.keys(fields).forEach(fc => {
            populatedFields[fc] = existingValues?.[fc] || DocumentStore.generateRandomValue(fc);
        });
        return populatedFields;
    },

    // ═══════════════════════════════════════════════════════════════
    //  FULL PIPELINE: Start Generation (orchestrates Steps 1-3)
    // ═══════════════════════════════════════════════════════════════
    async startGeneration() {
        if (this.isGenerating) return;

        const templateSelect = document.getElementById('gen-source-select');
        const templateId = templateSelect.value;
        const template = DocumentStore.uploadedTemplates.find(t => t.id === templateId);

        if (!template) {
            alert('Please select a source template. Upload documents via Import & Config first.');
            return;
        }

        const requestedCount = Math.min(parseInt(document.getElementById('gen-count-input').value) || 10, 500);
        const format = document.getElementById('gen-format-select').value;
        const dataSource = document.getElementById('gen-data-source').value;
        const preserveFormat = document.getElementById('gen-preserve-format').checked;
        const useAI = dataSource === 'AI Engine' || document.getElementById('gen-ai-enhance')?.checked;

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

        // ════════════════════════════════════════════════
        //  PRE-CHECK: Ensure AI is connected if needed
        // ════════════════════════════════════════════════
        if (useAI) {
            this._log(log, `Checking AI Engine connection...`, 'info');
            const aiOk = await AIEngine.checkConnection();
            if (aiOk) {
                this._log(log, `✓ AI Engine online — model: ${AIEngine.model}`, 'success');
            } else {
                this._log(log, `⚠ AI Engine offline — will use synthetic data as fallback`, 'info');
            }
        }

        // ════════════════════════════════════════════════
        //  STEP 1: Extract Template Structure
        // ════════════════════════════════════════════════
        this._log(log, `\n═══ STEP 1: Template Extraction ═══`, 'info');
        this._log(log, `Loading template: "${template.name}"`, 'msg');

        // Force re-extraction if AI just came online but we haven't used it yet
        const needsAIExtraction = useAI && AIEngine.isAvailable && !this._aiExtractedFields;
        if (!this._templateStructure || needsAIExtraction) {
            this._log(log, needsAIExtraction ? `Re-extracting with AI assistance (this may take 2-5 minutes)...` : `Extracting template structure...`, 'info');
            await this._extractTemplateStructure();
        }

        if (this._templateStructure) {
            const ts = this._templateStructure;
            this._log(log, `✓ Template: ${ts.totalPages} page(s), type: ${ts.documentType}`, 'success');
            this._log(log, `✓ Detected fields: ${Object.keys(ts.extractedFields).length}${this._aiExtractedFields ? ' (AI-powered)' : ' (regex)'}`, 'success');
            if (this._aiExtractedFields) {
                const fieldNames = this._aiExtractedFields.slice(0, 8).map(f => f.label || f.code).join(', ');
                const more = this._aiExtractedFields.length > 8 ? ` +${this._aiExtractedFields.length - 8} more` : '';
                this._log(log, `  Fields: ${fieldNames}${more}`, 'msg');
            }
        } else {
            this._log(log, `⚠ Template structure not extracted — using schema-based generation`, 'info');
        }

        fill.style.width = '10%';

        // ════════════════════════════════════════════════
        //  STEP 2: Build Content (Excel Data / AI / Random)
        // ════════════════════════════════════════════════
        this._log(log, `\n═══ STEP 2: Content Population ═══`, 'info');
        this._log(log, `Data source: ${dataSource} | AI: ${useAI && AIEngine.isAvailable ? '✓ Active' : '✗ Inactive'}`, 'info');

        const fields = DocumentStore.ocrData;
        const generatedDocs = [];

        const masterSchema = this._masterFieldSchema;
        const allDocInstances = this._allDocInstances;
        const maxTemplatePage = this._maxTemplatePage || 1;
        const templateStruct = this._templateStructure;

        // Merge template-extracted fields into master schema if no Excel schema exists
        let effectiveSchema = masterSchema;
        if (!effectiveSchema && templateStruct && Object.keys(templateStruct.extractedFields).length > 0) {
            effectiveSchema = {};
            Object.entries(templateStruct.extractedFields).forEach(([code, f]) => {
                effectiveSchema[code] = {
                    x: f.x,
                    y: f.y,
                    width: f.width,
                    height: f.height,
                    page: f.page,
                    label: f.label,
                    sampleValue: f.sampleValue || ''
                };
            });
            this._log(log, `Using ${Object.keys(effectiveSchema).length} fields from template extraction`, 'info');
        }

        // Determine document context for AI
        const docContext = templateStruct?.documentType || 'Business Document';

        // Build generation queue
        const useExcelData = allDocInstances && Object.keys(allDocInstances).length > 0;
        let generationQueue = [];

        if (useExcelData) {
            const instKeys = Object.keys(allDocInstances);
            instKeys.forEach(key => {
                const inst = allDocInstances[key];
                generationQueue.push({
                    caseId: inst.caseId,
                    docInstance: inst.docInstance,
                    fields: inst.fields
                });
            });
            this._log(log, `Excel data loaded: ${generationQueue.length} document instance(s)`, 'info');
        } else if (effectiveSchema && Object.keys(effectiveSchema).length > 0) {
            this._log(log, `Schema: ${Object.keys(effectiveSchema).length} field(s)`, 'info');
            this._log(log, `Generating ${requestedCount} document(s) with ${useAI && AIEngine.isAvailable ? 'AI' : 'synthetic'} data...`, 'info');

            for (let i = 0; i < requestedCount; i++) {
                const synthFields = {};
                Object.keys(effectiveSchema).forEach(fc => {
                    const coord = effectiveSchema[fc];
                    synthFields[fc] = {
                        value: '', // Will be populated in Step 2
                        x: coord.x,
                        y: coord.y,
                        width: coord.width,
                        height: coord.height,
                        page: coord.page,
                        label: coord.label,
                        sampleValue: coord.sampleValue || ''
                    };
                });
                generationQueue.push({
                    caseId: `CASE_${String(i + 1).padStart(4, '0')}`,
                    docInstance: `DOC_${String(i + 1).padStart(3, '0')}`,
                    fields: synthFields,
                    needsPopulation: true
                });
            }
        } else {
            // No schema at all — use ocrData fields
            for (let i = 0; i < requestedCount; i++) {
                const synthFields = {};
                fields.forEach(f => {
                    synthFields[f.field] = {
                        value: '',
                        x: 50,
                        y: 50 + Object.keys(synthFields).length * 20,
                        width: 200,
                        height: 16,
                        page: 1,
                        label: f.label
                    };
                });
                generationQueue.push({
                    caseId: `CASE_${String(i + 1).padStart(4, '0')}`,
                    docInstance: `DOC_${String(i + 1).padStart(3, '0')}`,
                    fields: synthFields,
                    needsPopulation: true
                });
            }
            this._log(log, `No schema — generating ${requestedCount} documents with ${fields.length} fields`, 'info');
        }

        const actualCount = generationQueue.length;
        countEl.textContent = `0 / ${actualCount}`;

        // ════════════════════════════════════════════════
        //  Populate content for items that need it
        // ════════════════════════════════════════════════
        for (let i = 0; i < generationQueue.length; i++) {
            const item = generationQueue[i];
            
            // Update progress
            const populatePct = 10 + (i / generationQueue.length) * 40;
            fill.style.width = populatePct + '%';
            countEl.textContent = `Populating ${i + 1} / ${actualCount}`;

            if (item.needsPopulation || (useAI && AIEngine.isAvailable)) {
                const existingValues = {};
                Object.keys(item.fields).forEach(fc => {
                    if (item.fields[fc].value) existingValues[fc] = item.fields[fc].value;
                });

                const populated = await this._populateContentForDocument(
                    item.fields, docContext,
                    item.needsPopulation && useAI, // Use full AI generation for items without data
                    existingValues,
                    log,
                    i  // Pass document index for unique generation
                );

                // Update field values
                Object.keys(populated).forEach(fc => {
                    if (item.fields[fc]) {
                        item.fields[fc].value = populated[fc];
                    }
                });
            }

            // Ensure all fields have values (random fallback)
            Object.keys(item.fields).forEach(fc => {
                if (!item.fields[fc].value) {
                    item.fields[fc].value = DocumentStore.generateRandomValue(fc);
                }
            });
        }

        // ════════════════════════════════════════════════
        //  STEP 3: Generate Documents (PDF Output)
        // ════════════════════════════════════════════════
        this._log(log, `\n═══ STEP 3: PDF Generation ═══`, 'info');
        this._log(log, `Total documents to generate: ${actualCount}`, 'info');

        if (effectiveSchema) {
            const fieldList = Object.keys(effectiveSchema).slice(0, 15).join(', ');
            const more = Object.keys(effectiveSchema).length > 15 ? ` ... +${Object.keys(effectiveSchema).length - 15} more` : '';
            this._log(log, `Fields: ${fieldList}${more}`, 'info');
        }

        let current = 0;
        const processNext = () => {
            return new Promise(resolve => {
                const interval = setInterval(() => {
                    if (current >= actualCount) {
                        clearInterval(interval);
                        resolve();
                        return;
                    }

                    const item = generationQueue[current];
                    current++;

                    const pct = (current / actualCount) * 100;
                    fill.style.width = pct + '%';
                    countEl.textContent = `${current} / ${actualCount}`;

                    const docData = {};
                    const docMeta = {};

                    Object.keys(item.fields).forEach(fc => {
                        const fieldInfo = item.fields[fc];
                        docData[fc] = fieldInfo.value;
                        docMeta[fc] = [{
                            value: fieldInfo.value,
                            x: fieldInfo.x,
                            y: fieldInfo.y,
                            width: fieldInfo.width,
                            height: fieldInfo.height,
                            page: fieldInfo.page,
                            label: fieldInfo.label || fc,
                            rotation: 0
                        }];
                    });

                    const docName = `${item.caseId} - ${item.docInstance}`;

                    generatedDocs.push({
                        id: `doc_${Date.now()}_${current}`,
                        templateId: templateId,
                        name: docName,
                        caseId: item.caseId,
                        docInstance: item.docInstance,
                        data: docData,
                        meta: docMeta,
                        format: format,
                        generatedAt: new Date(),
                        preserveFormat: preserveFormat,
                        maxTemplatePage: maxTemplatePage,
                    });

                    this._log(log, `Generated ${docName} (${Object.keys(item.fields).length} fields)`, 'success');
                }, 120);
            });
        };

        await processNext();

        // Store all generated documents
        DocumentStore.addGeneratedDocuments(generatedDocs);

        this._log(log, `\n✓ All ${actualCount} documents generated successfully!`, 'success');
        this._log(log, `→ ${actualCount} documents ready for export. Go to Export page to download.`, 'info');

        // Show completion actions
        this._showCompletionActions(actualCount);

        this.isGenerating = false;
        generateBtn.disabled = false;
        generateBtn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 2v14M2 9h14"/></svg>
            Generate Documents
        `;
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
        const headers = ['case_id', 'document_instance', 'document_name', 'template', 'format', 'generated_at', ...fields];
        const rows = docs.map(doc => {
            const tpl = DocumentStore.uploadedTemplates.find(t => t.id === doc.templateId);
            const tplName = tpl ? tpl.name : 'N/A';
            const base = [
                `"${doc.caseId || ''}"`,
                `"${doc.docInstance || ''}"`,
                `"${doc.name}"`,
                `"${tplName}"`,
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
            documents: docs.map(doc => {
                const tpl = DocumentStore.uploadedTemplates.find(t => t.id === doc.templateId);
                return {
                    name: doc.name,
                    caseId: doc.caseId,
                    docInstance: doc.docInstance,
                    template: tpl ? tpl.name : 'N/A',
                    format: doc.format,
                    generatedAt: doc.generatedAt.toISOString(),
                    preserveFormat: doc.preserveFormat,
                    extractedData: doc.data,
                    fieldMeta: doc.meta,
                };
            }),
        };
        const json = JSON.stringify(exportData, null, 2);
        const filename = 'ocr_structured_data.json';
        this._downloadFile(json, filename, 'application/json');
        return filename;
    },

    _buildSummary() {
        const docs = DocumentStore.generatedDocuments;
        const firstTpl = DocumentStore.uploadedTemplates.find(t => t.id === docs[0]?.templateId);
        let report = '═══════════════════════════════════════════════════════════\n';
        report += '  INSIGHT LEDGER — Document Generation Report\n';
        report += '═══════════════════════════════════════════════════════════\n\n';
        report += `Generated: ${new Date().toLocaleString()}\n`;
        report += `Total Documents: ${docs.length}\n`;
        report += `Template: ${firstTpl?.name || 'N/A'}\n`;
        report += `Format: ${docs[0]?.format || 'N/A'}\n\n`;
        report += '───────────────────────────────────────────────────────────\n\n';
        docs.forEach((doc, i) => {
            report += `Document ${i + 1}: ${doc.name}\n`;
            report += `  Case ID: ${doc.caseId || 'N/A'}\n`;
            report += `  Doc Instance: ${doc.docInstance || 'N/A'}\n`;
            report += `  Generated: ${doc.generatedAt.toLocaleString()}\n`;
            Object.entries(doc.data).forEach(([key, val]) => {
                const meta = doc.meta && doc.meta[key] && doc.meta[key][0];
                const pageInfo = meta ? ` [Page ${meta.page || 1}]` : '';
                report += `  ${key}: ${val}${pageInfo}\n`;
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
            if (!window.JSZip) {
                alert('JSZip is not currently loaded.');
                return null;
            }

            const { PDFDocument, rgb, StandardFonts } = window.PDFLib;
            const zip = new window.JSZip();

            // Fetch custom font buffer once
            let fontBuffer = null;
            try {
                const fontUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf';
                fontBuffer = await fetch(fontUrl).then(res => res.arrayBuffer());
            } catch(e) {
                console.warn('Custom Roboto font fetch failed:', e);
            }

            // Pre-load the template PDF bytes once to avoid reading the file N times
            let cachedTemplateBytes = null;
            const firstDoc = docs[0];
            const templateRef = DocumentStore.uploadedTemplates.find(t => t.id === firstDoc.templateId);
            if (templateRef && templateRef._file) {
                cachedTemplateBytes = await new Promise((resolve, reject) => {
                    const r = new FileReader();
                    r.onload = e => resolve(e.target.result);
                    r.onerror = e => reject(new Error('FileReader Error'));
                    r.readAsArrayBuffer(templateRef._file);
                });
            }

            for (let i = 0; i < docs.length; i++) {
                const doc = docs[i];
                const caseId = doc.caseId || 'Uncategorized';
                const docInstance = doc.docInstance || `doc_${i+1}`;
                
                const outPdf = await PDFDocument.create();
                
                if (window.fontkit) {
                    outPdf.registerFontkit(window.fontkit);
                }
                
                let font;
                if (fontBuffer) {
                    font = await outPdf.embedFont(fontBuffer, { subset: true });
                } else {
                    font = await outPdf.embedFont(StandardFonts?.Helvetica || 'Helvetica');
                }

                let templatePdfDoc;
                if (cachedTemplateBytes) {
                     templatePdfDoc = await PDFDocument.load(cachedTemplateBytes.slice(0));
                } else {
                     templatePdfDoc = await PDFDocument.create();
                     const maxPage = doc.maxTemplatePage || 1;
                     for (let p = 0; p < maxPage; p++) {
                         templatePdfDoc.addPage([595.28, 841.89]); // A4 fallback
                     }
                }
                
                // Copy ALL pages from template
                const templatePageIndices = templatePdfDoc.getPageIndices();
                const copiedPages = await outPdf.copyPages(templatePdfDoc, templatePageIndices);
                const startIdx = outPdf.getPageCount();
                copiedPages.forEach(p => outPdf.addPage(p));
                
                const allPages = outPdf.getPages();
                
                // ── STEP 1: WHITE-BOX over old field values ──
                // For each field's coordinate, draw a white rectangle to cover the original content
                // This ensures old template values are replaced, not overlaid
                Object.keys(doc.data || {}).forEach(field => {
                     const metas = doc.meta && doc.meta[field];
                     if (!metas || !Array.isArray(metas)) return;
                     
                     metas.forEach(meta => {
                         if (typeof meta.x === 'number' && !isNaN(meta.x) && typeof meta.y === 'number' && !isNaN(meta.y)) {
                              const pageOffset = Math.max(0, (parseInt(meta.page) || 1) - 1);
                              const targetIdx = startIdx + pageOffset;
                              
                              if (targetIdx < allPages.length) {
                                   const page = allPages[targetIdx];
                                   const { height: pageH } = page.getSize();
                                   
                                   const w = meta.width || 80;
                                   const h = meta.height || 16;
                                   const pdfY = pageH - meta.y - h;
                                   
                                   // White-box: cover old content completely
                                   page.drawRectangle({
                                       x: meta.x - 1,
                                       y: pdfY - 1,
                                       width: w + 2,
                                       height: h + 2,
                                       color: rgb(1, 1, 1),  // Pure white
                                       borderWidth: 0
                                   });
                              }
                         }
                     });
                });

                // ── STEP 2: Draw new values on top of white-boxed areas ──
                Object.keys(doc.data || {}).forEach(field => {
                     const metas = doc.meta && doc.meta[field];
                     if (!metas || !Array.isArray(metas)) return;
                     
                     metas.forEach(meta => {
                         if (typeof meta.x === 'number' && !isNaN(meta.x) && typeof meta.y === 'number' && !isNaN(meta.y)) {
                              const pageOffset = Math.max(0, (parseInt(meta.page) || 1) - 1);
                              const targetIdx = startIdx + pageOffset;
                              
                              if (targetIdx < allPages.length && font) {
                                   const page = allPages[targetIdx];
                                   const { height: pageH, width: pageW } = page.getSize();
                                   
                                   const w = meta.width || 80;
                                   const h = meta.height || 16;
                                   const pdfY = pageH - meta.y - h;
                                   
                                   // Calculate appropriate font size to fit the bounding box
                                   const textValue = String(meta.value || '');
                                   let fontSize = Math.min(h * 0.7, 11);
                                   if (fontSize < 6) fontSize = 6;
                                   
                                   // Truncate text to fit width
                                   let displayText = textValue;
                                   try {
                                       const textWidth = font.widthOfTextAtSize(displayText, fontSize);
                                       if (textWidth > w - 4) {
                                           // Reduce font size or truncate
                                           const ratio = (w - 4) / textWidth;
                                           if (ratio > 0.6) {
                                               fontSize = Math.max(6, fontSize * ratio);
                                           } else {
                                               const maxChars = Math.floor(displayText.length * ratio);
                                               displayText = displayText.substring(0, Math.max(1, maxChars - 1)) + '…';
                                           }
                                       }
                                   } catch(e) { /* ignore font measurement errors */ }
                                   
                                   // Draw the new value text
                                   page.drawText(displayText, {
                                       x: meta.x + 2,
                                       y: pdfY + (h - fontSize) / 2 + 1,
                                       font: font,
                                       size: fontSize,
                                       color: rgb(0.05, 0.05, 0.15)
                                   });
                              }
                         }
                     });
                });

                // ── STEP 3: Draw OCR LABEL overlays (colored tags above fields) ──
                const labelColors = [
                    [0.22, 0.56, 0.87],  // Blue
                    [0.16, 0.71, 0.37],  // Green
                    [0.90, 0.30, 0.25],  // Red
                    [0.57, 0.35, 0.82],  // Purple
                    [0.95, 0.61, 0.07],  // Orange
                    [0.10, 0.74, 0.61],  // Teal
                    [0.83, 0.18, 0.42],  // Pink
                    [0.40, 0.65, 0.12],  // Lime
                ];
                let colorIdx = 0;
                Object.keys(doc.data || {}).forEach(field => {
                     const metas = doc.meta && doc.meta[field];
                     if (!metas || !Array.isArray(metas)) return;
                     const lc = labelColors[colorIdx % labelColors.length];
                     colorIdx++;
                     metas.forEach(meta => {
                         if (typeof meta.x !== 'number' || typeof meta.y !== 'number') return;
                         const pageOffset = Math.max(0, (parseInt(meta.page) || 1) - 1);
                         const targetIdx = startIdx + pageOffset;
                         if (targetIdx >= allPages.length || !font) return;
                         const page = allPages[targetIdx];
                         const { height: pageH } = page.getSize();
                         const w = meta.width || 80;
                         const h = meta.height || 16;
                         const pdfY = pageH - meta.y - h;
                         // Colored border around field value
                         page.drawRectangle({
                             x: meta.x - 1, y: pdfY - 1,
                             width: w + 2, height: h + 2,
                             borderColor: rgb(lc[0], lc[1], lc[2]),
                             borderWidth: 0.8, opacity: 0.6
                         });
                         // Label tag above field
                         const labelText = (meta.label || field).substring(0, 30);
                         const labelFontSize = 6;
                         let labelW;
                         try { labelW = font.widthOfTextAtSize(labelText, labelFontSize) + 6; } catch(e) { labelW = 40; }
                         labelW = Math.max(labelW, 20);
                         const labelH = 9;
                         // Label background
                         page.drawRectangle({
                             x: meta.x - 1, y: pdfY + h + 1,
                             width: labelW, height: labelH,
                             color: rgb(lc[0], lc[1], lc[2]), borderWidth: 0
                         });
                         // Label text (white)
                         try {
                             page.drawText(labelText, {
                                 x: meta.x + 2, y: pdfY + h + 3,
                                 font, size: labelFontSize, color: rgb(1, 1, 1)
                             });
                         } catch(e) {}
                     });
                });
                
                const pdfBytes = await outPdf.save();
                
                // Add to ZIP under a folder named by case_id
                const safeCaseId = caseId.replace(/[^a-zA-Z0-9.\-_]/g, '_');
                const safeDocInstance = docInstance.replace(/[^a-zA-Z0-9.\-_]/g, '_');
                zip.folder(safeCaseId).file(`${safeDocInstance}.pdf`, pdfBytes);
            }

            // Also generate an Excel summary file with all document data
            if (window.XLSX) {
                const summaryData = [];
                docs.forEach(doc => {
                    Object.keys(doc.data || {}).forEach(field => {
                        const metas = doc.meta && doc.meta[field];
                        const meta = (metas && metas[0]) || {};
                        summaryData.push({
                            case_id: doc.caseId,
                            document_instance: doc.docInstance,
                            field_code: field,
                            label: meta.label || field,
                            value: doc.data[field],
                            page: meta.page || 1,
                            x: meta.x || 0,
                            y: meta.y || 0,
                            width: meta.width || 80,
                            height: meta.height || 16
                        });
                    });
                });
                const ws = XLSX.utils.json_to_sheet(summaryData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Generated_Data');
                const xlsxBytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
                zip.file('generated_data_summary.xlsx', xlsxBytes);
            }

            const zipBlob = await zip.generateAsync({ type: "blob" });
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'business_documents_archive.zip';
            a.click();
            
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            return 'business_documents_archive.zip';
        } catch (e) {
            alert("Error generating PDF Archive: " + e.message + "\n" + (e.stack ? e.stack.substring(0, 300) : ""));
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
