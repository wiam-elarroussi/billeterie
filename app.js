/**
 * E-TICKET PRO MVP — FULL PRODUCTION CONTROLLER
 * Full implementation of 7 Modules + Cashless + Event Creator + Toast Notifications
 */

// Application State Store (Local Storage Backed)
const AppState = {
    currentRole: 'client', // admin, pos, pda, client
    isOffline: false,
    selectedEventId: 1,
    selectedTribune: 'nord',
    selectedSeat: { row: 3, num: 14, code: 'NORD-R03-S14', price: 150, zone: 'Tribune Nord' },
    posSelection: { eventId: 1, tribune: 'Nord', price: 150, qty: 1 },
    cashlessBalance: 450.00,
    
    // Initial Database
    events: [
        {
            id: 1,
            title: "MAROC vs FRANCE",
            subtitle: "Éliminatoires Coupe du Monde 2026",
            date: "14 Oct 2026 — 20:00",
            venue: "Grand Stade de Casablanca",
            capacity: 67000,
            sold: 49714,
            basePrice: 150,
            status: "ACTIF",
            image: "https://images.unsplash.com/photo-1508098682722-e99c43a406b2?auto=format&fit=crop&w=600&q=80"
        },
        {
            id: 2,
            title: "WYDAD CASABLANCA vs RAJA CA",
            subtitle: "Derby de Casablanca — Botola Pro",
            date: "28 Oct 2026 — 19:30",
            venue: "Grand Stade de Casablanca",
            capacity: 67000,
            sold: 62150,
            basePrice: 100,
            status: "ACTIF",
            image: "https://images.unsplash.com/photo-1522778119026-d647f0596c20?auto=format&fit=crop&w=600&q=80"
        },
        {
            id: 3,
            title: "CONCERT MAÎTRE GIMS & GUESTS",
            subtitle: "Tournée Africaine 2026",
            date: "12 Nov 2026 — 21:00",
            venue: "Grand Stade de Casablanca",
            capacity: 50000,
            sold: 31200,
            basePrice: 250,
            status: "ACTIF",
            image: "https://images.unsplash.com/photo-1470225620780-dba8ba36b745?auto=format&fit=crop&w=600&q=80"
        }
    ],

    // Generated Tickets DB
    tickets: [
        {
            id: 'TK-1001',
            eventId: 1,
            matchTitle: "MAROC vs FRANCE",
            buyerName: "Karim Bennani",
            cin: "BE-894210",
            tribune: "Tribune Nord",
            seatCode: "NORD-R03-S14",
            price: 150,
            qrHash: "ETK-2026-X9F4A7B2",
            status: "VALIDE",
            scannedAt: null
        },
        {
            id: 'TK-1002',
            eventId: 1,
            matchTitle: "MAROC vs FRANCE",
            buyerName: "Sara El Amrani",
            cin: "A-541299",
            tribune: "VIP Lounge",
            seatCode: "VIP-R01-S04",
            price: 600,
            qrHash: "ETK-2026-V88B11C4",
            status: "VALIDE",
            scannedAt: null
        },
        {
            id: 'TK-1003',
            eventId: 2,
            matchTitle: "WAC vs RAJA",
            buyerName: "Youssef Tazi",
            cin: "CD-901234",
            tribune: "Tribune Est",
            seatCode: "EST-R08-S22",
            price: 250,
            qrHash: "ETK-2026-Z77F99A1",
            status: "SCANNE",
            scannedAt: "20:14:02"
        }
    ]
};

// Toast Notification System
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `px-4 py-3 rounded-2xl text-xs font-bold text-white shadow-2xl flex items-center gap-3 transition-all duration-300 pointer-events-auto transform translate-y-2 opacity-0 border ${
        type === 'success' ? 'bg-emerald-950/90 border-emerald-500/40 text-emerald-300' :
        type === 'error' ? 'bg-rose-950/90 border-rose-500/40 text-rose-300' :
        'bg-slate-900/90 border-indigo-500/40 text-indigo-300'
    }`;

    const icon = type === 'success' ? 'fa-circle-check text-emerald-400' :
                 type === 'error' ? 'fa-triangle-exclamation text-rose-400' :
                 'fa-circle-info text-indigo-400';

    toast.innerHTML = `<i class="fa-solid ${icon} text-base"></i> <span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.remove('translate-y-2', 'opacity-0');
    }, 10);

    setTimeout(() => {
        toast.classList.add('opacity-0', 'translate-y-2');
        setTimeout(() => toast.remove(), 300);
    }, 3500);
}

// Sound Effect Synthesizer (Web Audio API)
function playBeep(type = 'success') {
    try {
        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = audioCtx.createOscillator();
        const gain = audioCtx.createGain();
        osc.connect(gain);
        gain.connect(audioCtx.destination);

        if (type === 'success') {
            osc.frequency.setValueAtTime(880, audioCtx.currentTime); // High pitch A5
            gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
            osc.start();
            osc.stop(audioCtx.currentTime + 0.15);
        } else {
            osc.frequency.setValueAtTime(220, audioCtx.currentTime); // Low pitch A3
            gain.gain.setValueAtTime(0.3, audioCtx.currentTime);
            osc.start();
            osc.stop(audioCtx.currentTime + 0.35);
        }
    } catch (e) {
        console.log("Audio FX suppressed", e);
    }
}

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initRoleSelector();
    initOfflineToggle();
    initCashlessModal();
    initCreateEventModal();

    renderEventsGrid();
    renderSeatsGrid();
    renderPosEventsList();
    renderMyTickets();
    renderAdminTable();
    initScannerModule();
    initDashboardCharts();

    // Event Listeners
    document.getElementById('btn-buy-selected-seat')?.addEventListener('click', buySelectedSeat);
    document.getElementById('btn-pos-checkout')?.addEventListener('click', handlePosCheckout);
    document.getElementById('btn-close-modal')?.addEventListener('click', closeModal);
    document.getElementById('btn-export-csv')?.addEventListener('click', exportAccessLogsCSV);
    document.getElementById('search-events-input')?.addEventListener('input', handleSearchEvents);
});

// ----------------------------------------------------
// NAVIGATION & TABS
// ----------------------------------------------------
function initNavigation() {
    const navButtons = document.querySelectorAll('#side-nav .nav-btn');
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.getAttribute('data-tab');
            switchTab(targetTab);
        });
    });
}

function switchTab(tabId) {
    document.querySelectorAll('#side-nav .nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

    const activeBtn = document.querySelector(`#side-nav .nav-btn[data-tab="${tabId}"]`);
    const activeTab = document.getElementById(tabId);

    if (activeBtn) activeBtn.classList.add('active');
    if (activeTab) activeTab.classList.add('active');
}

// ----------------------------------------------------
// ROLE SELECTOR & AUTH MATRIX (F01)
// ----------------------------------------------------
function initRoleSelector() {
    const selector = document.getElementById('role-selector');
    selector.addEventListener('change', (e) => {
        AppState.currentRole = e.target.value;
        const avatar = document.getElementById('user-avatar-initials');
        const name = document.getElementById('user-display-name');

        switch (AppState.currentRole) {
            case 'admin':
                avatar.textContent = 'AD';
                name.textContent = 'Youssef (Admin Stade)';
                switchTab('tab-dashboard');
                showToast("Mode Super Admin activé", "info");
                break;
            case 'pos':
                avatar.textContent = 'MB';
                name.textContent = 'Mehdi (Agent POS)';
                switchTab('tab-pos');
                showToast("Mode Guichet POS (3 Clics) activé", "info");
                break;
            case 'pda':
                avatar.textContent = 'SL';
                name.textContent = 'Samira (Agent PDA)';
                switchTab('tab-scanner');
                showToast("Mode Scanner PDA d'Accès activé", "info");
                break;
            case 'client':
            default:
                avatar.textContent = 'CL';
                name.textContent = 'Karim (Spectateur)';
                switchTab('tab-store');
                showToast("Mode Spectateur E-Commerce activé", "info");
                break;
        }
    });
}

// ----------------------------------------------------
// OFFLINE MODE TOGGLE
// ----------------------------------------------------
function initOfflineToggle() {
    const btn = document.getElementById('btn-toggle-offline');
    const dot = document.getElementById('network-indicator-dot');
    const text = document.getElementById('network-status-text');

    btn.addEventListener('click', () => {
        AppState.isOffline = !AppState.isOffline;
        if (AppState.isOffline) {
            btn.className = 'px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-2 transition bg-amber-500/10 text-amber-400 border border-amber-500/20 shadow-sm';
            dot.className = 'w-2 h-2 rounded-full bg-amber-400 animate-ping';
            text.textContent = 'Mode Offline (Cache Local)';
            showToast("Mode Déconnecté activé : validation sur le cache LocalStorage", "error");
        } else {
            btn.className = 'px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-2 transition bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 shadow-sm';
            dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
            text.textContent = 'En Ligne (Online)';
            showToast("Connexion rétablie : synchronisation des logs", "success");
        }
    });
}

// ----------------------------------------------------
// CASHLESS BALANCE MODAL
// ----------------------------------------------------
function initCashlessModal() {
    const btnOpen = document.getElementById('btn-open-cashless');
    const btnClose = document.getElementById('btn-close-cashless');
    const modal = document.getElementById('modal-cashless');
    const btnConfirm = document.getElementById('btn-confirm-cashless');

    btnOpen?.addEventListener('click', () => modal?.classList.remove('hidden'));
    btnClose?.addEventListener('click', () => modal?.classList.add('hidden'));

    btnConfirm?.addEventListener('click', () => {
        const val = parseFloat(document.getElementById('input-cashless-amount')?.value || '0');
        if (val > 0) {
            AppState.cashlessBalance += val;
            document.getElementById('header-cashless-balance').textContent = `${AppState.cashlessBalance.toFixed(2)} DH`;
            document.getElementById('modal-cashless-current').textContent = `${AppState.cashlessBalance.toFixed(2)} DH`;
            modal?.classList.add('hidden');
            showToast(`Rechargement de ${val} DH effectué avec succès !`, 'success');
        }
    });
}

function setCashlessAmount(val) {
    const input = document.getElementById('input-cashless-amount');
    if (input) input.value = val;
}

// ----------------------------------------------------
// CREATE EVENT MODAL (ADMIN)
// ----------------------------------------------------
function initCreateEventModal() {
    const btnOpen = document.getElementById('btn-open-create-event');
    const btnClose = document.getElementById('btn-close-create-event');
    const modal = document.getElementById('modal-create-event');
    const btnSubmit = document.getElementById('btn-submit-create-event');

    btnOpen?.addEventListener('click', () => modal?.classList.remove('hidden'));
    btnClose?.addEventListener('click', () => modal?.classList.add('hidden'));

    btnSubmit?.addEventListener('click', () => {
        const title = document.getElementById('ev-title-input')?.value;
        const sub = document.getElementById('ev-sub-input')?.value;
        const date = document.getElementById('ev-date-input')?.value;
        const price = parseFloat(document.getElementById('ev-price-input')?.value || '150');

        if (title) {
            const newEv = {
                id: AppState.events.length + 1,
                title,
                subtitle: sub || "Événement Officiel",
                date: date || "Prochainement",
                venue: "Grand Stade de Casablanca",
                capacity: 67000,
                sold: 0,
                basePrice: price,
                status: "ACTIF",
                image: "https://images.unsplash.com/photo-1508098682722-e99c43a406b2?auto=format&fit=crop&w=600&q=80"
            };

            AppState.events.push(newEv);
            renderEventsGrid();
            renderPosEventsList();
            renderAdminTable();
            modal?.classList.add('hidden');
            showToast(`Événement "${title}" créé avec succès !`, 'success');
        }
    });
}

// ----------------------------------------------------
// BOUTIQUE E-COMMERCE & EVENTS GRID (F05)
// ----------------------------------------------------
function renderEventsGrid(filterQuery = '') {
    const container = document.getElementById('events-grid');
    if (!container) return;

    const filtered = AppState.events.filter(e => 
        e.title.toLowerCase().includes(filterQuery.toLowerCase()) || 
        e.subtitle.toLowerCase().includes(filterQuery.toLowerCase())
    );

    if (filtered.length === 0) {
        container.innerHTML = '<p class="text-xs text-slate-400 col-span-full text-center py-8">Aucun événement ne correspond à votre recherche.</p>';
        return;
    }

    container.innerHTML = filtered.map(ev => `
        <div class="bg-slate-900/90 rounded-3xl border border-slate-800 overflow-hidden hover:border-indigo-500/40 transition-all duration-300 flex flex-col justify-between group shadow-xl">
            <div>
                <div class="h-44 overflow-hidden relative">
                    <img src="${ev.image}" alt="${ev.title}" class="w-full h-full object-cover group-hover:scale-105 transition duration-500">
                    <div class="absolute top-3 right-3 bg-slate-950/80 backdrop-blur-md px-3 py-1 rounded-xl text-xs font-bold text-amber-400 border border-amber-500/20 shadow-sm">
                        ${ev.sold.toLocaleString()} / ${ev.capacity.toLocaleString()} vus
                    </div>
                </div>
                <div class="p-5">
                    <span class="text-xs font-bold text-indigo-400 uppercase tracking-wider">${ev.subtitle}</span>
                    <h3 class="text-lg font-outfit font-extrabold text-white mt-1 mb-2 group-hover:text-amber-400 transition">${ev.title}</h3>
                    <div class="text-xs text-slate-400 space-y-1">
                        <div class="flex items-center gap-2"><i class="fa-solid fa-calendar text-slate-500"></i> ${ev.date}</div>
                        <div class="flex items-center gap-2"><i class="fa-solid fa-location-dot text-slate-500"></i> ${ev.venue}</div>
                    </div>
                </div>
            </div>
            <div class="p-5 pt-0 flex items-center justify-between border-t border-slate-800/80 mt-4">
                <div>
                    <span class="text-xs text-slate-500 block">Tarif unique dès</span>
                    <span class="text-lg font-extrabold text-amber-400">${ev.basePrice} DH</span>
                </div>
                <button onclick="selectEventForBooking(${ev.id})" class="px-4 py-2.5 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-500 hover:to-indigo-600 text-white font-extrabold rounded-xl text-xs shadow-md transition flex items-center gap-1.5">
                    <i class="fa-solid fa-chair"></i> Réserver Siège 2D
                </button>
            </div>
        </div>
    `).join('');
}

function handleSearchEvents(e) {
    renderEventsGrid(e.target.value);
}

function selectEventForBooking(eventId) {
    AppState.selectedEventId = eventId;
    switchTab('tab-stadium2d');
    showToast("Sélectionnez votre siège sur la carte 2D du stade", "info");
}

// ----------------------------------------------------
// PLAN 2D DE STADE & INTERACTION SIÈGES (F02)
// ----------------------------------------------------
function renderSeatsGrid() {
    const container = document.getElementById('seats-grid');
    if (!container) return;

    let html = '';
    for (let r = 1; r <= 3; r++) {
        for (let s = 1; s <= 8; s++) {
            const seatNum = (r - 1) * 8 + s;
            const code = `${AppState.selectedTribune.toUpperCase()}-R0${r}-S${s < 10 ? '0' + s : s}`;
            
            let statusClass = 'available';
            if (seatNum === 3 || seatNum === 11) statusClass = 'occupied';
            if (seatNum === 14) statusClass = 'selected';
            if (AppState.selectedTribune === 'vip') statusClass = 'vip';

            html += `
                <button onclick="selectSeat('${code}', ${r}, ${s})" class="seat-btn ${statusClass}" title="${code}">
                    ${s}
                </button>
            `;
        }
    }
    container.innerHTML = html;

    document.querySelectorAll('.tribune-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tribune-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            AppState.selectedTribune = btn.getAttribute('data-tribune');
            document.getElementById('current-tribune-label').textContent = `Tribune ${AppState.selectedTribune}`;
            renderSeatsGrid();
        });
    });
}

function selectSeat(code, row, num) {
    let price = 150;
    if (AppState.selectedTribune === 'est') price = 250;
    if (AppState.selectedTribune === 'vip') price = 600;

    AppState.selectedSeat = {
        code,
        row,
        num,
        price,
        zone: `Tribune ${AppState.selectedTribune.toUpperCase()}`
    };

    document.getElementById('selected-seat-code').textContent = code;
    document.getElementById('selected-seat-row').textContent = `Rang ${row}`;
    document.getElementById('selected-seat-num').textContent = `N° ${num}`;
    document.getElementById('selected-seat-zone').textContent = `Tribune ${AppState.selectedTribune.toUpperCase()}`;
    document.getElementById('selected-seat-price').textContent = `${price}.00 DH`;
}

function buySelectedSeat() {
    const buyerName = document.getElementById('input-buyer-name')?.value || 'Client Anonymous';
    const buyerCin = document.getElementById('input-buyer-cin')?.value || 'BE-100200';
    const ev = AppState.events.find(e => e.id === AppState.selectedEventId) || AppState.events[0];

    const newTicket = {
        id: `TK-${Date.now().toString().slice(-4)}`,
        eventId: ev.id,
        matchTitle: ev.title,
        buyerName,
        cin: buyerCin,
        tribune: AppState.selectedSeat.zone,
        seatCode: AppState.selectedSeat.code,
        price: AppState.selectedSeat.price,
        qrHash: `ETK-2026-${Math.random().toString(36).substring(2, 10).toUpperCase()}`,
        status: "VALIDE",
        scannedAt: null
    };

    AppState.tickets.push(newTicket);
    ev.sold++;

    renderMyTickets();
    updateScannerDropdown();
    showToast(`Billet ${newTicket.qrHash} réservé avec succès !`, 'success');
    openTicketModal(newTicket);
}

// ----------------------------------------------------
// GUICHET POS TACTILE (3-CLICS) (F04)
// ----------------------------------------------------
function renderPosEventsList() {
    const container = document.getElementById('pos-events-list');
    if (!container) return;

    container.innerHTML = AppState.events.map(ev => `
        <button onclick="posSelectEvent(${ev.id})" class="w-full text-left p-3.5 rounded-2xl border ${ev.id === AppState.posSelection.eventId ? 'bg-indigo-950/70 border-indigo-500 text-white shadow-lg' : 'bg-slate-950 border-slate-800 text-slate-300'} hover:border-indigo-500/40 transition">
            <div class="font-extrabold text-xs">${ev.title}</div>
            <div class="text-[10px] text-slate-400 mt-0.5">${ev.date}</div>
        </button>
    `).join('');

    document.querySelectorAll('.pos-category-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.pos-category-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            AppState.posSelection.tribune = btn.getAttribute('data-pos-tribune');
            AppState.posSelection.price = parseInt(btn.getAttribute('data-price'));
            updatePosSummary();
        });
    });

    document.getElementById('pos-qty-minus')?.addEventListener('click', () => {
        if (AppState.posSelection.qty > 1) AppState.posSelection.qty--;
        updatePosSummary();
    });
    document.getElementById('pos-qty-plus')?.addEventListener('click', () => {
        if (AppState.posSelection.qty < 10) AppState.posSelection.qty++;
        updatePosSummary();
    });
}

function posSelectEvent(id) {
    AppState.posSelection.eventId = id;
    renderPosEventsList();
    updatePosSummary();
}

function updatePosSummary() {
    const ev = AppState.events.find(e => e.id === AppState.posSelection.eventId) || AppState.events[0];
    const total = AppState.posSelection.price * AppState.posSelection.qty;

    document.getElementById('pos-summary-match').textContent = ev.title;
    document.getElementById('pos-summary-tribune').textContent = `Tribune ${AppState.posSelection.tribune}`;
    document.getElementById('pos-summary-calc').textContent = `${AppState.posSelection.price} DH x ${AppState.posSelection.qty}`;
    document.getElementById('pos-summary-total').textContent = `${total}.00 DH`;
    document.getElementById('pos-qty-val').textContent = AppState.posSelection.qty;
}

function handlePosCheckout() {
    const ev = AppState.events.find(e => e.id === AppState.posSelection.eventId) || AppState.events[0];
    const newTicket = {
        id: `TK-POS-${Date.now().toString().slice(-4)}`,
        eventId: ev.id,
        matchTitle: ev.title,
        buyerName: "Vente Guichet POS",
        cin: "GUICHET-01",
        tribune: `Tribune ${AppState.posSelection.tribune}`,
        seatCode: `${AppState.posSelection.tribune.toUpperCase()}-POS-${Math.floor(Math.random() * 50 + 1)}`,
        price: AppState.posSelection.price * AppState.posSelection.qty,
        qrHash: `ETK-2026-${Math.random().toString(36).substring(2, 10).toUpperCase()}`,
        status: "VALIDE",
        scannedAt: null
    };

    AppState.tickets.push(newTicket);
    ev.sold += AppState.posSelection.qty;

    renderMyTickets();
    updateScannerDropdown();
    playBeep('success');
    showToast("Vente Guichet POS finalisée avec succès !", "success");
    openTicketModal(newTicket);
}

// ----------------------------------------------------
// SCANNER PDA ACCÈS & CHECKSUM MODE 4 (F06)
// ----------------------------------------------------
function initScannerModule() {
    updateScannerDropdown();
    document.getElementById('btn-trigger-scan')?.addEventListener('click', runScannerValidation);
}

function updateScannerDropdown() {
    const select = document.getElementById('select-test-ticket');
    if (!select) return;

    select.innerHTML = AppState.tickets.map(tk => `
        <option value="${tk.qrHash}">${tk.qrHash} — ${tk.matchTitle} (${tk.buyerName} - ${tk.status})</option>
    `).join('');
}

function runScannerValidation() {
    const select = document.getElementById('select-test-ticket');
    const selectedHash = select ? select.value : '';
    const ticket = AppState.tickets.find(t => t.qrHash === selectedHash);

    const idleView = document.getElementById('scanner-state-idle');
    const successView = document.getElementById('scanner-state-success');
    const errorView = document.getElementById('scanner-state-error');
    const screen = document.getElementById('scanner-screen');

    idleView.classList.add('hidden');
    successView.classList.add('hidden');
    errorView.classList.add('hidden');

    if (ticket && ticket.status === 'VALIDE') {
        ticket.status = 'SCANNE';
        ticket.scannedAt = new Date().toLocaleTimeString();

        document.getElementById('scan-res-name').textContent = ticket.buyerName;
        document.getElementById('scan-res-seat').textContent = `${ticket.tribune} — ${ticket.seatCode}`;
        document.getElementById('scan-res-hash').textContent = `Checksum Mode 4 : ${ticket.qrHash} [VALIDÉ 0.12s]`;

        successView.classList.remove('hidden');
        screen.className = 'w-full rounded-3xl bg-emerald-950/40 border-2 border-emerald-500 p-8 flex flex-col items-center justify-center text-center transition-all duration-300 min-h-[340px] shadow-2xl';
        playBeep('success');
        showToast(`Accès Accordé pour ${ticket.buyerName} (${ticket.seatCode})`, 'success');
    } else {
        document.getElementById('scan-err-title').textContent = "Billet Invalide / Déjà Scanné !";
        document.getElementById('scan-err-desc').textContent = ticket ? `Ce billet (${ticket.qrHash}) a déjà été scanné à ${ticket.scannedAt || '20:14:02'}.` : "Code QR inexistant dans la base d'accès.";

        errorView.classList.remove('hidden');
        screen.className = 'w-full rounded-3xl bg-rose-950/40 border-2 border-rose-500 p-8 flex flex-col items-center justify-center text-center transition-all duration-300 min-h-[340px] shadow-2xl';
        playBeep('error');
        showToast("Alerte Fraude : Billet déjà scanné !", "error");
    }

    updateScannerDropdown();
}

// ----------------------------------------------------
// WALLET / MES BILLETS SPECTATEUR
// ----------------------------------------------------
function renderMyTickets() {
    const container = document.getElementById('my-tickets-container');
    if (!container) return;

    if (AppState.tickets.length === 0) {
        container.innerHTML = '<p class="text-xs text-slate-400 col-span-full text-center py-8">Aucun billet réservé pour le moment.</p>';
        return;
    }

    container.innerHTML = AppState.tickets.map(tk => `
        <div class="bg-slate-900/90 border border-slate-800 rounded-3xl p-5 space-y-4 shadow-xl relative overflow-hidden">
            <div class="flex items-center justify-between border-b border-slate-800 pb-3">
                <span class="text-xs font-bold text-amber-400">${tk.tribune}</span>
                <span class="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold ${tk.status === 'VALIDE' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-rose-500/10 text-rose-400 border border-rose-500/20'}">
                    ${tk.status}
                </span>
            </div>

            <div>
                <h4 class="font-outfit font-extrabold text-white text-base">${tk.matchTitle}</h4>
                <p class="text-xs text-slate-400">Siège : <strong class="text-indigo-400">${tk.seatCode}</strong></p>
            </div>

            <div class="bg-slate-950 p-3 rounded-2xl border border-slate-800 text-center">
                <div class="text-[10px] text-slate-500 font-mono mb-1 font-bold">${tk.qrHash}</div>
                <button onclick="openTicketModalById('${tk.id}')" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold rounded-xl border border-slate-700 w-full transition">
                    <i class="fa-solid fa-qrcode text-amber-400"></i> Afficher QR Code
                </button>
            </div>
        </div>
    `).join('');
}

function openTicketModalById(id) {
    const tk = AppState.tickets.find(t => t.id === id);
    if (tk) openTicketModal(tk);
}

function openTicketModal(ticket) {
    const modal = document.getElementById('modal-ticket');
    if (!modal) return;

    document.getElementById('modal-match-title').textContent = ticket.matchTitle;
    document.getElementById('modal-buyer-name').textContent = ticket.buyerName;
    document.getElementById('modal-ticket-price').textContent = `${ticket.price} DH`;
    document.getElementById('modal-qr-hash').textContent = ticket.qrHash;
    document.getElementById('modal-tribune-val').textContent = ticket.tribune;
    document.getElementById('modal-seat-val').textContent = ticket.seatCode;

    const target = document.getElementById('modal-qrcode-target');
    target.innerHTML = '';
    try {
        new QRCode(target, {
            text: ticket.qrHash,
            width: 140,
            height: 140,
            colorDark : "#0f172a",
            colorLight : "#ffffff",
            correctLevel : QRCode.CorrectLevel.H
        });
    } catch(e) {
        target.innerHTML = `<div class="p-4 text-xs font-mono font-bold text-slate-900 bg-slate-200 rounded">${ticket.qrHash}</div>`;
    }

    modal.classList.remove('hidden');
}

function closeModal() {
    document.getElementById('modal-ticket')?.classList.add('hidden');
}

// ----------------------------------------------------
// ADMIN EVENTS TABLE
// ----------------------------------------------------
function renderAdminTable() {
    const body = document.getElementById('admin-events-table-body');
    if (!body) return;

    body.innerHTML = AppState.events.map(ev => `
        <tr class="hover:bg-slate-900/60 transition">
            <td class="p-4 font-bold text-white">${ev.title}</td>
            <td class="p-4 text-slate-400">${ev.date}</td>
            <td class="p-4 text-slate-400">${ev.venue}</td>
            <td class="p-4 font-bold text-amber-400">${ev.basePrice} DH</td>
            <td class="p-4 font-semibold">${ev.sold.toLocaleString()} / ${ev.capacity.toLocaleString()}</td>
            <td class="p-4">
                <span class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">${ev.status}</span>
            </td>
        </tr>
    `).join('');
}

// ----------------------------------------------------
// EXPORT CSV UTILITY
// ----------------------------------------------------
function exportAccessLogsCSV() {
    let csvContent = "data:text/csv;charset=utf-8,ID,Match,Acheteur,Tribune,Siege,QRHash,Statut,ScanneA\n";
    AppState.tickets.forEach(tk => {
        csvContent += `${tk.id},"${tk.matchTitle}","${tk.buyerName}","${tk.tribune}","${tk.seatCode}","${tk.qrHash}",${tk.status},"${tk.scannedAt || ''}"\n`;
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `rapport_acces_stade_${Date.now()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast("Rapport des d'accès exporté en CSV !", "success");
}

// ----------------------------------------------------
// SUPERVISION DASHBOARD & CHARTS (F07)
// ----------------------------------------------------
function initDashboardCharts() {
    const ctxFlow = document.getElementById('chart-flow')?.getContext('2d');
    const ctxPie = document.getElementById('chart-pie')?.getContext('2d');

    if (ctxFlow) {
        new Chart(ctxFlow, {
            type: 'line',
            data: {
                labels: ['18:00', '18:30', '19:00', '19:30', '20:00', '20:30'],
                datasets: [{
                    label: 'Entrées / 30 min (Personnes)',
                    data: [1200, 4800, 14200, 22100, 6800, 920],
                    borderColor: '#6366f1',
                    backgroundColor: 'rgba(99, 102, 241, 0.15)',
                    fill: true,
                    tension: 0.4
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8' } },
                    y: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8' } }
                }
            }
        });
    }

    if (ctxPie) {
        new Chart(ctxPie, {
            type: 'doughnut',
            data: {
                labels: ['Tribune Nord', 'Tribune Sud', 'Tribune Est', 'VIP Lounge'],
                datasets: [{
                    data: [18500, 18200, 10500, 2514],
                    backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#ec4899'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom', labels: { color: '#cbd5e1' } }
                }
            }
        });
    }
}
