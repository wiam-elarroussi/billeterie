/* E-Ticket Pro - Main Application Controller */

window.EProApp = {
  occupancy: 67420,
  maxCapacity: 80000,
  frauds: 14,
  revenue: 4850200,
  entriesChart: null,
  salesPieChart: null,
  currentRole: 'super-admin',

  init() {
    this.startClock();
    this.bindTabNavigation();
    this.initRoleSelector();
    this.initCharts();
    this.renderGatesGrid();
    this.initPosForm();
    this.drawTicketQr('TCK-8891-VAL-NORD-N45');

    // Init sub-engines
    window.Stadium2D?.init();
    window.AccessScanner?.init();
    window.CashlessEngine?.init();

    // Export report
    document.getElementById('export-report-btn')?.addEventListener('click', () => this.exportCpsReport());
    document.getElementById('simulate-entries-btn')?.addEventListener('click', () => this.simulateFlux());

    // Auto update live simulation ticks
    setInterval(() => this.liveTick(), 3000);
  },

  startClock() {
    const clockEl = document.getElementById('clock-display');
    const updateTime = () => {
      const now = new Date();
      clockEl.textContent = now.toLocaleTimeString('fr-FR');
    };
    updateTime();
    setInterval(updateTime, 1000);
  },

  bindTabNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    const tabPanes = document.querySelectorAll('.tab-pane');

    const tabSubtitles = {
      'tab-dashboard': 'Monitoring du Grand Stade de Casablanca • Match de Gala 2026',
      'tab-stadium': 'Configuration des Tribunes, Zones d\'accès et Numérotation des Sièges',
      'tab-pos': 'Point de Vente Guichet Vente Rapide en 3 Clics pour Écran Tactile',
      'tab-cashless': 'Gestion du Portefeuille Dématérialisé, NFC & Mode Hors-Ligne',
      'tab-access': 'Validation des Billets aux Tourniquets PDA (Exigence FIFA < 0.5s)',
      'tab-backoffice': 'Gestion des Rôles, Matrice des Droits & Serveur Dell PowerEdge R360'
    };

    navItems.forEach(item => {
      item.addEventListener('click', () => {
        const targetTab = item.dataset.tab;
        
        navItems.forEach(n => n.classList.remove('active'));
        tabPanes.forEach(p => p.classList.remove('active'));

        item.classList.add('active');
        document.getElementById(targetTab)?.classList.add('active');

        // Update Title & Subtitle
        const titleText = item.querySelector('span').textContent;
        document.getElementById('current-tab-title').textContent = titleText;
        document.getElementById('current-tab-sub').textContent = tabSubtitles[targetTab] || '';
      });
    });
  },

  initRoleSelector() {
    const selector = document.getElementById('user-role-select');
    if (!selector) return;

    selector.addEventListener('change', (e) => {
      this.currentRole = e.target.value;
      this.showToast(`Rôle basculé vers : ${selector.options[selector.selectedIndex].text}`);
    });
  },

  initCharts() {
    // 1. Line Chart: Entries per Minute
    const ctxLine = document.getElementById('entriesChart')?.getContext('2d');
    if (ctxLine) {
      this.entriesChart = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: ['21:30', '21:40', '21:50', '22:00', '22:10', '22:20', '22:30'],
          datasets: [{
            label: 'Entrées / Minute (Toutes Portes)',
            data: [320, 450, 680, 1120, 950, 780, 640],
            borderColor: '#06B6D4',
            backgroundColor: 'rgba(6, 182, 212, 0.15)',
            fill: true,
            tension: 0.4,
            borderWidth: 3
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { labels: { color: '#9CA3AF' } } },
          scales: {
            x: { ticks: { color: '#6B7280' }, grid: { color: 'rgba(255,255,255,0.05)' } },
            y: { ticks: { color: '#6B7280' }, grid: { color: 'rgba(255,255,255,0.05)' } }
          }
        }
      });
    }

    // 2. Pie Chart: Sales Breakdown
    const ctxPie = document.getElementById('salesPieChart')?.getContext('2d');
    if (ctxPie) {
      this.salesPieChart = new Chart(ctxPie, {
        type: 'doughnut',
        data: {
          labels: ['Cat. 1 (Nord/Ouest)', 'Cat. 2 (Sud)', 'VIP & Loges (Est)', 'Abonnés Season'],
          datasets: [{
            data: [38000, 24000, 5420, 12000],
            backgroundColor: ['#06B6D4', '#3B82F6', '#FFB800', '#10B981'],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { position: 'bottom', labels: { color: '#9CA3AF', padding: 14 } } }
        }
      });
    }
  },

  renderGatesGrid() {
    const grid = document.getElementById('gates-grid');
    if (!grid) return;

    const gates = [
      { name: 'PORTE A (Nord)', status: 'Optimal', throughput: '1,240 p/h', color: '#10B981' },
      { name: 'PORTE B (Sud)', status: 'Optimal', throughput: '980 p/h', color: '#10B981' },
      { name: 'PORTE C (Est - VIP)', status: 'Fluidifié', throughput: '320 p/h', color: '#FFB800' },
      { name: 'PORTE D (Ouest)', status: 'Optimal', throughput: '1,150 p/h', color: '#10B981' }
    ];

    grid.innerHTML = gates.map(g => `
      <div style="background: rgba(255,255,255,0.03); padding: 14px; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
        <div style="display: flex; justify-content: space-between; font-size: 13px; font-weight: 700;">
          <span>${g.name}</span>
          <span style="color: ${g.color};">● ${g.status}</span>
        </div>
        <div style="font-size: 11px; color: var(--text-dim); margin-top: 6px;">Débit moyen: <strong>${g.throughput}</strong></div>
      </div>
    `).join('');
  },

  initPosForm() {
    const form = document.getElementById('pos-sale-form');
    const catSelect = document.getElementById('pos-category-select');
    const qtyInput = document.getElementById('pos-qty-input');
    const buyerName = document.getElementById('pos-buyer-name');
    const totalDisplay = document.getElementById('pos-total-display');

    const updatePrice = () => {
      const price = parseInt(catSelect.value.split('-')[1], 10);
      const qty = parseInt(qtyInput.value, 10) || 1;
      const total = price * qty;
      totalDisplay.textContent = `${total} DH`;

      // Update thermal preview
      const catText = catSelect.options[catSelect.selectedIndex].text;
      document.getElementById('ticket-stand-txt').textContent = catText.split(' (')[0];
      document.getElementById('ticket-price-txt').textContent = `${total} DH`;
    };

    catSelect?.addEventListener('change', updatePrice);
    qtyInput?.addEventListener('input', updatePrice);
    buyerName?.addEventListener('input', () => {
      document.getElementById('ticket-name-txt').textContent = buyerName.value.toUpperCase();
    });

    form?.addEventListener('submit', (e) => {
      e.preventDefault();
      
      const ticketCode = 'TCK-' + Math.floor(1000 + Math.random() * 9000) + '-VAL';
      document.getElementById('ticket-checksum-txt').textContent = `CHECKSUM: ${ticketCode}-MODE4`;
      this.drawTicketQr(ticketCode);

      this.revenue += parseInt(totalDisplay.textContent, 10);
      document.getElementById('stat-revenue').textContent = `${this.revenue.toLocaleString()} DH`;

      this.addLog(`[VENTE GUICHET] Billet ${ticketCode} vendu à ${buyerName.value}`, 'success');
      this.showToast(`Billet généré et validé avec succès! Emplacement prêt pour impression.`);

      // Trigger auto window print preview option
      window.print();
    });
  },

  drawTicketQr(text) {
    const canvas = document.getElementById('ticket-qr-canvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    
    // Simple fast QR code pattern generator on canvas
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, 120, 120);
    ctx.fillStyle = '#000000';

    const size = 10;
    for (let r = 0; r < 12; r++) {
      for (let c = 0; c < 12; c++) {
        if ((r < 3 && c < 3) || (r < 3 && c > 8) || (r > 8 && c < 3) || (r + c) % 3 === 0) {
          ctx.fillRect(c * size, r * size, size, size);
        }
      }
    }
  },

  addLog(msg, type = 'info') {
    const container = document.getElementById('live-scan-logs');
    if (!container) return;

    const time = new Date().toLocaleTimeString('fr-FR');
    const colorMap = { success: 'var(--status-success)', danger: 'var(--status-danger)', info: 'var(--accent-cyan)' };

    const item = document.createElement('div');
    item.style.color = colorMap[type] || 'var(--text-main)';
    item.innerHTML = `[${time}] ${msg}`;
    
    container.prepend(item);
    if (container.children.length > 20) container.removeChild(container.lastChild);
  },

  incrementOccupancy() {
    this.occupancy += 1;
    document.getElementById('stat-occupancy').textContent = this.occupancy.toLocaleString();
    const pct = ((this.occupancy / this.maxCapacity) * 100).toFixed(1);
    document.getElementById('live-gauge-badge').textContent = `${pct}%`;
  },

  incrementFrauds() {
    this.frauds += 1;
    document.getElementById('stat-frauds').textContent = this.frauds;
  },

  simulateFlux() {
    if (this.entriesChart) {
      const data = this.entriesChart.data.datasets[0].data;
      data.shift();
      data.push(Math.floor(500 + Math.random() * 700));
      this.entriesChart.update();
    }
    this.incrementOccupancy();
    this.addLog(`[FLUX] Traitement rafale 14 billets scannés en 0.41s`, 'info');
    this.showToast('Flux d\'entrée rafraîchi en direct');
  },

  liveTick() {
    if (Math.random() > 0.4) {
      this.incrementOccupancy();
      const code = 'TCK-' + Math.floor(1000 + Math.random() * 9000);
      this.addLog(`[AUTO-SCAN] Tourniquet P-02 scanné avec succès (0.39s) - Billet ${code}`, 'success');
    }
  },

  exportCpsReport() {
    const reportData = {
      system: 'E-Ticket Pro CMS Billettique',
      standard: 'FIFA Ready / Appel d\'Offres N° 02/2026/GSC',
      timestamp: new Date().toISOString(),
      stadium: 'Grand Stade de Casablanca',
      totalOccupancy: this.occupancy,
      capacity: this.maxCapacity,
      totalRevenue: `${this.revenue} DH`,
      fraudCount: this.frauds,
      serverHardware: 'Dell PowerEdge R360 (iDRAC9 OK)'
    };

    const blob = new Blob([JSON.stringify(reportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `E-Ticket-Pro-Rapport-CPS-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);

    this.showToast('Rapport de conformité CPS exporté avec succès !');
  },

  showToast(message, isError = false) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
      toast.style.cssText = `
        position: fixed;
        bottom: 24px;
        right: 24px;
        padding: 14px 22px;
        border-radius: 10px;
        background: #111827;
        color: #FFF;
        font-family: var(--font-main);
        font-size: 13px;
        font-weight: 600;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
        border: 1px solid var(--primary-gold);
        z-index: 999;
        transition: all 0.3s ease;
      `;
      document.body.appendChild(toast);
    }

    toast.style.borderColor = isError ? 'var(--status-danger)' : 'var(--primary-gold)';
    toast.innerHTML = `<i class="fa-solid ${isError ? 'fa-triangle-exclamation' : 'fa-circle-check'}" style="color: ${isError ? 'var(--status-danger)' : 'var(--primary-gold)'}; margin-right: 8px;"></i> ${message}`;
    toast.style.opacity = '1';

    setTimeout(() => {
      toast.style.opacity = '0';
    }, 3500);
  }
};

document.addEventListener('DOMContentLoaded', () => {
  window.EProApp.init();
});
