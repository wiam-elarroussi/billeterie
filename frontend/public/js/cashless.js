/* E-Ticket-Pay Cashless Wallet & NFC Engine */

window.CashlessEngine = {
  balance: 450.00,
  history: [
    { id: 'TX-9012', time: '22:15:00', type: 'Rechargement Web', amount: '+200.00 DH', status: 'Complété' },
    { id: 'TX-8910', time: '21:40:12', type: 'Buvette Trib. Nord', amount: '-45.00 DH', status: 'Complété' },
    { id: 'TX-7811', time: '20:10:05', type: 'Achat Programme', amount: '-30.00 DH', status: 'Complété' }
  ],

  init() {
    this.bindEvents();
    this.renderHistory();
  },

  bindEvents() {
    document.getElementById('btn-recharge-100')?.addEventListener('click', () => this.recharge(100));
    document.getElementById('btn-recharge-200')?.addEventListener('click', () => this.recharge(200));
    document.getElementById('btn-recharge-500')?.addEventListener('click', () => this.recharge(500));

    const nfcTrigger = document.getElementById('nfc-tap-trigger');
    if (nfcTrigger) {
      nfcTrigger.addEventListener('click', () => {
        this.payNfc(45.00, 'Buvette Trib. Est');
      });
    }
  },

  recharge(amount) {
    this.balance += amount;
    this.updateBalanceDisplay();
    
    const now = new Date().toLocaleTimeString('fr-FR');
    const txId = 'TX-' + Math.floor(1000 + Math.random() * 9000);
    this.history.unshift({
      id: txId,
      time: now,
      type: 'Rechargement Carte NFC',
      amount: `+${amount.toFixed(2)} DH`,
      status: 'Complété'
    });

    this.renderHistory();
    window.EProApp.showToast(`Solde Cashless rechargé de +${amount} DH avec succès!`);
  },

  payNfc(amount, location) {
    if (this.balance < amount) {
      window.EProApp.showToast('Solde Cashless insuffisant pour effectuer le paiement !', true);
      return;
    }

    this.balance -= amount;
    this.updateBalanceDisplay();

    const nfcBtn = document.getElementById('nfc-tap-trigger');
    nfcBtn.style.transform = 'scale(1.2) rotate(15deg)';
    setTimeout(() => { nfcBtn.style.transform = 'none'; }, 300);

    const now = new Date().toLocaleTimeString('fr-FR');
    const txId = 'TX-' + Math.floor(1000 + Math.random() * 9000);
    this.history.unshift({
      id: txId,
      time: now,
      type: location,
      amount: `-${amount.toFixed(2)} DH`,
      status: 'Complété'
    });

    this.renderHistory();
    window.EProApp.showToast(`Paiement Cashless ${amount} DH validé sur ${location}`);
  },

  updateBalanceDisplay() {
    const el = document.getElementById('cashless-balance');
    if (el) el.textContent = `${this.balance.toFixed(2)} DH`;
  },

  renderHistory() {
    const tbody = document.getElementById('cashless-history-rows');
    if (!tbody) return;

    tbody.innerHTML = '';
    this.history.forEach(tx => {
      const tr = document.createElement('tr');
      const isPositive = tx.amount.startsWith('+');
      tr.innerHTML = `
        <td><strong style="font-family: var(--font-mono);">${tx.id}</strong></td>
        <td>${tx.time}</td>
        <td><span class="badge" style="background: rgba(255,255,255,0.05);">Card #NFC-8824</span></td>
        <td>${tx.type}</td>
        <td style="font-weight: 700; color: ${isPositive ? 'var(--status-success)' : 'var(--status-danger)'};">${tx.amount}</td>
        <td><span style="color: var(--status-success); font-weight: 600;">● ${tx.status}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }
};
