/* E-Ticket Access Control Scanner Engine (0.5s Latency) */

window.AccessScanner = {
  scannedTickets: new Set(['TCK-ALREADY-USED']),

  init() {
    this.bindEvents();
  },

  bindEvents() {
    const scanBtn = document.getElementById('btn-trigger-scan');
    const input = document.getElementById('pda-scan-input');
    
    if (scanBtn && input) {
      scanBtn.addEventListener('click', () => {
        this.verifyTicket(input.value || 'TCK-8891-VAL');
      });
      
      input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
          this.verifyTicket(input.value || 'TCK-8891-VAL');
        }
      });
    }

    const testFraud = document.getElementById('btn-test-fraud');
    if (testFraud) {
      testFraud.addEventListener('click', () => {
        this.verifyTicket('TCK-ALREADY-USED');
      });
    }

    const testValid = document.getElementById('btn-test-valid');
    if (testValid) {
      testValid.addEventListener('click', () => {
        const randomCode = 'TCK-' + Math.floor(1000 + Math.random() * 9000) + '-VAL';
        this.verifyTicket(randomCode);
      });
    }
  },

  verifyTicket(code) {
    const feedbackBox = document.getElementById('scanner-feedback-box');
    const startTime = performance.now();

    feedbackBox.className = 'scanner-feedback';
    feedbackBox.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Traitement Checksum Mode 4 en cours...`;

    // Simulate 0.38s - 0.48s execution latency (< 0.5s FIFA Requirement)
    setTimeout(() => {
      const elapsed = ((performance.now() - startTime) / 1000).toFixed(2);

      if (this.scannedTickets.has(code)) {
        // Anti-passback triggered / Already scanned
        feedbackBox.className = 'scanner-feedback denied';
        feedbackBox.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> REFUS ACCÈS : FRAUDE ANTI-PASSBACK ! Code [${code}] déjà utilisé sur Porte A. (${elapsed}s)`;
        this.playBeep(false);
        window.EProApp.addLog(`[FRAUDE] Code ${code} refusé aux tourniquets Porte A`, 'danger');
        window.EProApp.incrementFrauds();
      } else {
        // Valid Ticket
        this.scannedTickets.add(code);
        feedbackBox.className = 'scanner-feedback granted';
        feedbackBox.innerHTML = `<i class="fa-solid fa-circle-check"></i> ACCÈS AUTORISÉ • PORTE A • CODE [${code}] (${elapsed}s)`;
        this.playBeep(true);
        window.EProApp.addLog(`[VALIDE] Entry scanné code ${code} sur Porte A`, 'success');
        window.EProApp.incrementOccupancy();
      }
    }, 420);
  },

  playBeep(success) {
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);

      if (success) {
        osc.frequency.value = 880; // High beep A5
        gain.gain.setValueAtTime(0.1, ctx.currentTime);
        osc.start();
        osc.stop(ctx.currentTime + 0.15);
      } else {
        osc.frequency.value = 220; // Low beep A3
        gain.gain.setValueAtTime(0.2, ctx.currentTime);
        osc.start();
        osc.stop(ctx.currentTime + 0.4);
      }
    } catch(e) {
      console.log('Audio Context not available');
    }
  }
};
