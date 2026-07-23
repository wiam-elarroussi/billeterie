/* Stadium 2D Interactive Map Engine */

window.Stadium2D = {
  selectedSeat: null,
  currentStand: 'Nord',
  currentCategory: 'Catégorie 1',
  currentPrice: 350,

  init() {
    this.bindEvents();
    this.renderSeats('Nord', 'Catégorie 1', 350);
  },

  bindEvents() {
    const sectionPaths = document.querySelectorAll('.section-path');
    sectionPaths.forEach(path => {
      path.addEventListener('click', (e) => {
        const stand = e.currentTarget.dataset.stand;
        const cat = e.currentTarget.dataset.cat;
        const price = parseInt(e.currentTarget.dataset.price, 10);

        this.currentStand = stand;
        this.currentCategory = cat;
        this.currentPrice = price;

        document.getElementById('selected-stand-name').textContent = `Sélection de Siège - Tribune ${stand}`;
        document.getElementById('stand-cat-display').textContent = cat;
        document.getElementById('stand-price-display').textContent = `${price} DH`;

        this.renderSeats(stand, cat, price);
      });
    });

    const bookBtn = document.getElementById('btn-book-selected-seat');
    if (bookBtn) {
      bookBtn.addEventListener('click', () => {
        if (!this.selectedSeat) return;
        
        // Populate POS form with selected seat info
        document.getElementById('pos-category-select').value = this.getPosCategoryValue(this.currentStand);
        document.getElementById('ticket-seat-txt').textContent = this.selectedSeat;
        document.getElementById('ticket-stand-txt').textContent = `${this.currentStand.toUpperCase()} (${this.currentCategory.toUpperCase()})`;
        document.getElementById('ticket-price-txt').textContent = `${this.currentPrice} DH`;
        
        // Switch to POS tab
        document.querySelector('[data-tab="tab-pos"]').click();
        
        // Toast
        window.EProApp.showToast(`Siège ${this.selectedSeat} sélectionné pour le guichet de vente!`);
      });
    }
  },

  getPosCategoryValue(stand) {
    switch (stand) {
      case 'Nord': return 'nord-350';
      case 'Sud': return 'sud-200';
      case 'Est': return 'est-1200';
      case 'Ouest': return 'ouest-350';
      default: return 'nord-350';
    }
  },

  renderSeats(stand, category, price) {
    const grid = document.getElementById('seat-grid');
    grid.innerHTML = '';
    this.selectedSeat = null;
    document.getElementById('btn-book-selected-seat').disabled = true;

    const prefix = stand.charAt(0).toUpperCase();

    for (let i = 1; i <= 30; i++) {
      const seatId = `${prefix}-${i < 10 ? '0' + i : i}`;
      const item = document.createElement('div');
      item.className = 'seat-item';
      
      // Simulate occupied seats (e.g. seats divisible by 4 or 7)
      if (i % 4 === 0 || i % 7 === 0) {
        item.classList.add('taken');
        item.textContent = `${seatId} (OCC)`;
      } else if (category.includes('VIP')) {
        item.classList.add('vip');
        item.textContent = `${seatId} (VIP)`;
      } else {
        item.textContent = seatId;
      }

      if (!item.classList.contains('taken')) {
        item.addEventListener('click', () => {
          document.querySelectorAll('.seat-item').forEach(s => s.classList.remove('selected'));
          item.classList.add('selected');
          this.selectedSeat = seatId;
          document.getElementById('btn-book-selected-seat').disabled = false;
        });
      }

      grid.appendChild(item);
    }
  }
};
