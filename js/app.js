/**
 * NIKE MEN'S EXCLUSIVE 2026 - APPLICATION CONTROLLER
 * Full Interactive Logic: Catalog, Filters, Search, Cart, Wishlist, Modals, Currency
 */

(function () {
  'use strict';

  // State Management
  const state = {
    products: (typeof NIKE_PRODUCTS !== 'undefined' ? NIKE_PRODUCTS : (window.NIKE_PRODUCTS || [])),
    filteredProducts: [],
    theme: localStorage.getItem('nike_theme') || 'dark', // 'dark' or 'light'
    currency: localStorage.getItem('nike_currency') || 'USD', // 'USD' or 'UZS'
    exchangeRate: 13000, // 1 USD = 13,000 UZS
    activeCategory: 'all', // 'all', 'shoes', 'autumn', 'summer', 'winter'
    activeSubCategory: 'all',
    activeSeason: 'all',
    activeSize: 'all',
    searchQuery: '',
    sortBy: 'featured',
    displayedCount: 24,
    itemsPerPage: 24,
    cart: JSON.parse(localStorage.getItem('nike_cart') || '[]'),
    wishlist: JSON.parse(localStorage.getItem('nike_wishlist') || '[]'),
    appliedPromo: null, // { code: 'NIKE2026', discountPercent: 10 }
    selectedProduct: null,
    selectedSize: null
  };

  // Promo Codes Directory
  const PROMO_CODES = {
    'NIKE2026': { type: 'percent', value: 10, label: '10% Yangi mavsum chegirmasi' },
    'ERKAKLAR': { type: 'percent', value: 15, label: '15% Erkaklar kuni chegirmasi' },
    'JUSTDOIT': { type: 'fixed', value: 20, label: '$20 chegirma' }
  };

  // DOM Elements Cache
  const els = {
    topBar: document.getElementById('topBar'),
    themeToggleBtn: document.getElementById('themeToggleBtn'),
    topBarThemeBtn: document.getElementById('topBarThemeBtn'),
    topThemeText: document.getElementById('topThemeText'),
    currencyBtns: document.querySelectorAll('.currency-btn'),
    navLinks: document.querySelectorAll('.nav-link'),
    catalogTabs: document.querySelectorAll('.catalog-tab'),
    quickFilters: document.querySelectorAll('.quick-filter-btn'),
    // Hero Video Elements
    heroVideoPlayer: document.getElementById('heroVideoPlayer'),
    heroVideoIframe: document.getElementById('heroVideoIframe'),
    videoActiveTitle: document.getElementById('videoActiveTitle'),
    videoMuteToggleBtn: document.getElementById('videoMuteToggleBtn'),
    soundIcon: document.getElementById('soundIcon'),
    soundText: document.getElementById('soundText'),
    videoPlaylistItems: document.querySelectorAll('.video-playlist-item'),
    searchInput: document.getElementById('searchInput'),
    searchClearBtn: document.getElementById('searchClearBtn'),
    sortSelect: document.getElementById('sortSelect'),
    productsGrid: document.getElementById('productsGrid'),
    loadMoreContainer: document.getElementById('loadMoreContainer'),
    loadMoreBtn: document.getElementById('loadMoreBtn'),
    loadMoreInfo: document.getElementById('loadMoreInfo'),
    resultsCount: document.getElementById('resultsCount'),
    activeFilterBadge: document.getElementById('activeFilterBadge'),
    cartBtn: document.getElementById('cartBtn'),
    cartBadge: document.getElementById('cartBadge'),
    wishlistBtn: document.getElementById('wishlistBtn'),
    wishlistBadge: document.getElementById('wishlistBadge'),
    // Drawers
    cartDrawer: document.getElementById('cartDrawer'),
    cartBackdrop: document.getElementById('cartBackdrop'),
    cartCloseBtn: document.getElementById('cartCloseBtn'),
    cartItemsList: document.getElementById('cartItemsList'),
    cartSubtotal: document.getElementById('cartSubtotal'),
    cartDiscountRow: document.getElementById('cartDiscountRow'),
    cartDiscountVal: document.getElementById('cartDiscountVal'),
    cartTotal: document.getElementById('cartTotal'),
    promoInput: document.getElementById('promoInput'),
    promoApplyBtn: document.getElementById('promoApplyBtn'),
    checkoutBtn: document.getElementById('checkoutBtn'),
    // Wishlist Drawer
    wishlistDrawer: document.getElementById('wishlistDrawer'),
    wishlistCloseBtn: document.getElementById('wishlistCloseBtn'),
    wishlistItemsList: document.getElementById('wishlistItemsList'),
    // Modals
    detailModal: document.getElementById('detailModal'),
    detailModalClose: document.getElementById('detailModalClose'),
    detailModalContent: document.getElementById('detailModalContent'),
    checkoutModal: document.getElementById('checkoutModal'),
    checkoutModalClose: document.getElementById('checkoutModalClose'),
    checkoutForm: document.getElementById('checkoutForm'),
    checkoutTotalDisplay: document.getElementById('checkoutTotalDisplay'),
    successModal: document.getElementById('successModal'),
    successModalClose: document.getElementById('successModalClose'),
    successOrderNum: document.getElementById('successOrderNum'),
    // QR Code Elements
    heroQrBadge: document.getElementById('heroQrBadge'),
    floatingQrTrigger: document.getElementById('floatingQrTrigger'),
    qrModal: document.getElementById('qrModal'),
    qrModalClose: document.getElementById('qrModalClose'),
    qrDismissBtn: document.getElementById('qrDismissBtn'),
    qrCopyBtn: document.getElementById('qrCopyBtn'),
    toastContainer: document.getElementById('toastContainer'),
    // Filter Drawer Toggle & Subfilters
    filterToggleBtn: document.getElementById('filterToggleBtn'),
    filterDrawerPanel: document.getElementById('filterDrawerPanel'),
    subCategoryPills: document.getElementById('subCategoryPills'),
    sizePills: document.getElementById('sizePills'),
    resetFiltersBtn: document.getElementById('resetFiltersBtn')
  };

  // ================= UTILITIES =================
  function formatMoney(amountUSD) {
    if (state.currency === 'UZS') {
      const uzs = Math.round(amountUSD * state.exchangeRate);
      return uzs.toLocaleString('uz-UZ') + " so'm";
    }
    return "$" + Number(amountUSD).toFixed(0);
  }

  // Theme Management (Stealth Dark / Studio Light)
  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('nike_theme', theme);

    if (els.topThemeText) {
      els.topThemeText.textContent = theme === 'dark' ? "☀️ Yorug' fon" : "🌙 Qora fon";
    }
  }

  function toggleTheme() {
    const nextTheme = state.theme === 'dark' ? 'light' : 'dark';
    applyTheme(nextTheme);
    showToast(nextTheme === 'dark' ? "🌙 Qora stealth fon yoqildi" : "☀️ Yorug' studio fon yoqildi", "✨");
  }

  function showToast(message, icon = '✓') {
    if (!els.toastContainer) return;
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `<span style="font-weight:900; color:#10b981;">${icon}</span> <span>${message}</span>`;
    els.toastContainer.appendChild(toast);
    
    // Animate in
    setTimeout(() => toast.classList.add('show'), 10);

    // Remove after 3s
    setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => toast.remove(), 300);
    }, 3200);
  }

  function getSeasonLabel(season) {
    switch (season) {
      case 'kuz': return '🍂 Kuzgi';
      case 'yoz': return '☀️ Yozgi';
      case 'qish': return '❄️ Qishgi';
      default: return '⚡ Barcha fasl';
    }
  }

  function getBadgeClass(badge) {
    if (!badge) return '';
    const map = {
      'Bestseller': 'badge-bestseller',
      'Yangi': 'badge-yangi',
      'Chegirma': 'badge-chegirma',
      'Premium': 'badge-premium',
      'Trend': 'badge-trend',
      'Top reyting': 'badge-top-reyting'
    };
    return map[badge] || 'badge-bestseller';
  }

  // Generate SVG fallback image with Nike Swoosh & item name
  function getFallbackSvg(name, category) {
    const isShoe = category === 'shoes';
    const icon = isShoe ? '👟' : '👕';
    return `data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="600" height="600" viewBox="0 0 600 600">
      <rect width="100%" height="100%" fill="%23f3f4f6"/>
      <path d="M120 280 C 200 340, 320 360, 480 200 C 420 280, 240 380, 140 330 Z" fill="%23111111"/>
      <text x="300" y="440" font-family="sans-serif" font-size="28" font-weight="900" fill="%23111111" text-anchor="middle">${encodeURIComponent(name.slice(0, 26))}</text>
      <text x="300" y="480" font-family="sans-serif" font-size="18" font-weight="900" fill="%2310b981" text-anchor="middle">100% FAQAT ERKAKLAR UCHUN</text>
      <text x="300" y="220" font-size="64" text-anchor="middle">${icon}</text>
    </svg>`;
  }

  // ================= FILTERING & SORTING =================
  function applyFilters() {
    let result = [...state.products];

    // Main Category filter
    if (state.activeCategory === 'shoes') {
      result = result.filter(p => p.category === 'shoes');
    } else if (state.activeCategory === 'clothing') {
      result = result.filter(p => p.category === 'clothing');
    } else if (state.activeCategory === 'autumn') {
      result = result.filter(p => p.season === 'kuz');
    } else if (state.activeCategory === 'summer') {
      result = result.filter(p => p.season === 'yoz');
    } else if (state.activeCategory === 'winter') {
      result = result.filter(p => p.season === 'qish');
    } else if (state.activeCategory === 'bestseller') {
      result = result.filter(p => p.badge === 'Bestseller' || p.badge === 'Top reyting');
    }

    // SubCategory filter
    if (state.activeSubCategory !== 'all') {
      result = result.filter(p => p.subCategory === state.activeSubCategory);
    }

    // Size filter
    if (state.activeSize !== 'all') {
      result = result.filter(p => p.sizes && p.sizes.includes(state.activeSize));
    }

    // Search query filter
    if (state.searchQuery.trim()) {
      const q = state.searchQuery.toLowerCase().trim();
      result = result.filter(p => 
        p.name.toLowerCase().includes(q) ||
        p.description.toLowerCase().includes(q) ||
        (p.colorNames && p.colorNames.toLowerCase().includes(q)) ||
        (p.features && p.features.some(f => f.toLowerCase().includes(q)))
      );
    }

    // Sorting
    switch (state.sortBy) {
      case 'price-asc':
        result.sort((a, b) => a.priceUSD - b.priceUSD);
        break;
      case 'price-desc':
        result.sort((a, b) => b.priceUSD - a.priceUSD);
        break;
      case 'rating':
        result.sort((a, b) => b.rating - a.rating || b.reviewsCount - a.reviewsCount);
        break;
      case 'name':
        result.sort((a, b) => a.name.localeCompare(b.name));
        break;
      case 'featured':
      default:
        // Keep order or prioritize bestsellers
        result.sort((a, b) => (b.badge === 'Bestseller' ? 1 : 0) - (a.badge === 'Bestseller' ? 1 : 0));
        break;
    }

    state.filteredProducts = result;
    state.displayedCount = state.itemsPerPage;
    renderCatalog();
    updateResultsInfo();
  }

  // ================= CATALOG RENDERING =================
  function renderCatalog() {
    if (!els.productsGrid) return;

    const visibleItems = state.filteredProducts.slice(0, state.displayedCount);

    if (visibleItems.length === 0) {
      els.productsGrid.innerHTML = `
        <div class="empty-state">
          <div style="font-size: 3rem; margin-bottom: 12px;">🔍</div>
          <h3 class="empty-state-title">Hech qanday mahsulot topilmadi</h3>
          <p class="empty-state-desc">Qidiruv so'zini yoki tanlangan filtrlarni o'zgartirib ko'ring.</p>
          <button class="btn-primary" id="resetEmptyBtn" style="width: auto; padding: 10px 24px;">Filtrlarni tozalash</button>
        </div>
      `;
      const resetBtn = document.getElementById('resetEmptyBtn');
      if (resetBtn) resetBtn.addEventListener('click', resetAllFilters);
      els.loadMoreContainer.style.display = 'none';
      return;
    }

    els.productsGrid.innerHTML = visibleItems.map(p => {
      const isWish = state.wishlist.some(w => w.id === p.id);
      const fallbackUrl = getFallbackSvg(p.name, p.category);

      return `
        <article class="product-card" data-id="${p.id}">
          <div class="card-image-box" onclick="NikeApp.openDetail('${p.id}')">
            <img class="card-img" 
                 src="${p.image}" 
                 alt="${p.name}" 
                 loading="lazy" 
                 onerror="this.onerror=null; this.src='${fallbackUrl}'" />
            
            <div class="card-badges">
              ${p.badge ? `<span class="card-badge ${getBadgeClass(p.badge)}">${p.badge}</span>` : ''}
              <span class="season-tag">${getSeasonLabel(p.season)}</span>
            </div>

            <button class="card-wish-btn ${isWish ? 'active' : ''}" 
                    title="Sevimlilarga qo'shish" 
                    onclick="event.stopPropagation(); NikeApp.toggleWishlist('${p.id}')">
              <svg viewBox="0 0 24 24"><path d="M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z"/></svg>
            </button>
          </div>

          <div class="card-content">
            <div class="card-meta-row">
              <span class="gender-indicator">Erkaklar</span>
              <div class="card-rating">
                <span>★</span>
                <span>${p.rating}</span>
                <span style="color:#9ca3af; font-size:0.75rem;">(${p.reviewsCount})</span>
              </div>
            </div>

            <h3 class="card-title" onclick="NikeApp.openDetail('${p.id}')" title="${p.name}">${p.name}</h3>
            
            <div class="card-color-names">Rang: ${p.colorNames || 'Standart'}</div>

            <div class="card-price-row">
              <span class="current-price">${formatMoney(p.priceUSD)}</span>
              ${p.originalPriceUSD ? `<span class="original-price">${formatMoney(p.originalPriceUSD)}</span>` : ''}
            </div>

            <div class="card-actions-row">
              <button class="card-add-btn" onclick="NikeApp.quickAddToCart('${p.id}')">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M7 18c-1.1 0-1.99.9-1.99 2S5.9 22 7 22s2-.9 2-2-.9-2-2-2zM1 2v2h2l3.6 7.59-1.35 2.45c-.16.28-.25.61-.25.96 0 1.1.9 2 2 2h12v-2H7.42c-.14 0-.25-.11-.25-.25l.03-.12.9-1.63h7.45c.75 0 1.41-.41 1.75-1.03l3.58-6.49c.08-.14.12-.31.12-.48 0-.55-.45-1-1-1H5.21l-.94-2H1zm16 16c-1.1 0-1.99.9-1.99 2s.89 2 1.99 2 2-.9 2-2-.9-2-2-2z"/></svg>
                <span>Savatga</span>
              </button>
              <button class="card-quick-view-btn" title="Batafsil ko'rish" onclick="NikeApp.openDetail('${p.id}')">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
              </button>
            </div>
          </div>
        </article>
      `;
    }).join('');

    // Load More visibility
    if (state.displayedCount >= state.filteredProducts.length) {
      els.loadMoreContainer.style.display = 'none';
    } else {
      els.loadMoreContainer.style.display = 'flex';
      els.loadMoreInfo.textContent = `${state.filteredProducts.length} ta mahsulotdan ${visibleItems.length} tasi ko'rsatilmoqda`;
    }
  }

  function updateResultsInfo() {
    if (!els.resultsCount) return;
    const count = state.filteredProducts.length;
    els.resultsCount.innerHTML = `Natija: <strong>${count}</strong> ta erkaklar mahsuloti <span>(Jami: ${state.products.length})</span>`;
  }

  // ================= DETAIL MODAL =================
  function openDetail(productId) {
    const product = state.products.find(p => p.id === productId);
    if (!product) return;

    state.selectedProduct = product;
    state.selectedSize = product.sizes ? product.sizes[0] : 'Standart';

    const fallbackUrl = getFallbackSvg(product.name, product.category);
    const isWish = state.wishlist.some(w => w.id === product.id);

    els.detailModalContent.innerHTML = `
      <div class="modal-detail-grid">
        <div class="modal-image-box">
          <img src="${product.image}" alt="${product.name}" onerror="this.onerror=null; this.src='${fallbackUrl}'" />
        </div>

        <div class="modal-info-col">
          <div class="modal-category-badge">${product.gender} • ${getSeasonLabel(product.season)} • ${product.category === 'shoes' ? 'Oyoq kiyim' : 'Kiyim'}</div>
          <h2 class="modal-title">${product.name}</h2>
          
          <div class="card-rating" style="margin-bottom: 12px;">
            <span>★</span>
            <span style="font-weight: 800;">${product.rating}</span>
            <span style="color:#71717a;">(${product.reviewsCount} ta erkaklar fikri)</span>
            ${product.badge ? `<span class="card-badge ${getBadgeClass(product.badge)}" style="margin-left:8px;">${product.badge}</span>` : ''}
          </div>

          <div class="modal-price-box">
            <span class="current-price">${formatMoney(product.priceUSD)}</span>
            ${product.originalPriceUSD ? `<span class="original-price">${formatMoney(product.originalPriceUSD)}</span>` : ''}
          </div>

          <p class="modal-desc">${product.description}</p>

          ${product.features && product.features.length ? `
            <ul class="modal-feature-list">
              ${product.features.map(f => `<li>${f}</li>`).join('')}
            </ul>
          ` : ''}

          <div class="size-selector-box">
            <div class="size-header-row">
              <span>O'lchamni tanlang:</span>
              <span style="color:var(--nike-gray-text); font-weight:400;">(Erkaklar andozasi)</span>
            </div>
            <div class="size-pills-row">
              ${product.sizes.map((s, idx) => `
                <button class="size-pill ${idx === 0 ? 'selected' : ''}" 
                        onclick="NikeApp.selectModalSize(this, '${s}')">
                  ${s}
                </button>
              `).join('')}
            </div>
          </div>

          <div class="modal-cta-row">
            <button class="btn-primary" onclick="NikeApp.addCurrentToCart()">
              Savatchaga qo'shish
            </button>
            <button class="btn-secondary" onclick="NikeApp.buyNowCurrent()">
              Hoziroq sotib olish
            </button>
          </div>
        </div>
      </div>
    `;

    els.detailModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeDetailModal() {
    els.detailModal.classList.remove('active');
    document.body.style.overflow = '';
  }

  function selectModalSize(btnEl, size) {
    state.selectedSize = size;
    const parent = btnEl.closest('.size-pills-row');
    if (parent) {
      parent.querySelectorAll('.size-pill').forEach(b => b.classList.remove('selected'));
      btnEl.classList.add('selected');
    }
  }

  function addCurrentToCart() {
    if (!state.selectedProduct) return;
    addToCart(state.selectedProduct.id, state.selectedSize);
    closeDetailModal();
  }

  function buyNowCurrent() {
    if (!state.selectedProduct) return;
    addToCart(state.selectedProduct.id, state.selectedSize);
    closeDetailModal();
    openCheckout();
  }

  // ================= CART SYSTEM =================
  function quickAddToCart(productId) {
    const product = state.products.find(p => p.id === productId);
    if (!product) return;
    const defaultSize = product.sizes ? product.sizes[0] : 'Standart';
    addToCart(productId, defaultSize);
  }

  function addToCart(productId, size) {
    const product = state.products.find(p => p.id === productId);
    if (!product) return;

    const existingIndex = state.cart.findIndex(item => item.id === productId && item.size === size);
    if (existingIndex > -1) {
      state.cart[existingIndex].qty += 1;
    } else {
      state.cart.push({
        id: product.id,
        name: product.name,
        category: product.category,
        priceUSD: product.priceUSD,
        image: product.image,
        colorNames: product.colorNames,
        size: size || (product.sizes ? product.sizes[0] : 'Standart'),
        qty: 1
      });
    }

    saveCart();
    renderCart();
    updateBadges();
    showToast(`"${product.name}" (${size}) savatchaga qo'shildi!`);
  }

  function updateCartQty(index, delta) {
    if (!state.cart[index]) return;
    state.cart[index].qty += delta;
    if (state.cart[index].qty <= 0) {
      state.cart.splice(index, 1);
      showToast("Mahsulot savatchadan o'chirildi");
    }
    saveCart();
    renderCart();
    updateBadges();
  }

  function removeFromCart(index) {
    state.cart.splice(index, 1);
    saveCart();
    renderCart();
    updateBadges();
    showToast("Mahsulot savatchadan olib tashlandi");
  }

  function saveCart() {
    localStorage.setItem('nike_cart', JSON.stringify(state.cart));
  }

  function renderCart() {
    if (!els.cartItemsList) return;

    if (state.cart.length === 0) {
      els.cartItemsList.innerHTML = `
        <div style="text-align: center; padding: 40px 10px; color: var(--nike-gray-text);">
          <div style="font-size: 2.5rem; margin-bottom: 8px;">🛒</div>
          <p style="font-weight: 700; color: var(--nike-black);">Savatchangiz bo'sh</p>
          <p style="font-size: 0.85rem; margin-top: 4px;">Nike erkaklar kolleksiyasidan xarid qilishni boshlang!</p>
        </div>
      `;
      els.cartSubtotal.textContent = formatMoney(0);
      els.cartDiscountRow.style.display = 'none';
      els.cartTotal.textContent = formatMoney(0);
      return;
    }

    els.cartItemsList.innerHTML = state.cart.map((item, idx) => {
      const fallbackUrl = getFallbackSvg(item.name, item.category);
      return `
        <div class="drawer-item">
          <img class="drawer-item-img" 
               src="${item.image}" 
               alt="${item.name}" 
               onerror="this.onerror=null; this.src='${fallbackUrl}'" />
          
          <div class="drawer-item-info">
            <h4 class="drawer-item-title">${item.name}</h4>
            <div class="drawer-item-sub">O'lcham: <strong>${item.size}</strong> • Erkaklar</div>
            <div class="drawer-item-price">${formatMoney(item.priceUSD * item.qty)}</div>
            
            <div class="drawer-qty-row">
              <div class="qty-control">
                <button class="qty-btn" onclick="NikeApp.updateCartQty(${idx}, -1)">−</button>
                <span class="qty-val">${item.qty}</span>
                <button class="qty-btn" onclick="NikeApp.updateCartQty(${idx}, 1)">+</button>
              </div>
              <button class="item-remove-btn" onclick="NikeApp.removeFromCart(${idx})">O'chirish</button>
            </div>
          </div>
        </div>
      `;
    }).join('');

    // Totals calculation
    const rawSubtotalUSD = state.cart.reduce((sum, i) => sum + (i.priceUSD * i.qty), 0);
    let discountUSD = 0;

    if (state.appliedPromo) {
      if (state.appliedPromo.type === 'percent') {
        discountUSD = (rawSubtotalUSD * state.appliedPromo.value) / 100;
      } else if (state.appliedPromo.type === 'fixed') {
        discountUSD = Math.min(rawSubtotalUSD, state.appliedPromo.value);
      }
    }

    const finalTotalUSD = Math.max(0, rawSubtotalUSD - discountUSD);

    els.cartSubtotal.textContent = formatMoney(rawSubtotalUSD);

    if (discountUSD > 0) {
      els.cartDiscountRow.style.display = 'flex';
      els.cartDiscountVal.textContent = `-${formatMoney(discountUSD)}`;
    } else {
      els.cartDiscountRow.style.display = 'none';
    }

    els.cartTotal.textContent = formatMoney(finalTotalUSD);
  }

  function applyPromoCode() {
    const code = (els.promoInput.value || '').trim().toUpperCase();
    if (!code) {
      showToast("Promokodni kiriting", "⚠️");
      return;
    }
    if (PROMO_CODES[code]) {
      state.appliedPromo = { code, ...PROMO_CODES[code] };
      showToast(`Promokod qo'llandi: ${PROMO_CODES[code].label}!`, "🎉");
      renderCart();
    } else {
      showToast("Noto'g'ri promokod. 'NIKE2026' kodini sinab ko'ring", "❌");
    }
  }

  // ================= WISHLIST SYSTEM =================
  function toggleWishlist(productId) {
    const product = state.products.find(p => p.id === productId);
    if (!product) return;

    const existingIdx = state.wishlist.findIndex(w => w.id === productId);
    if (existingIdx > -1) {
      state.wishlist.splice(existingIdx, 1);
      showToast(`"${product.name}" sevimlilardan o'chirildi`);
    } else {
      state.wishlist.push({
        id: product.id,
        name: product.name,
        priceUSD: product.priceUSD,
        image: product.image,
        category: product.category,
        season: product.season
      });
      showToast(`"${product.name}" sevimlilarga qo'shildi!`, "❤️");
    }

    localStorage.setItem('nike_wishlist', JSON.stringify(state.wishlist));
    updateBadges();
    renderCatalog();
    renderWishlist();
  }

  function renderWishlist() {
    if (!els.wishlistItemsList) return;

    if (state.wishlist.length === 0) {
      els.wishlistItemsList.innerHTML = `
        <div style="text-align: center; padding: 40px 10px; color: var(--nike-gray-text);">
          <div style="font-size: 2.5rem; margin-bottom: 8px;">🤍</div>
          <p style="font-weight: 700; color: var(--nike-black);">Sevimlilar ro'yxati bo'sh</p>
          <p style="font-size: 0.85rem; margin-top: 4px;">Yoqqan krossovka va kiyimlarni yurakcha orqali saqlang.</p>
        </div>
      `;
      return;
    }

    els.wishlistItemsList.innerHTML = state.wishlist.map((item, idx) => {
      const fallbackUrl = getFallbackSvg(item.name, item.category);
      return `
        <div class="drawer-item">
          <img class="drawer-item-img" 
               src="${item.image}" 
               alt="${item.name}" 
               onerror="this.onerror=null; this.src='${fallbackUrl}'" />
          
          <div class="drawer-item-info">
            <h4 class="drawer-item-title">${item.name}</h4>
            <div class="drawer-item-sub">Erkaklar • ${getSeasonLabel(item.season)}</div>
            <div class="drawer-item-price">${formatMoney(item.priceUSD)}</div>
            
            <div style="display:flex; gap:8px; margin-top:8px;">
              <button class="btn-primary" 
                      style="padding:6px 14px; font-size:0.8rem;"
                      onclick="NikeApp.quickAddToCart('${item.id}'); NikeApp.toggleWishlist('${item.id}');">
                Savatga olish
              </button>
              <button class="item-remove-btn" onclick="NikeApp.toggleWishlist('${item.id}')">
                O'chirish
              </button>
            </div>
          </div>
        </div>
      `;
    }).join('');
  }

  function updateBadges() {
    const totalCartCount = state.cart.reduce((s, i) => s + i.qty, 0);
    els.cartBadge.textContent = totalCartCount;
    els.wishlistBadge.textContent = state.wishlist.length;
  }

  // ================= CHECKOUT ORDER SYSTEM =================
  function openCheckout() {
    if (state.cart.length === 0) {
      showToast("Savatchangiz bo'sh. Avval mahsulot tanlang!", "⚠️");
      return;
    }
    closeCart();

    const rawSubtotalUSD = state.cart.reduce((sum, i) => sum + (i.priceUSD * i.qty), 0);
    let discountUSD = 0;
    if (state.appliedPromo) {
      discountUSD = state.appliedPromo.type === 'percent'
        ? (rawSubtotalUSD * state.appliedPromo.value) / 100
        : Math.min(rawSubtotalUSD, state.appliedPromo.value);
    }
    const finalTotalUSD = Math.max(0, rawSubtotalUSD - discountUSD);

    els.checkoutTotalDisplay.textContent = formatMoney(finalTotalUSD);
    els.checkoutModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function handleOrderSubmit(e) {
    e.preventDefault();
    const name = document.getElementById('custName').value.trim();
    const phone = document.getElementById('custPhone').value.trim();
    const address = document.getElementById('custAddress').value.trim();

    if (!name || !phone || !address) {
      showToast("Iltimos barcha maydonlarni to'ldiring", "⚠️");
      return;
    }

    const orderId = 'NK-' + Math.floor(100000 + Math.random() * 900000);

    // Clear Cart
    state.cart = [];
    state.appliedPromo = null;
    saveCart();
    renderCart();
    updateBadges();

    // Close checkout and open success modal
    els.checkoutModal.classList.remove('active');
    els.successOrderNum.textContent = orderId;
    els.successModal.classList.add('active');
  }

  // ================= UI EVENT BINDINGS =================
  function initEvents() {
    // Theme Switcher Buttons
    if (els.themeToggleBtn) {
      els.themeToggleBtn.addEventListener('click', toggleTheme);
    }
    if (els.topBarThemeBtn) {
      els.topBarThemeBtn.addEventListener('click', toggleTheme);
    }

    // Currency Switcher
    els.currencyBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        els.currencyBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.currency = btn.dataset.currency;
        localStorage.setItem('nike_currency', state.currency);
        renderCatalog();
        renderCart();
        renderWishlist();
        if (state.selectedProduct) openDetail(state.selectedProduct.id);
      });
    });

    // Navigation Links
    els.navLinks.forEach(link => {
      link.addEventListener('click', () => {
        els.navLinks.forEach(l => l.classList.remove('active'));
        link.classList.add('active');
        const cat = link.dataset.category;
        state.activeCategory = cat;
        // Sync catalog tabs
        els.catalogTabs.forEach(t => {
          t.classList.toggle('active', t.dataset.category === cat);
        });
        applyFilters();
      });
    });

    // Catalog Tabs
    els.catalogTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        els.catalogTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const cat = tab.dataset.category;
        state.activeCategory = cat;
        // Sync nav links
        els.navLinks.forEach(l => {
          l.classList.toggle('active', l.dataset.category === cat);
        });
        applyFilters();
      });
    });

    // Hero Quick Filter Buttons
    els.quickFilters.forEach(btn => {
      btn.addEventListener('click', () => {
        const cat = btn.dataset.category;
        state.activeCategory = cat;
        els.catalogTabs.forEach(t => t.classList.toggle('active', t.dataset.category === cat));
        els.navLinks.forEach(l => l.classList.toggle('active', l.dataset.category === cat));
        applyFilters();
        const catSection = document.getElementById('catalogSection');
        if (catSection) catSection.scrollIntoView({ behavior: 'smooth' });
      });
    });

    // Video Mute / Unmute Toggle Button
    if (els.videoMuteToggleBtn && els.heroVideoPlayer) {
      els.videoMuteToggleBtn.addEventListener('click', () => {
        const isMuted = els.heroVideoPlayer.muted;
        els.heroVideoPlayer.muted = !isMuted;
        if (els.heroVideoPlayer.muted) {
          if (els.soundIcon) els.soundIcon.textContent = '🔇';
          if (els.soundText) els.soundText.textContent = 'Ovozsiz';
          showToast('🔇 Video ovozi o\'chirildi');
        } else {
          if (els.soundIcon) els.soundIcon.textContent = '🔊';
          if (els.soundText) els.soundText.textContent = 'Ovozli';
          showToast('🔊 Video ovozi yoqildi!');
          els.heroVideoPlayer.play().catch(() => {});
        }
      });
    }

    // Hero Video Playlist Switcher (Supports Local HD MP4 & YouTube Embeds)
    if (els.videoPlaylistItems) {
      els.videoPlaylistItems.forEach(item => {
        item.addEventListener('click', () => {
          const type = item.dataset.type || 'local';
          const title = item.dataset.title;

          // Update active button state
          els.videoPlaylistItems.forEach(b => b.classList.remove('active'));
          item.classList.add('active');

          // Update active title in topbar
          if (els.videoActiveTitle && title) {
            els.videoActiveTitle.textContent = title;
          }

          if (type === 'local' && els.heroVideoPlayer) {
            const src = item.dataset.src;
            if (!src) return;

            // Hide iframe, show native video player
            if (els.heroVideoIframe) {
              els.heroVideoIframe.style.display = 'none';
              els.heroVideoIframe.src = '';
            }
            els.heroVideoPlayer.style.display = 'block';

            // Set new source and play
            els.heroVideoPlayer.src = src;
            els.heroVideoPlayer.load();
            els.heroVideoPlayer.play().catch(e => {
              console.log('Autoplay handled:', e);
            });

            // Show sound toggle button for local video
            if (els.videoMuteToggleBtn) els.videoMuteToggleBtn.style.display = 'inline-flex';

            showToast(`🎬 Video qo'yildi: ${title}`);
          } else if (type === 'youtube' && els.heroVideoIframe) {
            const ytId = item.dataset.yt;
            if (!ytId) return;

            // Pause native video player, hide it
            if (els.heroVideoPlayer) {
              els.heroVideoPlayer.pause();
              els.heroVideoPlayer.style.display = 'none';
            }

            // Show iframe, load YouTube embed
            els.heroVideoIframe.style.display = 'block';
            els.heroVideoIframe.src = `https://www.youtube.com/embed/${ytId}?autoplay=1&mute=0&enablejsapi=1&rel=0`;

            // Hide sound toggle button because YouTube has built-in volume
            if (els.videoMuteToggleBtn) els.videoMuteToggleBtn.style.display = 'none';

            showToast(`🎬 YouTube video qo'yildi: ${title}`);
          }
        });
      });
    }

    // Search Input
    let searchTimeout;
    els.searchInput.addEventListener('input', (e) => {
      clearTimeout(searchTimeout);
      const val = e.target.value;
      els.searchClearBtn.style.display = val ? 'block' : 'none';
      searchTimeout = setTimeout(() => {
        state.searchQuery = val;
        applyFilters();
      }, 250);
    });

    els.searchClearBtn.addEventListener('click', () => {
      els.searchInput.value = '';
      state.searchQuery = '';
      els.searchClearBtn.style.display = 'none';
      applyFilters();
    });

    // Sort Select
    els.sortSelect.addEventListener('change', (e) => {
      state.sortBy = e.target.value;
      applyFilters();
    });

    // Filter Toggle Button (opens subfilters panel)
    els.filterToggleBtn.addEventListener('click', () => {
      const isHidden = els.filterDrawerPanel.style.display === 'none';
      els.filterDrawerPanel.style.display = isHidden ? 'flex' : 'none';
    });

    // SubCategory Pills
    if (els.subCategoryPills) {
      els.subCategoryPills.querySelectorAll('.pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          els.subCategoryPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          state.activeSubCategory = btn.dataset.sub;
          applyFilters();
        });
      });
    }

    // Size Pills
    if (els.sizePills) {
      els.sizePills.querySelectorAll('.pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          els.sizePills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          state.activeSize = btn.dataset.size;
          applyFilters();
        });
      });
    }

    // Reset Filters
    if (els.resetFiltersBtn) {
      els.resetFiltersBtn.addEventListener('click', resetAllFilters);
    }

    // Load More Button
    els.loadMoreBtn.addEventListener('click', () => {
      state.displayedCount += state.itemsPerPage;
      renderCatalog();
    });

    // Cart Open / Close
    els.cartBtn.addEventListener('click', openCart);
    els.cartCloseBtn.addEventListener('click', closeCart);
    els.cartBackdrop.addEventListener('click', () => {
      closeCart();
      closeWishlist();
    });

    // Promo Code
    els.promoApplyBtn.addEventListener('click', applyPromoCode);
    els.checkoutBtn.addEventListener('click', openCheckout);

    // Wishlist Open / Close
    els.wishlistBtn.addEventListener('click', openWishlist);
    els.wishlistCloseBtn.addEventListener('click', closeWishlist);

    // Detail Modal Close
    els.detailModalClose.addEventListener('click', closeDetailModal);
    els.detailModal.addEventListener('click', (e) => {
      if (e.target === els.detailModal) closeDetailModal();
    });

    // Checkout Modal
    els.checkoutModalClose.addEventListener('click', () => {
      els.checkoutModal.classList.remove('active');
      document.body.style.overflow = '';
    });
    els.checkoutForm.addEventListener('submit', handleOrderSubmit);

    // Payment Cards Selection
    document.querySelectorAll('.payment-radio-card').forEach(card => {
      card.addEventListener('click', () => {
        document.querySelectorAll('.payment-radio-card').forEach(c => c.classList.remove('active'));
        card.classList.add('active');
        const radio = card.querySelector('input[type="radio"]');
        if (radio) radio.checked = true;
      });
    });

    // Success Modal Close
    els.successModalClose.addEventListener('click', () => {
      els.successModal.classList.remove('active');
      document.body.style.overflow = '';
    });

    // QR Code Modal Handlers
    function openQrModal() {
      if (els.qrModal) {
        els.qrModal.classList.add('active');
        document.body.style.overflow = 'hidden';
      }
    }

    function closeQrModal() {
      if (els.qrModal) {
        els.qrModal.classList.remove('active');
        document.body.style.overflow = '';
      }
    }

    if (els.heroQrBadge) {
      els.heroQrBadge.addEventListener('click', openQrModal);
    }

    if (els.floatingQrTrigger) {
      els.floatingQrTrigger.addEventListener('click', openQrModal);
    }

    if (els.qrModalClose) {
      els.qrModalClose.addEventListener('click', closeQrModal);
    }

    if (els.qrDismissBtn) {
      els.qrDismissBtn.addEventListener('click', closeQrModal);
    }

    if (els.qrModal) {
      els.qrModal.addEventListener('click', (e) => {
        if (e.target === els.qrModal) closeQrModal();
      });
    }

    if (els.qrCopyBtn) {
      els.qrCopyBtn.addEventListener('click', () => {
        const urlToCopy = window.location.href.startsWith('http') 
          ? window.location.href 
          : 'https://nike-men-uzbekistan.shop';
        navigator.clipboard.writeText(urlToCopy).then(() => {
          showToast('📋 Sayt havolasi nusxalandi!', '✓');
          const originalText = els.qrCopyBtn.innerHTML;
          els.qrCopyBtn.innerHTML = '✓ Nusxalandi!';
          setTimeout(() => {
            els.qrCopyBtn.innerHTML = originalText;
          }, 2000);
        }).catch(() => {
          showToast('Havola: ' + urlToCopy);
        });
      });
    }

    // Keyboard ESC listener
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeDetailModal();
        closeCart();
        closeWishlist();
        closeQrModal();
        els.checkoutModal.classList.remove('active');
        els.successModal.classList.remove('active');
        document.body.style.overflow = '';
      }
    });
  }

  function resetAllFilters() {
    state.activeCategory = 'all';
    state.activeSubCategory = 'all';
    state.activeSize = 'all';
    state.searchQuery = '';
    state.sortBy = 'featured';
    if (els.searchInput) els.searchInput.value = '';
    if (els.searchClearBtn) els.searchClearBtn.style.display = 'none';
    if (els.sortSelect) els.sortSelect.value = 'featured';

    els.catalogTabs.forEach(t => t.classList.toggle('active', t.dataset.category === 'all'));
    els.navLinks.forEach(l => l.classList.toggle('active', l.dataset.category === 'all'));
    
    if (els.subCategoryPills) {
      els.subCategoryPills.querySelectorAll('.pill-btn').forEach(b => b.classList.toggle('active', b.dataset.sub === 'all'));
    }
    if (els.sizePills) {
      els.sizePills.querySelectorAll('.pill-btn').forEach(b => b.classList.toggle('active', b.dataset.size === 'all'));
    }

    applyFilters();
    showToast("Barcha filtrlar tozalandi");
  }

  function openCart() {
    closeWishlist();
    els.cartDrawer.classList.add('active');
    els.cartBackdrop.classList.add('active');
    document.body.style.overflow = 'hidden';
    renderCart();
  }

  function closeCart() {
    els.cartDrawer.classList.remove('active');
    els.cartBackdrop.classList.remove('active');
    document.body.style.overflow = '';
  }

  function openWishlist() {
    closeCart();
    els.wishlistDrawer.classList.add('active');
    els.cartBackdrop.classList.add('active');
    document.body.style.overflow = 'hidden';
    renderWishlist();
  }

  function closeWishlist() {
    els.wishlistDrawer.classList.remove('active');
    els.cartBackdrop.classList.remove('active');
    document.body.style.overflow = '';
  }

  // ================= INITIALIZATION =================
  function init() {
    // Apply initial theme
    applyTheme(state.theme);

    // Set initial currency button state
    els.currencyBtns.forEach(btn => {
      btn.classList.toggle('active', btn.dataset.currency === state.currency);
    });

    initEvents();
    applyFilters();
    renderCart();
    renderWishlist();
    updateBadges();

    console.log(`Nike Men's Store 2026 loaded: ${state.products.length} exclusive products.`);
  }

  // Global exposure for inline HTML events
  window.NikeApp = {
    openDetail,
    selectModalSize,
    addCurrentToCart,
    buyNowCurrent,
    quickAddToCart,
    updateCartQty,
    removeFromCart,
    toggleWishlist,
    openCart,
    closeCart,
    openWishlist,
    closeWishlist,
    toggleTheme
  };

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
