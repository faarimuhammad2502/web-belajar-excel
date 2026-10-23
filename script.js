/* ==========================================================
   rumus • — Master JavaScript (Vanilla JS)
   - Tab Modul Switcher (Finance vs HR)
   - Animated Pulse Favicon (Canvas)
   - Live SUMIF Calculation on Hero Sheet
   - Click-to-Copy Excel Formulas
   - Live PMT Loan Simulator Calculator
   - Live TRIM & PROPER Data Cleansing Demo
   ========================================================== */

document.addEventListener("DOMContentLoaded", () => {
  initAnimatedFavicon();
  initNavbar();
  initModulTabs();
  initCopyFormula();
  initPmtCalculator();
  initCleanToggle();
  initHeroSumif();
});

/* ---------- Helper Functions ---------- */
const $ = (selector, scope = document) => scope.querySelector(selector);
const $$ = (selector, scope = document) => [...scope.querySelectorAll(selector)];

const formatRupiah = (val) =>
  "Rp" + Math.round(val).toLocaleString("id-ID");

function showToast(message) {
  const toast = $("#toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add("is-visible");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => {
    toast.classList.remove("is-visible");
  }, 2400);
}

/* ==========================================================
   1. Animated Favicon (Canvas Pulse Animation)
   ========================================================== */
function initAnimatedFavicon() {
  const link = $("#favicon");
  if (!link) return;

  const size = 64;
  const canvas = document.createElement("canvas");
  canvas.width = canvas.height = size;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  let frame = 0;
  const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  function roundRect(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function draw() {
    const pulse = (Math.sin(frame / 7) + 1) / 2;
    const scale = 0.85 + pulse * 0.15;
    const s = size * scale;
    const offset = (size - s) / 2;

    ctx.clearRect(0, 0, size, size);

    // Green rounded container
    ctx.fillStyle = "#19734D";
    roundRect(offset, offset, s, s, s * 0.22);
    ctx.fill();

    // Excel grid inner lines
    ctx.strokeStyle = "rgba(255, 255, 255, 0.35)";
    ctx.lineWidth = 2;
    for (let i = 1; i < 3; i++) {
      const p = offset + (s / 3) * i;
      ctx.beginPath(); ctx.moveTo(offset + 4, p); ctx.lineTo(offset + s - 4, p); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(p, offset + 4); ctx.lineTo(p, offset + s - 4); ctx.stroke();
    }

    // Bold "X"
    ctx.fillStyle = "#FFFFFF";
    ctx.font = `800 ${Math.round(s * 0.58)}px 'Plus Jakarta Sans', system-ui, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("X", size / 2, size / 2 + s * 0.04);

    link.href = canvas.toDataURL("image/png");
    link.type = "image/png";
    frame++;
  }

  draw();
  if (!reduceMotion) {
    setInterval(draw, 120);
  }
}

/* ==========================================================
   2. Navbar Responsiveness & Smooth Scroll
   ========================================================== */
function initNavbar() {
  const navbar = $("#navbar");
  const toggle = $("#navToggle");
  const menu = $("#navMenu");

  const closeMenu = () => {
    menu.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
  };

  toggle.addEventListener("click", () => {
    const isOpen = menu.classList.toggle("is-open");
    toggle.setAttribute("aria-expanded", String(isOpen));
  });

  // Close menu on link click
  $$("a", menu).forEach((link) => {
    link.addEventListener("click", closeMenu);
  });

  // Links with data-goto-tab switch tab automatically
  $$("[data-goto-tab]").forEach((link) => {
    link.addEventListener("click", () => {
      const tabTarget = link.dataset.gotoTab;
      switchTab(tabTarget);
    });
  });

  // Navbar shadow on scroll
  window.addEventListener("scroll", () => {
    navbar.classList.toggle("is-scrolled", window.scrollY > 15);
  }, { passive: true });
}

/* ==========================================================
   3. Modul Tab Switcher (Finance vs HR)
   ========================================================== */
function initModulTabs() {
  const tabBtnFinance = $("#tabBtnFinance");
  const tabBtnHr = $("#tabBtnHr");

  tabBtnFinance.addEventListener("click", () => switchTab("finance"));
  tabBtnHr.addEventListener("click", () => switchTab("hr"));

  // Check URL hash (e.g. index.html#hr)
  if (window.location.hash === "#hr") {
    switchTab("hr");
  }
}

function switchTab(tab) {
  const tabBtnFinance = $("#tabBtnFinance");
  const tabBtnHr = $("#tabBtnHr");
  const contentFinance = $("#tabContentFinance");
  const contentHr = $("#tabContentHr");

  if (tab === "finance") {
    tabBtnFinance.classList.add("is-active");
    tabBtnFinance.setAttribute("aria-selected", "true");
    tabBtnHr.classList.remove("is-active");
    tabBtnHr.setAttribute("aria-selected", "false");

    contentFinance.hidden = false;
    contentFinance.classList.add("is-active");
    contentHr.hidden = true;
    contentHr.classList.remove("is-active");
  } else {
    tabBtnHr.classList.add("is-active");
    tabBtnHr.setAttribute("aria-selected", "true");
    tabBtnFinance.classList.remove("is-active");
    tabBtnFinance.setAttribute("aria-selected", "false");

    contentHr.hidden = false;
    contentHr.classList.add("is-active");
    contentFinance.hidden = true;
    contentFinance.classList.remove("is-active");
  }
}

/* ==========================================================
   4. Click-to-Copy Formula
   ========================================================== */
function initCopyFormula() {
  $$(".formula-box").forEach((box) => {
    box.addEventListener("click", async () => {
      const codeToCopy = box.dataset.copy;
      if (!codeToCopy) return;

      try {
        await navigator.clipboard.writeText(codeToCopy);
      } catch {
        // Fallback older browsers
        const textarea = document.createElement("textarea");
        textarea.value = codeToCopy;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand("copy");
        textarea.remove();
      }

      const copyTag = box.querySelector(".formula-copy-tag");
      if (copyTag) {
        copyTag.textContent = "Tersalin ✓";
        setTimeout(() => {
          copyTag.textContent = "Salin";
        }, 1800);
      }

      showToast("Rumus berhasil disalin ke clipboard! 📋");
    });
  });
}

/* ==========================================================
   5. Interactive PMT Calculator
   Excel Formula: =PMT(rate, nper, pv)
   ========================================================== */
function pmt(annualRate, months, principal) {
  const r = annualRate / 12;
  if (r === 0) return principal / months;
  return (principal * r) / (1 - Math.pow(1 + r, -months));
}

function initPmtCalculator() {
  const principalEl = $("#pmtPrincipal");
  const rateEl = $("#pmtRate");
  const tenorEl = $("#pmtTenor");
  const resultEl = $("#pmtResult");
  const loanRows = $$("#loanTable tbody tr");

  const calc = () => {
    const p = parseFloat(principalEl.value) || 0;
    const r = (parseFloat(rateEl.value) || 0) / 100;
    const t = parseInt(tenorEl.value, 10) || 0;
    if (t > 0 && p > 0) {
      resultEl.textContent = formatRupiah(pmt(r, t, p));
    } else {
      resultEl.textContent = "Rp 0";
    }
  };

  [principalEl, rateEl, tenorEl].forEach((input) => {
    input.addEventListener("input", () => {
      loanRows.forEach((r) => r.classList.remove("is-selected"));
      calc();
    });
  });

  // Click row to auto-fill
  loanRows.forEach((row) => {
    row.addEventListener("click", () => {
      loanRows.forEach((r) => r.classList.toggle("is-selected", r === row));
      principalEl.value = row.dataset.principal;
      rateEl.value = (parseFloat(row.dataset.rate) * 100).toString();
      tenorEl.value = row.dataset.tenor;
      calc();
      showToast(`Memuat data ${row.children[1].textContent} ke kalkulator`);
    });
  });

  // Initial highlight row 1
  if (loanRows[0]) loanRows[0].classList.add("is-selected");
  calc();
}

/* ==========================================================
   6. TRIM & PROPER Data Cleansing Interactive Demo (HR)
   ========================================================== */
function initCleanToggle() {
  const btn = $("#btnToggleClean");
  const cells = $$("#tableEmployeeRaw .raw-cell");
  if (!btn || !cells.length) return;

  let isCleaned = false;

  const excelTrim = (text) => text.trim().replace(/\s+/g, " ");
  const excelProper = (text) =>
    text.toLowerCase().replace(/(^|\s)\S/g, (t) => t.toUpperCase());

  btn.addEventListener("click", () => {
    isCleaned = !isCleaned;
    cells.forEach((cell) => {
      const rawVal = cell.dataset.raw;
      if (isCleaned) {
        cell.textContent = `"${excelProper(excelTrim(rawVal))}"`;
        cell.classList.add("is-cleaned");
      } else {
        cell.textContent = `"${rawVal}"`;
        cell.classList.remove("is-cleaned");
      }
    });

    btn.textContent = isCleaned ? "↩️ Tampilkan Data Mentah" : "✨ Bersihkan dengan TRIM + PROPER";
    showToast(isCleaned ? "Data dibersihkan dengan =TRIM(PROPER(...))! ✨" : "Kembali ke data input mentah");
  });
}

/* ==========================================================
   7. Live Hero SUMIF Highlight & Calculation
   ========================================================== */
function initHeroSumif() {
  const rows = $$(".excel-row[data-city]");
  const sumEl = $("#heroLiveSum");
  if (!rows.length || !sumEl) return;

  const criteria = "Jakarta";
  let total = 0;

  rows.forEach((row) => {
    if (row.dataset.city === criteria) {
      row.classList.add("is-matched");
      total += parseInt(row.dataset.val, 10) || 0;
    }
  });

  sumEl.textContent = formatRupiah(total);
}
