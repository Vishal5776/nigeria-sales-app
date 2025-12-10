// === CONFIG: Google Apps Script Web App URL ===
const API_URL =
  "https://script.google.com/macros/s/AKfycbw6WSGTrnh_6k9pNPNOYzAmx-9vbOvfdq2uqfy22u40EnjlWCmfZb8KnMRuuKx5nlIk/exec";

// === Dashboard chart (Chart.js) ===
let sales7dChart = null;

// ==== Inventory → Sales bridge (in-memory cart) ====
let hasAutoLoadedQueuedBales = false; // only auto-load once

// ======================================
// CHART
// ======================================
function renderSalesChart(salesByDate) {
  const ctx = document.getElementById("sales-7d-chart");
  if (!ctx) return;

  const labels = (salesByDate || []).map((p) => p.date);
  const data = (salesByDate || []).map((p) => p.amount);

  if (!labels.length) {
    if (sales7dChart) {
      sales7dChart.destroy();
      sales7dChart = null;
    }
    return;
  }

  if (sales7dChart) {
    sales7dChart.data.labels = labels;
    sales7dChart.data.datasets[0].data = data;
    sales7dChart.update();
    return;
  }

  sales7dChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Sales (₦)",
          data,
          borderWidth: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          ticks: {
            maxRotation: 0,
            autoSkip: true,
            font: { size: 10 }
          }
        },
        y: {
          beginAtZero: true,
          ticks: {
            font: { size: 10 },
            callback: (v) => v.toLocaleString()
          }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const v = ctx.parsed.y || 0;
              return "₦ " + v.toFixed(2);
            }
          }
        }
      }
    }
  });
}

// ======================================
// TOAST
// ======================================
function showToast(msg, ok = true) {
  const box = document.getElementById("toast");
  if (!box) return;
  const el = document.createElement("div");
  el.textContent = msg;
  el.style.padding = "10px 14px";
  el.style.background = ok ? "#10B981" : "#EF4444";
  el.style.color = "#fff";
  el.style.borderRadius = "8px";
  el.style.marginTop = "6px";
  el.style.fontSize = "12px";
  el.style.maxWidth = "260px";
  el.style.wordWrap = "break-word";
  box.appendChild(el);
  setTimeout(() => el.remove(), 3000);
}

// ======================================
// GENERIC API CALL
// ======================================
async function apiCall(action, payload = {}) {
  const url = new URL(API_URL);
  url.searchParams.set("action", action);
  url.searchParams.set("payload", JSON.stringify(payload));

  const res = await fetch(url.toString(), { method: "GET" });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "API error");
  return json.data;
}

// ======================================
// DASHBOARD
// ======================================
async function loadDashboard(filters = {}) {
  try {
    const d = await apiCall("dashboard", filters);
    const t = d.totals || {};

    document.getElementById("db-total-invoices").textContent =
      t.totalInvoices || 0;
    document.getElementById("db-total-qty").textContent = t.totalQty || 0;
    document.getElementById("db-total-amount").textContent = Number(
      t.totalAmount || 0
    ).toFixed(2);
    document.getElementById("db-today-amount").textContent = Number(
      t.todayAmount || 0
    ).toFixed(2);

    const tcBody = document.getElementById("db-top-customers-body");
    tcBody.innerHTML = "";
    (d.topCustomers || []).forEach((c) => {
      tcBody.innerHTML += `
        <tr>
          <td class="p-2">${c.customer}</td>
          <td class="p-2 text-right">${c.amount.toFixed(2)}</td>
        </tr>`;
    });

    const riBody = document.getElementById("db-recent-invoices-body");
    riBody.innerHTML = "";
    (d.recentInvoices || []).forEach((inv) => {
      riBody.innerHTML += `
        <tr>
          <td class="p-2">${inv.invoiceNo}</td>
          <td class="p-2">${inv.date}</td>
          <td class="p-2">${inv.customer}</td>
          <td class="p-2 text-right">${inv.amount.toFixed(2)}</td>
        </tr>`;
    });

    renderSalesChart(d.salesByDate || []);
  } catch (err) {
    console.error(err);
    showToast("Dashboard load failed: " + err.message, false);
  }
}

function loadDashboardFromUI() {
  const from = document.getElementById("dash-from")?.value || null;
  const to = document.getElementById("dash-to")?.value || null;
  loadDashboard({ dateFrom: from, dateTo: to });
}

// ======================================
// OFFLINE QUEUE (Inward verify)
// ======================================
function loadVerifyQueue() {
  try {
    return JSON.parse(localStorage.getItem("pendingVerifications") || "[]");
  } catch (e) {
    return [];
  }
}
function saveVerifyQueue(q) {
  localStorage.setItem("pendingVerifications", JSON.stringify(q));
}

async function processVerifyQueue() {
  if (!navigator.onLine) return;
  let q = loadVerifyQueue();
  if (!q.length) return;
  showToast("Syncing offline verifications...");
  while (q.length && navigator.onLine) {
    const item = q[0];
    try {
      await apiCall("verifyBaleRow", { baleNo: item.baleNo });
      q.shift();
      saveVerifyQueue(q);
    } catch (err) {
      showToast("Sync error: " + err.message, false);
      break;
    }
  }
  if (!q.length) showToast("Verification sync complete");
}

// ======================================
// TABS + INIT
// ======================================
document.addEventListener("DOMContentLoaded", () => {
  const tabs = ["dashboard", "inward", "inventory", "sales"];

  function activate(tabName) {
    tabs.forEach((t) => {
      document
        .getElementById("tab-" + t)
        .classList.toggle("hidden", t !== tabName);
      document
        .getElementById("btn-" + t)
        .classList.toggle("tab-active", t === tabName);
    });

    if (tabName === "dashboard") loadDashboardFromUI();
    if (tabName === "inward") loadPendingInward();
    if (tabName === "inventory") loadInventorySummary();
    if (tabName === "sales") {
      loadRecentInvoices();
      autoLoadQueuedBales();
    }
  }

  tabs.forEach((t) => {
    document
      .getElementById("btn-" + t)
      .addEventListener("click", () => activate(t));
  });

  const btnApply = document.getElementById("dash-apply");
  const btnClear = document.getElementById("dash-clear");
  if (btnApply) btnApply.addEventListener("click", () => loadDashboardFromUI());
  if (btnClear) {
    btnClear.addEventListener("click", () => {
      const from = document.getElementById("dash-from");
      const to = document.getElementById("dash-to");
      if (from) from.value = "";
      if (to) to.value = "";
      loadDashboard({});
    });
  }

  window.addEventListener("online", processVerifyQueue);

  initSalesTab();

  activate("dashboard");
  loadDashboardFromUI();
  processVerifyQueue();
});

// ======================================
// INWARD (grouped by Bale)
// ======================================
async function loadPendingInward() {
  const list = document.getElementById("verify-list");
  list.innerHTML =
    '<div class="text-sm text-gray-500">Loading...</div>';
  try {
    const rows = await apiCall("getPendingInward", {});
    if (!rows.length) {
      list.innerHTML =
        '<div class="text-sm text-gray-500">No pending bales for Nigeria Godown.</div>';
      return;
    }
    list.innerHTML = "";
    rows.forEach((b) => {
      const div = document.createElement("div");
      div.className =
        "p-3 border rounded bg-gray-50 flex flex-col md:flex-row md:justify-between md:items-center gap-2 text-xs md:text-sm";

      const totalYard = Number(b.totalQtyYard || 0);

      div.innerHTML = `
        <div>
          <div><strong>Bale ${b.baleNo || ""}</strong> • ${b.fabricName || ""} • ${
        b.fabricQuality || ""
      }</div>
          <div class="text-gray-500 mt-1">
            Total Qty: ${totalYard.toLocaleString()} Yard • 
            Designs: ${b.designCount || 0} • 
            Colors: ${b.colorCount || 0}
          </div>
        </div>
        <div class="flex md:block">
          <button class="bg-emerald-500 text-white px-3 py-1 rounded text-xs w-full md:w-auto"
                  onclick="verifyFromApp('${(b.baleNo || "")
                    .replace(/'/g, "\\'")}')">
            Verify
          </button>
        </div>
      `;
      list.appendChild(div);
    });
  } catch (err) {
    console.error(err);
    list.innerHTML =
      '<div class="text-sm text-red-500">Error loading: ' +
      err.message +
      "</div>";
  }
}

async function verifyFromApp(baleNo) {
  if (!navigator.onLine) {
    const q = loadVerifyQueue();
    q.push({ baleNo });
    saveVerifyQueue(q);
    showToast("Saved offline. Will verify when online.");
    return;
  }
  try {
    await apiCall("verifyBaleRow", { baleNo });
    showToast("Verified bale " + baleNo);
    loadPendingInward();
  } catch (err) {
    showToast("Verify failed: " + err.message, false);
  }
}

// ======================================
// INVENTORY
// ======================================
async function loadInventorySummary() {
  const summaryBox = document.getElementById("inventory-summary");
  const tableBody = document.getElementById("inventory-fabric-body");
  const detailBox = document.getElementById("inventory-detail");

  if (!summaryBox || !tableBody) return;

  summaryBox.innerHTML = "Loading inventory summary...";
  tableBody.innerHTML =
    '<tr><td class="p-2 text-sm text-gray-500" colspan="3">Loading...</td></tr>';
  if (detailBox) {
    detailBox.innerHTML =
      '<div class="text-sm text-gray-400">Click a fabric row to see bale details.</div>';
  }

  try {
    const s = await apiCall("getInventorySummary", {});

    summaryBox.innerHTML = `
      <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
        <div class="p-3 rounded-lg bg-white shadow-sm">
          <div class="text-xs text-gray-500">Total Bales</div>
          <div class="text-xl font-semibold">${s.totalBales || 0}</div>
        </div>
        <div class="p-3 rounded-lg bg-white shadow-sm">
          <div class="text-xs text-gray-500">Total Qty (Yard)</div>
          <div class="text-xl font-semibold">${(
            s.totalQty || 0
          ).toLocaleString()}</div>
        </div>
        <div class="p-3 rounded-lg bg-white shadow-sm">
          <div class="text-xs text-gray-500">Total Designs</div>
          <div class="text-xl font-semibold">${s.designCount || 0}</div>
        </div>
      </div>
    `;

    const fabrics = (s.fabrics || []).sort((a, b) =>
      a.fabricName.localeCompare(b.fabricName)
    );

    if (!fabrics.length) {
      tableBody.innerHTML =
        '<tr><td class="p-2 text-sm text-gray-500" colspan="3">No inventory yet.</td></tr>';
    } else {
      tableBody.innerHTML = "";
      fabrics.forEach((f) => {
        const tr = document.createElement("tr");
        tr.className = "hover:bg-gray-50 cursor-pointer";
        tr.onclick = () => loadInventoryDetails(f.fabricName);
        tr.innerHTML = `
          <td class="p-2 text-sm">${f.fabricName}</td>
          <td class="p-2 text-sm text-center">${f.baleCount || 0}</td>
          <td class="p-2 text-sm text-right">${(f.totalQty || 0).toLocaleString()}</td>
        `;
        tableBody.appendChild(tr);
      });
    }
  } catch (err) {
    summaryBox.innerHTML =
      '<div class="text-sm text-red-500">Error: ' + err.message + "</div>";
    tableBody.innerHTML = "";
  }
}

async function loadInventoryDetails(fabricName) {
  const detailBox = document.getElementById("inventory-detail");
  if (!detailBox) return;

  detailBox.innerHTML = `<div class="text-sm text-gray-500">Loading ${fabricName}...</div>`;

  try {
    const rows = await apiCall("getInventoryDetailsByFabric", { fabricName });

    if (!rows.length) {
      detailBox.innerHTML = `<div class="text-sm text-gray-500">No rows found for ${fabricName}.</div>`;
      return;
    }

    const baleMap = {};
    rows.forEach((r) => {
      const bale = r.baleNo || "";
      if (!bale) return;

      if (!baleMap[bale]) {
        baleMap[bale] = {
          baleNo: bale,
          availableQty: 0,
          designs: new Set(),
          colors: new Set(),
          lines: []
        };
      }

      const bm = baleMap[bale];
      const avail = Number(r.availableQty || 0);
      bm.availableQty += avail;
      if (r.designNo) bm.designs.add(String(r.designNo));
      if (r.color) bm.colors.add(String(r.color));
      bm.lines.push(r);
    });

    const bales = Object.values(baleMap);

    const wrapper = document.createElement("div");
    wrapper.className = "space-y-3";

    bales.forEach((b) => {
      const card = document.createElement("div");
      card.className = "border rounded-lg bg-white";

      const header = document.createElement("div");
      header.className =
        "flex items-center justify-between p-2 cursor-pointer hover:bg-gray-50";

      header.innerHTML = `
        <div>
          <div class="font-semibold text-sm">Bale ${b.baleNo}</div>
          <div class="text-xs text-gray-500">
            Designs: ${b.designs.size} • Colors: ${b.colors.size}
          </div>
        </div>
        <div class="flex items-center gap-3">
          <div class="text-xs md:text-sm text-gray-700">
            Available: <strong>${b.availableQty}</strong> Yard
          </div>
          <button type="button"
            class="px-3 py-1 rounded-md bg-emerald-600 text-white text-xs md:text-sm add-to-sales-btn">
            Add to Sales
          </button>
        </div>
      `;

      const details = document.createElement("div");
      details.className = "border-t px-3 pb-2 pt-2 space-y-1 hidden";

      b.lines.forEach((r) => {
        const div = document.createElement("div");
        div.className = "text-xs md:text-sm text-gray-600";
        div.innerHTML = `
          <div><strong>Design ${r.designNo || ""}</strong> • ${r.color || ""}</div>
          <div>
            Inward: ${r.inwardQty || 0} ${r.uom || ""} •
            Sold: ${r.soldQty || 0} •
            <strong>Available: ${r.availableQty || 0}</strong> •
            Rate: ${r.rate || ""} •
            Total: ${r.total || ""}
          </div>
        `;
        details.appendChild(div);
      });

      header.addEventListener("click", () => {
        details.classList.toggle("hidden");
      });

      const btn = header.querySelector(".add-to-sales-btn");
      if (btn) {
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          addBaleToSalesFromInventory(b.baleNo);
        });
      }

      card.appendChild(header);
      card.appendChild(details);
      wrapper.appendChild(card);
    });

    detailBox.innerHTML = "";
    detailBox.appendChild(wrapper);
  } catch (err) {
    detailBox.innerHTML =
      '<div class="text-sm text-red-500">Error loading details: ' +
      err.message +
      "</div>";
  }
}

// Add bale number into the Sales "cart" input (does not switch tab)
function addBaleToSalesFromInventory(baleNo) {
  const bale = String(baleNo || "").trim();
  if (!bale) return;

  const salesBaleInput = document.getElementById("sales-bale");
  if (salesBaleInput) {
    const existing = salesBaleInput.value
      .split(",")
      .map((b) => b.trim())
      .filter(Boolean);
    if (!existing.includes(bale)) existing.push(bale);
    salesBaleInput.value = existing.join(", ");
  }

  showToast(`Bale ${bale} added to Sales list. Open Sales & click Load.`, true);
}

// ======================================
// SALES – AUTO LOAD QUEUED BALES
// ======================================
function autoLoadQueuedBales() {
  if (hasAutoLoadedQueuedBales) return;
  const salesBaleInput = document.getElementById("sales-bale");
  const tbody = document.getElementById("sales-lines-body");
  if (!salesBaleInput || !tbody) return;

  const value = (salesBaleInput.value || "").trim();
  if (!value) return;
  if (tbody.children.length > 0) return;

  hasAutoLoadedQueuedBales = true;
  loadBaleStock();
}

// ======================================
// SALES TAB LOGIC
// ======================================
let editingInvoiceNo = null; // null = new invoice
let isSavingInvoice = false; // avoid double-save

function initSalesTab() {
  const btnLoadBale = document.getElementById("sales-bale-load");
  const btnLoadCart = document.getElementById("sales-load-cart");
  const btnAddLine = document.getElementById("sales-add-line");
  const btnSave = document.getElementById("sales-save");

  if (!btnAddLine || !btnSave) return;

  const handleLoadClick = () => {
    if (editingInvoiceNo) {
      showToast(
        "In edit mode, bale list comes from existing invoice.",
        false
      );
      return;
    }
    loadBaleStock();
  };

  if (btnLoadBale) btnLoadBale.addEventListener("click", handleLoadClick);
  if (btnLoadCart) btnLoadCart.addEventListener("click", handleLoadClick);

  btnAddLine.addEventListener("click", () => {
    if (editingInvoiceNo) {
      showToast("Add rows directly when editing invoice.", true);
    }
    addSalesLine();
  });

  btnSave.addEventListener("click", saveSalesEntry);

  const dateInput = document.getElementById("sales-date");
  if (dateInput && !dateInput.value) {
    const d = new Date();
    dateInput.value = d.toISOString().slice(0, 10);
  }

  loadRecentInvoices();
}

// Create rows from #sales-bale (comma-separated bale numbers)
async function loadBaleStock() {
  const input = document.getElementById("sales-bale");
  if (!input) return;

  const raw = (input.value || "").trim();
  if (!raw) {
    showToast("Enter at least one Bale No (or add from Inventory).", false);
    return;
  }

  const bales = raw
    .split(",")
    .map((b) => b.trim())
    .filter(Boolean);

  if (!bales.length) {
    showToast("No valid bale numbers found.", false);
    return;
  }

  const tbody = document.getElementById("sales-lines-body");
  if (!tbody) return;

  for (const bale of bales) {
    const tr = addSalesLine();
    const baleInput = tr.querySelector(".baleInput");
    if (baleInput) baleInput.value = bale;
    await loadRowBaleStock(tr);
  }

  showToast("Loaded bale stock into Sales.", true);
}

// === Helper: create one sales line row ===
function addSalesLine() {
  const tbody = document.getElementById("sales-lines-body");
  if (!tbody) return null;

  const tr = document.createElement("tr");
  tr.className = "border-b";

  tr.innerHTML = `
    <td class="p-2">
      <input type="text" class="input baleInput" placeholder="Bale" />
    </td>
    <td class="p-2 fabricNameCell"></td>
    <td class="p-2 fabricQualityCell"></td>

    <td class="p-2">
      <div class="flex flex-col gap-1">
        <select class="input designSelect" multiple size="3"></select>
        <label class="flex items-center gap-1 text-[11px] text-gray-500">
          <input type="checkbox" class="designAllChk">
          <span>Select all</span>
        </label>
      </div>
    </td>

    <td class="p-2">
      <div class="flex flex-col gap-1">
        <select class="input colorSelect" multiple size="3"></select>
        <label class="flex items-center gap-1 text-[11px] text-gray-500">
          <input type="checkbox" class="colorAllChk">
          <span>Select all</span>
        </label>
      </div>
    </td>

    <td class="p-2 text-right onePcsQtyCell"></td>
    <td class="p-2 text-right availQtyCell"></td>
    <td class="p-2 text-right">
      <input type="number" class="input text-right qtyInput" min="0" step="0.01" />
    </td>
    <td class="p-2 uomCell"></td>
    <td class="p-2 text-right inwardRateCell"></td>
    <td class="p-2 text-right">
      <input type="number" class="input text-right sellingRateInput" min="0" step="0.01" />
    </td>
    <td class="p-2 text-right amountCell"></td>
    <td class="p-2 text-center">
      <button class="text-xs text-red-500" type="button">✕</button>
    </td>
  `;

  tbody.appendChild(tr);

  const baleInput = tr.querySelector(".baleInput");
  const designSelect = tr.querySelector(".designSelect");
  const colorSelect = tr.querySelector(".colorSelect");
  const designAllChk = tr.querySelector(".designAllChk");
  const colorAllChk = tr.querySelector(".colorAllChk");
  const deleteBtn = tr.querySelector("button");
  const qtyInput = tr.querySelector(".qtyInput");
  const rateInput = tr.querySelector(".sellingRateInput");

  baleInput.addEventListener("change", () => loadRowBaleStock(tr));
  baleInput.addEventListener("blur", () => {
    if (!tr._stockList) loadRowBaleStock(tr);
  });

  designSelect.addEventListener("change", () => {
    designAllChk.checked =
      designSelect.selectedOptions.length === designSelect.options.length;
    applySelectionToRow(tr);
  });
  designAllChk.addEventListener("change", () => {
    const all = designAllChk.checked;
    Array.from(designSelect.options).forEach((opt) => {
      opt.selected = all;
    });
    applySelectionToRow(tr);
  });

  colorSelect.addEventListener("change", () => {
    colorAllChk.checked =
      colorSelect.selectedOptions.length === colorSelect.options.length;
    applySelectionToRow(tr);
  });
  colorAllChk.addEventListener("change", () => {
    const all = colorAllChk.checked;
    Array.from(colorSelect.options).forEach((opt) => {
      opt.selected = all;
    });
    applySelectionToRow(tr);
  });

  if (qtyInput) {
    qtyInput.addEventListener("input", () => {
      enforceQtyLimit(tr);
      recalcInvoiceTotal();
    });
  }
  if (rateInput) {
    rateInput.addEventListener("input", () => {
      recalcInvoiceTotal();
    });
  }

  deleteBtn.addEventListener("click", () => {
    tr.remove();
    recalcInvoiceTotal();
  });

  return tr;
}

// === Multi-select helpers ===
function getSelectedValues(selectEl) {
  if (!selectEl) return [];
  return Array.from(selectEl.selectedOptions || []).map((o) => o.value);
}
function setAllOptionsSelected(selectEl) {
  if (!selectEl) return;
  Array.from(selectEl.options).forEach((o) => (o.selected = true));
}

// recompute row summary from selected designs / colors
function applySelectionToRow(row) {
  const list = row._stockList || [];
  if (!list.length) {
    clearRowDetail(row);
    return;
  }

  const designSelect = row.querySelector(".designSelect");
  const colorSelect = row.querySelector(".colorSelect");

  let selectedDesigns = getSelectedValues(designSelect);
  let selectedColors = getSelectedValues(colorSelect);

  if (!selectedDesigns.length) {
    selectedDesigns = Array.from(new Set(list.map((i) => String(i.designNo))));
  }
  if (!selectedColors.length) {
    selectedColors = Array.from(new Set(list.map((i) => String(i.color))));
  }

  const chosen = list.filter(
    (i) =>
      selectedDesigns.includes(String(i.designNo)) &&
      selectedColors.includes(String(i.color))
  );

  if (!chosen.length) {
    clearRowDetail(row);
    return;
  }

  row._selectedStockItems = chosen;
  const first = chosen[0];

  const totalAvailable = chosen.reduce(
    (sum, i) => sum + Number(i.availableQty || 0),
    0
  );

  row.dataset.fabricName = first.fabricName || "";
  row.dataset.fabricQuality = first.fabricQuality || "";
  row.dataset.uom = first.uom || "";
  row.dataset.inwardRate = Number(first.inwardRate || 0);
  row.dataset.onePcsQty = Number(first.onePcsQty || 0);
  row.dataset.availableQty = totalAvailable;
  row.dataset.color =
    selectedColors.length === 1 ? selectedColors[0] : "MULTI";

  row.querySelector(".fabricNameCell").textContent = first.fabricName || "";
  row.querySelector(".fabricQualityCell").textContent =
    first.fabricQuality || "";
  row.querySelector(".onePcsQtyCell").textContent = first.onePcsQty || "";
  row.querySelector(".availQtyCell").textContent = totalAvailable || 0;
  row.querySelector(".uomCell").textContent = first.uom || "";
  row.querySelector(".inwardRateCell").textContent = first.inwardRate || "";

  const rateInput = row.querySelector(".sellingRateInput");
  if (rateInput && !rateInput.value) {
    rateInput.value = first.inwardRate || 0;
    rateInput.dataset.inwardRate = first.inwardRate || 0;
  }

  const qtyInput = row.querySelector(".qtyInput");
  if (qtyInput && !qtyInput.value) {
    qtyInput.value = totalAvailable || 0;
  }

  recalcRowAmount(row);
  recalcInvoiceTotal();
}

// Load stock only for THIS row's bale (multi design / color)
async function loadRowBaleStock(row) {
  const baleInput = row.querySelector(".baleInput");
  const baleNo = (baleInput?.value || "").trim();
  if (!baleNo) return;

  try {
    const data = await apiCall("getStockByBale", { baleNo });
    const list = data || [];
    row._stockList = list;

    const designSelect = row.querySelector(".designSelect");
    const colorSelect = row.querySelector(".colorSelect");
    const designAllChk = row.querySelector(".designAllChk");
    const colorAllChk = row.querySelector(".colorAllChk");

    designSelect.innerHTML = "";
    colorSelect.innerHTML = "";
    clearRowDetail(row);

    if (!list.length) {
      showToast(`No stock for bale ${baleNo}`, false);
      return;
    }

    const designs = Array.from(new Set(list.map((i) => String(i.designNo))));
    const colors = Array.from(new Set(list.map((i) => String(i.color))));

    designs.forEach((d) => {
      const opt = document.createElement("option");
      opt.value = d;
      opt.textContent = d;
      opt.selected = true;
      designSelect.appendChild(opt);
    });

    colors.forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      opt.selected = true;
      colorSelect.appendChild(opt);
    });

    if (designAllChk) designAllChk.checked = true;
    if (colorAllChk) colorAllChk.checked = true;

    const first = list[0];
    row.dataset.fabricName = first.fabricName || "";
    row.dataset.fabricQuality = first.fabricQuality || "";
    row.querySelector(".fabricNameCell").textContent = first.fabricName || "";
    row.querySelector(".fabricQualityCell").textContent =
      first.fabricQuality || "";

    applySelectionToRow(row);
  } catch (err) {
    console.error(err);
    showToast(`Failed to load bale ${baleNo}: ${err.message}`, false);
  }
}

// Ensure qty never exceeds available for that row’s selection
function enforceQtyLimit(row) {
  const qtyInput = row.querySelector(".qtyInput");
  if (!qtyInput) return;
  const available = Number(row.dataset.availableQty || 0);
  if (!available) return;

  let qty = Number(qtyInput.value || 0);
  if (qty > available) {
    qtyInput.value = available;
    showToast(
      `Qty cannot exceed available (${available}) for this bale / selection`,
      false
    );
  } else if (qty < 0) {
    qtyInput.value = 0;
  }

  recalcRowAmount(row);
}

// Clear row details
function clearRowDetail(row) {
  const setText = (selector, value = "") => {
    const el = row.querySelector(selector);
    if (el) el.textContent = value;
  };
  setText(".fabricNameCell");
  setText(".fabricQualityCell");
  setText(".onePcsQtyCell");
  setText(".availQtyCell");
  setText(".uomCell");
  setText(".inwardRateCell");
  setText(".amountCell");

  const rateInput = row.querySelector(".sellingRateInput");
  if (rateInput) {
    rateInput.value = "";
    rateInput.dataset.inwardRate = "";
  }

  row.dataset.fabricName = "";
  row.dataset.fabricQuality = "";
  row.dataset.color = "";
  row.dataset.onePcsQty = "";
  row.dataset.uom = "";
  row.dataset.inwardRate = "";
  row.dataset.availableQty = "";
}

// Amount & totals
function recalcRowAmount(row) {
  const qtyInput = row.querySelector(".qtyInput");
  const rateInput = row.querySelector(".sellingRateInput");
  const amountCell = row.querySelector(".amountCell");
  if (!qtyInput || !rateInput || !amountCell) return 0;

  const qty = Number(qtyInput.value || 0);
  const rate = Number(rateInput.value || 0);
  const amount = qty * rate;
  amountCell.textContent = amount > 0 ? amount.toFixed(2) : "";
  return amount;
}

function recalcInvoiceTotal() {
  let total = 0;
  const rows = document.querySelectorAll("#sales-lines-body tr");
  rows.forEach((row) => {
    total += recalcRowAmount(row);
  });
  const totalEl = document.getElementById("sales-total-value");
  if (totalEl) totalEl.textContent = total.toFixed(2);
}

// ===============================
// SAVE INVOICE (with validations)
// ===============================
async function saveSalesEntry() {
  // block double-click while saving
  if (isSavingInvoice) return;

  const customerInput = document.getElementById("sales-customer");
  const customerName = (customerInput?.value || "").trim();

  // 1) Customer required
  if (!customerName) {
    showToast("Please enter Customer Name.", false);
    if (customerInput) customerInput.focus();
    return;
  }

  // 2) At least one row present
  const rows = Array.from(
    document.querySelectorAll("#sales-lines-body tr")
  );
  if (!rows.length) {
    showToast("Please add at least one item before saving.", false);
    return;
  }

  // 3) Build items
  const items = [];

  for (const row of rows) {
    const designSelect = row.querySelector(".designSelect");
    const colorSelect = row.querySelector(".colorSelect");
    const qtyInput = row.querySelector(".qtyInput");
    const sellingInput = row.querySelector(".sellingRateInput");
    const baleInput = row.querySelector(".baleInput");

    const rowQty = Number(qtyInput?.value || 0);
    const selectedDesigns = Array.from(designSelect.selectedOptions)
      .map((o) => o.value)
      .filter(Boolean);
    const selectedColors = Array.from(colorSelect.selectedOptions)
      .map((o) => o.value)
      .filter(Boolean);

    if (!selectedDesigns.length || !selectedColors.length || rowQty <= 0)
      continue;

    const rowBales = (baleInput?.value || "").trim();

    const designCsv = selectedDesigns.join(",");
    const colorCsv = selectedColors.join(",");

    const inwardRate = Number(
      sellingInput.dataset.inwardRate || row.dataset.inwardRate || 0
    );
    const sellingRate = Number(sellingInput.value || 0);

    if (sellingRate < inwardRate) {
      showToast(`Selling price below inward for ${designCsv}`, false);
      return;
    }

    items.push({
      baleNos: rowBales,
      fabricName: row.dataset.fabricName,
      fabricQuality: row.dataset.fabricQuality,
      designNo: designCsv,
      color: colorCsv,
      onePcsQty: Number(row.dataset.onePcsQty || 0),
      qty: rowQty,
      uom: row.dataset.uom,
      inwardRate,
      sellingRate
    });
  }

  if (!items.length) {
    showToast("Add at least one valid item", false);
    return;
  }

  // 4) Header fields
  const date = document.getElementById("sales-date").value;
  const companyName = document
    .getElementById("sales-company")
    .value.trim();
  const invoiceNoInput = document
    .getElementById("sales-invoice")
    .value.trim();
  const email = document.getElementById("sales-email").value.trim();
  const website = document.getElementById("sales-website").value.trim();
  const vatNo = document.getElementById("sales-vat").value.trim();
  const customerImageUrl = document
    .getElementById("sales-image")
    .value.trim();

  const header = {
    date,
    customerName,
    companyName,
    email,
    website,
    vatNo,
    customerImageUrl,
    invoiceNo: editingInvoiceNo || invoiceNoInput
  };

  // 5) Saving state (disable button)
  const btnSave = document.getElementById("sales-save");
  let oldLabel = "";
  isSavingInvoice = true;
  if (btnSave) {
    oldLabel = btnSave.textContent;
    btnSave.disabled = true;
    btnSave.textContent = "Saving...";
  }

  try {
    const action = editingInvoiceNo ? "updateInvoice" : "saveOutwardEntry";
    const data = await apiCall(action, { header, items });

    showToast(
      (editingInvoiceNo ? "Invoice updated: " : "Invoice saved: ") +
        data.invoiceNo,
      true
    );

    if (data.pdfUrl) {
      window.open(data.pdfUrl, "_blank");
    }

    document.getElementById("sales-lines-body").innerHTML = "";
    editingInvoiceNo = null;
    document.getElementById("sales-mode-label").textContent =
      "New invoice mode";
    document.getElementById("sales-invoice").value = data.invoiceNo;
    recalcInvoiceTotal();
    loadRecentInvoices();
    loadDashboardFromUI();
  } catch (err) {
    console.error(err);
    showToast(err.message || "Failed to save invoice", false);
  } finally {
    isSavingInvoice = false;
    if (btnSave) {
      btnSave.disabled = false;
      btnSave.textContent = oldLabel || "Save Invoice";
    }
  }
}

// ===============================
// Recent invoices + edit support
// ===============================
async function loadRecentInvoices() {
  const tbody = document.getElementById("recent-invoices-body");
  if (!tbody) return;

  tbody.innerHTML =
    '<tr><td colspan="6" class="p-2 text-xs text-gray-500">Loading...</td></tr>';

  try {
    const list = await apiCall("getRecentInvoices", {});
    if (!list.length) {
      tbody.innerHTML =
        '<tr><td colspan="6" class="p-2 text-xs text-gray-500">No invoices yet.</td></tr>';
      return;
    }

    tbody.innerHTML = "";
    list.forEach((inv) => {
      const tr = document.createElement("tr");
      tr.className = "border-b text-xs md:text-sm";
      const dateStr = inv.date
        ? new Date(inv.date).toISOString().slice(0, 10)
        : "";
      tr.innerHTML = `
        <td class="p-2">${inv.invoiceNo}</td>
        <td class="p-2">${dateStr}</td>
        <td class="p-2">${inv.customerName || ""}</td>
        <td class="p-2 text-right">${(inv.total || 0).toFixed(2)}</td>
        <td class="p-2">
          ${
            inv.pdfUrl
              ? `<a href="${inv.pdfUrl}" target="_blank" class="text-emerald-600 underline">View</a>`
              : ""
          }
        </td>
        <td class="p-2 text-center">
          <button class="text-xs text-blue-600 underline" type="button"
                  onclick="startEditInvoice('${inv.invoiceNo.replace(
                    /'/g,
                    "\\'"
                  )}')">
            Edit
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  } catch (err) {
    console.error(err);
    tbody.innerHTML =
      '<tr><td colspan="6" class="p2 text-xs text-red-500">Error: ' +
      (err.message || err) +
      "</td></tr>";
  }
}

async function startEditInvoice(invoiceNo) {
  try {
    const res = await apiCall("getInvoiceDetails", { invoiceNo });
    const h = res.header;
    const items = res.items || [];

    editingInvoiceNo = invoiceNo;
    document.getElementById("sales-mode-label").textContent =
      "Editing invoice " + invoiceNo;

    document.getElementById("sales-date").value = h.date
      ? new Date(h.date).toISOString().slice(0, 10)
      : "";
    document.getElementById("sales-customer").value = h.customerName || "";
    document.getElementById("sales-company").value = h.companyName || "";
    document.getElementById("sales-email").value = h.email || "";
    document.getElementById("sales-website").value = h.website || "";
    document.getElementById("sales-vat").value = h.vatNo || "";
    document.getElementById("sales-image").value = h.customerImageUrl || "";
    document.getElementById("sales-invoice").value = h.invoiceNo || "";
    document.getElementById("sales-bale").value = h.baleNos || "";

    const tbody = document.getElementById("sales-lines-body");
    tbody.innerHTML = "";

    for (const item of items) {
      const tr = addSalesLine();

      const baleInput = tr.querySelector(".baleInput");
      baleInput.value = item.baleNos || "";
      await loadRowBaleStock(tr);

      const designSelect = tr.querySelector(".designSelect");
      const colorSelect = tr.querySelector(".colorSelect");

      const designVals = String(item.designNo || "")
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
      const colorVals = String(item.color || "")
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);

      Array.from(designSelect.options).forEach(
        (opt) => (opt.selected = designVals.includes(opt.value))
      );
      Array.from(colorSelect.options).forEach(
        (opt) => (opt.selected = colorVals.includes(opt.value))
      );

      applySelectionToRow(tr);

      const qtyInput = tr.querySelector(".qtyInput");
      const rateInput = tr.querySelector(".sellingRateInput");

      if (qtyInput) {
        qtyInput.value = item.qty || 0;
      }
      if (rateInput) {
        rateInput.value = item.sellingRate || 0;
        rateInput.dataset.inwardRate = item.inwardRate || 0;

        rateInput.onblur = () => {
          const inward = Number(rateInput.dataset.inwardRate || 0);
          const selling = Number(rateInput.value || 0);
          if (selling < inward) {
            showToast(
              `Selling price cannot be below inward price (${inward})`,
              false
            );
            rateInput.value = inward;
            rateInput.focus();
          }
        };
      }

      recalcRowAmount(tr);
    }

    recalcInvoiceTotal();
    showToast("Loaded invoice " + invoiceNo + " for editing", true);
  } catch (err) {
    console.error(err);
    showToast(err.message || "Failed to load invoice", false);
  }
}

// Refresh App button
function resetApp() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.getRegistrations().then((regs) => {
      regs.forEach((reg) => reg.unregister());
      caches.keys().then((keys) => keys.forEach((key) => caches.delete(key)));
      showToast("Refreshing…");
      setTimeout(() => location.reload(true), 800);
    });
  } else {
    location.reload(true);
  }
}
