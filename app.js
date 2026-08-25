(() => {
  "use strict";

  const STORAGE_KEY = "expenseTrackerData_v1";

  const PRESET_CATEGORIES = [
    "Acessórios", "Apostas", "Cabeleireiro", "Comida", "Gasolina", "Geral",
    "Ginásio", "Investimentos", "IRS", "Lazer", "Ordenado", "Prendas",
    "Produtos de Beleza", "Restaurantes", "Roupa", "Saídas à Noite", "Saúde",
    "Shein", "Subscrições", "Supermercado", "Tecnologia", "Transportes",
    "Unhas", "Viagens"
  ];
  const PRESET_ACCOUNTS = ["Novo Banco", "Revolut", "Trading 212"];

  // ---------- storage ----------

  function loadData() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error("Failed to load stored data", e);
      return [];
    }
  }

  function saveData() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.data));
  }

  function uuid() {
    if (crypto.randomUUID) return crypto.randomUUID();
    return "id-" + Date.now() + "-" + Math.random().toString(16).slice(2);
  }

  const state = {
    data: loadData(),
    editId: null,
    sort: { key: "date", dir: "desc" },
    search: "",
    pendingDelete: null,
  };

  // ---------- helpers ----------

  function normalize(s) {
    return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  }

  function fmtMoney(n) {
    const v = Number(n) || 0;
    const sign = v < 0 ? "-" : "";
    return sign + "€" + Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function toast(msg, opts = {}) {
    const el = document.getElementById("toast");
    const msgEl = document.getElementById("toastMessage");
    const actionEl = document.getElementById("toastAction");

    msgEl.textContent = msg;
    if (opts.actionLabel && opts.onAction) {
      actionEl.textContent = opts.actionLabel;
      actionEl.classList.remove("hidden");
      actionEl.onclick = () => {
        opts.onAction();
        el.classList.add("hidden");
        clearTimeout(toast._t);
      };
    } else {
      actionEl.classList.add("hidden");
      actionEl.onclick = null;
    }

    el.classList.remove("hidden");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.add("hidden"), opts.duration || 3000);
  }

  function todayISO() {
    const d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }

  function typeFromRaw(raw) {
    const n = normalize(raw);
    if (n === "rendimento" || n === "income") return "income";
    if (n === "investimento" || n === "investment") return "investment";
    return "expense";
  }

  // Excel serial date (or JS Date from SheetJS cellDates) -> "YYYY-MM-DD"
  function excelValueToISO(val) {
    if (val instanceof Date) {
      return val.getUTCFullYear() + "-" + String(val.getUTCMonth() + 1).padStart(2, "0") + "-" + String(val.getUTCDate()).padStart(2, "0");
    }
    if (typeof val === "number") {
      const ms = Date.UTC(1899, 11, 30) + val * 86400000;
      const d = new Date(ms);
      return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" + String(d.getUTCDate()).padStart(2, "0");
    }
    if (typeof val === "string") {
      const parsed = new Date(val);
      if (!isNaN(parsed)) {
        return parsed.getFullYear() + "-" + String(parsed.getMonth() + 1).padStart(2, "0") + "-" + String(parsed.getDate()).padStart(2, "0");
      }
    }
    return null;
  }

  // ---------- categories / accounts datalists ----------

  function refreshDatalists() {
    const cats = new Set(PRESET_CATEGORIES);
    const accs = new Set(PRESET_ACCOUNTS);
    state.data.forEach((e) => {
      if (e.category) cats.add(e.category);
      if (e.account) accs.add(e.account);
    });
    const catList = document.getElementById("categoryList");
    catList.innerHTML = [...cats].sort().map((c) => `<option value="${escapeHtml(c)}">`).join("");
    const accList = document.getElementById("accountList");
    accList.innerHTML = [...accs].sort().map((a) => `<option value="${escapeHtml(a)}">`).join("");
  }

  function escapeHtml(s) {
    return (s || "").toString().replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  // ---------- form: add / edit ----------

  const form = document.getElementById("expenseForm");
  const formTitle = document.getElementById("formTitle");
  const submitBtn = document.getElementById("submitBtn");
  const cancelEditBtn = document.getElementById("cancelEditBtn");

  document.getElementById("date").value = todayISO();

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const amount = parseFloat(document.getElementById("amount").value);
    const category = document.getElementById("category").value.trim();
    const date = document.getElementById("date").value;
    const note = document.getElementById("note").value.trim();
    const type = document.getElementById("type").value;
    const account = document.getElementById("account").value.trim();

    if (!amount || amount <= 0 || !category || !date) {
      toast("Please fill in amount, category and date.");
      return;
    }

    if (state.editId) {
      const rec = state.data.find((e) => e.id === state.editId);
      if (rec) {
        Object.assign(rec, { amount, category, date, note, type, account, updatedAt: Date.now() });
        toast("Expense updated.");
      }
    } else {
      state.data.push({
        id: uuid(), amount, category, date, note, type, account,
        createdAt: Date.now(), updatedAt: Date.now(),
      });
      toast("Expense added.");
    }

    saveData();
    exitEditMode();
    renderAll();
  });

  cancelEditBtn.addEventListener("click", () => exitEditMode());

  function enterEditMode(rec) {
    state.editId = rec.id;
    document.getElementById("expenseId").value = rec.id;
    document.getElementById("amount").value = rec.amount;
    document.getElementById("category").value = rec.category;
    document.getElementById("date").value = rec.date;
    document.getElementById("note").value = rec.note || "";
    document.getElementById("type").value = rec.type || "expense";
    document.getElementById("account").value = rec.account || "";
    formTitle.textContent = "Edit Expense";
    submitBtn.textContent = "Save Changes";
    cancelEditBtn.classList.remove("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function exitEditMode() {
    state.editId = null;
    form.reset();
    document.getElementById("date").value = todayISO();
    document.getElementById("type").value = "expense";
    formTitle.textContent = "Add Expense";
    submitBtn.textContent = "Add Expense";
    cancelEditBtn.classList.add("hidden");
  }

  function deleteExpense(id) {
    const index = state.data.findIndex((e) => e.id === id);
    if (index === -1) return;

    const [record] = state.data.splice(index, 1);
    if (state.editId === id) exitEditMode();
    saveData();
    renderAll();

    const pending = { record, index };
    pending.timer = setTimeout(() => {
      if (state.pendingDelete === pending) state.pendingDelete = null;
    }, 10000);
    state.pendingDelete = pending;

    toast("Expense deleted.", {
      duration: 10000,
      actionLabel: "Undo",
      onAction: () => {
        if (state.pendingDelete !== pending) return;
        clearTimeout(pending.timer);
        state.data.splice(Math.min(pending.index, state.data.length), 0, pending.record);
        state.pendingDelete = null;
        saveData();
        renderAll();
        toast("Expense restored.");
      },
    });
  }

  // ---------- date range filter (dashboard only) ----------

  const rangeSelect = document.getElementById("rangeSelect");
  const customRange = document.getElementById("customRange");
  const customFrom = document.getElementById("customFrom");
  const customTo = document.getElementById("customTo");

  rangeSelect.addEventListener("change", () => {
    customRange.classList.toggle("hidden", rangeSelect.value !== "custom");
    renderDashboard();
  });
  customFrom.addEventListener("change", renderDashboard);
  customTo.addEventListener("change", renderDashboard);

  function getRangeBounds() {
    const now = new Date();
    const mode = rangeSelect.value;
    if (mode === "all") return { from: null, to: null };
    if (mode === "year") {
      return { from: `${now.getFullYear()}-01-01`, to: `${now.getFullYear()}-12-31` };
    }
    if (mode === "month") {
      const y = now.getFullYear(), m = now.getMonth() + 1;
      const lastDay = new Date(y, m, 0).getDate();
      return { from: `${y}-${String(m).padStart(2, "0")}-01`, to: `${y}-${String(m).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}` };
    }
    if (mode === "custom") {
      return { from: customFrom.value || null, to: customTo.value || null };
    }
    return { from: null, to: null };
  }

  function filterByRange(data) {
    const { from, to } = getRangeBounds();
    return data.filter((e) => (!from || e.date >= from) && (!to || e.date <= to));
  }

  // ---------- dashboard ----------

  let categoryChart = null;
  let timeChart = null;

  // "Accessible Orchid" data colors, taken from the theme used in 2026.pbix
  const BASE_PALETTE = ["#364B59", "#C480A7", "#663466", "#6E98B5", "#67002E", "#5D8099", "#3F213F", "#EC64A9"];
  const SHADE_STEPS = [0, 0.35, -0.3, 0.6, -0.55];

  function shadeColor(hex, percent) {
    const num = parseInt(hex.slice(1), 16);
    const r = (num >> 16) & 0xff, g = (num >> 8) & 0xff, b = num & 0xff;
    const mix = (c) => percent >= 0 ? c + (255 - c) * percent : c * (1 + percent);
    return "#" + [r, g, b].map((c) => Math.round(mix(c)).toString(16).padStart(2, "0")).join("");
  }

  function paletteColor(i) {
    const base = BASE_PALETTE[i % BASE_PALETTE.length];
    const cycle = Math.floor(i / BASE_PALETTE.length);
    if (cycle === 0) return base;
    return shadeColor(base, SHADE_STEPS[cycle % SHADE_STEPS.length]);
  }

  function renderDashboard() {
    const filtered = filterByRange(state.data);
    const expenses = filtered.filter((e) => e.type === "expense");
    const income = filtered.filter((e) => e.type === "income");
    const investments = filtered.filter((e) => e.type === "investment");

    const totalSpending = expenses.reduce((s, e) => s + e.amount, 0);
    const totalIncome = income.reduce((s, e) => s + e.amount, 0);
    const totalInvested = investments.reduce((s, e) => s + e.amount, 0);
    const savings = totalIncome - totalSpending;

    document.getElementById("statTotalSpending").textContent = fmtMoney(totalSpending);
    document.getElementById("statTotalIncome").textContent = fmtMoney(totalIncome);
    document.getElementById("statSavings").textContent = fmtMoney(savings);
    document.getElementById("statInvested").textContent = fmtMoney(totalInvested);

    renderCategoryChart(expenses);
    renderTimeChart(expenses);
  }

  function renderCategoryChart(expenses) {
    const ctx = document.getElementById("categoryChart");
    const emptyNote = document.getElementById("categoryEmpty");

    const totals = {};
    expenses.forEach((e) => { totals[e.category] = (totals[e.category] || 0) + e.amount; });
    const labels = Object.keys(totals).sort((a, b) => totals[b] - totals[a]);

    if (categoryChart) { categoryChart.destroy(); categoryChart = null; }

    if (labels.length === 0) {
      ctx.classList.add("hidden");
      emptyNote.classList.remove("hidden");
      return;
    }
    ctx.classList.remove("hidden");
    emptyNote.classList.add("hidden");

    categoryChart = new Chart(ctx, {
      type: "pie",
      data: {
        labels,
        datasets: [{
          data: labels.map((l) => totals[l]),
          backgroundColor: labels.map((_, i) => paletteColor(i)),
        }],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: window.innerWidth < 600 ? "bottom" : "right", labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${fmtMoney(ctx.parsed)}` } },
        },
      },
    });
  }

  function renderTimeChart(expenses) {
    const ctx = document.getElementById("timeChart");
    const emptyNote = document.getElementById("timeEmpty");

    if (timeChart) { timeChart.destroy(); timeChart = null; }

    if (expenses.length === 0) {
      ctx.classList.add("hidden");
      emptyNote.classList.remove("hidden");
      return;
    }
    ctx.classList.remove("hidden");
    emptyNote.classList.add("hidden");

    const dates = expenses.map((e) => e.date).sort();
    const spanDays = (new Date(dates[dates.length - 1]) - new Date(dates[0])) / 86400000;
    const byDay = spanDays <= 62;

    const totals = {};
    expenses.forEach((e) => {
      const key = byDay ? e.date : e.date.slice(0, 7);
      totals[key] = (totals[key] || 0) + e.amount;
    });
    const labels = Object.keys(totals).sort();

    timeChart = new Chart(ctx, {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: "Spending",
          data: labels.map((l) => totals[l]),
          borderColor: "#364B59",
          backgroundColor: "rgba(54, 75, 89, 0.12)",
          fill: true,
          tension: 0.25,
          pointRadius: labels.length > 40 ? 0 : 3,
        }],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (ctx) => fmtMoney(ctx.parsed.y) } },
        },
        scales: {
          y: { ticks: { callback: (v) => fmtMoney(v) } },
        },
      },
    });
  }

  // ---------- table ----------

  const tbody = document.getElementById("expenseTableBody");
  const tableEmpty = document.getElementById("tableEmpty");
  const searchInput = document.getElementById("searchInput");

  searchInput.addEventListener("input", () => {
    state.search = searchInput.value;
    renderTable();
  });

  document.querySelectorAll("#expenseTable thead th[data-sort]").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.sort;
      if (state.sort.key === key) {
        state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
      } else {
        state.sort = { key, dir: "asc" };
      }
      renderTable();
    });
  });

  function renderTable() {
    const q = normalize(state.search);
    let rows = state.data.filter((e) => !q || normalize(e.note).includes(q) || normalize(e.category).includes(q));

    const { key, dir } = state.sort;
    rows = rows.slice().sort((a, b) => {
      let av = a[key], bv = b[key];
      if (key === "amount") { av = Number(av); bv = Number(bv); }
      else { av = (av || "").toString().toLowerCase(); bv = (bv || "").toString().toLowerCase(); }
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });

    if (rows.length === 0) {
      tbody.innerHTML = "";
      tableEmpty.classList.remove("hidden");
      return;
    }
    tableEmpty.classList.add("hidden");

    tbody.innerHTML = rows.map((e) => `
      <tr>
        <td data-label="Date">${e.date}</td>
        <td class="amount-${e.type}" data-label="Amount">${fmtMoney(e.amount)}</td>
        <td data-label="Category">${escapeHtml(e.category)}</td>
        <td data-label="Type"><span class="type-badge ${e.type}">${e.type}</span></td>
        <td data-label="Account">${escapeHtml(e.account || "")}</td>
        <td data-label="Note">${escapeHtml(e.note || "")}</td>
        <td class="row-actions">
          <button class="btn btn-icon" data-edit="${e.id}">Edit</button>
          <button class="btn btn-danger" data-delete="${e.id}">Delete</button>
        </td>
      </tr>
    `).join("");

    tbody.querySelectorAll("[data-edit]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const rec = state.data.find((e) => e.id === btn.dataset.edit);
        if (rec) enterEditMode(rec);
      });
    });
    tbody.querySelectorAll("[data-delete]").forEach((btn) => {
      btn.addEventListener("click", () => deleteExpense(btn.dataset.delete));
    });
  }

  // ---------- export / restore JSON ----------

  document.getElementById("exportBtn").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state.data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `expense-tracker-backup-${todayISO()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  });

  const TYPE_LABELS_PT = { expense: "Despesa", income: "Rendimento", investment: "Investimento" };

  document.getElementById("exportExcelBtn").addEventListener("click", () => {
    if (state.data.length === 0) {
      toast("No transactions to export.");
      return;
    }
    const rows = state.data
      .slice()
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0))
      .map((e) => ({
        Tipo: TYPE_LABELS_PT[e.type] || "Despesa",
        Data: new Date(e.date + "T00:00:00Z"),
        Valor: e.amount,
        "Descrição": e.note || "",
        Categoria: e.category,
        App: e.account || "",
      }));
    const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Finanças");
    XLSX.writeFile(wb, `expense-tracker-${todayISO()}.xlsx`, { cellDates: true });
  });

  document.getElementById("restoreInput").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!Array.isArray(parsed)) throw new Error("Invalid backup file");
        if (!confirm(`Restore ${parsed.length} transactions from this backup? This will REPLACE all current data (${state.data.length} transactions).`)) return;
        state.data = parsed;
        saveData();
        renderAll();
        toast("Backup restored.");
      } catch (err) {
        toast("Could not read backup file: " + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  });

  // ---------- Excel import ----------

  const importInput = document.getElementById("importInput");
  const importModal = document.getElementById("importModal");
  const importSheetSelect = document.getElementById("importSheetSelect");
  const importPreview = document.getElementById("importPreview");
  const importConfirmBtn = document.getElementById("importConfirmBtn");

  let importWorkbook = null;
  let importPendingRows = [];

  document.getElementById("importBtn").addEventListener("click", () => importInput.click());

  importInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        importWorkbook = XLSX.read(evt.target.result, { type: "array", cellDates: true });
      } catch (err) {
        toast("Could not read Excel file: " + err.message);
        return;
      }
      const names = importWorkbook.SheetNames;
      importSheetSelect.innerHTML = names.map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
      const preferred = names.find((n) => normalize(n).includes("financ")) || names[0];
      importSheetSelect.value = preferred;
      updateImportPreview();
      importModal.classList.remove("hidden");
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  });

  importSheetSelect.addEventListener("change", updateImportPreview);

  const COLUMN_SYNONYMS = {
    date: ["data", "date"],
    amount: ["valor", "amount", "value"],
    category: ["categoria", "category"],
    note: ["descricao", "description", "note", "notes"],
    type: ["tipo", "type"],
    account: ["app", "account", "conta"],
  };

  function detectColumns(headerKeys) {
    const map = {};
    const normalizedKeys = headerKeys.map((h) => ({ raw: h, norm: normalize(h) }));
    for (const [field, synonyms] of Object.entries(COLUMN_SYNONYMS)) {
      const match = normalizedKeys.find((h) => synonyms.includes(h.norm));
      if (match) map[field] = match.raw;
    }
    return map;
  }

  function updateImportPreview() {
    const ws = importWorkbook.Sheets[importSheetSelect.value];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (rows.length === 0) {
      importPreview.textContent = "This sheet has no data rows.";
      importPendingRows = [];
      importConfirmBtn.disabled = true;
      return;
    }
    const cols = detectColumns(Object.keys(rows[0]));
    if (!cols.date || !cols.amount) {
      importPreview.textContent = `Couldn't find recognizable "date" and "amount" columns in this sheet. Columns found: ${Object.keys(rows[0]).join(", ")}`;
      importPendingRows = [];
      importConfirmBtn.disabled = true;
      return;
    }

    const mapped = [];
    for (const row of rows) {
      const iso = excelValueToISO(row[cols.date]);
      const amount = parseFloat(row[cols.amount]);
      if (!iso || isNaN(amount)) continue;
      mapped.push({
        date: iso,
        amount: Math.abs(amount),
        category: (cols.category ? row[cols.category] : "").toString().trim() || "Uncategorized",
        note: (cols.note ? row[cols.note] : "").toString().trim(),
        type: cols.type ? typeFromRaw(row[cols.type]) : "expense",
        account: (cols.account ? row[cols.account] : "").toString().trim(),
      });
    }

    importPendingRows = mapped;
    const skipped = rows.length - mapped.length;
    importPreview.textContent = `Found ${mapped.length} valid transaction(s) using columns: ${Object.entries(cols).map(([k, v]) => `${k}=${v}`).join(", ")}.` +
      (skipped > 0 ? ` (${skipped} row(s) skipped: missing date/amount.)` : "");
    importConfirmBtn.disabled = mapped.length === 0;
  }

  function rowKey(e) {
    return [e.date, e.amount.toFixed(2), normalize(e.category), normalize(e.note), e.type].join("|");
  }

  importConfirmBtn.addEventListener("click", () => {
    const existingKeys = new Set(state.data.map(rowKey));
    let added = 0, skipped = 0;
    for (const row of importPendingRows) {
      const key = rowKey(row);
      if (existingKeys.has(key)) { skipped++; continue; }
      existingKeys.add(key);
      state.data.push({ id: uuid(), ...row, createdAt: Date.now(), updatedAt: Date.now() });
      added++;
    }
    saveData();
    renderAll();
    importModal.classList.add("hidden");
    toast(`Imported ${added} transaction(s).` + (skipped > 0 ? ` Skipped ${skipped} duplicate(s).` : ""));
  });

  document.getElementById("importCancelBtn").addEventListener("click", () => {
    importModal.classList.add("hidden");
  });

  // ---------- init ----------

  function renderAll() {
    refreshDatalists();
    renderDashboard();
    renderTable();
  }

  renderAll();

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch((err) => console.error("SW registration failed", err));
    });
  }
})();
