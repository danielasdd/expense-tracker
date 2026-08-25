(() => {
  "use strict";

  const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  const LEGACY_STORAGE_KEY = "expenseTrackerData_v1";
  const LAST_TRACK_KEY = "expenseTrackerLastTrack_v1";

  const PRESET_CATEGORIES = [
    "Acessórios", "Apostas", "Cabeleireiro", "Comida", "Gasolina", "Geral",
    "Ginásio", "Investimentos", "IRS", "Lazer", "Ordenado", "Prendas",
    "Produtos de Beleza", "Restaurantes", "Roupa", "Saídas à Noite", "Saúde",
    "Shein", "Subscrições", "Supermercado", "Tecnologia", "Transportes",
    "Unhas", "Viagens"
  ];
  const PRESET_ACCOUNTS = ["Novo Banco", "Revolut", "Trading 212"];

  const state = {
    session: null,
    tracks: [],
    currentTrackId: null,
    data: [],
    editId: null,
    sort: { key: "date", dir: "desc" },
    search: "",
    categoryFilterValue: "",
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

  function slugify(s) {
    return (s || "export").toString().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "export";
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

  function escapeHtml(s) {
    return (s || "").toString().replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function fromDbRow(row) {
    return {
      id: row.id,
      amount: Number(row.amount),
      category: row.category,
      date: row.date,
      note: row.note || "",
      type: row.type,
      account: row.account || "",
    };
  }

  // ---------- categories / accounts datalists ----------

  function refreshDatalists() {
    const cats = new Set(PRESET_CATEGORIES);
    const accs = new Set(PRESET_ACCOUNTS);
    state.data.forEach((e) => {
      if (e.category) cats.add(e.category);
      if (e.account) accs.add(e.account);
    });
    const sortedCats = [...cats].sort();

    const categorySelectEl = document.getElementById("category");
    const currentCategoryValue = categorySelectEl.value;
    categorySelectEl.innerHTML =
      `<option value="" disabled${currentCategoryValue ? "" : " selected"}>Select a category</option>` +
      sortedCats.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("") +
      `<option value="__new__">+ Add new category…</option>`;
    if (sortedCats.includes(currentCategoryValue)) categorySelectEl.value = currentCategoryValue;

    const accList = document.getElementById("accountList");
    accList.innerHTML = [...accs].sort().map((a) => `<option value="${escapeHtml(a)}">`).join("");

    const categoryFilterEl = document.getElementById("categoryFilter");
    const currentFilterValue = categoryFilterEl.value;
    categoryFilterEl.innerHTML =
      `<option value="">All Categories</option>` +
      sortedCats.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
    categoryFilterEl.value = sortedCats.includes(currentFilterValue) ? currentFilterValue : "";
  }

  // ---------- auth ----------

  const authScreen = document.getElementById("authScreen");
  const appContent = document.getElementById("appContent");
  const authForm = document.getElementById("authForm");
  const authEmail = document.getElementById("authEmail");
  const authPassword = document.getElementById("authPassword");
  const authError = document.getElementById("authError");
  const authSubmitBtn = document.getElementById("authSubmitBtn");
  const authToggleBtn = document.getElementById("authToggleBtn");
  const authToggleText = document.getElementById("authToggleText");
  const authTitle = document.getElementById("authTitle");
  const logoutBtn = document.getElementById("logoutBtn");

  let authMode = "signin";

  function updateAuthUI() {
    if (authMode === "signin") {
      authTitle.textContent = "Sign In";
      authSubmitBtn.textContent = "Sign In";
      authToggleText.textContent = "Don't have an account?";
      authToggleBtn.textContent = "Sign Up";
    } else {
      authTitle.textContent = "Sign Up";
      authSubmitBtn.textContent = "Sign Up";
      authToggleText.textContent = "Already have an account?";
      authToggleBtn.textContent = "Sign In";
    }
    authError.classList.add("hidden");
  }

  authToggleBtn.addEventListener("click", () => {
    authMode = authMode === "signin" ? "signup" : "signin";
    updateAuthUI();
  });

  authForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    authError.classList.add("hidden");
    authSubmitBtn.disabled = true;
    const email = authEmail.value.trim();
    const password = authPassword.value;
    try {
      const { error } = authMode === "signin"
        ? await sb.auth.signInWithPassword({ email, password })
        : await sb.auth.signUp({ email, password });
      if (error) throw error;
    } catch (err) {
      authError.textContent = err.message || "Something went wrong.";
      authError.classList.remove("hidden");
    } finally {
      authSubmitBtn.disabled = false;
    }
  });

  logoutBtn.addEventListener("click", async () => {
    await sb.auth.signOut();
  });

  let initialized = false;

  sb.auth.onAuthStateChange((_event, session) => {
    state.session = session;
    if (session) {
      authScreen.classList.add("hidden");
      appContent.classList.remove("hidden");
      logoutBtn.classList.remove("hidden");
      initAfterLogin();
    } else {
      initialized = false;
      authScreen.classList.remove("hidden");
      appContent.classList.add("hidden");
      logoutBtn.classList.add("hidden");
      authForm.reset();
    }
  });

  async function initAfterLogin() {
    if (initialized) return;
    initialized = true;
    try {
      state.tracks = await loadTracks();
      const generalTrack = state.tracks.find((t) => t.is_general);
      const lastId = localStorage.getItem(LAST_TRACK_KEY);
      state.currentTrackId = (lastId && state.tracks.some((t) => t.id === lastId)) ? lastId : generalTrack.id;
      await maybeMigrateLegacyData(generalTrack.id);
      await refreshCurrentTrackData();
      renderTrackSelect();
      renderAll();
    } catch (err) {
      toast("Could not load your data: " + err.message);
    }
  }

  // ---------- tracks ----------

  async function loadTracks() {
    const { data, error } = await sb.from("tracks").select("*").order("is_general", { ascending: false }).order("name");
    if (error) throw error;
    let tracks = data || [];
    if (!tracks.some((t) => t.is_general)) {
      const { data: created, error: createErr } = await sb.from("tracks")
        .insert({ name: "General", is_general: true, status: "open", user_id: state.session.user.id })
        .select().single();
      if (createErr) throw createErr;
      tracks = [created, ...tracks];
    }
    return tracks;
  }

  async function refreshCurrentTrackData() {
    const requestedTrackId = state.currentTrackId;
    const { data, error } = await sb.from("expenses").select("*").eq("track_id", requestedTrackId).order("date", { ascending: false });
    if (error) throw error;
    // Ignore stale responses: if the user switched tracks again while this
    // request was in flight, a slower earlier response must not clobber data
    // that already belongs to the now-current track.
    if (state.currentTrackId !== requestedTrackId) return;
    state.data = (data || []).map(fromDbRow);
  }

  function currentTrack() {
    return state.tracks.find((t) => t.id === state.currentTrackId);
  }

  function isCurrentTrackClosed() {
    const t = currentTrack();
    return !!t && t.status === "closed";
  }

  const trackSelect = document.getElementById("trackSelect");
  const closeTrackBtn = document.getElementById("closeTrackBtn");
  const reopenTrackBtn = document.getElementById("reopenTrackBtn");

  function renderTrackSelect() {
    trackSelect.innerHTML = state.tracks
      .map((t) => `<option value="${t.id}">${escapeHtml(t.name)}${t.status === "closed" ? " (closed)" : ""}</option>`)
      .join("") + `<option value="__new__">+ New Track…</option>`;
    trackSelect.value = state.currentTrackId;
    updateTrackControls();
  }

  function updateTrackControls() {
    const track = currentTrack();
    const isClosed = !!track && track.status === "closed";
    closeTrackBtn.classList.toggle("hidden", !track || track.is_general || isClosed);
    reopenTrackBtn.classList.toggle("hidden", !track || track.is_general || !isClosed);
    document.body.classList.toggle("track-closed", isClosed);
  }

  trackSelect.addEventListener("change", async () => {
    if (trackSelect.value === "__new__") {
      const name = (prompt('Name this track (e.g. "Paris Expenses"):') || "").trim();
      trackSelect.value = state.currentTrackId;
      if (!name) return;
      try {
        const { data, error } = await sb.from("tracks")
          .insert({ name, is_general: false, status: "open", user_id: state.session.user.id })
          .select().single();
        if (error) throw error;
        state.tracks.push(data);
        await switchTrack(data.id);
        renderTrackSelect();
        toast(`Track "${name}" created.`);
      } catch (err) {
        toast("Could not create track: " + err.message);
      }
      return;
    }
    await switchTrack(trackSelect.value);
  });

  async function switchTrack(id) {
    if (state.editId) exitEditMode();
    state.currentTrackId = id;
    localStorage.setItem(LAST_TRACK_KEY, id);
    try {
      await refreshCurrentTrackData();
    } catch (err) {
      toast("Could not load track data: " + err.message);
    }
    updateTrackControls();
    renderAll();
  }

  closeTrackBtn.addEventListener("click", async () => {
    const track = currentTrack();
    if (!track || track.is_general) return;
    if (!confirm(`Close "${track.name}"? Its net total will be added to General.`)) return;
    try {
      const { data, error } = await sb.rpc("close_track", { p_track_id: track.id });
      if (error) throw error;
      Object.assign(track, data);
      renderTrackSelect();
      toast(`"${track.name}" closed.`);
    } catch (err) {
      toast("Could not close track: " + err.message);
    }
  });

  reopenTrackBtn.addEventListener("click", async () => {
    const track = currentTrack();
    if (!track || track.is_general) return;
    try {
      const { data, error } = await sb.rpc("reopen_track", { p_track_id: track.id });
      if (error) throw error;
      Object.assign(track, data);
      renderTrackSelect();
      toast(`"${track.name}" reopened.`);
    } catch (err) {
      toast("Could not reopen track: " + err.message);
    }
  });

  // ---------- legacy localStorage migration ----------

  async function maybeMigrateLegacyData(generalTrackId) {
    const raw = localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!raw) return;
    let legacy;
    try {
      legacy = JSON.parse(raw);
    } catch {
      return;
    }
    if (!Array.isArray(legacy) || legacy.length === 0) return;

    const { count, error: countErr } = await sb.from("expenses").select("id", { count: "exact", head: true }).eq("track_id", generalTrackId);
    if (countErr || (count && count > 0)) return;

    if (!confirm(`Found ${legacy.length} transaction(s) saved locally in this browser. Import them into your account's General track now?`)) {
      return;
    }
    try {
      const rows = legacy.map((e) => ({
        amount: e.amount, category: e.category, date: e.date,
        note: e.note || "", type: e.type || "expense", account: e.account || "",
      }));
      const { data: importedCount, error } = await sb.rpc("import_legacy_data", { p_rows: rows });
      if (error) throw error;
      localStorage.setItem(LEGACY_STORAGE_KEY + "_imported_" + todayISO(), raw);
      localStorage.removeItem(LEGACY_STORAGE_KEY);
      toast(`Imported ${importedCount} transaction(s) from this browser's local data.`);
    } catch (err) {
      toast("Could not import local data: " + err.message);
    }
  }

  // ---------- form: add / edit ----------

  const form = document.getElementById("expenseForm");
  const formTitle = document.getElementById("formTitle");
  const submitBtn = document.getElementById("submitBtn");
  const cancelEditBtn = document.getElementById("cancelEditBtn");

  document.getElementById("date").value = todayISO();

  const categorySelect = document.getElementById("category");
  let lastCategoryValue = "";

  categorySelect.addEventListener("change", () => {
    if (categorySelect.value === "__new__") {
      const name = (prompt("New category name:") || "").trim();
      if (!name) {
        categorySelect.value = lastCategoryValue;
        return;
      }
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      categorySelect.insertBefore(opt, categorySelect.querySelector('option[value="__new__"]'));
      categorySelect.value = name;
    }
    lastCategoryValue = categorySelect.value;
  });

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (isCurrentTrackClosed()) {
      toast("This track is closed. Reopen it to make changes.");
      return;
    }
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

    submitBtn.disabled = true;
    try {
      if (state.editId) {
        const { data, error } = await sb.from("expenses")
          .update({ amount, category, date, note, type, account })
          .eq("id", state.editId)
          .select().single();
        if (error) throw error;
        const idx = state.data.findIndex((e) => e.id === state.editId);
        if (idx !== -1) state.data[idx] = fromDbRow(data);
        toast("Expense updated.");
      } else {
        const { data, error } = await sb.from("expenses")
          .insert({ amount, category, date, note, type, account, track_id: state.currentTrackId, user_id: state.session.user.id })
          .select().single();
        if (error) throw error;
        state.data.push(fromDbRow(data));
        toast("Expense added.");
      }
      exitEditMode();
      renderAll();
    } catch (err) {
      toast("Could not save: " + err.message);
    } finally {
      submitBtn.disabled = false;
    }
  });

  cancelEditBtn.addEventListener("click", () => exitEditMode());

  function enterEditMode(rec) {
    state.editId = rec.id;
    document.getElementById("expenseId").value = rec.id;
    document.getElementById("amount").value = rec.amount;
    document.getElementById("category").value = rec.category;
    lastCategoryValue = rec.category;
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
    lastCategoryValue = "";
    formTitle.textContent = "Add Expense";
    submitBtn.textContent = "Add Expense";
    cancelEditBtn.classList.add("hidden");
  }

  function deleteExpense(id) {
    if (isCurrentTrackClosed()) {
      toast("This track is closed. Reopen it to make changes.");
      return;
    }
    const index = state.data.findIndex((e) => e.id === id);
    if (index === -1) return;

    const [record] = state.data.splice(index, 1);
    if (state.editId === id) exitEditMode();
    renderAll();

    const pending = { record, index };
    pending.timer = setTimeout(async () => {
      if (state.pendingDelete !== pending) return;
      state.pendingDelete = null;
      const { error } = await sb.from("expenses").delete().eq("id", id);
      if (error) toast("Could not delete on server: " + error.message);
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

  const categoryFilter = document.getElementById("categoryFilter");
  categoryFilter.addEventListener("change", () => {
    state.categoryFilterValue = categoryFilter.value;
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
    let rows = state.data.filter((e) =>
      (!q || normalize(e.note).includes(q) || normalize(e.category).includes(q)) &&
      (!state.categoryFilterValue || e.category === state.categoryFilterValue)
    );

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
    const trackName = currentTrack()?.name || "export";
    const blob = new Blob([JSON.stringify(state.data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `expense-tracker-${slugify(trackName)}-${todayISO()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  });

  const TYPE_LABELS_PT = { expense: "Despesa", income: "Rendimento", investment: "Investimento" };

  document.getElementById("exportExcelBtn").addEventListener("click", () => {
    if (state.data.length === 0) {
      toast("No transactions to export.");
      return;
    }
    const trackName = currentTrack()?.name || "export";
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
    XLSX.writeFile(wb, `expense-tracker-${slugify(trackName)}-${todayISO()}.xlsx`, { cellDates: true });
  });

  document.getElementById("restoreInput").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!Array.isArray(parsed)) throw new Error("Invalid backup file");
        const trackName = currentTrack()?.name || "this track";
        if (!confirm(`Restore ${parsed.length} transaction(s) from this backup? This will REPLACE all entries in the "${trackName}" track (${state.data.length} transactions).`)) return;
        if (isCurrentTrackClosed()) {
          toast("This track is closed. Reopen it to make changes.");
          return;
        }

        const { error: delError } = await sb.from("expenses").delete().eq("track_id", state.currentTrackId);
        if (delError) throw delError;

        if (parsed.length > 0) {
          const rows = parsed.map((r) => ({
            amount: r.amount, category: r.category, date: r.date, note: r.note || "",
            type: r.type || "expense", account: r.account || "",
            track_id: state.currentTrackId, user_id: state.session.user.id,
          }));
          const { data, error: insError } = await sb.from("expenses").insert(rows).select();
          if (insError) throw insError;
          state.data = data.map(fromDbRow);
        } else {
          state.data = [];
        }
        renderAll();
        toast("Backup restored.");
      } catch (err) {
        toast("Could not restore backup: " + err.message);
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

  importConfirmBtn.addEventListener("click", async () => {
    if (isCurrentTrackClosed()) {
      toast("This track is closed. Reopen it to make changes.");
      importModal.classList.add("hidden");
      return;
    }
    const existingKeys = new Set(state.data.map(rowKey));
    const toInsert = [];
    let skipped = 0;
    for (const row of importPendingRows) {
      const key = rowKey(row);
      if (existingKeys.has(key)) { skipped++; continue; }
      existingKeys.add(key);
      toInsert.push({ ...row, track_id: state.currentTrackId, user_id: state.session.user.id });
    }
    importModal.classList.add("hidden");
    if (toInsert.length === 0) {
      toast(`No new transactions to import.` + (skipped > 0 ? ` Skipped ${skipped} duplicate(s).` : ""));
      return;
    }
    try {
      const { data, error } = await sb.from("expenses").insert(toInsert).select();
      if (error) throw error;
      state.data.push(...data.map(fromDbRow));
      renderAll();
      toast(`Imported ${data.length} transaction(s).` + (skipped > 0 ? ` Skipped ${skipped} duplicate(s).` : ""));
    } catch (err) {
      toast("Import failed: " + err.message);
    }
  });

  document.getElementById("importCancelBtn").addEventListener("click", () => {
    importModal.classList.add("hidden");
  });

  // ---------- background refresh ----------

  async function refreshOnFocus() {
    if (!state.session || !state.currentTrackId || state.editId) return;
    try {
      await refreshCurrentTrackData();
      renderAll();
    } catch (err) {
      console.error("Background refresh failed", err);
    }
  }

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refreshOnFocus();
  });

  // ---------- init ----------

  function renderAll() {
    refreshDatalists();
    renderDashboard();
    renderTable();
  }

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch((err) => console.error("SW registration failed", err));
    });
  }
})();
