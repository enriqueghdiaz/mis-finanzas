/* Mis Finanzas — lógica de la interfaz */
(function () {
  const { LS } = Auth;
  const { eur, SERIES } = Charts;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const round2 = v => Math.round(v * 100) / 100;
  const todayISO = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  const MES = Excel.MESES;
  const monthLabel = k => { const [y, m] = k.split("-"); return `${MES[+m - 1]} ${y}`; };
  const monthShort = k => MES[+k.split("-")[1] - 1].slice(0, 3);
  const addMonths = (k, n) => { let [y, m] = k.split("-").map(Number); m += n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return `${y}-${String(m).padStart(2, "0")}`; };

  const S = {
    provider: null, model: null, mode: null, view: "inicio",
    month: todayISO().slice(0, 7), queue: [], flushing: false,
    status: { kind: "", text: "" }, file: null,
    movFilter: { q: "", tipo: "Todos", limit: 60 },
    budgetSeg: "Gasto", patSeg: "Balance", balMonth: null
  };
  const settings = () => LS.get("fin_settings") || {};
  const saveSettings = patch => LS.set("fin_settings", Object.assign(settings(), patch));

  // ── Tema ──────────────────────────────────────────────────
  function applyTheme() {
    const t = settings().theme || "auto";
    if (t === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", t);
  }

  // ── Estado de sincronización ──────────────────────────────
  function setStatus(kind, text) {
    S.status = { kind, text };
    const b = $("#syncBadge");
    b.className = "sync-badge " + kind; b.textContent = text;
  }
  function toast(msg, ms = 2600) {
    const t = $("#toast"); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => t.hidden = true, ms);
  }

  // ── Arranque ──────────────────────────────────────────────
  async function start() {
    applyTheme();
    if (!Excel.Expr || !Excel.noteKey || !Charts.sparkbars || !Charts.diverging) {
      document.getElementById("main").innerHTML = `<div class="alert err" style="margin-top:10px"><span class="i">!</span><div><b>Actualización incompleta.</b><br>
        Faltan archivos de la versión nueva en GitHub. Sube <b>todos</b> los archivos del paquete de actualización (carpeta <code>js</code> completa, <code>style.css</code> y <code>sw.js</code>) y vuelve a abrir la app.</div></div>`;
      return;
    }
    if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});
    bindChrome();
    try { await Auth.handleRedirect(); } catch (e) { toast("Inicio de sesión: " + e.message, 5000); }

    if (new URLSearchParams(location.search).get("nuevo")) { S.pendingNew = true; history.replaceState(null, "", Auth.redirectUri()); }
    const st = settings();
    if (st.mode === "demo") return startDemo();
    if (!Auth.isSignedIn()) return showLogin();
    if (!st.fileId) return showFilePicker();
    S.mode = "graph";
    S.provider = new Excel.GraphProvider(st.fileId, st.driveId);
    S.file = { name: st.fileName, webUrl: st.fileUrl };
    S.queue = LS.get("fin_queue") || [];
    const cache = LS.get("fin_cache");
    if (cache && cache.fileId === st.fileId) {
      try { S.model = Excel.build(cache.raw); S.queue.forEach(op => Excel.applyLocal(S.model, op)); showApp(); } catch (e) { S.model = null; }
    }
    if (!S.model) { showApp(true); }
    await refresh();
  }

  function startDemo() {
    S.mode = "demo"; S.provider = new DemoProvider(); S.queue = [];
    S.file = { name: "Modo demo (datos ficticios)" };
    showApp(true);
    refresh();
  }

  async function refresh() {
    if (!S.provider) return;
    if (S.queue.length) { await flush(); if (S.queue.length) return; }
    setStatus("", "Sincronizando…");
    try {
      const raw = await Excel.loadRaw(S.provider);
      S.model = Excel.build(raw);
      if (S.mode === "graph") LS.set("fin_cache", { fileId: settings().fileId, raw });
      setStatus("ok", S.mode === "demo" ? "Demo" : "Al día · " + new Date().toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" }));
      if (S.model.warnings.length) toast("Aviso: no he podido leer " + S.model.warnings.join(" · "), 6000);
      render();
    } catch (e) { handleError(e); if (!S.model) renderError(e); }
  }

  function handleError(e) {
    console.error(e);
    if (e instanceof Auth.AuthError) { setStatus("err", "Inicia sesión"); return; }
    if (!navigator.onLine || e instanceof TypeError) { setStatus("pending", S.queue.length ? `Sin conexión · ${S.queue.length} pendiente(s)` : "Sin conexión"); return; }
    setStatus("err", "Error de sincronización");
    if (S.model) toast(e.message, 5000);
  }

  // ── Cola de cambios (funciona sin conexión) ───────────────
  function saveQueue() { if (S.mode === "graph") LS.set("fin_queue", S.queue); }
  async function doOps(list, okMsg) {
    if (settings().autoSnap !== false && S.model) {
      const pre = [];
      list.forEach(op => {
        if (op.type !== "setBudgetCells") return;
        monthsOfCells(op.data.cells).forEach(k => { if (!snapsOf(k).length && !pre.some(x => x.data.mes === k)) pre.push(snapOp(k, "Inicial (automática)")); });
      });
      list = pre.concat(list);
    }
    list.forEach(op => { op.t = Date.now(); Excel.applyLocal(S.model, op); S.queue.push(op); });
    saveQueue();
    render();
    if (okMsg) toast(okMsg);
    await flush();
    if (!S.queue.length) await refresh();
  }
  function doOp(type, data, okMsg) { return doOps([{ type, data }], okMsg); }
  async function flush() {
    if (S.flushing || !S.queue.length) return;
    S.flushing = true;
    setStatus("pending", `Guardando ${S.queue.length}…`);
    try {
      while (S.queue.length) {
        const op = S.queue[0];
        try { await Excel.ops[op.type](S.provider, S.model, op.data); }
        catch (e) {
          if (e instanceof Auth.AuthError || e instanceof TypeError || !navigator.onLine || e.status === 429 || e.status >= 500) throw e;
          toast("No se pudo guardar: " + e.message, 6000); // error de datos: se descarta ese cambio
        }
        S.queue.shift(); saveQueue();
      }
    } catch (e) { handleError(e); }
    finally { S.flushing = false; }
  }
  window.addEventListener("online", () => { flush().then(() => { if (!S.queue.length) refresh(); }); });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && S.provider && S.model && Date.now() - (refresh._last || 0) > 60000) { refresh._last = Date.now(); refresh(); }
  });

  // ── Cálculos ──────────────────────────────────────────────
  function txOfMonth(k) { return S.model.tx.list.filter(t => t.fecha.startsWith(k)); }
  function sumBy(list, keyFn) { const o = {}; list.forEach(t => { const k = keyFn(t); o[k] = (o[k] || 0) + t.importe; }); return o; }
  function monthStats(k) {
    const l = txOfMonth(k);
    const ing = l.filter(t => t.tipo === "Ingreso").reduce((a, t) => a + t.importe, 0);
    const gas = l.filter(t => t.tipo === "Gasto").reduce((a, t) => a + t.importe, 0);
    const aho = l.filter(t => t.tipo === "Ahorro").reduce((a, t) => a + t.importe, 0);
    const esen = l.filter(t => t.tipo === "Gasto" && t.nat === "Esencial").reduce((a, t) => a + t.importe, 0);
    const b = S.model.budget;
    const inB = b.months.some(m => m.key === k);
    const bGas = inB ? b.items.filter(i => i.tipo === "Gasto").reduce((a, i) => a + (i.vals[k] || 0), 0) : null;
    const bIng = inB ? b.items.filter(i => i.tipo === "Ingreso").reduce((a, i) => a + (i.vals[k] || 0), 0) : null;
    const bAho = inB && b.savings ? b.savings.vals[k] : null;
    return { ing, gas, aho, esen, pres: gas - esen, margen: ing - gas - aho, bGas, bIng, bAho, n: l.length };
  }
  function allMonths() {
    const set = new Set(S.model.budget.months.map(m => m.key));
    S.model.tx.list.forEach(t => t.fecha && set.add(t.fecha.slice(0, 7)));
    return Array.from(set).sort();
  }
  function txMonths() { return Array.from(new Set(S.model.tx.list.map(t => t.fecha.slice(0, 7)).filter(Boolean))).sort(); }
  function clampMonth(list) { if (!list.includes(S.month)) { const past = list.filter(k => k <= todayISO().slice(0, 7)); S.month = past.length ? past[past.length - 1] : list[0]; } }
  function cleanNote(n) { return (n || "").replace(/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*-?\s*/, "").replace(/\s+-\s+(Edenred|Bizum.*)$/i, "").trim() || "(sin nota)"; }

  // ── Cuentas ──
  const restrictedAccounts = () => Excel.accounts().filter(a => !a.principal);
  function allowedAccounts(tipo, cat) {
    return Excel.accounts().filter(a => a.principal || (tipo === "Gasto" && (a.gastoCats || []).includes(cat)) || (tipo === "Ingreso" && (a.ingresoCats || []).includes(cat)));
  }
  const accCats = a => (a.gastoCats || []).concat(a.ingresoCats || []);
  function accPart(k, cat, acc) { const n = (S.model.notes || {})[Excel.noteKey(k, cat)]; return n && n.cuentas && +n.cuentas[acc] || 0; }
  function realAcc(k, cat, acc) { return txOfMonth(k).filter(t => t.cat === cat && Excel.accountOf(t) === acc).reduce((a, t) => a + t.importe, 0); }

  const CAT_ICON = {
    "Nómina": "💼", "Dietas": "🎫", "Otros ingresos": "➕", "Financiación": "🏦", "Ahorro": "🐷",
    "Vivienda": "🏠", "Préstamos": "💳", "Facturas básicas": "💡", "Transporte": "⛽", "Seguros obligatorios": "🛡️", "Medicamentos - r": "💊",
    "Alimentación": "🛒", "Medicamentos - nr": "💊", "Reparaciones": "🔧", "Servicios médicos": "🩺", "Educación": "📚",
    "Suscripciones": "🔁", "Cuotas": "🏋️", "Clases": "🎓", "Donaciones": "🤝",
    "Restaurantes": "🍽️", "Entretenimiento": "🎬", "Viajes": "✈️", "Ropa": "👕", "Caprichos": "🎁", "Regalos": "🎀", "Otros": "📦"
  };
  const icon = c => CAT_ICON[c] || "•";

  // ── Navegación ────────────────────────────────────────────
  const TITLES = { inicio: "Inicio", movimientos: "Movimientos", presupuesto: "Presupuesto", analisis: "Análisis", patrimonio: "Patrimonio" };
  let chromeBound = false;
  function bindChrome() {
    if (chromeBound) return; chromeBound = true;
    $$(".tab").forEach(b => b.addEventListener("click", () => go(b.dataset.view)));
    $("#fab").addEventListener("click", () => openTxForm());
    $("#btnSettings").addEventListener("click", openSettings);
    $("#syncBadge").addEventListener("click", () => {
      if (S.status.kind === "err" && S.status.text === "Inicia sesión") return Auth.login().catch(e => toast(e.message));
      refresh();
    });
    $$("[data-close]").forEach(b => b.addEventListener("click", closeSheet));
    let resizeT; window.addEventListener("resize", () => { clearTimeout(resizeT); resizeT = setTimeout(() => S.model && render(), 200); });
  }
  function go(view) {
    S.view = view; Charts.hideTip();
    $$(".tab").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    window.scrollTo(0, 0);
    render();
  }
  function showApp(loading) {
    $("#view-login").hidden = true;
    $("#tabbar").hidden = false; $("#fab").hidden = false; $("#btnSettings").hidden = false;
    if (loading && !S.model) {
      $$(".view").forEach(v => v.hidden = true);
      const v = $("#view-inicio"); v.hidden = false; v.innerHTML = `<div class="spinner"></div><p class="empty">Leyendo tu Excel…</p>`;
      return;
    }
    render();
  }
  function renderError(e) {
    $("#viewTitle").textContent = "Mis Finanzas";
    const v = $("#view-" + S.view); $$(".view").forEach(x => x.hidden = true); v.hidden = false;
    v.innerHTML = `<div class="alert err"><span class="i">!</span><div><b>No he podido leer el Excel.</b><br>${esc(e.message)}</div></div>
      <button class="btn" id="retry">Reintentar</button><button class="btn secondary" id="chg">Elegir otro archivo</button>`;
    $("#retry", v).onclick = refresh; $("#chg", v).onclick = showFilePicker;
  }

  function render() {
    if (!S.model) return;
    $("#viewTitle").textContent = TITLES[S.view];
    $$(".view").forEach(v => v.hidden = v.id !== "view-" + S.view);
    const v = $("#view-" + S.view);
    ({ inicio: renderInicio, movimientos: renderMovs, presupuesto: renderBudget, analisis: renderAnalisis, patrimonio: renderPatrimonio })[S.view](v);
    $$(".tab-badge").forEach(x => x.remove());
    if (S.pendingNew) { S.pendingNew = false; setTimeout(() => openTxForm(), 200); }
  }

  function monthSwitch(list, onChange) {
    const i = list.indexOf(S.month);
    const wrap = document.createElement("div"); wrap.className = "month-switch";
    wrap.innerHTML = `<button type="button" aria-label="Mes anterior" ${i <= 0 ? "disabled" : ""}>‹</button><span class="m">${monthLabel(S.month)}</span><button type="button" aria-label="Mes siguiente" ${i >= list.length - 1 ? "disabled" : ""}>›</button>`;
    const [prev, next] = $$("button", wrap);
    prev.onclick = () => { S.month = list[i - 1]; onChange(); };
    next.onclick = () => { S.month = list[i + 1]; onChange(); };
    return wrap;
  }

  // ── INICIO ────────────────────────────────────────────────
  function renderInicio(v) {
    const months = allMonths(); clampMonth(months);
    const k = S.month, st = monthStats(k);
    v.innerHTML = "";
    v.appendChild(monthSwitch(months, render));
    const pct = st.bGas ? Math.min(100, st.gas / st.bGas * 100) : 0;
    const hero = document.createElement("div"); hero.className = "card hero";
    hero.innerHTML = `<div class="label">Gastado en ${MES[+k.split("-")[1] - 1]}</div>
      <div class="big num">${eur(st.gas)}</div>
      ${st.bGas != null ? `<div class="small">${st.gas <= st.bGas ? `Te quedan <b>${eur(st.bGas - st.gas)}</b> de ${eur(st.bGas)} presupuestados` : `Te has pasado <b>${eur(st.gas - st.bGas)}</b> del presupuesto (${eur(st.bGas)})`}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div>` : `<div class="small">Este mes no tiene columna de presupuesto en la hoja PPTO.</div>`}
      <div class="hero-row" style="margin-top:14px"><div><span>Ingresos</span><b class="num">${eur(st.ing)}</b></div><div><span>Margen (tras ahorro)</span><b class="num">${eur(st.margen)}</b></div></div>`;
    v.appendChild(hero);

    const nw = netWorthSeries();
    const last = nw.filter(x => x.v != null);
    const cur = last[last.length - 1], prev = last[last.length - 2];
    const kp = document.createElement("div"); kp.className = "kpis";
    kp.innerHTML = `<div class="kpi"><span>Ahorro del mes</span><b class="num">${eur(st.aho)}</b><small class="muted">${st.bAho != null ? "Plan: " + eur(st.bAho) : "&nbsp;"}</small></div>
      <div class="kpi"><span>Patrimonio neto${cur ? " · " + cur.name.toLowerCase() : ""}</span><b class="num">${cur ? eur(cur.v) : "—"}</b>
      <small class="${cur && prev ? (cur.v >= prev.v ? "pos" : "neg") : "muted"}">${cur && prev ? (cur.v >= prev.v ? "▲ " : "▼ ") + eur(Math.abs(cur.v - prev.v)) + " vs " + prev.name.toLowerCase() : "&nbsp;"}</small></div>`;
    v.appendChild(kp);

    // Contador de lo que queda por asignar
    const cIni = counter(k);
    if (cIni) {
      const row = document.createElement("button"); row.type = "button"; row.className = "counter-pill " + (Math.abs(cIni.libre) < 1 ? "ok" : cIni.libre > 0 ? "free" : "over");
      row.innerHTML = `<span>${Math.abs(cIni.libre) < 1 ? "Presupuesto de " + MES[+k.split("-")[1] - 1] + " cuadrado ✓" : cIni.libre > 0 ? "Por asignar tras el ahorro" : "Asignado de más"}</span><b class="num">${Math.abs(cIni.libre) < 1 ? "" : eur(Math.abs(cIni.libre))} ›</b>`;
      row.onclick = () => { S.budgetSeg = "Gasto"; go("presupuesto"); };
      v.appendChild(row);
    }

    // Alertas: categorías por encima de presupuesto
    const over = budgetRows(k, "Gasto").filter(r => r.budget != null && r.real > r.budget + 0.005).sort((a, b) => (b.real - b.budget) - (a.real - a.budget));
    if (over.length) {
      const al = document.createElement("div"); al.className = "alert";
      al.innerHTML = `<span class="i">!</span><div><b>Revisar:</b> ${over.slice(0, 4).map(r => `${esc(r.cat)} (+${eur(r.real - r.budget)})`).join(", ")}${over.length > 4 ? ` y ${over.length - 4} más` : ""}</div>`;
      al.style.cursor = "pointer"; al.onclick = () => { S.budgetSeg = "Gasto"; go("presupuesto"); };
      v.appendChild(al);
    }

    // Gráfico mensual
    const c1 = document.createElement("div"); c1.className = "card";
    c1.innerHTML = `<div class="card-head"><h3>Gasto por mes</h3><button class="link" type="button">Ver análisis</button></div><div class="ch"></div>`;
    $(".link", c1).onclick = () => go("analisis");
    v.appendChild(c1);
    const tm = txMonths().filter(m => m <= todayISO().slice(0, 7)).slice(-6);
    requestAnimationFrame(() => Charts.stackedBars($(".ch", c1), tm.map(monthShort), [
      { name: "Esenciales", color: SERIES[0], values: tm.map(m => monthStats(m).esen) },
      { name: "Prescindibles", color: SERIES[1], values: tm.map(m => monthStats(m).pres) }
    ], { target: tm.map(m => monthStats(m).bGas), fullLabels: tm.map(monthLabel) }));

    // Últimos movimientos
    const c2 = document.createElement("div"); c2.className = "card";
    c2.innerHTML = `<div class="card-head"><h3>Últimos movimientos</h3><button class="link" type="button">Ver todos</button></div><div class="list"></div>`;
    $(".link", c2).onclick = () => go("movimientos");
    const list = $(".list", c2);
    S.model.tx.list.slice(0, 6).forEach(t => list.appendChild(txRow(t)));
    if (!S.model.tx.list.length) list.innerHTML = `<div class="empty">Aún no hay movimientos.</div>`;
    v.appendChild(c2);
  }

  function txRow(t) {
    const b = document.createElement("button"); b.type = "button"; b.className = "row" + (t.pending ? " pending" : "");
    const out = t.tipo === "Ahorro" && t.importe < 0;
    const sign = t.tipo === "Ingreso" ? "+" : t.tipo === "Ahorro" ? (out ? "← " : "→ ") : "−";
    const cls = t.tipo === "Ingreso" || out ? "pos" : "";
    const d = t.fecha ? new Date(t.fecha + "T12:00").toLocaleDateString("es-ES", { day: "numeric", month: "short" }) : "";
    b.innerHTML = `<span class="ico">${icon(t.cat)}</span><span class="mid"><div class="t">${esc(t.notas || t.cat)}</div><div class="s">${esc(t.cat)} · ${d}${Excel.accountOf(t) !== Excel.mainAccount() ? ` · <span class="acc-tag">${esc(Excel.accountOf(t))}</span>` : (t.nat && t.nat !== "NA" && t.tipo === "Gasto" ? " · " + esc(t.nat) : "")}</div></span><span class="amt num ${cls}">${sign}${eur(Math.abs(t.importe))}</span>`;
    b.onclick = () => openTxForm(t);
    return b;
  }

  // ── MOVIMIENTOS ───────────────────────────────────────────
  function goMovs(filter) {
    Object.assign(S.movFilter, { q: "", tipo: "Todos", cat: "", month: "", acc: "", limit: 60 }, filter || {});
    go("movimientos");
  }
  function renderMovs(v) {
    const F = S.movFilter;
    F.cat = F.cat || ""; F.month = F.month || ""; F.sort = F.sort || "fecha"; F.acc = F.acc || "";
    const months = txMonths().slice().reverse();
    const catOpts = Excel.CATS.map(g => `<optgroup label="${g.tipo === "Gasto" ? esc(g.nat + " · " + g.per) : "Ingresos " + g.per.toLowerCase() + "s"}">${g.cats.map(c => `<option value="${esc(c)}" ${F.cat === c ? "selected" : ""}>${esc(c)}</option>`).join("")}</optgroup>`).join("")
      + `<optgroup label="Ahorro"><option value="Ahorro" ${F.cat === "Ahorro" ? "selected" : ""}>Ahorro</option></optgroup>`;
    v.innerHTML = `<input class="search" type="search" placeholder="Buscar concepto, categoría o importe" value="${esc(F.q)}" aria-label="Buscar">
      <div class="chips">${["Todos", "Gasto", "Ingreso", "Ahorro"].map(x => `<button type="button" class="chip ${F.tipo === x ? "on" : ""}" data-t="${x}">${x === "Todos" ? "Todos" : x === "Gasto" ? "Gastos" : x === "Ingreso" ? "Ingresos" : "Ahorro"}</button>`).join("")}</div>
      <div class="filters">
        <select id="fCat" aria-label="Categoría"><option value="">Todas las categorías</option>${catOpts}</select>
        <select id="fMes" aria-label="Mes"><option value="">Todos los meses</option>${months.map(m => `<option value="${m}" ${F.month === m ? "selected" : ""}>${monthLabel(m)}</option>`).join("")}</select>
        <select id="fAcc" aria-label="Cuenta"><option value="">Todas las cuentas</option>${Excel.accounts().map(a => `<option value="${esc(a.name)}" ${F.acc === a.name ? "selected" : ""}>${esc(a.name)}</option>`).join("")}</select>
        <select id="fSort" aria-label="Ordenar"><option value="fecha" ${F.sort === "fecha" ? "selected" : ""}>Más recientes</option><option value="importe" ${F.sort === "importe" ? "selected" : ""}>Mayor importe</option><option value="concepto" ${F.sort === "concepto" ? "selected" : ""}>Concepto A-Z</option></select>
      </div>
      <div id="movSum" class="mov-sum"></div>
      <div class="card" style="padding-top:4px"><div class="list" id="movList"></div></div>`;
    const inp = $(".search", v);
    inp.oninput = () => { F.q = inp.value; F.limit = 60; drawList(); };
    $$(".chip", v).forEach(c => c.onclick = () => { F.tipo = c.dataset.t; F.limit = 60; $$(".chip", v).forEach(x => x.classList.toggle("on", x === c)); drawList(); });
    $("#fCat", v).onchange = e => { F.cat = e.target.value; F.limit = 60; drawList(); };
    $("#fMes", v).onchange = e => { F.month = e.target.value; F.limit = 60; drawList(); };
    $("#fSort", v).onchange = e => { F.sort = e.target.value; F.limit = 60; drawList(); };
    $("#fAcc", v).onchange = e => { F.acc = e.target.value; F.limit = 60; drawList(); };
    function matches(t, q) {
      if (!q) return true;
      const hay = (t.notas + " " + t.cat).toLowerCase();
      const r = Expr.tryEval(q);
      if (r.ok && /^[\d\s.,]+$/.test(q) && Math.abs(t.importe - r.value) < 0.005) return true;
      return q.split(/\s+/).every(w => hay.includes(w));
    }
    function drawList() {
      const q = F.q.trim().toLowerCase();
      let l = S.model.tx.list.filter(t => (F.tipo === "Todos" || t.tipo === F.tipo) && (!F.cat || t.cat === F.cat) && (!F.month || t.fecha.startsWith(F.month)) && (!F.acc || Excel.accountOf(t) === F.acc) && matches(t, q));
      if (F.sort === "importe") l = l.slice().sort((a, b) => b.importe - a.importe);
      else if (F.sort === "concepto") l = l.slice().sort((a, b) => cleanNote(a.notas).localeCompare(cleanNote(b.notas), "es"));
      else l = l.slice().sort((a, b) => b.fecha.localeCompare(a.fecha));
      const g = l.filter(x => x.tipo === "Gasto").reduce((a, x) => a + x.importe, 0);
      const i = l.filter(x => x.tipo === "Ingreso").reduce((a, x) => a + x.importe, 0);
      const ah = l.filter(x => x.tipo === "Ahorro").reduce((a, x) => a + x.importe, 0);
      const active = q || F.tipo !== "Todos" || F.cat || F.month || F.acc;
      $("#movSum", v).innerHTML = `<span><b>${l.length}</b> movimiento${l.length === 1 ? "" : "s"}${g ? ` · gastos <b class="num">${eur(g)}</b>` : ""}${i ? ` · ingresos <b class="num pos">${eur(i)}</b>` : ""}${ah ? ` · ahorro <b class="num">${eur(ah)}</b>` : ""}</span>
        ${active ? `<button type="button" class="link" id="fClear">Quitar filtros</button>` : ""}`;
      if ($("#fClear", v)) $("#fClear", v).onclick = () => { Object.assign(F, { q: "", tipo: "Todos", cat: "", month: "", acc: "", limit: 60 }); renderMovs(v); };
      const box = $("#movList", v); box.innerHTML = "";
      if (!l.length) { box.innerHTML = `<div class="empty">No hay movimientos con ese filtro.</div>`; return; }
      let cur = null;
      l.slice(0, F.limit).forEach(t => {
        const m = t.fecha.slice(0, 7);
        if (F.sort === "fecha" && m !== cur) {
          cur = m;
          const tot = l.filter(x => x.fecha.startsWith(m));
          const gg = tot.filter(x => x.tipo === "Gasto").reduce((a, x) => a + x.importe, 0);
          const ii = tot.filter(x => x.tipo === "Ingreso").reduce((a, x) => a + x.importe, 0);
          const h = document.createElement("div"); h.className = "group-h";
          h.innerHTML = `<span>${monthLabel(m)}</span><span class="num">${ii ? "+" + eur(ii) + " · " : ""}${gg ? "−" + eur(gg) : ""}</span>`;
          box.appendChild(h);
        }
        box.appendChild(txRow(t));
      });
      if (l.length > F.limit) {
        const more = document.createElement("button"); more.className = "btn secondary"; more.style.marginTop = "10px";
        more.textContent = `Ver más (${l.length - F.limit})`; more.onclick = () => { F.limit += 100; drawList(); };
        box.appendChild(more);
      }
    }
    drawList();
  }

  // ── Formulario de movimiento ─────────────────────────────
  function openTxForm(t) {
    const editing = !!t;
    const f = t ? Object.assign({}, t) : { tipo: "Gasto", cat: "", importe: "", fecha: todayISO(), notas: "" };
    f.retiro = f.tipo === "Ahorro" && +f.importe < 0;
    f.cuenta = t ? Excel.accountOf(t) : "";
    if (f.importe !== "") f.importe = Math.abs(+f.importe);
    const body = document.createElement("div");
    const draw = () => {
      const groups = f.tipo === "Ingreso" ? Excel.CATS.filter(g => g.tipo === "Ingreso") : f.tipo === "Gasto" ? Excel.CATS.filter(g => g.tipo === "Gasto") : [];
      body.innerHTML = `
        <div class="seg" style="margin-bottom:14px">${["Gasto", "Ingreso", "Ahorro"].map(x => `<button type="button" data-t="${x}" class="${f.tipo === x ? "on" : ""}">${x}</button>`).join("")}</div>
        <div class="field"><input id="fImp" class="amount-input num" type="text" inputmode="decimal" placeholder="0,00 €" value="${f.importe !== "" ? String(f.importe).replace(".", ",") : ""}" aria-label="Importe"></div>
        ${groups.length ? `<div class="field"><label>Categoría</label>${groups.map(g => `<div class="cat-sec">${g.tipo === "Gasto" ? g.nat + " · " + g.per : g.per}</div><div class="cat-grid">${g.cats.map(c => `<button type="button" class="chip ${f.cat === c ? "on" : ""}" data-c="${esc(c)}">${icon(c)} ${esc(c)}</button>`).join("")}</div>`).join("")}</div>` : `<div class="seg" style="margin-bottom:8px"><button type="button" data-r="0" class="${!f.retiro ? "on" : ""}">→ Aportar</button><button type="button" data-r="1" class="${f.retiro ? "on" : ""}">← Retirar</button></div>
          <p class="small muted" style="margin-top:0">${f.retiro ? "Se registra en negativo en la categoría Ahorro (el dinero vuelve a estar disponible). Para que sume a «Por asignar», baja también el ahorro previsto en Presupuesto → Ahorro → «Retirar del ahorro»." : "Se registra como <b>Gasto · Esencial · Fijo · Ahorro</b>, igual que en tu Excel."}</p>`}
        ${(() => {
          const opts = f.cat ? allowedAccounts(f.tipo, f.cat) : [];
          if (opts.length < 2) return "";
          const lastAcc = (settings().lastAcc || {})[f.cat];
          if (!f.cuenta || !opts.some(a => a.name === f.cuenta)) f.cuenta = opts.some(a => a.name === lastAcc) ? lastAcc : Excel.mainAccount();
          return `<div class="field"><label>Pagado con</label><div class="chips">${opts.map(a => `<button type="button" class="chip acc ${f.cuenta === a.name ? "on" : ""}" data-acc="${esc(a.name)}">${a.principal ? "🏦" : "🎫"} ${esc(a.name)}</button>`).join("")}</div></div>`;
        })()}
        <div class="field"><label for="fFecha">Fecha (mes en el que cuenta)</label><input id="fFecha" type="date" value="${esc(f.fecha)}"></div>
        <div class="field"><label for="fNotas">Notas</label><input id="fNotas" type="text" list="noteSug" value="${esc(f.notas)}" placeholder="Ej.: Compra semanal"><datalist id="noteSug"></datalist></div>
        <button class="btn" id="fSave" type="button">${editing ? "Guardar cambios" : "Añadir"}</button>
        ${editing ? `<div style="height:10px"></div><button class="btn danger" id="fDel" type="button">Eliminar</button>` : ""}`;
      $$(".seg button[data-r]", body).forEach(b => b.onclick = () => { keep(); f.retiro = b.dataset.r === "1"; draw(); });
      $$(".seg button[data-t]", body).forEach(b => b.onclick = () => { keep(); f.tipo = b.dataset.t; if (f.tipo === "Ahorro") f.cat = "Ahorro"; else if (Excel.CAT_INFO[f.cat] && Excel.CAT_INFO[f.cat].tipo !== f.tipo) f.cat = ""; draw(); });
      $$(".cat-grid .chip", body).forEach(b => b.onclick = () => { keep(); f.cat = b.dataset.c; draw(); });
      $$(".chip.acc", body).forEach(b => b.onclick = () => { keep(); f.cuenta = b.dataset.acc; draw(); });
      const sug = $("#noteSug", body);
      const notes = {}; S.model.tx.list.filter(x => !f.cat || x.cat === f.cat).forEach(x => { const n = x.notas; if (n) notes[n] = (notes[n] || 0) + 1; });
      Object.entries(notes).sort((a, b) => b[1] - a[1]).slice(0, 12).forEach(([n]) => { const o = document.createElement("option"); o.value = n; sug.appendChild(o); });
      $("#fSave", body).onclick = save;
      if (editing) $("#fDel", body).onclick = async () => {
        if (!confirm("¿Eliminar este movimiento también del Excel?")) return;
        closeSheet(); await doOp("deleteTx", { old: strip(t) }, "Movimiento eliminado");
      };
      if (!editing && !f.importe) setTimeout(() => $("#fImp", body).focus(), 250);
    };
    const keep = () => {
      const imp = $("#fImp", body); if (imp) f.importe = imp.value;
      const fe = $("#fFecha", body); if (fe) f.fecha = fe.value;
      const no = $("#fNotas", body); if (no) f.notas = no.value;
    };
    const strip = x => ({ tipo: x.tipo, tipoRaw: x.tipoRaw, nat: x.nat, per: x.per, cat: x.cat, fecha: x.fecha, importe: x.importe, notas: x.notas });
    async function save() {
      keep();
      const imp = parseFloat(String(f.importe).replace(/\./g, "").replace(",", ".")) || parseFloat(String(f.importe));
      if (!imp || imp <= 0) return toast("Escribe un importe");
      if (f.tipo === "Ahorro") f.cat = "Ahorro";
      if (!f.cat) return toast("Elige una categoría");
      if (!f.fecha) return toast("Elige una fecha");
      const info = f.tipo === "Ahorro" ? Excel.SAVING : Excel.CAT_INFO[f.cat];
      const accOk = allowedAccounts(f.tipo, f.cat).some(a => a.name === f.cuenta);
      const cuenta = accOk ? f.cuenta : Excel.mainAccount();
      if (allowedAccounts(f.tipo, f.cat).length > 1) saveSettings({ lastAcc: Object.assign({}, settings().lastAcc, { [f.cat]: cuenta }) });
      const tx = { tipo: f.tipo, nat: info.nat, per: info.per, cat: f.cat, fecha: f.fecha, importe: round2(f.tipo === "Ahorro" && f.retiro ? -imp : imp), notas: (f.notas || "").trim(), cuenta };
      // ¿Este ingreso supera lo previsto para su categoría ese mes?
      let extraPart = 0;
      if (!editing && tx.tipo === "Ingreso" && settings().askSavings !== false) {
        const mk = tx.fecha.slice(0, 7), bud = budgetOf(tx.cat, mk);
        if (bud != null) { const before = realCat(mk, tx.cat); extraPart = Math.max(0, before + tx.importe - bud) - Math.max(0, before - bud); }
      }
      closeSheet();
      if (editing) await doOp("updateTx", { old: strip(t), tx }, "Cambios guardados");
      else {
        const pending = doOp("addTx", { tx }, "Movimiento añadido");
        if (extraPart > 0.009) setTimeout(() => askSavings({ extra: round2(extraPart), fecha: tx.fecha, origen: tx.notas || tx.cat }), 350);
        await pending;
      }
    }
    draw();
    openSheet(editing ? "Editar movimiento" : "Nuevo movimiento", body);
  }

  // ── PRESUPUESTO ───────────────────────────────────────────
  const Expr = Excel.Expr;
  const nowMonth = () => todayISO().slice(0, 7);
  const parseAmount = s => { const r = Expr.tryEval(s); return r.ok ? r.value : NaN; };
  function realCat(k, cat) { return txOfMonth(k).filter(t => t.cat === cat).reduce((a, t) => a + t.importe, 0); }
  function budgetOf(cat, k) {
    const b = S.model.budget;
    if (cat === "Ahorro") return b.savings && b.months.some(m => m.key === k) ? (b.savings.vals[k] || 0) : null;
    const it = b.items.find(i => i.cat === cat);
    return it && b.months.some(m => m.key === k) ? (it.vals[k] || 0) : null;
  }
  // La fila «Ahorro» de PPTO tratada como una partida más
  function savingsItem() {
    const sv = S.model.budget.savings;
    return sv ? { cat: "Ahorro", row: sv.row, tipo: "Ahorro", nat: "", per: "", vals: sv.vals, f: sv.f } : null;
  }
  function isRateFormula(f) {
    const b = S.model.budget;
    return !!(b.rate && typeof f === "string" && f[0] === "=" && f.replace(/\$/g, "").includes(b.rate.addr));
  }
  // Fórmula «ingresos × tasa» para una columna, copiando la que ya usa tu Excel en otro mes
  function rateFormulaFor(col) {
    const b = S.model.budget, sv = b.savings; if (!sv || !b.rate) return null;
    const src = b.months.find(m => isRateFormula(sv.f[m.key]));
    if (!src) return null;
    return sv.f[src.key].replace(new RegExp(`(^|[^$A-Z])${src.col}(?=\\$?\\d)`, "g"), `$1${col}`);
  }
  function plannedIncome(k) { return S.model.budget.items.filter(i => i.tipo === "Ingreso").reduce((a, i) => a + (i.vals[k] || 0), 0); }
  // Construye la celda (número o fórmula) a partir de las líneas del desglose
  function linesToCell(lines) {
    const used = lines.map(l => ({ e: (l.e || "").trim(), l: (l.l || "").trim() })).filter(l => l.e || l.l);
    const parts = used.filter(l => l.e).map(l => (used.length > 1 && /[*/]/.test(l.e) ? `(${l.e})` : l.e));
    const exprAll = parts.join(" + ").replace(/\+\s*-\s*/g, "- ");
    return { used, cell: parts.length ? Expr.toCell(exprAll) : 0, desglose: (used.length > 1 || used.some(l => l.l)) ? used : [] };
  }
  // Meses de referencia: los completos (o con datos) anteriores al mes que se presupuesta
  function refMonths(k, n) { const lim = k < nowMonth() ? k : nowMonth(); return txMonths().filter(m => m < lim).slice(-(n || 3)); }
  function catStats(cat, k) {
    const ref = refMonths(k, 3), vals = ref.map(m => realCat(m, cat));
    const avg = ref.length ? vals.reduce((a, x) => a + x, 0) / ref.length : null;
    const bIdx = S.model.budget.months.findIndex(m => m.key === k);
    const prevBudgetMonth = bIdx > 0 ? S.model.budget.months[bIdx - 1].key : null;
    const withB = ref.filter(m => budgetOf(cat, m) != null);
    const six = refMonths(k, 6);
    return {
      ref, vals, avg,
      last: ref.length ? vals[vals.length - 1] : null,
      max: vals.length ? Math.max(...vals) : null,
      prevBudgetMonth, prevBudget: prevBudgetMonth ? budgetOf(cat, prevBudgetMonth) : null,
      overs: withB.filter(m => realCat(m, cat) > budgetOf(cat, m) * 1.05 + 5).length,
      nRefB: withB.length,
      dev: withB.length ? withB.reduce((a, m) => a + realCat(m, cat) - budgetOf(cat, m), 0) / withB.length : null,
      devInit: withB.length ? withB.reduce((a, m) => a + realCat(m, cat) - initialBudget(cat, m), 0) / withB.length : null,
      six, sixVals: six.map(m => realCat(m, cat))
    };
  }
  function noteOf(k, cat) { return (S.model.notes || {})[Excel.noteKey(k, cat)] || null; }
  function hasNote(n) { return n && (n.nota || (n.desglose && n.desglose.length)); }

  // Contador: lo que queda por asignar tras el ahorro (se mueve con los ingresos reales)
  function counter(k, over) {
    const b = S.model.budget;
    if (!b.months.some(m => m.key === k)) return null;
    const val = it => (over && over[it.row] != null) ? over[it.row] : (it.vals[k] || 0);
    const inc = b.items.filter(i => i.tipo === "Ingreso");
    const realInc = sumBy(txOfMonth(k).filter(t => t.tipo === "Ingreso"), t => t.cat);
    const ingPlan = inc.reduce((a, i) => a + val(i), 0);
    let ingCons = inc.reduce((a, i) => a + Math.max(val(i), realInc[i.cat] || 0), 0);
    Object.keys(realInc).forEach(c => { if (!inc.some(i => i.cat === c)) ingCons += realInc[c]; });
    const extra = Math.max(0, ingCons - ingPlan);
    const rate = b.rate ? b.rate.value : 0;
    const sv = b.savings, sf = sv && sv.f ? sv.f[k] : null;
    const byRate = !!(b.rate && typeof sf === "string" && sf.replace(/\$/g, "").includes(b.rate.addr));
    const ahoPlan = sv ? (over && over[sv.row] != null ? over[sv.row] : (byRate && over ? ingPlan * rate : (sv.vals[k] || 0))) : 0;
    // El ahorro extra solo cuenta si lo apartas tú (movimiento de Ahorro con nota «Ahorro extra…»)
    const ahoTx = txOfMonth(k).filter(t => t.tipo === "Ahorro");
    const ahoReal = ahoTx.reduce((a, t) => a + t.importe, 0);
    const ahoExtra = ahoTx.filter(t => /^ahorro extra/i.test(t.notas || "")).reduce((a, t) => a + t.importe, 0);
    const aho = Math.max(ahoPlan, ahoReal - ahoExtra) + ahoExtra;
    const gasPres = b.items.filter(i => i.tipo === "Gasto").reduce((a, i) => a + val(i), 0);
    const gasReal = txOfMonth(k).filter(t => t.tipo === "Gasto").reduce((a, t) => a + t.importe, 0);
    // Lo que queda por gastar de lo presupuestado (partida a partida; lo que te pasas no resta de otras)
    let pend = 0, exceso = 0, nPend = 0;
    b.items.filter(i => i.tipo === "Gasto").forEach(i => {
      const d = val(i) - realCat(k, i.cat);
      if (d > 0.004) { pend += d; nPend++; } else if (d < -0.004) exceso -= d;
    });
    // Cuentas restringidas (Edenred): su dinero solo vale para sus categorías
    let ingR = 0, gasR = 0, pendR = 0;
    const accs = restrictedAccounts().map(a => {
      const ingP = inc.reduce((x, i) => x + Math.min(accPart(k, i.cat, a.name), val(i)), 0);
      const ingRe = txOfMonth(k).filter(t => t.tipo === "Ingreso" && Excel.accountOf(t) === a.name).reduce((x, t) => x + t.importe, 0);
      const ingC = Math.max(ingP, ingRe);
      const gasP = b.items.filter(i => i.tipo === "Gasto").reduce((x, i) => x + Math.min(accPart(k, i.cat, a.name), val(i)), 0);
      const pendA = b.items.filter(i => i.tipo === "Gasto").reduce((x, i) => x + Math.max(0, Math.min(accPart(k, i.cat, a.name), val(i)) - realAcc(k, i.cat, a.name)), 0);
      ingR += ingC; gasR += gasP; pendR += pendA;
      return { name: a.name, cats: (a.gastoCats || []).join(", "), ingC, gasP, pend: pendA, libre: ingC - gasP, used: ingC > 0.004 || gasP > 0.004 };
    }).filter(a => a.used);
    return { accs, pendMain: pend - pendR, pend, exceso, nPend, ingPlan, ingCons, extra, rate, byRate, ahoPlan, ahoExtra, ahoReal, aho, gasPres, gasReal, libre: (ingCons - ingR) - aho - (gasPres - gasR) };
  }
  function counterHTML(c, k, compact) {
    if (!c) return "";
    const L = c.libre, done = Math.abs(L) < 1;
    const head = done ? "Todo asignado ✓" : L > 0 ? "Por asignar" : "Asignado de más";
    const pct = x => c.ingCons > 0 ? Math.max(0, Math.min(100, x / c.ingCons * 100)) : 0;
    return `<div class="counter ${done ? "ok" : L > 0 ? "free" : "over"}">
      <div class="counter-top"><div><div class="label">${head} · ${esc(MES[+k.split("-")[1] - 1])}</div>
        <div class="big num">${done ? eur(0) : (L > 0 ? "" : "−") + eur(Math.abs(L))}</div>
        ${c.accs.map(a => `<div class="acc-free ${a.libre < -0.5 ? "neg" : ""}">🎫 ${esc(a.name)}: ${a.libre >= 0 ? eur(a.libre) + " por asignar" : eur(-a.libre) + " de más"} <small>(solo ${esc(a.cats)})</small></div>`).join("")}</div>
        ${compact ? "" : `<div class="pend"><div class="label">Pendiente de gastar</div><div class="pend-num num">${eur(c.pend)}</div>
          ${c.accs.some(a => a.pend > 0.004) ? `<div class="xsmall muted">de ello ${c.accs.filter(a => a.pend > 0.004).map(a => `${eur(a.pend)} ${esc(a.name)}`).join(", ")}</div>` : ""}
          ${c.exceso > 0.004 ? `<div class="xsmall neg">${eur(c.exceso)} de exceso</div>` : `<div class="xsmall muted">en ${c.nPend} partida${c.nPend === 1 ? "" : "s"}</div>`}</div>`}</div>
      <div class="stackbar"><span style="width:${pct(c.aho)}%;background:var(--s1)"></span><span style="width:${pct(c.gasPres)}%;background:var(--s2)"></span></div>
      <div class="counter-lines">
        <div><span>Ingresos${c.extra > 0.004 ? ` <small class="pos">(+${eur(c.extra)} extra)</small>` : ""}${c.accs.length ? ` <small>(sin ${c.accs.map(a => esc(a.name)).join(", ")})</small>` : ""}</span><b class="num">${eur(c.ingCons - c.accs.reduce((x, a) => x + a.ingC, 0))}</b></div>
        <div><span><i class="dot" style="background:var(--s1)"></i>Ahorro${c.ahoExtra > 0.004 ? ` <small>(+${eur(c.ahoExtra)} apartado del extra)</small>` : ""}</span><b class="num">−${eur(c.aho)}</b></div>
        <div><span><i class="dot" style="background:var(--s2)"></i>Gastos presupuestados${c.accs.length ? ` <small>(sin ${c.accs.map(a => esc(a.name)).join(", ")})</small>` : ""}</span><b class="num">−${eur(c.gasPres - c.accs.reduce((x, a) => x + a.gasP, 0))}</b></div>
        ${c.accs.map(a => `<div class="acc-line"><span>🎫 ${esc(a.name)}: entra ${eur(a.ingC)} · presupuestado ${eur(a.gasP)}</span><b class="num">${a.libre >= 0 ? "" : "−"}${eur(Math.abs(a.libre))}</b></div>`).join("")}
      </div>
      ${!compact && c.extra - c.ahoExtra > 0.5 ? `<button type="button" class="link counter-save" data-extra="${round2(c.extra - c.ahoExtra)}">Tienes ${eur(c.extra - c.ahoExtra)} de ingresos extra. ¿Apartar una parte para ahorro?</button>` : ""}
      </div>`;
  }

  // Pregunta si apartar ahorro de un ingreso no previsto (nunca se hace solo)
  function askSavings(o) {
    const rate = S.model.budget.rate ? S.model.budget.rate.value : 0;
    const sug = rate ? Math.round(o.extra * rate) : 0;
    const body = document.createElement("div");
    body.innerHTML = `<p class="small" style="margin-top:0">Has recibido <b>${eur(o.extra)}</b> más de lo previsto (${esc(o.origen)}). Si no apartas nada, queda <b>disponible para asignar</b> en el presupuesto.</p>
      ${sug ? `<button class="btn" id="aRate" type="button">Apartar ${Math.round(rate * 100)}% · ${eur(sug)}</button><div style="height:10px"></div>` : ""}
      <div class="field" style="margin-bottom:8px"><label for="aAmt">Otra cantidad</label><input id="aAmt" class="num" inputmode="decimal" placeholder="0,00"></div>
      <button class="btn secondary" id="aOther" type="button">Apartar esta cantidad</button><div style="height:10px"></div>
      <button class="btn secondary" id="aNo" type="button">No, dejarlo para asignar</button>
      <label class="small" style="display:flex;gap:8px;align-items:center;margin-top:14px"><input type="checkbox" id="aNever"> No volver a preguntar (se puede cambiar en Ajustes)</label>
      <p class="xsmall muted">Se registra como un movimiento de Ahorro con la nota «Ahorro extra», que el contador descuenta.</p>`;
    const done = async amt => {
      if ($("#aNever", body).checked) saveSettings({ askSavings: false });
      closeSheet();
      if (!(amt > 0)) return;
      const tx = { tipo: "Ahorro", nat: Excel.SAVING.nat, per: Excel.SAVING.per, cat: "Ahorro", fecha: o.fecha, importe: round2(amt), notas: "Ahorro extra · " + o.origen };
      await doOp("addTx", { tx }, `Apartados ${eur(amt)} para ahorro`);
    };
    if ($("#aRate", body)) $("#aRate", body).onclick = () => done(sug);
    $("#aOther", body).onclick = () => { const a = parseAmount($("#aAmt", body).value); if (!(a > 0)) return toast("Escribe una cantidad"); done(a); };
    $("#aNo", body).onclick = () => done(0);
    openSheet("¿Apartas algo para ahorro?", body);
  }
  function budgetRows(k, tipo) {
    const b = S.model.budget, inB = b.months.some(m => m.key === k);
    const real = sumBy(txOfMonth(k).filter(t => t.tipo === tipo), t => t.cat);
    return b.items.filter(i => i.tipo === tipo).map(i => ({ cat: i.cat, nat: i.nat, per: i.per, row: i.row, budget: inB ? (i.vals[k] || 0) : null, real: real[i.cat] || 0 }));
  }
  function renderBudget(v) {
    const b = S.model.budget;
    const months = b.months.map(m => m.key);
    if (!months.length) { v.innerHTML = `<div class="empty">No encuentro columnas de meses en la hoja PPTO.</div>`; return; }
    if (!months.includes(S.month)) { const p = months.filter(x => x <= nowMonth()); S.month = p.length ? p[p.length - 1] : months[0]; }
    const k = S.month, tipo = S.budgetSeg;
    v.innerHTML = "";
    v.appendChild(monthSwitch(months, render));
    const cObj = counter(k);
    if (cObj) {
      const cnt = document.createElement("div"); cnt.className = "card counter-card" + (settings().counterOpen ? " open" : "");
      cnt.innerHTML = counterHTML(cObj, k) + `<button type="button" class="counter-toggle" aria-label="Ver detalle">${settings().counterOpen ? "Ocultar detalle ▴" : "Ver detalle ▾"}</button>`;
      $(".counter-toggle", cnt).onclick = () => { saveSettings({ counterOpen: !settings().counterOpen }); render(); };
      const sv = $(".counter-save", cnt);
      if (sv) sv.onclick = () => askSavings({ extra: +sv.dataset.extra, fecha: k === nowMonth() ? todayISO() : k + "-01", origen: "ingresos extra de " + MES[+k.split("-")[1] - 1] });
      v.appendChild(cnt);
    }
    const seg = document.createElement("div"); seg.className = "seg";
    seg.innerHTML = [["Gasto", "Gastos"], ["Ingreso", "Ingresos"], ["Plan", "Planificar"]].map(([s, l]) => `<button type="button" class="${tipo === s ? "on" : ""}" data-s="${s}">${l}</button>`).join("");
    $$("button", seg).forEach(x => x.onclick = () => { S.budgetSeg = x.dataset.s; render(); });
    v.appendChild(seg);
    if (tipo === "Plan") return renderPlan(v, k);

    const rows = budgetRows(k, tipo);
    const tb = rows.reduce((a, r) => a + r.budget, 0), tr = rows.reduce((a, r) => a + r.real, 0);
    const sum = document.createElement("div"); sum.className = "kpis";
    sum.innerHTML = `<div class="kpi"><span>Presupuestado</span><b class="num">${eur(tb)}</b></div>
      <div class="kpi"><span>${tipo === "Gasto" ? "Gastado" : "Recibido"}</span><b class="num">${eur(tr)}</b>
      <small class="${tipo === "Gasto" ? (tr > tb ? "neg" : "pos") : (tr >= tb ? "pos" : "muted")}">${tipo === "Gasto" ? (cObj ? `Pendiente ${eur(cObj.pend)}${cObj.exceso > 0.004 ? ` <span class="neg">· exceso ${eur(cObj.exceso)}</span>` : ""}` : (tr > tb ? "Exceso " + eur(tr - tb) : "Quedan " + eur(tb - tr))) : (tr >= tb ? "Objetivo cumplido" : "Faltan " + eur(tb - tr))}</small></div>`;
    v.appendChild(sum);

    const card = document.createElement("div"); card.className = "card";
    const svIt = savingsItem();
    let savHTML = "";
    if (svIt) {
      const plan = svIt.vals[k] || 0, realS = realCat(k, "Ahorro"), ns = noteOf(k, "Ahorro");
      const p = plan > 0 ? Math.max(0, realS) / plan : 0;
      const sub = hasNote(ns) ? (ns.nota || ns.desglose.map(x => x.l || x.e).filter(Boolean).join(" · ")) : (isRateFormula(svIt.f[k]) ? `${Math.round(S.model.budget.rate.value * 100)}% de tus ingresos previstos` : "");
      savHTML = `<div class="bgroup-h"><h4>Ahorro</h4><span class="xsmall muted">toca para cambiarlo o retirar</span></div>
        <div class="brow" data-row="${svIt.row}" style="cursor:pointer"><div class="brow-top"><span class="name">${icon("Ahorro")} Ahorro previsto ${hasNote(ns) ? `<span class="note-ico">📝</span>` : ""}</span><span class="vals num">${eur(realS)} / <b>${eur(plan)}</b></span></div>
        ${sub ? `<div class="brow-sub">${esc(sub)}</div>` : ""}
        <div class="prog sav"><div style="width:${Math.min(100, p * 100)}%"></div></div></div>`;
    }
    const groups = tipo === "Gasto" ? [["Esencial", "Fijo"], ["Esencial", "Variable"], ["Prescindible", "Fijo"], ["Prescindible", "Variable"]] : [["NA", "Fijo"], ["NA", "Variable"]];
    let html = "";
    groups.forEach(([nat, per]) => {
      const gr = rows.filter(r => r.nat === nat && r.per === per);
      if (!gr.length) return;
      const gb = gr.reduce((a, r) => a + r.budget, 0), greal = gr.reduce((a, r) => a + r.real, 0);
      html += `<div class="bgroup-h"><h4>${tipo === "Gasto" ? nat + " · " + per : "Ingresos " + per.toLowerCase() + "s"}</h4><span class="xsmall muted num">${eur(greal)} / ${eur(gb)}</span></div>`;
      gr.forEach(r => {
        const p = r.budget ? r.real / r.budget : (r.real ? 2 : 0);
        const cls = tipo === "Gasto" ? (p > 1.0001 ? "over" : p > 0.85 ? "near" : "") : "";
        const tag = tipo === "Gasto" && r.real > r.budget + 0.005 ? `<span class="tag over">Revisar</span>` : "";
        const n = noteOf(k, r.cat);
        const it = b.items.find(i => i.row === r.row);
        const ex = Expr.fromFormula(it.f[k]);
        let sub = hasNote(n) ? (n.nota || n.desglose.map(x => x.l || x.e).filter(Boolean).join(" · ")) : (ex && !Expr.isPlainNumber(ex) ? ex : "");
        const parts = restrictedAccounts().filter(a => accPart(k, r.cat, a.name) > 0);
        if (parts.length) sub = parts.map(a => `🎫 ${a.name} ${eur(accPart(k, r.cat, a.name))} (gastado ${eur(realAcc(k, r.cat, a.name))})`).join(" · ") + ` · ${Excel.mainAccount()} ${eur(Math.max(0, (r.budget || 0) - parts.reduce((x, a) => x + accPart(k, r.cat, a.name), 0)))}` + (sub ? " · " + sub : "");
        html += `<div class="brow" data-row="${r.row}" style="cursor:pointer"><div class="brow-top"><span class="name">${icon(r.cat)} ${esc(r.cat)} ${hasNote(n) ? `<span class="note-ico" title="Tiene notas">📝</span>` : ""} ${tag}</span><span class="vals num">${eur(r.real)} / <b>${eur(r.budget)}</b></span></div>
          ${sub ? `<div class="brow-sub">${esc(sub)}</div>` : ""}
          <div class="prog ${cls}"><div style="width:${Math.min(100, p * 100)}%"></div></div></div>`;
      });
    });
    card.innerHTML = savHTML + html;
    $$(".brow", card).forEach(el => el.onclick = () => editBudgetItem(+el.dataset.row, k));
    v.appendChild(card);

    const actions = document.createElement("div"); actions.className = "btn-row";
    actions.innerHTML = `<button class="btn secondary" type="button" id="bAll">Editar mes completo</button><button class="btn secondary" type="button" id="bSnap">📸 Congelar</button>`;
    $("#bAll", actions).onclick = () => editBudgetMonth(k);
    $("#bSnap", actions).onclick = () => freezeBudget(k);
    const sn = snapsOf(k);
    const hist = document.createElement("button"); hist.type = "button"; hist.className = "counter-pill";
    hist.innerHTML = sn.length ? `<span>📸 ${sn.length} foto${sn.length === 1 ? "" : "s"} este mes · última ${esc(snapDate(sn[sn.length - 1].id))} (${esc(sn[sn.length - 1].label)})</span><b>›</b>` : `<span>📸 Aún no has congelado el presupuesto de ${esc(MES[+k.split("-")[1] - 1])}</span><b>›</b>`;
    hist.onclick = () => snapHistory(k);
    v.appendChild(actions);
    v.appendChild(hist);
    const note = document.createElement("p"); note.className = "xsmall muted"; note.style.margin = "0 4px";
    note.textContent = `Toca una partida para editarla: admite sumas y restas (p. ej. 8,25 + 144,09 − 30) y un concepto por línea. Se guarda en la hoja PPTO (columna ${b.months.find(m => m.key === k).col}) y las notas en la hoja «App_Notas».`;
    v.appendChild(note);
  }

  // Barra de operadores para el teclado numérico del móvil
  function opBar(container, getTarget, extra) {
    const bar = document.createElement("div"); bar.className = "opbar";
    bar.innerHTML = (extra || "") + ["+", "−", "(", ")"].map(o => `<button type="button" class="opk" data-o="${o}">${o}</button>`).join("");
    $$(".opk[data-o]", bar).forEach(bt => {
      bt.addEventListener("pointerdown", e => e.preventDefault());
      bt.addEventListener("click", () => {
        const inp = getTarget(); if (!inp) return;
        const o = bt.dataset.o === "−" ? "-" : bt.dataset.o;
        const ins = /[+\-]/.test(o) ? ` ${o} ` : o;
        const a = inp.selectionStart != null ? inp.selectionStart : inp.value.length, z = inp.selectionEnd != null ? inp.selectionEnd : a;
        inp.value = inp.value.slice(0, a) + ins + inp.value.slice(z);
        const pos = a + ins.length; try { inp.setSelectionRange(pos, pos); } catch (e) {}
        inp.focus(); inp.dispatchEvent(new Event("input"));
      });
    });
    container.appendChild(bar);
    return bar;
  }

  function editBudgetItem(row, k) {
    const b = S.model.budget;
    const it = (b.savings && row === b.savings.row) ? savingsItem() : b.items.find(i => i.row === row);
    const isSav = it.tipo === "Ahorro";
    const m = b.months.find(x => x.key === k);
    const st = catStats(it.cat, k);
    const real = realCat(k, it.cat);
    const nTx = txOfMonth(k).filter(t => t.cat === it.cat).length;
    const later = b.months.filter(x => x.key > k);
    const note = noteOf(k, it.cat);
    const cur = it.vals[k] || 0;
    // Líneas del desglose: notas guardadas → fórmula de la celda → importe
    let lines = null, mismatch = false;
    const ex = Expr.fromFormula(it.f[k]);
    if (note && note.desglose.length) {
      const sum = note.desglose.reduce((a, x) => a + (parseAmount(x.e) || 0), 0);
      if (Math.abs(sum - cur) < 0.01) lines = note.desglose.map(x => ({ e: x.e || "", l: x.l || "" }));
      else mismatch = true;
    }
    if (!lines && ex) {
      const terms = Expr.splitTerms(ex);
      lines = terms.map((t, i) => ({ e: t, l: note && note.desglose.length === terms.length ? (note.desglose[i].l || "") : "" }));
    }
    if (!lines) lines = [{ e: cur ? Expr.fmt(cur) : "", l: "" }];
    let lastFocus = null;
    const origLines = lines.map(l => Object.assign({}, l));
    // Cuentas restringidas que aplican a esta partida (p. ej. Edenred en Restaurantes y Dietas)
    const accOpts = isSav ? [] : restrictedAccounts().filter(a => it.tipo === "Gasto" ? (a.gastoCats || []).includes(it.cat) : (a.ingresoCats || []).includes(it.cat));
    const sugAcc = a => {
      if (it.tipo === "Gasto") { // lo que entra ese mes en la cuenta (partida de ingreso con reparto)
        const inc = b.items.filter(i => i.tipo === "Ingreso").reduce((x, i) => x + accPart(k, i.cat, a.name), 0);
        return inc > 0 ? round2(inc) : 0;
      }
      const prevK = b.months.filter(x => x.key < k).map(x => x.key).pop();
      return prevK && accPart(prevK, it.cat, a.name) ? accPart(prevK, it.cat, a.name) : 0;
    };
    // Ahorro: ¿se calcula con tu tasa? (=ingresos × B17)
    const rate = b.rate ? b.rate.value : 0;
    const canRate = isSav && rate && !!rateFormulaFor(m.col);
    let rateMode = isSav && isRateFormula(it.f[k]);

    const body = document.createElement("div");
    body.innerHTML = `
      <div class="mini-kpis">
        <div><span>Real ${esc(monthShort(k))}</span><b class="num">${eur(real)}</b></div>
        <div><span>Media ${st.ref.length || ""} meses</span><b class="num">${st.avg != null ? eur(st.avg) : "—"}</b></div>
        <div><span>Presup. ${st.prevBudgetMonth ? esc(monthShort(st.prevBudgetMonth)) : "anterior"}</span><b class="num">${st.prevBudget != null ? eur(st.prevBudget) : "—"}</b></div>
      </div>
      <div id="bHist"></div>
      ${mismatch ? `<div class="alert" style="margin-bottom:10px"><span class="i">!</span><div>El importe del Excel (${eur(cur)}) no coincide con el desglose que guardaste. Revisa las líneas.</div></div>` : ""}
      <div class="chips" id="quick" style="margin:10px 0">
        ${st.avg != null ? `<button type="button" class="chip" data-v="${Math.round(st.avg)}">Media: ${eur(Math.round(st.avg))}</button>` : ""}
        ${st.last != null ? `<button type="button" class="chip" data-v="${round2(st.last)}">Real ${esc(monthShort(st.ref[st.ref.length - 1]))}: ${eur(st.last)}</button>` : ""}
        ${st.prevBudget != null ? `<button type="button" class="chip" data-v="${round2(st.prevBudget)}">Como ${esc(monthShort(st.prevBudgetMonth))}: ${eur(st.prevBudget)}</button>` : ""}
        ${st.max != null && st.max > (st.avg || 0) + 1 ? `<button type="button" class="chip" data-v="${round2(st.max)}">Máximo: ${eur(st.max)}</button>` : ""}
        ${canRate ? `<button type="button" class="chip" id="rateChip">Tu tasa: ${Math.round(rate * 100)}% de ingresos = ${eur(plannedIncome(k) * rate)}</button>` : ""}
      </div>
      ${isSav && real > 0.004 ? `<div class="alert" style="margin-bottom:8px"><span class="i">i</span><div>En ${esc(MES[+k.split("-")[1] - 1])} ya has apartado <b>${eur(real)}</b>. Bajar el previsto por debajo de eso no libera dinero: para usarlo, utiliza <b>«Retirar del ahorro»</b> (abajo).</div></div>` : ""}
      ${isSav ? `<p class="xsmall muted" id="rateInfo" style="margin:0 0 8px">${rateMode ? `Ahora se calcula solo: <b>${Math.round(rate * 100)}% de tus ingresos previstos</b>. Si escribes otro importe, ese mes deja de calcularse con la tasa.` : "Importe fijo para este mes."}</p>` : ""}
      <div class="field" style="margin-bottom:6px"><label>Importe y desglose <span class="muted">(puedes sumar y restar)</span></label></div>
      <div id="lines"></div>
      <div id="ops"></div>
      <div class="total-row"><span>Total partida</span><b class="num" id="tot"></b></div>
      ${accOpts.map(a => `<div class="acc-box"><label for="acc_${esc(a.name)}">🎫 De ello con <b>${esc(a.name)}</b> <span class="muted xsmall">(${it.tipo === "Ingreso" ? "lo que entra en " + esc(a.name) : "solo se puede gastar en " + esc((a.gastoCats || []).join(", ")) + "; no se reasigna a otras partidas"})</span></label>
        <div class="acc-row"><input id="acc_${esc(a.name)}" class="acc-in num" data-acc="${esc(a.name)}" inputmode="decimal" value="${accPart(k, it.cat, a.name) ? Expr.fmt(accPart(k, it.cat, a.name)) : ""}" placeholder="0">
        ${sugAcc(a) ? `<button type="button" class="chip acc-sug" data-acc="${esc(a.name)}" data-v="${sugAcc(a)}">${it.tipo === "Gasto" ? "Lo que entra" : "Como antes"}: ${eur(sugAcc(a))}</button>` : ""}</div>
        <div class="xsmall muted acc-rest" data-acc="${esc(a.name)}"></div></div>`).join("")}
      <div class="field"><label for="bNote">Recordatorio / comentario</label><textarea id="bNote" rows="2" placeholder="Ej.: incluye seguro moto 2º pago">${esc(note ? note.nota : "")}</textarea></div>
      ${later.length ? `<label class="small" style="display:flex;gap:8px;align-items:center;margin-bottom:14px"><input type="checkbox" id="bLater"> Aplicar también a ${later.map(x => monthShort(x.key)).join(", ")}</label>` : ""}
      <button class="btn" id="bSave" type="button">Guardar</button>
      ${nTx ? `<div style="height:10px"></div><button class="btn secondary" id="bMovs" type="button">Ver los ${nTx} movimientos de ${esc(monthShort(k))}</button>` : ""}
      ${isSav ? `<div class="withdraw">
        <h4>Retirar del ahorro para gastar</h4>
        <p class="xsmall muted" style="margin:0 0 10px">Resta la cantidad del ahorro previsto de ${esc(MES[+k.split("-")[1] - 1])}, así sube tu «Por asignar» y puedes repartirla entre las partidas.</p>
        <div class="w-grid"><input id="wAmt" class="num" inputmode="decimal" placeholder="Importe"><input id="wWhy" placeholder="Motivo (p. ej. reparación moto)"></div>
        <label class="small" style="display:flex;gap:8px;align-items:center;margin:10px 0"><input type="checkbox" id="wTx" checked> Registrar también la retirada como movimiento (si ya habías traspasado el dinero)</label>
        <button class="btn secondary" id="wGo" type="button">Retirar del ahorro</button></div>` : ""}`;

    const box = $("#lines", body);
    const drawLines = () => {
      box.innerHTML = lines.map((l, i) => `
        <div class="dline" data-i="${i}">
          <div class="dline-top"><input class="expr num" inputmode="decimal" value="${esc(l.e)}" placeholder="0,00" aria-label="Importe línea ${i + 1}">
            <span class="eq num"></span>
            ${lines.length > 1 ? `<button type="button" class="icon-btn del" aria-label="Quitar línea">✕</button>` : ""}</div>
          <input class="lbl" value="${esc(l.l)}" placeholder="Concepto (opcional)" aria-label="Concepto línea ${i + 1}">
        </div>`).join("");
      $$(".dline", box).forEach(d => {
        const i = +d.dataset.i, e = $(".expr", d), lb = $(".lbl", d);
        e.addEventListener("focus", () => lastFocus = e);
        e.addEventListener("input", () => { lines[i].e = e.value; if (rateMode) { rateMode = false; const ri = $("#rateInfo", body); if (ri) ri.textContent = "Importe fijo para este mes."; } recalc(); });
        lb.addEventListener("input", () => lines[i].l = lb.value);
        const del = $(".del", d); if (del) del.onclick = () => { lines.splice(i, 1); drawLines(); };
      });
      recalc();
    };
    const recalc = () => {
      let tot = 0, bad = false;
      $$(".dline", box).forEach(d => {
        const e = $(".expr", d), eq = $(".eq", d), r = Expr.tryEval(e.value);
        if (!e.value.trim()) { eq.textContent = ""; return; }
        if (r.ok) { tot += r.value; eq.textContent = Expr.isPlainNumber(e.value) ? "" : "= " + eur(r.value); eq.classList.remove("neg"); }
        else { bad = true; eq.textContent = "¿?"; eq.classList.add("neg"); }
      });
      $("#tot", body).textContent = bad ? "revisa" : eur(round2(tot));
      $$(".acc-rest", body).forEach(el => {
        const inp = $(`.acc-in[data-acc="${el.dataset.acc}"]`, body), v2 = parseAmount(inp.value || "0");
        el.textContent = isNaN(v2) ? "¿?" : (v2 > tot + 0.004 ? `Es más que el total de la partida (${eur(tot)}).` : `Resto con ${Excel.mainAccount()}: ${eur(round2(tot - v2))}`);
      });
    };
    opBar($("#ops", body), () => lastFocus || $(".expr", box), `<button type="button" class="opk wide" id="addLine">＋ Línea</button>`);
    $("#addLine", body).onclick = () => { lines.push({ e: "", l: "" }); drawLines(); const all = $$(".expr", box); all[all.length - 1].focus(); };
    $$("#quick .chip[data-v]", body).forEach(c => c.onclick = () => { lines = [{ e: Expr.fmt(+c.dataset.v), l: "" }]; rateMode = false; const ri = $("#rateInfo", body); if (ri) ri.textContent = "Importe fijo para este mes."; drawLines(); });
    $$(".acc-in", body).forEach(inp => { inp.addEventListener("focus", () => lastFocus = inp); inp.addEventListener("input", recalc); });
    $$(".acc-sug", body).forEach(c => c.onclick = () => { $(`.acc-in[data-acc="${c.dataset.acc}"]`, body).value = Expr.fmt(+c.dataset.v); recalc(); });
    if ($("#rateChip", body)) $("#rateChip", body).onclick = () => {
      lines = [{ e: Expr.fmt(round2(plannedIncome(k) * rate)), l: "" }]; drawLines(); rateMode = true;
      $("#rateInfo", body).innerHTML = `Se calculará solo: <b>${Math.round(rate * 100)}% de tus ingresos previstos</b> (como en tu Excel).`;
    };
    drawLines();
    requestAnimationFrame(() => {
      const ms = st.six.concat(k);
      Charts.stackedBars($("#bHist", body), ms.map(monthShort), [{ name: "Real", color: SERIES[0], values: ms.map(x => realCat(x, it.cat)) }],
        { height: 130, target: ms.map(x => budgetOf(it.cat, x)), targetName: isSav ? "Previsto" : "Presupuesto", fullLabels: ms.map(monthLabel) });
    });
    if ($("#bMovs", body)) $("#bMovs", body).onclick = () => { closeSheet(); goMovs({ cat: it.cat, month: k }); };

    $("#bSave", body).onclick = async () => {
      for (const l of lines) if ((l.e || "").trim() && !Expr.tryEval(l.e).ok) return toast("Revisa el importe: " + l.e);
      const { cell, desglose } = linesToCell(lines);
      const notaTxt = $("#bNote", body).value.trim();
      const targets = [m].concat($("#bLater", body) && $("#bLater", body).checked ? later : []);
      const cells = targets.map(x => ({ addr: x.col + row, value: (rateMode && canRate) ? rateFormulaFor(x.col) : cell }));
      const cuentas = {};
      for (const inp of $$(".acc-in", body)) {
        if (!inp.value.trim()) continue;
        const v2 = parseAmount(inp.value);
        if (isNaN(v2)) return toast("Revisa el importe de " + inp.dataset.acc);
        if (v2 > 0) cuentas[inp.dataset.acc] = round2(v2);
      }
      const hadCu = x => { const n0 = noteOf(x.key, it.cat); return n0 && n0.cuentas && Object.keys(n0.cuentas).length; };
      const notes = targets.filter(x => notaTxt || desglose.length || hasNote(noteOf(x.key, it.cat)) || Object.keys(cuentas).length || hadCu(x))
        .map(x => ({ mes: x.key, cat: it.cat, nota: notaTxt, desglose, cuentas: accOpts.length ? cuentas : undefined }));
      closeSheet();
      const list = [{ type: "setBudgetCells", data: { cells } }];
      if (notes.length) list.push({ type: "setNotes", data: { notes } });
      await doOps(list, isSav ? "Ahorro previsto actualizado" : "Partida guardada");
    };
    if ($("#wGo", body)) $("#wGo", body).onclick = async () => {
      const amt = parseAmount($("#wAmt", body).value);
      if (!(amt > 0)) return toast("Escribe cuánto quieres retirar");
      if (amt > cur + 0.004 && !confirm(`Vas a retirar ${eur(amt)} y el ahorro previsto de este mes es ${eur(cur)}. Quedará en negativo. ¿Seguir?`)) return;
      const why = $("#wWhy", body).value.trim();
      const base = origLines.filter(l => (l.e || "").trim());
      const wl = base.concat([{ e: "-" + Expr.fmt(amt), l: "Retirada" + (why ? ": " + why : "") }]);
      const { cell, desglose } = linesToCell(wl);
      const prevNota = note ? note.nota : "";
      const nota = (prevNota ? prevNota + " · " : "") + `Retirados ${eur(amt)}${why ? " (" + why + ")" : ""}`;
      const list = [
        { type: "setBudgetCells", data: { cells: [{ addr: m.col + row, value: cell }] } },
        { type: "setNotes", data: { notes: [{ mes: k, cat: "Ahorro", nota, desglose }] } }
      ];
      if ($("#wTx", body).checked) list.push({ type: "addTx", data: { tx: { tipo: "Ahorro", nat: Excel.SAVING.nat, per: Excel.SAVING.per, cat: "Ahorro", fecha: k === nowMonth() ? todayISO() : k + "-01", importe: -round2(amt), notas: "Retirada de ahorro" + (why ? " · " + why : "") } } });
      closeSheet();
      await doOps(list, `Retirados ${eur(amt)} del ahorro`);
    };
    openSheet(`${icon(it.cat)} ${it.cat} · ${monthShort(k)}`, body);
  }

  function editBudgetMonth(k, base) {
    const b = S.model.budget, m = b.months.find(x => x.key === k);
    const idx = b.months.indexOf(m), prevM = b.months[idx - 1];
    const stats = {}; b.items.forEach(i => stats[i.row] = catStats(i.cat, k));
    const lastRef = refMonths(k, 1)[0];
    let lastFocus = null;
    const body = document.createElement("div");
    const startVal = i => {
      const own = Expr.fromFormula(i.f[k]) || Expr.fmt(i.vals[k] || 0);
      // Las partidas con desglose (conceptos) se respetan al rellenar con medias o reales
      if ((base === "avg" || base === "last") && hasNote(noteOf(k, i.cat)) && noteOf(k, i.cat).desglose.length) return own;
      if (base === "prev" && prevM) return Expr.fromFormula(i.f[prevM.key]) || Expr.fmt(i.vals[prevM.key] || 0);
      if (base === "avg" && i.tipo === "Gasto" && stats[i.row].avg != null) return Expr.fmt(Math.round(stats[i.row].avg));
      if (base === "last" && i.tipo === "Gasto" && lastRef) return Expr.fmt(round2(realCat(lastRef, i.cat)));
      return Expr.fromFormula(i.f[k]) || Expr.fmt(i.vals[k] || 0);
    };
    body.innerHTML = `<div id="mCounter"></div>
      <div class="small muted" style="margin:10px 0 6px">Rellenar con:</div>
      <div class="chips" style="margin-bottom:10px">
        <button type="button" class="chip ${!base ? "on" : ""}" data-b="">Lo actual</button>
        ${prevM ? `<button type="button" class="chip ${base === "prev" ? "on" : ""}" data-b="prev">Presupuesto de ${esc(monthShort(prevM.key))}</button>` : ""}
        <button type="button" class="chip ${base === "avg" ? "on" : ""}" data-b="avg">Media real (3 meses)</button>
        ${lastRef ? `<button type="button" class="chip ${base === "last" ? "on" : ""}" data-b="last">Real de ${esc(monthShort(lastRef))}</button>` : ""}
      </div>
      ${["Ingreso", "Ahorro", "Gasto"].map(tp => tp === "Ahorro" ? (b.savings ? `<div class="bgroup-h"><h4>Ahorro</h4></div>
        <div class="brow"><div class="brow-top" style="align-items:center"><span class="name small">${icon("Ahorro")} Ahorro previsto
          <div class="xsmall muted">${isRateFormula(b.savings.f[k]) ? `${Math.round(b.rate.value * 100)}% de ingresos (se recalcula solo)` : "importe fijo"}</div></span>
          <span style="text-align:right"><input class="budget-edit num" data-row="${b.savings.row}" data-sav="1" inputmode="decimal" value="${esc(Expr.fromFormula(b.savings.f[k]) || Expr.fmt(b.savings.vals[k] || 0))}"><div class="eq xsmall num"></div></span></div></div>` : "") :
        `<div class="bgroup-h"><h4>${tp === "Gasto" ? "Gastos" : "Ingresos"}</h4></div>` + b.items.filter(i => i.tipo === tp).map(i => {
        const s2 = stats[i.row];
        return `<div class="brow"><div class="brow-top" style="align-items:center"><span class="name small">${icon(i.cat)} ${esc(i.cat)}
          ${s2.avg != null ? `<div class="xsmall muted">media ${eur(s2.avg)}${s2.prevBudget != null ? ` · ant. ${eur(s2.prevBudget)}` : ""}</div>` : ""}</span>
          <span style="text-align:right"><input class="budget-edit num" data-row="${i.row}" inputmode="decimal" value="${esc(startVal(i))}"><div class="eq xsmall num"></div></span></div></div>`;
      }).join("")).join("")}
      <div id="ops" style="position:sticky;bottom:0;background:var(--surface);padding:8px 0"></div>
      <button class="btn" id="save" type="button">Guardar ${monthLabel(k)}</button>`;
    const inputs = $$(".budget-edit", body);
    const initial = new Map(inputs.map(inp => [inp, inp.value]));
    const recalc = () => {
      const over = {}; let bad = false;
      inputs.forEach(inp => {
        const r = Expr.tryEval(inp.value), eq = inp.parentElement.querySelector(".eq");
        if (r.ok) { if (!inp.dataset.sav || inp.value !== initial.get(inp)) over[+inp.dataset.row] = r.value; eq.textContent = Expr.isPlainNumber(inp.value) || !inp.value.trim() ? "" : "= " + eur(r.value); eq.classList.remove("neg"); }
        else { bad = true; eq.textContent = "¿?"; eq.classList.add("neg"); }
      });
      $("#mCounter", body).innerHTML = counterHTML(counter(k, over), k, true);
      return bad;
    };
    inputs.forEach(inp => { inp.addEventListener("focus", () => lastFocus = inp); inp.addEventListener("input", recalc); });
    opBar($("#ops", body), () => lastFocus);
    $$(".chips .chip", body).forEach(c => c.onclick = () => { closeSheet(); setTimeout(() => editBudgetMonth(k, c.dataset.b || undefined), 50); });
    recalc();
    $("#save", body).onclick = async () => {
      if (recalc()) return toast("Hay importes que no se entienden (¿?)");
      const cells = [];
      inputs.forEach(inp => {
        const row = +inp.dataset.row;
        if (inp.dataset.sav) { if (inp.value !== initial.get(inp)) cells.push({ addr: m.col + row, value: inp.value.trim() ? Expr.toCell(inp.value) : 0 }); return; }
        const it = b.items.find(i => i.row === row);
        const cell = inp.value.trim() ? Expr.toCell(inp.value) : 0;
        const oldCell = typeof it.f[k] === "string" && it.f[k][0] === "=" ? it.f[k].replace(/\s+/g, "").replace(/^=\+/, "=") : round2(it.vals[k] || 0);
        const same = typeof cell === "number" ? (typeof oldCell === "number" && Math.abs(cell - oldCell) < 0.004) : cell === oldCell;
        if (!same) cells.push({ addr: m.col + row, value: cell });
      });
      closeSheet();
      if (!cells.length) return toast("Sin cambios");
      await doOps([{ type: "setBudgetCells", data: { cells } }], `Guardadas ${cells.length} partidas`);
    };
    openSheet(`Presupuesto · ${monthLabel(k)}`, body);
  }

  // ── Fotos del presupuesto (congelar) ──────────────────────
  function snapsOf(k) { return (S.model.snaps || []).filter(x => x.mes === k).sort((a, b) => a.id.localeCompare(b.id)); }
  function nowStamp() { const d = new Date(); const p2 = n => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`; }
  function snapDate(id) { const [d, t] = id.split(" "); const dt = new Date(d + "T12:00"); return dt.toLocaleDateString("es-ES", { day: "numeric", month: "short" }) + (t ? " " + t.slice(0, 5) : ""); }
  // Estado actual del presupuesto de un mes, en el mismo formato que una foto
  function currentRows(k) {
    const b = S.model.budget, rows = [];
    b.items.forEach(i => {
      const cu = {}; restrictedAccounts().forEach(a => { const v = accPart(k, i.cat, a.name); if (v) cu[a.name] = v; });
      rows.push({ tipo: i.tipo, cat: i.cat, v: round2(i.vals[k] || 0), cuentas: cu });
    });
    if (b.savings) rows.push({ tipo: "Ahorro", cat: "Ahorro", v: round2(b.savings.vals[k] || 0), cuentas: {} });
    return rows;
  }
  function rowsToItems(rows) { const o = {}; rows.forEach(r => o[r.cat] = { tipo: r.tipo, v: r.v, cuentas: r.cuentas }); return o; }
  function snapTotals(items) {
    let ing = 0, gas = 0, aho = 0;
    Object.values(items).forEach(x => { if (x.tipo === "Ingreso") ing += x.v; else if (x.tipo === "Gasto") gas += x.v; else if (x.tipo === "Ahorro") aho += x.v; });
    return { ing, gas, aho, libre: ing - aho - gas };
  }
  function snapDiff(a, b) { // a → b
    const out = [];
    Object.keys(b).filter(c => a[c]).forEach(c => { const x = a[c].v, y = b[c].v; if (Math.abs(y - x) > 0.004) out.push({ cat: c, from: x, to: y, d: round2(y - x), tipo: (b[c] || a[c]).tipo }); });
    return out.sort((p, q) => Math.abs(q.d) - Math.abs(p.d));
  }
  function snapOp(k, label) { return { type: "addSnapshot", data: { id: `${nowStamp()} ${k}#${Math.random().toString(36).slice(2, 6)}`, mes: k, label, rows: currentRows(k) } }; }
  function monthsOfCells(cells) {
    const b = S.model.budget, set = new Set();
    cells.forEach(c => { const col = Excel.parseAddr(c.addr).c1; const m = b.months.find(x => Excel.colNumber(x.col) === col); if (m) set.add(m.key); });
    return Array.from(set);
  }
  function diffChips(diff, max) {
    return diff.slice(0, max || 99).map(x => `<span class="dchip ${x.d > 0 ? "up" : "down"}">${x.d > 0 ? "▲" : "▼"} ${esc(x.cat)} ${x.d > 0 ? "+" : "−"}${eur(Math.abs(x.d))}</span>`).join("") + (diff.length > (max || 99) ? `<span class="dchip">+${diff.length - max} más</span>` : "");
  }

  function freezeBudget(k) {
    const snaps = snapsOf(k), last = snaps[snaps.length - 1];
    const cur = rowsToItems(currentRows(k));
    const diff = last ? snapDiff(last.items, cur) : [];
    const t = snapTotals(cur);
    const def = snaps.length ? `Reajuste ${new Date().toLocaleDateString("es-ES", { day: "numeric", month: "short" })}` : "Inicial";
    const body = document.createElement("div");
    body.innerHTML = `<p class="small" style="margin-top:0">Guarda una foto del presupuesto de <b>${esc(monthLabel(k))}</b> tal como está ahora. Así podrás ver después qué partidas reorganizas y comparar lo gastado con lo que tenías previsto al principio.</p>
      <div class="mini-kpis"><div><span>Ingresos</span><b class="num">${eur(t.ing)}</b></div><div><span>Ahorro</span><b class="num">${eur(t.aho)}</b></div><div><span>Gastos</span><b class="num">${eur(t.gas)}</b></div></div>
      ${last ? `<div class="small" style="margin:10px 0 6px">${diff.length ? `Cambios desde la última foto (${esc(snapDate(last.id))}):` : `Sin cambios desde la última foto (${esc(snapDate(last.id))}).`}</div><div class="dchips">${diffChips(diff, 12)}</div>` : `<p class="xsmall muted">Es la primera foto de este mes.</p>`}
      <div class="field" style="margin-top:12px"><label for="sLabel">Nombre de la foto</label><input id="sLabel" value="${esc(def)}"></div>
      <button class="btn" id="sGo" type="button">📸 Congelar presupuesto</button>`;
    $("#sGo", body).onclick = async () => {
      const op = snapOp(k, $("#sLabel", body).value.trim() || def);
      closeSheet();
      await doOps([op], "Presupuesto congelado");
    };
    openSheet("Congelar presupuesto", body);
  }

  function snapHistory(k) {
    const snaps = snapsOf(k);
    const cur = { id: "", label: "Ahora", items: rowsToItems(currentRows(k)), now: true };
    const all = snaps.concat(snaps.length && snapDiff(snaps[snaps.length - 1].items, cur.items).length ? [cur] : []);
    const body = document.createElement("div");
    if (!snaps.length) body.innerHTML = `<p class="small muted">Aún no hay fotos de ${esc(monthLabel(k))}.</p>`;
    else {
      const first = snaps[0], tf = snapTotals(first.items), tc = snapTotals(cur.items);
      body.innerHTML = `<div class="mini-kpis"><div><span>Gastos al inicio</span><b class="num">${eur(tf.gas)}</b></div><div><span>Gastos ahora</span><b class="num">${eur(tc.gas)}</b></div><div><span>Reajustes</span><b class="num">${all.length - 1}</b></div></div>
        <div class="timeline">${all.map((s2, i) => {
          const prev = all[i - 1], diff = prev ? snapDiff(prev.items, s2.items) : [], t = snapTotals(s2.items);
          return `<div class="tl-item ${s2.now ? "now" : ""}"><div class="tl-dot"></div><div class="tl-body">
            <div class="tl-head"><b>${esc(s2.label || "Foto")}</b><span class="xsmall muted">${s2.now ? "sin congelar" : esc(snapDate(s2.id))}</span></div>
            <div class="xsmall muted">Gastos ${eur(t.gas)} · ahorro ${eur(t.aho)} · por asignar ${eur(t.libre)}</div>
            ${prev ? (diff.length ? `<div class="dchips">${diffChips(diff)}</div>` : `<div class="xsmall muted">Sin cambios</div>`) : ""}
            ${s2.now ? `<button class="link" type="button" data-freeze="1">Congelar ahora</button>` : `<button class="link del-snap" type="button" data-id="${esc(s2.id)}">Borrar foto</button>`}
          </div></div>`;
        }).join("")}</div>`;
    }
    body.insertAdjacentHTML("beforeend", `<div style="height:12px"></div><button class="btn" id="hFreeze" type="button">📸 Congelar presupuesto</button>`);
    $("#hFreeze", body).onclick = () => { closeSheet(); setTimeout(() => freezeBudget(k), 60); };
    $$("[data-freeze]", body).forEach(b2 => b2.onclick = () => { closeSheet(); setTimeout(() => freezeBudget(k), 60); });
    $$(".del-snap", body).forEach(b2 => b2.onclick = async () => {
      if (!confirm("¿Borrar esta foto del presupuesto?")) return;
      closeSheet(); await doOps([{ type: "deleteSnapshot", data: { id: b2.dataset.id } }], "Foto borrada");
    });
    openSheet(`Fotos · ${monthLabel(k)}`, body);
  }

  // Presupuesto inicial (primera foto) o, si no hay fotos, el actual
  function initialBudget(cat, k) {
    const s0 = snapsOf(k)[0];
    if (s0 && s0.items[cat]) return s0.items[cat].v;
    return budgetOf(cat, k);
  }
  // Estadísticas de reorganización por partida en todos los meses con fotos
  function reorgStats() {
    const b = S.model.budget, months = Array.from(new Set((S.model.snaps || []).map(x => x.mes))).sort();
    const st = {};
    months.forEach(k => {
      const snaps = snapsOf(k), cur = rowsToItems(currentRows(k));
      const seq = snaps.map(x => x.items).concat([cur]);
      for (let i = 1; i < seq.length; i++) snapDiff(seq[i - 1], seq[i]).forEach(d => {
        const o = st[d.cat] = st[d.cat] || { cat: d.cat, tipo: d.tipo, changes: 0, up: 0, down: 0, net: 0, months: new Set() };
        o.changes++; if (d.d > 0) o.up += d.d; else o.down -= d.d; o.months.add(k);
      });
      snapDiff(seq[0], cur).forEach(d => { const o = st[d.cat]; if (o) o.net += d.d; });
    });
    Object.values(st).forEach(o => { o.nMonths = o.months.size; o.avgNet = o.net / Math.max(1, months.length); });
    return { months, list: Object.values(st).sort((a, b2) => b2.changes - a.changes || (b2.up + b2.down) - (a.up + a.down)) };
  }

  // ── Planificar: consejos y referencias para presupuestar ──
  function insights(k) {
    const b = S.model.budget, out = [], mk = MES[+k.split("-")[1] - 1];
    const exp = b.items.filter(i => i.tipo === "Gasto");
    exp.forEach(i => {
      const s = catStats(i.cat, k), bud = budgetOf(i.cat, k);
      if (s.avg == null) return;
      if (s.overs >= 2 && s.nRefB >= 2) {
        out.push({ ico: "⚠️", impact: Math.abs(s.dev || 0) + 30, cat: i.cat, text: `<b>${esc(i.cat)}</b>: te has pasado ${s.overs} de los últimos ${s.nRefB} meses (gastas de media ${eur(s.avg)}). ${bud != null && bud < s.avg ? `Para ${mk} tienes ${eur(bud)}: súbelo a ~${eur(Math.round(s.avg))} o recorta.` : "Vigila esta partida."}` });
      } else if (bud != null && s.avg > 20 && bud < s.avg * 0.8) {
        out.push({ ico: "⚠️", impact: s.avg - bud, cat: i.cat, text: `<b>${esc(i.cat)}</b>: para ${mk} presupuestas ${eur(bud)}, pero tu media es ${eur(s.avg)}. Te faltarían ~${eur(s.avg - bud)}.` });
      }
      if (bud != null && bud >= 20 && s.avg < bud * 0.35 && s.ref.length >= 2) {
        out.push({ ico: "💡", impact: bud - s.avg, cat: i.cat, text: `<b>${esc(i.cat)}</b>: presupuestas ${eur(bud)} y sueles gastar ${eur(s.avg)}. Podrías liberar ~${eur(bud - Math.round(s.avg))} para ahorro u otras partidas.` });
      }
      if (s.vals.length >= 3) {
        const prev = (s.vals[0] + s.vals[1]) / 2, last = s.vals[2];
        if (last > prev * 1.35 && last - prev > 25) out.push({ ico: "📈", impact: last - prev, cat: i.cat, text: `<b>${esc(i.cat)}</b> sube: ${eur(last)} en ${esc(monthShort(s.ref[2]))} frente a ${eur(prev)} de media antes.` });
      }
    });
    reorgStats().list.filter(o => o.tipo === "Gasto" && o.nMonths >= 2 && o.avgNet >= 15).forEach(o => {
      const bud = budgetOf(o.cat, k);
      out.push({ ico: "🔁", impact: o.avgNet + 20, cat: o.cat, text: `<b>${esc(o.cat)}</b>: la subes durante el mes en ${o.nMonths} meses (+${eur(o.avgNet)} de media). ${bud != null ? `Empieza ${mk} con ~${eur(Math.round(bud + o.avgNet))} en vez de ${eur(bud)}.` : ""}` });
    });
    reorgStats().list.filter(o => o.tipo === "Gasto" && o.nMonths >= 2 && o.avgNet <= -15).forEach(o => {
      out.push({ ico: "🔁", impact: -o.avgNet, cat: o.cat, text: `<b>${esc(o.cat)}</b>: sueles bajarla durante el mes (${eur(o.avgNet)} de media). Es una buena fuente para reasignar desde el principio.` });
    });
    const c = counter(k);
    if (c) {
      if (c.libre < -1) out.push({ ico: "⚠️", impact: 1e6, text: `Para ${mk} has asignado <b>${eur(-c.libre)}</b> más de lo que ingresas después de ahorrar. Ajusta partidas o el ahorro.` });
      else if (c.libre > 30) out.push({ ico: "💡", impact: c.libre, text: `Te quedan <b>${eur(c.libre)}</b> sin asignar en ${mk}. Puedes reservarlos como colchón, sumarlos al ahorro o repartirlos.` });
    }
    const ref = refMonths(k, 3);
    if (ref.length) {
      const ing = ref.reduce((a, m) => a + monthStats(m).ing, 0), aho = ref.reduce((a, m) => a + monthStats(m).aho, 0);
      const gas = ref.reduce((a, m) => a + monthStats(m).gas, 0), pres = ref.reduce((a, m) => a + monthStats(m).pres, 0);
      if (b.rate && ing > 0 && aho / ing < b.rate.value - 0.03) out.push({ ico: "🐷", impact: 500, text: `En los últimos ${ref.length} meses has ahorrado el <b>${Math.round(aho / ing * 100)}%</b> de tus ingresos; tu objetivo es el ${Math.round(b.rate.value * 100)}%.` });
      else if (b.rate && ing > 0) out.push({ ico: "✅", impact: 1, text: `Ahorro en línea con tu objetivo: <b>${Math.round(aho / ing * 100)}%</b> de los ingresos en los últimos ${ref.length} meses.` });
      if (gas > 0 && pres / gas > 0.5) out.push({ ico: "💡", impact: 200, text: `Los gastos prescindibles son el <b>${Math.round(pres / gas * 100)}%</b> de lo que gastas. Ahí está tu mayor margen de ajuste.` });
    }
    return out.sort((a, b2) => b2.impact - a.impact);
  }

  function renderPlan(v, k) {
    const b = S.model.budget, mk = MES[+k.split("-")[1] - 1];
    const ref = refMonths(k, 3);
    // Consejos
    const ins = insights(k);
    const ci = document.createElement("div"); ci.className = "card";
    ci.innerHTML = `<h3>Consejos para ${esc(mk)}</h3><p class="xsmall muted" style="margin:-6px 0 8px">Según ${ref.length ? `${ref.map(monthShort).join(", ")}` : "tus datos"}.</p>
      ${ins.length ? `<div class="tips">${ins.slice(0, 6).map(t => `<div class="tip" ${t.cat ? `data-cat="${esc(t.cat)}"` : ""}><span class="tip-ico">${t.ico}</span><span>${t.text}</span></div>`).join("")}</div>` : `<p class="small muted">Aún no hay meses suficientes para darte consejos.</p>`}`;
    $$(".tip[data-cat]", ci).forEach(t => { t.style.cursor = "pointer"; t.onclick = () => { const it = b.items.find(i => i.cat === t.dataset.cat); if (it) editBudgetItem(it.row, k); }; });
    v.appendChild(ci);

    // Referencia por partida
    const cr = document.createElement("div"); cr.className = "card";
    const exp = b.items.filter(i => i.tipo === "Gasto").map(i => ({ i, s: catStats(i.cat, k), bud: budgetOf(i.cat, k) }))
      .filter(x => (x.s.avg || 0) > 0 || (x.bud || 0) > 0).sort((a, b2) => (b2.s.avg || 0) - (a.s.avg || 0));
    cr.innerHTML = `<h3>Referencia por partida</h3>
      <p class="xsmall muted" style="margin:-6px 0 8px">Barras: lo gastado en los últimos meses (azul, el más reciente). Raya: lo presupuestado para ${esc(mk)}. Toca para ajustar.</p>
      <div class="refs">${exp.map(x => {
        const tag = x.bud == null || x.s.avg == null ? "" : (x.bud < x.s.avg * 0.85 && x.s.avg - x.bud > 10 ? `<span class="tag warn">Corto</span>` : (x.bud > x.s.avg * 1.5 && x.bud - x.s.avg > 20 ? `<span class="tag">Holgado</span>` : ""));
        return `<button type="button" class="ref-row" data-row="${x.i.row}"><span class="ref-name">${icon(x.i.cat)} ${esc(x.i.cat)} ${tag}</span>
          ${Charts.sparkbars(x.s.sixVals.slice(-4), x.bud)}
          <span class="ref-vals num"><b>${x.bud != null ? eur(x.bud) : "—"}</b><small>media ${x.s.avg != null ? eur(x.s.avg) : "—"}</small></span></button>`;
      }).join("")}</div>`;
    $$(".ref-row", cr).forEach(r => r.onclick = () => editBudgetItem(+r.dataset.row, k));
    v.appendChild(cr);

    // Desviación media (frente al presupuesto inicial o al final)
    const hasSnapsRef = ref.some(m => snapsOf(m).length);
    const base = hasSnapsRef && S.devBase === "init" ? "init" : "final";
    const dkey = base === "init" ? "devInit" : "dev";
    const devs = b.items.filter(i => i.tipo === "Gasto").map(i => ({ i, s: catStats(i.cat, k) })).filter(x => x.s[dkey] != null && Math.abs(x.s[dkey]) >= 1)
      .sort((a, b2) => b2.s[dkey] - a.s[dkey]);
    if (devs.length) {
      const cd = card("¿Dónde fallaba el presupuesto?", v, `Diferencia media entre lo gastado y lo presupuestado ${base === "init" ? "<b>al inicio de cada mes</b>" : "(presupuesto final)"} · ${ref.map(monthShort).join(", ")}.`);
      if (hasSnapsRef) {
        const sg = document.createElement("div"); sg.className = "seg"; sg.style.margin = "0 0 10px";
        sg.innerHTML = `<button type="button" class="${base === "final" ? "on" : ""}" data-b="final">Presupuesto final</button><button type="button" class="${base === "init" ? "on" : ""}" data-b="init">Presupuesto inicial</button>`;
        $$("button", sg).forEach(x => x.onclick = () => { S.devBase = x.dataset.b; render(); });
        $(".ch", cd).before(sg);
      }
      requestAnimationFrame(() => Charts.diverging($(".ch", cd), devs.map(x => ({
        label: x.i.cat, value: x.s[dkey],
        tip: `De media gastaste ${eur(x.s.avg)}; ${x.s[dkey] > 0 ? "te pasaste " + eur(x.s[dkey]) : "te sobraron " + eur(-x.s[dkey])} al mes${base === "init" ? " respecto a lo que presupuestaste al empezar" : ""}.`,
        onClick: () => editBudgetItem(x.i.row, k)
      }))));
    }

    // Reorganizaciones del presupuesto
    const rs = reorgStats();
    const cro = document.createElement("div"); cro.className = "card";
    if (!rs.months.length) {
      cro.innerHTML = `<h3>Reorganizaciones del presupuesto</h3><p class="small muted" style="margin:0">Cuando congeles el presupuesto (📸 en Gastos) verás aquí qué partidas reorganizas más y cuánto se mueven a lo largo del mes.</p>`;
      v.appendChild(cro);
    } else {
      const sk = snapsOf(k), curK = rowsToItems(currentRows(k));
      const dK = sk.length ? snapDiff(sk[0].items, curK) : [];
      cro.innerHTML = `<h3>Reorganizaciones del presupuesto</h3>
        <p class="xsmall muted" style="margin:-6px 0 8px">${rs.months.length} mes${rs.months.length === 1 ? "" : "es"} con fotos (${rs.months.map(monthShort).join(", ")}).</p>
        ${sk.length ? `<div class="small" style="margin-bottom:6px">En <b>${esc(mk)}</b>, del inicio a ahora (${sk.length} foto${sk.length === 1 ? "" : "s"}):</div><div class="dchips" style="margin-bottom:10px">${dK.length ? diffChips(dK, 8) : `<span class="xsmall muted">sin cambios</span>`}</div>` : ""}
        <table class="table"><thead><tr><th>Partida</th><th>Cambios</th><th>Subidas</th><th>Bajadas</th></tr></thead><tbody>
        ${rs.list.slice(0, 8).map(o => `<tr><td>${icon(o.cat)} ${esc(o.cat)}</td><td class="num">${o.changes}</td><td class="num pos">${o.up ? "+" + eur(o.up) : "—"}</td><td class="num neg">${o.down ? "−" + eur(o.down) : "—"}</td></tr>`).join("") || `<tr><td colspan="4" class="muted">Sin cambios todavía</td></tr>`}
        </tbody></table><div class="ch" style="margin-top:12px"></div>`;
      $(".table", cro).insertAdjacentHTML("beforebegin", `<button type="button" class="link" id="seeHist" style="margin-bottom:8px">Ver fotos de ${esc(mk)} ›</button>`);
      $("#seeHist", cro).onclick = () => snapHistory(k);
      v.appendChild(cro);
      const nets = rs.list.filter(o => Math.abs(o.avgNet) >= 1 && o.tipo === "Gasto").sort((a, b2) => b2.avgNet - a.avgNet);
      if (nets.length) {
        $(".ch", cro).insertAdjacentHTML("beforebegin", `<div class="small" style="margin-top:10px"><b>Cuánto cambia cada partida durante el mes</b> <span class="xsmall muted">(media del inicio al final)</span></div>`);
        requestAnimationFrame(() => Charts.diverging($(".ch", cro), nets.map(o => ({
          label: o.cat, value: o.avgNet,
          tip: `La has cambiado ${o.changes} ${o.changes === 1 ? "vez" : "veces"} en ${o.nMonths} mes${o.nMonths === 1 ? "" : "es"}: subidas ${eur(o.up)}, bajadas ${eur(o.down)}.`
        })), { leftLabel: "← la bajas", rightLabel: "la subes →" }));
      }
    }

    // Fijos vs variables
    const ms = refMonths(k, 6);
    if (ms.length) {
      const fij = ms.map(m => txOfMonth(m).filter(t => t.tipo === "Gasto" && t.per === "Fijo").reduce((a, t) => a + t.importe, 0));
      const vari = ms.map(m => txOfMonth(m).filter(t => t.tipo === "Gasto" && t.per !== "Fijo").reduce((a, t) => a + t.importe, 0));
      const avgF = fij.reduce((a, x) => a + x, 0) / ms.length, avgV = vari.reduce((a, x) => a + x, 0) / ms.length;
      const cf = card("Fijos frente a variables", v, `Tus gastos fijos rondan <b>${eur(avgF)}</b>/mes: es la base de cada presupuesto. Lo variable (media ${eur(avgV)}) es donde puedes ajustar.`);
      requestAnimationFrame(() => Charts.stackedBars($(".ch", cf), ms.map(monthShort), [
        { name: "Fijos", color: SERIES[0], values: fij }, { name: "Variables", color: SERIES[1], values: vari }
      ], { fullLabels: ms.map(monthLabel), target: ms.map(m => b.months.some(x => x.key === m) ? b.items.filter(i => i.tipo === "Gasto").reduce((a, i) => a + (i.vals[m] || 0), 0) : null) }));
    }

    const bt = document.createElement("div"); bt.className = "btn-row";
    bt.innerHTML = `<button class="btn" type="button" id="pAvg">Preparar ${esc(mk)} con tus medias</button>`;
    $("#pAvg", bt).onclick = () => editBudgetMonth(k, "avg");
    v.appendChild(bt);
  }

  // ── ANÁLISIS ──────────────────────────────────────────────
  function renderAnalisis(v) {
    const months = txMonths(); if (!months.length) { v.innerHTML = `<div class="empty">Sin datos todavía.</div>`; return; }
    clampMonth(months);
    const k = S.month, st = monthStats(k);
    v.innerHTML = "";
    v.appendChild(monthSwitch(months, render));

    const kp = document.createElement("div"); kp.className = "kpis";
    const days = k === todayISO().slice(0, 7) ? +todayISO().slice(8, 10) : new Date(+k.slice(0, 4), +k.slice(5, 7), 0).getDate();
    kp.innerHTML = `<div class="kpi"><span>Gasto total</span><b class="num">${eur(st.gas)}</b><small class="muted">${eur(st.gas / days)} / día</small></div>
      <div class="kpi"><span>Tasa de ahorro</span><b class="num">${st.ing ? Math.round(st.aho / st.ing * 100) : 0}%</b><small class="muted">${eur(st.aho)} de ${eur(st.ing)}</small></div>
      <div class="kpi"><span>Esencial / prescindible</span><b class="num">${st.gas ? Math.round(st.esen / st.gas * 100) : 0}% / ${st.gas ? Math.round(st.pres / st.gas * 100) : 0}%</b></div>
      <div class="kpi"><span>Fijo / variable</span>${(() => { const f = txOfMonth(k).filter(t => t.tipo === "Gasto" && t.per === "Fijo").reduce((a, t) => a + t.importe, 0); return `<b class="num">${st.gas ? Math.round(f / st.gas * 100) : 0}% / ${st.gas ? Math.round((st.gas - f) / st.gas * 100) : 0}%</b>`; })()}</div>`;
    v.appendChild(kp);

    // Donut por categoría
    const byCat = Object.entries(sumBy(txOfMonth(k).filter(t => t.tipo === "Gasto"), t => t.cat)).sort((a, b) => b[1] - a[1]);
    const top = byCat.slice(0, 7), rest = byCat.slice(7).reduce((a, x) => a + x[1], 0);
    const data = top.map(([c, val], i) => ({ label: c, value: val, color: SERIES[i] }));
    if (rest > 0) data.push({ label: "Resto", value: rest, color: "var(--muted-mark)" });
    const c1 = card("Gastos por categoría", v);
    Charts.donut($(".ch", c1), data, { center: "gastado" });

    // Real vs presupuesto
    const hasSn = snapsOf(k).length > 0;
    const rows = budgetRows(k, "Gasto").filter(r => r.real > 0 || (r.budget || 0) > 0).sort((a, b) => b.real - a.real).map(r => ({ label: r.cat, real: r.real, budget: r.budget, budget0: hasSn ? initialBudget(r.cat, k) : null, onClick: () => goMovs({ cat: r.cat, month: k }) }));
    const c2 = card("Real vs presupuesto", v, rows.some(r => r.budget != null) ? `La raya negra es el presupuesto${hasSn ? "; la gris discontinua, el que tenías al empezar el mes" : ""}. En rojo, lo que se pasa. Toca una barra para ver sus movimientos.` : "Este mes no tiene presupuesto en PPTO.");
    requestAnimationFrame(() => Charts.bullet($(".ch", c2), rows));

    // Acceso a la planificación del siguiente mes
    const nextB = S.model.budget.months.map(m => m.key).find(m => m > nowMonth()) || S.model.budget.months.map(m => m.key).find(m => m >= nowMonth());
    if (nextB) {
      const pl = document.createElement("button"); pl.type = "button"; pl.className = "counter-pill free";
      pl.innerHTML = `<span>💡 Consejos y referencias para presupuestar ${MES[+nextB.split("-")[1] - 1]}</span><b>›</b>`;
      pl.onclick = () => { S.month = nextB; S.budgetSeg = "Plan"; go("presupuesto"); };
      v.appendChild(pl);
    }

    // Evolución mensual
    const tm = months.filter(m => m <= todayISO().slice(0, 7)).slice(-12);
    const c3 = card("Ingresos, gastos y ahorro", v);
    requestAnimationFrame(() => Charts.line($(".ch", c3), tm.map(monthShort), [
      { name: "Ingresos", color: SERIES[2], values: tm.map(m => monthStats(m).ing) },
      { name: "Gastos", color: SERIES[1], values: tm.map(m => monthStats(m).gas) },
      { name: "Ahorro", color: SERIES[0], values: tm.map(m => monthStats(m).aho) }
    ], { fullLabels: tm.map(monthLabel) }));

    // Tendencia de las categorías principales
    const topCats = Object.entries(sumBy(S.model.tx.list.filter(t => t.tipo === "Gasto" && tm.includes(t.fecha.slice(0, 7))), t => t.cat)).sort((a, b) => b[1] - a[1]).slice(0, 4).map(x => x[0]);
    const c4 = card("Tendencia de tus 4 mayores categorías", v);
    requestAnimationFrame(() => Charts.line($(".ch", c4), tm.map(monthShort), topCats.map((c, i) => ({
      name: c, color: SERIES[i], values: tm.map(m => txOfMonth(m).filter(t => t.cat === c).reduce((a, t) => a + t.importe, 0))
    })), { fullLabels: tm.map(monthLabel) }));

    // Top conceptos
    const notes = Object.entries(sumBy(txOfMonth(k).filter(t => t.tipo === "Gasto"), t => cleanNote(t.notas))).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const c5 = card("¿En qué se fue el dinero?", v);
    $(".ch", c5).innerHTML = `<table class="table"><tbody>${notes.map(([n, val]) => `<tr><td>${esc(n)}</td><td class="num">${eur(val)}</td></tr>`).join("")}</tbody></table>`;
  }
  function card(title, parent, sub) {
    const c = document.createElement("div"); c.className = "card";
    c.innerHTML = `<h3>${title}</h3>${sub ? `<p class="xsmall muted" style="margin:-6px 0 8px">${sub}</p>` : ""}<div class="ch"></div>`;
    parent.appendChild(c); return c;
  }

  // ── PATRIMONIO ────────────────────────────────────────────
  function netWorthSeries() {
    const bal = S.model.balance;
    return bal.months.map((m, x) => {
      let any = false, a = 0, p = 0;
      bal.sections.forEach(s => s.items.forEach(it => { const val = it.vals[x]; if (val != null) { any = true; if (s.side === "pasivo") p += val; else a += val; } }));
      return { name: m.name, key: m.key, x, v: any ? a - p : null, a, p };
    });
  }
  function renderPatrimonio(v) {
    v.innerHTML = "";
    const seg = document.createElement("div"); seg.className = "seg";
    seg.innerHTML = ["Balance", "Inversiones", "Ahorro"].map(x => `<button type="button" class="${S.patSeg === x ? "on" : ""}">${x}</button>`).join("");
    $$("button", seg).forEach(b => b.onclick = () => { S.patSeg = b.textContent; render(); });
    v.appendChild(seg);
    ({ Balance: renderBalance, Inversiones: renderInvest, Ahorro: renderSavings })[S.patSeg](v);
  }

  function renderBalance(v) {
    const bal = S.model.balance, nw = netWorthSeries();
    if (!bal.months.length) { const e = document.createElement("div"); e.className = "empty"; e.textContent = `No he podido leer la hoja «${window.APP_CONFIG.SHEETS.balance}» del Excel.`; v.appendChild(e); return; }
    const withData = nw.filter(x => x.v != null);
    if (S.balMonth == null || !withData.find(x => x.x === S.balMonth)) S.balMonth = withData.length ? withData[withData.length - 1].x : 0;
    const cur = nw[S.balMonth], prev = withData.filter(x => x.x < S.balMonth).pop();
    const hero = document.createElement("div"); hero.className = "card hero";
    hero.innerHTML = `<div class="label">Patrimonio neto · ${esc(cur.name.toLowerCase())}</div><div class="big num">${eur(cur.v || 0)}</div>
      <div class="hero-row"><div><span>Activos</span><b class="num">${eur(cur.a)}</b></div><div><span>Pasivos</span><b class="num">${eur(cur.p)}</b></div></div>
      ${prev ? `<div class="small" style="margin-top:10px">${cur.v >= prev.v ? "▲" : "▼"} ${eur(Math.abs(cur.v - prev.v))} respecto a ${esc(prev.name.toLowerCase())}</div>` : ""}`;
    v.appendChild(hero);

    const chips = document.createElement("div"); chips.className = "chips";
    chips.innerHTML = withData.map(x => `<button type="button" class="chip ${x.x === S.balMonth ? "on" : ""}" data-x="${x.x}">${esc(x.name)}</button>`).join("");
    $$(".chip", chips).forEach(c => c.onclick = () => { S.balMonth = +c.dataset.x; render(); });
    v.appendChild(chips);

    if (withData.length > 1) {
      const c = card("Evolución del patrimonio neto", v);
      requestAnimationFrame(() => Charts.line($(".ch", c), withData.map(x => x.name.slice(0, 3)), [{ name: "Patrimonio neto", color: SERIES[0], values: withData.map(x => x.v) }], { zero: false, fullLabels: withData.map(x => x.name) }));
    }
    bal.sections.forEach(s => {
      const c = document.createElement("div"); c.className = "card";
      const tot = s.items.reduce((a, it) => a + (it.vals[S.balMonth] || 0), 0);
      const pt = prev ? s.items.reduce((a, it) => a + (it.vals[prev.x] || 0), 0) : null;
      c.innerHTML = `<div class="card-head"><h3>${esc(s.name)}${s.side === "pasivo" ? " (pasivo)" : ""}</h3><b class="num">${eur(tot)}</b></div>
        <table class="table"><tbody>${s.items.map(it => {
          const val = it.vals[S.balMonth], pv = prev ? it.vals[prev.x] : null;
          const d = val != null && pv != null ? val - pv : null;
          return `<tr><td>${esc(it.label)}</td><td class="num">${val == null ? "—" : eur(val)}</td><td class="num xsmall ${d > 0 ? "pos" : d < 0 ? "neg" : "muted"}" style="width:84px">${d == null || Math.abs(d) < 0.005 ? "" : (d > 0 ? "+" : "") + eur(d)}</td></tr>`;
        }).join("")}</tbody></table>`;
      v.appendChild(c);
    });
    const next = bal.months.findIndex((m, x) => x > (withData.length ? withData[withData.length - 1].x : -1));
    const btns = document.createElement("div"); btns.className = "btn-row";
    btns.innerHTML = `${next >= 0 ? `<button class="btn" id="bNew" type="button">Registrar ${esc(bal.months[next].name.toLowerCase())}</button>` : ""}<button class="btn secondary" id="bEdit" type="button">Editar ${esc(cur.name.toLowerCase())}</button>`;
    if (next >= 0) $("#bNew", btns).onclick = () => editBalance(next, withData.length ? withData[withData.length - 1].x : null);
    $("#bEdit", btns).onclick = () => editBalance(S.balMonth, null);
    v.appendChild(btns);
  }

  function editBalance(x, copyFrom) {
    const bal = S.model.balance, m = bal.months[x];
    const body = document.createElement("div");
    body.innerHTML = `<p class="small muted" style="margin-top:0">${copyFrom != null ? `Valores precargados de ${esc(bal.months[copyFrom].name.toLowerCase())}. Actualiza los que hayan cambiado.` : "Actualiza los saldos de este mes."}</p>` +
      bal.sections.map((s, si) => `<div class="bgroup-h"><h4>${esc(s.name)}</h4></div>` + s.items.map((it, ii) => {
        const val = it.vals[x] != null ? it.vals[x] : (copyFrom != null ? it.vals[copyFrom] : null);
        return `<div class="brow"><div class="brow-top" style="align-items:center"><span class="name small">${esc(it.label)}</span><input class="budget-edit num" data-s="${si}" data-i="${ii}" inputmode="decimal" value="${val == null ? "" : String(round2(val)).replace(".", ",")}" placeholder="—"></div></div>`;
      }).join("")).join("") + `<div style="height:12px"></div><button class="btn" id="save" type="button">Guardar ${esc(m.name.toLowerCase())}</button>`;
    $("#save", body).onclick = async () => {
      const cells = [];
      $$(".budget-edit", body).forEach(inp => {
        const s = bal.sections[+inp.dataset.s], it = s.items[+inp.dataset.i];
        const raw = inp.value.trim();
        const val = raw === "" ? null : round2(parseFloat(raw.replace(/\./g, "").replace(",", ".")) || 0);
        if (val === it.vals[x] || (val != null && it.vals[x] != null && Math.abs(val - it.vals[x]) < 0.004)) return;
        const col = s.side === "pasivo" ? s.cols[x] : m.col;
        cells.push({ addr: col + it.row, value: val });
      });
      closeSheet();
      if (!cells.length) return toast("Sin cambios");
      S.balMonth = x;
      await doOp("setBalance", { cells }, "Balance actualizado");
    };
    openSheet(`Balance · ${m.name}`, body);
  }

  function renderInvest(v) {
    const bal = S.model.balance;
    const sec = bal.sections.find(s => /inversi/i.test(s.name));
    if (sec) {
      const idxs = bal.months.map((m, x) => x).filter(x => sec.items.some(it => it.vals[x] != null));
      const last = idxs[idxs.length - 1], prev = idxs[idxs.length - 2];
      const tot = x => sec.items.reduce((a, it) => a + (it.vals[x] || 0), 0);
      const hero = document.createElement("div"); hero.className = "card hero";
      const d = prev != null ? tot(last) - tot(prev) : null;
      hero.innerHTML = `<div class="label">Cartera · ${last != null ? esc(bal.months[last].name.toLowerCase()) : ""}</div><div class="big num">${last != null ? eur(tot(last)) : "—"}</div>
        ${d != null ? `<div class="small">${d >= 0 ? "▲" : "▼"} ${eur(Math.abs(d))} (${tot(prev) ? (d / tot(prev) * 100).toFixed(1) : 0}%) vs ${esc(bal.months[prev].name.toLowerCase())} · incluye aportaciones</div>` : ""}`;
      v.appendChild(hero);
      if (last != null) {
        const c = card("Distribución actual", v);
        Charts.donut($(".ch", c), sec.items.filter(it => it.vals[last]).map((it, i) => ({ label: it.label, value: it.vals[last], color: SERIES[i % 8] })), { center: "invertido" });
      }
      if (idxs.length > 1) {
        const c = card("Evolución por producto", v);
        requestAnimationFrame(() => Charts.line($(".ch", c), idxs.map(x => bal.months[x].name.slice(0, 3)), sec.items.slice(0, 8).map((it, i) => ({ name: it.label, color: SERIES[i], values: idxs.map(x => it.vals[x]) })), { fullLabels: idxs.map(x => bal.months[x].name) }));
      }
      const c = card("Detalle", v);
      $(".ch", c).innerHTML = `<div class="scroll-x"><table class="table"><thead><tr><th>Producto</th>${idxs.slice(-3).map(x => `<th>${esc(bal.months[x].name.slice(0, 3))}</th>`).join("")}</tr></thead><tbody>${sec.items.map(it => `<tr><td>${esc(it.label)}</td>${idxs.slice(-3).map(x => `<td class="num">${it.vals[x] == null ? "—" : eur(it.vals[x])}</td>`).join("")}</tr>`).join("")}</tbody>
        <tfoot><tr><td>Total</td>${idxs.slice(-3).map(x => `<td class="num">${eur(tot(x))}</td>`).join("")}</tr></tfoot></table></div>
        <p class="xsmall muted">Para actualizar los valores usa <b>Balance → Registrar/Editar mes</b>.</p>`;
    }
    const plan = S.model.invest;
    if (plan) {
      const c = document.createElement("div"); c.className = "card";
      const split = Excel.investSplit(plan, plan.amount);
      c.innerHTML = `<div class="card-head"><h3>Aportación de este mes</h3><button class="link" type="button">Cambiar importe</button></div>
        <div style="font-size:26px;font-weight:700" class="num">${eur(plan.amount)}</div>
        <table class="table" style="margin-top:8px"><tbody>${split.map(r => `<tr><td>${esc(r.name)}<div class="xsmall muted">${esc(r.kind)}${r.isin ? " · " + esc(r.isin) : ""}</div></td><td class="num xsmall muted">${Math.round(r.pct * 100)}%</td><td class="num"><b>${eur(r.eur)}</b></td></tr>`).join("")}</tbody></table>
        <p class="xsmall muted">Reparto según tu hoja «Ahorro - Inversión».</p>`;
      $(".link", c).onclick = () => {
        const body = document.createElement("div");
        body.innerHTML = `<div class="field"><label>Importe a invertir</label><input id="amt" class="amount-input num" inputmode="decimal" value="${String(round2(plan.amount)).replace(".", ",")}"></div><div id="prev"></div><button class="btn" id="sv" type="button">Guardar en el Excel</button>`;
        const upd = () => { const a = parseFloat($("#amt", body).value.replace(",", ".")) || 0; $("#prev", body).innerHTML = `<table class="table">${Excel.investSplit(plan, a).map(r => `<tr><td>${esc(r.name)}</td><td class="num">${eur(r.eur)}</td></tr>`).join("")}</table><div style="height:12px"></div>`; };
        $("#amt", body).oninput = upd; upd();
        $("#sv", body).onclick = async () => { const a = round2(parseFloat($("#amt", body).value.replace(",", ".")) || 0); closeSheet(); await doOp("setInvestAmount", { value: a }, "Aportación actualizada"); };
        openSheet("Aportación mensual", body);
      };
      v.appendChild(c);
    }
  }

  function renderSavings(v) {
    const b = S.model.budget;
    const months = b.months.map(m => m.key);
    const realAho = k => txOfMonth(k).filter(t => t.tipo === "Ahorro").reduce((a, t) => a + t.importe, 0);
    const pastMonths = months.filter(k => k <= todayISO().slice(0, 7));
    const accReal = pastMonths.reduce((a, k) => a + realAho(k), 0);
    const planYear = b.savings ? months.reduce((a, k) => a + (b.savings.vals[k] || 0), 0) : 0;
    const hero = document.createElement("div"); hero.className = "card hero";
    hero.innerHTML = `<div class="label">Ahorrado ${months.length ? "desde " + monthLabel(months[0]).split(" ")[0] : ""}</div><div class="big num">${eur(accReal)}</div>
      <div class="hero-row"><div><span>Plan hasta ${months.length ? monthLabel(months[months.length - 1]).split(" ")[0] : ""}</span><b class="num">${eur(planYear)}</b></div>
      <div><span>Tasa de ahorro objetivo</span><b class="num">${b.rate ? Math.round(b.rate.value * 100) + "%" : "—"}</b></div></div>`;
    v.appendChild(hero);

    const c1 = card("Ahorro real vs plan", v, "Barras: lo ahorrado cada mes. Raya: lo previsto en PPTO.");
    requestAnimationFrame(() => Charts.stackedBars($(".ch", c1), months.map(monthShort), [{ name: "Ahorro real", color: SERIES[0], values: months.map(realAho) }], { target: b.savings ? months.map(k => b.savings.vals[k] || 0) : null, targetName: "Plan", fullLabels: months.map(monthLabel), legend: true }));

    if (b.rate) {
      const c = document.createElement("div"); c.className = "card";
      c.innerHTML = `<div class="card-head"><h3>Tasa de ahorro</h3><button class="link" type="button">Cambiar</button></div>
        <p class="small muted" style="margin:0">En tu Excel el ahorro previsto de los meses futuros es <b>ingresos × ${Math.round(b.rate.value * 100)}%</b> (celda ${b.rate.addr} de PPTO). Si la cambias, se recalculan esos meses.</p>`;
      $(".link", c).onclick = () => {
        const body = document.createElement("div");
        body.innerHTML = `<div class="field"><label>Porcentaje de ahorro sobre ingresos</label><input id="r" class="amount-input num" inputmode="decimal" value="${Math.round(b.rate.value * 1000) / 10}"></div><button class="btn" id="sv" type="button">Guardar</button>`;
        $("#sv", body).onclick = async () => { const r = (parseFloat($("#r", body).value.replace(",", ".")) || 0) / 100; closeSheet(); await doOp("setBudgetCells", { cells: [{ addr: b.rate.addr, value: r }] }, "Tasa de ahorro actualizada"); };
        openSheet("Tasa de ahorro", body);
      };
      v.appendChild(c);
    }

    // Metas de ahorro (guardadas en el móvil)
    const goals = LS.get("fin_goals") || [];
    const avgSave = pastMonths.length ? pastMonths.reduce((a, k) => a + realAho(k), 0) / pastMonths.length : 0;
    const cg = document.createElement("div"); cg.className = "card";
    cg.innerHTML = `<div class="card-head"><h3>Metas</h3><button class="link" type="button">Añadir meta</button></div><div class="gl"></div>`;
    const gl = $(".gl", cg);
    if (!goals.length) gl.innerHTML = `<p class="small muted" style="margin:0">Crea una meta (colchón de emergencia, viaje, entrada de una casa…) y te digo cuánto apartar al mes.</p>`;
    goals.forEach((g, i) => {
      const monthsLeft = Math.max(1, (new Date(g.date).getFullYear() - new Date().getFullYear()) * 12 + new Date(g.date).getMonth() - new Date().getMonth());
      const need = Math.max(0, (g.target - g.saved) / monthsLeft);
      const p = g.target ? Math.min(100, g.saved / g.target * 100) : 0;
      const d = document.createElement("div"); d.className = "goal";
      d.innerHTML = `<div class="brow-top"><span class="name">${esc(g.name)}</span><span class="vals num">${eur(g.saved)} / <b>${eur(g.target)}</b></span></div>
        <div class="prog"><div style="width:${p}%"></div></div>
        <div class="xsmall muted" style="margin-top:6px">${g.saved >= g.target ? "¡Meta conseguida!" : `Necesitas <b>${eur(need)}/mes</b> durante ${monthsLeft} meses (hasta ${new Date(g.date).toLocaleDateString("es-ES", { month: "long", year: "numeric" })})`}</div>`;
      d.style.cursor = "pointer"; d.onclick = () => editGoal(i);
      gl.appendChild(d);
    });
    $(".link", cg).onclick = () => editGoal(-1);
    v.appendChild(cg);

    // Proyección
    const sec = S.model.balance.sections;
    const nw = netWorthSeries().filter(x => x.v != null);
    const lastX = nw.length ? nw[nw.length - 1].x : null;
    const fin = lastX == null ? 0 : sec.filter(s => /inversi|cash/i.test(s.name)).reduce((a, s) => a + s.items.reduce((q, it) => q + (it.vals[lastX] || 0), 0), 0);
    const pj = settings().proj || { monthly: Math.round(avgSave) || 300, rate: 5, years: 10 };
    const cp = document.createElement("div"); cp.className = "card";
    cp.innerHTML = `<h3>Proyección de tu dinero</h3>
      <div class="kpis" style="grid-template-columns:1fr 1fr 1fr;gap:8px">
        <div class="field" style="margin:0"><label>€/mes</label><input id="pm" inputmode="decimal" value="${pj.monthly}"></div>
        <div class="field" style="margin:0"><label>Rent. anual %</label><input id="pr" inputmode="decimal" value="${pj.rate}"></div>
        <div class="field" style="margin:0"><label>Años</label><input id="py" inputmode="numeric" value="${pj.years}"></div></div>
      <div class="ch" style="margin-top:10px"></div><p class="small" id="pres" style="margin:6px 0 0"></p>
      <p class="xsmall muted" style="margin:4px 0 0">Parte de tus inversiones + liquidez actuales (${eur(fin)}). Rentabilidad supuesta, no garantizada.</p>`;
    v.appendChild(cp);
    const drawProj = () => {
      const mo = parseFloat($("#pm", cp).value.replace(",", ".")) || 0, r = (parseFloat($("#pr", cp).value.replace(",", ".")) || 0) / 100, y = Math.min(40, Math.max(1, parseInt($("#py", cp).value) || 10));
      saveSettings({ proj: { monthly: mo, rate: r * 100, years: y } });
      const labels = [], withRet = [], noRet = []; let val = fin, plain = fin;
      for (let i = 0; i <= y; i++) {
        labels.push(String(new Date().getFullYear() + i)); withRet.push(Math.round(val)); noRet.push(Math.round(plain));
        for (let mth = 0; mth < 12; mth++) { val = val * (1 + r / 12) + mo; plain += mo; }
      }
      Charts.line($(".ch", cp), labels, [{ name: "Con rentabilidad", color: SERIES[0], values: withRet, noDots: true }, { name: "Solo aportaciones", color: SERIES[1], values: noRet, dash: true, noDots: true }], {});
      $("#pres", cp).innerHTML = `En ${y} años: <b>${eur(withRet[withRet.length - 1])}</b> (de ellos ${eur(withRet[withRet.length - 1] - noRet[noRet.length - 1])} de rentabilidad).`;
    };
    $$("input", cp).forEach(i => i.oninput = drawProj);
    requestAnimationFrame(drawProj);
  }

  function editGoal(i) {
    const goals = LS.get("fin_goals") || [];
    const g = i >= 0 ? goals[i] : { name: "", target: "", saved: 0, date: `${new Date().getFullYear() + 1}-06-30` };
    const body = document.createElement("div");
    body.innerHTML = `<div class="field"><label>Nombre</label><input id="gn" value="${esc(g.name)}" placeholder="Fondo de emergencia"></div>
      <div class="field"><label>Objetivo (€)</label><input id="gt" inputmode="decimal" value="${g.target}"></div>
      <div class="field"><label>Ya tengo (€)</label><input id="gs" inputmode="decimal" value="${g.saved}"></div>
      <div class="field"><label>Fecha límite</label><input id="gd" type="date" value="${g.date}"></div>
      <button class="btn" id="sv" type="button">Guardar</button>${i >= 0 ? `<div style="height:10px"></div><button class="btn danger" id="dl" type="button">Eliminar meta</button>` : ""}
      <p class="xsmall muted">Las metas se guardan en este móvil (no en el Excel).</p>`;
    $("#sv", body).onclick = () => {
      const n = { name: $("#gn", body).value.trim() || "Meta", target: parseFloat($("#gt", body).value.replace(",", ".")) || 0, saved: parseFloat($("#gs", body).value.replace(",", ".")) || 0, date: $("#gd", body).value };
      if (i >= 0) goals[i] = n; else goals.push(n);
      LS.set("fin_goals", goals); closeSheet(); render();
    };
    if (i >= 0) $("#dl", body).onclick = () => { goals.splice(i, 1); LS.set("fin_goals", goals); closeSheet(); render(); };
    openSheet(i >= 0 ? "Editar meta" : "Nueva meta", body);
  }

  // ── Hoja inferior ─────────────────────────────────────────
  function openSheet(title, body) {
    $("#sheetTitle").textContent = title;
    const b = $("#sheetBody"); b.innerHTML = ""; b.appendChild(body);
    $("#sheet").hidden = false; document.body.style.overflow = "hidden"; document.body.classList.add("sheet-open");
    history.pushState({ sheet: 1 }, "");
  }
  function closeSheet() {
    if ($("#sheet").hidden) return;
    $("#sheet").hidden = true; document.body.style.overflow = ""; document.body.classList.remove("sheet-open");
    if (history.state && history.state.sheet) history.back();
  }
  window.addEventListener("popstate", () => { if (!$("#sheet").hidden) { $("#sheet").hidden = true; document.body.style.overflow = ""; document.body.classList.remove("sheet-open"); } });

  // ── Login, selección de archivo y ajustes ────────────────
  function showLogin() {
    $$(".view").forEach(v => v.hidden = true);
    $("#tabbar").hidden = true; $("#fab").hidden = true;
    $("#viewTitle").textContent = "Mis Finanzas"; setStatus("", "Sin conectar");
    const v = $("#view-login"); v.hidden = false;
    const cid = Auth.cfg().clientId;
    v.innerHTML = `<div class="login-hero"><img class="logo" src="icons/icon-192.png" alt=""><h2>Tu contabilidad, en el móvil</h2>
      <p class="muted">Registra gastos, prepara el presupuesto y revisa tu balance. Todo se guarda en tu Excel de OneDrive.</p></div>
      <div class="card">
        ${cid ? "" : `<div class="field"><label for="cid">Client ID de tu app de Microsoft (ver guía)</label><input id="cid" placeholder="00000000-0000-0000-0000-000000000000"></div>`}
        <button class="btn" id="lg" type="button">Conectar con OneDrive</button>
        <div style="height:10px"></div>
        <button class="btn secondary" id="dm" type="button">Probar con datos de ejemplo</button>
      </div>
      <div class="card small"><b>Dirección de esta app</b> (la necesitas al registrarla en Microsoft):<br><code>${esc(Auth.redirectUri())}</code></div>`;
    $("#lg", v).onclick = () => {
      const inp = $("#cid", v);
      if (inp) { if (!inp.value.trim()) return toast("Pega primero el Client ID"); saveSettings({ clientId: inp.value.trim() }); }
      Auth.login().catch(e => toast(e.message, 5000));
    };
    $("#dm", v).onclick = () => { saveSettings({ mode: "demo" }); startDemo(); };
  }

  async function showFilePicker() {
    $$(".view").forEach(v => v.hidden = true);
    $("#tabbar").hidden = true; $("#fab").hidden = true;
    $("#viewTitle").textContent = "Elige tu Excel";
    const v = $("#view-login"); v.hidden = false;
    v.innerHTML = `<div class="card"><h3>Buscando en tu OneDrive…</h3><div class="spinner"></div></div>`;
    let who = "";
    try { const me = await Excel.GraphProvider.me(); who = me.userPrincipalName || me.displayName; } catch (e) { if (e instanceof Auth.AuthError) return showLogin(); }
    const [recent, found] = await Promise.all([
      Excel.GraphProvider.recent().catch(() => []),
      Excel.GraphProvider.searchFiles(window.APP_CONFIG.FILE_SEARCH || "Finanzas").catch(() => [])
    ]);
    v.innerHTML = `<p class="small muted" style="margin:0">Conectado como <b>${esc(who)}</b></p>
      <div class="card"><h3>Opción más fiable: pega el enlace</h3>
        <p class="small muted" style="margin-top:0">En OneDrive (web o app) toca los <b>⋯</b> del archivo → <b>Compartir</b> → <b>Copiar vínculo</b>, y pégalo aquí.</p>
        <div class="field"><input id="fl1" placeholder="https://1drv.ms/x/…  o  https://onedrive.live.com/…"></div>
        <button class="btn" id="flb" type="button">Usar este archivo</button></div>
      <div class="card"><h3>Abiertos recientemente</h3><div class="list" id="fr"></div></div>
      <div class="card"><h3>Resultados de búsqueda</h3><div class="list" id="fl"></div>
        <div class="field" style="margin-top:10px"><input id="fq" placeholder="Buscar por nombre"></div><button class="btn secondary" id="fs" type="button">Buscar</button></div>
      <div class="card"><h3>O escribe la ruta</h3><p class="xsmall muted" style="margin-top:0">Relativa a OneDrive, p. ej. <code>Documents/20260925 Finanzas Personales.xlsm</code>. También vale la ruta copiada del PC.</p>
        <div class="field"><input id="fp" placeholder="Carpeta/archivo.xlsm"></div><button class="btn secondary" id="fpb" type="button">Usar esta ruta</button></div>
      <button class="btn secondary" id="lo" type="button">Cerrar sesión</button>`;
    const fill = (box, list, emptyMsg) => {
      box.innerHTML = list.length ? "" : `<div class="empty">${emptyMsg}</div>`;
      list.slice().sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime))).forEach(f => {
        const b = document.createElement("button"); b.className = "row"; b.type = "button";
        const folder = (f.parentReference && f.parentReference.path || "").replace(/^\/drive(s\/[^/]+)?\/root:?/, "") || "OneDrive";
        b.innerHTML = `<span class="ico">📗</span><span class="mid"><div class="t">${esc(f.name)}</div><div class="s">${esc(folder)} · modificado ${new Date(f.lastModifiedDateTime).toLocaleString("es-ES", { dateStyle: "short", timeStyle: "short" })}</div></span>`;
        b.onclick = () => choose(f); box.appendChild(b);
      });
    };
    fill($("#fr", v), recent, "No hay Excel recientes.");
    fill($("#fl", v), found, "La búsqueda no encuentra nada (OneDrive puede tardar en indexar archivos nuevos).");
    $("#flb", v).onclick = async () => { const u = $("#fl1", v).value.trim(); if (!u) return toast("Pega el enlace"); try { choose(await Excel.GraphProvider.byShareLink(u)); } catch (e) { toast("No puedo abrir ese enlace: " + e.message, 6000); } };
    $("#fs", v).onclick = async () => { try { fill($("#fl", v), await Excel.GraphProvider.searchFiles($("#fq", v).value || "xls"), "Sin resultados."); } catch (e) { toast(e.message); } };
    $("#fpb", v).onclick = async () => { try { choose(await Excel.GraphProvider.byPath($("#fp", v).value)); } catch (e) { toast("No encuentro esa ruta: " + e.message, 6000); } };
    $("#lo", v).onclick = () => { Auth.logout(); showLogin(); };
    function choose(f) {
      if (!/\.xls[xm]$/i.test(f.name || "")) return toast("Eso no parece un Excel: " + (f.name || ""));
      saveSettings({ fileId: f.id, driveId: f.parentReference && f.parentReference.driveId || null, fileName: f.name, fileUrl: f.webUrl, mode: "graph" });
      LS.del("fin_cache"); LS.del("fin_queue");
      toast("Conectado a " + f.name);
      S.model = null; start();
    }
  }

  function openSettings() {
    const st = settings();
    const body = document.createElement("div");
    body.innerHTML = `
      <div class="card" style="margin-bottom:12px"><div class="small muted">Archivo</div><b>${esc(S.file ? S.file.name : "—")}</b>
        <div class="xsmall muted">${S.model ? `${S.model.tx.list.length} movimientos · leído ${new Date(S.model.loadedAt).toLocaleString("es-ES")}` : ""}</div>
        ${S.queue.length ? `<div class="xsmall" style="color:var(--warn)">${S.queue.length} cambio(s) pendiente(s) de subir</div>` : ""}</div>
      <button class="btn" id="sRef" type="button">Actualizar desde el Excel</button><div style="height:10px"></div>
      ${S.file && S.file.webUrl ? `<a class="btn secondary" style="display:block;text-align:center;text-decoration:none;box-sizing:border-box" href="${esc(S.file.webUrl)}" target="_blank" rel="noopener">Abrir el Excel en OneDrive</a><div style="height:10px"></div>` : ""}
      <label class="small" style="display:flex;gap:8px;align-items:center;margin-bottom:10px"><input type="checkbox" id="sAuto" ${st.autoSnap !== false ? "checked" : ""}> Congelar automáticamente el presupuesto inicial antes del primer cambio de cada mes</label>
      <label class="small" style="display:flex;gap:8px;align-items:center;margin-bottom:14px"><input type="checkbox" id="sAsk" ${st.askSavings !== false ? "checked" : ""}> Preguntarme si aparto ahorro cuando entra un ingreso no previsto</label>
      <div class="field"><label>Tema</label><select id="sTheme"><option value="auto">Automático</option><option value="light">Claro</option><option value="dark">Oscuro</option></select></div>
      ${S.mode === "graph" ? `<button class="btn secondary" id="sFile" type="button">Cambiar de archivo</button><div style="height:10px"></div>` : ""}
      <details style="margin-bottom:14px"><summary class="small muted">Conexión con Microsoft</summary>
        <div class="field" style="margin-top:10px"><label>Client ID</label><input id="sCid" value="${esc(Auth.cfg().clientId)}"></div>
        <div class="field"><label>Tipo de cuenta</label><select id="sAuth"><option value="consumers">Personal (Outlook, Hotmail…)</option><option value="common">Personal o de trabajo</option><option value="organizations">Trabajo / escuela</option></select></div>
        <p class="xsmall muted">Dirección de redirección: <code>${esc(Auth.redirectUri())}</code></p>
        <button class="btn secondary" id="sSaveC" type="button">Guardar</button></details>
      <button class="btn danger" id="sOut" type="button">${S.mode === "demo" ? "Salir del modo demo" : "Cerrar sesión"}</button>
      <p class="xsmall muted" style="text-align:center">Mis Finanzas · v1.8</p>`;
    $("#sTheme", body).value = st.theme || "auto";
    $("#sAsk", body).onchange = e => saveSettings({ askSavings: e.target.checked });
    $("#sAuto", body).onchange = e => saveSettings({ autoSnap: e.target.checked });
    $("#sTheme", body).onchange = e => { saveSettings({ theme: e.target.value }); applyTheme(); render(); };
    $("#sAuth", body).value = Auth.cfg().authority;
    $("#sRef", body).onclick = () => { closeSheet(); refresh(); };
    if ($("#sFile", body)) $("#sFile", body).onclick = () => { closeSheet(); saveSettings({ fileId: null, driveId: null }); showFilePicker(); };
    $("#sSaveC", body).onclick = () => { saveSettings({ clientId: $("#sCid", body).value.trim(), authority: $("#sAuth", body).value }); toast("Guardado"); };
    $("#sOut", body).onclick = () => {
      if (S.mode === "demo") { saveSettings({ mode: null }); }
      else { if (S.queue.length && !confirm("Hay cambios sin subir. ¿Cerrar sesión igualmente?")) return; Auth.logout(); LS.del("fin_cache"); LS.del("fin_queue"); saveSettings({ mode: null }); }
      closeSheet(); S.provider = null; S.model = null; showLogin();
    };
    openSheet("Ajustes", body);
  }

  start();
})();
