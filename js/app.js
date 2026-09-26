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
  async function doOp(type, data, okMsg) {
    const op = { type, data, t: Date.now() };
    Excel.applyLocal(S.model, op);
    S.queue.push(op); saveQueue();
    render();
    if (okMsg) toast(okMsg);
    await flush();
    if (!S.queue.length) await refresh();
  }
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
    const sign = t.tipo === "Ingreso" ? "+" : t.tipo === "Ahorro" ? "→ " : "−";
    const cls = t.tipo === "Ingreso" ? "pos" : "";
    const d = t.fecha ? new Date(t.fecha + "T12:00").toLocaleDateString("es-ES", { day: "numeric", month: "short" }) : "";
    b.innerHTML = `<span class="ico">${icon(t.cat)}</span><span class="mid"><div class="t">${esc(t.notas || t.cat)}</div><div class="s">${esc(t.cat)} · ${d}${t.nat && t.nat !== "NA" && t.tipo === "Gasto" ? " · " + esc(t.nat) : ""}</div></span><span class="amt num ${cls}">${sign}${eur(t.importe)}</span>`;
    b.onclick = () => openTxForm(t);
    return b;
  }

  // ── MOVIMIENTOS ───────────────────────────────────────────
  function renderMovs(v) {
    const F = S.movFilter;
    v.innerHTML = `<input class="search" type="search" placeholder="Buscar por nota o categoría" value="${esc(F.q)}">
      <div class="chips">${["Todos", "Gasto", "Ingreso", "Ahorro"].map(x => `<button type="button" class="chip ${F.tipo === x ? "on" : ""}" data-t="${x}">${x === "Todos" ? "Todos" : x === "Gasto" ? "Gastos" : x === "Ingreso" ? "Ingresos" : "Ahorro"}</button>`).join("")}</div>
      <div class="card" style="padding-top:4px"><div class="list" id="movList"></div></div>`;
    const inp = $(".search", v);
    inp.oninput = () => { F.q = inp.value; F.limit = 60; drawList(); };
    $$(".chip", v).forEach(c => c.onclick = () => { F.tipo = c.dataset.t; F.limit = 60; $$(".chip", v).forEach(x => x.classList.toggle("on", x === c)); drawList(); });
    function drawList() {
      const q = F.q.trim().toLowerCase();
      let l = S.model.tx.list.filter(t => (F.tipo === "Todos" || t.tipo === F.tipo) && (!q || (t.notas + " " + t.cat).toLowerCase().includes(q)));
      l = l.slice().sort((a, b) => b.fecha.localeCompare(a.fecha));
      const box = $("#movList", v); box.innerHTML = "";
      if (!l.length) { box.innerHTML = `<div class="empty">No hay movimientos con ese filtro.</div>`; return; }
      let cur = null;
      l.slice(0, F.limit).forEach(t => {
        const m = t.fecha.slice(0, 7);
        if (m !== cur) {
          cur = m;
          const tot = l.filter(x => x.fecha.startsWith(m));
          const g = tot.filter(x => x.tipo === "Gasto").reduce((a, x) => a + x.importe, 0);
          const i = tot.filter(x => x.tipo === "Ingreso").reduce((a, x) => a + x.importe, 0);
          const h = document.createElement("div"); h.className = "group-h";
          h.innerHTML = `<span>${monthLabel(m)}</span><span class="num">${i ? "+" + eur(i) + " · " : ""}−${eur(g)}</span>`;
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
    const body = document.createElement("div");
    const draw = () => {
      const groups = f.tipo === "Ingreso" ? Excel.CATS.filter(g => g.tipo === "Ingreso") : f.tipo === "Gasto" ? Excel.CATS.filter(g => g.tipo === "Gasto") : [];
      body.innerHTML = `
        <div class="seg" style="margin-bottom:14px">${["Gasto", "Ingreso", "Ahorro"].map(x => `<button type="button" data-t="${x}" class="${f.tipo === x ? "on" : ""}">${x}</button>`).join("")}</div>
        <div class="field"><input id="fImp" class="amount-input num" type="text" inputmode="decimal" placeholder="0,00 €" value="${f.importe !== "" ? String(f.importe).replace(".", ",") : ""}" aria-label="Importe"></div>
        ${groups.length ? `<div class="field"><label>Categoría</label>${groups.map(g => `<div class="cat-sec">${g.tipo === "Gasto" ? g.nat + " · " + g.per : g.per}</div><div class="cat-grid">${g.cats.map(c => `<button type="button" class="chip ${f.cat === c ? "on" : ""}" data-c="${esc(c)}">${icon(c)} ${esc(c)}</button>`).join("")}</div>`).join("")}</div>` : `<p class="small muted" style="margin-top:0">Se registra como <b>Gasto · Esencial · Fijo · Ahorro</b>, igual que en tu Excel.</p>`}
        <div class="field"><label for="fFecha">Fecha (mes en el que cuenta)</label><input id="fFecha" type="date" value="${esc(f.fecha)}"></div>
        <div class="field"><label for="fNotas">Notas</label><input id="fNotas" type="text" list="noteSug" value="${esc(f.notas)}" placeholder="Ej.: Compra semanal"><datalist id="noteSug"></datalist></div>
        <button class="btn" id="fSave" type="button">${editing ? "Guardar cambios" : "Añadir"}</button>
        ${editing ? `<div style="height:10px"></div><button class="btn danger" id="fDel" type="button">Eliminar</button>` : ""}`;
      $$(".seg button", body).forEach(b => b.onclick = () => { keep(); f.tipo = b.dataset.t; if (f.tipo === "Ahorro") f.cat = "Ahorro"; else if (Excel.CAT_INFO[f.cat] && Excel.CAT_INFO[f.cat].tipo !== f.tipo) f.cat = ""; draw(); });
      $$(".cat-grid .chip", body).forEach(b => b.onclick = () => { keep(); f.cat = b.dataset.c; draw(); });
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
      const tx = { tipo: f.tipo, nat: info.nat, per: info.per, cat: f.cat, fecha: f.fecha, importe: round2(imp), notas: (f.notas || "").trim() };
      closeSheet();
      if (editing) await doOp("updateTx", { old: strip(t), tx }, "Cambios guardados");
      else await doOp("addTx", { tx }, "Movimiento añadido");
    }
    draw();
    openSheet(editing ? "Editar movimiento" : "Nuevo movimiento", body);
  }

  // ── PRESUPUESTO ───────────────────────────────────────────
  function budgetRows(k, tipo) {
    const b = S.model.budget, inB = b.months.some(m => m.key === k);
    const real = sumBy(txOfMonth(k).filter(t => t.tipo === tipo), t => t.cat);
    return b.items.filter(i => i.tipo === tipo).map(i => ({ cat: i.cat, nat: i.nat, per: i.per, row: i.row, budget: inB ? (i.vals[k] || 0) : null, real: real[i.cat] || 0 }));
  }
  function renderBudget(v) {
    const b = S.model.budget;
    const months = b.months.map(m => m.key);
    if (!months.length) { v.innerHTML = `<div class="empty">No encuentro columnas de meses en la hoja PPTO.</div>`; return; }
    if (!months.includes(S.month)) { const p = months.filter(x => x <= todayISO().slice(0, 7)); S.month = p.length ? p[p.length - 1] : months[0]; }
    const k = S.month, tipo = S.budgetSeg;
    v.innerHTML = "";
    v.appendChild(monthSwitch(months, render));
    const seg = document.createElement("div"); seg.className = "seg";
    seg.innerHTML = `<button type="button" class="${tipo === "Gasto" ? "on" : ""}" data-s="Gasto">Gastos</button><button type="button" class="${tipo === "Ingreso" ? "on" : ""}" data-s="Ingreso">Ingresos</button>`;
    $$("button", seg).forEach(x => x.onclick = () => { S.budgetSeg = x.dataset.s; render(); });
    v.appendChild(seg);

    const rows = budgetRows(k, tipo);
    const tb = rows.reduce((a, r) => a + r.budget, 0), tr = rows.reduce((a, r) => a + r.real, 0);
    const sum = document.createElement("div"); sum.className = "kpis";
    sum.innerHTML = `<div class="kpi"><span>Presupuestado</span><b class="num">${eur(tb)}</b></div>
      <div class="kpi"><span>${tipo === "Gasto" ? "Gastado" : "Recibido"}</span><b class="num">${eur(tr)}</b>
      <small class="${tipo === "Gasto" ? (tr > tb ? "neg" : "pos") : (tr >= tb ? "pos" : "muted")}">${tipo === "Gasto" ? (tr > tb ? "Exceso " + eur(tr - tb) : "Quedan " + eur(tb - tr)) : (tr >= tb ? "Objetivo cumplido" : "Faltan " + eur(tb - tr))}</small></div>`;
    v.appendChild(sum);

    const card = document.createElement("div"); card.className = "card";
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
        html += `<div class="brow" data-row="${r.row}" style="cursor:pointer"><div class="brow-top"><span class="name">${icon(r.cat)} ${esc(r.cat)} ${tag}</span><span class="vals num">${eur(r.real)} / <b>${eur(r.budget)}</b></span></div>
          <div class="prog ${cls}"><div style="width:${Math.min(100, p * 100)}%"></div></div></div>`;
      });
    });
    card.innerHTML = html;
    $$(".brow", card).forEach(el => el.onclick = () => editBudgetItem(+el.dataset.row, k));
    v.appendChild(card);

    const actions = document.createElement("div"); actions.className = "btn-row";
    actions.innerHTML = `<button class="btn secondary" type="button" id="bAll">Editar mes completo</button>`;
    $("#bAll", actions).onclick = () => editBudgetMonth(k);
    v.appendChild(actions);
    const note = document.createElement("p"); note.className = "xsmall muted"; note.style.margin = "0 4px";
    note.textContent = `Los cambios se escriben en la hoja PPTO (columna ${b.months.find(m => m.key === k).col}). Para presupuestar meses nuevos, añade sus columnas (con la fecha en la cabecera) en el Excel.`;
    v.appendChild(note);
  }

  function editBudgetItem(row, k) {
    const b = S.model.budget, it = b.items.find(i => i.row === row);
    const m = b.months.find(x => x.key === k);
    const real = txOfMonth(k).filter(t => t.cat === it.cat).reduce((a, t) => a + t.importe, 0);
    const past = txMonths().filter(x => x < k).slice(-3);
    const avg = past.length ? past.reduce((a, x) => a + txOfMonth(x).filter(t => t.cat === it.cat).reduce((s, t) => s + t.importe, 0), 0) / past.length : null;
    const later = b.months.filter(x => x.key > k);
    const body = document.createElement("div");
    body.innerHTML = `<p class="muted small" style="margin-top:0">${monthLabel(k)} · real hasta hoy: <b>${eur(real)}</b>${avg != null ? ` · media últimos ${past.length} meses: <b>${eur(avg)}</b>` : ""}</p>
      <div class="field"><label for="bVal">Presupuesto</label><input id="bVal" class="amount-input num" type="text" inputmode="decimal" value="${String(round2(it.vals[k] || 0)).replace(".", ",")}"></div>
      ${typeof it.f[k] === "string" && it.f[k][0] === "=" ? `<p class="xsmall muted">La celda tenía la fórmula <code>${esc(it.f[k])}</code>; se sustituirá por el importe.</p>` : ""}
      ${later.length ? `<label class="small" style="display:flex;gap:8px;align-items:center;margin-bottom:14px"><input type="checkbox" id="bLater"> Aplicar también a ${later.map(x => monthShort(x.key)).join(", ")}</label>` : ""}
      <button class="btn" id="bSave" type="button">Guardar</button>`;
    $("#bSave", body).onclick = async () => {
      const val = round2(parseFloat($("#bVal", body).value.replace(/\./g, "").replace(",", ".")) || 0);
      const cells = [{ addr: m.col + row, value: val }];
      if ($("#bLater", body) && $("#bLater", body).checked) later.forEach(x => cells.push({ addr: x.col + row, value: val }));
      closeSheet();
      await doOp("setBudgetCells", { cells }, "Presupuesto actualizado");
    };
    openSheet(`${icon(it.cat)} ${it.cat}`, body);
  }

  function editBudgetMonth(k) {
    const b = S.model.budget, m = b.months.find(x => x.key === k);
    const idx = b.months.indexOf(m), prevM = b.months[idx - 1];
    const body = document.createElement("div");
    const draw = (src) => {
      body.innerHTML = `${prevM ? `<button class="btn secondary" id="cp" type="button" style="margin-bottom:12px">Copiar de ${monthLabel(prevM.key)}</button>` : ""}
        ${["Ingreso", "Gasto"].map(tp => `<div class="bgroup-h"><h4>${tp === "Gasto" ? "Gastos" : "Ingresos"}</h4></div>` + b.items.filter(i => i.tipo === tp).map(i => `
          <div class="brow"><div class="brow-top" style="align-items:center"><span class="name">${icon(i.cat)} ${esc(i.cat)}</span>
          <input class="budget-edit num" data-row="${i.row}" inputmode="decimal" value="${String(round2((src || i.vals)[src ? i.row : k] || 0)).replace(".", ",")}"></div></div>`).join("")).join("")}
        <div style="height:12px"></div><button class="btn" id="save" type="button">Guardar ${monthLabel(k)}</button>`;
      if (prevM) $("#cp", body).onclick = () => { const src2 = {}; b.items.forEach(i => src2[i.row] = i.vals[prevM.key]); draw(src2); toast("Copiado. Revisa y guarda."); };
      $("#save", body).onclick = async () => {
        const cells = [];
        $$(".budget-edit", body).forEach(inp => {
          const row = +inp.dataset.row, it = b.items.find(i => i.row === row);
          const val = round2(parseFloat(inp.value.replace(/\./g, "").replace(",", ".")) || 0);
          if (Math.abs(val - (it.vals[k] || 0)) > 0.004) cells.push({ addr: m.col + row, value: val });
        });
        closeSheet();
        if (!cells.length) return toast("Sin cambios");
        await doOp("setBudgetCells", { cells }, `Guardadas ${cells.length} partidas`);
      };
    };
    draw(null);
    openSheet(`Presupuesto · ${monthLabel(k)}`, body);
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
    const rows = budgetRows(k, "Gasto").filter(r => r.real > 0 || (r.budget || 0) > 0).sort((a, b) => b.real - a.real).map(r => ({ label: r.cat, real: r.real, budget: r.budget }));
    const c2 = card("Real vs presupuesto", v, rows.some(r => r.budget != null) ? "La raya negra es el presupuesto; en rojo, lo que se pasa." : "Este mes no tiene presupuesto en PPTO.");
    requestAnimationFrame(() => Charts.bullet($(".ch", c2), rows));

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
    $("#sheet").hidden = false; document.body.style.overflow = "hidden";
    history.pushState({ sheet: 1 }, "");
  }
  function closeSheet() {
    if ($("#sheet").hidden) return;
    $("#sheet").hidden = true; document.body.style.overflow = "";
    if (history.state && history.state.sheet) history.back();
  }
  window.addEventListener("popstate", () => { if (!$("#sheet").hidden) { $("#sheet").hidden = true; document.body.style.overflow = ""; } });

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
      <div class="field"><label>Tema</label><select id="sTheme"><option value="auto">Automático</option><option value="light">Claro</option><option value="dark">Oscuro</option></select></div>
      ${S.mode === "graph" ? `<button class="btn secondary" id="sFile" type="button">Cambiar de archivo</button><div style="height:10px"></div>` : ""}
      <details style="margin-bottom:14px"><summary class="small muted">Conexión con Microsoft</summary>
        <div class="field" style="margin-top:10px"><label>Client ID</label><input id="sCid" value="${esc(Auth.cfg().clientId)}"></div>
        <div class="field"><label>Tipo de cuenta</label><select id="sAuth"><option value="consumers">Personal (Outlook, Hotmail…)</option><option value="common">Personal o de trabajo</option><option value="organizations">Trabajo / escuela</option></select></div>
        <p class="xsmall muted">Dirección de redirección: <code>${esc(Auth.redirectUri())}</code></p>
        <button class="btn secondary" id="sSaveC" type="button">Guardar</button></details>
      <button class="btn danger" id="sOut" type="button">${S.mode === "demo" ? "Salir del modo demo" : "Cerrar sesión"}</button>
      <p class="xsmall muted" style="text-align:center">Mis Finanzas · v1.3</p>`;
    $("#sTheme", body).value = st.theme || "auto";
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
