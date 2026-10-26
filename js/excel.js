/* Conexión con el Excel (Microsoft Graph) y traducción de sus hojas al modelo de la app. */
(function () {
  const GRAPH = "https://graph.microsoft.com/v1.0";

  // ── Utilidades de Excel ────────────────────────────────────
  const XL_EPOCH = Date.UTC(1899, 11, 30);
  function serialToISO(n) {
    if (typeof n === "string" && /^\d{4}-\d{2}-\d{2}/.test(n)) return n.slice(0, 10);
    if (typeof n !== "number" || !isFinite(n)) return null;
    return new Date(XL_EPOCH + Math.round(n) * 86400000).toISOString().slice(0, 10);
  }
  function isoToSerial(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - XL_EPOCH) / 86400000);
  }
  function colLetter(n) { // 1 -> A
    let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s;
  }
  function colNumber(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function parseAddr(a) {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
    const c1 = colNumber(m[1]), r1 = +m[2];
    return { c1, r1, c2: m[3] ? colNumber(m[3]) : c1, r2: m[4] ? +m[4] : r1 };
  }
  const num = v => (typeof v === "number" && isFinite(v)) ? v : (v === "" || v == null ? 0 : (isFinite(+v) ? +v : 0));

  // ── Calculadora: "8,25 + 144,09 − 30" ⇄ "=8.25+144.09-30" ──
  const Expr = (function () {
    function normNum(t) {
      if (t.includes(",")) return t.replace(/\./g, "").replace(",", ".");
      if (/^\d{1,3}(\.\d{3})+$/.test(t)) return t.replace(/\./g, "");
      return t;
    }
    function tokenize(str) {
      const s = String(str).replace(/[−–]/g, "-").replace(/[×x]/gi, "*").replace(/÷/g, "/").replace(/€/g, "").replace(/\s+/g, "");
      const toks = []; let i = 0;
      while (i < s.length) {
        const ch = s[i];
        if ("+-*/()".includes(ch)) { toks.push(ch); i++; continue; }
        const m = /^[\d.,]+/.exec(s.slice(i));
        if (!m) throw new Error("Carácter no válido: " + ch);
        const n = parseFloat(normNum(m[0]));
        if (!isFinite(n)) throw new Error("Número no válido: " + m[0]);
        toks.push(n); i += m[0].length;
      }
      return toks;
    }
    function evaluate(str) {
      const toks = tokenize(str); let pos = 0;
      if (!toks.length) return 0;
      const peek = () => toks[pos], next = () => toks[pos++];
      function expr() { let v = term(); while (peek() === "+" || peek() === "-") { const op = next(); const r = term(); v = op === "+" ? v + r : v - r; } return v; }
      function term() { let v = factor(); while (peek() === "*" || peek() === "/") { const op = next(); const r = factor(); v = op === "*" ? v * r : v / r; } return v; }
      function factor() {
        const t = next();
        if (t === "-") return -factor();
        if (t === "+") return factor();
        if (t === "(") { const v = expr(); if (next() !== ")") throw new Error("Falta un paréntesis"); return v; }
        if (typeof t === "number") return t;
        throw new Error("Expresión incompleta");
      }
      const v = expr();
      if (pos < toks.length) throw new Error("Expresión no válida");
      if (!isFinite(v)) throw new Error("Resultado no válido");
      return Math.round(v * 100) / 100;
    }
    function tryEval(str) { try { return { ok: true, value: evaluate(str) }; } catch (e) { return { ok: false, error: e.message }; } }
    function isPlainNumber(str) { const t = tokenize(str); return t.length === 1 && typeof t[0] === "number" || (t.length === 2 && t[0] === "-" && typeof t[1] === "number"); }
    // Expresión del usuario → contenido de celda (número o fórmula con punto decimal)
    function toCell(str) {
      const toks = tokenize(str);
      if (!toks.length) return 0;
      if (isPlainNumber(str)) return evaluate(str);
      const f = toks.map(t => typeof t === "number" ? String(t) : t).join("");
      return "=" + f.replace(/^\+/, "");
    }
    // Fórmula de celda solo con números (=8.25+144.09) → "8,25 + 144,09"; si no, null
    function fromFormula(f) {
      if (typeof f === "number") return fmt(f);
      if (typeof f !== "string" || f[0] !== "=") return null;
      const body = f.slice(1).replace(/\s+/g, "");
      if (!/^[+\-]?[\d.+\-*/()]+$/.test(body)) return null;
      return body.replace(/^\+/, "").replace(/\d+(\.\d+)?/g, n => n.replace(".", ",")).replace(/([+\-*/])/g, " $1 ").replace(/\s+/g, " ").replace(/^ - /, "-").trim();
    }
    // Divide en sumandos de primer nivel: "8,25 + 144,09 - 30" → ["8,25", "144,09", "-30"]
    function splitTerms(expr) {
      const out = []; let depth = 0, cur = "";
      const s = String(expr).replace(/[−–]/g, "-");
      for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (ch === "(") depth++;
        if (ch === ")") depth--;
        if (depth === 0 && (ch === "+" || ch === "-") && cur.trim() && !/[*/(]\s*$/.test(cur)) { out.push(cur.trim()); cur = ch === "-" ? "-" : ""; continue; }
        cur += ch;
      }
      if (cur.trim()) out.push(cur.trim());
      return out.map(t => t.replace(/^-\s+/, "-"));
    }
    function fmt(n) { return String(Math.round(n * 100) / 100).replace(".", ","); }
    return { evaluate, tryEval, toCell, fromFormula, splitTerms, fmt, isPlainNumber };
  })();

  // ── Estructura de categorías (hoja Registro) ──────────────
  const CATS = [
    { tipo: "Ingreso", nat: "NA", per: "Fijo", cats: ["Nómina", "Dietas"] },
    { tipo: "Ingreso", nat: "NA", per: "Variable", cats: ["Otros ingresos", "Financiación"] },
    { tipo: "Gasto", nat: "Esencial", per: "Fijo", cats: ["Vivienda", "Préstamos", "Facturas básicas", "Transporte", "Seguros obligatorios", "Medicamentos - r"] },
    { tipo: "Gasto", nat: "Esencial", per: "Variable", cats: ["Alimentación", "Medicamentos - nr", "Reparaciones", "Servicios médicos", "Educación"] },
    { tipo: "Gasto", nat: "Prescindible", per: "Fijo", cats: ["Suscripciones", "Cuotas", "Clases", "Donaciones"] },
    { tipo: "Gasto", nat: "Prescindible", per: "Variable", cats: ["Restaurantes", "Entretenimiento", "Viajes", "Ropa", "Caprichos", "Regalos", "Otros"] }
  ];
  // El traspaso a ahorro se registra como en tu Excel: Gasto · Esencial · Fijo · Ahorro
  const SAVING = { tipo: "Gasto", nat: "Esencial", per: "Fijo", cat: "Ahorro" };
  const CAT_INFO = {};
  CATS.forEach(g => g.cats.forEach(c => CAT_INFO[c] = { tipo: g.tipo, nat: g.nat, per: g.per }));
  CAT_INFO["Ahorro"] = { tipo: "Ahorro", nat: SAVING.nat, per: SAVING.per };
  const INCOME_CATS = CATS.filter(g => g.tipo === "Ingreso").flatMap(g => g.cats);
  const EXPENSE_CATS = CATS.filter(g => g.tipo === "Gasto").flatMap(g => g.cats);

  // ── Cuentas (configurables en config.js con ACCOUNTS) ──
  function accounts() {
    const c = window.APP_CONFIG.ACCOUNTS;
    return Array.isArray(c) && c.length ? c : [
      { name: "Openbank", principal: true },
      { name: "Edenred", gastoCats: ["Restaurantes"], ingresoCats: ["Dietas"], nota: "ticket restaurante: solo Restaurantes" }
    ];
  }
  const mainAccount = () => (accounts().find(a => a.principal) || accounts()[0]).name;
  // Cuenta de un movimiento: la columna «Cuenta» o, en los antiguos, la que aparezca en la nota
  function accountOf(t) {
    if (t.cuenta) return t.cuenta;
    const hit = accounts().find(a => !a.principal && new RegExp(a.name, "i").test(t.notas || ""));
    return hit ? hit.name : mainAccount();
  }

  const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  // ── Proveedor Microsoft Graph ──────────────────────────────
  class GraphProvider {
    constructor(fileId, driveId) { this.fileId = fileId; this.driveId = driveId || null; this.session = null; this.kind = "graph"; }
    get item() { return this.driveId ? `${GRAPH}/drives/${this.driveId}/items/${this.fileId}` : `${GRAPH}/me/drive/items/${this.fileId}`; }
    get base() { return this.item + "/workbook"; }
    static async call(method, url, body, extraHeaders) {
      for (let attempt = 0; attempt < 4; attempt++) {
        const token = await Auth.getToken();
        const r = await fetch(url, {
          method,
          headers: Object.assign({ Authorization: "Bearer " + token, "Content-Type": "application/json" }, extraHeaders || {}),
          body: body ? JSON.stringify(body) : undefined
        });
        if (r.status === 204) return null;
        if ([429, 503, 504].includes(r.status) && attempt < 3) {
          const wait = (+r.headers.get("Retry-After") || (1 + attempt * 2)) * 1000;
          await new Promise(res => setTimeout(res, wait)); continue;
        }
        const j = await r.json().catch(() => null);
        if (!r.ok) {
          const e = new Error((j && j.error && j.error.message) || ("Error " + r.status));
          e.status = r.status; e.code = j && j.error && (j.error.innerError && j.error.innerError.code || j.error.code);
          throw e;
        }
        return j;
      }
    }
    async req(method, path, body) {
      if (!this.session) await this.openSession();
      try {
        return await GraphProvider.call(method, this.base + path, body, { "workbook-session-id": this.session });
      } catch (e) {
        if (/session/i.test((e.code || "") + e.message) || e.status === 404 && /session/i.test(e.message)) {
          await this.openSession();
          return GraphProvider.call(method, this.base + path, body, { "workbook-session-id": this.session });
        }
        throw e;
      }
    }
    async openSession() {
      const j = await GraphProvider.call("POST", this.base + "/createSession", { persistChanges: true });
      this.session = j.id;
    }
    ws(sheet) { return `/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`; }
    async readRange(sheet, addr) {
      const j = await this.req("GET", `${this.ws(sheet)}/range(address='${addr}')?$select=values,formulas`);
      return { values: j.values, formulas: j.formulas };
    }
    async writeRange(sheet, addr, data) { await this.req("PATCH", `${this.ws(sheet)}/range(address='${addr}')`, data); }
    async addTableColumn(name, values) { await this.req("POST", `${this.tbl(name)}/columns/add`, { values }); }
    async addSheet(name) { await this.req("POST", "/worksheets/add", { name }); }
    tbl(name) { return `/tables('${encodeURIComponent(name)}')`; }
    async readTable(name) { const j = await this.req("GET", `${this.tbl(name)}/range?$select=values`); return j.values; }
    async addTableRow(name, index, row) { await this.req("POST", `${this.tbl(name)}/rows/add`, { index, values: [row] }); }
    async updateTableRow(name, index, row) { await this.req("PATCH", `${this.tbl(name)}/rows/itemAt(index=${index})`, { values: [row] }); }
    async deleteTableRow(name, index) { await this.req("DELETE", `${this.tbl(name)}/rows/itemAt(index=${index})`); }
    async info() {
      return GraphProvider.call("GET", `${this.item}?$select=id,name,webUrl,lastModifiedDateTime`);
    }
    static async searchFiles(q) {
      const j = await GraphProvider.call("GET", `${GRAPH}/me/drive/root/search(q='${encodeURIComponent(q.replace(/'/g, "''"))}')?$select=id,name,webUrl,lastModifiedDateTime,parentReference&$top=50`);
      return (j.value || []).filter(f => /\.xls[xm]$/i.test(f.name));
    }
    static async byPath(path) {
      // Admite rutas copiadas del PC (C:\Users\...\OneDrive\...) y la carpeta "Documentos" (en OneDrive se llama "Documents")
      let p = path.trim().replace(/^["']|["']$/g, "").replace(/\\/g, "/");
      const k = p.toLowerCase().lastIndexOf("onedrive/");
      if (k >= 0) p = p.slice(k + 9);
      p = p.replace(/^\/+/, "");
      const tryPath = q => GraphProvider.call("GET", `${GRAPH}/me/drive/root:/${q.split("/").map(encodeURIComponent).join("/")}?$select=id,name,webUrl,lastModifiedDateTime,parentReference`);
      try { return await tryPath(p); }
      catch (e) {
        if (/^documentos\//i.test(p)) return tryPath(p.replace(/^documentos\//i, "Documents/"));
        if (/^documents\//i.test(p)) return tryPath(p.replace(/^documents\//i, "Documentos/"));
        throw e;
      }
    }
    static async byShareLink(url) {
      const b = btoa(unescape(encodeURIComponent(url.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
      return GraphProvider.call("GET", `${GRAPH}/shares/u!${b}/driveItem?$select=id,name,webUrl,lastModifiedDateTime,parentReference`);
    }
    static async recent() {
      const j = await GraphProvider.call("GET", `${GRAPH}/me/drive/recent?$top=60`);
      return (j.value || []).map(it => {
        const r = it.remoteItem || it;
        return { id: r.id, name: it.name || r.name, webUrl: it.webUrl || r.webUrl, lastModifiedDateTime: it.lastModifiedDateTime || r.lastModifiedDateTime, parentReference: r.parentReference || it.parentReference };
      }).filter(f => /\.xls[xm]$/i.test(f.name || ""));
    }
    static async me() { return GraphProvider.call("GET", `${GRAPH}/me?$select=displayName,userPrincipalName`); }
  }

  // ── Lectura de todo el libro ───────────────────────────────
  const RANGES = { budget: "B6:AN50", balance: "B8:R31", invest: "B2:F21", notes: "A1:G600", snaps: "A1:H6000" };
  const SNAP_HEAD = ["Foto", "Mes", "Etiqueta", "Tipo", "Categoría", "Importe", "Cuentas", "Creada"];
  const snapsSheet = () => (window.APP_CONFIG.SHEETS.snaps = window.APP_CONFIG.SHEETS.snaps || "App_Fotos");
  const NOTES_HEAD = ["Clave", "Mes", "Categoría", "Nota", "Desglose", "Actualizado", "Cuentas"];
  const notesSheet = () => (window.APP_CONFIG.SHEETS.notes = window.APP_CONFIG.SHEETS.notes || "App_Notas");

  // Comprueba paso a paso el archivo para dar un error claro si algo no cuadra
  async function diagnose(p) {
    let item;
    try { item = await p.info(); }
    catch (e) {
      if (e instanceof Auth.AuthError) throw e;
      if (p.driveId) { p.driveId = null; try { item = await p.info(); } catch (e2) { item = null; } }
      if (!item) throw new Error("Paso 1 (archivo): OneDrive dice que el archivo no existe o no tienes acceso. Pulsa «Elegir otro archivo» y vuelve a seleccionarlo. Detalle: " + e.message);
    }
    try { await p.openSession(); }
    catch (e) {
      if (e instanceof Auth.AuthError) throw e;
      throw new Error(`Paso 2 (abrir el libro «${item.name}»): OneDrive no deja abrirlo con la API de Excel. Suele pasar con archivos .xlsm o muy pesados. Prueba a borrar la columna I de «BDD P&G» o a guardarlo como .xlsx. Detalle: ${e.message}`);
    }
    let sheets = [], tables = [];
    try { sheets = ((await p.req("GET", "/worksheets?$select=name")).value || []).map(x => x.name); } catch (e) { throw new Error("Paso 3 (leer hojas): " + e.message); }
    try { tables = ((await p.req("GET", "/tables?$select=name")).value || []).map(x => x.name); } catch (e) { throw new Error("Paso 3 (leer tablas): " + e.message); }
    return { item, sheets, tables };
  }

  async function loadRaw(p) {
    const S = window.APP_CONFIG.SHEETS;
    const warnings = [];
    let sheets = null, tables = null;
    if (p.kind === "graph") {
      const d = await diagnose(p); sheets = d.sheets; tables = d.tables;
      const tbl = tables.find(t => t.toLowerCase() === S.txTable.toLowerCase());
      if (!tbl) throw new Error(`En «${d.item.name}» no encuentro la tabla «${S.txTable}» (la de los movimientos). Tablas que hay: ${tables.join(", ") || "ninguna"}. ¿Es la versión correcta del archivo?`);
      S.txTable = tbl;
      notesSheet();
      const nh = sheets.find(n => n.toLowerCase() === S.notes.toLowerCase()); if (nh) S.notes = nh;
      snapsSheet(); const sh = sheets.find(n => n.toLowerCase() === S.snaps.toLowerCase()); if (sh) S.snaps = sh;
      ["budget", "balance", "invest"].forEach(k => {
        const hit = sheets.find(n => n.trim().toLowerCase() === S[k].trim().toLowerCase());
        if (hit) S[k] = hit; else warnings.push(S[k]);
      });
      if (warnings.includes(S.budget)) throw new Error(`En «${d.item.name}» no encuentro la hoja «${S.budget}». Hojas que hay: ${sheets.join(", ")}.`);
    }
    notesSheet();
    const opt = (k, rng) => (sheets && !sheets.includes(S[k])) ? Promise.resolve(null) : p.readRange(S[k], rng).catch(e => { warnings.push(S[k] + ": " + e.message); return null; });
    const tx = await p.readTable(S.txTable).catch(e => { throw new Error(`Leyendo la tabla «${S.txTable}»: ${e.message}`); });
    const budget = await p.readRange(S.budget, RANGES.budget).catch(e => { throw new Error(`Leyendo la hoja «${S.budget}» (${RANGES.budget}): ${e.message}`); });
    const [balance, invest] = await Promise.all([opt("balance", RANGES.balance), opt("invest", RANGES.invest)]);
    const notes = (sheets && !sheets.includes(S.notes)) ? null : await p.readRange(S.notes, RANGES.notes).catch(() => null);
    snapsSheet();
    const snaps = (sheets && !sheets.includes(S.snaps)) ? null : await p.readRange(S.snaps, RANGES.snaps).catch(() => null);
    return { tx, budget, balance, invest, notes, snaps, warnings, loadedAt: new Date().toISOString() };
  }

  // ── Traducción a modelo ────────────────────────────────────
  const TX_COLS = ["Transacción", "Naturaleza", "Periodicidad", "Categoría", "Fecha", "Importe", "Notas", "Cuenta"];

  function parseTx(values) {
    const head = (values[0] || []).map(h => String(h).trim());
    const idx = {}; TX_COLS.forEach(c => idx[c] = head.indexOf(c));
    const list = [];
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      if (!r || r.every(v => v === "" || v == null || (typeof v === "string" && !v.trim()))) continue;
      const cat = String(r[idx["Categoría"]] || "").trim();
      const tipoRaw = String(r[idx["Transacción"]] || "").trim();
      list.push({
        idx: i - 1,
        tipo: cat === "Ahorro" ? "Ahorro" : tipoRaw,
        tipoRaw,
        nat: String(r[idx["Naturaleza"]] || "").trim(),
        per: String(r[idx["Periodicidad"]] || "").trim(),
        cat,
        fecha: serialToISO(r[idx["Fecha"]]) || "",
        importe: num(r[idx["Importe"]]),
        notas: String(r[idx["Notas"]] == null ? "" : r[idx["Notas"]]).trim(),
        cuenta: idx["Cuenta"] >= 0 ? String(r[idx["Cuenta"]] == null ? "" : r[idx["Cuenta"]]).trim() : ""
      });
    }
    return { head, idx, list, accCol: idx["Cuenta"] >= 0 };
  }

  function txToRow(head, t) {
    const map = {
      "Transacción": t.tipo === "Ahorro" ? SAVING.tipo : t.tipo,
      "Naturaleza": t.nat, "Periodicidad": t.per, "Categoría": t.cat,
      "Fecha": isoToSerial(t.fecha), "Importe": +t.importe, "Notas": t.notas || "", "Cuenta": t.cuenta || ""
    };
    return head.map(h => (h in map ? map[h] : ""));
  }
  function sameTx(a, b) {
    return a.cat === b.cat && a.fecha === b.fecha && Math.abs(a.importe - b.importe) < 0.005 && (a.notas || "") === (b.notas || "") && a.tipoRaw === b.tipoRaw;
  }

  function isSubtotalFormula(f) {
    if (typeof f !== "string" || f[0] !== "=") return false;
    const g = f.replace(/\s+/g, "");
    return /^=\+?SUM\(/i.test(g) || /^=\+?-?\$?[A-Z]{1,2}\$?\d+([+-]\$?[A-Z]{1,2}\$?\d+)*$/.test(g);
  }

  function parseBudget(raw) {
    const V = raw.values, F = raw.formulas || raw.values, R0 = 6, C0 = 2; // B6
    const months = [];
    // Meses = cabeceras con fecha del bloque principal (se para en la primera columna vacía,
    // para no confundir otras tablas de la hoja, como «Foto», con meses)
    const head = V[0] || [];
    for (let j = 1; j < head.length; j++) {
      const v = head[j];
      if (v === "" || v == null) { if (months.length) break; else continue; }
      if (typeof v === "number" && v > 30000) {
        const key = serialToISO(v).slice(0, 7);
        if (!months.some(m => m.key === key)) months.push({ key, col: colLetter(C0 + j), j });
      }
    }
    let gi = V.findIndex(r => String(r[0]).trim().toUpperCase() === "GASTOS");
    if (gi < 0) gi = 13;
    const items = []; let savings = null, rate = null;
    const firstJ = months.length ? months[0].j : 1;
    const seen = new Set();
    V.forEach((r, i) => {
      if (i === 0 || i === gi) return;
      const label = String(r[0]).trim();
      const inIncome = i < gi;
      if (inIncome && typeof r[0] === "number" && r[0] > 0 && r[0] < 1) { rate = { value: r[0], addr: "B" + (R0 + i) }; return; }
      if (inIncome && label === "Ahorro") {
        savings = { row: R0 + i, vals: {}, f: {} };
        months.forEach(m => { savings.vals[m.key] = num(r[m.j]); savings.f[m.key] = F[i][m.j]; });
        return;
      }
      const list = inIncome ? INCOME_CATS : EXPENSE_CATS;
      if (!list.includes(label) || seen.has((inIncome ? "I" : "G") + label)) return;
      if (isSubtotalFormula(F[i][firstJ])) return;
      seen.add((inIncome ? "I" : "G") + label);
      const info = CAT_INFO[label];
      const it = { cat: label, row: R0 + i, tipo: inIncome ? "Ingreso" : "Gasto", nat: info.nat, per: info.per, vals: {}, f: {} };
      months.forEach(m => { it.vals[m.key] = num(r[m.j]); it.f[m.key] = F[i][m.j]; });
      items.push(it);
    });
    return { months, items, savings, rate, minRow: Math.min(...items.map(i => i.row)), maxRow: Math.max(...items.map(i => i.row)), formulas: F, R0, C0 };
  }

  function parseBalance(raw, year) {
    const V = raw.values, R0 = 8, C0 = 2;
    const monthNames = [];
    for (let j = 1; j <= 7; j++) monthNames.push(String(V[0][j] || "").trim());
    const months = monthNames.map((n, k) => {
      const mi = MESES.indexOf(n.toLowerCase());
      return { name: n, key: mi >= 0 ? `${year}-${String(mi + 1).padStart(2, "0")}` : null, col: colLetter(C0 + 1 + k), j: 1 + k };
    });
    const sections = [];
    for (let i = 0; i < V.length; i++) {
      const second = String(V[i][1] || "").trim();
      if (monthNames.includes(second) && String(V[i][0]).trim()) {
        const sec = { name: String(V[i][0]).trim(), side: "activo", items: [] };
        for (let k = i + 1; k < V.length; k++) {
          const lab = String(V[k][0] || "").trim();
          if (/^total/i.test(lab)) break;
          if (!lab) continue;
          sec.items.push({ label: lab, row: R0 + k, vals: months.map(m => V[k][m.j] === "" ? null : num(V[k][m.j])) });
        }
        sections.push(sec);
      }
    }
    // Pasivos (bloque K:R)
    const kj = colNumber("K") - C0;
    if (V[0] && String(V[0][kj] || "").trim()) {
      const sec = { name: String(V[0][kj]).trim(), side: "pasivo", items: [] };
      for (let k = 1; k < V.length; k++) {
        const lab = String(V[k][kj] || "").trim();
        if (/^total/i.test(lab)) break;
        const hasVals = months.some((m, x) => V[k][kj + 1 + x] !== "" && V[k][kj + 1 + x] != null);
        if (!lab && !hasVals && k > 1) continue;
        sec.items.push({ label: lab || "Préstamo", row: R0 + k, vals: months.map((m, x) => { const v = V[k][kj + 1 + x]; return v === "" || v == null ? null : num(v); }) });
      }
      sec.cols = months.map((m, x) => colLetter(C0 + kj + 1 + x));
      sections.push(sec);
    }
    return { months, sections };
  }

  function parseInvest(raw) {
    if (!raw) return null;
    const V = raw.values, R0 = 2;
    const out = { amount: num(V[0][1]), amountAddr: "C2", cash: null, funds: [] };
    let kind = "";
    for (let i = 1; i < V.length; i++) {
      const lab = String(V[i][0] || "").trim();
      if (/renta variable/i.test(lab)) kind = "Renta variable";
      else if (/renta fija/i.test(lab)) kind = "Renta fija";
      else if (/cuenta/i.test(lab) && typeof V[i][1] === "number") out.cash = { name: lab, pct: V[i][1], row: R0 + i };
      else if (lab && lab !== "Nombre" && typeof V[i][1] === "number" && !/fondos|efectivo/i.test(lab)) {
        out.funds.push({ name: lab, pct: V[i][1], isin: String(V[i][2] || ""), kind, row: R0 + i });
      }
    }
    return out;
  }

  function investSplit(plan, amount) {
    if (!plan) return [];
    const cashPct = plan.cash ? plan.cash.pct : 0;
    const cash = Math.round(amount * cashPct);
    const rest = amount - cash;
    const rows = [];
    if (plan.cash) rows.push({ name: plan.cash.name, kind: "Efectivo", pct: cashPct, eur: cash });
    plan.funds.forEach(f => rows.push({ name: f.name, isin: f.isin, kind: f.kind, pct: f.pct, eur: Math.round(rest * (f.pct / (1 - cashPct || 1))) }));
    return rows;
  }

  function parseNotes(raw) {
    const out = {};
    if (!raw || !raw.values) return out;
    raw.values.forEach((r, i) => {
      if (i === 0 || !r[0]) return;
      let des = [];
      try { des = r[4] ? JSON.parse(r[4]) : []; } catch (e) { des = []; }
      let cu = {};
      try { cu = r[6] ? JSON.parse(r[6]) : {}; } catch (e) { cu = {}; }
      out[String(r[0])] = { mes: String(r[1] || ""), cat: String(r[2] || ""), nota: String(r[3] || ""), desglose: Array.isArray(des) ? des : [], cuentas: cu && typeof cu === "object" ? cu : {}, row: i + 1 };
    });
    return out;
  }
  const noteKey = (mes, cat) => `${mes}|${cat}`;

  // Fotos del presupuesto: una fila por partida, agrupadas por la columna «Foto»
  function parseSnaps(raw) {
    const out = [], byId = {};
    if (!raw || !raw.values) return out;
    raw.values.forEach((r, i) => {
      if (i === 0 || !r[0] || !r[4]) return;
      const id = String(r[0]);
      let s = byId[id];
      if (!s) { s = byId[id] = { id, mes: String(r[1] || ""), label: String(r[2] || ""), creada: String(r[7] || id), items: {} }; out.push(s); }
      let cu = {};
      try { cu = r[6] ? JSON.parse(r[6]) : {}; } catch (e) { cu = {}; }
      s.items[String(r[4])] = { tipo: String(r[3] || ""), v: typeof r[5] === "number" ? r[5] : (parseFloat(r[5]) || 0), cuentas: cu || {} };
    });
    return out.sort((a, b) => a.id.localeCompare(b.id));
  }

  function build(raw) {
    const tx = parseTx(raw.tx);
    const budget = parseBudget(raw.budget);
    const year = budget.months.length ? +budget.months[0].key.slice(0, 4) : new Date().getFullYear();
    return {
      tx, budget,
      balance: raw.balance ? parseBalance(raw.balance, year) : { months: [], sections: [] },
      warnings: raw.warnings || [],
      notes: parseNotes(raw.notes),
      snaps: parseSnaps(raw.snaps),
      invest: parseInvest(raw.invest),
      loadedAt: raw.loadedAt
    };
  }

  // ── Escrituras ─────────────────────────────────────────────
  async function findTxIndex(p, head, t) {
    const values = await p.readTable(window.APP_CONFIG.SHEETS.txTable);
    const parsed = parseTx(values);
    const hit = parsed.list.find(x => sameTx(x, t));
    if (!hit) throw new Error("No encuentro ese movimiento en el Excel (¿lo cambiaste desde el ordenador?). Actualiza y vuelve a intentarlo.");
    return hit.idx;
  }

  async function ensureAccountCol(p, model, t) {
    if (!t || !t.cuenta || model.tx.accCol) return;
    const T = window.APP_CONFIG.SHEETS.txTable;
    const vals = await p.readTable(T);
    const head = (vals[0] || []).map(h => String(h).trim());
    if (!head.includes("Cuenta")) {
      if (!p.addTableColumn) return;
      await p.addTableColumn(T, [["Cuenta"]].concat(vals.slice(1).map(() => [""])));
      head.push("Cuenta");
    }
    model.tx.head = head; model.tx.accCol = true;
  }

  const ops = {
    async addTx(p, model, d) {
      await ensureAccountCol(p, model, d.tx);
      await p.addTableRow(window.APP_CONFIG.SHEETS.txTable, 0, txToRow(model.tx.head, d.tx));
    },
    async updateTx(p, model, d) {
      await ensureAccountCol(p, model, d.tx);
      const i = await findTxIndex(p, model.tx.head, d.old);
      await p.updateTableRow(window.APP_CONFIG.SHEETS.txTable, i, txToRow(model.tx.head, d.tx));
    },
    async deleteTx(p, model, d) {
      const i = await findTxIndex(p, model.tx.head, d.old);
      await p.deleteTableRow(window.APP_CONFIG.SHEETS.txTable, i);
    },
    async setBudgetCells(p, model, d) { // d.cells: [{addr, value}]
      const S = window.APP_CONFIG.SHEETS.budget;
      // Agrupa por columna en un solo rango conservando las fórmulas del resto de filas
      const byCol = {};
      d.cells.forEach(c => { const a = parseAddr(c.addr); (byCol[a.c1] = byCol[a.c1] || []).push({ row: a.r1, value: c.value }); });
      for (const c of Object.keys(byCol)) {
        const cells = byCol[c];
        if (cells.length === 1) {
          await p.writeRange(S, colLetter(+c) + cells[0].row, { formulas: [[cells[0].value]] });
          continue;
        }
        const r1 = Math.min(...cells.map(x => x.row)), r2 = Math.max(...cells.map(x => x.row));
        const cur = await p.readRange(S, `${colLetter(+c)}${r1}:${colLetter(+c)}${r2}`);
        const f = cur.formulas.map(r => [r[0]]);
        cells.forEach(x => f[x.row - r1][0] = x.value);
        await p.writeRange(S, `${colLetter(+c)}${r1}:${colLetter(+c)}${r2}`, { formulas: f });
      }
    },
    async setBalance(p, model, d) { // d.cells: [{addr, value}]
      const S = window.APP_CONFIG.SHEETS.balance;
      for (const c of d.cells) await p.writeRange(S, c.addr, { values: [[c.value === null ? "" : c.value]] });
    },
    async setNotes(p, model, d) { // d.notes: [{mes, cat, nota, desglose}]
      const S = notesSheet();
      let cur;
      try { cur = await p.readRange(S, RANGES.notes); }
      catch (e) {
        if (!p.addSheet) throw e;
        await p.addSheet(S);
        cur = { values: [] };
      }
      const vals = cur.values || [];
      if (!vals.length || String(vals[0][0] || "") !== "Clave" || String(vals[0][6] || "") !== "Cuentas") await p.writeRange(S, "A1:G1", { values: [NOTES_HEAD] });
      const keys = vals.map(r => String(r[0] || ""));
      let nextFree = Math.max(2, keys.reduce((last, k, i) => k ? i + 1 : last, 1) + 1);
      for (const n of d.notes) {
        const key = noteKey(n.mes, n.cat);
        let row = keys.indexOf(key) + 1;
        if (row < 2) { row = nextFree++; keys[row - 1] = key; }
        const des = (n.desglose || []).filter(x => x && (x.e || x.l));
        // Si no se indica el reparto por cuentas, se conserva el que hubiera
        const prevCu = vals[row - 1] ? String(vals[row - 1][6] || "") : "";
        const cu = n.cuentas === undefined ? prevCu : (n.cuentas && Object.keys(n.cuentas).length ? JSON.stringify(n.cuentas) : "");
        await p.writeRange(S, `A${row}:G${row}`, { values: [[key, n.mes, n.cat, n.nota || "", des.length ? JSON.stringify(des) : "", new Date().toISOString().slice(0, 16).replace("T", " "), cu]] });
      }
    },
    async addSnapshot(p, model, d) { // d: {id, mes, label, rows:[{tipo, cat, v, cuentas}]}
      const S = snapsSheet();
      let cur;
      try { cur = await p.readRange(S, RANGES.snaps); }
      catch (e) { if (!p.addSheet) throw e; await p.addSheet(S); cur = { values: [] }; }
      const vals = cur.values || [];
      if (!vals.length || String(vals[0][0] || "") !== "Foto") await p.writeRange(S, "A1:H1", { values: [SNAP_HEAD] });
      const last = vals.reduce((l, r, i) => (r[0] ? i + 1 : l), 1);
      const r1 = Math.max(2, last + 1), r2 = r1 + d.rows.length - 1;
      const creada = String(d.id).split(" ").slice(0, 2).join(" ");
      await p.writeRange(S, `A${r1}:H${r2}`, { values: d.rows.map(x => [d.id, d.mes, d.label || "", x.tipo, x.cat, x.v, x.cuentas && Object.keys(x.cuentas).length ? JSON.stringify(x.cuentas) : "", creada]) });
    },
    async deleteSnapshot(p, model, d) {
      const S = snapsSheet();
      const cur = await p.readRange(S, RANGES.snaps);
      const rows = []; (cur.values || []).forEach((r, i) => { if (String(r[0]) === d.id) rows.push(i + 1); });
      if (!rows.length) return;
      // borra por bloques consecutivos
      let a = rows[0], prev = rows[0];
      const flushBlock = async (x, y) => p.writeRange(S, `A${x}:H${y}`, { values: Array.from({ length: y - x + 1 }, () => ["", "", "", "", "", "", "", ""]) });
      for (const r of rows.slice(1).concat([null])) {
        if (r === prev + 1) { prev = r; continue; }
        await flushBlock(a, prev); if (r != null) { a = r; prev = r; }
      }
    },
    async setInvestAmount(p, model, d) {
      await p.writeRange(window.APP_CONFIG.SHEETS.invest, "C2", { values: [[d.value]] });
    }
  };

  // Aplica la operación también en local (para verla al instante y sin conexión)
  function applyLocal(model, op) {
    const d = op.data;
    if (op.type === "addTx") {
      model.tx.list.unshift(Object.assign({ idx: -1, pending: true, tipoRaw: d.tx.tipo === "Ahorro" ? SAVING.tipo : d.tx.tipo }, d.tx));
    } else if (op.type === "updateTx") {
      const i = model.tx.list.findIndex(x => sameTx(x, d.old));
      if (i >= 0) model.tx.list[i] = Object.assign({}, model.tx.list[i], d.tx, { pending: true, tipoRaw: d.tx.tipo === "Ahorro" ? SAVING.tipo : d.tx.tipo });
    } else if (op.type === "deleteTx") {
      const i = model.tx.list.findIndex(x => sameTx(x, d.old));
      if (i >= 0) model.tx.list.splice(i, 1);
    } else if (op.type === "setBudgetCells") {
      d.cells.forEach(c => {
        const a = parseAddr(c.addr);
        if (model.budget.rate && c.addr === model.budget.rate.addr) { model.budget.rate.value = c.value; return; }
        const m = model.budget.months.find(m => colNumber(m.col) === a.c1);
        if (!m) return;
        const it = model.budget.items.find(i => i.row === a.r1);
        let val = c.value;
        if (typeof c.value === "string") {
          const r = Expr.tryEval(c.value.replace(/^=/, ""));
          if (r.ok) val = r.value;
          else if (model.budget.rate && c.value.replace(/\$/g, "").includes(model.budget.rate.addr)) // =ingresos × tasa
            val = Math.round(model.budget.items.filter(i => i.tipo === "Ingreso").reduce((x, i) => x + (i.vals[m.key] || 0), 0) * model.budget.rate.value * 100) / 100;
          else val = (it ? it.vals[m.key] : model.budget.savings && model.budget.savings.vals[m.key]) || 0;
        }
        if (it) { it.vals[m.key] = val; it.f[m.key] = c.value; }
        if (model.budget.savings && model.budget.savings.row === a.r1) { model.budget.savings.vals[m.key] = val; model.budget.savings.f[m.key] = c.value; }
      });
    } else if (op.type === "setBalance") {
      d.cells.forEach(c => {
        const a = parseAddr(c.addr);
        model.balance.sections.forEach(s => s.items.forEach(it => {
          if (it.row !== a.r1) return;
          if (s.side === "pasivo") { const x = s.cols.indexOf(colLetter(a.c1)); if (x >= 0) it.vals[x] = c.value; }
          else { const x = model.balance.months.findIndex(m => m.col === colLetter(a.c1)); if (x >= 0) it.vals[x] = c.value; }
        }));
      });
    } else if (op.type === "setNotes") {
      model.notes = model.notes || {};
      d.notes.forEach(n => {
        const old = model.notes[noteKey(n.mes, n.cat)];
        model.notes[noteKey(n.mes, n.cat)] = { mes: n.mes, cat: n.cat, nota: n.nota || "", desglose: (n.desglose || []).filter(x => x && (x.e || x.l)), cuentas: n.cuentas === undefined ? (old ? old.cuentas || {} : {}) : (n.cuentas || {}) };
      });
    } else if (op.type === "addSnapshot") {
      model.snaps = (model.snaps || []).filter(x => x.id !== d.id);
      const items = {}; d.rows.forEach(x => items[x.cat] = { tipo: x.tipo, v: x.v, cuentas: x.cuentas || {} });
      model.snaps.push({ id: d.id, mes: d.mes, label: d.label || "", creada: d.id, items, pending: true });
      model.snaps.sort((a, b) => a.id.localeCompare(b.id));
    } else if (op.type === "deleteSnapshot") {
      model.snaps = (model.snaps || []).filter(x => x.id !== d.id);
    } else if (op.type === "setInvestAmount" && model.invest) {
      model.invest.amount = d.value;
    }
  }

  window.Excel = {
    accounts, mainAccount, accountOf,
    GraphProvider, loadRaw, build, ops, applyLocal, investSplit, txToRow, Expr, noteKey,
    CATS, CAT_INFO, SAVING, INCOME_CATS, EXPENSE_CATS, MESES,
    serialToISO, isoToSerial, colLetter, colNumber, parseAddr, sameTx
  };
})();
