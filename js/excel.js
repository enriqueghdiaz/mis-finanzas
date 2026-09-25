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

  const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  // ── Proveedor Microsoft Graph ──────────────────────────────
  class GraphProvider {
    constructor(fileId) { this.fileId = fileId; this.session = null; this.kind = "graph"; }
    get base() { return `${GRAPH}/me/drive/items/${this.fileId}/workbook`; }
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
    tbl(name) { return `/tables('${encodeURIComponent(name)}')`; }
    async readTable(name) { const j = await this.req("GET", `${this.tbl(name)}/range?$select=values`); return j.values; }
    async addTableRow(name, index, row) { await this.req("POST", `${this.tbl(name)}/rows/add`, { index, values: [row] }); }
    async updateTableRow(name, index, row) { await this.req("PATCH", `${this.tbl(name)}/rows/itemAt(index=${index})`, { values: [row] }); }
    async deleteTableRow(name, index) { await this.req("DELETE", `${this.tbl(name)}/rows/itemAt(index=${index})`); }
    async info() {
      return GraphProvider.call("GET", `${GRAPH}/me/drive/items/${this.fileId}?$select=id,name,webUrl,lastModifiedDateTime`);
    }
    static async searchFiles(q) {
      const j = await GraphProvider.call("GET", `${GRAPH}/me/drive/root/search(q='${encodeURIComponent(q.replace(/'/g, "''"))}')?$select=id,name,webUrl,lastModifiedDateTime,parentReference&$top=50`);
      return (j.value || []).filter(f => /\.xls[xm]$/i.test(f.name));
    }
    static async byPath(path) {
      const clean = path.replace(/^\/+/, "").split("/").map(encodeURIComponent).join("/");
      return GraphProvider.call("GET", `${GRAPH}/me/drive/root:/${clean}?$select=id,name,webUrl,lastModifiedDateTime`);
    }
    static async me() { return GraphProvider.call("GET", `${GRAPH}/me?$select=displayName,userPrincipalName`); }
  }

  // ── Lectura de todo el libro ───────────────────────────────
  const RANGES = { budget: "B6:AN50", balance: "B8:R31", invest: "B2:F21" };

  async function loadRaw(p) {
    const S = window.APP_CONFIG.SHEETS;
    const [tx, budget, balance, invest] = await Promise.all([
      p.readTable(S.txTable),
      p.readRange(S.budget, RANGES.budget),
      p.readRange(S.balance, RANGES.balance),
      p.readRange(S.invest, RANGES.invest).catch(() => null)
    ]);
    return { tx, budget, balance, invest, loadedAt: new Date().toISOString() };
  }

  // ── Traducción a modelo ────────────────────────────────────
  const TX_COLS = ["Transacción", "Naturaleza", "Periodicidad", "Categoría", "Fecha", "Importe", "Notas"];

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
        notas: String(r[idx["Notas"]] == null ? "" : r[idx["Notas"]]).trim()
      });
    }
    return { head, idx, list };
  }

  function txToRow(head, t) {
    const map = {
      "Transacción": t.tipo === "Ahorro" ? SAVING.tipo : t.tipo,
      "Naturaleza": t.nat, "Periodicidad": t.per, "Categoría": t.cat,
      "Fecha": isoToSerial(t.fecha), "Importe": +t.importe, "Notas": t.notas || ""
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
    (V[0] || []).forEach((v, j) => {
      if (j > 0 && typeof v === "number" && v > 30000) {
        const iso = serialToISO(v);
        months.push({ key: iso.slice(0, 7), col: colLetter(C0 + j), j });
      }
    });
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
        savings = { row: R0 + i, vals: {} };
        months.forEach(m => savings.vals[m.key] = num(r[m.j]));
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

  function build(raw) {
    const tx = parseTx(raw.tx);
    const budget = parseBudget(raw.budget);
    const year = budget.months.length ? +budget.months[0].key.slice(0, 4) : new Date().getFullYear();
    return {
      tx, budget,
      balance: parseBalance(raw.balance, year),
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

  const ops = {
    async addTx(p, model, d) {
      await p.addTableRow(window.APP_CONFIG.SHEETS.txTable, 0, txToRow(model.tx.head, d.tx));
    },
    async updateTx(p, model, d) {
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
          await p.writeRange(S, colLetter(+c) + cells[0].row, { values: [[cells[0].value]] });
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
        if (it) it.vals[m.key] = c.value;
        if (model.budget.savings && model.budget.savings.row === a.r1) model.budget.savings.vals[m.key] = c.value;
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
    } else if (op.type === "setInvestAmount" && model.invest) {
      model.invest.amount = d.value;
    }
  }

  window.Excel = {
    GraphProvider, loadRaw, build, ops, applyLocal, investSplit, txToRow,
    CATS, CAT_INFO, SAVING, INCOME_CATS, EXPENSE_CATS, MESES,
    serialToISO, isoToSerial, colLetter, colNumber, parseAddr, sameTx
  };
})();
