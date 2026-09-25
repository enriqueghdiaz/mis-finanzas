/* Modo demo: un libro ficticio en memoria con la misma estructura que tu Excel.
   Sirve para probar la app antes de conectarla a OneDrive. Nada se guarda en la nube. */
(function () {
  const { colLetter, colNumber, parseAddr, isoToSerial } = window.Excel;

  class DemoProvider {
    constructor() { this.kind = "demo"; this.wb = makeDemoWorkbook(); }
    grid(sheet) { return this.wb.sheets[sheet] || (this.wb.sheets[sheet] = {}); }
    async readRange(sheet, addr) {
      const g = this.grid(sheet), a = parseAddr(addr), values = [], formulas = [];
      for (let r = a.r1; r <= a.r2; r++) {
        const vr = [], fr = [];
        for (let c = a.c1; c <= a.c2; c++) {
          const cell = g[colLetter(c) + r];
          const v = cell == null ? "" : (typeof cell === "object" ? cell.v : cell);
          vr.push(v); fr.push(cell && typeof cell === "object" ? cell.f : v);
        }
        values.push(vr); formulas.push(fr);
      }
      return { values, formulas };
    }
    async writeRange(sheet, addr, data) {
      const g = this.grid(sheet), a = parseAddr(addr), arr = data.values || data.formulas;
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
        const v = arr[r - a.r1][c - a.c1];
        if (typeof v === "string" && v[0] === "=") { const old = g[colLetter(c) + r]; g[colLetter(c) + r] = { f: v, v: old && typeof old === "object" ? old.v : 0 }; }
        else g[colLetter(c) + r] = v;
      }
    }
    async readTable() { return [this.wb.txHead].concat(this.wb.tx.map(r => r.slice())); }
    async addTableRow(n, index, row) { this.wb.tx.splice(index, 0, row); }
    async updateTableRow(n, index, row) { this.wb.tx[index] = row; }
    async deleteTableRow(n, index) { this.wb.tx.splice(index, 1); }
    async info() { return { name: "Demo (datos ficticios)", webUrl: null, lastModifiedDateTime: new Date().toISOString() }; }
  }

  function rng(seed) { return () => { seed = (seed * 16807) % 2147483647; return (seed - 1) / 2147483646; }; }

  function makeDemoWorkbook() {
    const rand = rng(42);
    const sheets = { PPTO: {}, BCE: {}, "Ahorro - Inversión": {} };
    const P = sheets.PPTO;
    const months = ["2026-07-01", "2026-08-01", "2026-09-01", "2026-10-01", "2026-11-01", "2026-12-01"];
    const cols = ["C", "D", "E", "F", "G", "H"];
    P.B6 = "INGRESOS"; P.B19 = "GASTOS";
    cols.forEach((c, i) => { P[c + "6"] = isoToSerial(months[i]); P[c + "19"] = isoToSerial(months[i]); });
    const sumRow = (row, label, f) => { P["B" + row] = label; cols.forEach(c => P[c + row] = { f: f.replace(/#/g, c), v: 0 }); };
    const catRow = (row, label, vals) => { P["B" + row] = label; cols.forEach((c, i) => P[c + row] = vals[i]); };
    sumRow(7, "Fijos", "=+#8"); sumRow(8, "Nómina", "=+SUM(#9:#10)");
    catRow(9, "Nómina", [1850, 1850, 1850, 1850, 1850, 1850]);
    catRow(10, "Dietas", [180, 180, 180, 180, 180, 180]);
    sumRow(11, "Variables", "=+#12"); sumRow(12, "Otros", "=+SUM(#13:#14)");
    catRow(13, "Otros ingresos", [50, 0, 60, 0, 0, 0]);
    catRow(14, "Financiación", [0, 0, 0, 0, 0, 0]);
    sumRow(15, "Totales", "=+#7+#11");
    catRow(16, "Ahorro", [520, 508, 522, 508, 508, 508]);
    P.B17 = 0.25;
    sumRow(20, "Esenciales", "=+#21+#28"); sumRow(21, "Fijos", "=+SUM(#22:#27)");
    const budgetExp = {
      22: ["Vivienda", 450], 23: ["Préstamos", 0], 24: ["Facturas básicas", 60], 25: ["Transporte", 70], 26: ["Seguros obligatorios", 35], 27: ["Medicamentos - r", 5],
      29: ["Alimentación", 220], 30: ["Medicamentos - nr", 10], 31: ["Reparaciones", 20], 32: ["Servicios médicos", 0], 33: ["Educación", 15],
      36: ["Suscripciones", 25], 37: ["Cuotas", 40], 38: ["Clases", 0], 39: ["Donaciones", 10],
      41: ["Restaurantes", 180], 42: ["Entretenimiento", 30], 43: ["Viajes", 100], 44: ["Ropa", 40], 45: ["Caprichos", 20], 46: ["Regalos", 30], 47: ["Otros", 40]
    };
    sumRow(28, "Variables", "=+SUM(#29:#33)"); sumRow(34, "Prescindibles", "=+#35+#40"); sumRow(35, "Fijos", "=+SUM(#36:#39)");
    sumRow(40, "Variables", "=+SUM(#41:#47)"); sumRow(48, "Totales", "=+#20+#34");
    Object.entries(budgetExp).forEach(([row, [label, v]]) => catRow(+row, label, cols.map(() => v)));

    // Movimientos ficticios (julio–septiembre)
    const txHead = ["Transacción", "Naturaleza", "Periodicidad", "Categoría", "Fecha", "Importe", "Notas"];
    const tx = [];
    const add = (tipo, nat, per, cat, iso, imp, notas) => tx.push([tipo, nat, per, cat, isoToSerial(iso), Math.round(imp * 100) / 100, notas]);
    const plan = [
      ["Gasto", "Esencial", "Variable", "Alimentación", ["Supermercado", "Mercado", "Panadería", "Frutería"], 6, 8, 45],
      ["Gasto", "Prescindible", "Variable", "Restaurantes", ["Cena con amigos", "Café", "Menú del día", "Pizza"], 7, 6, 38],
      ["Gasto", "Esencial", "Fijo", "Transporte", ["Gasolina", "Parking"], 3, 12, 28],
      ["Gasto", "Prescindible", "Variable", "Entretenimiento", ["Cine", "Concierto"], 1, 8, 30],
      ["Gasto", "Prescindible", "Variable", "Caprichos", ["Libro", "Revista"], 1, 5, 20],
      ["Gasto", "Prescindible", "Variable", "Otros", ["Ferretería", "Farmacia"], 1, 8, 30]
    ];
    ["2026-07", "2026-08", "2026-09"].forEach((m, mi) => {
      const last = mi === 2 ? 25 : 28;
      add("Ingreso", "NA", "Fijo", "Nómina", m + "-01", 1850, "Nómina");
      add("Ingreso", "NA", "Fijo", "Dietas", m + "-01", 180, "Tickets comida");
      add("Gasto", "Esencial", "Fijo", "Ahorro", m + "-01", [520, 508, 522][mi], "Traspaso a inversión");
      add("Gasto", "Esencial", "Fijo", "Vivienda", m + "-02", 450, "Alquiler");
      add("Gasto", "Esencial", "Fijo", "Facturas básicas", m + "-05", 48 + rand() * 20, "Luz y agua");
      add("Gasto", "Esencial", "Fijo", "Seguros obligatorios", m + "-06", 35, "Seguro coche");
      add("Gasto", "Prescindible", "Fijo", "Suscripciones", m + "-03", 12.99, "Streaming");
      add("Gasto", "Prescindible", "Fijo", "Suscripciones", m + "-03", 1.99, "Nube");
      add("Gasto", "Prescindible", "Fijo", "Cuotas", m + "-04", 40, "Gimnasio");
      plan.forEach(([t, n, p, c, notes, times, lo, hi]) => {
        for (let k = 0; k < times; k++) {
          const d = String(1 + Math.floor(rand() * last)).padStart(2, "0");
          add(t, n, p, c, `${m}-${d}`, lo + rand() * (hi - lo), notes[k % notes.length]);
        }
      });
      if (mi === 1) { add("Gasto", "Prescindible", "Variable", "Viajes", m + "-12", 285.4, "Escapada fin de semana"); add("Gasto", "Prescindible", "Variable", "Ropa", m + "-18", 59.9, "Zapatillas"); }
      if (mi === 0) add("Ingreso", "NA", "Variable", "Otros ingresos", m + "-15", 50, "Venta segunda mano");
      if (mi === 2) add("Gasto", "Esencial", "Variable", "Educación", m + "-10", 15, "Curso online");
    });
    tx.sort((a, b) => b[4] - a[4]);

    // Balance
    const B = sheets.BCE;
    const mNames = ["Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    const lc = ["C", "D", "E", "F", "G", "H", "I"], rc = ["L", "M", "N", "O", "P", "Q", "R"];
    const header = (row, label) => { B["B" + row] = label; lc.forEach((c, i) => B[c + row] = mNames[i]); };
    header(8, "Motos"); B.K8 = "Préstamos"; rc.forEach((c, i) => B[c + "8"] = mNames[i]);
    B.B9 = "Coche"; B.C9 = 8200; B.D9 = 8100; B.E9 = 8000; B.B10 = "Bicicleta"; B.C10 = 600; B.D10 = 590; B.E10 = 580; B.B11 = "Total";
    B.K9 = "Préstamo coche"; B.L9 = 2400; B.M9 = 2200; B.N9 = 2000; B.K31 = "Total Pasivos";
    header(13, "Inversiones");
    const inv = [["Fondo indexado mundial", [1500, 1620, 1740]], ["Fondo bonos gobierno", [400, 450, 500]], ["Fondo bonos corporativos", [390, 440, 492]], ["Oro", [300, 310, 322]]];
    inv.forEach(([n, v], i) => { B["B" + (14 + i)] = n; v.forEach((x, k) => B[lc[k] + (14 + i)] = x); });
    B.B19 = "Total";
    header(21, "Cash & Receivables");
    const cash = [["Cuenta corriente", [820, 760, 910]], ["Cuenta remunerada", [1200, 1400, 1605]], ["Efectivo", [80, 60, 120]]];
    cash.forEach(([n, v], i) => { B["B" + (22 + i)] = n; v.forEach((x, k) => B[lc[k] + (22 + i)] = x); });
    B.B29 = "Total";

    const A = sheets["Ahorro - Inversión"];
    A.B2 = "Broker"; A.C2 = 508;
    A.B4 = "Efectivo"; A.B6 = "Cuenta remunerada - 2%"; A.C6 = 0.4;
    A.B8 = "Renta variable"; A.B10 = "Fondos"; A.B12 = "Nombre"; A.C12 = "%"; A.D12 = "ID"; A.E12 = "Eur";
    A.B13 = "Fondo indexado mundial"; A.C13 = 0.42; A.D13 = "DEMO0000001";
    A.B15 = "Renta fija"; A.B17 = "Fondos"; A.B19 = "Nombre"; A.C19 = "%"; A.D19 = "ID"; A.E19 = "Eur";
    A.B20 = "Fondo bonos gobierno"; A.C20 = 0.09; A.D20 = "DEMO0000002";
    A.B21 = "Fondo bonos corporativos"; A.C21 = 0.09; A.D21 = "DEMO0000003";

    return { sheets, txHead, tx };
  }

  window.DemoProvider = DemoProvider;
})();
