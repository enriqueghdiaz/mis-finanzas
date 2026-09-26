/* Gráficos SVG ligeros (sin librerías) con tooltip al tocar. */
(function () {
  const NS = "http://www.w3.org/2000/svg";
  const SERIES = ["--s1", "--s2", "--s3", "--s4", "--s5", "--s6", "--s7", "--s8"].map(v => `var(${v})`);
  // Formato español con separador de miles siempre (1.707 €, 12.345,67 €)
  function fmtNum(v, dec) {
    v = +v || 0;
    const neg = v < 0 && Math.abs(v).toFixed(dec) !== (0).toFixed(dec);
    const [i, d] = Math.abs(v).toFixed(dec).split(".");
    return (neg ? "-" : "") + i.replace(/\B(?=(\d{3})+(?!\d))/g, ".") + (d ? "," + d : "");
  }
  const eur = v => fmtNum(v, Math.abs(v || 0) >= 1000 ? 0 : 2) + "\u00a0€";
  const eurShort = v => Math.abs(v) >= 1000 ? fmtNum(v / 1000, Math.abs(v) >= 10000 || Math.round(v / 100) % 10 === 0 ? 0 : 1) + "k" : fmtNum(Math.round(v), 0);

  function el(tag, attrs, parent) {
    const e = document.createElementNS(NS, tag);
    for (const k in attrs) e.setAttribute(k, attrs[k]);
    if (parent) parent.appendChild(e);
    return e;
  }
  function svg(container, h) {
    container.innerHTML = "";
    const w = Math.max(260, container.clientWidth || 320);
    const s = el("svg", { viewBox: `0 0 ${w} ${h}`, width: "100%", height: h, class: "chart", role: "img" }, container);
    return { s, w, h };
  }
  function niceMax(v) {
    if (v <= 0) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(v)));
    const n = v / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
  }
  // Tooltip
  const tip = () => document.getElementById("tooltip");
  function showTip(html, x, y) {
    const t = tip(); t.innerHTML = html; t.hidden = false;
    const r = t.getBoundingClientRect();
    let left = Math.min(window.innerWidth - r.width - 8, Math.max(8, x - r.width / 2));
    let top = y - r.height - 12; if (top < 8) top = y + 16;
    t.style.left = left + "px"; t.style.top = top + "px";
  }
  function hideTip() { const t = tip(); if (t) t.hidden = true; }
  document.addEventListener("pointerdown", e => { if (!e.target.closest(".chart")) hideTip(); }, { passive: true });
  window.addEventListener("scroll", hideTip, { passive: true });

  function roundedTopRect(x, y, w, h, r) {
    if (h <= 0) return "";
    r = Math.min(r, w / 2, h);
    return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
  }
  function roundedRightRect(x, y, w, h, r) {
    if (w <= 0) return "";
    r = Math.min(r, h / 2, w);
    return `M${x},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h - r}Q${x + w},${y + h} ${x + w - r},${y + h}H${x}Z`;
  }

  // ── Donut ────────────────────────────────────────────────
  function donut(container, data, opts = {}) {
    const total = data.reduce((a, d) => a + d.value, 0);
    container.innerHTML = "";
    const wrap = document.createElement("div"); wrap.className = "donut-wrap"; container.appendChild(wrap);
    const box = document.createElement("div"); wrap.appendChild(box);
    const size = 150, r = 62, sw = 20, cx = size / 2, cy = size / 2;
    const s = el("svg", { viewBox: `0 0 ${size} ${size}`, width: size, height: size, class: "chart", role: "img", "aria-label": opts.label || "Gráfico circular" }, box);
    if (!total) { el("circle", { cx, cy, r, fill: "none", stroke: "var(--grid)", "stroke-width": sw }, s); }
    let a0 = -Math.PI / 2;
    const gap = data.length > 1 ? 0.025 : 0;
    data.forEach((d, i) => {
      const frac = d.value / total; if (!frac) return;
      const a1 = a0 + frac * Math.PI * 2;
      const s0 = a0 + gap / 2, s1 = Math.max(s0 + 0.001, a1 - gap / 2);
      const large = s1 - s0 > Math.PI ? 1 : 0;
      const p = el("path", {
        d: `M${cx + r * Math.cos(s0)},${cy + r * Math.sin(s0)} A${r},${r} 0 ${large} 1 ${cx + r * Math.cos(s1)},${cy + r * Math.sin(s1)}`,
        fill: "none", stroke: d.color, "stroke-width": sw, "stroke-linecap": "butt", style: "cursor:pointer"
      }, s);
      p.addEventListener("pointerdown", e => showTip(`<b>${d.label}</b>${eur(d.value)} · ${(frac * 100).toFixed(1)}%`, e.clientX, e.clientY));
      a0 = a1;
    });
    const t1 = el("text", { x: cx, y: cy - 2, "text-anchor": "middle", style: "font-size:17px;font-weight:700;fill:var(--text)" }, s); t1.textContent = eur(total).replace(/,\d\d(?=\s)/, "");
    const t2 = el("text", { x: cx, y: cy + 16, "text-anchor": "middle", style: "font-size:11px" }, s); t2.textContent = opts.center || "Total";
    const lg = document.createElement("div"); lg.className = "donut-legend"; wrap.appendChild(lg);
    data.forEach(d => {
      const row = document.createElement("div");
      row.innerHTML = `<i style="background:${d.color}"></i><span class="n">${d.label}</span><b class="num">${total ? Math.round(d.value / total * 100) : 0}%</b>`;
      lg.appendChild(row);
    });
  }

  // ── Barras apiladas por mes (con marca de presupuesto opcional) ──
  function stackedBars(container, labels, series, opts = {}) {
    const H = opts.height || 210;
    const { s, w, h } = svg(container, H);
    const m = { l: 38, r: 8, t: 10, b: 24 };
    const iw = w - m.l - m.r, ih = h - m.t - m.b;
    const totals = labels.map((_, i) => series.reduce((a, se) => a + (se.values[i] || 0), 0));
    const maxV = niceMax(Math.max(...totals, ...(opts.target || []).map(v => v || 0)));
    for (let k = 0; k <= 4; k++) {
      const y = m.t + ih - ih * k / 4;
      el("line", { x1: m.l, x2: w - m.r, y1: y, y2: y, class: "gridline" }, s);
      const t = el("text", { x: m.l - 6, y: y + 3, "text-anchor": "end" }, s); t.textContent = eurShort(maxV * k / 4);
    }
    const band = iw / labels.length, bw = Math.min(34, band * 0.55);
    labels.forEach((lab, i) => {
      const x = m.l + band * i + (band - bw) / 2;
      let y0 = m.t + ih;
      const visible = series.map((se, k) => ({ se, v: se.values[i] || 0, k })).filter(o => o.v > 0);
      visible.forEach((o, n) => {
        const bh = ih * o.v / maxV;
        const gapPx = n > 0 ? 2 : 0;
        const top = n === visible.length - 1;
        const y = y0 - bh;
        const d = top ? roundedTopRect(x, y, bw, Math.max(0, bh - gapPx), 4) : `M${x},${y}h${bw}v${Math.max(0, bh - gapPx)}h${-bw}Z`;
        el("path", { d, fill: o.se.color }, s);
        y0 = y;
      });
      if (opts.target && opts.target[i] != null) {
        const ty = m.t + ih - ih * opts.target[i] / maxV;
        el("line", { x1: x - 5, x2: x + bw + 5, y1: ty, y2: ty, stroke: "var(--text)", "stroke-width": 2, "stroke-dasharray": "4 3" }, s);
      }
      const t = el("text", { x: x + bw / 2, y: h - 6, "text-anchor": "middle" }, s); t.textContent = lab;
      const hit = el("rect", { x: m.l + band * i, y: m.t, width: band, height: ih, fill: "transparent", style: "cursor:pointer" }, s);
      hit.addEventListener("pointerdown", e => {
        let html = `<b>${opts.fullLabels ? opts.fullLabels[i] : lab}</b>`;
        series.forEach(se => html += `<div><span style="color:${se.color}">■</span> ${se.name}: ${eur(se.values[i] || 0)}</div>`);
        html += `<div>Total: <b style="display:inline">${eur(totals[i])}</b></div>`;
        if (opts.target && opts.target[i] != null) html += `<div>${opts.targetName || "Presupuesto"}: ${eur(opts.target[i])}</div>`;
        showTip(html, e.clientX, e.clientY);
      });
    });
    el("line", { x1: m.l, x2: w - m.r, y1: m.t + ih, y2: m.t + ih, class: "axis" }, s);
    if (opts.legend !== false) legend(container, series, opts.target ? opts.targetName || "Presupuesto" : null);
  }

  function legend(container, series, targetName) {
    const lg = document.createElement("div"); lg.className = "legend";
    lg.innerHTML = series.map(se => `<span><i style="background:${se.color}"></i>${se.name}</span>`).join("") +
      (targetName ? `<span><i style="background:none;border-top:2px dashed var(--text);border-radius:0;height:0;width:14px"></i>${targetName}</span>` : "");
    container.appendChild(lg);
  }

  // ── Barras horizontales real vs presupuesto ──
  function bullet(container, rows, opts = {}) {
    const rowH = 30, H = rows.length * rowH + 8;
    const { s, w } = svg(container, H);
    const lw = Math.min(118, w * 0.36), rw = 64, iw = w - lw - rw - 8;
    const maxV = Math.max(1, ...rows.map(r => Math.max(r.real, r.budget || 0)));
    rows.forEach((r, i) => {
      const y = 4 + i * rowH;
      const t = el("text", { x: 0, y: y + 17 }, s); t.textContent = r.label.length > 17 ? r.label.slice(0, 16) + "…" : r.label;
      el("rect", { x: lw, y: y + 7, width: iw, height: 12, rx: 4, fill: "var(--surface-2)" }, s);
      const over = r.budget != null && r.real > r.budget + 0.005;
      el("path", { d: roundedRightRect(lw, y + 7, iw * r.real / maxV, 12, 4), fill: over ? "var(--neg)" : (opts.color || "var(--s1)") }, s);
      if (r.budget) {
        const bx = lw + iw * r.budget / maxV;
        el("line", { x1: bx, x2: bx, y1: y + 3, y2: y + 23, stroke: "var(--text)", "stroke-width": 2 }, s);
      }
      const v = el("text", { x: w, y: y + 17, "text-anchor": "end", style: "fill:var(--text);font-weight:600" }, s); v.textContent = eurShort(r.real) + " €";
      const hit = el("rect", { x: 0, y, width: w, height: rowH, fill: "transparent" }, s);
      hit.addEventListener("pointerdown", e => showTip(`<b>${r.label}</b>Real: ${eur(r.real)}${r.budget != null ? `<br>Presupuesto: ${eur(r.budget)}<br>${over ? "Exceso: " + eur(r.real - r.budget) : "Disponible: " + eur(r.budget - r.real)}` : ""}`, e.clientX, e.clientY));
    });
  }

  // ── Líneas ──
  function line(container, labels, series, opts = {}) {
    const H = opts.height || 200;
    const { s, w, h } = svg(container, H);
    const m = { l: 40, r: 10, t: 12, b: 24 };
    const iw = w - m.l - m.r, ih = h - m.t - m.b;
    const all = series.flatMap(se => se.values.filter(v => v != null));
    let minV = opts.zero === false ? Math.min(...all) : Math.min(0, ...all);
    let maxV = niceMax(Math.max(...all, 1));
    if (opts.zero === false) { const span = maxV - minV; minV = Math.max(0, Math.floor((minV - span * 0.15) / 100) * 100); maxV = niceMax(maxV); }
    const X = i => m.l + (labels.length === 1 ? iw / 2 : iw * i / (labels.length - 1));
    const Y = v => m.t + ih - ih * (v - minV) / (maxV - minV || 1);
    for (let k = 0; k <= 4; k++) {
      const v = minV + (maxV - minV) * k / 4, y = Y(v);
      el("line", { x1: m.l, x2: w - m.r, y1: y, y2: y, class: "gridline" }, s);
      const t = el("text", { x: m.l - 6, y: y + 3, "text-anchor": "end" }, s); t.textContent = eurShort(v);
    }
    labels.forEach((lab, i) => {
      if (labels.length > 8 && i % Math.ceil(labels.length / 6) !== 0 && i !== labels.length - 1) return;
      const t = el("text", { x: X(i), y: h - 6, "text-anchor": "middle" }, s); t.textContent = lab;
    });
    series.forEach(se => {
      let d = "", pen = false;
      se.values.forEach((v, i) => { if (v == null) { pen = false; return; } d += (pen ? "L" : "M") + X(i) + "," + Y(v); pen = true; });
      el("path", { d, fill: "none", stroke: se.color, "stroke-width": 2, "stroke-linejoin": "round", "stroke-linecap": "round", "stroke-dasharray": se.dash ? "5 4" : "" }, s);
      se.values.forEach((v, i) => { if (v != null && !se.noDots && (labels.length <= 12)) el("circle", { cx: X(i), cy: Y(v), r: 4, fill: se.color, stroke: "var(--surface)", "stroke-width": 2 }, s); });
    });
    const cross = el("line", { x1: 0, x2: 0, y1: m.t, y2: m.t + ih, stroke: "var(--text-3)", "stroke-width": 1, visibility: "hidden" }, s);
    const hit = el("rect", { x: m.l, y: m.t, width: iw, height: ih, fill: "transparent", style: "touch-action:pan-y" }, s);
    const onMove = e => {
      const r = s.getBoundingClientRect();
      const px = (e.clientX - r.left) * (w / r.width);
      const i = Math.max(0, Math.min(labels.length - 1, Math.round((px - m.l) / (iw / Math.max(1, labels.length - 1)))));
      cross.setAttribute("x1", X(i)); cross.setAttribute("x2", X(i)); cross.setAttribute("visibility", "visible");
      let html = `<b>${opts.fullLabels ? opts.fullLabels[i] : labels[i]}</b>`;
      series.forEach(se => { if (se.values[i] != null) html += `<div><span style="color:${se.color}">■</span> ${se.name}: ${eur(se.values[i])}</div>`; });
      showTip(html, e.clientX, e.clientY);
    };
    hit.addEventListener("pointerdown", onMove); hit.addEventListener("pointermove", e => { if (e.buttons || e.pointerType === "mouse") onMove(e); });
    hit.addEventListener("pointerleave", () => cross.setAttribute("visibility", "hidden"));
    if (series.length > 1 && opts.legend !== false) legend(container, series);
  }

  window.Charts = { fmtNum, donut, stackedBars, bullet, line, SERIES, eur, eurShort, hideTip };
})();
