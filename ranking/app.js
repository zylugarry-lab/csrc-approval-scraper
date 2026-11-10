const GROUPS = [
  ["income", "资管收入（亿元）", "g-income", false],
  ["scale", "资管规模（亿元）", "g-scale", false],
  ["private", "私募不含ABS（亿元）", "g-private", true],
  ["public", "公募（亿元）", "g-public", true],
  ["abs", "ABS（亿元）", "g-abs", true],
  ["non_mmf", "非货公募（亿元）", "g-non", true],
];

const state = { year: 2026, period: "H1", q: "" };

function dataset() {
  return window.RANKING.datasets.find((d) => d.year === state.year && d.period === state.period);
}

function fmt(v, growth) {
  if (v === "/" ) return { text: "/", cls: "slash" };
  if (v === null || v === undefined || Number.isNaN(v)) return { text: "", cls: "" };
  const n = Number(v);
  const text = n.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return { text, cls: n < 0 ? "neg" : "" };
}

function cell(v, growth) {
  const f = fmt(v, growth);
  return `<td class="num ${f.cls}">${f.text}</td>`;
}

function renderControls() {
  const years = document.getElementById("years");
  const periods = document.getElementById("periods");
  years.innerHTML = window.RANKING.years.map((y) =>
    `<button type="button" data-year="${y}" aria-pressed="${y === state.year}">${y}</button>`
  ).join("");
  periods.innerHTML = window.RANKING.periods.map((p) =>
    `<button type="button" data-period="${p.id}" aria-pressed="${p.id === state.period}">${p.name}</button>`
  ).join("");
}

function render() {
  const data = dataset();
  const empty = document.getElementById("empty");
  const wrap = document.getElementById("wrap");
  const sub = document.getElementById("subtitle");
  if (!data) {
    wrap.hidden = true;
    empty.hidden = false;
    const have = window.RANKING.datasets.map((d) => `${d.year}年${d.period === "FY" ? "年报" : "半年报"}`);
    sub.textContent = "该报告期还没有生成";
    empty.innerHTML = `<h2>${state.year}年${window.RANKING.periods.find((p) => p.id === state.period).name}还没有数据</h2>
      <p>已入库的报告可以直接打开。</p>
      <div>${window.RANKING.datasets.map((d) =>
        `<button type="button" data-jump="${d.year}|${d.period}">${d.title}</button>`
      ).join("")}</div>`;
    return;
  }
  empty.hidden = true;
  wrap.hidden = false;
  const nScale = data.rows.filter((r) => typeof r.scale[0] === "number").length;
  const nIncome = data.rows.filter((r) => typeof r.income[0] === "number").length;
  sub.textContent = `${data.title} · 截至 ${data.asof} · 规模 ${nScale}/49 · 收入 ${nIncome}/49${data.note ? " · " + data.note : ""}`;
  const head1 = [`<th class="g-id name" rowspan="2">券商名称</th>`];
  const head2 = [];
  for (const [key, title, cls] of GROUPS) {
    const labs = key === "income" ? data.labels.income : key === "scale" ? data.labels.scale : data.labels.block;
    head1.push(`<th class="${cls}" colspan="${labs.length}">${title}</th>`);
    labs.forEach((lab) => head2.push(`<th class="${cls}">${lab}</th>`));
  }
  document.getElementById("head").innerHTML = `<tr>${head1.join("")}</tr><tr>${head2.join("")}</tr>`;
  const q = state.q.trim();
  const rows = data.rows.filter((r) => !q || r.name.includes(q) || r.code.includes(q));
  const html = rows.map((r, i) => {
    const cls = r.ours ? "ours" : (i % 2 ? "zebra" : "");
    const tds = GROUPS.map(([key]) => r[key].map((v, idx) => cell(v, idx >= 2)).join("")).join("");
    return `<tr class="${cls}"><td class="name">${r.name}</td>${tds}</tr>`;
  }).join("");
  const beta = `<tr class="beta"><td class="name">行业β</td>${
    GROUPS.map(([key]) => data.beta[key].map((v, idx) => cell(v, idx >= 2)).join("")).join("")
  }</tr>`;
  document.getElementById("body").innerHTML = html + (q ? "" : beta);
}

function toCsv() {
  const data = dataset();
  if (!data) return;
  const headers = ["券商名称"];
  for (const [key, title] of GROUPS) {
    const labs = key === "income" ? data.labels.income : key === "scale" ? data.labels.scale : data.labels.block;
    labs.forEach((lab) => headers.push(`${title}${lab}`));
  }
  const lines = [headers.join(",")];
  const all = data.rows.concat([{ name: "行业β", ...data.beta }]);
  for (const r of all) {
    const vals = [r.name];
    for (const [key] of GROUPS) {
      r[key].forEach((v) => vals.push(v === null || v === undefined ? "" : v));
    }
    lines.push(vals.map((v) => `"${String(v).replaceAll('"', '""')}"`).join(","));
  }
  const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `券商资管业务排名_${data.year}_${data.period}.csv`;
  a.click();
}

document.body.addEventListener("click", (e) => {
  const y = e.target.closest("[data-year]");
  const p = e.target.closest("[data-period]");
  const j = e.target.closest("[data-jump]");
  if (y) state.year = Number(y.dataset.year);
  if (p) state.period = p.dataset.period;
  if (j) {
    const [yy, pp] = j.dataset.jump.split("|");
    state.year = Number(yy);
    state.period = pp;
  }
  if (y || p || j) {
    renderControls();
    render();
  }
  if (e.target.id === "csv") toCsv();
});
document.getElementById("q").addEventListener("input", (e) => {
  state.q = e.target.value;
  render();
});

renderControls();
render();
