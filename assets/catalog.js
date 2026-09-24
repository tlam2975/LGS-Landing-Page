const WORKBOOK_FILE = "../28082026. Total product on MiniApp.xlsx";
const IMAGE_BY_MODEL = {};
const PAGE = window.CATEGORY_PAGE || { key: "all", title: "Products", groups: [] };

const els = {
  search: document.getElementById("search"),
  type: document.getElementById("type"),
  sort: document.getElementById("sort"),
  plan: document.getElementById("plan"),
  reset: document.getElementById("reset"),
  grid: document.getElementById("grid"),
  empty: document.getElementById("empty"),
  status: document.getElementById("status"),
  title: document.getElementById("page-title"),
  copy: document.getElementById("page-copy")
};

let products = [];
const moneyFormat = new Intl.NumberFormat("vi-VN", { style: "currency", currency: "VND", maximumFractionDigits: 0 });
const money = value => value ? moneyFormat.format(value) : "n.a.";
const toNumber = value => Number.isFinite(+value) ? Math.round(+value) : null;
const text = (key, params) => window.I18N ? window.I18N.t(key, params) : key;

function sheetRows(workbook, name) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" });
}

function normalizeCategory(value, code) {
  if (!value || value === "#N/A") {
    return ({ ES: "Air Solution", AV: "Audio/Video", AO: "Audio/Video" })[code] || "Audio/Video";
  }
  if (value === "ES") return "Air Solution";
  if (value === "AV" || value === "AO") return "Audio/Video";
  if (value === "Appliance" || value === "Audio / Video" || value === "Other") return "Audio/Video";
  return value;
}

function routeCategory(category) {
  if (category === "Refrigerator") return "refrigerator";
  if (category === "TV") return "tv";
  if (category === "Audio/Video") return "audio-video";
  return "appliances";
}

function displayType(category, level, fallback) {
  if (category !== "Refrigerator") return level[9] || fallback;

  const family = [level[11], level[13], fallback].join(" ").toLowerCase();
  if (family.includes("french door") || family.includes("f/d")) return "FD";
  if ((level[9] || "").toLowerCase() === "bottom freezer" || fallback === "B/F") return "BF";

  return level[9] || fallback;
}

function parseWorkbook(workbook) {
  const priceRows = sheetRows(workbook, "Total MiniApp");
  const levelRows = sheetRows(workbook, "Product level");
  const levelBySuffix = new Map(levelRows.slice(3).filter(row => row[4]).map(row => [row[4], row]));
  return priceRows.slice(3).filter(row => row[3] && row[4]).map(row => {
    const level = levelBySuffix.get(row[4]) || [];
    const category = normalizeCategory(level[7], row[1]);
    return {
      productCode: row[1],
      subCode: row[2],
      model: row[3],
      suffix: row[4],
      outright: toNumber(row[5]) || toNumber(row[10]) || toNumber(row[6]) || toNumber(row[11]),
      monthly12: toNumber(row[7]),
      monthly24: toNumber(row[8]),
      monthly36: toNumber(row[9]),
      total12: toNumber(row[12]),
      total24: toNumber(row[13]),
      total36: toNumber(row[14]),
      category,
      route: routeCategory(category),
      type: displayType(category, level, row[2]),
      series: level[11] || "",
      variant: level[13] || ""
    };
  }).filter(product => PAGE.key === "all" || product.route === PAGE.key);
}

function lgSearch(value) {
  return `https://www.lg.com/vn/search/?search=${encodeURIComponent(value)}`;
}

function imageFor(product) {
  return IMAGE_BY_MODEL[product.model] || IMAGE_BY_MODEL[product.suffix] || "";
}

function setOptions() {
  const currentSort = els.sort.value || "monthly36";
  const currentPlan = els.plan.value || "36";
  els.sort.innerHTML = `
    <option value="monthly36">${text("lowestMonthly")}</option>
    <option value="outright">${text("lowestOutright")}</option>
    <option value="model">${text("modelAZ")}</option>
  `;
  els.plan.innerHTML = `
    <option value="36">${text("months36")}</option>
    <option value="24">${text("months24")}</option>
    <option value="12">${text("months12")}</option>
  `;
  els.sort.value = currentSort;
  els.plan.value = currentPlan;
  renderTypes(products);
}

function renderTypes(list) {
  const current = els.type.value;
  const types = [...new Set(list.map(product => product.type).filter(Boolean))].sort();
  els.type.innerHTML = `<option value="all">${text("allTypes")}</option>` + types.map(type => `<option value="${type}">${type}</option>`).join("");
  els.type.value = types.includes(current) ? current : "all";
}

function card(product) {
  const planMonths = els.plan.value;
  const source = imageFor(product);
  const image = source
    ? `<img src="${source}" alt="${product.model}" onerror="this.remove();this.parentElement.insertAdjacentHTML('beforeend','<div class=&quot;placeholder&quot;>${text("imagePending")}<br>${product.model}</div>')">`
    : `<div class="placeholder">${text("imagePending")}<br>${product.model}</div>`;
  return `
    <article class="card">
      <div class="media"><span class="badge">${product.category}</span>${image}</div>
      <div class="body">
        <div class="kicker">${product.type || product.subCode}</div>
        <div class="model">${product.model}</div>
        <div class="meta">${product.variant || product.series || product.suffix}<br>${product.suffix}</div>
        <div class="price-row">
          <div class="price"><span>${text("outright")}</span><strong>${money(product.outright)}</strong></div>
          <div class="price"><span>${text("monthTotal", { months: planMonths })}</span><strong>${money(product[`total${planMonths}`])}</strong></div>
        </div>
        <div class="plans">
          <div class="plan"><span>${text("months12")}</span><strong>${money(product.monthly12)}</strong></div>
          <div class="plan"><span>${text("months24")}</span><strong>${money(product.monthly24)}</strong></div>
          <div class="plan"><span>${text("months36")}</span><strong>${money(product.monthly36)}</strong></div>
        </div>
        <div class="links">
          <a class="primary" href="${lgSearch(product.model)}" target="_blank" rel="noopener">${text("findImage")}</a>
          <a href="${lgSearch(product.suffix)}" target="_blank" rel="noopener">LG.com</a>
        </div>
      </div>
    </article>`;
}

function render() {
  const query = els.search.value.trim().toLowerCase();
  let list = products.slice();
  renderTypes(list);
  if (els.type.value !== "all") list = list.filter(product => product.type === els.type.value);
  if (query) {
    list = list.filter(product => [product.model, product.suffix, product.category, product.type, product.variant, product.series].join(" ").toLowerCase().includes(query));
  }
  const key = els.sort.value;
  list.sort((a, b) => key === "model" ? a.model.localeCompare(b.model) : (a[key] || Infinity) - (b[key] || Infinity));
  els.grid.innerHTML = list.map(card).join("");
  els.empty.hidden = list.length > 0;
  els.status.textContent = text("productsShown", { shown: list.length, total: products.length });
}

async function loadWorkbook() {
  els.title.textContent = text(PAGE.titleKey) || PAGE.title;
  els.copy.textContent = text(PAGE.copyKey) || PAGE.copy;
  setOptions();
  try {
    const response = await fetch(WORKBOOK_FILE);
    if (!response.ok) throw new Error(WORKBOOK_FILE);
    const buffer = await response.arrayBuffer();
    products = parseWorkbook(XLSX.read(buffer));
    setOptions();
    render();
  } catch (error) {
    els.status.textContent = text("workbookError");
    els.grid.innerHTML = `<div class="empty">${text("workbookHelp")}</div>`;
  }
}

[els.search, els.type, els.sort, els.plan].forEach(element => element.addEventListener("input", render));
els.reset.addEventListener("click", () => {
  els.search.value = "";
  els.type.value = "all";
  els.sort.value = "monthly36";
  els.plan.value = "36";
  render();
});

loadWorkbook();

document.addEventListener("languagechange", () => {
  if (els.search.dataset.i18nPlaceholder) els.search.placeholder = text(els.search.dataset.i18nPlaceholder);
  if (els.reset.dataset.i18n) els.reset.textContent = text(els.reset.dataset.i18n);
  els.title.textContent = text(PAGE.titleKey) || PAGE.title;
  els.copy.textContent = text(PAGE.copyKey) || PAGE.copy;
  setOptions();
  render();
});
