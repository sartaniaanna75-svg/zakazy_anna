"""Собирает одну HTML-страницу, которую можно открыть двойным щелчком."""

import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PRICES = ROOT / "data" / "prices.json"
XLSX_LIB = ROOT / "out" / "xlsx.full.min.js"
TARGETS = [
    Path(r"C:\Users\1\Desktop\zakazy.html"),
    ROOT / "out" / "zakazy.html",
]

PAGE = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<title>Заказы поставщикам</title>
<script src="xlsx.full.min.js"></script>
<style>
  body { font-family: Segoe UI, sans-serif; margin: 0; background: #f4f6f8; color: #1c2430; }
  header { background: #16324f; color: white; padding: 16px 24px; }
  nav button { background: transparent; color: white; border: 0; font-size: 18px; margin-right: 18px; cursor: pointer; }
  nav button.active { text-decoration: underline; }
  main { padding: 24px; max-width: 1100px; }
  .card { background: white; border-radius: 10px; padding: 16px 18px; margin-bottom: 16px; }
  button.action { background: #16324f; color: white; border: 0; border-radius: 8px; padding: 10px 16px; font-size: 16px; cursor: pointer; }
  button.secondary { background: #6b7280; }
  button:disabled { background: #b9c0c8; cursor: not-allowed; }
  input[type="text"] { font-size: 16px; padding: 8px; width: 280px; }
  table { width: 100%; border-collapse: collapse; }
  td, th { text-align: left; padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }
  .warn { color: #9a3412; }
  .ok { color: #166534; }
  .muted { color: #4b5563; }
  .hidden { display: none; }
</style>
</head>
<body>
<header>
  <nav>
    <button type="button" data-tab="today" class="active">Сегодня</button>
    <button type="button" data-tab="prices">Прайсы</button>
    <button type="button" data-tab="match">Номенклатура</button>
    <button type="button" data-tab="order">Заказ</button>
  </nav>
</header>
<main>
  <section id="today">
    <h1>Сегодня</h1>
    <div class="card">
      <p><b>1. Наш прайс</b> — Excel с вашими товарами: код, штрихкод, наименование, единица.</p>
      <input id="our-file" type="file">
      <p><b>2. Новый прайс поставщика</b> — после загрузки программа сама свяжет строки.</p>
      <p>Поставщик, если в файле нет такой колонки <input id="supplier-name" type="text" placeholder="Например, Петя"></p>
      <input id="supplier-file" type="file">
      <p id="report" class="muted">Можно загрузить Excel, CSV, ODS или страницу из Google Таблиц. Программа приведёт строки к нашим колонкам: название, цена, штрихкод, код, единица.</p>
    </div>
    <div id="mapper" class="card hidden">
      <h2>Укажите колонки</h2>
      <p>Этот формат программа видит первый раз. Выберите колонки один раз. Следующий такой прайс загрузится сам.</p>
      <p>Строка заголовков <input id="header-row" type="number" min="1" value="1"></p>
      <p>Название <select id="map-name"></select></p>
      <p>Цена <select id="map-price"></select></p>
      <p>Штрихкод <select id="map-barcode"></select></p>
      <p>Код <select id="map-code"></select></p>
      <p>Единица <select id="map-unit"></select></p>
      <div id="map-preview"></div>
      <p><button id="map-apply" class="action" type="button">Загрузить так</button></p>
    </div>
    <div class="card" id="today-stats"></div>
  </section>
  <section id="prices" class="hidden">
    <h1>Прайсы</h1>
    <div class="card" id="supplier-summary"></div>
    <div class="card">
      <p><input id="price-search" type="text" placeholder="Найти товар, например сайра"></p>
      <div id="price-results"></div>
    </div>
  </section>
  <section id="match" class="hidden">
    <h1>Номенклатура</h1>
    <div class="card" id="auto-box"></div>
    <div class="card">
      <h2>Проверьте, похоже на ваш товар</h2>
      <div id="suggest-box"></div>
    </div>
    <div class="card">
      <h2>Не связалось</h2>
      <p class="warn">Для этих строк программа не нашла ваш товар. Их можно связать вручную.</p>
      <p><input id="miss-search" type="text" placeholder="Найти среди несвязанных"> Ваш товар
      <select id="manual-catalog"></select></p>
      <div id="miss-box"></div>
    </div>
  </section>
  <section id="order" class="hidden">
    <h1>Заказ</h1>
    <div id="order-body"></div>
  </section>
</main>
<script id="price-data" type="application/json">__DATA__</script>
<script>
const rows = JSON.parse(document.getElementById("price-data").textContent);
const KEY = "zakazy-html-v2";
const state = JSON.parse(localStorage.getItem(KEY) || "{\"catalog\":[],\"links\":{},\"methods\":{},\"skip\":{},\"absent\":{},\"confirmed\":{},\"templates\":{}}");
state.catalog = state.catalog || [];
state.links = state.links || {};
state.methods = state.methods || {};
state.skip = state.skip || {};
state.absent = state.absent || {};
state.confirmed = state.confirmed || {};
state.templates = state.templates || {};
let pending = null;
let suggestions = [];
let unmatched = [];
function save() {
  localStorage.setItem(KEY, JSON.stringify({
    catalog: state.catalog, links: state.links, methods: state.methods,
    skip: state.skip, absent: state.absent, confirmed: state.confirmed, templates: state.templates
  }));
}
function today() {
  const d = new Date();
  return String(d.getDate()).padStart(2,"0") + "." + String(d.getMonth()+1).padStart(2,"0") + "." + d.getFullYear();
}
function offerKey(row) { return [row[0], row[2], row[1]].join("|"); }
function norm(value) {
  return String(value || "").toLowerCase().replace(/ё/g, "е").replace(/[^a-zа-я0-9]+/gi, " ").replace(/\s+/g, " ").trim();
}
function tokens(value) { return norm(value).split(" ").filter(word => word.length >= 3); }
function barcode(value) {
  const parts = String(value || "").match(/\d{8,14}/g) || [];
  if (!parts.length) return "";
  const ean = parts.find(part => part.replace(/^0+/, "").length === 13) || parts[parts.length - 1];
  return ean.replace(/^0+/, "");
}
function normUnit(value) {
  const text = String(value || "").trim().toLowerCase().replace(".", "");
  const map = {шт:"шт", штука:"шт", штук:"шт", упак:"упак", уп:"упак", упаковка:"упак"};
  return map[text] || text;
}
function unitProblem(catalogUnit, offerUnit) {
  const left = normUnit(catalogUnit);
  const right = normUnit(offerUnit);
  if (left && right && left !== right) return "единица не совпала";
  if (left && !right) return "единица не указана";
  return "";
}
function uniqueOffers() {
  const seen = new Map();
  rows.forEach(row => { const key = offerKey(row); if (!seen.has(key)) seen.set(key, row); });
  return seen;
}
function score(left, right) {
  const a = new Set(tokens(left));
  const b = new Set(tokens(right));
  if (!a.size || !b.size) return 0;
  let hit = 0;
  a.forEach(token => { if (b.has(token)) hit += 1; });
  return hit / Math.max(a.size, b.size);
}
function analyze() {
  suggestions = [];
  unmatched = [];
  const byBarcode = new Map();
  const byName = new Map();
  const index = new Map();
  state.catalog.forEach((item, indexItem) => {
    if (item.barcode) byBarcode.set(item.barcode, indexItem);
    const name = norm(item.name);
    if (name) byName.set(name, indexItem);
    tokens(item.name).forEach(token => {
      const list = index.get(token) || [];
      if (list.length < 40) list.push(indexItem);
      index.set(token, list);
    });
  });
  let auto = 0;
  uniqueOffers().forEach(row => {
    const key = offerKey(row);
    if (state.links[key]) return;
    const code = barcode(row[3]);
    const exact = byName.get(norm(row[1]));
    let hit = code && byBarcode.has(code) ? byBarcode.get(code) : null;
    let method = hit == null ? "" : "штрихкод";
    if (hit == null && exact != null) { hit = exact; method = "точное название"; }
    if (hit != null) {
      const item = state.catalog[hit];
      const problem = unitProblem(item.unit, row[4]);
      const pair = key + "\n" + item.code;
      if (state.skip[pair]) { unmatched.push(key); return; }
      if (problem === "единица не совпала") suggestions.push({key, catalogIndex: hit, reason: method + ", " + problem});
      else { state.links[key] = item.code; state.methods[key] = method; auto += 1; }
      return;
    }
    const candidates = new Map();
    tokens(row[1]).forEach(token => (index.get(token) || []).forEach(indexItem => candidates.set(indexItem, 1)));
    let best = -1, bestScore = 0, second = 0;
    candidates.forEach((_, indexItem) => {
      const item = state.catalog[indexItem];
      if (state.skip[key + "\n" + item.code]) return;
      const value = score(row[1], item.name);
      if (value > bestScore) { second = bestScore; bestScore = value; best = indexItem; }
      else if (value > second) second = value;
    });
    if (best >= 0 && bestScore >= 0.5 && bestScore - second >= 0.1) suggestions.push({key, catalogIndex: best, reason: "похожее название"});
    else unmatched.push(key);
  });
  save();
  return auto;
}
function reportText(auto) {
  if (!state.catalog.length) return "Сначала загрузите наш прайс.";
  return "Связано автоматически: " + Object.keys(state.links).length
    + ". Сомнения, нужно проверить: " + suggestions.length
    + ". Не связалось: " + unmatched.length + ".";
}
function show(tab) {
  document.querySelectorAll("main section").forEach(node => node.classList.add("hidden"));
  document.getElementById(tab).classList.remove("hidden");
  document.querySelectorAll("nav button").forEach(button => button.classList.toggle("active", button.dataset.tab === tab));
  if (tab === "today") renderToday();
  if (tab === "prices") renderPrices();
  if (tab === "match") renderMatch();
  if (tab === "order") renderOrder();
}
document.querySelectorAll("nav button").forEach(button => button.addEventListener("click", () => show(button.dataset.tab)));
function renderToday() {
  const offers = uniqueOffers();
  document.getElementById("today-stats").innerHTML = "<p>Строк цен: <b>" + rows.length + "</b>. Ваших товаров: <b>" + state.catalog.length + "</b>.</p>"
    + "<p>Связано: <b>" + Object.keys(state.links).length + "</b> из " + offers.size + ".</p>"
    + (Object.keys(state.links).length ? "<p><button class='action' id='go-order' type='button'>Собрать заказ</button></p>" : "<p><button class='action' disabled>Собрать заказ</button></p><p class='warn'>Сначала свяжите товары.</p>");
  const go = document.getElementById("go-order");
  if (go) go.addEventListener("click", () => show("order"));
  document.getElementById("report").textContent = reportText(0);
}
function cell(value) { return String(value == null ? "" : value).trim(); }
function findHeader(grid) {
  for (let index = 0; index < Math.min(grid.length, 20); index += 1) {
    const labels = grid[index].map(item => cell(item).toLowerCase()).join(" | ");
    if (labels.includes("наименование") || labels.includes("номенклатура")) return index;
  }
  return -1;
}
function column(labels, needles, reject) {
  for (let index = 0; index < labels.length; index += 1) {
    const label = labels[index];
    if (reject && label.includes(reject)) continue;
    if (needles.some(needle => label.includes(needle))) return index;
  }
  return -1;
}
function cyrCount(text) { return (String(text).match(/[А-Яа-яЁё]/g) || []).length; }
function decodeText(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder("utf-16le").decode(buffer);
  if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder("utf-16be").decode(buffer);
  let offset = 0;
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) offset = 3;
  const head = new TextDecoder("latin1").decode(bytes.slice(offset, offset + 240));
  const declared = (head.match(/encoding=["']([^"']+)["']/i) || [])[1] || "";
  const utf8 = new TextDecoder("utf-8").decode(bytes.slice(offset));
  const cp1251 = new TextDecoder("windows-1251").decode(bytes.slice(offset));
  if (/1251|866|koi8/i.test(declared)) return cp1251;
  return cyrCount(cp1251) > cyrCount(utf8) ? cp1251 : utf8;
}
function filledCount(grid) {
  return grid.reduce((sum, row) => sum + row.filter(item => cell(item)).length, 0);
}
function bestGrid(candidates) {
  return candidates.filter(Boolean).sort((a, b) => filledCount(b) + cyrCount(b.flat().join(" ")) - filledCount(a) - cyrCount(a.flat().join(" ")))[0] || [];
}
function sheetGrid(book) {
  if (!book || !book.SheetNames || !book.SheetNames.length) return [];
  const grids = book.SheetNames.map(name => XLSX.utils.sheet_to_json(book.Sheets[name], {header: 1, defval: "", raw: false}));
  return bestGrid(grids);
}
function gridFromHtml(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const tables = [...doc.querySelectorAll("table")].map(table => [...table.rows].map(row => [...row.cells].map(item => item.textContent.replace(/\s+/g, " ").trim())));
  return bestGrid(tables);
}
function readGrid(file, done) {
  const reader = new FileReader();
  reader.onload = () => {
    const bytes = new Uint8Array(reader.result);
    const text = decodeText(reader.result);
    if (/\.gsheet$/i.test(file.name)) {
      done(null, "Это ярлык Google Таблицы, в нём нет строк прайса. В самой таблице нажмите Файл, затем Скачать, и выберите CSV или Excel. Этот скачанный файл уже откроется здесь.");
      return;
    }
    const ole = bytes[0] === 0xD0 && bytes[1] === 0xCF;
    const zip = bytes[0] === 0x50 && bytes[1] === 0x4B;
    const variants = [];
    const tryBook = (data, options) => {
      try { variants.push(sheetGrid(XLSX.read(data, options))); } catch (error) { /* этот способ не подошёл */ }
    };
    if (ole) tryBook(reader.result, {type: "array", codepage: 1251});
    else if (zip) tryBook(reader.result, {type: "array"});
    else {
      if (/<table/i.test(text)) variants.push(gridFromHtml(text));
      tryBook(text, {type: "string"});
      tryBook(text, {type: "string", FS: ";", codepage: 1251});
      tryBook(text, {type: "string", FS: "\t"});
    }
    const grid = bestGrid(variants);
    if (filledCount(grid) < 2) {
      done(null, "В файле не нашлась таблица с товарами. Если это Google Таблица, скачайте её через Файл → Скачать → CSV и загрузите ещё раз.");
      return;
    }
    done(grid);
  };
  reader.readAsArrayBuffer(file);
}
function headerKey(labels) {
  return labels.map(item => cell(item).toLowerCase()).filter(Boolean).join("|");
}
function guessHeader(grid) {
  let best = 0;
  let bestScore = -1;
  for (let index = 0; index < Math.min(grid.length, 30); index += 1) {
    const joined = grid[index].map(item => cell(item).toLowerCase()).join(" ");
    const filled = grid[index].filter(item => cell(item)).length;
    if (filled < 2) continue;
    let score = filled;
    if (/наимен|номенклат|товар|назван|product|name/.test(joined)) score += 6;
    if (/цен|руб|price|стоим|прайс/.test(joined)) score += 6;
    score += Math.min((joined.match(/[а-яё]/g) || []).length, 12);
    if (score > bestScore) { bestScore = score; best = index; }
  }
  return best;
}
function guessMap(labels) {
  const low = labels.map(item => cell(item).toLowerCase());
  return {
    name: column(low, ["наименование", "номенклатура", "название", "товар", "product", "name"]),
    price: column(low, ["цена", "price", "стоимость", "цпрайс", "прайс"]),
    barcode: column(low, ["штрихкод", "штрих", "barcode", "ean"]),
    code: column(low, ["код", "артикул", "sku"], "штрих"),
    unit: column(low, ["единица", "ед.", "ед ", "unit"])
  };
}
function fillMapper(grid, header, map) {
  const labels = grid[header].map((item, index) => cell(item) || ("Колонка " + (index + 1)));
  ["name", "price", "barcode", "code", "unit"].forEach(field => {
    const select = document.getElementById("map-" + field);
    const empty = field === "name" || (field === "price" && pending.kind === "supplier") ? "" : "<option value='-1'>—</option>";
    select.innerHTML = empty + labels.map((label, index) => "<option value='" + index + "'>" + label + "</option>").join("");
    if (map[field] >= 0) select.value = String(map[field]);
  });
  document.getElementById("header-row").value = header + 1;
  const sample = grid.slice(header + 1, header + 4).map(row => "<tr>" + labels.map((_, index) => "<td>" + cell(row[index]) + "</td>").join("") + "</tr>").join("");
  document.getElementById("map-preview").innerHTML = "<table><tr>" + labels.map(label => "<th>" + label + "</th>").join("") + "</tr>" + sample + "</table>";
  document.getElementById("mapper").classList.remove("hidden");
}
function chosenMap() {
  const read = id => Number(document.getElementById(id).value);
  return {name: read("map-name"), price: read("map-price"), barcode: read("map-barcode"), code: read("map-code"), unit: read("map-unit")};
}
function applyRows(grid, header, map, file, kind) {
  const typed = document.getElementById("supplier-name").value.trim() || file.name.replace(/\.[^.]+$/, "");
  if (kind === "our") {
    state.catalog = [];
    state.links = {};
    state.methods = {};
    for (let index = header + 1; index < grid.length; index += 1) {
      const row = grid[index];
      const name = cell(row[map.name]);
      if (!name) continue;
      state.catalog.push({
        code: map.code >= 0 ? cell(row[map.code]) || name : name,
        barcode: map.barcode >= 0 ? barcode(row[map.barcode]) : "",
        name, unit: map.unit >= 0 ? cell(row[map.unit]) : ""
      });
    }
    analyze();
    document.getElementById("report").textContent = "Наш прайс загружен, " + state.catalog.length + " товаров. " + reportText(0);
    renderToday();
    show("match");
    return;
  }
  let added = 0;
  for (let index = header + 1; index < grid.length; index += 1) {
    const row = grid[index];
    const name = cell(row[map.name]);
    const price = map.price >= 0 ? Number(String(row[map.price] || "").replace(/\s/g, "").replace(",", ".").replace(/[^\d.\-]/g, "")) : 0;
    if (!name || !price) continue;
    rows.push([typed, name, map.code >= 0 ? cell(row[map.code]) : "", map.barcode >= 0 ? barcode(row[map.barcode]) : "", map.unit >= 0 ? cell(row[map.unit]) : "", price, today(), "Цена", file.name]);
    added += 1;
  }
  if (state.catalog.length) analyze();
  document.getElementById("report").textContent = "Прайс загружен: " + added + " строк. " + (state.catalog.length ? reportText(0) : "Наш прайс ещё не загружен, связка будет после него.");
  if (state.catalog.length) show("match");
  else renderToday();
}
function startImport(file, kind) {
  readGrid(file, (grid, error) => {
    if (!grid) {
      document.getElementById("report").textContent = error;
      return;
    }
    const header = findHeader(grid) >= 0 ? findHeader(grid) : guessHeader(grid);
    const labels = grid[header].map(item => cell(item));
    const saved = state.templates[headerKey(labels)];
    const map = saved || guessMap(labels.map(item => item.toLowerCase()));
    pending = {grid, header, file, kind, labels};
    const ready = map.name >= 0 && (kind === "our" || map.price >= 0);
    if (saved || ready) {
      applyRows(grid, header, map, file, kind);
      document.getElementById("mapper").classList.add("hidden");
      return;
    }
    fillMapper(grid, header, map);
    document.getElementById("report").textContent = "Формат незнакомый. Укажите, где название" + (kind === "supplier" ? " и цена" : "") + ", и нажмите «Загрузить так».";
  });
}
document.getElementById("header-row").addEventListener("change", () => {
  if (!pending) return;
  const header = Math.max(0, Number(document.getElementById("header-row").value) - 1);
  pending.header = header;
  pending.labels = pending.grid[header].map(item => cell(item));
  fillMapper(pending.grid, header, guessMap(pending.labels.map(item => item.toLowerCase())));
});
document.getElementById("map-apply").addEventListener("click", () => {
  if (!pending) return;
  const map = chosenMap();
  if (!(map.name >= 0) || (pending.kind === "supplier" && !(map.price >= 0))) {
    alert("Выберите колонку названия" + (pending.kind === "supplier" ? " и цены" : "") + ".");
    return;
  }
  state.templates[headerKey(pending.labels)] = map;
  save();
  applyRows(pending.grid, pending.header, map, pending.file, pending.kind);
  document.getElementById("mapper").classList.add("hidden");
});
document.getElementById("our-file").addEventListener("change", event => {
  const file = event.target.files[0];
  if (file) startImport(file, "our");
});
document.getElementById("supplier-file").addEventListener("change", event => {
  const file = event.target.files[0];
  if (file) startImport(file, "supplier");
});
function renderPrices() {
  const counts = new Map();
  rows.forEach(row => counts.set(row[0], (counts.get(row[0]) || 0) + 1));
  document.getElementById("supplier-summary").innerHTML = "<table><tr><th>Поставщик</th><th>Строк</th></tr>"
    + [...counts.entries()].sort((a,b) => b[1] - a[1]).map(([name, count]) => "<tr><td>" + name + "</td><td>" + count + "</td></tr>").join("") + "</table>";
  drawPriceSearch();
}
function drawPriceSearch() {
  const query = (document.getElementById("price-search").value || "").trim().toLowerCase();
  const found = query ? rows.filter(row => row[1].toLowerCase().includes(query)).slice(0, 40) : [];
  document.getElementById("price-results").innerHTML = query
    ? "<p class='muted'>Показаны первые " + found.length + ".</p><table><tr><th>Поставщик</th><th>Название</th><th>Цена</th><th>Ед.</th></tr>"
      + found.map(row => "<tr><td>" + row[0] + "</td><td>" + row[1] + "</td><td>" + row[5] + "</td><td>" + (row[4] || "—") + "</td></tr>").join("") + "</table>"
    : "<p class='muted'>Введите название, чтобы сравнить поставщиков.</p>";
}
document.getElementById("price-search").addEventListener("input", drawPriceSearch);
function renderMatch() {
  const offers = uniqueOffers();
  document.getElementById("auto-box").innerHTML = "<p class='ok'>Связано автоматически: " + Object.keys(state.links).length + "</p>";
  document.getElementById("suggest-box").innerHTML = suggestions.length ? suggestions.slice(0, 30).map(item => {
    const row = offers.get(item.key);
    const catalog = state.catalog[item.catalogIndex];
    return "<p><b>" + catalog.name + "</b> <span class='muted'>" + catalog.code + "</span><br>"
      + row[1] + " <span class='muted'>" + row[0] + "</span><br><span class='warn'>" + item.reason + "</span><br>"
      + "<button class='action yes' type='button' data-key='" + encodeURIComponent(item.key) + "' data-code='" + encodeURIComponent(catalog.code) + "'>Связать</button> "
      + "<button class='secondary no' type='button' data-key='" + encodeURIComponent(item.key) + "' data-code='" + encodeURIComponent(catalog.code) + "'>Это не тот товар</button></p>";
  }).join("") : "<p class='muted'>Сомнений нет.</p>";
  document.querySelectorAll(".yes").forEach(button => button.addEventListener("click", () => {
    const key = decodeURIComponent(button.dataset.key);
    state.links[key] = decodeURIComponent(button.dataset.code);
    state.methods[key] = "проверка";
    suggestions = suggestions.filter(item => item.key !== key);
    save(); renderMatch();
  }));
  document.querySelectorAll(".no").forEach(button => button.addEventListener("click", () => {
    const key = decodeURIComponent(button.dataset.key);
    const code = decodeURIComponent(button.dataset.code);
    state.skip[key + "\n" + code] = true;
    suggestions = suggestions.filter(item => item.key !== key);
    unmatched.push(key);
    save(); renderMatch();
  }));
  const select = document.getElementById("manual-catalog");
  select.innerHTML = state.catalog.slice(0, 300).map(item => "<option value='" + encodeURIComponent(item.code) + "'>" + item.name + "</option>").join("");
  drawMiss();
}
function drawMiss() {
  const offers = uniqueOffers();
  const query = (document.getElementById("miss-search").value || "").trim().toLowerCase();
  const list = unmatched.filter(key => !query || key.toLowerCase().includes(query)).slice(0, 30);
  document.getElementById("miss-box").innerHTML = "<p>Не связалось: <b>" + unmatched.length + "</b>. Показаны первые " + list.length + ".</p>"
    + list.map(key => {
      const row = offers.get(key);
      return "<p>" + (row ? row[1] + " <span class='muted'>" + row[0] + "</span>" : key)
        + " <button class='action manual' type='button' data-key='" + encodeURIComponent(key) + "'>Связать</button></p>";
    }).join("");
  document.querySelectorAll(".manual").forEach(button => button.addEventListener("click", () => {
    const code = decodeURIComponent(document.getElementById("manual-catalog").value || "");
    if (!code) { alert("Сначала загрузите наш прайс."); return; }
    const key = decodeURIComponent(button.dataset.key);
    state.links[key] = code;
    state.methods[key] = "вручную";
    unmatched = unmatched.filter(item => item !== key);
    save(); renderMatch();
  }));
}
document.getElementById("miss-search").addEventListener("input", drawMiss);
function renderOrder() {
  const box = document.getElementById("order-body");
  const codes = [...new Set(Object.values(state.links))];
  if (!codes.length) {
    box.innerHTML = "<div class='card'><button class='action' disabled>Собрать заказ</button><p class='warn'>Сначала свяжите товары.</p></div>";
    return;
  }
  const catalog = new Map(state.catalog.map(item => [item.code, item]));
  const offers = uniqueOffers();
  const day = today();
  box.innerHTML = codes.slice(0, 80).map(code => {
    const item = catalog.get(code) || {name: code, unit: ""};
    const group = [...offers.values()].filter(row => state.links[offerKey(row)] === code);
    const problem = group.find(row => unitProblem(item.unit, row[4]));
    if (problem && group.some(row => normUnit(row[4]) && normUnit(item.unit) && normUnit(row[4]) !== normUnit(item.unit))) {
      return "<div class='card'><p><b>" + item.name + "</b></p><p class='warn'>единица не совпала</p></div>";
    }
    const usable = group.filter(row => state.absent[code + "|" + row[0]] !== day && !unitProblem(item.unit, row[4]));
    if (!usable.length) return "<div class='card'><p><b>" + item.name + "</b></p><p class='warn'>Нет строки с той же единицей.</p></div>";
    usable.sort((a, b) => a[5] - b[5]);
    const fixed = state.confirmed[code];
    let winner = usable[0];
    if (fixed && fixed.date === day) winner = usable.find(row => row[0] === fixed.supplier) || winner;
    const next = usable.find(row => row[0] !== winner[0]);
    return "<div class='card'><p><b>" + item.name + "</b></p><p>" + winner[0] + " — " + winner[5] + "</p><p class='muted'>" + winner[1] + "</p>"
      + (next ? "<p class='muted'>Следующий: " + next[0] + " — " + next[5] + "</p>" : "")
      + "<button class='action confirm' type='button' data-code='" + encodeURIComponent(code) + "' data-supplier='" + encodeURIComponent(winner[0]) + "' data-price='" + winner[5] + "'>Подтвердить</button> "
      + "<button class='secondary absent' type='button' data-code='" + encodeURIComponent(code) + "' data-supplier='" + encodeURIComponent(winner[0]) + "'>Нет в наличии</button></div>";
  }).join("");
  document.querySelectorAll(".absent").forEach(button => button.addEventListener("click", () => {
    const code = decodeURIComponent(button.dataset.code);
    const supplier = decodeURIComponent(button.dataset.supplier);
    state.absent[code + "|" + supplier] = today();
    if (state.confirmed[code] && state.confirmed[code].supplier === supplier) delete state.confirmed[code];
    save(); renderOrder();
  }));
  document.querySelectorAll(".confirm").forEach(button => button.addEventListener("click", () => {
    state.confirmed[decodeURIComponent(button.dataset.code)] = {date: today(), supplier: decodeURIComponent(button.dataset.supplier), price: Number(button.dataset.price)};
    save(); renderOrder();
  }));
}
if (state.catalog.length) analyze();
renderToday();
</script>
</body>
</html>
"""


def main() -> None:
    source = json.loads(PRICES.read_text(encoding="utf-8"))
    compact = [
        [
            row.get("Поставщик") or "",
            row.get("Наименование") or "",
            row.get("Код товара") or "",
            row.get("Штрихкод") or "",
            row.get("Единица измерения") or "",
            row.get("Цена") or 0,
            row.get("Дата обновления") or "",
            row.get("Тип цены") or "",
            row.get("Исходный файл") or "",
        ]
        for row in source
    ]
    payload = json.dumps(compact, ensure_ascii=False).replace("<", "\\u003c")
    html = PAGE.replace("__DATA__", payload)
    for target in TARGETS:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(html, encoding="utf-8")
        library = target.with_name("xlsx.full.min.js")
        if XLSX_LIB.exists() and XLSX_LIB.resolve() != library.resolve():
            shutil.copyfile(XLSX_LIB, library)
        print(target)


if __name__ == "__main__":
    main()
