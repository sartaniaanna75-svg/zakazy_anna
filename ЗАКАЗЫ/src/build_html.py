"""Собирает одну HTML-страницу, которую можно открыть двойным щелчком."""

import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PRICES = ROOT / "data" / "prices.json"
XLSX_LIB = ROOT / "out" / "xlsx.full.min.js"
TARGETS = [
    ROOT / "zakazy.html",
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
  main { padding: 24px; max-width: none; }
  .card { background: white; border-radius: 10px; padding: 16px 18px; margin-bottom: 16px; }
  button.action { background: #16324f; color: white; border: 0; border-radius: 8px; padding: 10px 16px; font-size: 16px; cursor: pointer; }
  button.secondary { background: #6b7280; }
  button.quiet { background: transparent; color: #16324f; border: 0; text-decoration: underline; cursor: pointer; font-size: 16px; padding: 0; }
  button.filter { background: #e5e7eb; color: #1c2430; border: 0; border-radius: 8px; padding: 8px 12px; margin-right: 8px; cursor: pointer; }
  button.filter.active { background: #16324f; color: white; }
  button:disabled { background: #b9c0c8; cursor: not-allowed; }
  input[type="text"], select { font-size: 16px; padding: 8px; }
  table { width: 100%; border-collapse: collapse; }
  td, th { text-align: left; padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }
  tr.need { background: #fff7ed; }
  tr.done { background: #f0fdf4; }
  tr.doubt { background: #fff1f2; }
  textarea.row-search.doubt { background: #fee2e2; border: 2px solid #b91c1c; }
  .warn { color: #9a3412; }
  .ok { color: #166534; }
  .muted { color: #4b5563; }
  .hidden { display: none; }
  #picker-list { max-height: 320px; overflow: auto; }
  #match-search { width: 360px; }
  button.no { margin-left: 8px; padding: 6px 12px; }
  #match table { table-layout: fixed; }
  #match th:nth-child(1), #match td:nth-child(1) { width: 22%; }
  #match th:nth-child(2), #match td:nth-child(2) { width: 12%; }
  #match th:nth-child(3), #match td:nth-child(3) { width: 42%; }
  #match th:nth-child(4), #match td:nth-child(4) { width: 14%; }
  #match td { overflow-wrap: anywhere; }
  textarea.row-search { width: 100%; min-height: 5.2em; box-sizing: border-box; font: inherit; font-size: 16px; line-height: 1.35; resize: vertical; }
  #inline-drop { position: fixed; z-index: 30; background: white; border: 1px solid #d1d5db; border-radius: 8px; max-height: 420px; overflow: auto; box-shadow: 0 8px 24px rgba(0,0,0,.12); padding: 8px 10px; }
  #inline-drop button { display: block; width: 100%; text-align: left; margin: 4px 0; white-space: normal; line-height: 1.35; }
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
    <div class="card">
      <p>Совпавшие строки уже идут в заказ: на них только «Нет», если товар чужой. Красное поле — программа не уверена, проверьте и нажмите «Да». Если своего товара нет, нажмите «Нет» или наберите название в строке.</p>
      <p>
        <select id="match-supplier"></select>
        <input id="match-search" type="text" placeholder="Найти в прайсе этого поставщика">
      </p>
      <p id="match-filters"></p>
      <p id="match-counts" class="muted"></p>
    </div>
    <div class="card">
      <table>
        <thead>
          <tr>
            <th>Номенклатура поставщика</th>
            <th>Поставщик</th>
            <th>Наша номенклатура</th>
            <th>Статус</th>
            <th>Подтверждение</th>
          </tr>
        </thead>
        <tbody id="match-body"></tbody>
      </table>
      <p><button id="match-more" class="secondary" type="button">Показать ещё</button></p>
    </div>
    <div id="inline-drop" class="hidden"></div>
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
state.verified = state.verified || {};
state.draft = state.draft || {};
state.cleared = state.cleared || {};
state.supplierPick = state.supplierPick || "";
let pending = null;
let judged = [];
let judgedByKey = new Map();
let wordIndex = new Map();
const STOP = new Set("для или при под без шт штук упак упаковка набор все это про над".split(" "));
let matchFilter = "all";
let matchShown = 40;
let lastRead = "";
const HUMAN = {вручную: 1, проверка: 1, подтверждено: 1};
if (!state.migratedV3) {
  Object.keys(state.links).forEach(key => {
    if (HUMAN[state.methods[key]]) state.verified[key] = true;
    else {
      delete state.links[key];
      delete state.methods[key];
    }
  });
  state.migratedV3 = true;
}
if (!state.migratedV4) {
  state.cleared = {};
  state.migratedV4 = true;
  save();
}
function save() {
  localStorage.setItem(KEY, JSON.stringify({
    catalog: state.catalog, links: state.links, methods: state.methods,
    skip: state.skip, absent: state.absent, confirmed: state.confirmed, templates: state.templates,
    verified: state.verified, draft: state.draft, cleared: state.cleared, migratedV3: state.migratedV3, migratedV4: state.migratedV4, supplierPick: state.supplierPick || ""
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
function measures(value) {
  const text = String(value || "").toLowerCase().replace(/ё/g, "е").replace(/,/g, ".");
  const found = [];
  const pattern = /(\d+(?:\.\d+)?)\s*(килограммов|килограмма|килограмм|кило|кг|граммов|грамма|грамм|гр|г|миллилитров|миллилитра|миллилитр|мл|литров|литра|литр|л)(?![а-яa-z])/gi;
  let match = pattern.exec(text);
  while (match) {
    let amount = parseFloat(match[1]);
    const unit = match[2];
    let kind = "г";
    if (/^кило/.test(unit) || unit === "кг") amount *= 1000;
    else if (/^литр/.test(unit) || unit === "л") { amount *= 1000; kind = "мл"; }
    else if (/^милли/.test(unit) || unit === "мл") kind = "мл";
    found.push(kind + ":" + Math.round(amount));
    match = pattern.exec(text);
  }
  return [...new Set(found)];
}
function measureRelation(left, right) {
  const offer = measures(left);
  const ours = measures(right);
  if (offer.length && ours.length) return offer.some(item => ours.includes(item)) ? "same" : "conflict";
  return "open";
}
const MEASURE_WORDS = new Set("кг гр мл литр литра литров кило килограмм килограмма килограммов грамм грамма граммов".split(" "));
function sigTokens(value) {
  return tokens(value).filter(token => token.length >= 3 && !STOP.has(token) && !MEASURE_WORDS.has(token) && !/^\d/.test(token));
}
const SHORT_STEM = new Set("пос пор кон сти бел дет жид гел".split(" "));
function tokenFits(left, right) {
  if (left === right) return true;
  const short = left.length <= right.length ? left : right;
  const long = left.length <= right.length ? right : left;
  if (!long.startsWith(short)) return false;
  const extra = long.length - short.length;
  if (short.length >= 4 && extra <= 3) return true;
  return short.length === 3 && SHORT_STEM.has(short) && extra <= 5;
}
function indexKeys(token) {
  const keys = [token];
  if (token.length > 4) keys.push(token.slice(0, 4));
  if (token.length >= 5 && SHORT_STEM.has(token.slice(0, 3))) keys.push(token.slice(0, 3));
  return keys;
}
function nameScore(left, right) {
  const offer = sigTokens(left);
  const ours = sigTokens(right);
  if (!offer.length || !ours.length) return 0;
  let hit = 0;
  offer.forEach(token => { if (ours.some(item => tokenFits(token, item))) hit += 1; });
  return hit / offer.length;
}
function seedWord(name) {
  const words = tokens(name).filter(token => token.length >= 4 && !STOP.has(token) && !/^\d/.test(token));
  words.sort((a, b) => {
    const as = (wordIndex.get(a) || []).length || 9999;
    const bs = (wordIndex.get(b) || []).length || 9999;
    return as - bs || b.length - a.length;
  });
  return words[0] || "";
}
function numberList(value) {
  return (String(value || "").match(/\d+(?:[.,]\d+)?/g) || []).map(part => part.replace(",", "."));
}
function sameNumbers(left, right) {
  const a = numberList(left).sort().join("|");
  const b = numberList(right).sort().join("|");
  return Boolean(a) && a === b;
}
function numbersDiffer(left, right) {
  const a = numberList(left).sort().join("|");
  const b = numberList(right).sort().join("|");
  return Boolean(a) && Boolean(b) && a !== b;
}
function catalogByCode(code) {
  return state.catalog.find(item => item.code === code) || null;
}
function uniqueCatalog() {
  const seen = new Set();
  const list = [];
  state.catalog.forEach(item => {
    const name = String(item.name || "").trim();
    const code = String(item.code || "").trim();
    const mark = norm(name) + "\n" + (code && code !== name ? code : "");
    if (!name || seen.has(mark)) return;
    seen.add(mark);
    list.push(item);
  });
  return list;
}
function choiceLabel(item) {
  const name = String(item.name || "").trim();
  const code = String(item.code || "").trim();
  return esc(name) + (code && code !== name ? " <span class='muted'>" + esc(code) + "</span>" : "");
}
function tally(list) {
  const counts = {sure: 0, doubt: 0, none: 0, confirmed: 0, picked: 0};
  list.forEach(item => { counts[item.status] = (counts[item.status] || 0) + 1; });
  return counts;
}
function analyze() {
  judged = [];
  judgedByKey = new Map();
  const catalog = uniqueCatalog();
  const byBarcode = new Map();
  const byName = new Map();
  const index = new Map();
  catalog.forEach((item, indexItem) => {
    if (item.barcode) byBarcode.set(item.barcode, indexItem);
    const name = norm(item.name);
    if (name && !byName.has(name)) byName.set(name, indexItem);
    sigTokens(item.name).forEach(token => {
      indexKeys(token).forEach(keyName => {
        const list = index.get(keyName) || [];
        list.push(indexItem);
        index.set(keyName, list);
      });
    });
  });
  wordIndex = index;
  function push(item) {
    judged.push(item);
    judgedByKey.set(item.key, item);
  }
  uniqueOffers().forEach(row => {
    const key = offerKey(row);
    if (state.verified[key] && state.links[key]) {
      push({key, status: "confirmed", code: state.links[key], reason: ""});
      return;
    }
    if (state.draft[key]) {
      push({key, status: "picked", code: state.draft[key], reason: "вы выбрали, осталось подтвердить"});
      return;
    }
    if (state.cleared[key]) {
      push({key, status: "none", code: "", reason: ""});
      return;
    }
    const code = barcode(row[3]);
    let hit = null;
    let method = "";
    if (code && byBarcode.has(code)) {
      const indexItem = byBarcode.get(code);
      const item = catalog[indexItem];
      if (!state.skip[key + "\n" + item.code]) { hit = indexItem; method = "штрихкод"; }
    }
    if (hit == null) {
      const exact = byName.get(norm(row[1]));
      if (exact != null) {
        const item = catalog[exact];
        if (!state.skip[key + "\n" + item.code]) { hit = exact; method = "точное название"; }
      }
    }
    if (hit != null) {
      const item = catalog[hit];
      const problem = unitProblem(item.unit, row[4]);
      if (problem === "единица не совпала") push({key, status: "doubt", code: item.code, reason: method + ", единица не совпала"});
      else push({key, status: "sure", code: item.code, reason: method});
      return;
    }
    const candidates = new Map();
    const offerWords = sigTokens(row[1]);
    offerWords.forEach(token => indexKeys(token).forEach(keyName => (index.get(keyName) || []).forEach(indexItem => candidates.set(indexItem, 1))));
    let best = -1, bestScore = 0, bestPack = "open";
    candidates.forEach((_, indexItem) => {
      const item = catalog[indexItem];
      if (state.skip[key + "\n" + item.code]) return;
      const pack = measureRelation(row[1], item.name);
      if (pack === "conflict") return;
      const value = nameScore(row[1], item.name);
      if (value < 0.65) return;
      if (value > bestScore) { bestScore = value; best = indexItem; bestPack = pack; }
    });
    if (best < 0) { push({key, status: "none", code: "", reason: ""}); return; }
    const item = catalog[best];
    const problem = unitProblem(item.unit, row[4]);
    if (problem === "единица не совпала") { push({key, status: "doubt", code: item.code, reason: "единица не совпала"}); return; }
    const sure = bestPack === "same" && bestScore >= 0.9 && offerWords.length >= 2;
    push({key, status: sure ? "sure" : "doubt", code: item.code, reason: sure ? "название и фасовка совпали" : "похоже, проверьте"});
  });
}
function reportText() {
  if (!state.catalog.length) return "Сначала загрузите наш прайс.";
  const counts = tally(judged);
  return "Совпало, в заказе: " + (counts.sure || 0)
    + ". Сомнение, нужна галочка: " + (counts.doubt || 0)
    + ". Не сопоставлено: " + (counts.none || 0)
    + ". Подтверждено: " + ((counts.confirmed || 0) + (counts.picked || 0)) + ".";
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
  const counts = tally(judged);
  const ready = (counts.sure || 0) + (counts.confirmed || 0) + (counts.picked || 0);
  document.getElementById("today-stats").innerHTML = "<p>Строк цен: <b>" + rows.length + "</b>. Ваших товаров: <b>" + state.catalog.length + "</b>.</p>"
    + "<p>Совпало, в заказе: <b>" + (counts.sure || 0) + "</b>.</p>"
    + "<p>Сомнение, нужна галочка: <b>" + (counts.doubt || 0) + "</b>.</p>"
    + "<p>Не сопоставлено: <b>" + (counts.none || 0) + "</b>.</p>"
    + "<p>Подтверждено: <b>" + ((counts.confirmed || 0) + (counts.picked || 0)) + "</b>.</p>"
    + "<p><button class='action' id='go-match' type='button'>Открыть сверку</button> "
    + (ready ? "<button class='action' id='go-order' type='button'>Собрать заказ</button></p>" : "<button class='action' disabled>Собрать заказ</button></p><p class='warn'>В заказ попадают совпавшие строки и те, где на красном поле стоит «Да».</p>");
  const openMatch = document.getElementById("go-match");
  if (openMatch) openMatch.addEventListener("click", () => show("match"));
  const go = document.getElementById("go-order");
  if (go) go.addEventListener("click", () => show("order"));
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
function pickColumn(labels, needles, reject) {
  const skip = reject || [];
  for (let needleIndex = 0; needleIndex < needles.length; needleIndex += 1) {
    const needle = needles[needleIndex];
    for (let index = 0; index < labels.length; index += 1) {
      const label = String(labels[index] || "").replace(/\s+/g, " ");
      if (skip.some(word => label.includes(word))) continue;
      if (label.includes(needle)) return index;
    }
  }
  return -1;
}
function guessMap(labels) {
  const low = labels.map(item => cell(item).toLowerCase());
  return {
    name: pickColumn(low, ["наименование", "название", "product", "name", "номенклатура", "товар"], ["идентификатор", "штрих", "шк", "группа", "подгруппа", "марка", "код"]),
    price: pickColumn(low, ["цена", "price", "стоимость"], ["ндс"]),
    barcode: pickColumn(low, ["штрихкод", "штрих", "шк", "barcode", "ean"], []),
    code: pickColumn(low, ["артикул", "sku", "код"], ["штрих", "шк"]),
    unit: pickColumn(low, ["единица", "ед.", "ед ", "unit"], [])
  };
}
function headerLabels(grid, header) {
  const top = (grid[header] || []).map(item => cell(item));
  const next = grid[header + 1] || [];
  const nextJoined = next.map(item => cell(item).toLowerCase()).join(" ");
  const nameCol = column(top.map(item => item.toLowerCase()), ["товар", "наименование", "название"], "идентификатор");
  const nextName = nameCol >= 0 ? cell(next[nameCol]) : "x";
  if (!nextName && /цена|ед|штрих/.test(nextJoined)) return top.map((label, index) => label || cell(next[index]));
  return top;
}
function usableMap(labels, saved) {
  const guessed = guessMap(labels.map(item => item.toLowerCase()));
  if (!saved) return guessed;
  const nameLabel = (labels[saved.name] || "").toLowerCase().replace(/\s+/g, " ");
  if (nameLabel.includes("идентификатор")) return guessed;
  if (nameLabel.includes("код") && !nameLabel.includes("наимен")) return guessed;
  return saved;
}
function looksLikeHeader(row) {
  const joined = (row || []).map(item => cell(item).toLowerCase()).join(" ");
  if (/\d{8,14}/.test(joined.replace(/\s/g, ""))) return false;
  return /наимен|номенклат|товар|назван|штрих|артикул|цена|прайс/.test(joined);
}
function isBarcodeValue(value) {
  const text = cell(value).replace(/\s/g, "");
  return /^\d{8,14}$/.test(text);
}
function isPriceValue(value) {
  const text = cell(value).replace(/\s/g, "").replace(",", ".");
  if (!/^\d+(\.\d{1,2})?$/.test(text)) return false;
  if (isBarcodeValue(text)) return false;
  const number = Number(text);
  return number > 0 && number < 1000000;
}
function isNameValue(value) {
  const text = cell(value);
  return /[а-яёa-z]/i.test(text) && text.length >= 4 && !isBarcodeValue(text);
}
function detectPlain(grid, kind) {
  for (let index = 0; index < Math.min(grid.length, 15); index += 1) {
    if (looksLikeHeader(grid[index])) return null;
  }
  const sample = [];
  for (let index = 0; index < grid.length && sample.length < 40; index += 1) {
    if ((grid[index] || []).some(item => cell(item))) sample.push(grid[index]);
  }
  if (sample.length < 2) return null;
  const width = sample.reduce((max, row) => Math.max(max, row.length), 0);
  const cols = [];
  for (let index = 0; index < width; index += 1) {
    let names = 0, bars = 0, prices = 0, filled = 0;
    sample.forEach(row => {
      const value = row[index];
      if (!cell(value)) return;
      filled += 1;
      if (isBarcodeValue(value)) bars += 1;
      else if (isPriceValue(value)) prices += 1;
      else if (isNameValue(value)) names += 1;
    });
    cols.push({index, names, bars, prices, filled});
  }
  const nameCol = cols.filter(col => col.filled && col.names >= 3 && col.names >= col.filled * 0.6).sort((a, b) => b.names - a.names)[0];
  const barCol = cols.filter(col => col.filled && col.bars >= 3 && col.bars >= col.filled * 0.6).sort((a, b) => b.bars - a.bars)[0];
  const priceCol = cols.filter(col => col.filled && (!barCol || col.index !== barCol.index) && col.prices >= 3 && col.prices >= col.filled * 0.6).sort((a, b) => b.prices - a.prices)[0];
  if (!nameCol) return null;
  if (kind === "supplier" && !priceCol) return null;
  return {header: -1, map: {name: nameCol.index, price: priceCol ? priceCol.index : -1, barcode: barCol ? barCol.index : -1, code: -1, unit: -1}};
}
function sampleValues(grid, header, col) {
  const start = header < 0 ? 0 : header + 1;
  const values = [];
  for (let index = start; index < grid.length && values.length < 24; index += 1) {
    const value = cell((grid[index] || [])[col]);
    if (value) values.push(value);
  }
  return values;
}
function looksLikeCode(value) {
  return /^\d{4,14}$/.test(cell(value).replace(/\s/g, ""));
}
function nameColumnOk(values) {
  if (values.length < 3) return false;
  const words = values.filter(isNameValue).length;
  const codes = values.filter(looksLikeCode).length;
  return words >= values.length * 0.6 && codes < values.length * 0.5;
}
function priceColumnOk(values) {
  if (values.length < 3) return false;
  const prices = values.map(value => {
    const text = cell(value).replace(/\s/g, "").replace(",", ".");
    return isPriceValue(text) ? Number(text) : null;
  }).filter(value => value != null);
  if (prices.length < values.length * 0.4) return false;
  const tiny = prices.filter(value => value < 1).length;
  return tiny < prices.length * 0.8;
}
function barcodeColumnOk(values) {
  if (values.length < 3) return false;
  return values.filter(isBarcodeValue).length >= values.length * 0.6;
}
function gridWidth(grid) {
  return grid.reduce((max, row) => Math.max(max, (row || []).length), 0);
}
function bestNameColumn(grid, header) {
  let best = -1;
  let bestScore = 0;
  for (let index = 0; index < gridWidth(grid); index += 1) {
    const values = sampleValues(grid, header, index);
    const words = values.filter(isNameValue).length;
    if (!nameColumnOk(values) || words <= bestScore) continue;
    bestScore = words;
    best = index;
  }
  return best;
}
function bestPriceColumn(grid, header, avoid) {
  let best = -1;
  let bestScore = 0;
  for (let index = 0; index < gridWidth(grid); index += 1) {
    if (avoid.indexOf(index) >= 0) continue;
    const values = sampleValues(grid, header, index);
    if (!priceColumnOk(values)) continue;
    const prices = values.filter(isPriceValue).length;
    if (prices <= bestScore) continue;
    bestScore = prices;
    best = index;
  }
  return best;
}
function repairMap(grid, header, map, kind) {
  const copy = {name: map.name, price: map.price, barcode: map.barcode, code: map.code, unit: map.unit};
  if (!nameColumnOk(sampleValues(grid, header, copy.name))) {
    const better = bestNameColumn(grid, header);
    if (better >= 0) copy.name = better;
  }
  if (copy.barcode >= 0 && !barcodeColumnOk(sampleValues(grid, header, copy.barcode))) copy.barcode = -1;
  if (kind === "supplier" && !priceColumnOk(sampleValues(grid, header, copy.price))) {
    const better = bestPriceColumn(grid, header, [copy.barcode]);
    if (better >= 0) copy.price = better;
  }
  return copy;
}
function layoutOk(grid, header, map, kind) {
  if (!nameColumnOk(sampleValues(grid, header, map.name))) return false;
  if (kind === "supplier" && !priceColumnOk(sampleValues(grid, header, map.price))) return false;
  return true;
}
function catalogLooksLikeIds() {
  const sample = state.catalog.slice(0, 30);
  if (sample.length < 5) return false;
  const ids = sample.filter(item => /^[0-9a-f]{8}-[0-9a-f-]{20,}$/i.test(String(item.name).trim())).length;
  return ids >= sample.length / 2;
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
    state.verified = {};
    state.draft = {};
    state.cleared = {};
    state.skip = {};
    const seen = new Set();
    for (let index = header + 1; index < grid.length; index += 1) {
      const row = grid[index];
      const name = cell(row[map.name]);
      if (!name) continue;
      const code = map.code >= 0 ? cell(row[map.code]) || name : name;
      const mark = norm(name) + "\n" + (code.trim() === name.trim() ? "" : code.trim());
      if (seen.has(mark)) continue;
      seen.add(mark);
      state.catalog.push({
        code,
        barcode: map.barcode >= 0 ? barcode(row[map.barcode]) : "",
        name, unit: map.unit >= 0 ? cell(row[map.unit]) : ""
      });
    }
    analyze();
    save();
    const examples = state.catalog.slice(0, 3).map(item => item.name).join(" · ");
    lastRead = examples ? "Пример: " + examples : "";
    document.getElementById("report").textContent = "Наш прайс загружен, " + state.catalog.length + " товаров. " + lastRead + " " + reportText();
    renderToday();
    show("match");
    return;
  }
  let added = 0;
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (rows[index][8] === file.name) rows.splice(index, 1);
  }
  for (let index = header + 1; index < grid.length; index += 1) {
    const row = grid[index];
    const name = cell(row[map.name]);
    const price = map.price >= 0 ? Number(String(row[map.price] || "").replace(/\s/g, "").replace(",", ".").replace(/[^\d.\-]/g, "")) : 0;
    if (!name || !price) continue;
    rows.push([typed, name, map.code >= 0 ? cell(row[map.code]) : "", map.barcode >= 0 ? barcode(row[map.barcode]) : "", map.unit >= 0 ? cell(row[map.unit]) : "", price, today(), "Цена", file.name]);
    added += 1;
  }
  if (state.catalog.length) analyze();
  const examples = rows.slice(Math.max(rows.length - added, 0), Math.max(rows.length - added, 0) + 3).map(row => row[1]).join(" · ");
  lastRead = examples ? "Пример: " + examples : "";
  document.getElementById("report").textContent = "Прайс загружен: " + added + " строк. " + lastRead + " " + (state.catalog.length ? reportText() : "Наш прайс ещё не загружен, связка будет после него.");
  if (state.catalog.length) show("match");
  else renderToday();
}
function startImport(file, kind) {
  readGrid(file, (grid, error) => {
    if (!grid) {
      document.getElementById("report").textContent = error;
      return;
    }
    const plain = detectPlain(grid, kind);
    if (plain) {
      const map = repairMap(grid, plain.header, plain.map, kind);
      if (layoutOk(grid, plain.header, map, kind)) {
        applyRows(grid, plain.header, map, file, kind);
        document.getElementById("mapper").classList.add("hidden");
        return;
      }
    }
    const header = findHeader(grid) >= 0 ? findHeader(grid) : guessHeader(grid);
    const labels = headerLabels(grid, header);
    const saved = state.templates[headerKey(labels)];
    const map = repairMap(grid, header, usableMap(labels, saved), kind);
    pending = {grid, header, file, kind, labels};
    if (layoutOk(grid, header, map, kind)) {
      applyRows(grid, header, map, file, kind);
      document.getElementById("mapper").classList.add("hidden");
      return;
    }
    fillMapper(grid, header, map);
    document.getElementById("report").textContent = "Не получилось уверенно прочитать колонки. Проверьте, где название" + (kind === "supplier" ? " и цена" : "") + ". Ниже первые строки файла.";
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
function esc(value) {
  return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}
const STATUS_TEXT = {
  sure: "Совпало",
  doubt: "Сомнение",
  none: "Не сопоставлено",
  confirmed: "Подтверждено",
  picked: "Выбрано"
};
function passesFilter(item) {
  if (matchFilter === "all") return true;
  if (matchFilter === "sure") return item.status === "sure" || item.status === "picked";
  return item.status === matchFilter;
}
function renderMatch() {
  const offers = uniqueOffers();
  const suppliers = [...new Set([...offers.values()].map(row => row[0]))].sort();
  const select = document.getElementById("match-supplier");
  if (!state.supplierPick || !suppliers.includes(state.supplierPick)) state.supplierPick = suppliers[0] || "";
  select.innerHTML = suppliers.map(name => "<option value=\"" + esc(name) + "\"" + (name === state.supplierPick ? " selected" : "") + ">" + esc(name) + "</option>").join("");
  const filters = [
    ["all", "Все"],
    ["sure", "Совпало"],
    ["doubt", "Сомнение"],
    ["none", "Не сопоставлено"],
    ["confirmed", "Подтверждено"]
  ];
  document.getElementById("match-filters").innerHTML = filters.map(([id, label]) => "<button class='filter" + (matchFilter === id ? " active" : "") + "' type='button' data-filter='" + id + "'>" + label + "</button>").join("");
  const query = (document.getElementById("match-search").value || "").trim().toLowerCase();
  const visible = judged.filter(item => {
    const row = offers.get(item.key);
    if (!row || row[0] !== state.supplierPick) return false;
    if (!passesFilter(item)) return false;
    if (query && !row[1].toLowerCase().includes(query)) return false;
    return true;
  });
  const counts = tally(visible);
  document.getElementById("match-counts").textContent = state.catalog.length
    ? (lastRead ? lastRead + " " : "") + "У этого поставщика в фильтре: " + visible.length + ". Совпало: " + (counts.sure || 0) + ". Сомнение: " + (counts.doubt || 0) + ". Не сопоставлено: " + (counts.none || 0) + ". Подтверждено: " + ((counts.confirmed || 0) + (counts.picked || 0)) + "."
      + (catalogLooksLikeIds() ? " Названия не загрузились: в справочнике коды вида 444da418-…. Откройте «Сегодня», в пункте «1. Наш прайс» ещё раз выберите файл Остатки.xls." : "")
    : "Сначала загрузите наш прайс на вкладке «Сегодня».";
  const slice = visible.slice(0, matchShown);
  document.getElementById("match-body").innerHTML = slice.map(item => {
    const row = offers.get(item.key);
    const ours = item.code ? catalogByCode(item.code) : null;
    const oursName = ours && !catalogLooksLikeIds() ? ours.name : "";
    const oursText = "<textarea class='row-search" + (item.status === "doubt" ? " doubt" : "") + "' rows='3' autocomplete='off' data-key='" + encodeURIComponent(item.key) + "' placeholder='чайка'>" + esc(oursName) + "</textarea>";
    const reason = item.reason ? "<br><span class='warn'>" + esc(item.reason) + "</span>" : "";
    const noButton = "<button class='secondary no' type='button' data-key='" + encodeURIComponent(item.key) + "' data-code='" + encodeURIComponent(item.code || "") + "'>Нет</button>";
    const tick = item.status === "doubt" || item.status === "picked"
      ? "<label><input class='tick' type='checkbox' data-key='" + encodeURIComponent(item.key) + "' data-code='" + encodeURIComponent(item.code) + "'" + (item.status === "confirmed" ? " checked" : "") + "> Да</label> " + noButton
      : (item.status === "confirmed"
        ? "<label><input class='tick' type='checkbox' data-key='" + encodeURIComponent(item.key) + "' data-code='" + encodeURIComponent(item.code) + "' checked> Да</label> " + noButton
        : noButton);
    const rowClass = item.status === "doubt" ? "doubt" : (item.status === "sure" || item.status === "confirmed" ? "done" : "need");
    return "<tr class='" + rowClass + "'><td>" + esc(row[1]) + "</td><td>" + esc(row[0]) + "</td><td>" + oursText + "</td><td>"
      + esc(STATUS_TEXT[item.status] || item.status) + reason + "</td><td>" + tick + "</td></tr>";
  }).join("") || "<tr><td colspan='5'>В этом фильтре пусто.</td></tr>";
  const more = document.getElementById("match-more");
  more.classList.toggle("hidden", visible.length <= slice.length);
  more.textContent = "Показать ещё (" + Math.max(visible.length - slice.length, 0) + ")";
  document.querySelectorAll("textarea.row-search").forEach(growSearch);
}
function growSearch(field) {
  field.style.height = "auto";
  field.style.height = Math.max(field.scrollHeight + 2, 72) + "px";
}
document.getElementById("match-supplier").addEventListener("change", () => {
  state.supplierPick = document.getElementById("match-supplier").value;
  matchShown = 40;
  save();
  renderMatch();
});
document.getElementById("match-filters").addEventListener("click", event => {
  const button = event.target.closest("button");
  if (!button) return;
  matchFilter = button.dataset.filter;
  matchShown = 40;
  renderMatch();
});
document.getElementById("match-search").addEventListener("input", () => { matchShown = 40; renderMatch(); });
document.getElementById("match-more").addEventListener("click", () => { matchShown += 40; renderMatch(); });
document.getElementById("match-body").addEventListener("click", event => {
  const button = event.target.closest("button.no");
  if (!button) return;
  const key = decodeURIComponent(button.dataset.key);
  const code = decodeURIComponent(button.dataset.code || "");
  if (code) state.skip[key + "\n" + code] = true;
  state.cleared[key] = true;
  delete state.draft[key];
  delete state.verified[key];
  delete state.links[key];
  delete state.methods[key];
  save();
  analyze();
  renderMatch();
});
document.getElementById("match-body").addEventListener("change", event => {
  const box = event.target;
  if (!box.classList.contains("tick")) return;
  const key = decodeURIComponent(box.dataset.key);
  const code = decodeURIComponent(box.dataset.code || "");
  if (box.checked && code) {
    state.links[key] = code;
    state.verified[key] = true;
    state.methods[key] = "подтверждено";
    delete state.draft[key];
    delete state.cleared[key];
    const item = judgedByKey.get(key);
    if (item) { item.status = "confirmed"; item.code = code; item.reason = ""; }
  } else {
    delete state.verified[key];
    delete state.links[key];
    delete state.methods[key];
    delete state.draft[key];
    save();
    analyze();
    renderMatch();
    return;
  }
  save();
  renderMatch();
});
let activeKey = "";
function hideDrop() {
  document.getElementById("inline-drop").classList.add("hidden");
}
function showDrop(field) {
  const drop = document.getElementById("inline-drop");
  const query = field.value.trim();
  const parts = norm(query).split(" ").filter(Boolean);
  const rect = field.getBoundingClientRect();
  drop.style.top = (rect.bottom + 4) + "px";
  drop.style.left = rect.left + "px";
  drop.style.width = Math.max(rect.width, 380) + "px";
  if (catalogLooksLikeIds()) {
    drop.innerHTML = "<p class='warn'>«" + esc(query) + "» не находится: в справочник попали коды, а не названия. Откройте «Сегодня» и в пункте «1. Наш прайс» ещё раз выберите Остатки.xls.</p>";
    drop.classList.remove("hidden");
    return;
  }
  if (!parts.length || parts.join("").length < 3) { hideDrop(); return; }
  if (!state.catalog.length) {
    drop.innerHTML = "<p class='warn'>Справочник пуст. Загрузите наш прайс.</p>";
    drop.classList.remove("hidden");
    return;
  }
  const found = uniqueCatalog().filter(item => {
    const blob = (item.name + " " + item.code + " " + (item.barcode || "")).toLowerCase();
    return parts.every(part => blob.includes(part));
  });
  const slice = found.slice(0, 30);
  drop.innerHTML = found.length
    ? "<p class='muted'>Найдено " + found.length + (found.length > slice.length ? ", показаны первые " + slice.length : "") + ".</p>"
      + slice.map(item => "<button class='quiet choose' type='button' data-code='" + encodeURIComponent(item.code) + "'>" + choiceLabel(item) + "</button>").join("")
    : "<p class='muted'>В справочнике нет «" + esc(query) + "».</p>";
  drop.classList.remove("hidden");
}
document.getElementById("match-body").addEventListener("input", event => {
  const field = event.target.closest("textarea.row-search");
  if (!field) return;
  activeKey = decodeURIComponent(field.dataset.key);
  growSearch(field);
  showDrop(field);
});
document.getElementById("match-body").addEventListener("focusin", event => {
  const field = event.target.closest("textarea.row-search");
  if (!field) return;
  activeKey = decodeURIComponent(field.dataset.key);
  showDrop(field);
});
document.getElementById("inline-drop").addEventListener("mousedown", event => {
  const button = event.target.closest("button.choose");
  if (!button || !activeKey) return;
  event.preventDefault();
  const code = decodeURIComponent(button.dataset.code);
  delete state.verified[activeKey];
  delete state.cleared[activeKey];
  state.links[activeKey] = code;
  state.verified[activeKey] = true;
  state.methods[activeKey] = "подтверждено";
  delete state.draft[activeKey];
  const item = judgedByKey.get(activeKey);
  if (item) { item.status = "confirmed"; item.code = code; item.reason = ""; }
  save();
  hideDrop();
  renderMatch();
});
document.addEventListener("mousedown", event => {
  if (event.target.closest("#inline-drop") || event.target.closest("textarea.row-search")) return;
  hideDrop();
});
function linkedOffers() {
  const map = new Map();
  judged.forEach(item => {
    if (state.cleared[item.key]) return;
    if ((item.status === "sure" || item.status === "confirmed" || item.status === "picked") && item.code) map.set(item.key, item.code);
  });
  return map;
}
function renderOrder() {
  const box = document.getElementById("order-body");
  const links = linkedOffers();
  const codes = [...new Set([...links.values()])];
  if (!codes.length) {
    box.innerHTML = "<div class='card'><button class='action' disabled>Собрать заказ</button><p class='warn'>В заказ попадают совпавшие строки и те, где на красном поле стоит «Да».</p></div>";
    return;
  }
  const catalog = new Map(state.catalog.map(item => [item.code, item]));
  const offers = uniqueOffers();
  const day = today();
  box.innerHTML = codes.slice(0, 80).map(code => {
    const item = catalog.get(code) || {name: code, unit: ""};
    const group = [...offers.values()].filter(row => links.get(offerKey(row)) === code);
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
save();
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
