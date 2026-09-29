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
  tr.review { background: #fff7ed; }
  tr.conflict { background: #fff1f2; }
  tr.shift { background: #fff7ed; }
  button.yes { margin: 0 8px 8px 0; }
  button.pick { margin: 0 8px 8px 0; }
  #price-detail table td { vertical-align: top; }
  button.linkish { background: transparent; border: 0; color: #16324f; text-decoration: underline; cursor: pointer; font: inherit; font-size: 16px; padding: 0; }
  .warn { color: #9a3412; }
  .ok { color: #166534; }
  h2 { font-size: 20px; margin: 0 0 8px; }
  .muted { color: #4b5563; }
  .hidden { display: none; }
  #picker-list { max-height: 320px; overflow: auto; }
  #match-search { width: 360px; }
  button.no { margin-left: 8px; padding: 6px 12px; }
  #match table { table-layout: fixed; }
  #match th:nth-child(1), #match td:nth-child(1) { width: 18%; }
  #match th:nth-child(2), #match td:nth-child(2) { width: 11%; }
  #match th:nth-child(3), #match td:nth-child(3) { width: 20%; }
  #match th:nth-child(4), #match td:nth-child(4) { width: 8%; }
  #match th:nth-child(5), #match td:nth-child(5) { width: 18%; }
  #match th:nth-child(6), #match td:nth-child(6) { width: 11%; }
  #match th:nth-child(7), #match td:nth-child(7) { width: 14%; }
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
    <div class="card" id="today-feed"></div>
    <div class="card">
      <button id="show-upload" class="quiet" type="button">+ Загрузить прайс вручную</button>
      <div id="upload-box" class="hidden">
        <h2>Прайс поставщика</h2>
        <p>Файл может быть в любом привычном виде: программа сама найдёт название и цену. Это не список ваших товаров.</p>
        <p>Поставщик, если в файле нет такой колонки <input id="supplier-name" type="text" placeholder="Например, Петя"></p>
        <input id="supplier-file" type="file">
        <p id="report" class="muted">Excel, CSV, ODS или скачанная страница из Google Таблиц.</p>
      </div>
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
    <div class="card" id="catalog-card">
      <h2>Наша номенклатура</h2>
      <p class="muted">Список товаров из 1С. Его загружают один раз, сопоставления остаются. Файл выбирают снова только если в 1С появились новые товары.</p>
      <div id="catalog-status"></div>
      <p id="our-file-wrap"><input id="our-file" type="file"></p>
      <p><button id="replace-catalog" class="quiet hidden" type="button">Обновить номенклатуру из 1С</button></p>
    </div>
  </section>
  <section id="prices" class="hidden">
    <h1>Прайсы</h1>
    <div class="card" id="supplier-summary"></div>
    <div class="card hidden" id="price-detail">
      <h2 id="price-title"></h2>
      <p><input id="price-search" type="text" placeholder="Найти в этом прайсе"></p>
      <p id="price-filters"></p>
      <table>
        <thead><tr><th>Номенклатура поставщика</th><th>Цена</th><th>Наша номенклатура</th><th>Состояние</th></tr></thead>
        <tbody id="price-body"></tbody>
      </table>
      <p><button id="price-more" class="secondary" type="button">Показать ещё</button></p>
    </div>
  </section>
  <section id="match" class="hidden">
    <h1>Номенклатура</h1>
    <div class="card">
      <p>Здесь только то, что нужно решить. Уже подтверждённые товары программа узнаёт сама и повторно не спрашивает. Их можно посмотреть кнопкой «Подтверждено».</p>
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
            <th>Предлагаемая наша номенклатура</th>
            <th>Уверенность</th>
            <th>Почему</th>
            <th>Статус</th>
            <th>Действие</th>
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
state.bonds = state.bonds || {};
state.missing = state.missing || {};
state.seen = state.seen || {};
state.fresh = state.fresh || {};
state.inbox = state.inbox || [];
state.uploads = state.uploads || [];
state.supplierPick = state.supplierPick || "";
let pending = null;
let judged = [];
let judgedByKey = new Map();
let latestByKey = new Map();
let wordIndex = new Map();
let analyzed = false;
const STOP = new Set("для или при под без шт штук упак упаковка набор все это про над".split(" "));
const CYR_SOUND = {а:"a", б:"b", в:"v", г:"g", д:"d", е:"e", ё:"e", ж:"zh", з:"z", и:"i", й:"y", к:"k", л:"l", м:"m", н:"n", о:"o", п:"p", р:"r", с:"s", т:"t", у:"u", ф:"f", х:"h", ц:"ts", ч:"ch", ш:"sh", щ:"sch", ъ:"", ы:"i", ь:"", э:"e", ю:"yu", я:"ya"};
let matchFilter = "attention";
let priceFilter = "all";
let catalogReplaceOpen = false;
let uploadOpen = false;
let priceOpen = "";
let matchShown = 40;
let priceShown = 40;
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
  const payload = {
    catalog: state.catalog, links: state.links, methods: state.methods,
    skip: state.skip, absent: state.absent, confirmed: state.confirmed, templates: state.templates,
    verified: state.verified, draft: state.draft, cleared: state.cleared,
    bonds: state.bonds, missing: state.missing, seen: state.seen, fresh: state.fresh,
    inbox: state.inbox, uploads: state.uploads,
    migratedV3: state.migratedV3, migratedV4: state.migratedV4, migratedV5: state.migratedV5, seenReady: state.seenReady,
    supplierPick: state.supplierPick || ""
  };
  try {
    localStorage.setItem(KEY, JSON.stringify(payload));
  } catch (error) {
    payload.uploads = [];
    try { localStorage.setItem(KEY, JSON.stringify(payload)); } catch (again) { /* связи важнее копий прайсов */ }
  }
}
function bondRecord(catalogCode, row) {
  return {catalogCode: catalogCode, nameNorm: norm(row[1]), barcode: barcode(row[3]), supplierCode: cell(row[2]), at: today(), who: ""};
}
function putBond(row, catalogCode) {
  const supplier = row[0];
  const record = bondRecord(catalogCode, row);
  const bar = barcode(row[3]);
  const code = cell(row[2]);
  state.bonds[supplier + "\n" + stableId(row)] = record;
  if (bar) state.bonds[supplier + "\nb:" + bar] = record;
  if (code) state.bonds[supplier + "\nc:" + norm(code)] = record;
  if (record.nameNorm) state.bonds[supplier + "\nn:" + record.nameNorm] = record;
  const key = offerKey(row);
  state.links[key] = catalogCode;
  state.verified[key] = true;
  state.methods[key] = "подтверждено";
  delete state.draft[key];
  delete state.cleared[key];
  delete state.missing[supplier + "\n" + stableId(row)];
  delete state.fresh[supplier + "\n" + stableId(row)];
}
function findBond(row) {
  const supplier = row[0];
  const bar = barcode(row[3]);
  const code = cell(row[2]);
  const keys = [stableId(row)];
  if (bar) keys.push("b:" + bar);
  if (code) keys.push("c:" + norm(code));
  for (let index = 0; index < keys.length; index += 1) {
    const hit = state.bonds[supplier + "\n" + keys[index]];
    if (hit && hit.catalogCode) return hit;
  }
  const byName = state.bonds[supplier + "\nn:" + norm(row[1])];
  if (byName && byName.catalogCode) {
    const codeChanged = Boolean(byName.supplierCode && code && norm(byName.supplierCode) !== norm(code));
    const barChanged = Boolean(byName.barcode && bar && byName.barcode !== bar);
    if (!codeChanged && !barChanged) return byName;
  }
  const legacy = offerKey(row);
  if (state.verified[legacy] && state.links[legacy]) return {catalogCode: state.links[legacy], nameNorm: norm(row[1]), barcode: bar, supplierCode: code, at: "", who: ""};
  return null;
}
function shiftedBond(row) {
  const hit = state.bonds[row[0] + "\nn:" + norm(row[1])];
  if (!hit || !hit.catalogCode) return null;
  const code = cell(row[2]);
  const bar = barcode(row[3]);
  const codeChanged = Boolean(hit.supplierCode && code && norm(hit.supplierCode) !== norm(code));
  const barChanged = Boolean(hit.barcode && bar && hit.barcode !== bar);
  if (!codeChanged && !barChanged) return null;
  return hit;
}
function isMissing(row) {
  return Boolean(state.missing[row[0] + "\n" + stableId(row)]);
}
if (!state.migratedV5) {
  Object.keys(state.verified || {}).forEach(key => {
    if (!state.verified[key] || !state.links[key]) return;
    const parts = key.split("|");
    const supplier = parts[0] || "";
    const code = parts[1] || "";
    const name = parts.slice(2).join("|");
    const record = {catalogCode: state.links[key], nameNorm: norm(name), barcode: "", supplierCode: code, at: "", who: ""};
    if (code) state.bonds[supplier + "\nc:" + norm(code)] = record;
    if (name) state.bonds[supplier + "\nn:" + norm(name)] = record;
  });
  Object.keys(state.cleared || {}).forEach(key => {
    if (!state.cleared[key]) return;
    const parts = key.split("|");
    const supplier = parts[0] || "";
    const code = parts[1] || "";
    const name = parts.slice(2).join("|");
    if (code) state.missing[supplier + "\nc:" + norm(code)] = {at: ""};
    if (name) state.missing[supplier + "\nn:" + norm(name)] = {at: ""};
  });
  state.migratedV5 = true;
  save();
}
function today() {
  const d = new Date();
  return String(d.getDate()).padStart(2,"0") + "." + String(d.getMonth()+1).padStart(2,"0") + "." + d.getFullYear();
}
function clock() {
  const d = new Date();
  return String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
}
function offerKey(row) { return [row[0], row[2], row[1]].join("|"); }
function stableId(row) {
  const bar = barcode(row[3]);
  if (bar) return "b:" + bar;
  const code = cell(row[2]);
  if (code) return "c:" + norm(code);
  return "n:" + norm(row[1]);
}
function rowStamp(row, index) {
  const parts = String(row[6] || "").split(".");
  const day = parts.length === 3 ? (Number(parts[2]) * 10000 + Number(parts[1]) * 100 + Number(parts[0])) : 0;
  return day * 1000000 + index;
}
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
  const pattern = /(\d+(?:\.\d+)?)\s*(килограммов|килограмма|килограмм|кило|кг|kg|граммов|грамма|грамм|гр|г|g|миллилитров|миллилитра|миллилитр|мл|ml|литров|литра|литр|л|l)(?![а-яa-z])/gi;
  let match = pattern.exec(text);
  while (match) {
    let amount = parseFloat(match[1]);
    const unit = match[2].toLowerCase();
    let kind = "г";
    if (/^кило/.test(unit) || unit === "кг" || unit === "kg") amount *= 1000;
    else if (/^литр/.test(unit) || unit === "л" || unit === "l") { amount *= 1000; kind = "мл"; }
    else if (/^милли/.test(unit) || unit === "мл" || unit === "ml") kind = "мл";
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
const PACK_WORDS = new Set("pe bag пэ мешок пакет кор коробка пленка".split(" "));
function sigTokens(value) {
  return tokens(value).filter(token => token.length >= 3 && !STOP.has(token) && !MEASURE_WORDS.has(token) && !PACK_WORDS.has(token) && !/^\d/.test(token) && !/^[xх]\d+$/.test(token));
}
const SHORT_STEM = new Set("пос пор кон сти бел дет жид гел авт".split(" "));
function soundKey(token) {
  let out = "";
  const text = String(token || "").toLowerCase();
  for (let index = 0; index < text.length; index += 1) {
    const letter = text[index];
    out += Object.prototype.hasOwnProperty.call(CYR_SOUND, letter) ? CYR_SOUND[letter] : letter;
  }
  out = out.replace(/ph/g, "f").replace(/ck/g, "k").replace(/qu/g, "kv").replace(/j/g, "y").replace(/w/g, "v").replace(/x/g, "ks").replace(/c/g, "k");
  out = out.replace(/ai/g, "ey").replace(/ay/g, "ey").replace(/ei/g, "ey").replace(/y$/g, "i");
  return out.replace(/([a-z])\1+/g, "$1");
}
function tokenFits(left, right) {
  if (left === right) return true;
  const rawShort = left.length <= right.length ? left : right;
  const rawLong = left.length <= right.length ? right : left;
  if (rawLong.startsWith(rawShort) && rawShort.length === 3 && SHORT_STEM.has(rawShort) && rawLong.length - rawShort.length <= 5) return true;
  if (left.length >= 4 && right.length >= 4 && left.slice(0, 4) === right.slice(0, 4)) return true;
  const leftSound = soundKey(left);
  const rightSound = soundKey(right);
  if (leftSound && leftSound === rightSound) return true;
  const short = leftSound.length <= rightSound.length ? leftSound : rightSound;
  const long = leftSound.length <= rightSound.length ? rightSound : leftSound;
  if (!short || !long.startsWith(short)) return false;
  const extra = long.length - short.length;
  if (short.length >= 4 && extra <= 3) return true;
  return short.length >= 5 && extra <= 6;
}
function indexKeys(token) {
  const sound = soundKey(token);
  const keys = [];
  [token, sound].forEach(value => {
    if (!value || keys.includes(value)) return;
    keys.push(value);
    if (value.length > 4) keys.push(value.slice(0, 4));
  });
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
function pieceCount(value) {
  const text = String(value || "").toLowerCase().replace(/ё/g, "е");
  const after = text.match(/(?:x|х|\*|pcs)\s*(\d+)/);
  if (after) return Number(after[1]);
  const before = text.match(/(\d+)\s*(?:шт|штук|pcs)(?![а-яa-z])/);
  return before ? Number(before[1]) : 0;
}
function explainMatch(offerName, ourName, relation) {
  if (relation === "conflict" || relation === "pieces") return conflictReason(offerName, ourName);
  const bits = [];
  sigTokens(offerName).forEach(token => {
    const hit = sigTokens(ourName).find(item => tokenFits(token, item));
    if (!hit) return;
    if (norm(token) !== norm(hit)) bits.push(token.toUpperCase() + " совпадает с " + hit);
    else bits.push(hit);
  });
  if (relation === "same") bits.push(measureLabel(measures(offerName)));
  return bits.slice(0, 6).join(", ");
}
function criticalRelation(left, right) {
  const pack = measureRelation(left, right);
  if (pack === "conflict") return "conflict";
  const offerPieces = pieceCount(left);
  const ourPieces = pieceCount(right);
  if (offerPieces && ourPieces && offerPieces !== ourPieces) return "pieces";
  return pack;
}
function measureLabel(list) {
  return list.map(item => {
    const parts = item.split(":");
    const amount = Number(parts[1]);
    if (parts[0] === "г" && amount >= 1000 && amount % 1000 === 0) return (amount / 1000) + " кг";
    if (parts[0] === "мл" && amount >= 1000 && amount % 1000 === 0) return (amount / 1000) + " л";
    return amount + " " + parts[0];
  }).join(", ");
}
function conflictReason(offerName, ourName) {
  const offer = measures(offerName);
  const ours = measures(ourName);
  if (offer.length && ours.length && !offer.some(item => ours.includes(item))) {
    return "Фасовка не совпадает: поставщик — " + measureLabel(offer) + ", наша номенклатура — " + measureLabel(ours) + ".";
  }
  const offerPieces = pieceCount(offerName);
  const ourPieces = pieceCount(ourName);
  if (offerPieces && ourPieces && offerPieces !== ourPieces) {
    return "Количество не совпадает: поставщик — " + offerPieces + " шт, наша номенклатура — " + ourPieces + " шт.";
  }
  return "Критическая характеристика не совпадает.";
}
function rebuildLatest() {
  const byIdentity = new Map();
  rows.forEach((row, index) => {
    const id = row[0] + "\n" + stableId(row);
    const prev = byIdentity.get(id);
    if (!prev || rowStamp(row, index) >= rowStamp(prev.row, prev.index)) byIdentity.set(id, {row, index});
  });
  latestByKey = new Map();
  byIdentity.forEach(item => latestByKey.set(offerKey(item.row), item.row));
}
function isFresh(row) {
  return Boolean(row && state.fresh[row[0] + "\n" + stableId(row)]);
}
function needsPerson(item, row) {
  if (!item || !row) return false;
  if (item.status === "confirmed" || item.status === "missing") return false;
  if (item.status === "review" || item.status === "conflict" || item.status === "shift") return true;
  return isFresh(row);
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
  rebuildLatest();
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
  latestByKey.forEach(row => {
    const key = offerKey(row);
    const saved = findBond(row);
    if (saved) {
      push({key, status: "confirmed", code: saved.catalogCode, score: 100, reason: "сопоставление уже подтверждено"});
      return;
    }
    if (isMissing(row)) {
      push({key, status: "missing", code: "", score: 0, reason: "отметили, что такого товара у нас нет"});
      return;
    }
    const moved = shiftedBond(row);
    if (moved) {
      push({key, status: "shift", code: moved.catalogCode, score: 0, reason: "Возможно, изменился идентификатор товара поставщика. Прежнее соответствие сохранено, подтвердите его или выберите другой товар."});
      return;
    }
    if (state.draft[key] && state.links[key]) {
      push({key, status: "review", code: state.links[key], score: 0, reason: "товар выбран, его ещё не подтвердили"});
      return;
    }
    const code = barcode(row[3]);
    if (code && byBarcode.has(code)) {
      const item = catalog[byBarcode.get(code)];
      push({key, status: "sure", code: item.code, score: 100, reason: "совпал штрихкод"});
      return;
    }
    const exact = byName.get(norm(row[1]));
    if (exact != null) {
      const item = catalog[exact];
      const relation = criticalRelation(row[1], item.name);
      if (relation === "conflict" || relation === "pieces") push({key, status: "conflict", code: item.code, score: 100, reason: conflictReason(row[1], item.name)});
      else push({key, status: "sure", code: item.code, score: 100, reason: "название совпало"});
      return;
    }
    const candidates = new Set();
    const offerWords = sigTokens(row[1]);
    const take = (list, limit) => {
      const size = limit || list.length;
      for (let index = 0; index < list.length && index < size; index += 1) candidates.add(list[index]);
    };
    offerWords.forEach(token => indexKeys(token).forEach(keyName => {
      const list = index.get(keyName) || [];
      if (list.length && list.length <= 250) take(list);
    }));
    if (!candidates.size && offerWords.length) {
      let rarest = null;
      offerWords.forEach(token => indexKeys(token).forEach(keyName => {
        const list = index.get(keyName) || [];
        if (list.length && (!rarest || list.length < rarest.length)) rarest = list;
      }));
      if (rarest) take(rarest, 400);
    }
    const ranked = [];
    candidates.forEach(indexItem => {
      const item = catalog[indexItem];
      const value = nameScore(row[1], item.name);
      if (value < 0.4) return;
      ranked.push({indexItem, value, relation: criticalRelation(row[1], item.name)});
    });
    ranked.sort((a, b) => {
      const aBad = a.relation === "conflict" || a.relation === "pieces" ? 1 : 0;
      const bBad = b.relation === "conflict" || b.relation === "pieces" ? 1 : 0;
      return aBad - bBad || b.value - a.value;
    });
    const samePack = ranked.filter(item => item.relation !== "conflict" && item.relation !== "pieces");
    const pool = samePack.length ? samePack : ranked;
    const best = pool[0];
    if (!best) { push({key, status: "none", code: "", score: 0, reason: "", options: []}); return; }
    const close = pool.filter((item, index) => index > 0 && item.value >= best.value - 0.12).length;
    const options = pool.slice(0, 5).map(item => ({code: catalog[item.indexItem].code, score: Math.round(item.value * 100)}));
    const chosen = catalog[best.indexItem];
    const percent = Math.round(best.value * 100);
    const why = explainMatch(row[1], chosen.name, best.relation);
    if (best.relation === "conflict" || best.relation === "pieces") {
      push({key, status: "conflict", code: chosen.code, score: percent, reason: why, options});
      return;
    }
    if (best.value >= 0.95 && offerWords.length >= 2 && !close) {
      push({key, status: "sure", code: chosen.code, score: percent, reason: why || "название и фасовка совпали", options});
      return;
    }
    const several = close ? "Найдено несколько возможных вариантов. " : "";
    push({key, status: "review", code: chosen.code, score: percent, reason: several + (why || "похоже, нужна проверка"), options});
  });
  analyzed = true;
}
function attentionFor(supplier) {
  let count = 0;
  judged.forEach(item => {
    const row = latestByKey.get(item.key);
    if (row && row[0] === supplier && needsPerson(item, row)) count += 1;
  });
  return count;
}
function freshFor(supplier) {
  let count = 0;
  latestByKey.forEach(row => { if (row[0] === supplier && isFresh(row)) count += 1; });
  return count;
}
function refreshInbox() {
  if (!analyzed) return;
  state.inbox.forEach(item => {
    item.attention = attentionFor(item.supplier);
    item.fresh = freshFor(item.supplier);
  });
}
function reportText() {
  if (!state.catalog.length) return "Сначала загрузите номенклатуру.";
  const counts = tally(judged);
  return "Автосопоставлено: " + (counts.sure || 0)
    + ". На проверку: " + ((counts.review || 0) + (counts.shift || 0))
    + ". Конфликт: " + (counts.conflict || 0)
    + ". Нет соответствия: " + (counts.none || 0)
    + ". Подтверждено: " + (counts.confirmed || 0) + ".";
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
function renderCatalogCard() {
  const status = document.getElementById("catalog-status");
  const wrap = document.getElementById("our-file-wrap");
  const replace = document.getElementById("replace-catalog");
  if (!state.catalog.length) {
    catalogReplaceOpen = false;
    status.innerHTML = "<p class='warn'>Список пуст. Выберите файл номенклатуры из 1С: код, наименование, единица.</p>";
    wrap.classList.remove("hidden");
    replace.classList.add("hidden");
    return;
  }
  const sample = state.catalog.slice(0, 3).map(item => esc(item.name)).join(" · ");
  status.innerHTML = "<p class='ok'>Номенклатура на месте: <b>" + state.catalog.length + "</b> товаров. Повторно файл выбирать не нужно.</p>"
    + (sample ? "<p class='muted'>Например: " + sample + "</p>" : "");
  wrap.classList.toggle("hidden", !catalogReplaceOpen);
  replace.classList.remove("hidden");
}
function renderToday() {
  renderCatalogCard();
  const box = document.getElementById("upload-box");
  if (box) box.classList.toggle("hidden", !uploadOpen);
  const feed = document.getElementById("today-feed");
  const todayItems = state.inbox.filter(item => item.at === today());
  const waiting = todayItems.reduce((sum, item) => sum + Math.max(item.attention || 0, item.fresh || 0), 0);
  const lines = todayItems.map(item => {
    let note = "обработан";
    let mark = "✓";
    if (item.fresh && item.attention) {
      note = item.fresh + " новых товаров, " + item.attention + " требуют решения";
      mark = "⚠";
    } else if (item.fresh) {
      note = item.fresh + " новых товаров";
      mark = "⚠";
    } else if (item.attention) {
      note = item.attention + " позиций требуют сопоставления";
      mark = "⚠";
    }
    return "<p><button class='linkish open-inbox' type='button' data-supplier=\"" + esc(item.supplier) + "\">" + esc(item.supplier) + "</button> — " + note + " " + mark + " <span class='muted'>" + esc(item.time || "") + "</span></p>";
  }).join("");
  let summary = "<p>Сегодня прайсы ещё не загружали. Когда файл придёт, положите его кнопкой ниже.</p>";
  if (todayItems.length && !waiting) summary = "<p>Сегодня получено прайсов: <b>" + todayItems.length + "</b>.</p>" + lines + "<p class='ok'>Все полученные прайсы обработаны. Действий не требуется.</p>";
  if (todayItems.length && waiting) summary = "<p>Сегодня получено прайсов: <b>" + todayItems.length + "</b>.</p>" + lines + "<p><b>Требуют вашего решения: " + waiting + " позиций</b></p><p><button class='action' id='go-match' type='button'>Проверить " + waiting + " позиций</button></p>";
  feed.innerHTML = summary;
  const openMatch = document.getElementById("go-match");
  if (openMatch) openMatch.addEventListener("click", () => { matchFilter = "attention"; show("match"); });
  feed.querySelectorAll(".open-inbox").forEach(button => button.addEventListener("click", () => {
    state.supplierPick = button.dataset.supplier;
    matchFilter = "attention";
    save();
    show("match");
  }));
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
    const seen = new Set();
    const next = [];
    for (let index = header + 1; index < grid.length; index += 1) {
      const row = grid[index];
      const name = cell(row[map.name]);
      if (!name) continue;
      const code = map.code >= 0 ? cell(row[map.code]) || name : name;
      const mark = norm(name) + "\n" + (code.trim() === name.trim() ? "" : code.trim());
      if (seen.has(mark)) continue;
      seen.add(mark);
      next.push({
        code,
        barcode: map.barcode >= 0 ? barcode(row[map.barcode]) : "",
        name, unit: map.unit >= 0 ? cell(row[map.unit]) : ""
      });
    }
    state.catalog = next;
    analyzed = false;
    analyze();
    save();
    catalogReplaceOpen = false;
    const examples = state.catalog.slice(0, 3).map(item => item.name).join(" · ");
    lastRead = examples ? "Пример: " + examples : "";
    document.getElementById("report").textContent = "Номенклатура загружена, " + state.catalog.length + " товаров. Сопоставления сохранены. " + lastRead;
    renderToday();
    return;
  }
  let added = 0;
  const batch = [];
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (rows[index][8] === file.name) rows.splice(index, 1);
  }
  for (let index = header + 1; index < grid.length; index += 1) {
    const row = grid[index];
    const name = cell(row[map.name]);
    const price = map.price >= 0 ? Number(String(row[map.price] || "").replace(/\s/g, "").replace(",", ".").replace(/[^\d.\-]/g, "")) : 0;
    if (!name || !price) continue;
    batch.push([typed, name, map.code >= 0 ? cell(row[map.code]) : "", map.barcode >= 0 ? barcode(row[map.barcode]) : "", map.unit >= 0 ? cell(row[map.unit]) : "", price, today(), "Цена", file.name]);
    added += 1;
  }
  let freshCount = 0;
  batch.forEach(row => {
    const id = row[0] + "\n" + stableId(row);
    if (!state.seen[id]) { state.fresh[id] = today(); freshCount += 1; }
    state.seen[id] = 1;
    rows.push(row);
  });
  state.uploads = state.uploads.filter(item => item.file !== file.name);
  state.uploads.push({file: file.name, supplier: typed, at: today(), rows: batch});
  state.inbox = state.inbox.filter(item => item.file !== file.name);
  state.inbox.push({supplier: typed, file: file.name, at: today(), time: clock(), count: added, fresh: freshCount, attention: 0});
  analyzed = false;
  if (state.catalog.length) analyze();
  else { rebuildLatest(); judged = []; analyzed = true; }
  const note = state.inbox[state.inbox.length - 1];
  if (note && analyzed) note.attention = attentionFor(typed);
  save();
  const examples = batch.slice(0, 3).map(row => row[1]).join(" · ");
  lastRead = (freshCount ? "В прайсе " + typed + " обнаружено " + freshCount + " новых позиций. " : "") + (examples ? "Пример: " + examples + ". " : "");
  document.getElementById("report").textContent = "Прайс загружен: " + added + " строк. " + lastRead + (state.catalog.length ? reportText() : "Номенклатура ещё не загружена, связка будет после неё.");
  state.supplierPick = typed;
  matchFilter = freshCount ? "fresh" : "attention";
  uploadOpen = true;
  renderToday();
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
document.getElementById("replace-catalog").addEventListener("click", () => {
  catalogReplaceOpen = true;
  renderCatalogCard();
});
document.getElementById("show-upload").addEventListener("click", () => {
  uploadOpen = !uploadOpen;
  renderToday();
});
document.getElementById("our-file").addEventListener("change", event => {
  const file = event.target.files[0];
  event.target.value = "";
  if (file) startImport(file, "our");
});
document.getElementById("supplier-file").addEventListener("change", event => {
  const file = event.target.files[0];
  if (file) startImport(file, "supplier");
});
function whenLabel(value) {
  if (!value) return "дата не указана";
  if (value === today()) return "сегодня";
  const parts = String(value).split(".");
  if (parts.length !== 3) return value;
  const then = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  if (then.getFullYear() === yesterday.getFullYear() && then.getMonth() === yesterday.getMonth() && then.getDate() === yesterday.getDate()) return "вчера";
  return value;
}
function supplierRows(name) {
  const list = [];
  latestByKey.forEach(row => { if (row[0] === name) list.push(row); });
  return list;
}
function supplierState(name) {
  const items = judged.filter(item => {
    const row = latestByKey.get(item.key);
    return row && row[0] === name;
  });
  const fresh = items.filter(item => isFresh(latestByKey.get(item.key)) && item.status !== "confirmed" && item.status !== "missing").length;
  const review = items.filter(item => item.status === "review" || item.status === "shift" || item.status === "conflict").length;
  if (fresh) return "⚠ " + fresh + " новых";
  if (review) return "⚠ " + review + " на проверку";
  if (!items.length) return "нет строк";
  return "✓ Обработан";
}
function renderPrices() {
  if (!analyzed) analyze();
  refreshInbox();
  const latestDate = new Map();
  const latestStamp = new Map();
  const counts = new Map();
  latestByKey.forEach(row => {
    counts.set(row[0], (counts.get(row[0]) || 0) + 1);
    const stamp = rowStamp(row, 0);
    if (!latestStamp.has(row[0]) || stamp >= latestStamp.get(row[0])) {
      latestStamp.set(row[0], stamp);
      latestDate.set(row[0], row[6]);
    }
  });
  const names = [...counts.keys()].sort((a, b) => a.localeCompare(b, "ru"));
  document.getElementById("supplier-summary").innerHTML = "<table><tr><th>Поставщик</th><th>Последний прайс</th><th>Позиций</th><th>Состояние</th></tr>"
    + names.map(name => "<tr><td><button class='linkish open-supplier' type='button' data-supplier=\"" + esc(name) + "\">" + esc(name) + "</button></td><td>"
      + esc(whenLabel(latestDate.get(name) || "")) + "</td><td>" + counts.get(name) + "</td><td>" + supplierState(name) + "</td></tr>").join("")
    + "</table>";
  document.querySelectorAll(".open-supplier").forEach(button => button.addEventListener("click", () => {
    priceOpen = button.dataset.supplier;
    priceShown = 40;
    priceFilter = "all";
    renderPriceDetail();
  }));
  if (priceOpen) renderPriceDetail();
}
function renderPriceDetail() {
  const box = document.getElementById("price-detail");
  if (!priceOpen) { box.classList.add("hidden"); return; }
  box.classList.remove("hidden");
  document.getElementById("price-title").textContent = priceOpen;
  const filters = [["all", "Все"], ["linked", "Сопоставленные"], ["review", "На проверку"], ["conflict", "Конфликт"], ["fresh", "Новинки"], ["none", "Нет соответствия"]];
  document.getElementById("price-filters").innerHTML = filters.map(([id, label]) => "<button class='filter" + (priceFilter === id ? " active" : "") + "' type='button' data-price-filter='" + id + "'>" + label + "</button>").join("");
  const query = (document.getElementById("price-search").value || "").trim().toLowerCase();
  const visible = judged.filter(item => {
    const row = latestByKey.get(item.key);
    if (!row || row[0] !== priceOpen) return false;
    if (priceFilter === "linked" && item.status !== "sure" && item.status !== "confirmed") return false;
    if (priceFilter === "review" && item.status !== "review" && item.status !== "shift") return false;
    if (priceFilter === "conflict" && item.status !== "conflict") return false;
    if (priceFilter === "fresh" && !isFresh(row)) return false;
    if (priceFilter === "none" && item.status !== "none" && item.status !== "missing") return false;
    if (query && !row[1].toLowerCase().includes(query)) return false;
    return true;
  });
  const slice = visible.slice(0, priceShown);
  document.getElementById("price-body").innerHTML = slice.map(item => {
    const row = latestByKey.get(item.key);
    const ours = item.code ? catalogByCode(item.code) : null;
    return "<tr><td>" + esc(row[1]) + "</td><td>" + esc(row[5]) + "</td><td>" + esc(ours ? ours.name : "—") + "</td><td>" + esc(STATUS_TEXT[item.status] || item.status) + "</td></tr>";
  }).join("") || "<tr><td colspan='4'>В этом фильтре пусто.</td></tr>";
  const more = document.getElementById("price-more");
  more.classList.toggle("hidden", visible.length <= slice.length);
  more.textContent = "Показать ещё (" + Math.max(visible.length - slice.length, 0) + ")";
}
document.getElementById("price-filters").addEventListener("click", event => {
  const button = event.target.closest("button");
  if (!button || !button.dataset.priceFilter) return;
  priceFilter = button.dataset.priceFilter;
  priceShown = 40;
  renderPriceDetail();
});
document.getElementById("price-search").addEventListener("input", () => { priceShown = 40; renderPriceDetail(); });
document.getElementById("price-more").addEventListener("click", () => { priceShown += 40; renderPriceDetail(); });
function esc(value) {
  return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}
const STATUS_TEXT = {
  sure: "Автосопоставлено",
  review: "На проверку",
  conflict: "Конфликт",
  shift: "На проверку",
  none: "Нет соответствия",
  missing: "Нет у нас",
  confirmed: "Подтверждено",
  doubt: "На проверку",
  picked: "На проверку"
};
function passesFilter(item, row) {
  if (matchFilter === "all") return true;
  if (matchFilter === "attention") return needsPerson(item, row);
  if (matchFilter === "sure") return item.status === "sure";
  if (matchFilter === "review") return item.status === "review" || item.status === "shift" || item.status === "doubt";
  if (matchFilter === "fresh") return isFresh(row);
  if (matchFilter === "none") return item.status === "none" || item.status === "missing";
  return item.status === matchFilter;
}
function renderMatch() {
  if (!analyzed) analyze();
  refreshInbox();
  const suppliers = [...new Set([...latestByKey.values()].map(row => row[0]))].sort((a, b) => a.localeCompare(b, "ru"));
  const select = document.getElementById("match-supplier");
  if (!state.supplierPick || !suppliers.includes(state.supplierPick)) state.supplierPick = suppliers[0] || "";
  select.innerHTML = suppliers.map(name => "<option value=\"" + esc(name) + "\"" + (name === state.supplierPick ? " selected" : "") + ">" + esc(name) + "</option>").join("");
  const filters = [
    ["attention", "Нужно решить"],
    ["all", "Все"],
    ["sure", "Автосопоставлено"],
    ["review", "На проверку"],
    ["conflict", "Конфликт"],
    ["fresh", "Новинки"],
    ["none", "Нет соответствия"],
    ["confirmed", "Подтверждено"]
  ];
  document.getElementById("match-filters").innerHTML = filters.map(([id, label]) => "<button class='filter" + (matchFilter === id ? " active" : "") + "' type='button' data-filter='" + id + "'>" + label + "</button>").join("");
  const query = (document.getElementById("match-search").value || "").trim().toLowerCase();
  const visible = judged.filter(item => {
    const row = latestByKey.get(item.key);
    if (!row || row[0] !== state.supplierPick) return false;
    if (!passesFilter(item, row)) return false;
    if (query && !row[1].toLowerCase().includes(query)) return false;
    return true;
  });
  document.getElementById("match-counts").textContent = state.catalog.length
    ? (lastRead ? lastRead + " " : "") + "В этом списке: " + visible.length + "."
      + (catalogLooksLikeIds() ? " Названия не загрузились: в справочнике коды, а не товары. Обновите номенклатуру из 1С на вкладке «Сегодня»." : "")
    : "Сначала загрузите номенклатуру на вкладке «Сегодня».";
  const slice = visible.slice(0, matchShown);
  document.getElementById("match-body").innerHTML = slice.map(item => {
    const row = latestByKey.get(item.key);
    const ours = item.code ? catalogByCode(item.code) : null;
    const oursName = ours && !catalogLooksLikeIds() ? ours.name : (item.code ? "товар не найден в текущем списке" : "—");
    const key = encodeURIComponent(item.key);
    const code = encodeURIComponent(item.code || "");
    const confirm = item.code && item.status !== "confirmed" && item.status !== "missing"
      ? "<button class='action yes' type='button' data-key='" + key + "' data-code='" + code + "'>Подтвердить</button>"
      : "";
    const alts = (item.options || []).slice(1, 5).map(option => {
      const alt = catalogByCode(option.code);
      return alt ? "<button class='quiet alt' type='button' data-key='" + key + "' data-code='" + encodeURIComponent(option.code) + "'>" + esc(alt.name) + " · " + option.score + "%</button>" : "";
    }).join("");
    const other = "<button class='secondary pick' type='button' data-key='" + key + "'>Выбрать другой товар</button>";
    const gone = item.status === "missing" ? "" : "<button class='quiet gone' type='button' data-key='" + key + "'>Такого товара у нас нет</button>";
    const rowClass = item.status === "conflict" ? "conflict" : (item.status === "review" || item.status === "shift" ? "review" : (item.status === "sure" || item.status === "confirmed" ? "done" : "need"));
    const score = item.status === "none" || item.status === "missing" ? "—" : ((item.score || 0) + "%");
    return "<tr class='" + rowClass + "'><td>" + esc(row[1]) + "</td><td>" + esc(row[0]) + "</td><td>" + esc(oursName) + (alts ? "<div>" + alts + "</div>" : "") + "</td><td>" + score + "</td><td>"
      + esc(item.reason || "") + "</td><td>" + esc(STATUS_TEXT[item.status] || item.status) + "</td><td>" + confirm + other + gone + "</td></tr>";
  }).join("") || "<tr><td colspan='7'>В этом списке пусто.</td></tr>";
  const more = document.getElementById("match-more");
  more.classList.toggle("hidden", visible.length <= slice.length);
  more.textContent = "Показать ещё (" + Math.max(visible.length - slice.length, 0) + ")";
}
function acceptMatch(key, code) {
  const row = latestByKey.get(key);
  if (!row || !code) return;
  putBond(row, code);
  const item = judgedByKey.get(key);
  if (item) { item.status = "confirmed"; item.code = code; item.score = 100; item.reason = "сопоставление уже подтверждено"; }
  refreshInbox();
  save();
  renderMatch();
}
function markMissing(key) {
  const row = latestByKey.get(key);
  if (!row) return;
  const id = row[0] + "\n" + stableId(row);
  state.missing[id] = {at: today()};
  delete state.fresh[id];
  delete state.bonds[id];
  delete state.verified[key];
  delete state.links[key];
  delete state.methods[key];
  delete state.draft[key];
  const item = judgedByKey.get(key);
  if (item) { item.status = "missing"; item.code = ""; item.score = 0; item.reason = "отметили, что такого товара у нас нет"; }
  refreshInbox();
  save();
  renderMatch();
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
  const yes = event.target.closest("button.yes, button.alt");
  if (yes) { acceptMatch(decodeURIComponent(yes.dataset.key), decodeURIComponent(yes.dataset.code || "")); return; }
  const gone = event.target.closest("button.gone");
  if (gone) { markMissing(decodeURIComponent(gone.dataset.key)); return; }
  const pick = event.target.closest("button.pick");
  if (!pick) return;
  activeKey = decodeURIComponent(pick.dataset.key);
  openPicker(pick);
});
let activeKey = "";
function hideDrop() {
  document.getElementById("inline-drop").classList.add("hidden");
}
function placeDrop(anchor) {
  const drop = document.getElementById("inline-drop");
  const rect = anchor.getBoundingClientRect();
  drop.style.top = (rect.bottom + 4) + "px";
  drop.style.left = Math.max(rect.left - 160, 16) + "px";
  drop.style.width = "420px";
  return drop;
}
function fillPicker(query) {
  const list = document.getElementById("picker-list");
  if (!list) return;
  if (catalogLooksLikeIds()) {
    list.innerHTML = "<p class='warn'>В справочнике коды, а не названия. Обновите номенклатуру из 1С на вкладке «Сегодня».</p>";
    return;
  }
  if (!state.catalog.length) {
    list.innerHTML = "<p class='warn'>Справочник пуст. Загрузите номенклатуру на вкладке «Сегодня».</p>";
    return;
  }
  const parts = sigTokens(query);
  const wanted = measures(query);
  if ((!parts.length || parts.join("").length < 3) && !wanted.length) {
    list.innerHTML = "<p class='muted'>Введите часть названия, например ALLEY или шуманит.</p>";
    return;
  }
  const found = uniqueCatalog().filter(item => {
    const ours = sigTokens(item.name);
    const wordsOk = !parts.length || parts.every(part => ours.some(token => tokenFits(part, token)));
    const has = measures(item.name);
    const weightOk = !wanted.length || !has.length || wanted.some(piece => has.includes(piece));
    return wordsOk && weightOk;
  });
  const slice = found.slice(0, 30);
  list.innerHTML = found.length
    ? "<p class='muted'>Найдено " + found.length + (found.length > slice.length ? ", показаны первые " + slice.length : "") + ".</p>"
      + slice.map(item => "<button class='quiet choose' type='button' data-code='" + encodeURIComponent(item.code) + "'>" + choiceLabel(item) + "</button>").join("")
    : "<p class='muted'>В справочнике нет «" + esc(query) + "».</p>";
}
function openPicker(anchor) {
  const drop = placeDrop(anchor);
  drop.innerHTML = "<p><input id='picker-query' type='text' placeholder='шуманит'></p><div id='picker-list'><p class='muted'>Введите часть названия, например шуманит.</p></div>";
  drop.classList.remove("hidden");
  const field = document.getElementById("picker-query");
  field.focus();
  field.addEventListener("input", () => fillPicker(field.value));
}
document.getElementById("inline-drop").addEventListener("mousedown", event => {
  const button = event.target.closest("button.choose");
  if (!button || !activeKey) return;
  event.preventDefault();
  const code = decodeURIComponent(button.dataset.code);
  hideDrop();
  acceptMatch(activeKey, code);
});
document.addEventListener("mousedown", event => {
  if (event.target.closest("#inline-drop") || event.target.closest("button.pick")) return;
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
  if (!analyzed && state.catalog.length) analyze();
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
(state.uploads || []).forEach(pack => {
  if (!pack || !pack.rows || rows.some(row => row[8] === pack.file)) return;
  pack.rows.forEach(row => rows.push(row));
});
if (!state.seenReady) {
  rows.forEach(row => { state.seen[row[0] + "\n" + stableId(row)] = 1; });
  state.seenReady = true;
}
save();
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
