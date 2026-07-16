'use strict';

const APP_VERSION = '1.1.0';
const KEY = 'jh-assist-v09';
const CLOUD_SCHEMA = 'jh-assist-cloud-v1';
const CLOUD_FILE_DEFAULT = 'jh-assist-data.json';
const SYNC_CONFIG_KEY = 'jh-assist-sync-config-v1';
const DEVICE_ID_KEY = 'jh-assist-device-id-v1';
const MIGRATION_AT_KEY = 'jh-assist-migration-at-v1';
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPES = ['Files.ReadWrite.AppFolder'];

const $ = id => document.getElementById(id);
const esc = (s = '') => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uuid = () => crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random();
const pad2 = n => String(n).padStart(2, '0');
const nowIso = () => new Date().toISOString();
const validIso = value => typeof value === 'string' && !Number.isNaN(Date.parse(value));
const timestampOf = (item, fallback = '') => validIso(item?.updatedAt) ? item.updatedAt : validIso(item?.createdAt) ? item.createdAt : validIso(item?.timestamp) ? item.timestamp : fallback;

function localIsoDate(date = new Date()) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}
function monthLastDate(year, month) {
  return `${year}-${pad2(month)}-${pad2(new Date(year, month, 0).getDate())}`;
}
function duration(start, end) {
  if (!start || !end) return 0;
  const [sh, sm] = start.split(':').map(Number);
  const [eh, em] = end.split(':').map(Number);
  return Math.max(0, (eh * 60 + em - sh * 60 - sm) / 60);
}
function formatHours(n) {
  return Number(n || 0).toLocaleString('de-DE', {minimumFractionDigits: 0, maximumFractionDigits: 2}) + ' Std.';
}
function weeksBetween(a, b) {
  if (!a || !b) return 0;
  const d1 = new Date(a + 'T12:00:00');
  const d2 = new Date(b + 'T12:00:00');
  return Math.max(0, (d2 - d1) / (7 * 86400000) + 1 / 7);
}
function monthKey(dateOrYear, month) {
  return month === undefined ? String(dateOrYear).slice(0, 7) : `${dateOrYear}-${String(month).padStart(2, '0')}`;
}
function isAppleMobile() {
  return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}
function isWindows() {
  return /Windows/i.test(navigator.userAgent);
}

function defaultState() {
  return {
    cases: [],
    services: [],
    prices: [],
    auditLog: [],
    billedMonths: {},
    deletedServiceIds: [],
    tombstones: {services: {}},
    meta: {schema: CLOUD_SCHEMA, version: APP_VERSION, updatedAt: nowIso(), updatedBy: ''}
  };
}

const old = JSON.parse(
  localStorage.getItem('jh-assist-v08') ||
  localStorage.getItem('jh-assist-v07') ||
  localStorage.getItem('jh-assist-v06') ||
  localStorage.getItem('jh-assist-v03') ||
  localStorage.getItem('jh-assist-v02') ||
  localStorage.getItem('jh-assist-v01') ||
  'null'
);
const migrationAt = localStorage.getItem(MIGRATION_AT_KEY) || nowIso();
localStorage.setItem(MIGRATION_AT_KEY, migrationAt);

function normalizeState(input, fallbackTimestamp = migrationAt) {
  const source = input && typeof input === 'object' ? input : defaultState();
  const result = defaultState();
  const fallback = validIso(source?.meta?.updatedAt) ? source.meta.updatedAt : fallbackTimestamp;

  result.cases = (Array.isArray(source.cases) ? source.cases : []).filter(Boolean).map(item => {
    const id = item.id || uuid();
    const createdAt = validIso(item.createdAt) ? item.createdAt : validIso(item.updatedAt) ? item.updatedAt : fallback;
    const updatedAt = validIso(item.updatedAt) ? item.updatedAt : createdAt;
    return {...item, id, createdAt, updatedAt};
  });
  result.services = (Array.isArray(source.services) ? source.services : []).filter(Boolean).map(item => {
    const id = item.id || uuid();
    const createdAt = validIso(item.createdAt) ? item.createdAt : validIso(item.updatedAt) ? item.updatedAt : fallback;
    const updatedAt = validIso(item.updatedAt) ? item.updatedAt : createdAt;
    return {...item, id, createdAt, updatedAt};
  });
  result.prices = (Array.isArray(source.prices) ? source.prices : []).filter(Boolean).map(item => {
    const year = Number(item.year);
    const id = item.id || `price-${year}`;
    const createdAt = validIso(item.createdAt) ? item.createdAt : validIso(item.updatedAt) ? item.updatedAt : fallback;
    const updatedAt = validIso(item.updatedAt) ? item.updatedAt : createdAt;
    return {...item, id, year, city: Number(item.city || 0), county: Number(item.county || 0), createdAt, updatedAt};
  });
  result.auditLog = (Array.isArray(source.auditLog) ? source.auditLog : []).filter(Boolean).map(item => ({
    ...item,
    id: item.id || uuid(),
    timestamp: validIso(item.timestamp) ? item.timestamp : fallback
  })).sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp))).slice(0, 1000);

  result.billedMonths = {};
  if (source.billedMonths && typeof source.billedMonths === 'object') {
    for (const [key, value] of Object.entries(source.billedMonths)) {
      if (!value || typeof value !== 'object') continue;
      result.billedMonths[key] = {
        ...value,
        updatedAt: validIso(value.updatedAt) ? value.updatedAt : validIso(value.exportedAt) ? value.exportedAt : fallback
      };
    }
  }

  result.tombstones = {services: {}};
  const sourceTombstones = source.tombstones?.services && typeof source.tombstones.services === 'object' ? source.tombstones.services : {};
  for (const [id, deletedAt] of Object.entries(sourceTombstones)) {
    result.tombstones.services[id] = validIso(deletedAt) ? deletedAt : fallback;
  }
  for (const id of Array.isArray(source.deletedServiceIds) ? source.deletedServiceIds : []) {
    if (!result.tombstones.services[id]) result.tombstones.services[id] = fallback;
  }
  result.deletedServiceIds = Object.keys(result.tombstones.services);

  result.meta = {
    schema: CLOUD_SCHEMA,
    version: APP_VERSION,
    updatedAt: validIso(source?.meta?.updatedAt) ? source.meta.updatedAt : fallback,
    updatedBy: String(source?.meta?.updatedBy || '')
  };
  return result;
}

let state = normalizeState(
  JSON.parse(localStorage.getItem(KEY) || 'null') || {
    cases: old?.cases || [],
    services: old?.services || [],
    prices: old?.prices || [],
    auditLog: old?.auditLog || [],
    billedMonths: old?.billedMonths || {},
    deletedServiceIds: old?.deletedServiceIds || []
  }
);

const deviceId = localStorage.getItem(DEVICE_ID_KEY) || uuid();
localStorage.setItem(DEVICE_ID_KEY, deviceId);

function persistState({touch = false, scheduleSync = false, render = true} = {}) {
  state = normalizeState(state);
  if (touch) {
    state.meta.updatedAt = nowIso();
    state.meta.updatedBy = deviceId;
  }
  localStorage.setItem(KEY, JSON.stringify(state));
  if (render) renderAll();
  if (scheduleSync) queueCloudSync();
}
function save() {
  persistState({touch: true, scheduleSync: true, render: true});
}
function isBilledMonth(key) {
  return Boolean(state.billedMonths[key]);
}
function setChangedAfterBilling(key) {
  if (state.billedMonths[key]) {
    state.billedMonths[key].status = 'changed_after_billing';
    state.billedMonths[key].updatedAt = nowIso();
  }
}
function logChange({caseId = '', month = '', action, field = '', oldValue = '', newValue = ''}) {
  state.auditLog.unshift({
    id: uuid(), timestamp: nowIso(), caseId, month, action, field,
    oldValue: String(oldValue ?? ''), newValue: String(newValue ?? '')
  });
  if (state.auditLog.length > 1000) state.auditLog.length = 1000;
  if (month) setChangedAfterBilling(month);
}
function allCases() {
  return [...state.cases].sort((a, b) => a.name.localeCompare(b.name, 'de'));
}
function activeCases() {
  return allCases().filter(c => c.active !== false);
}
function caseById(id) {
  return state.cases.find(c => c.id === id);
}

function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
  $('view-' + name).classList.remove('hidden');
  document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.view === name));
}
document.querySelectorAll('.tab').forEach(b => b.onclick = () => showView(b.dataset.view));

function renderSelects() {
  const active = activeCases();
  const all = allCases();
  const build = list => list.length
    ? list.map(c => `<option value="${c.id}">${esc(c.name)}${c.active === false ? ' (archiviert)' : ''}</option>`).join('')
    : '<option value="">Zuerst Fall anlegen</option>';
  const lists = {serviceCase: active, hoursCase: all, reportCase: all};
  for (const [id, list] of Object.entries(lists)) {
    const el = $(id);
    if (!el) continue;
    const current = el.value;
    el.innerHTML = build(list);
    if ([...el.options].some(o => o.value === current)) el.value = current;
  }
}
function renderCases() {
  const q = $('caseSearch').value.toLowerCase();
  const filter = $('caseFilter').value;
  const list = state.cases.filter(c =>
    (filter === 'all' || (filter === 'active' && c.active !== false) || (filter === 'inactive' && c.active === false)) &&
    [c.name, c.authority, c.officer, c.guardians].join(' ').toLowerCase().includes(q)
  ).sort((a, b) => a.name.localeCompare(b.name, 'de'));
  $('cases').innerHTML = list.length ? list.map(c => `
    <div class="item"><strong>${esc(c.name)}</strong>${c.active === false ? '<span class="badge">archiviert</span>' : ''}
    <div class="muted">${esc(c.authority)} · ${formatHours(c.weeklyHours)} pro Woche · ${esc(c.officer || 'keine Sachbearbeitung')}</div>
    <div class="muted">Bewilligung: ${esc(c.approvalFrom || 'offen')} bis ${esc(c.approvalTo || 'offen')}</div>
    <div class="row" style="margin-top:8px"><button class="secondary small" onclick="editCase('${c.id}')">Bearbeiten</button><button class="${c.active === false ? 'secondary' : 'danger'} small" onclick="toggleCase('${c.id}')">${c.active === false ? 'Reaktivieren' : 'Archivieren'}</button></div></div>`).join('') : '<p class="muted">Keine passenden Fälle.</p>';
}
function renderServices() {
  const sorted = [...state.services].sort((a, b) => (b.date + b.start).localeCompare(a.date + a.start)).slice(0, 15);
  $('services').innerHTML = sorted.length ? sorted.map(s => {
    const c = caseById(s.caseId);
    const changed = isBilledMonth(monthKey(s.date)) && state.billedMonths[monthKey(s.date)].status === 'changed_after_billing';
    return `<div class="item"><strong>${esc(c?.name || 'Unbekannter Fall')}</strong><span class="badge">${formatHours(duration(s.start, s.end))}</span>${changed ? '<span class="badge warn">nach Abrechnung geändert</span>' : ''}<div>${esc(s.date)} · ${esc(s.start)}–${esc(s.end)}</div><div class="muted">${esc(s.activity)}${s.note ? ' · ' + esc(s.note) : ''}</div><div class="row" style="margin-top:7px"><button class="secondary small" onclick="editService('${s.id}')">Bearbeiten</button><button class="danger small" onclick="deleteService('${s.id}')">Löschen</button></div></div>`;
  }).join('') : '<p class="muted">Noch keine Leistungen erfasst.</p>';
}
function renderPrices() {
  $('prices').innerHTML = state.prices.length ? [...state.prices].sort((a, b) => b.year - a.year).map(p => `<div class="item"><strong>${p.year}</strong><div class="muted">Stadt: ${Number(p.city).toLocaleString('de-DE', {style: 'currency', currency: 'EUR'})} · Landkreis: ${Number(p.county).toLocaleString('de-DE', {style: 'currency', currency: 'EUR'})}</div></div>`).join('') : '<p class="muted">Noch keine Preise hinterlegt.</p>';
}
function renderHours() {
  const id = $('hoursCase').value;
  const c = caseById(id);
  const year = Number($('hoursYear').value);
  if (!c) {
    $('hoursSummary').innerHTML = '<p class="muted">Kein aktiver Fall vorhanden.</p>';
    $('notesList').innerHTML = '';
    return;
  }
  const sv = state.services.filter(s => s.caseId === id && Number(s.date.slice(0, 4)) === year);
  const done = sv.reduce((n, s) => n + duration(s.start, s.end), 0);
  let from = c.approvalFrom || year + '-01-01';
  let to = c.approvalTo || year + '-12-31';
  if (Number(from.slice(0, 4)) < year) from = year + '-01-01';
  if (Number(to.slice(0, 4)) > year) to = year + '-12-31';
  if (Number(from.slice(0, 4)) > year || Number(to.slice(0, 4)) < year) { from = ''; to = ''; }
  const approved = weeksBetween(from, to) * Number(c.weeklyHours || 0);
  const remaining = approved - done;
  const p = state.prices.find(x => Number(x.year) === year);
  const rate = p ? (c.authority === 'Stadt' ? p.city : p.county) : 0;
  $('hoursSummary').innerHTML = `<div class="stats" style="margin-top:16px"><div class="stat"><span class="muted">Jahreskontingent rechnerisch</span><strong>${formatHours(approved)}</strong></div><div class="stat"><span class="muted">Geleistet</span><strong>${formatHours(done)}</strong></div><div class="stat"><span class="muted">Rest</span><strong>${formatHours(remaining)}</strong></div></div><p class="muted">Grundlage: ${c.weeklyHours} Wochenstunden innerhalb des im Kalenderjahr liegenden Bewilligungszeitraums.${rate ? ' Aktueller Satz: ' + Number(rate).toLocaleString('de-DE', {style: 'currency', currency: 'EUR'}) + '.' : ''}</p>${remaining < 0 ? '<p class="message" style="color:#8b3030">Achtung: Das rechnerische Kontingent ist überschritten.</p>' : ''}`;
  const notes = sv.filter(s => s.note).sort((a, b) => b.date.localeCompare(a.date));
  $('notesList').innerHTML = notes.length ? notes.map(s => `<div class="item"><strong>${esc(s.date)}</strong><div>${esc(s.note)}</div><div class="muted">${esc(s.start)}–${esc(s.end)}</div></div>`).join('') : '<p class="muted">Für dieses Jahr sind keine Verlaufsnotizen gespeichert.</p>';
}
function renderBillingStatus() {
  const key = monthKey(Number($('billingYear').value), Number($('billingMonth').value));
  const entry = state.billedMonths[key];
  $('billingStatus').innerHTML = !entry ? 'Status: noch nicht abgerechnet' : entry.status === 'changed_after_billing'
    ? `<span class="badge warn">geändert nach Abrechnung</span> Letzter Export: ${esc(new Date(entry.exportedAt).toLocaleString('de-DE'))}`
    : `<span class="badge ok">abgerechnet</span> Export: ${esc(new Date(entry.exportedAt).toLocaleString('de-DE'))}`;
}
function renderAuditLog() {
  const items = state.auditLog.slice(0, 100);
  $('auditLog').innerHTML = items.length ? items.map(x => {
    const c = caseById(x.caseId);
    return `<div class="item"><strong>${esc(new Date(x.timestamp).toLocaleString('de-DE'))}</strong>${x.month ? ` <span class="badge">${esc(x.month)}</span>` : ''}<div>${esc(c?.name || 'Allgemein')} – ${esc(x.action)}${x.field ? ` – ${esc(x.field)}` : ''}</div>${x.oldValue || x.newValue ? `<div class="muted">${esc(x.oldValue)} → ${esc(x.newValue)}</div>` : ''}</div>`;
  }).join('') : '<p class="muted">Noch keine Änderungen nach einer Abrechnung protokolliert.</p>';
}
function renderAll() {
  renderSelects();
  renderCases();
  renderServices();
  renderPrices();
  renderHours();
  renderBillingStatus();
  renderAuditLog();
  renderSyncPanel();
}

$('caseForm').onsubmit = e => {
  e.preventDefault();
  const id = $('caseId').value;
  const existing = id ? caseById(id) : null;
  const name = $('caseName').value.trim();
  const sheetName = $('sheetName').value.trim() || name;
  if (/[\\/?*\[\]:]/.test(sheetName) || sheetName.length > 31) {
    alert('Der Tabellenblattname darf höchstens 31 Zeichen lang sein und keines der Zeichen \\ / ? * [ ] : enthalten.');
    return;
  }
  if (state.cases.some(c => c.id !== id && String(c.sheetName || c.name).trim().toLocaleLowerCase('de-DE') === sheetName.toLocaleLowerCase('de-DE'))) {
    alert('Dieser Tabellenblattname ist bereits einem anderen Fall zugeordnet.');
    return;
  }
  const timestamp = nowIso();
  const obj = {
    id: id || uuid(), name, authority: $('authority').value,
    weeklyHours: Number($('weeklyHours').value), officer: $('officer').value.trim(),
    guardians: $('guardians').value.trim(), approvalFrom: $('approvalFrom').value,
    approvalTo: $('approvalTo').value, sheetName,
    active: id ? existing?.active !== false : true,
    createdAt: existing?.createdAt || timestamp, updatedAt: timestamp
  };
  if (obj.approvalFrom && obj.approvalTo && obj.approvalTo < obj.approvalFrom) {
    alert('Das Bewilligungsende liegt vor dem Beginn.');
    return;
  }
  if (id) state.cases[state.cases.findIndex(c => c.id === id)] = obj;
  else state.cases.push(obj);
  resetCaseForm();
  save();
};
function resetCaseForm() {
  $('caseForm').reset();
  $('caseId').value = '';
  $('caseHeading').textContent = 'Fall anlegen';
  $('cancelCaseEdit').classList.add('hidden');
}
window.editCase = id => {
  const c = caseById(id);
  if (!c) return;
  showView('cases');
  $('caseId').value = c.id;
  $('caseName').value = c.name;
  $('authority').value = c.authority;
  $('weeklyHours').value = c.weeklyHours;
  $('officer').value = c.officer || '';
  $('guardians').value = c.guardians || '';
  $('approvalFrom').value = c.approvalFrom || '';
  $('approvalTo').value = c.approvalTo || '';
  $('sheetName').value = c.sheetName || c.name;
  $('caseHeading').textContent = 'Fall bearbeiten';
  $('cancelCaseEdit').classList.remove('hidden');
  scrollTo({top: 0, behavior: 'smooth'});
};
window.toggleCase = id => {
  const c = caseById(id);
  if (!c) return;
  c.active = c.active === false;
  c.updatedAt = nowIso();
  save();
};
$('cancelCaseEdit').onclick = resetCaseForm;
$('caseSearch').oninput = renderCases;
$('caseFilter').onchange = renderCases;

function minutesOf(time) {
  const [h, m] = time.split(':').map(Number);
  return h * 60 + m;
}
function overlaps(aStart, aEnd, bStart, bEnd) {
  return minutesOf(aStart) < minutesOf(bEnd) && minutesOf(bStart) < minutesOf(aEnd);
}
function isQuarterHour(time) {
  return /^(?:[01]\d|2[0-3]):(?:00|15|30|45)$/.test(time);
}
$('serviceForm').onsubmit = e => {
  e.preventDefault();
  const id = $('serviceId').value;
  const caseId = $('serviceCase').value;
  const date = $('date').value;
  const start = $('start').value;
  const end = $('end').value;
  const c = caseById(caseId);
  const existing = id ? state.services.find(s => s.id === id) : null;
  $('serviceMessage').textContent = '';
  if (!c) return;
  if (!isQuarterHour(start) || !isQuarterHour(end)) {
    $('serviceMessage').textContent = 'Beginn und Ende müssen im Viertelstundentakt liegen.';
    return;
  }
  if ((c.approvalFrom && date < c.approvalFrom) || (c.approvalTo && date > c.approvalTo)) {
    if (!confirm('Das Datum liegt außerhalb des hinterlegten Bewilligungszeitraums. Trotzdem speichern?')) return;
  }
  if (state.services.some(s => s.id !== id && s.caseId === caseId && s.date === date)) {
    $('serviceMessage').textContent = 'Für diesen Fall ist an diesem Tag bereits ein Eintrag vorhanden.';
    return;
  }
  if (minutesOf(end) <= minutesOf(start)) {
    $('serviceMessage').textContent = 'Die Endzeit muss nach der Beginnzeit liegen.';
    return;
  }
  if (duration(start, end) > 4 && !confirm(`Der Termin dauert ${formatHours(duration(start, end))}. Ist das korrekt?`)) return;
  const conflict = state.services.find(s => s.id !== id && s.date === date && overlaps(start, end, s.start, s.end));
  if (conflict) {
    const other = caseById(conflict.caseId);
    $('serviceMessage').textContent = `Terminüberschneidung mit ${other?.name || 'einem anderen Fall'} (${conflict.start}–${conflict.end}). Direkte Anschlusstermine sind erlaubt.`;
    return;
  }
  const timestamp = nowIso();
  const obj = {
    id: id || uuid(), caseId, date, start, end,
    activity: $('activity').value.trim(), note: $('note').value.trim(),
    useForReport: $('reportUse').checked,
    createdAt: existing?.createdAt || timestamp, updatedAt: timestamp
  };
  if (existing) {
    const tracked = {
      Fall: [existing.caseId, obj.caseId], Datum: [existing.date, obj.date],
      Beginn: [existing.start, obj.start], Ende: [existing.end, obj.end],
      Tätigkeit: [existing.activity, obj.activity], Verlaufsnotiz: [existing.note, obj.note],
      'Für Verlaufsdokumentation': [existing.useForReport !== false, obj.useForReport !== false]
    };
    const affectedMonths = new Set([monthKey(existing.date), monthKey(obj.date)]);
    for (const [field, [oldValue, newValue]] of Object.entries(tracked)) {
      if (String(oldValue ?? '') !== String(newValue ?? '')) {
        for (const key of affectedMonths) {
          if (isBilledMonth(key)) logChange({caseId: obj.caseId, month: key, action: 'Leistung geändert', field, oldValue, newValue});
        }
      }
    }
    state.services[state.services.findIndex(s => s.id === id)] = obj;
    $('serviceMessage').textContent = 'Leistung geändert.';
  } else {
    state.services.push(obj);
    delete state.tombstones.services[obj.id];
    const key = monthKey(date);
    if (isBilledMonth(key)) logChange({caseId, month: key, action: 'Leistung nach Abrechnung hinzugefügt', field: 'Termin', oldValue: '', newValue: `${date} ${start}–${end}`});
    $('serviceMessage').textContent = 'Leistung gespeichert.';
  }
  resetServiceForm();
  save();
};
function resetServiceForm() {
  $('serviceId').value = '';
  $('serviceHeading').textContent = 'Leistung erfassen';
  $('saveServiceBtn').textContent = 'Leistung speichern';
  $('cancelServiceEdit').classList.add('hidden');
  $('note').value = '';
  $('reportUse').checked = true;
}
window.editService = id => {
  const s = state.services.find(x => x.id === id);
  if (!s) return;
  showView('entry');
  $('serviceId').value = s.id;
  $('serviceCase').value = s.caseId;
  $('date').value = s.date;
  $('start').value = s.start;
  $('end').value = s.end;
  $('activity').value = s.activity || 'einzelfallbezogene Tätigkeit';
  $('note').value = s.note || '';
  $('reportUse').checked = s.useForReport !== false;
  $('serviceHeading').textContent = 'Leistung bearbeiten';
  $('saveServiceBtn').textContent = 'Änderung speichern';
  $('cancelServiceEdit').classList.remove('hidden');
  scrollTo({top: 0, behavior: 'smooth'});
};
$('cancelServiceEdit').onclick = resetServiceForm;
window.deleteService = id => {
  const s = state.services.find(x => x.id === id);
  if (!s || !confirm('Diesen Leistungseintrag wirklich löschen?')) return;
  const timestamp = nowIso();
  const key = monthKey(s.date);
  if (isBilledMonth(key)) logChange({caseId: s.caseId, month: key, action: 'Leistung nach Abrechnung gelöscht', field: 'Termin', oldValue: `${s.date} ${s.start}–${s.end}`, newValue: ''});
  state.services = state.services.filter(x => x.id !== id);
  state.tombstones.services[id] = timestamp;
  state.deletedServiceIds = Object.keys(state.tombstones.services);
  save();
};
$('priceForm').onsubmit = e => {
  e.preventDefault();
  const year = Number($('priceYear').value);
  const i = state.prices.findIndex(p => Number(p.year) === year);
  const existing = i >= 0 ? state.prices[i] : null;
  const timestamp = nowIso();
  const obj = {
    id: existing?.id || `price-${year}`, year,
    city: Number($('priceCity').value), county: Number($('priceCounty').value),
    createdAt: existing?.createdAt || timestamp, updatedAt: timestamp
  };
  if (i >= 0) state.prices[i] = obj;
  else state.prices.push(obj);
  save();
};
$('hoursCase').onchange = renderHours;
$('hoursYear').oninput = renderHours;

function downloadBlob(content, name, type = 'application/json') {
  const blob = new Blob([content], {type});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 3000);
}
async function exportJsonFile(payload, fileName) {
  downloadBlob(JSON.stringify(payload, null, 2), fileName);
}
$('exportBtn').onclick = async () => {
  await exportJsonFile({schema: 'jh-assist-backup-v1', version: APP_VERSION, exportedAt: nowIso(), ...state}, 'JH-Assist-Vollsicherung-' + localIsoDate() + '.json');
};
$('importFile').onchange = async e => {
  const f = e.target.files[0];
  if (!f) return;
  try {
    const d = JSON.parse(await f.text());
    if (!Array.isArray(d.cases) || !Array.isArray(d.services)) throw Error();
    if (confirm('Alle lokalen und synchronisierten Daten durch diese Vollsicherung ersetzen?')) {
      state = normalizeState(d, nowIso());
      save();
    }
  } catch {
    alert('Die Datei ist keine gültige JH-Assist-Vollsicherung.');
  }
  e.target.value = '';
};
$('clearBtn').onclick = () => {
  if (confirm('Wirklich nur die lokal gespeicherten Daten auf diesem Gerät löschen? Beim nächsten Synchronisieren werden die OneDrive-Daten wieder geladen.')) {
    state = defaultState();
    persistState({touch: true, scheduleSync: false, render: true});
  }
};

function reportServices() {
  const id = $('reportCase').value;
  const from = $('reportFrom').value;
  const to = $('reportTo').value;
  return state.services.filter(s => s.caseId === id && s.note && s.useForReport !== false && (!from || s.date >= from) && (!to || s.date <= to)).sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
}
function createReport() {
  const c = caseById($('reportCase').value);
  const items = reportServices();
  const msg = $('reportMessage');
  if (!c) { msg.textContent = 'Bitte einen Fall auswählen.'; return; }
  if (!items.length) {
    $('reportOutput').value = '';
    msg.textContent = 'Im gewählten Zeitraum liegen keine für den Bericht markierten Verlaufsnotizen vor.';
    return;
  }
  const from = $('reportFrom').value || items[0].date;
  const to = $('reportTo').value || items[items.length - 1].date;
  const lines = [];
  lines.push(`Verlaufsdokumentation – ${c.name}`);
  lines.push(`Berichtszeitraum: ${from} bis ${to}`);
  lines.push('');
  lines.push('Grundlage');
  lines.push(`Die folgende Verlaufsdokumentation beruht ausschließlich auf ${items.length} dokumentierten Verlaufsnotiz(en). Sie ist fachlich zu prüfen, zu gewichten und vor Verwendung zu überarbeiten.`);
  lines.push('');
  lines.push('Chronologischer Verlauf');
  for (const s of items) lines.push(`${s.date}, ${s.start}–${s.end}: ${s.note.trim()}`);
  lines.push('');
  lines.push('Zusammenfassende Verlaufsdarstellung');
  lines.push(`Im Berichtszeitraum fanden ${items.length} dokumentierte Kontakte im Rahmen der einzelfallbezogenen Tätigkeit statt. Die nachfolgenden Feststellungen geben die dokumentierten Beobachtungen und Arbeitsinhalte in chronologischer Form wieder:`);
  for (const s of items) lines.push(`- ${s.note.trim()}`);
  lines.push('');
  lines.push('Abschließende fachliche Einordnung');
  lines.push('Bitte ergänzen: Entwicklung im Berichtszeitraum, erreichte beziehungsweise nicht erreichte Hilfeplanziele, Ressourcen, bestehender Unterstützungsbedarf und fachliche Empfehlung für den weiteren Hilfeverlauf.');
  $('reportOutput').value = lines.join('\n');
  msg.textContent = 'Verlaufsdokumentation erstellt. Es wurden keine nicht dokumentierten Tatsachen ergänzt.';
}
$('createReportBtn').onclick = createReport;
$('copyReportBtn').onclick = async () => {
  if (!$('reportOutput').value) return;
  try {
    if (navigator.clipboard?.writeText) await navigator.clipboard.writeText($('reportOutput').value);
    else throw Error();
  } catch {
    $('reportOutput').focus();
    $('reportOutput').select();
    document.execCommand('copy');
  }
  $('reportMessage').textContent = 'Verlaufsdokumentation wurde kopiert.';
};
$('downloadReportBtn').onclick = () => {
  const c = caseById($('reportCase').value);
  if (!$('reportOutput').value) return;
  downloadBlob($('reportOutput').value, `Verlaufsdokumentation_${(c?.name || 'Fall').replace(/[^a-z0-9äöüß_-]+/gi, '_')}.txt`, 'text/plain;charset=utf-8');
};

/* Microsoft-Anmeldung und OneDrive-Synchronisierung */
let syncConfig = loadSyncConfig();
let msalApp = null;
let currentAccount = null;
let appFolderId = '';
let appFolderWebUrl = '';
let syncInFlight = null;
let syncTimer = null;
let syncStatus = {kind: 'idle', text: 'Nicht eingerichtet', detail: ''};

function loadSyncConfig() {
  const defaults = window.JH_ASSIST_SYNC_CONFIG && typeof window.JH_ASSIST_SYNC_CONFIG === 'object' ? window.JH_ASSIST_SYNC_CONFIG : {};
  let override = {};
  try { override = JSON.parse(localStorage.getItem(SYNC_CONFIG_KEY) || '{}'); } catch { override = {}; }
  return {
    clientId: String(override.clientId || defaults.clientId || '').trim(),
    tenantId: String(override.tenantId || defaults.tenantId || '').trim(),
    redirectUri: String(override.redirectUri || defaults.redirectUri || currentPageUrl()).trim(),
    fileName: String(defaults.fileName || CLOUD_FILE_DEFAULT).trim() || CLOUD_FILE_DEFAULT
  };
}
function currentPageUrl() {
  if (!location.protocol.startsWith('http')) return '';
  return `${location.origin}${location.pathname}`;
}
function isGuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}
function syncConfigured() {
  return isGuid(syncConfig.clientId) && isGuid(syncConfig.tenantId) && /^https:\/\//i.test(syncConfig.redirectUri);
}
function setSyncStatus(kind, text, detail = '') {
  syncStatus = {kind, text, detail};
  const badge = $('syncBadge');
  if (badge) {
    badge.textContent = text;
    badge.className = `sync-pill ${kind}`;
    badge.title = detail || text;
  }
  const status = $('syncStatusText');
  if (status) status.textContent = detail ? `${text} – ${detail}` : text;
}
function renderSyncPanel() {
  const client = $('syncClientId');
  const tenant = $('syncTenantId');
  const redirect = $('syncRedirectUri');
  if (client && document.activeElement !== client) client.value = syncConfig.clientId;
  if (tenant && document.activeElement !== tenant) tenant.value = syncConfig.tenantId;
  if (redirect) redirect.value = syncConfig.redirectUri || currentPageUrl();
  const accountText = $('syncAccount');
  if (accountText) accountText.textContent = currentAccount ? `Angemeldet als ${currentAccount.username || currentAccount.name || 'Microsoft-Konto'}` : 'Nicht bei Microsoft angemeldet.';
  const login = $('syncLoginBtn');
  const logout = $('syncLogoutBtn');
  const sync = $('syncNowBtn');
  if (login) login.classList.toggle('hidden', Boolean(currentAccount));
  if (logout) logout.classList.toggle('hidden', !currentAccount);
  if (sync) sync.disabled = !currentAccount || !navigator.onLine;
  const folder = $('cloudFolderLink');
  if (folder) {
    folder.classList.toggle('hidden', !appFolderWebUrl);
    if (appFolderWebUrl) folder.href = appFolderWebUrl;
  }
}
function stableObject(value) {
  if (Array.isArray(value)) return value.map(stableObject);
  if (value && typeof value === 'object') {
    const result = {};
    for (const key of Object.keys(value).sort()) result[key] = stableObject(value[key]);
    return result;
  }
  return value;
}
function comparableState(input) {
  const s = normalizeState(input);
  return stableObject({
    cases: [...s.cases].sort((a, b) => String(a.id).localeCompare(String(b.id))),
    services: [...s.services].sort((a, b) => String(a.id).localeCompare(String(b.id))),
    prices: [...s.prices].sort((a, b) => Number(a.year) - Number(b.year)),
    auditLog: [...s.auditLog].sort((a, b) => String(a.id).localeCompare(String(b.id))),
    billedMonths: s.billedMonths,
    tombstones: s.tombstones
  });
}
function fingerprint(input) {
  return JSON.stringify(comparableState(input));
}
function newerEntity(a, b, fallback = migrationAt) {
  if (!a) return b;
  if (!b) return a;
  const ta = Date.parse(timestampOf(a, fallback)) || 0;
  const tb = Date.parse(timestampOf(b, fallback)) || 0;
  if (ta !== tb) return ta > tb ? a : b;
  return JSON.stringify(stableObject(a)) >= JSON.stringify(stableObject(b)) ? a : b;
}
function mergeEntityArrays(localItems, remoteItems, keyFn) {
  const map = new Map();
  for (const item of localItems || []) map.set(keyFn(item), item);
  for (const item of remoteItems || []) {
    const key = keyFn(item);
    map.set(key, newerEntity(map.get(key), item));
  }
  return [...map.values()];
}
function mergeStates(localInput, remoteInput) {
  const local = normalizeState(localInput);
  const remote = normalizeState(remoteInput);
  const merged = defaultState();
  merged.tombstones.services = {...local.tombstones.services};
  for (const [id, deletedAt] of Object.entries(remote.tombstones.services)) {
    const current = merged.tombstones.services[id];
    if (!current || Date.parse(deletedAt) > Date.parse(current)) merged.tombstones.services[id] = deletedAt;
  }
  merged.cases = mergeEntityArrays(local.cases, remote.cases, item => item.id);
  merged.prices = mergeEntityArrays(local.prices, remote.prices, item => String(item.year));
  merged.services = mergeEntityArrays(local.services, remote.services, item => item.id).filter(item => {
    const deletedAt = merged.tombstones.services[item.id];
    return !deletedAt || Date.parse(timestampOf(item, migrationAt)) > Date.parse(deletedAt);
  });
  merged.deletedServiceIds = Object.keys(merged.tombstones.services);
  merged.auditLog = mergeEntityArrays(local.auditLog, remote.auditLog, item => item.id)
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp))).slice(0, 1000);
  merged.billedMonths = {};
  const keys = new Set([...Object.keys(local.billedMonths), ...Object.keys(remote.billedMonths)]);
  for (const key of keys) merged.billedMonths[key] = newerEntity(local.billedMonths[key], remote.billedMonths[key]);
  const localMetaTime = Date.parse(local.meta.updatedAt) || 0;
  const remoteMetaTime = Date.parse(remote.meta.updatedAt) || 0;
  merged.meta = localMetaTime >= remoteMetaTime ? {...local.meta} : {...remote.meta};
  merged.meta.schema = CLOUD_SCHEMA;
  merged.meta.version = APP_VERSION;
  return normalizeState(merged);
}
async function initializeMicrosoftAuth() {
  if (!syncConfigured()) {
    setSyncStatus('setup', 'Nicht eingerichtet', 'Client-ID und Mandanten-ID fehlen.');
    renderSyncPanel();
    return;
  }
  if (!location.protocol.startsWith('http')) {
    setSyncStatus('error', 'Online-App öffnen', 'Die Microsoft-Anmeldung funktioniert nur über die veröffentlichte HTTPS-App.');
    renderSyncPanel();
    return;
  }
  if (typeof msal === 'undefined') {
    setSyncStatus('error', 'Anmeldung nicht geladen', 'Die Microsoft-Anmeldebibliothek fehlt.');
    renderSyncPanel();
    return;
  }
  try {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: syncConfig.clientId,
        authority: `https://login.microsoftonline.com/${syncConfig.tenantId}`,
        redirectUri: syncConfig.redirectUri,
        postLogoutRedirectUri: syncConfig.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: {cacheLocation: 'localStorage'},
      system: {allowPlatformBroker: false}
    });
    await msalApp.initialize();
    const redirectResult = await msalApp.handleRedirectPromise();
    if (redirectResult?.account) msalApp.setActiveAccount(redirectResult.account);
    currentAccount = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
    if (currentAccount) msalApp.setActiveAccount(currentAccount);
    renderSyncPanel();
    if (currentAccount) {
      setSyncStatus(navigator.onLine ? 'syncing' : 'offline', navigator.onLine ? 'Synchronisiere…' : 'Offline', navigator.onLine ? 'Microsoft-Anmeldung erfolgreich.' : 'Änderungen werden lokal gespeichert.');
      if (navigator.onLine) await syncNow({interactive: false, reason: 'Start'});
    } else {
      setSyncStatus('signedout', 'Nicht angemeldet', 'Mit Microsoft anmelden, um PC und iPhone automatisch abzugleichen.');
    }
  } catch (error) {
    console.error(error);
    setSyncStatus('error', 'Anmeldefehler', error.message || String(error));
  }
}
async function loginMicrosoft() {
  if (!syncConfigured()) {
    alert('Bitte zuerst Client-ID und Mandanten-ID speichern.');
    return;
  }
  if (!msalApp) await initializeMicrosoftAuth();
  if (!msalApp) return;
  await msalApp.loginRedirect({scopes: GRAPH_SCOPES, redirectUri: syncConfig.redirectUri});
}
async function logoutMicrosoft() {
  if (!msalApp || !currentAccount) return;
  await msalApp.logoutRedirect({account: currentAccount, postLogoutRedirectUri: syncConfig.redirectUri});
}
async function acquireGraphToken(interactive = false) {
  if (!msalApp || !currentAccount) throw new Error('Nicht bei Microsoft angemeldet.');
  try {
    const result = await msalApp.acquireTokenSilent({scopes: GRAPH_SCOPES, account: currentAccount});
    return result.accessToken;
  } catch (error) {
    const needsInteraction = ['interaction_required', 'login_required', 'consent_required'].includes(error?.errorCode);
    if (interactive && needsInteraction) {
      await msalApp.acquireTokenRedirect({scopes: GRAPH_SCOPES, account: currentAccount, redirectUri: syncConfig.redirectUri});
      throw new Error('Microsoft-Anmeldung wird geöffnet.');
    }
    throw error;
  }
}
async function graphFetch(path, options = {}, interactive = false) {
  const token = await acquireGraphToken(interactive);
  const headers = new Headers(options.headers || {});
  headers.set('Authorization', `Bearer ${token}`);
  const response = await fetch(`${GRAPH_BASE}${path}`, {...options, headers});
  return response;
}
async function ensureAppFolder(interactive = false) {
  if (appFolderId) return {id: appFolderId, webUrl: appFolderWebUrl};
  const response = await graphFetch('/me/drive/special/approot?$select=id,name,webUrl', {}, interactive);
  if (!response.ok) throw new Error(`OneDrive-App-Ordner konnte nicht geöffnet werden (${response.status}).`);
  const folder = await response.json();
  appFolderId = folder.id;
  appFolderWebUrl = folder.webUrl || '';
  renderSyncPanel();
  return folder;
}
async function readRemoteState(interactive = false) {
  const folder = await ensureAppFolder(interactive);
  const name = encodeURIComponent(syncConfig.fileName || CLOUD_FILE_DEFAULT);
  const metadataResponse = await graphFetch(`/me/drive/items/${folder.id}:/${name}?$select=id,name,eTag,lastModifiedDateTime,webUrl`, {}, interactive);
  if (metadataResponse.status === 404) return {exists: false, etag: '', data: null};
  if (!metadataResponse.ok) throw new Error(`OneDrive-Datendatei konnte nicht gelesen werden (${metadataResponse.status}).`);
  const item = await metadataResponse.json();
  const contentResponse = await graphFetch(`/me/drive/items/${item.id}/content`, {}, interactive);
  if (!contentResponse.ok) throw new Error(`OneDrive-Datendatei konnte nicht heruntergeladen werden (${contentResponse.status}).`);
  const text = (await contentResponse.text()).replace(/^\uFEFF/, '');
  const data = JSON.parse(text);
  if (data.schema !== CLOUD_SCHEMA && !Array.isArray(data.cases)) throw new Error('Die OneDrive-Datei enthält kein gültiges JH-Assist-Datenformat.');
  return {exists: true, etag: item.eTag || '', data};
}
async function writeRemoteState(data, etag = '', interactive = false) {
  const folder = await ensureAppFolder(interactive);
  const name = encodeURIComponent(syncConfig.fileName || CLOUD_FILE_DEFAULT);
  const headers = {'Content-Type': 'application/json; charset=utf-8'};
  if (etag) headers['If-Match'] = etag;
  const response = await graphFetch(`/me/drive/items/${folder.id}:/${name}:/content`, {
    method: 'PUT', headers, body: JSON.stringify(data, null, 2)
  }, interactive);
  if (response.status === 412) {
    const error = new Error('Synchronisationskonflikt');
    error.code = 'etag-conflict';
    throw error;
  }
  if (!response.ok) throw new Error(`OneDrive-Datendatei konnte nicht gespeichert werden (${response.status}).`);
  return response.json();
}
async function performSync(interactive, retry = 0) {
  const remote = await readRemoteState(interactive);
  const localBefore = normalizeState(state);
  const remoteNormalized = remote.exists ? normalizeState(remote.data) : defaultState();
  const merged = remote.exists ? mergeStates(localBefore, remoteNormalized) : localBefore;
  const localChanged = fingerprint(merged) !== fingerprint(localBefore);
  const remoteChanged = !remote.exists || fingerprint(merged) !== fingerprint(remoteNormalized);
  state = merged;
  if (localChanged) persistState({touch: false, scheduleSync: false, render: true});
  if (remoteChanged) {
    state.meta.updatedAt = nowIso();
    state.meta.updatedBy = deviceId;
    try {
      await writeRemoteState(state, remote.etag, interactive);
    } catch (error) {
      if (error.code === 'etag-conflict' && retry < 1) return performSync(interactive, retry + 1);
      throw error;
    }
    persistState({touch: false, scheduleSync: false, render: false});
  }
}
async function syncNow({interactive = false, reason = 'Manuell'} = {}) {
  if (!syncConfigured() || !currentAccount) {
    setSyncStatus('signedout', 'Nicht angemeldet', 'Die Daten bleiben lokal gespeichert.');
    return false;
  }
  if (!navigator.onLine) {
    setSyncStatus('offline', 'Offline', 'Änderungen bleiben lokal und werden später übertragen.');
    return false;
  }
  if (syncInFlight) return syncInFlight;
  syncInFlight = (async () => {
    try {
      setSyncStatus('syncing', 'Synchronisiere…', reason);
      await performSync(interactive);
      const label = new Date().toLocaleTimeString('de-DE', {hour: '2-digit', minute: '2-digit'});
      setSyncStatus('synced', 'Synchronisiert', `Stand ${label}`);
      renderAll();
      return true;
    } catch (error) {
      console.error(error);
      const text = error?.message || String(error);
      if (/Anmeldung wird geöffnet/.test(text)) return false;
      setSyncStatus('error', 'Sync-Fehler', text);
      return false;
    } finally {
      syncInFlight = null;
    }
  })();
  return syncInFlight;
}
function queueCloudSync() {
  if (!syncConfigured() || !currentAccount) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => syncNow({interactive: false, reason: 'Lokale Änderung'}), 1200);
}

$('syncConfigForm').onsubmit = e => {
  e.preventDefault();
  const clientId = $('syncClientId').value.trim();
  const tenantId = $('syncTenantId').value.trim();
  const redirectUri = $('syncRedirectUri').value.trim();
  if (!isGuid(clientId) || !isGuid(tenantId)) {
    alert('Client-ID und Mandanten-ID müssen gültige GUIDs sein.');
    return;
  }
  if (!/^https:\/\//i.test(redirectUri)) {
    alert('Die Redirect-URI muss eine HTTPS-Adresse sein.');
    return;
  }
  localStorage.setItem(SYNC_CONFIG_KEY, JSON.stringify({clientId, tenantId, redirectUri}));
  alert('Microsoft-Konfiguration gespeichert. Die App wird neu geladen.');
  location.reload();
};
$('syncLoginBtn').onclick = () => loginMicrosoft().catch(error => { console.error(error); setSyncStatus('error', 'Anmeldefehler', error.message || String(error)); });
$('syncLogoutBtn').onclick = () => logoutMicrosoft().catch(error => { console.error(error); setSyncStatus('error', 'Abmeldefehler', error.message || String(error)); });
$('syncNowBtn').onclick = () => syncNow({interactive: true, reason: 'Manuell'});

window.addEventListener('online', () => syncNow({interactive: false, reason: 'Internetverbindung wiederhergestellt'}));
window.addEventListener('offline', () => setSyncStatus('offline', 'Offline', 'Änderungen bleiben lokal gespeichert.'));
window.addEventListener('focus', () => { if (document.visibilityState === 'visible') syncNow({interactive: false, reason: 'App geöffnet'}); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') syncNow({interactive: false, reason: 'App geöffnet'}); });
setInterval(() => { if (document.visibilityState === 'visible') syncNow({interactive: false, reason: 'Automatischer Abgleich'}); }, 60000);

/* Direkter Excel-Export über den lokalen Windows-Helfer */
function validateBillingSelection(year, month) {
  const monthStart = `${year}-${pad2(month)}-01`;
  const monthEnd = monthLastDate(year, month);
  const cases = state.cases.filter(c => (!c.approvalFrom || c.approvalFrom <= monthEnd) && (!c.approvalTo || c.approvalTo >= monthStart));
  if (!cases.length) return 'Für diesen Monat gibt es keine abzurechnenden Fälle.';
  if (!state.prices.some(p => Number(p.year) === year)) return `Für ${year} sind noch keine Preise hinterlegt.`;
  return '';
}
async function startExcelExport() {
  const year = Number($('billingYear').value);
  const month = Number($('billingMonth').value);
  const msg = $('billingMessage');
  const validation = validateBillingSelection(year, month);
  if (validation) { msg.textContent = validation; return; }
  if (!isWindows()) {
    msg.textContent = 'Die Excel-Abrechnung wird am Windows-PC erstellt. Öffne JH Assist dort und starte den Export erneut.';
    return;
  }
  if (!currentAccount) {
    msg.textContent = 'Bitte zuerst unter Einstellungen mit Microsoft anmelden.';
    return;
  }
  if (!navigator.onLine) {
    msg.textContent = 'Für die Excel-Abrechnung muss der aktuelle Datenstand zuerst mit OneDrive synchronisiert werden.';
    return;
  }
  const requiredUpdatedAt = state.meta.updatedAt || nowIso();
  void syncNow({interactive: false, reason: 'Vor Excel-Export'});
  msg.textContent = 'Der lokale Excel-Helfer wird geöffnet und wartet bei Bedarf kurz auf den OneDrive-Abgleich. Bestätige gegebenenfalls die Browserabfrage.';
  const protocolUrl = `jhassist://export?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&after=${encodeURIComponent(requiredUpdatedAt)}`;
  window.location.href = protocolUrl;
}
$('exportBillingDataBtn').onclick = startExcelExport;
$('billingMonth').onchange = renderBillingStatus;
$('billingYear').oninput = renderBillingStatus;

const quarterTimes = [];
for (let h = 0; h < 24; h++) for (const m of [0, 15, 30, 45]) quarterTimes.push(`${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`);
const timeOptions = quarterTimes.map(t => `<option value="${t}">${t}</option>`).join('');
$('start').innerHTML = timeOptions;
$('end').innerHTML = timeOptions;
$('start').value = '09:00';
$('end').value = '10:00';

const monthNames = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember'];
$('billingMonth').innerHTML = monthNames.map((m, i) => `<option value="${i + 1}">${m}</option>`).join('');
$('billingMonth').value = String(new Date().getMonth() + 1);
$('billingYear').value = new Date().getFullYear();
const reportStart = new Date();
reportStart.setMonth(reportStart.getMonth() - 6);
$('reportFrom').value = localIsoDate(reportStart);
$('reportTo').value = localIsoDate();
const today = localIsoDate();
$('date').value = today;
$('hoursYear').value = new Date().getFullYear();
$('priceYear').value = new Date().getFullYear();

persistState({touch: false, scheduleSync: false, render: true});
initializeMicrosoftAuth();
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(error => console.warn('Service Worker konnte nicht registriert werden.', error)));
}
