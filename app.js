'use strict';

const APP_VERSION = '1.1.0';
const STORAGE_KEY = 'jh-assist-v09';
const LEGACY_KEYS = [
  'jh-assist-v08',
  'jh-assist-v07',
  'jh-assist-v06',
  'jh-assist-v03',
  'jh-assist-v02',
  'jh-assist-v01'
];
const DEFAULT_ACTIVITY = 'einzelfallbezogene Tätigkeit';
const LEGACY_TIMESTAMP = '1970-01-01T00:00:00.000Z';

const $ = id => document.getElementById(id);
const esc = (value = '') => String(value).replace(/[&<>"']/g, character => ({
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
}[character]));
const uuid = () => globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const pad2 = value => String(value).padStart(2, '0');
const nowIso = () => new Date().toISOString();

function safeParse(text, fallback = null) {
  try {
    return text ? JSON.parse(text) : fallback;
  } catch (error) {
    console.warn('Gespeicherte Daten konnten nicht gelesen werden.', error);
    return fallback;
  }
}

function localDateString(date = new Date()) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function lastDateOfMonth(year, month) {
  const day = new Date(year, month, 0).getDate();
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

function validTimestamp(value, fallback = LEGACY_TIMESTAMP) {
  const parsed = Date.parse(String(value || ''));
  return Number.isFinite(parsed) ? new Date(parsed).toISOString() : fallback;
}

function defaultState() {
  return {
    cases: [],
    services: [],
    prices: [],
    auditLog: [],
    billedMonths: {},
    deletedServices: [],
    deletedServiceIds: [],
    syncMigratedAt: nowIso()
  };
}

function firstLegacyState() {
  for (const key of LEGACY_KEYS) {
    const parsed = safeParse(localStorage.getItem(key));
    if (parsed) return parsed;
  }
  return null;
}

function normalizeCase(item = {}) {
  const name = String(item.name || '').trim();
  const createdAt = validTimestamp(item.createdAt || item.updatedAt);
  return {
    id: String(item.id || uuid()),
    name,
    authority: item.authority === 'Stadt' ? 'Stadt' : 'Landkreis',
    weeklyHours: Number.isFinite(Number(item.weeklyHours)) ? Number(item.weeklyHours) : 0,
    officer: String(item.officer || ''),
    guardians: String(item.guardians || ''),
    approvalFrom: String(item.approvalFrom || '').slice(0, 10),
    approvalTo: String(item.approvalTo || '').slice(0, 10),
    sheetName: String(item.sheetName || name).trim() || name,
    active: item.active !== false,
    createdAt,
    updatedAt: validTimestamp(item.updatedAt || createdAt)
  };
}

function normalizeService(item = {}) {
  const createdAt = validTimestamp(item.createdAt || item.updatedAt);
  return {
    id: String(item.id || uuid()),
    caseId: String(item.caseId || ''),
    date: String(item.date || '').slice(0, 10),
    start: String(item.start || ''),
    end: String(item.end || ''),
    activity: String(item.activity || DEFAULT_ACTIVITY).trim() || DEFAULT_ACTIVITY,
    note: String(item.note || ''),
    useForReport: item.useForReport !== false,
    createdAt,
    updatedAt: validTimestamp(item.updatedAt || createdAt)
  };
}

function normalizePrice(item = {}) {
  const createdAt = validTimestamp(item.createdAt || item.updatedAt);
  return {
    year: Number(item.year),
    city: Number(item.city || 0),
    county: Number(item.county || 0),
    createdAt,
    updatedAt: validTimestamp(item.updatedAt || createdAt)
  };
}

function normalizeAudit(item = {}) {
  const timestamp = validTimestamp(item.timestamp, nowIso());
  return {
    id: String(item.id || `${timestamp}|${item.caseId || ''}|${item.action || ''}|${item.field || ''}`),
    timestamp,
    caseId: String(item.caseId || ''),
    month: String(item.month || ''),
    action: String(item.action || ''),
    field: String(item.field || ''),
    oldValue: String(item.oldValue ?? ''),
    newValue: String(item.newValue ?? '')
  };
}

function normalizeDeletedService(item = {}, fallbackTime = nowIso()) {
  if (typeof item === 'string') return { id: item, deletedAt: fallbackTime };
  return {
    id: String(item.id || ''),
    deletedAt: validTimestamp(item.deletedAt, fallbackTime)
  };
}

function normalizeBilledMonths(input) {
  const result = {};
  if (!input || typeof input !== 'object') return result;
  for (const [key, value] of Object.entries(input)) {
    if (!value || typeof value !== 'object') continue;
    const exportedAt = validTimestamp(value.exportedAt || value.updatedAt);
    result[key] = {
      exportedAt,
      status: value.status === 'changed_after_billing' ? 'changed_after_billing' : 'billed',
      updatedAt: validTimestamp(value.updatedAt || exportedAt)
    };
  }
  return result;
}

function normalizeState(input) {
  const source = input && typeof input === 'object' ? input : defaultState();
  const migrationTime = validTimestamp(source.syncMigratedAt, nowIso());
  const deletedSource = Array.isArray(source.deletedServices)
    ? source.deletedServices
    : Array.isArray(source.deletedServiceIds)
      ? source.deletedServiceIds
      : [];
  const deletedServices = deletedSource
    .map(item => normalizeDeletedService(item, migrationTime))
    .filter(item => item.id);
  const latestDeletion = new Map();
  for (const item of deletedServices) {
    const current = latestDeletion.get(item.id);
    if (!current || Date.parse(item.deletedAt) >= Date.parse(current.deletedAt)) latestDeletion.set(item.id, item);
  }
  const normalizedDeleted = [...latestDeletion.values()];
  return {
    cases: Array.isArray(source.cases) ? source.cases.map(normalizeCase) : [],
    services: Array.isArray(source.services) ? source.services.map(normalizeService) : [],
    prices: Array.isArray(source.prices) ? source.prices.map(normalizePrice).filter(price => Number.isFinite(price.year)) : [],
    auditLog: Array.isArray(source.auditLog) ? source.auditLog.map(normalizeAudit) : [],
    billedMonths: normalizeBilledMonths(source.billedMonths),
    deletedServices: normalizedDeleted,
    deletedServiceIds: normalizedDeleted.map(item => item.id),
    syncMigratedAt: migrationTime
  };
}

function timestampOf(item, field = 'updatedAt') {
  const parsed = Date.parse(item?.[field] || item?.createdAt || item?.timestamp || '');
  return Number.isFinite(parsed) ? parsed : 0;
}

function mergeRecords(localItems, remoteItems, keyOf, timestampField = 'updatedAt') {
  const map = new Map();
  for (const item of localItems || []) map.set(keyOf(item), item);
  for (const item of remoteItems || []) {
    const key = keyOf(item);
    if (key === undefined || key === null || key === '') continue;
    const current = map.get(key);
    if (!current || timestampOf(item, timestampField) >= timestampOf(current, timestampField)) map.set(key, item);
  }
  return [...map.values()];
}

function mergeStates(localInput, remoteInput) {
  const local = normalizeState(localInput);
  if (!remoteInput) return local;
  const remote = normalizeState(remoteInput);

  const deletedServices = mergeRecords(local.deletedServices, remote.deletedServices, item => item.id, 'deletedAt')
    .map(item => normalizeDeletedService(item));
  const deletionMap = new Map(deletedServices.map(item => [item.id, Date.parse(item.deletedAt)]));
  const services = mergeRecords(local.services, remote.services, item => item.id)
    .map(normalizeService)
    .filter(item => (deletionMap.get(item.id) || 0) < timestampOf(item));

  const billedMonths = {};
  for (const key of new Set([...Object.keys(local.billedMonths), ...Object.keys(remote.billedMonths)])) {
    const left = local.billedMonths[key];
    const right = remote.billedMonths[key];
    billedMonths[key] = !left ? right : !right ? left : timestampOf(right) >= timestampOf(left) ? right : left;
  }

  return normalizeState({
    cases: mergeRecords(local.cases, remote.cases, item => item.id),
    services,
    prices: mergeRecords(local.prices, remote.prices, item => String(item.year)),
    auditLog: mergeRecords(local.auditLog, remote.auditLog, item => item.id, 'timestamp')
      .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)))
      .slice(0, 1000),
    billedMonths,
    deletedServices,
    syncMigratedAt: local.syncMigratedAt || remote.syncMigratedAt || nowIso()
  });
}

const storedState = safeParse(localStorage.getItem(STORAGE_KEY));
let state = normalizeState(storedState || firstLegacyState() || defaultState());
localStorage.setItem(STORAGE_KEY, JSON.stringify(state));

function duration(start, end) {
  if (!start || !end) return 0;
  return Math.max(0, (minutesOf(end) - minutesOf(start)) / 60);
}

function formatHours(value) {
  return `${Number(value || 0).toLocaleString('de-DE', { minimumFractionDigits: 0, maximumFractionDigits: 2 })} Std.`;
}

function dateParts(value) {
  const [year, month, day] = String(value || '').split('-').map(Number);
  return { year, month, day };
}

function weeksBetween(from, to) {
  if (!from || !to) return 0;
  const first = dateParts(from);
  const last = dateParts(to);
  if (![first.year, first.month, first.day, last.year, last.month, last.day].every(Number.isFinite)) return 0;
  const firstUtc = Date.UTC(first.year, first.month - 1, first.day);
  const lastUtc = Date.UTC(last.year, last.month - 1, last.day);
  if (lastUtc < firstUtc) return 0;
  return ((lastUtc - firstUtc) / 86400000 + 1) / 7;
}

function persistLocal(options = {}) {
  const render = options === true || options?.render !== false;
  const sync = options !== true && options?.sync === true;
  state = normalizeState(state);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  if (render) renderAll();
  if (sync) globalThis.JHOneDrive?.scheduleSync?.();
}

function save(options = {}) {
  persistLocal({ render: options.render !== false, sync: options.sync !== false });
}

function applyCloudState(nextState) {
  state = normalizeState(nextState);
  persistLocal({ render: true, sync: false });
}

function monthKey(dateOrYear, month) {
  return month === undefined
    ? String(dateOrYear).slice(0, 7)
    : `${dateOrYear}-${pad2(month)}`;
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

function logChange({ caseId = '', month = '', action, field = '', oldValue = '', newValue = '' }) {
  state.auditLog.unshift(normalizeAudit({
    id: uuid(),
    timestamp: new Date().toISOString(),
    caseId,
    month,
    action,
    field,
    oldValue,
    newValue
  }));
  if (state.auditLog.length > 1000) state.auditLog.length = 1000;
  if (month) setChangedAfterBilling(month);
}

function allCases() {
  return [...state.cases].sort((a, b) => a.name.localeCompare(b.name, 'de'));
}

function activeCases() {
  return allCases().filter(item => item.active !== false);
}

function caseById(id) {
  return state.cases.find(item => item.id === id);
}

function showView(name) {
  document.querySelectorAll('.view').forEach(view => view.classList.add('hidden'));
  $(`view-${name}`).classList.remove('hidden');
  document.querySelectorAll('.tab').forEach(button => button.classList.toggle('active', button.dataset.view === name));
}

document.querySelectorAll('.tab').forEach(button => {
  button.onclick = () => showView(button.dataset.view);
});

function fillSelect(element, list, currentValue, includeArchiveLabel = false) {
  element.innerHTML = list.length
    ? list.map(item => `<option value="${esc(item.id)}">${esc(item.name)}${includeArchiveLabel && item.active === false ? ' (archiviert)' : ''}</option>`).join('')
    : '<option value="">Zuerst Fall anlegen</option>';
  if ([...element.options].some(option => option.value === currentValue)) element.value = currentValue;
}

function renderSelects() {
  const active = activeCases();
  const all = allCases();
  fillSelect($('serviceCase'), active, $('serviceCase').value);
  fillSelect($('hoursCase'), all, $('hoursCase').value, true);
  fillSelect($('reportCase'), all, $('reportCase').value, true);
}

function renderCases() {
  const query = $('caseSearch').value.toLowerCase();
  const filter = $('caseFilter').value;
  const list = state.cases
    .filter(item => (
      filter === 'all' ||
      (filter === 'active' && item.active !== false) ||
      (filter === 'inactive' && item.active === false)
    ) && [item.name, item.authority, item.officer, item.guardians].join(' ').toLowerCase().includes(query))
    .sort((a, b) => a.name.localeCompare(b.name, 'de'));

  $('cases').innerHTML = list.length
    ? list.map(item => `<div class="item">
        <strong>${esc(item.name)}</strong>${item.active === false ? '<span class="badge">archiviert</span>' : ''}
        <div class="muted">${esc(item.authority)} · ${formatHours(item.weeklyHours)} pro Woche · ${esc(item.officer || 'keine Sachbearbeitung')}</div>
        <div class="muted">Bewilligung: ${esc(item.approvalFrom || 'offen')} bis ${esc(item.approvalTo || 'offen')}</div>
        <div class="row" style="margin-top:8px">
          <button class="secondary small" onclick="editCase('${item.id}')">Bearbeiten</button>
          <button class="${item.active === false ? 'secondary' : 'danger'} small" onclick="toggleCase('${item.id}')">${item.active === false ? 'Reaktivieren' : 'Archivieren'}</button>
        </div>
      </div>`).join('')
    : '<p class="muted">Keine passenden Fälle.</p>';
}

function renderServices() {
  const sorted = [...state.services]
    .sort((a, b) => `${b.date}${b.start}`.localeCompare(`${a.date}${a.start}`))
    .slice(0, 15);
  $('services').innerHTML = sorted.length
    ? sorted.map(service => {
      const relatedCase = caseById(service.caseId);
      const key = monthKey(service.date);
      const changed = isBilledMonth(key) && state.billedMonths[key].status === 'changed_after_billing';
      return `<div class="item">
        <strong>${esc(relatedCase?.name || 'Unbekannter Fall')}</strong>
        <span class="badge">${formatHours(duration(service.start, service.end))}</span>
        ${changed ? '<span class="badge warn">nach Abrechnung geändert</span>' : ''}
        <div>${esc(service.date)} · ${esc(service.start)}–${esc(service.end)}</div>
        <div class="muted">${esc(service.activity)}${service.note ? ` · ${esc(service.note)}` : ''}</div>
        <div class="row" style="margin-top:7px">
          <button class="secondary small" onclick="editService('${service.id}')">Bearbeiten</button>
          <button class="danger small" onclick="deleteService('${service.id}')">Löschen</button>
        </div>
      </div>`;
    }).join('')
    : '<p class="muted">Noch keine Leistungen erfasst.</p>';
}

function renderPrices() {
  $('prices').innerHTML = state.prices.length
    ? [...state.prices].sort((a, b) => b.year - a.year).map(price => `<div class="item">
        <strong>${price.year}</strong>
        <div class="muted">Stadt: ${Number(price.city).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' })} · Landkreis: ${Number(price.county).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' })}</div>
      </div>`).join('')
    : '<p class="muted">Noch keine Preise hinterlegt.</p>';
}

function renderHours() {
  const caseId = $('hoursCase').value;
  const selectedCase = caseById(caseId);
  const year = Number($('hoursYear').value);
  if (!selectedCase) {
    $('hoursSummary').innerHTML = '<p class="muted">Kein Fall vorhanden.</p>';
    $('notesList').innerHTML = '';
    return;
  }

  const services = state.services.filter(service => service.caseId === caseId && Number(service.date.slice(0, 4)) === year);
  const completed = services.reduce((sum, service) => sum + duration(service.start, service.end), 0);
  let from = selectedCase.approvalFrom;
  let to = selectedCase.approvalTo;
  if (from && Number(from.slice(0, 4)) < year) from = `${year}-01-01`;
  if (to && Number(to.slice(0, 4)) > year) to = `${year}-12-31`;
  const approved = weeksBetween(from, to) * Number(selectedCase.weeklyHours || 0);
  const remaining = approved - completed;
  const price = state.prices.find(item => Number(item.year) === year);
  const rate = price ? (selectedCase.authority === 'Stadt' ? price.city : price.county) : 0;

  $('hoursSummary').innerHTML = `<div class="stats" style="margin-top:16px">
      <div class="stat"><span class="muted">Jahreskontingent rechnerisch</span><strong>${formatHours(approved)}</strong></div>
      <div class="stat"><span class="muted">Geleistet</span><strong>${formatHours(completed)}</strong></div>
      <div class="stat"><span class="muted">Rest</span><strong>${formatHours(remaining)}</strong></div>
    </div>
    <p class="muted">Grundlage: ${selectedCase.weeklyHours} Wochenstunden innerhalb des im Kalenderjahr liegenden Bewilligungszeitraums.${rate ? ` Aktueller Satz: ${Number(rate).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' })}.` : ''}</p>
    ${remaining < 0 ? '<p class="message" style="color:#8b3030">Achtung: Das rechnerische Kontingent ist überschritten.</p>' : ''}`;

  const notes = services.filter(service => service.note).sort((a, b) => b.date.localeCompare(a.date));
  $('notesList').innerHTML = notes.length
    ? notes.map(service => `<div class="item"><strong>${esc(service.date)}</strong><div>${esc(service.note)}</div><div class="muted">${esc(service.start)}–${esc(service.end)}</div></div>`).join('')
    : '<p class="muted">Für dieses Jahr sind keine Verlaufsnotizen gespeichert.</p>';
}

function renderBillingStatus() {
  const key = monthKey(Number($('billingYear').value), Number($('billingMonth').value));
  const entry = state.billedMonths[key];
  $('billingStatus').innerHTML = !entry
    ? 'Status: noch nicht abgerechnet'
    : entry.status === 'changed_after_billing'
      ? `<span class="badge warn">geändert nach Abrechnung</span> Letzter Export: ${esc(new Date(entry.exportedAt).toLocaleString('de-DE'))}`
      : `<span class="badge ok">abgerechnet</span> Export: ${esc(new Date(entry.exportedAt).toLocaleString('de-DE'))}`;
}

function renderAuditLog() {
  const items = state.auditLog.slice(0, 100);
  $('auditLog').innerHTML = items.length
    ? items.map(item => {
      const relatedCase = caseById(item.caseId);
      return `<div class="item">
        <strong>${esc(new Date(item.timestamp).toLocaleString('de-DE'))}</strong>${item.month ? ` <span class="badge">${esc(item.month)}</span>` : ''}
        <div>${esc(relatedCase?.name || 'Allgemein')} – ${esc(item.action)}${item.field ? ` – ${esc(item.field)}` : ''}</div>
        ${item.oldValue || item.newValue ? `<div class="muted">${esc(item.oldValue)} → ${esc(item.newValue)}</div>` : ''}
      </div>`;
    }).join('')
    : '<p class="muted">Noch keine Änderungen nach einer Abrechnung protokolliert.</p>';
}

function renderAll() {
  renderSelects();
  renderCases();
  renderServices();
  renderPrices();
  renderHours();
  renderBillingStatus();
  renderAuditLog();
}

$('caseForm').onsubmit = event => {
  event.preventDefault();
  const id = $('caseId').value;
  const name = $('caseName').value.trim();
  const sheetName = $('sheetName').value.trim() || name;
  if (/[\\/?*\[\]:]/.test(sheetName) || sheetName.length > 31) {
    alert('Der Tabellenblattname darf höchstens 31 Zeichen lang sein und keine Zeichen \\ / ? * [ ] : enthalten.');
    return;
  }
  const reservedSheetNames = new Set(['gesamt', 'stadt', 'landkreis', 'geldbedarf']);
  if (reservedSheetNames.has(sheetName.toLocaleLowerCase('de-DE'))) {
    alert('Dieser Tabellenblattname ist für ein Gesamt- oder Berechnungsblatt reserviert.');
    return;
  }
  if (state.cases.some(item => item.id !== id && item.sheetName.toLocaleLowerCase('de-DE') === sheetName.toLocaleLowerCase('de-DE'))) {
    alert('Dieser Tabellenblattname ist bereits einem anderen Fall zugeordnet.');
    return;
  }
  const existing = id ? caseById(id) : null;
  const item = normalizeCase({
    id: id || uuid(),
    name,
    authority: $('authority').value,
    weeklyHours: Number($('weeklyHours').value),
    officer: $('officer').value.trim(),
    guardians: $('guardians').value.trim(),
    approvalFrom: $('approvalFrom').value,
    approvalTo: $('approvalTo').value,
    sheetName,
    active: existing ? existing.active !== false : true,
    createdAt: existing?.createdAt || nowIso(),
    updatedAt: nowIso()
  });
  if (item.approvalFrom && item.approvalTo && item.approvalTo < item.approvalFrom) {
    alert('Das Bewilligungsende liegt vor dem Beginn.');
    return;
  }
  if (id) state.cases[state.cases.findIndex(entry => entry.id === id)] = item;
  else state.cases.push(item);
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
  const item = caseById(id);
  if (!item) return;
  showView('cases');
  $('caseId').value = item.id;
  $('caseName').value = item.name;
  $('authority').value = item.authority;
  $('weeklyHours').value = item.weeklyHours;
  $('officer').value = item.officer || '';
  $('guardians').value = item.guardians || '';
  $('approvalFrom').value = item.approvalFrom || '';
  $('approvalTo').value = item.approvalTo || '';
  $('sheetName').value = item.sheetName || item.name;
  $('caseHeading').textContent = 'Fall bearbeiten';
  $('cancelCaseEdit').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
};

window.toggleCase = id => {
  const item = caseById(id);
  if (!item) return;
  item.active = item.active === false;
  item.updatedAt = nowIso();
  save();
};

$('cancelCaseEdit').onclick = resetCaseForm;
$('caseSearch').oninput = renderCases;
$('caseFilter').onchange = renderCases;

function minutesOf(time) {
  const [hour, minute] = String(time).split(':').map(Number);
  return hour * 60 + minute;
}

function overlaps(firstStart, firstEnd, secondStart, secondEnd) {
  return minutesOf(firstStart) < minutesOf(secondEnd) && minutesOf(secondStart) < minutesOf(firstEnd);
}

function isQuarterHour(time) {
  return /^(?:[01]\d|2[0-3]):(?:00|15|30|45)$/.test(time);
}

$('serviceForm').onsubmit = event => {
  event.preventDefault();
  const id = $('serviceId').value;
  const caseId = $('serviceCase').value;
  const date = $('date').value;
  const start = $('start').value;
  const end = $('end').value;
  const selectedCase = caseById(caseId);
  const existing = id ? state.services.find(item => item.id === id) : null;
  $('serviceMessage').textContent = '';
  if (!selectedCase) return;
  if (!isQuarterHour(start) || !isQuarterHour(end)) {
    $('serviceMessage').textContent = 'Beginn und Ende müssen im Viertelstundentakt liegen.';
    return;
  }
  if ((selectedCase.approvalFrom && date < selectedCase.approvalFrom) || (selectedCase.approvalTo && date > selectedCase.approvalTo)) {
    if (!confirm('Das Datum liegt außerhalb des hinterlegten Bewilligungszeitraums. Trotzdem speichern?')) return;
  }
  if (state.services.some(item => item.id !== id && item.caseId === caseId && item.date === date)) {
    $('serviceMessage').textContent = 'Für diesen Fall ist an diesem Tag bereits ein Eintrag vorhanden.';
    return;
  }
  if (minutesOf(end) <= minutesOf(start)) {
    $('serviceMessage').textContent = 'Die Endzeit muss nach der Beginnzeit liegen.';
    return;
  }
  const conflict = state.services.find(item => item.id !== id && item.date === date && overlaps(start, end, item.start, item.end));
  if (conflict) {
    const otherCase = caseById(conflict.caseId);
    $('serviceMessage').textContent = `Terminüberschneidung mit ${otherCase?.name || 'einem anderen Fall'} (${conflict.start}–${conflict.end}). Direkte Anschlusstermine sind erlaubt.`;
    return;
  }

  const item = normalizeService({
    id: id || uuid(),
    caseId,
    date,
    start,
    end,
    activity: $('activity').value.trim(),
    note: $('note').value.trim(),
    useForReport: $('reportUse').checked,
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });

  if (existing) {
    const trackedFields = {
      Fall: [existing.caseId, item.caseId],
      Datum: [existing.date, item.date],
      Beginn: [existing.start, item.start],
      Ende: [existing.end, item.end],
      Tätigkeit: [existing.activity, item.activity],
      Verlaufsnotiz: [existing.note, item.note],
      'Für Verlaufsdokumentation': [existing.useForReport !== false, item.useForReport !== false]
    };
    const affectedMonths = new Set([monthKey(existing.date), monthKey(item.date)]);
    for (const [field, [oldValue, newValue]] of Object.entries(trackedFields)) {
      if (String(oldValue ?? '') !== String(newValue ?? '')) {
        for (const key of affectedMonths) {
          if (isBilledMonth(key)) logChange({ caseId: item.caseId, month: key, action: 'Leistung geändert', field, oldValue, newValue });
        }
      }
    }
    state.services[state.services.findIndex(service => service.id === id)] = item;
    $('serviceMessage').textContent = 'Leistung geändert.';
  } else {
    state.services.push(item);
    const key = monthKey(date);
    if (isBilledMonth(key)) {
      logChange({ caseId, month: key, action: 'Leistung nach Abrechnung hinzugefügt', field: 'Termin', oldValue: '', newValue: `${date} ${start}–${end}` });
    }
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
  $('activity').value = DEFAULT_ACTIVITY;
}

window.editService = id => {
  const item = state.services.find(service => service.id === id);
  if (!item) return;
  showView('entry');
  if (![...$('serviceCase').options].some(option => option.value === item.caseId)) {
    const archivedCase = caseById(item.caseId);
    if (archivedCase) $('serviceCase').add(new Option(`${archivedCase.name} (archiviert)`, archivedCase.id));
  }
  $('serviceId').value = item.id;
  $('serviceCase').value = item.caseId;
  $('date').value = item.date;
  $('start').value = item.start;
  $('end').value = item.end;
  $('activity').value = item.activity || DEFAULT_ACTIVITY;
  $('note').value = item.note || '';
  $('reportUse').checked = item.useForReport !== false;
  $('serviceHeading').textContent = 'Leistung bearbeiten';
  $('saveServiceBtn').textContent = 'Änderung speichern';
  $('cancelServiceEdit').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
};

$('cancelServiceEdit').onclick = resetServiceForm;

window.deleteService = id => {
  const item = state.services.find(service => service.id === id);
  if (!item || !confirm('Diesen Leistungseintrag wirklich löschen?')) return;
  const key = monthKey(item.date);
  if (isBilledMonth(key)) {
    logChange({ caseId: item.caseId, month: key, action: 'Leistung nach Abrechnung gelöscht', field: 'Termin', oldValue: `${item.date} ${item.start}–${item.end}`, newValue: '' });
  }
  state.services = state.services.filter(service => service.id !== id);
  const deletedAt = nowIso();
  state.deletedServices = state.deletedServices.filter(entry => entry.id !== id);
  state.deletedServices.push({ id, deletedAt });
  state.deletedServiceIds = state.deletedServices.map(entry => entry.id);
  save();
};

$('priceForm').onsubmit = event => {
  event.preventDefault();
  const year = Number($('priceYear').value);
  const existing = state.prices.find(price => Number(price.year) === year);
  const item = normalizePrice({
    year,
    city: $('priceCity').value,
    county: $('priceCounty').value,
    createdAt: existing?.createdAt || nowIso(),
    updatedAt: nowIso()
  });
  if (!Number.isFinite(item.year) || item.city < 0 || item.county < 0) {
    alert('Bitte gültige Preise eintragen.');
    return;
  }
  const index = state.prices.findIndex(price => Number(price.year) === year);
  if (index >= 0) state.prices[index] = item;
  else state.prices.push(item);
  save();
};

$('hoursCase').onchange = renderHours;
$('hoursYear').oninput = renderHours;

$('exportBtn').onclick = async () => {
  await exportJsonFile({
    schema: 'jh-assist-backup-v1',
    version: APP_VERSION,
    exportedAt: new Date().toISOString(),
    ...state
  }, `JH-Assist-Vollsicherung-${localDateString()}.json`);
};

$('importFile').onchange = async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const data = safeParse(await file.text());
    if (!Array.isArray(data?.cases) || !Array.isArray(data?.services)) throw new Error('Ungültiges Format');
    if (confirm('Alle lokalen Daten durch diese Vollsicherung ersetzen?')) {
      state = normalizeState(data);
      save();
      alert('Die Vollsicherung wurde importiert.');
    }
  } catch (error) {
    alert('Die Datei ist keine gültige JH-Assist-Vollsicherung.');
  }
  event.target.value = '';
};

$('clearBtn').onclick = () => {
  if (confirm('Wirklich alle lokal gespeicherten Daten löschen?')) {
    state = defaultState();
    save();
  }
};

function downloadBlob(content, name, type = 'application/json') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    anchor.remove();
  }, 3000);
}

function isAppleMobile() {
  return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}

async function deliverFile(content, name, type = 'application/json') {
  if (isAppleMobile() && typeof File !== 'undefined' && navigator.share && navigator.canShare) {
    try {
      const file = new File([content], name, { type });
      if (navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: name });
        return 'shared';
      }
    } catch (error) {
      if (error?.name === 'AbortError') throw error;
      console.warn('Teilen nicht verfügbar; Datei wird stattdessen heruntergeladen.', error);
    }
  }
  downloadBlob(content, name, type);
  return 'downloaded';
}

function billingContext() {
  const year = Number($('billingYear').value);
  const month = Number($('billingMonth').value);
  const monthStart = `${year}-${pad2(month)}-01`;
  const monthEnd = lastDateOfMonth(year, month);
  const cases = state.cases.filter(item =>
    (!item.approvalFrom || item.approvalFrom <= monthEnd) &&
    (!item.approvalTo || item.approvalTo >= monthStart)
  );
  const price = state.prices.find(item => Number(item.year) === year);
  return { year, month, monthStart, monthEnd, cases, price };
}

function isWindowsDevice() {
  return /Windows/i.test(navigator.userAgent || '') || /Win/i.test(navigator.platform || '');
}

function launchExportHelper(year, month, cloudUpdatedAt = '') {
  const config = globalThis.JHOneDrive?.getConfig?.() || {};
  const parameters = new URLSearchParams({
    year: String(year),
    month: String(month),
    clientId: String(config.clientId || ''),
    tenantId: String(config.tenantId || '')
  });
  if (cloudUpdatedAt) parameters.set('sync', cloudUpdatedAt);
  const protocolUrl = `jhassist://export?${parameters.toString()}`;
  globalThis.location.href = protocolUrl;
}

async function exportBillingData() {
  const message = $('billingMessage');
  const { year, month, cases, price } = billingContext();
  if (!cases.length) {
    message.textContent = 'Für diesen Monat gibt es keine abzurechnenden Fälle.';
    return;
  }
  if (!price) {
    message.textContent = `Für ${year} sind noch keine Preise hinterlegt.`;
    return;
  }
  if (!globalThis.JHOneDrive?.configured?.()) {
    message.textContent = 'Bitte zuerst unter Einstellungen die OneDrive-Synchronisierung einrichten.';
    showView('settings');
    return;
  }
  if (!isWindowsDevice()) {
    message.textContent = 'Die Excel-Abrechnung wird auf dem Windows-PC erstellt. Die Daten sind bereits über OneDrive verfügbar.';
    return;
  }

  if (!globalThis.JHOneDrive.signedIn?.()) {
    message.textContent = 'Microsoft-Anmeldung wird geöffnet. Nach der Anmeldung bitte den Export erneut starten.';
    try { await globalThis.JHOneDrive.syncNow({ interactive: true }); }
    catch (error) { message.textContent = `Anmeldung konnte nicht gestartet werden: ${error.message}`; }
    return;
  }

  // Der Protokollaufruf muss direkt innerhalb des Tastendrucks erfolgen,
  // damit Browser ihn nicht als unerwünschten externen App-Start blockieren.
  // Der Exporthelfer wartet gleichzeitig darauf, dass dieser Datenstand in
  // OneDrive sichtbar ist. So bleibt der Vorgang ein einziger Klick.
  const syncTarget = nowIso();
  message.textContent = 'Aktuelle Daten werden synchronisiert; der Windows-Exporthelfer wird geöffnet …';
  const syncPromise = globalThis.JHOneDrive.syncNow({ interactive: false, throwOnError: true });
  launchExportHelper(year, month, syncTarget);
  syncPromise
    .then(result => {
      if (result?.status === 'synced') {
        message.textContent = 'Daten synchronisiert. Der Exporthelfer erstellt die Excel-Abrechnung.';
      }
    })
    .catch(error => {
      message.textContent = `Synchronisierung fehlgeschlagen: ${error.message}`;
    });
}

function reportServices() {
  const caseId = $('reportCase').value;
  const from = $('reportFrom').value;
  const to = $('reportTo').value;
  return state.services
    .filter(item => item.caseId === caseId && item.note && item.useForReport !== false && (!from || item.date >= from) && (!to || item.date <= to))
    .sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));
}

function createReport() {
  const selectedCase = caseById($('reportCase').value);
  const items = reportServices();
  const message = $('reportMessage');
  if (!selectedCase) {
    message.textContent = 'Bitte einen Fall auswählen.';
    return;
  }
  if (!items.length) {
    $('reportOutput').value = '';
    message.textContent = 'Im gewählten Zeitraum liegen keine für die Verlaufsdokumentation markierten Verlaufsnotizen vor.';
    return;
  }
  const from = $('reportFrom').value || items[0].date;
  const to = $('reportTo').value || items[items.length - 1].date;
  const lines = [
    `Verlaufsdokumentation – ${selectedCase.name}`,
    `Berichtszeitraum: ${from} bis ${to}`,
    '',
    'Grundlage',
    `Die folgende Verlaufsdokumentation beruht ausschließlich auf ${items.length} dokumentierten Verlaufsnotiz(en). Sie ist fachlich zu prüfen, zu gewichten und vor Verwendung zu überarbeiten.`,
    '',
    'Chronologischer Verlauf'
  ];
  for (const item of items) lines.push(`${item.date}, ${item.start}–${item.end}: ${item.note.trim()}`);
  lines.push('', 'Zusammenfassende Verlaufsdarstellung');
  lines.push(`Im Berichtszeitraum fanden ${items.length} dokumentierte Kontakte im Rahmen der einzelfallbezogenen Tätigkeit statt. Die nachfolgenden Feststellungen geben die dokumentierten Beobachtungen und Arbeitsinhalte wieder:`);
  for (const item of items) lines.push(`- ${item.note.trim()}`);
  lines.push('', 'Abschließende fachliche Einordnung');
  lines.push('Bitte ergänzen: Entwicklung im Berichtszeitraum, erreichte beziehungsweise nicht erreichte Hilfeplanziele, Ressourcen, bestehender Unterstützungsbedarf und fachliche Empfehlung für den weiteren Hilfeverlauf.');
  $('reportOutput').value = lines.join('\n');
  message.textContent = 'Verlaufsdokumentation erstellt. Es wurden keine nicht dokumentierten Tatsachen ergänzt.';
}

async function copyText(text) {
  if (navigator.clipboard?.writeText && globalThis.isSecureContext) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  textarea.remove();
}

$('exportBillingDataBtn').onclick = exportBillingData;
$('billingMonth').onchange = renderBillingStatus;
$('billingYear').oninput = renderBillingStatus;
$('createReportBtn').onclick = createReport;
$('copyReportBtn').onclick = async () => {
  if (!$('reportOutput').value) return;
  try {
    await copyText($('reportOutput').value);
    $('reportMessage').textContent = 'Verlaufsdokumentation wurde kopiert.';
  } catch (error) {
    $('reportMessage').textContent = 'Der Text konnte nicht kopiert werden.';
  }
};
$('downloadReportBtn').onclick = async () => {
  const selectedCase = caseById($('reportCase').value);
  if (!$('reportOutput').value) return;
  const safeName = (selectedCase?.name || 'Fall').replace(/[^a-z0-9äöüß_-]+/gi, '_');
  await deliverFile($('reportOutput').value, `Verlaufsdokumentation_${safeName}.txt`, 'text/plain;charset=utf-8');
};

function updateSyncStatus(status = {}) {
  const badge = $('syncBadge');
  const message = $('syncStatus');
  if (badge) {
    badge.className = `sync-pill ${status.state || 'local'}`;
    const labels = {
      synced: 'Synchronisiert', syncing: 'Synchronisiert …', ready: 'Angemeldet',
      offline: 'Offline', error: 'Fehler', signed_out: 'Nicht angemeldet', local: 'Lokal'
    };
    badge.textContent = labels[status.state] || 'OneDrive';
  }
  if (message) {
    const account = status.account ? ` Konto: ${status.account}.` : '';
    message.textContent = `${status.message || ''}${account}`.trim();
  }
  const link = $('syncFolderLink');
  if (link) {
    if (status.appFolderUrl) {
      link.href = status.appFolderUrl;
      link.classList.remove('hidden');
    } else {
      link.classList.add('hidden');
    }
  }
}

function fillSyncConfiguration() {
  if (!globalThis.JHOneDrive) return;
  const config = globalThis.JHOneDrive.getConfig();
  $('syncClientId').value = config.clientId || '';
  $('syncTenantId').value = config.tenantId || 'organizations';
  $('syncRedirectUri').value = config.redirectUri || globalThis.JHOneDrive.defaultRedirectUri();
  $('syncAuto').checked = config.autoSync !== false;
  updateSyncStatus(globalThis.JHOneDrive.getStatus());
}

$('syncConfigForm').onsubmit = async event => {
  event.preventDefault();
  const clientId = $('syncClientId').value.trim();
  const tenantId = $('syncTenantId').value.trim() || 'organizations';
  const redirectUri = $('syncRedirectUri').value.trim();
  if (!/^[0-9a-f-]{36}$/i.test(clientId)) {
    $('syncStatus').textContent = 'Die Application (Client) ID hat kein gültiges Format.';
    return;
  }
  try {
    await globalThis.JHOneDrive.reconfigure({ clientId, tenantId, redirectUri, autoSync: $('syncAuto').checked });
    $('syncStatus').textContent = 'Konfiguration gespeichert. Jetzt mit Microsoft anmelden.';
  } catch (error) {
    $('syncStatus').textContent = `Konfiguration konnte nicht aktiviert werden: ${error.message}`;
  }
};

$('syncSignInBtn').onclick = async () => {
  try { await globalThis.JHOneDrive.signIn(); }
  catch (error) { $('syncStatus').textContent = `Anmeldung konnte nicht gestartet werden: ${error.message}`; }
};

$('syncSignOutBtn').onclick = async () => {
  try { await globalThis.JHOneDrive.signOut(); }
  catch (error) { $('syncStatus').textContent = `Abmeldung konnte nicht gestartet werden: ${error.message}`; }
};

$('syncNowBtn').onclick = async () => {
  try {
    const result = await globalThis.JHOneDrive.syncNow({ interactive: true });
    if (result?.status === 'synced') $('syncStatus').textContent = `Synchronisiert: ${new Date(result.lastSyncAt).toLocaleString('de-DE')}`;
  } catch (error) {
    $('syncStatus').textContent = `Synchronisierung fehlgeschlagen: ${error.message}`;
  }
};

async function initializeOneDrive() {
  if (!globalThis.JHOneDrive) return;
  fillSyncConfiguration();
  await globalThis.JHOneDrive.initialize({
    version: APP_VERSION,
    getState: () => normalizeState(state),
    normalizeState,
    mergeStates,
    setState: nextState => {
      state = normalizeState(nextState);
      persistLocal({ render: true, sync: false });
    },
    onStatus: updateSyncStatus
  });
  fillSyncConfiguration();
}

const quarterTimes = [];
for (let hour = 0; hour < 24; hour += 1) {
  for (const minute of [0, 15, 30, 45]) quarterTimes.push(`${pad2(hour)}:${pad2(minute)}`);
}
const timeOptions = quarterTimes.map(time => `<option value="${time}">${time}</option>`).join('');
$('start').innerHTML = timeOptions;
$('end').innerHTML = timeOptions;
$('start').value = '09:00';
$('end').value = '10:00';

const monthNames = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
$('billingMonth').innerHTML = monthNames.map((name, index) => `<option value="${index + 1}">${name}</option>`).join('');
$('billingMonth').value = String(new Date().getMonth() + 1);
$('billingYear').value = new Date().getFullYear();

const reportStart = new Date();
reportStart.setMonth(reportStart.getMonth() - 6);
$('reportFrom').value = localDateString(reportStart);
$('reportTo').value = localDateString();
$('date').value = localDateString();
$('hoursYear').value = new Date().getFullYear();
$('priceYear').value = new Date().getFullYear();

renderAll();
initializeOneDrive().catch(error => console.error('OneDrive-Initialisierung fehlgeschlagen.', error));

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(error => console.warn('Service Worker konnte nicht registriert werden.', error)));
}
