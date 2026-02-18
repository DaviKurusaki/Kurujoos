const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { shell } = require('electron');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const https = require('https');
const XLSX = require('xlsx');

const SETTINGS_FILE = () => path.join(app.getPath('userData'), 'settings.json');
const DEFAULT_SETTINGS = {
  rootFolder: '',
  fiscalRoot: '',
  trackingApiKey: '',
  theme: 'light',
  lastLoginUsername: ''
};
const ADMIN_SETTINGS_PASSWORD = '0604';

const XLSX_READ_OPTIONS = { cellDates: true };
const XLSX_WORKBOOK_CACHE = new Map(); // filePath -> { mtimeMs, workbook }
const WORKBOOK_AOA_CACHE = new WeakMap(); // workbook -> Map(sheetName -> aoa)
const WORKBOOK_OBJECTS_CACHE = new WeakMap(); // workbook -> Map(sheetName -> objects[])
const RH_HORAS_MONTH_INDEX_CACHE = new WeakMap(); // workbook -> { source: objects[], byMonth: Map<YYYY-MM, objects[]> }
const RH_ENSURED_WORKBOOKS = new WeakSet();

const FAST_CACHE_VERSION = 1;
const FAST_CACHE_ROOT_FOLDER = () => path.join(app.getPath('userData'), 'fast-cache');

function hashText(value) {
  return crypto.createHash('sha1').update(String(value || ''), 'utf8').digest('hex');
}

function getFastCacheFilePath(tag, sourceFilePath) {
  const safeTag = String(tag || '').trim() || 'generic';
  const safeSource = String(sourceFilePath || '').trim();
  const fileName = `${hashText(safeSource)}.json`;
  return path.join(FAST_CACHE_ROOT_FOLDER(), safeTag, fileName);
}

function readFastCache(tag, sourceFilePath, expectedMtimeMs) {
  try {
    const cachePath = getFastCacheFilePath(tag, sourceFilePath);
    if (!fs.existsSync(cachePath)) {
      return null;
    }
    const text = fs.readFileSync(cachePath, 'utf8');
    const parsed = text ? JSON.parse(text) : null;
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }
    if (Number(parsed.version) !== FAST_CACHE_VERSION) {
      return null;
    }
    if (String(parsed.sourceFilePath || '') !== String(sourceFilePath || '')) {
      return null;
    }
    if (Number(parsed.sourceMtimeMs) !== Number(expectedMtimeMs || 0)) {
      return null;
    }
    return parsed.payload ?? null;
  } catch (error) {
    return null;
  }
}

function writeFastCache(tag, sourceFilePath, sourceMtimeMs, payload) {
  try {
    const cachePath = getFastCacheFilePath(tag, sourceFilePath);
    ensureDirSync(path.dirname(cachePath));
    fs.writeFileSync(cachePath, JSON.stringify({
      version: FAST_CACHE_VERSION,
      sourceFilePath: String(sourceFilePath || ''),
      sourceMtimeMs: Number(sourceMtimeMs || 0),
      builtAt: new Date().toISOString(),
      payload
    }));
  } catch (error) {
    // cache é best-effort: ignora falhas (permissão/antivírus/etc)
  }
}

function getFileMtimeMs(filePath) {
  try {
    return fs.statSync(filePath).mtimeMs || 0;
  } catch (error) {
    return 0;
  }
}

function readWorkbookCached(filePath) {
  const safePath = String(filePath || '').trim();
  if (!safePath) {
    throw new Error('Caminho da planilha obrigatorio.');
  }
  const mtimeMs = getFileMtimeMs(safePath);
  const cached = XLSX_WORKBOOK_CACHE.get(safePath) || null;
  if (cached && cached.workbook && cached.mtimeMs === mtimeMs) {
    return { workbook: cached.workbook, mtimeMs };
  }
  const workbook = XLSX.readFile(safePath, XLSX_READ_OPTIONS);
  XLSX_WORKBOOK_CACHE.set(safePath, { workbook, mtimeMs });
  return { workbook, mtimeMs };
}

function invalidateFileCaches(filePath) {
  const safePath = String(filePath || '').trim();
  if (!safePath) {
    return;
  }
  XLSX_WORKBOOK_CACHE.delete(safePath);

  if (FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.filePath === safePath) {
    FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE = { filePath: '', mtimeMs: -1, pedidoToReqs: null, reqBaseToPedidos: null };
  }
  if (FISCAL_BANCO_PEDIDOS_ROWS_CACHE.filePath === safePath) {
    FISCAL_BANCO_PEDIDOS_ROWS_CACHE = { filePath: '', mtimeMs: -1, result: null };
  }
  if (FISCAL_BANCO_PEDIDOS_LITE_CACHE.filePath === safePath) {
    FISCAL_BANCO_PEDIDOS_LITE_CACHE = { filePath: '', mtimeMs: -1, rows: null, looksLikeHeader: false };
  }
  if (RH_SNAPSHOT_CACHE.filePath === safePath) {
    RH_SNAPSHOT_CACHE = { filePath: '', mtimeMs: -1, snapshot: null };
  }
}

function writeWorkbookFile(workbook, filePath) {
  XLSX.writeFile(workbook, filePath);
  invalidateFileCaches(filePath);
}

function invalidateWorkbookCaches(workbook, sheetName = null) {
  if (!workbook || typeof workbook !== 'object') {
    return;
  }
  if (!sheetName) {
    WORKBOOK_AOA_CACHE.delete(workbook);
    WORKBOOK_OBJECTS_CACHE.delete(workbook);
    return;
  }
  const aoaMap = WORKBOOK_AOA_CACHE.get(workbook);
  if (aoaMap) {
    aoaMap.delete(sheetName);
  }
  const objMap = WORKBOOK_OBJECTS_CACHE.get(workbook);
  if (objMap) {
    objMap.delete(sheetName);
  }
}

function getWorksheetEffectiveRange(worksheet) {
  if (!worksheet) {
    return null;
  }
  const keys = Object.keys(worksheet).filter((key) => key && key[0] !== '!');
  if (!keys.length) {
    return null;
  }
  let minR = Number.POSITIVE_INFINITY;
  let minC = Number.POSITIVE_INFINITY;
  let maxR = -1;
  let maxC = -1;
  keys.forEach((address) => {
    const cell = worksheet[address];
    if (!cell || (cell.v === undefined && cell.w === undefined && cell.f === undefined)) {
      return;
    }
    const decoded = XLSX.utils.decode_cell(address);
    if (!decoded) {
      return;
    }
    minR = Math.min(minR, decoded.r);
    minC = Math.min(minC, decoded.c);
    maxR = Math.max(maxR, decoded.r);
    maxC = Math.max(maxC, decoded.c);
  });
  if (maxR < 0 || maxC < 0 || !Number.isFinite(minR) || !Number.isFinite(minC)) {
    return null;
  }
  return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
}

function sheetToAoaFast(worksheet) {
  if (!worksheet) {
    return [];
  }
  const effectiveRange = getWorksheetEffectiveRange(worksheet);
  const range = effectiveRange || XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  return XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '', range });
}

function readUsersFile() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const { workbook } = readWorkbookCached(bancoPedidosPath);
  const sheetName = 'USERS';
  if (!workbook.Sheets[sheetName]) {
    const ws = XLSX.utils.aoa_to_sheet([
      ['USERNAME', 'SALT', 'ITERATIONS', 'HASH', 'CAN_FISCAL', 'CAN_PCP', 'CREATED_AT', 'PERMISSIONS']
    ]);
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = ws;
  }

  const ws = workbook.Sheets[sheetName];
  const aoa = sheetToAoaFast(ws);
  const headerRow = Array.isArray(aoa[0]) ? aoa[0] : [];
  const normalizedHeaders = headerRow.map((value) => String(value || '').trim().toUpperCase());
  const idx = {
    username: normalizedHeaders.indexOf('USERNAME'),
    salt: normalizedHeaders.indexOf('SALT'),
    iterations: normalizedHeaders.indexOf('ITERATIONS'),
    hash: normalizedHeaders.indexOf('HASH'),
    canFiscal: normalizedHeaders.indexOf('CAN_FISCAL'),
    canPcp: normalizedHeaders.indexOf('CAN_PCP'),
    createdAt: normalizedHeaders.indexOf('CREATED_AT'),
    permissions: normalizedHeaders.indexOf('PERMISSIONS')
  };
  const users = [];

  for (let i = 1; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const username = normalizeUsername(idx.username >= 0 ? row[idx.username] : row[0]);
    if (!username) {
      continue;
    }

    const salt = String(idx.salt >= 0 ? row[idx.salt] : row[1] || '').trim();
    const iterations = Number(idx.iterations >= 0 ? row[idx.iterations] : row[2] || 0);
    const hash = String(idx.hash >= 0 ? row[idx.hash] : row[3] || '').trim();
    const canFiscal = idx.canFiscal >= 0 ? row[idx.canFiscal] : row[4];
    let canPcp = idx.canPcp >= 0 ? row[idx.canPcp] : undefined;
    let createdAt = String(idx.createdAt >= 0 ? row[idx.createdAt] : row[5] || '').trim();
    const permissionsCell = idx.permissions >= 0 ? row[idx.permissions] : '';
    if (idx.canPcp < 0 && !createdAt && row[5] !== undefined) {
      const legacyValue = String(row[5] || '').trim();
      if (legacyValue) {
        createdAt = legacyValue;
      }
    }
    if (idx.canPcp < 0) {
      canPcp = '1';
    }

    const jsonPermissions = parsePermissionsCell(permissionsCell);
    const permissions = normalizePermissions({
      ...jsonPermissions,
      fiscal: parseUserPermission(
        canFiscal,
        Object.prototype.hasOwnProperty.call(jsonPermissions, 'fiscal') ? !!jsonPermissions.fiscal : false
      ),
      pcp: parseUserPermission(
        canPcp,
        Object.prototype.hasOwnProperty.call(jsonPermissions, 'pcp') ? !!jsonPermissions.pcp : true
      )
    });

    users.push({
      username,
      passwordHash: { salt, iterations, hash },
      permissions,
      createdAt
    });
  }

  return { workbook, sheetName, bancoPedidosPath, users };
}

function writeUsersFile(payload) {
  const workbook = payload && payload.workbook ? payload.workbook : null;
  const bancoPedidosPath = payload && payload.bancoPedidosPath ? payload.bancoPedidosPath : null;
  const sheetName = payload && payload.sheetName ? payload.sheetName : 'USERS';
  const users = Array.isArray(payload && payload.users) ? payload.users : [];

  if (!workbook || !bancoPedidosPath) {
    throw new Error('Falha ao salvar usuarios (workbook/path ausentes).');
  }

  const aoa = [['USERNAME', 'SALT', 'ITERATIONS', 'HASH', 'CAN_FISCAL', 'CAN_PCP', 'CREATED_AT', 'PERMISSIONS']];
  users.forEach((user) => {
    const normalizedPermissions = normalizePermissions(user.permissions || {});
    aoa.push([
      normalizeUsername(user.username),
      user.passwordHash?.salt || '',
      user.passwordHash?.iterations || '',
      user.passwordHash?.hash || '',
      normalizedPermissions.fiscal ? '1' : '0',
      normalizedPermissions.pcp ? '1' : '0',
      user.createdAt || '',
      JSON.stringify(normalizedPermissions)
    ]);
  });

  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
  if (!workbook.SheetNames.includes(sheetName)) {
    workbook.SheetNames.push(sheetName);
  }

  writeWorkbookFile(workbook, bancoPedidosPath);
  return { users };
}

function normalizeUsername(value) {
  return String(value || '').trim().toUpperCase();
}

function normalizePermissionKey(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function getDefaultPermissionMap() {
  return {
    pcp: true,
    fiscal: false,
    rh: false,
    rh_edit: false,
    settings: true,
    log: true
  };
}

function normalizePermissions(raw, fallback = null) {
  const base = fallback || getDefaultPermissionMap();
  const result = { ...base };
  if (!raw || typeof raw !== 'object') {
    return result;
  }

  Object.entries(raw).forEach(([key, value]) => {
    const normalizedKey = normalizePermissionKey(key);
    if (!normalizedKey) {
      return;
    }
    result[normalizedKey] = !!value;
  });
  return result;
}

function parsePermissionsCell(value) {
  const text = String(value || '').trim();
  if (!text) {
    return {};
  }
  try {
    const parsed = JSON.parse(text);
    if (parsed && typeof parsed === 'object') {
      return normalizePermissions(parsed, {});
    }
    return {};
  } catch (error) {
    return {};
  }
}

function parseUserPermission(value, defaultValue) {
  if (value === undefined || value === null || String(value).trim() === '') {
    return !!defaultValue;
  }
  const text = String(value || '').trim().toUpperCase();
  return text === '1' || text === 'TRUE' || text === 'SIM';
}

function hashPassword(password, options = {}) {
  const safePassword = String(password || '');
  if (!safePassword) {
    throw new Error('Senha obrigatoria.');
  }

  const iterations = Number(options.iterations || 120000);
  const salt = options.salt || crypto.randomBytes(16).toString('hex');
  const hash = crypto
    .pbkdf2Sync(safePassword, salt, iterations, 32, 'sha256')
    .toString('hex');

  return { salt, iterations, hash };
}

function verifyPassword(password, stored) {
  if (!stored || !stored.salt || !stored.iterations || !stored.hash) {
    return false;
  }

  try {
    const computed = hashPassword(password, { salt: stored.salt, iterations: stored.iterations });
    return computed.hash === stored.hash;
  } catch (error) {
    return false;
  }
}

function ensureDefaultAdminUser() {
  try {
    const data = readUsersFile();
    const users = data.users || [];
    if (users.length) {
      return;
    }

    const username = 'ADM';
    const passwordHash = hashPassword('1234');
    users.push({
      username,
      passwordHash,
      permissions: normalizePermissions({ fiscal: true, pcp: true, settings: true, log: true }),
      createdAt: new Date().toISOString()
    });

    writeUsersFile({
      workbook: data.workbook,
      bancoPedidosPath: data.bancoPedidosPath,
      sheetName: data.sheetName,
      users
    });
  } catch (error) {
    // Caminho fiscal/banco ainda nao configurado, nao força criacao aqui.
  }
}

function listUsers() {
  const data = readUsersFile();
  return (data.users || []).map((user) => ({
    username: user.username,
    permissions: normalizePermissions(user.permissions || {}),
    createdAt: user.createdAt || ''
  }));
}

function createUser({ username, password, permissions }) {
  const data = readUsersFile();
  const users = data.users || [];
  const normalized = normalizeUsername(username);
  if (!normalized) {
    throw new Error('Usuario obrigatorio.');
  }

  if (users.some((u) => normalizeUsername(u.username) === normalized)) {
    throw new Error(`Usuario ja existe: ${normalized}`);
  }

  const passwordHash = hashPassword(password);
  const next = {
    username: normalized,
    passwordHash,
    permissions: normalizePermissions(permissions || {}),
    createdAt: new Date().toISOString()
  };
  users.push(next);
  writeUsersFile({
    workbook: data.workbook,
    bancoPedidosPath: data.bancoPedidosPath,
    sheetName: data.sheetName,
    users
  });
  return { username: next.username, permissions: next.permissions };
}

function updateUserPermissions({ username, permissions }) {
  const data = readUsersFile();
  const users = data.users || [];
  const normalized = normalizeUsername(username);
  const user = users.find((u) => normalizeUsername(u.username) === normalized);
  if (!user) {
    throw new Error(`Usuario nao encontrado: ${normalized}`);
  }

  user.permissions = normalizePermissions(permissions || {}, normalizePermissions(user.permissions || {}));
  writeUsersFile({
    workbook: data.workbook,
    bancoPedidosPath: data.bancoPedidosPath,
    sheetName: data.sheetName,
    users
  });
  return { username: user.username, permissions: user.permissions };
}

function setUserPassword({ username, password }) {
  const data = readUsersFile();
  const users = data.users || [];
  const normalized = normalizeUsername(username);
  const user = users.find((u) => normalizeUsername(u.username) === normalized);
  if (!user) {
    throw new Error(`Usuario nao encontrado: ${normalized}`);
  }

  user.passwordHash = hashPassword(password);
  writeUsersFile({
    workbook: data.workbook,
    bancoPedidosPath: data.bancoPedidosPath,
    sheetName: data.sheetName,
    users
  });
  return { username: user.username };
}

function deleteUser({ username }) {
  const data = readUsersFile();
  const users = data.users || [];
  const normalized = normalizeUsername(username);
  const filtered = users.filter((u) => normalizeUsername(u.username) !== normalized);
  if (filtered.length === users.length) {
    throw new Error(`Usuario nao encontrado: ${normalized}`);
  }
  writeUsersFile({
    workbook: data.workbook,
    bancoPedidosPath: data.bancoPedidosPath,
    sheetName: data.sheetName,
    users: filtered
  });
  return true;
}

function verifyLogin({ username, password, requireFiscal }) {
  let data;
  try {
    data = readUsersFile();
  } catch (error) {
    return { ok: false, error: error.message || 'Erro ao abrir base de usuarios.' };
  }
  const users = data.users || [];
  const normalized = normalizeUsername(username);
  const user = users.find((u) => normalizeUsername(u.username) === normalized);
  if (!user) {
    return { ok: false, error: 'Usuario ou senha incorretos.' };
  }

  const okPass = verifyPassword(password, user.passwordHash);
  if (!okPass) {
    return { ok: false, error: 'Usuario ou senha incorretos.' };
  }

  const perms = normalizePermissions(user.permissions || {});
  if (requireFiscal && !perms.fiscal) {
    return { ok: false, error: 'Usuario sem permissao para acessar o FISCAL.' };
  }

  return { ok: true, username: user.username, permissions: perms };
}

const STANDARD_COLUMNS = [
  'LINHA',
  'QTD',
  'DESCRIÇÃO',
  'N PC',
  'N REQ',
  'CLIENTE',
  'PROGRAMA',
  'BUCHA',
  'ENTREGA',
  'OBS'
];

const FAGOR_1_COLUMNS = [
  ...STANDARD_COLUMNS,
  'MAQ./OPER.'
];

const MOLDAGEM_COLUMNS = [
  'CÓD',
  'QTD',
  'DESCRIÇÃO ITEM',
  'PC',
  'REQ',
  'CLIENTE',
  'MOLDES',
  'EIXO',
  'BUCHA PRENSADAS',
  'PESO UTILIZADO',
  'TOTAL BUCHAS',
  'PESO TOTAL',
  'DT DE ENTREGA',
  'STATUS',
  'Nº PRENSA',
  'OPERADOR'
];

const MESTRA_COLUMNS = [
  'LINHA',
  'QTD',
  'DESCRIÇÃO',
  'N PC',
  'N REQ',
  'CLIENTE',
  'BAIXA PROJETO',
  'BAIXA MAQUINA',
  'BAIXA ACABAMENTO',
  'OBS'
];

const LIST_TYPES = {
  acabamento: {
    label: 'Acabamento',
    relativePath: ['LISTAS DIÁRIAS', 'ACABAMENTO'],
    filePrefix: 'Lista de Acabamento',
    columns: STANDARD_COLUMNS
  },
  projeto: {
    label: 'Projeto',
    relativePath: ['LISTAS DIÁRIAS', 'PROJETO'],
    filePrefix: 'Lista de Projeto',
    columns: STANDARD_COLUMNS
  },
  fagor1: {
    label: 'FAGOR 1',
    relativePath: ['LISTAS DIÁRIAS', 'MÁQUINAS', 'FAGOR 1'],
    filePrefix: 'LISTA MAQUINAS FAGOR 1',
    columns: FAGOR_1_COLUMNS,
    machineValue: 'FAGOR 1'
  },
  fagor2: {
    label: 'FAGOR 2',
    relativePath: ['LISTAS DIÁRIAS', 'MÁQUINAS', 'FAGOR 2'],
    filePrefix: 'LISTA MAQUINAS FAGOR 2',
    columns: FAGOR_1_COLUMNS,
    machineValue: 'FAGOR 2'
  },
  mcs1: {
    label: 'MCS 1',
    relativePath: ['LISTAS DIÁRIAS', 'MÁQUINAS', 'MCS 1'],
    filePrefix: 'LISTA MAQUINAS MCS 1',
    columns: FAGOR_1_COLUMNS,
    machineValue: 'MCS 1'
  },
  mcs2: {
    label: 'MCS 2',
    relativePath: ['LISTAS DIÁRIAS', 'MÁQUINAS', 'MCS 2'],
    filePrefix: 'LISTA MAQUINAS MCS 2',
    columns: FAGOR_1_COLUMNS,
    machineValue: 'MCS 2'
  },
  mcs3: {
    label: 'MCS 3',
    relativePath: ['LISTAS DIÁRIAS', 'MÁQUINAS', 'MCS 3'],
    filePrefix: 'LISTA MAQUINAS MCS 3',
    columns: FAGOR_1_COLUMNS,
    machineValue: 'MCS 3'
  },
  moldagem: {
    label: 'Moldagem',
    relativePath: ['LISTAS DIÁRIAS', 'MOLDAGEM'],
    filePrefix: 'Lista de Moldagem',
    columns: MOLDAGEM_COLUMNS,
    preferredSheetName: 'Lista'
  },
  mestra: {
    label: 'MESTRA',
    columns: MESTRA_COLUMNS
  }
};

const READY_OBS_VALUES = new Set(['PRONTO', 'PALOMA', 'DARCI', 'REBECA', 'GUSTAVO', 'GUSTAVO S']);
const PROJECT_READY_QTD_COLORS = new Set(['00B050', '92D050']);
const PROJECT_NEW_PROGRAM_VALUES = new Set(['NAO TEM', 'SE']);
const ACABAMENTO_OPERATORS = ['GUSTAVO', 'DARCI', 'REBECA', 'PALOMA'];
const MOLDAGEM_STATUS_ESTOQUE = 'BUCHA ESTOQUE';
const MOLDAGEM_STATUS_PRENSADO = 'PRENSADO';
const MACHINE_LIST_TYPES = ['fagor1', 'fagor2', 'mcs1', 'mcs2', 'mcs3'];

function getCell(worksheet, row, col) {
  const address = XLSX.utils.encode_cell({ c: col, r: row });
  return worksheet[address] || null;
}

function setCell(worksheet, row, col, value) {
  const address = XLSX.utils.encode_cell({ c: col, r: row });
  const stringValue = value == null ? '' : String(value);
  worksheet[address] = { t: 's', v: stringValue };
}

function normalizeText(value) {
  return (value || '')
    .toString()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9\s\-\/]/gi, '')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function parseHeHoursFromCell(cell) {
  if (!cell) {
    return 0;
  }

  const raw = cell.v;
  if (raw == null) {
    const text = String(cell.w || '').trim();
    if (!text) {
      return 0;
    }
    return parseHeHoursFromCell({ v: text });
  }

  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    // Se vier como Date, usa hora/minuto como duração.
    return raw.getHours() + (raw.getMinutes() / 60) + (raw.getSeconds() / 3600);
  }

  if (typeof raw === 'number') {
    // Excel pode guardar duração como fração do dia (ex.: 0.0625 = 1h30m)
    const z = String(cell.z || '').toLowerCase();
    const w = String(cell.w || '').trim();
    const looksLikeTime = z.includes('h') || z.includes('m') || z.includes(':') || w.includes(':');
    if (raw >= 0 && raw <= 1 && looksLikeTime) {
      return raw * 24;
    }
    return raw;
  }

  const text = String(raw || '').trim();
  if (!text) {
    return 0;
  }
  const hhmm = text.match(/^(-?\d{1,3})\s*:\s*(\d{1,2})/);
  if (hhmm) {
    const h = Number(hhmm[1]);
    const m = Number(hhmm[2]);
    if (Number.isFinite(h) && Number.isFinite(m)) {
      return h + (m / 60);
    }
  }
  const cleaned = text.replace(/\s*h/i, '').replace(/\s*m/i, '').replace(',', '.');
  const num = Number(cleaned);
  if (Number.isFinite(num)) {
    return num;
  }
  return 0;
}

function parseRhAtrasoHoursFromCell(cell) {
  if (!cell) {
    return 0;
  }

  const raw = cell.v;
  if (raw == null) {
    const text = String(cell.w || '').trim();
    if (!text) {
      return 0;
    }
    return parseRhAtrasoHoursFromCell({ v: text });
  }

  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return raw.getHours() + (raw.getMinutes() / 60) + (raw.getSeconds() / 3600);
  }

  if (typeof raw === 'number') {
    const z = String(cell.z || '').toLowerCase();
    const w = String(cell.w || '').trim();
    const looksLikeTime = z.includes('h') || z.includes('m') || z.includes(':') || w.includes(':');
    if (raw >= 0 && raw <= 1 && looksLikeTime) {
      return raw * 24;
    }
    // Atraso geralmente vem em minutos quando é um número "solto" (ex.: 15, 30, 90).
    if (raw > 8 && raw <= 720 && !looksLikeTime) {
      return raw / 60;
    }
    return raw;
  }

  const text = String(raw || '').trim();
  if (!text) {
    return 0;
  }

  const minMatch = text.match(/^(-?\d+(?:[.,]\d+)?)\s*(min|mins|m)\b/i);
  if (minMatch) {
    const mins = Number(String(minMatch[1]).replace(',', '.'));
    if (Number.isFinite(mins)) {
      return mins / 60;
    }
  }
  const hMatch = text.match(/^(-?\d+(?:[.,]\d+)?)\s*(h|hora|horas)\b/i);
  if (hMatch) {
    const hours = Number(String(hMatch[1]).replace(',', '.'));
    if (Number.isFinite(hours)) {
      return hours;
    }
  }

  // Fallback para hh:mm, fração do dia e números simples.
  return parseHeHoursFromCell({ v: text });
}

function tokenizePersonName(value) {
  const normalized = normalizeText(String(value || '').replace(/['`]/g, ' '));
  const stop = new Set(['DE', 'DA', 'DO', 'DAS', 'DOS', 'E']);
  return normalized
    .split(' ')
    .map((t) => String(t || '').trim())
    .filter((t) => t.length >= 2 && !stop.has(t));
}

function matchColaboradorBySheetName(colaboradores, sheetName) {
  const list = Array.isArray(colaboradores) ? colaboradores : [];
  const sheet = String(sheetName || '').trim();
  if (!sheet) {
    return { match: null, score: 0 };
  }
  const sheetNorm = normalizeText(sheet.replace(/['`]/g, ' '));
  const sheetTokens = tokenizePersonName(sheetNorm);
  if (!sheetTokens.length) {
    return { match: null, score: 0 };
  }

  let best = null;
  let bestScore = 0;
  list.forEach((c) => {
    const name = String(c.nome || '').trim();
    if (!name) {
      return;
    }
    const collabNorm = normalizeText(name.replace(/['`]/g, ' '));
    const collabTokens = tokenizePersonName(collabNorm);
    if (!collabTokens.length) {
      return;
    }

    let hits = 0;
    collabTokens.forEach((t) => {
      if (sheetTokens.includes(t)) {
        hits += 1;
      }
    });

    const base = hits / collabTokens.length;
    const bonus = (sheetNorm.includes(collabNorm) || collabNorm.includes(sheetNorm)) ? 0.35 : 0;
    const score = base + bonus;
    if (score > bestScore) {
      best = c;
      bestScore = score;
    }
  });

  // exige pelo menos 2 tokens batendo ou substring forte
  if (!best) {
    return { match: null, score: 0 };
  }
  const bestTokens = tokenizePersonName(best.nome || '');
  const hits = bestTokens.filter((t) => sheetTokens.includes(t)).length;
  const strongSubstring = sheetNorm.includes(normalizeText(best.nome || ''));
  if (hits >= 2 || (strongSubstring && bestScore >= 0.6)) {
    return { match: best, score: bestScore };
  }
  if (bestScore >= 0.9) {
    return { match: best, score: bestScore };
  }
  return { match: null, score: bestScore };
}

const FISCAL_BANK_PEDIDOS_RELATIVE = ['Arquivos KuruJoos', 'Financeiro', 'banco Pedidos.xlsx'];
const FISCAL_BANCO_ORDENS_RELATIVE = ['PartsSeals', 'Banco de Ordens', 'BancoDeOrdens.xlsm'];
const FISCAL_RNC_RELATIVE = ['PartsSeals', 'Rnc Interna', 'Relatório de Não Conformidade BANCO DE DADOS.xlsm'];
const PCP_DASHBOARD_DAILY_RELATIVE = ['PartsSeals', 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'];
const PCP_DASHBOARD_DAILY_RELATIVE_LEGACY = ['Rede', 'PartsSeals', 'Programação de Produção Diaria'];
const RH_HE_DB_RELATIVE = ['Arquivos KuruJoos', 'RH', 'GestaoHorasExtras.xlsx'];

function ensureDirSync(folderPath) {
  if (!folderPath) {
    return;
  }
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath, { recursive: true });
  }
}

function generateRhId() {
  return Date.now() * 1000 + Math.floor(Math.random() * 1000);
}

function readSheetAoa(workbook, sheetName) {
  const ws = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
  if (!ws) {
    return [];
  }
  const safeSheet = String(sheetName || '').trim();
  if (!safeSheet) {
    return [];
  }
  let aoaMap = WORKBOOK_AOA_CACHE.get(workbook);
  if (!aoaMap) {
    aoaMap = new Map();
    WORKBOOK_AOA_CACHE.set(workbook, aoaMap);
  }
  if (aoaMap.has(safeSheet)) {
    return aoaMap.get(safeSheet);
  }
  const aoa = sheetToAoaFast(ws);
  aoaMap.set(safeSheet, aoa);
  return aoa;
}

function writeSheetAoa(workbook, sheetName, aoa) {
  const next = Array.isArray(aoa) ? aoa : [];
  if (!workbook.Sheets[sheetName]) {
    workbook.SheetNames.push(sheetName);
  }
  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(next);
  invalidateWorkbookCaches(workbook, sheetName);
}

function normalizeRhRole(value) {
  const text = String(value || '').trim().toUpperCase();
  if (text === 'RH') {
    return 'RH';
  }
  return 'GESTOR';
}

function resolveRhDbPath() {
  return resolveFiscalPath(RH_HE_DB_RELATIVE);
}

function createRhWorkbookIfMissing(filePath) {
  ensureDirSync(path.dirname(filePath));
  const wb = XLSX.utils.book_new();

  writeSheetAoa(wb, 'SETORES', [[
    'ID',
    'NOME_SETOR',
    'GESTOR_RESPONSAVEL',
    'CENTRO_CUSTO',
    'CREATED_AT',
    'UPDATED_AT',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'COLABORADORES', [[
    'ID',
    'NOME',
    'MATRICULA',
    'SETOR_ID',
    'CARGO',
    'STATUS',
    'DATA_ADMISSAO',
    'LIMITE_MENSAL',
    'CREATED_AT',
    'UPDATED_AT',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'HORAS_EXTRAS', [[
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'TIPO_HORA',
    'OBSERVACAO',
    'JUSTIFICATIVA',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'ATRASOS', [[
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'OBSERVACAO',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'SOLICITACOES_HE', [[
    'ID',
    'NUMERO_DOCUMENTO',
    'DATA_SOLICITACAO',
    'DATA_HE',
    'COLABORADOR_ID',
    'NOME_COLABORADOR',
    'SETOR',
    'FINALIDADE',
    'HORA_INICIO',
    'HORA_FIM',
    'TOTAL_HORAS',
    'SOLICITANTE',
    'DATA_CRIACAO',
    'ULTIMA_ATUALIZACAO',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'USERS_RH', [[
    'USERNAME',
    'ROLE',
    'SETOR_ID',
    'UPDATED_AT',
    'DELETED_AT'
  ]]);

  writeSheetAoa(wb, 'CONFIG', [
    ['KEY', 'VALUE'],
    ['justificativa_acima_horas', '2'],
    ['limite_mensal_padrao', '40'],
    ['margem_faixa_1_ate_horas', '25'],
    ['margem_faixa_1_percent', '50'],
    ['margem_faixa_2_ate_horas', '40'],
    ['margem_faixa_2_percent', '60'],
    ['margem_faixa_3_ate_horas', '60'],
    ['margem_faixa_3_percent', '80'],
    ['margem_faixa_4_percent', '100']
  ]);

  writeSheetAoa(wb, 'AUDIT_LOG', [[
    'ID',
    'DATA_HORA',
    'USERNAME',
    'ACTION',
    'ENTITY',
    'ENTITY_ID',
    'BEFORE_JSON',
    'AFTER_JSON'
  ]]);

  writeSheetAoa(wb, 'ACCESS_REQUESTS', [[
    'ID',
    'USERNAME',
    'REQUESTED_ROLE',
    'REQUESTED_SETOR_ID',
    'REQUESTED_AT',
    'STATUS',
    'DECIDED_AT',
    'DECIDED_BY'
  ]]);

  writeWorkbookFile(wb, filePath);
  return wb;
}

function ensureRhSheets(workbook) {
  const required = {
    SETORES: [
      'ID',
      'NOME_SETOR',
      'GESTOR_RESPONSAVEL',
      'CENTRO_CUSTO',
      'CREATED_AT',
      'UPDATED_AT',
      'DELETED_AT'
    ],
    COLABORADORES: [
      'ID',
      'NOME',
      'MATRICULA',
      'SETOR_ID',
      'CARGO',
      'STATUS',
      'DATA_ADMISSAO',
      'LIMITE_MENSAL',
      'CREATED_AT',
      'UPDATED_AT',
      'DELETED_AT'
    ],
    HORAS_EXTRAS: [
      'ID',
      'COLABORADOR_ID',
      'DATA',
      'QUANTIDADE_HORAS',
      'TIPO_HORA',
      'OBSERVACAO',
      'JUSTIFICATIVA',
      'CRIADO_POR',
      'DATA_REGISTRO',
      'UPDATED_AT',
      'DELETED_AT'
    ],
    ATRASOS: [
      'ID',
      'COLABORADOR_ID',
      'DATA',
      'QUANTIDADE_HORAS',
      'OBSERVACAO',
      'CRIADO_POR',
      'DATA_REGISTRO',
      'UPDATED_AT',
      'DELETED_AT'
    ],
    SOLICITACOES_HE: [
      'ID',
      'NUMERO_DOCUMENTO',
      'DATA_SOLICITACAO',
      'DATA_HE',
      'COLABORADOR_ID',
      'NOME_COLABORADOR',
      'SETOR',
      'FINALIDADE',
      'HORA_INICIO',
      'HORA_FIM',
      'TOTAL_HORAS',
      'SOLICITANTE',
      'DATA_CRIACAO',
      'ULTIMA_ATUALIZACAO',
      'DELETED_AT'
    ],
    USERS_RH: ['USERNAME', 'ROLE', 'SETOR_ID', 'UPDATED_AT', 'DELETED_AT'],
    CONFIG: ['KEY', 'VALUE'],
    AUDIT_LOG: [
      'ID',
      'DATA_HORA',
      'USERNAME',
      'ACTION',
      'ENTITY',
      'ENTITY_ID',
      'BEFORE_JSON',
      'AFTER_JSON'
    ],
    ACCESS_REQUESTS: [
      'ID',
      'USERNAME',
      'REQUESTED_ROLE',
      'REQUESTED_SETOR_ID',
      'REQUESTED_AT',
      'STATUS',
      'DECIDED_AT',
      'DECIDED_BY'
    ]
  };

  const ensureSheetColumns = (sheetName, headers) => {
    if (!workbook.Sheets[sheetName]) {
      writeSheetAoa(workbook, sheetName, [headers]);
      return;
    }
    const aoa = readSheetAoa(workbook, sheetName);
    const existingHeader = Array.isArray(aoa[0]) ? aoa[0].map((h) => String(h || '').trim().toUpperCase()) : [];
    const normalizedRequired = headers.map((h) => String(h || '').trim().toUpperCase()).filter(Boolean);
    const missing = normalizedRequired.filter((h) => !existingHeader.includes(h));
    if (!missing.length) {
      return;
    }
    const nextHeader = [...existingHeader, ...missing];
    const next = [nextHeader];
    for (let i = 1; i < aoa.length; i += 1) {
      const row = Array.isArray(aoa[i]) ? [...aoa[i]] : [];
      while (row.length < nextHeader.length) {
        row.push('');
      }
      next.push(row);
    }
    writeSheetAoa(workbook, sheetName, next);
  };

  Object.entries(required).forEach(([sheetName, headers]) => ensureSheetColumns(sheetName, headers));
}

function readRhDb() {
  const filePath = resolveRhDbPath();
  if (!fs.existsSync(filePath)) {
    const workbook = createRhWorkbookIfMissing(filePath);
    ensureRhSheets(workbook);
    RH_ENSURED_WORKBOOKS.add(workbook);
    const mtimeMs = getFileMtimeMs(filePath);
    XLSX_WORKBOOK_CACHE.set(filePath, { workbook, mtimeMs });
    return { workbook, filePath };
  }
  const { workbook } = readWorkbookCached(filePath);
  if (!RH_ENSURED_WORKBOOKS.has(workbook)) {
    ensureRhSheets(workbook);
    RH_ENSURED_WORKBOOKS.add(workbook);
  }
  return { workbook, filePath };
}

function buildRhSnapshotFromWorkbook(workbook) {
  const config = getRhConfig(workbook);
  const setores = listRhSetores({ allowedSectorId: '' }, workbook);
  const colaboradores = listRhColaboradores({ allowedSectorId: '' }, workbook);

  const rawHe = readSheetObjects(workbook, 'HORAS_EXTRAS');
  const horasByMonth = {};
  rawHe.forEach((row) => {
    if (String(row.DELETED_AT || '').trim()) {
      return;
    }
    const id = String(row.ID || '').trim();
    const colaborador_id = String(row.COLABORADOR_ID || '').trim();
    const data = String(row.DATA || '').trim();
    if (!id || !colaborador_id || !data) {
      return;
    }
    const month = data.length >= 7 ? data.slice(0, 7) : '';
    if (!/^\d{4}-\d{2}$/.test(month)) {
      return;
    }
    if (!horasByMonth[month]) {
      horasByMonth[month] = [];
    }
    horasByMonth[month].push({
      id,
      colaborador_id,
      data,
      quantidade_horas: Number(row.QUANTIDADE_HORAS || 0),
      tipo_hora: String(row.TIPO_HORA || '').trim(),
      observacao: String(row.OBSERVACAO || '').trim(),
      justificativa: String(row.JUSTIFICATIVA || '').trim(),
      criado_por: String(row.CRIADO_POR || '').trim(),
      data_registro: String(row.DATA_REGISTRO || '').trim()
    });
  });

  const rawAtrasos = readSheetObjects(workbook, 'ATRASOS');
  const atrasosByMonth = {};
  rawAtrasos.forEach((row) => {
    if (String(row.DELETED_AT || '').trim()) {
      return;
    }
    const id = String(row.ID || '').trim();
    const colaborador_id = String(row.COLABORADOR_ID || '').trim();
    const data = String(row.DATA || '').trim();
    if (!id || !colaborador_id || !data) {
      return;
    }
    const month = data.length >= 7 ? data.slice(0, 7) : '';
    if (!/^\d{4}-\d{2}$/.test(month)) {
      return;
    }
    if (!atrasosByMonth[month]) {
      atrasosByMonth[month] = [];
    }
    atrasosByMonth[month].push({
      id,
      colaborador_id,
      data,
      quantidade_horas: Number(row.QUANTIDADE_HORAS || 0),
      observacao: String(row.OBSERVACAO || '').trim(),
      criado_por: String(row.CRIADO_POR || '').trim(),
      data_registro: String(row.DATA_REGISTRO || '').trim()
    });
  });

  return { config, setores, colaboradores, horasByMonth, atrasosByMonth };
}

function readRhSnapshot() {
  const filePath = resolveRhDbPath();
  if (!fs.existsSync(filePath)) {
    const workbook = createRhWorkbookIfMissing(filePath);
    ensureRhSheets(workbook);
    RH_ENSURED_WORKBOOKS.add(workbook);
    const mtimeMs = getFileMtimeMs(filePath);
    XLSX_WORKBOOK_CACHE.set(filePath, { workbook, mtimeMs });
    const snapshot = buildRhSnapshotFromWorkbook(workbook);
    RH_SNAPSHOT_CACHE = { filePath, mtimeMs, snapshot };
    writeFastCache('rh-snapshot', filePath, mtimeMs, snapshot);
    return { filePath, mtimeMs, snapshot };
  }

  const mtimeMs = getFileMtimeMs(filePath);
  if (
    RH_SNAPSHOT_CACHE.snapshot
    && RH_SNAPSHOT_CACHE.filePath === filePath
    && RH_SNAPSHOT_CACHE.mtimeMs === mtimeMs
  ) {
    return { filePath, mtimeMs, snapshot: RH_SNAPSHOT_CACHE.snapshot };
  }

  const diskCache = readFastCache('rh-snapshot', filePath, mtimeMs);
  if (
    diskCache
    && typeof diskCache === 'object'
    && Object.prototype.hasOwnProperty.call(diskCache, 'atrasosByMonth')
  ) {
    RH_SNAPSHOT_CACHE = { filePath, mtimeMs, snapshot: diskCache };
    return { filePath, mtimeMs, snapshot: diskCache };
  }

  const { workbook } = readWorkbookCached(filePath);
  if (!RH_ENSURED_WORKBOOKS.has(workbook)) {
    ensureRhSheets(workbook);
    RH_ENSURED_WORKBOOKS.add(workbook);
  }
  const snapshot = buildRhSnapshotFromWorkbook(workbook);
  RH_SNAPSHOT_CACHE = { filePath, mtimeMs, snapshot };
  writeFastCache('rh-snapshot', filePath, mtimeMs, snapshot);
  return { filePath, mtimeMs, snapshot };
}

function mapAoaToObjects(aoa) {
  const rows = Array.isArray(aoa) ? aoa : [];
  const header = Array.isArray(rows[0]) ? rows[0] : [];
  const headers = header.map((h) => String(h || '').trim().toUpperCase());
  const result = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const obj = {};
    headers.forEach((key, idx) => {
      if (!key) {
        return;
      }
      obj[key] = row[idx] !== undefined ? row[idx] : '';
    });
    result.push(obj);
  }
  return result;
}

function readSheetObjects(workbook, sheetName) {
  const safeSheet = String(sheetName || '').trim();
  if (!safeSheet) {
    return [];
  }
  let objMap = WORKBOOK_OBJECTS_CACHE.get(workbook);
  if (!objMap) {
    objMap = new Map();
    WORKBOOK_OBJECTS_CACHE.set(workbook, objMap);
  }
  if (objMap.has(safeSheet)) {
    return objMap.get(safeSheet);
  }
  const aoa = readSheetAoa(workbook, safeSheet);
  const objs = mapAoaToObjects(aoa);
  objMap.set(safeSheet, objs);
  return objs;
}

function pickConfigMap(workbook) {
  const objs = readSheetObjects(workbook, 'CONFIG');
  const result = {};
  objs.forEach((row) => {
    const key = String(row.KEY || '').trim();
    if (!key) {
      return;
    }
    result[key] = String(row.VALUE ?? '').trim();
  });
  return result;
}

function getRhConfig(workbook) {
  const raw = pickConfigMap(workbook);
  const justificativa_acima_horas = Number(raw.justificativa_acima_horas || 2);
  const limite_mensal_padrao = Number(raw.limite_mensal_padrao || 40);
  const faixa1Ate = Number(raw.margem_faixa_1_ate_horas || 25);
  const faixa1Pct = Number(raw.margem_faixa_1_percent || 50);
  const faixa2Ate = Number(raw.margem_faixa_2_ate_horas || 40);
  const faixa2Pct = Number(raw.margem_faixa_2_percent || 60);
  const faixa3Ate = Number(raw.margem_faixa_3_ate_horas || 60);
  const faixa3Pct = Number(raw.margem_faixa_3_percent || 80);
  const faixa4Pct = Number(raw.margem_faixa_4_percent || 100);
  const safeFaixa1Ate = Number.isFinite(faixa1Ate) ? faixa1Ate : 25;
  const safeFaixa1Pct = Number.isFinite(faixa1Pct) ? faixa1Pct : 50;
  const safeFaixa2Ate = Number.isFinite(faixa2Ate) ? faixa2Ate : 40;
  const safeFaixa2Pct = Number.isFinite(faixa2Pct) ? faixa2Pct : 60;
  const safeFaixa3Ate = Number.isFinite(faixa3Ate) ? faixa3Ate : 60;
  const safeFaixa3Pct = Number.isFinite(faixa3Pct) ? faixa3Pct : 80;
  const safeFaixa4Pct = Number.isFinite(faixa4Pct) ? faixa4Pct : 100;

  const parseBool = (value, fallback = true) => {
    const text = String(value ?? '').trim().toLowerCase();
    if (!text) {
      return fallback;
    }
    if (['0', 'false', 'nao', 'não', 'n'].includes(text)) {
      return false;
    }
    if (['1', 'true', 'sim', 's', 'yes', 'y'].includes(text)) {
      return true;
    }
    return fallback;
  };

  const parseIsoDateList = (value) => {
    const text = String(value ?? '').trim();
    if (!text) {
      return [];
    }
    return text
      .split(/[\n,;|]+/g)
      .map((x) => String(x || '').trim())
      .filter((x) => /^\d{4}-\d{2}-\d{2}$/.test(x));
  };

  const descontar_feriados = parseBool(raw.descontar_feriados, true);
  const municipio_uf = String(raw.municipio_uf || '').trim() || `SANTA BARBARA D'OESTE-SP`;
  const feriados_extras = [
    ...parseIsoDateList(raw.feriados_extras),
    ...parseIsoDateList(raw.feriados_municipais)
  ];
  return {
    justificativa_acima_horas: Number.isFinite(justificativa_acima_horas) ? justificativa_acima_horas : 2,
    limite_mensal_padrao: Number.isFinite(limite_mensal_padrao) ? limite_mensal_padrao : 40,
    descontar_feriados,
    municipio_uf,
    feriados_extras,
    margem_faixa_1_ate_horas: safeFaixa1Ate,
    margem_faixa_1_percent: safeFaixa1Pct,
    margem_faixa_2_ate_horas: safeFaixa2Ate,
    margem_faixa_2_percent: safeFaixa2Pct,
    margem_faixa_3_ate_horas: safeFaixa3Ate,
    margem_faixa_3_percent: safeFaixa3Pct,
    margem_faixa_4_percent: safeFaixa4Pct,
    margem_faixas: [
      {
        ate_horas: safeFaixa1Ate,
        percent: safeFaixa1Pct
      },
      {
        ate_horas: safeFaixa2Ate,
        percent: safeFaixa2Pct
      },
      {
        ate_horas: safeFaixa3Ate,
        percent: safeFaixa3Pct
      },
      {
        ate_horas: null,
        percent: safeFaixa4Pct
      }
    ]
  };
}

function getRhMarginForHours(config, totalHours) {
  const safeHours = Number(totalHours || 0);
  const tiers = (config && Array.isArray(config.margem_faixas)) ? config.margem_faixas : [];
  for (let i = 0; i < tiers.length; i += 1) {
    const tier = tiers[i] || {};
    const ate = tier.ate_horas;
    if (ate == null) {
      return { percent: Number(tier.percent || 0), faixa: `ACIMA` };
    }
    if (safeHours <= Number(ate)) {
      return { percent: Number(tier.percent || 0), faixa: `ATÉ ${Number(ate)}h` };
    }
  }
  return { percent: 0, faixa: '' };
}

function setRhConfig({ username, config }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const next = config && typeof config === 'object' ? config : {};
  const current = getRhConfig(workbook);
  const aoa = [
    ['KEY', 'VALUE'],
    ['justificativa_acima_horas', String(Number(next.justificativa_acima_horas ?? current.justificativa_acima_horas))],
    ['limite_mensal_padrao', String(Number(next.limite_mensal_padrao ?? current.limite_mensal_padrao))],
    ['margem_faixa_1_ate_horas', String(Number(next.margem_faixa_1_ate_horas ?? (current.margem_faixas?.[0]?.ate_horas ?? 25)))],
    ['margem_faixa_1_percent', String(Number(next.margem_faixa_1_percent ?? (current.margem_faixas?.[0]?.percent ?? 50)))],
    ['margem_faixa_2_ate_horas', String(Number(next.margem_faixa_2_ate_horas ?? (current.margem_faixas?.[1]?.ate_horas ?? 40)))],
    ['margem_faixa_2_percent', String(Number(next.margem_faixa_2_percent ?? (current.margem_faixas?.[1]?.percent ?? 60)))],
    ['margem_faixa_3_ate_horas', String(Number(next.margem_faixa_3_ate_horas ?? (current.margem_faixas?.[2]?.ate_horas ?? 60)))],
    ['margem_faixa_3_percent', String(Number(next.margem_faixa_3_percent ?? (current.margem_faixas?.[2]?.percent ?? 80)))],
    ['margem_faixa_4_percent', String(Number(next.margem_faixa_4_percent ?? (current.margem_faixas?.[3]?.percent ?? 100)))]
  ];
  writeSheetAoa(workbook, 'CONFIG', aoa);
  appendRhAudit(workbook, {
    username: ctx.username,
    action: 'CONFIG_SET',
    entity: 'CONFIG',
    entityId: '',
    before: current,
    after: getRhConfig(workbook)
  });
  writeWorkbookFile(workbook, filePath);
  return { config: getRhConfig(workbook) };
}

function listRhUsers(workbook) {
  const objs = readSheetObjects(workbook, 'USERS_RH');
  return objs
    .filter((row) => !String(row.DELETED_AT || '').trim())
    .map((row) => ({
      username: normalizeUsername(row.USERNAME),
      role: normalizeRhRole(row.ROLE),
      setor_id: String(row.SETOR_ID || '').trim() || ''
    }))
    .filter((row) => row.username);
}

function getRhUserRule(workbook, username) {
  const safe = normalizeUsername(username);
  if (!safe) {
    return null;
  }
  const users = listRhUsers(workbook);
  return users.find((u) => normalizeUsername(u.username) === safe) || null;
}

function ensureRhAccess(username, { requireEdit = false } = {}) {
  const safeUser = normalizeUsername(username);
  if (!safeUser) {
    throw new Error('Usuário obrigatório.');
  }
  const user = listUsers().find((u) => normalizeUsername(u.username) === safeUser) || null;
  if (!user) {
    throw new Error('Usuário não encontrado.');
  }
  const perms = normalizePermissions(user.permissions || {});
  if (!perms.rh) {
    throw new Error('Usuário sem permissão para acessar RH.');
  }
  if (requireEdit && !perms.rh_edit) {
    throw new Error('Usuário sem permissão de edição no RH.');
  }

  const canEdit = !!perms.rh_edit;
  const { snapshot, filePath } = readRhSnapshot();
  return {
    username: safeUser,
    permissions: perms,
    canEdit,
    role: canEdit ? 'RH' : 'VISUALIZAR',
    allowedSectorId: '',
    dbFilePath: filePath,
    config: snapshot && snapshot.config ? snapshot.config : {}
  };
}

function listRhSetores(ctx, workbook) {
  const objs = readSheetObjects(workbook, 'SETORES');
  return objs
    .filter((row) => !String(row.DELETED_AT || '').trim())
    .map((row) => ({
      id: String(row.ID || '').trim(),
      nome_setor: String(row.NOME_SETOR || '').trim(),
      gestor_responsavel: String(row.GESTOR_RESPONSAVEL || '').trim(),
      centro_custo: String(row.CENTRO_CUSTO || '').trim()
    }))
    .filter((row) => row.id && row.nome_setor)
    .filter((row) => !ctx.allowedSectorId || String(row.id) === String(ctx.allowedSectorId))
    .sort((a, b) => String(a.nome_setor).localeCompare(String(b.nome_setor), 'pt-BR'));
}

function listRhColaboradores(ctx, workbook) {
  const objs = readSheetObjects(workbook, 'COLABORADORES');
  return objs
    .filter((row) => !String(row.DELETED_AT || '').trim())
    .map((row) => ({
      id: String(row.ID || '').trim(),
      nome: String(row.NOME || '').trim(),
      matricula: String(row.MATRICULA || '').trim(),
      setor_id: String(row.SETOR_ID || '').trim(),
      cargo: String(row.CARGO || '').trim(),
      status: String(row.STATUS || '').trim() || 'ativo',
      data_admissao: String(row.DATA_ADMISSAO || '').trim(),
      limite_mensal: String(row.LIMITE_MENSAL || '').trim()
    }))
    .filter((row) => row.id && row.nome)
    .filter((row) => !ctx.allowedSectorId || String(row.setor_id) === String(ctx.allowedSectorId))
    .sort((a, b) => String(a.nome).localeCompare(String(b.nome), 'pt-BR'));
}

function listRhSetoresFromSnapshot(ctx, snapshot) {
  const setores = snapshot && Array.isArray(snapshot.setores) ? snapshot.setores : [];
  return setores
    .filter((row) => row && row.id && row.nome_setor)
    .filter((row) => !ctx.allowedSectorId || String(row.id) === String(ctx.allowedSectorId));
}

function listRhColaboradoresFromSnapshot(ctx, snapshot) {
  const colaboradores = snapshot && Array.isArray(snapshot.colaboradores) ? snapshot.colaboradores : [];
  return colaboradores
    .filter((row) => row && row.id && row.nome)
    .filter((row) => !ctx.allowedSectorId || String(row.setor_id) === String(ctx.allowedSectorId));
}

function listRhHorasExtrasFromSnapshot(ctx, snapshot, filters) {
  const { month, setorId, colaboradorId, tipoHora } = filters || {};
  const safeMonth = String(month || '').trim();
  if (!safeMonth || !/^\d{4}-\d{2}$/.test(safeMonth)) {
    throw new Error('Mês inválido para RH. Use YYYY-MM.');
  }

  const setores = listRhSetoresFromSnapshot(ctx, snapshot);
  const colaboradores = listRhColaboradoresFromSnapshot(ctx, snapshot);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorSet = new Set(setores.map((s) => String(s.id)));

  const hoursByMonth = snapshot && snapshot.horasByMonth && typeof snapshot.horasByMonth === 'object'
    ? snapshot.horasByMonth
    : {};
  const monthRows = Array.isArray(hoursByMonth[safeMonth]) ? hoursByMonth[safeMonth] : [];

  const safeSetorId = String(setorId || '').trim();
  const safeColabId = String(colaboradorId || '').trim();
  const safeTipo = String(tipoHora || '').trim();
  const allowedSector = ctx.allowedSectorId ? String(ctx.allowedSectorId) : '';

  const rows = [];
  for (let i = 0; i < monthRows.length; i += 1) {
    const row = monthRows[i] || {};
    const id = String(row.id || '').trim();
    const colaborador_id = String(row.colaborador_id || '').trim();
    const data = String(row.data || '').trim();
    if (!id || !colaborador_id || !data) {
      continue;
    }

    const tipo_hora = String(row.tipo_hora || '').trim();
    if (safeTipo && tipo_hora !== safeTipo) {
      continue;
    }
    if (safeColabId && colaborador_id !== safeColabId) {
      continue;
    }

    const collab = collabById.get(String(colaborador_id));
    if (!collab) {
      continue;
    }
    const setorIdForCollab = String(collab.setor_id || '').trim();
    if (allowedSector && setorIdForCollab !== allowedSector) {
      continue;
    }
    if (safeSetorId && setorIdForCollab !== safeSetorId) {
      continue;
    }
    if (setorIdForCollab && setorSet.size && !setorSet.has(setorIdForCollab)) {
      continue;
    }

    rows.push({
      id,
      colaborador_id,
      data,
      quantidade_horas: Number(row.quantidade_horas || 0),
      tipo_hora,
      observacao: String(row.observacao || '').trim(),
      justificativa: String(row.justificativa || '').trim(),
      criado_por: String(row.criado_por || '').trim(),
      data_registro: String(row.data_registro || '').trim()
    });
  }

  rows.sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.id).localeCompare(String(a.id)));
  return rows;
}

function listRhAtrasosFromSnapshot(ctx, snapshot, filters) {
  const { month, setorId, colaboradorId } = filters || {};
  const safeMonth = String(month || '').trim();
  if (!safeMonth || !/^\d{4}-\d{2}$/.test(safeMonth)) {
    throw new Error('Mês inválido para RH. Use YYYY-MM.');
  }

  const setores = listRhSetoresFromSnapshot(ctx, snapshot);
  const colaboradores = listRhColaboradoresFromSnapshot(ctx, snapshot);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorSet = new Set(setores.map((s) => String(s.id)));

  const atrasosByMonth = snapshot && snapshot.atrasosByMonth && typeof snapshot.atrasosByMonth === 'object'
    ? snapshot.atrasosByMonth
    : {};
  const monthRows = Array.isArray(atrasosByMonth[safeMonth]) ? atrasosByMonth[safeMonth] : [];

  const safeSetorId = String(setorId || '').trim();
  const safeColabId = String(colaboradorId || '').trim();
  const allowedSector = ctx.allowedSectorId ? String(ctx.allowedSectorId) : '';

  const rows = [];
  for (let i = 0; i < monthRows.length; i += 1) {
    const row = monthRows[i] || {};
    const id = String(row.id || '').trim();
    const colaborador_id = String(row.colaborador_id || '').trim();
    const data = String(row.data || '').trim();
    if (!id || !colaborador_id || !data) {
      continue;
    }
    if (safeColabId && colaborador_id !== safeColabId) {
      continue;
    }

    const collab = collabById.get(String(colaborador_id));
    if (!collab) {
      continue;
    }
    const setorIdForCollab = String(collab.setor_id || '').trim();
    if (allowedSector && setorIdForCollab !== allowedSector) {
      continue;
    }
    if (safeSetorId && setorIdForCollab !== safeSetorId) {
      continue;
    }
    if (setorIdForCollab && setorSet.size && !setorSet.has(setorIdForCollab)) {
      continue;
    }

    rows.push({
      id,
      colaborador_id,
      data,
      quantidade_horas: Number(row.quantidade_horas || 0),
      observacao: String(row.observacao || '').trim(),
      criado_por: String(row.criado_por || '').trim(),
      data_registro: String(row.data_registro || '').trim()
    });
  }

  rows.sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.id).localeCompare(String(a.id)));
  return rows;
}

function getRhHorasExtrasMonthIndex(workbook) {
  const objs = readSheetObjects(workbook, 'HORAS_EXTRAS');
  const cached = RH_HORAS_MONTH_INDEX_CACHE.get(workbook) || null;
  if (cached && cached.source === objs && cached.byMonth) {
    return cached.byMonth;
  }
  const byMonth = new Map();
  for (let i = 0; i < objs.length; i += 1) {
    const row = objs[i] || {};
    if (String(row.DELETED_AT || '').trim()) {
      continue;
    }
    const data = String(row.DATA || '').trim();
    const month = data.length >= 7 ? data.slice(0, 7) : '';
    if (!/^\d{4}-\d{2}$/.test(month)) {
      continue;
    }
    if (!byMonth.has(month)) {
      byMonth.set(month, []);
    }
    byMonth.get(month).push(row);
  }
  RH_HORAS_MONTH_INDEX_CACHE.set(workbook, { source: objs, byMonth });
  return byMonth;
}

function listRhHorasExtras(ctx, workbook, filters) {
  const { month, setorId, colaboradorId, tipoHora } = filters || {};
  const safeMonth = String(month || '').trim();
  if (!safeMonth || !/^\d{4}-\d{2}$/.test(safeMonth)) {
    throw new Error('Mês inválido para RH. Use YYYY-MM.');
  }

  const setores = listRhSetores(ctx, workbook);
  const colaboradores = listRhColaboradores(ctx, workbook);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorSet = new Set(setores.map((s) => String(s.id)));

  const monthRows = (getRhHorasExtrasMonthIndex(workbook).get(safeMonth) || []);
  const safeSetorId = String(setorId || '').trim();
  const safeColabId = String(colaboradorId || '').trim();
  const safeTipo = String(tipoHora || '').trim();
  const allowedSector = ctx.allowedSectorId ? String(ctx.allowedSectorId) : '';

  const rows = [];
  for (let i = 0; i < monthRows.length; i += 1) {
    const row = monthRows[i] || {};

    const id = String(row.ID || '').trim();
    const colaborador_id = String(row.COLABORADOR_ID || '').trim();
    const data = String(row.DATA || '').trim();
    if (!id || !colaborador_id || !data) {
      continue;
    }

    const tipo_hora = String(row.TIPO_HORA || '').trim();
    if (safeTipo && tipo_hora !== safeTipo) {
      continue;
    }
    if (safeColabId && colaborador_id !== safeColabId) {
      continue;
    }

    const collab = collabById.get(String(colaborador_id));
    if (!collab) {
      continue;
    }
    const setorIdForCollab = String(collab.setor_id || '').trim();
    if (allowedSector && setorIdForCollab !== allowedSector) {
      continue;
    }
    if (safeSetorId && setorIdForCollab !== safeSetorId) {
      continue;
    }
    if (setorIdForCollab && setorSet.size && !setorSet.has(setorIdForCollab)) {
      continue;
    }

    rows.push({
      id,
      colaborador_id,
      data,
      quantidade_horas: Number(row.QUANTIDADE_HORAS || 0),
      tipo_hora,
      observacao: String(row.OBSERVACAO || '').trim(),
      justificativa: String(row.JUSTIFICATIVA || '').trim(),
      criado_por: String(row.CRIADO_POR || '').trim(),
      data_registro: String(row.DATA_REGISTRO || '').trim()
    });
  }

  rows.sort((a, b) => String(b.data).localeCompare(String(a.data)) || String(b.id).localeCompare(String(a.id)));
  return rows;
}

function computePreviousMonth(month) {
  const [y, m] = String(month).split('-').map((n) => Number(n));
  if (!Number.isFinite(y) || !Number.isFinite(m)) {
    return '';
  }
  const dt = new Date(Date.UTC(y, m - 1, 1));
  dt.setUTCMonth(dt.getUTCMonth() - 1);
  const yyyy = String(dt.getUTCFullYear()).padStart(4, '0');
  const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  return `${yyyy}-${mm}`;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function ymdFromDateUtc(dt) {
  const yyyy = String(dt.getUTCFullYear()).padStart(4, '0');
  const mm = pad2(dt.getUTCMonth() + 1);
  const dd = pad2(dt.getUTCDate());
  return `${yyyy}-${mm}-${dd}`;
}

function addDaysUtc(dt, days) {
  const copy = new Date(dt.getTime());
  copy.setUTCDate(copy.getUTCDate() + Number(days || 0));
  return copy;
}

function computeEasterSundayUtc(year) {
  // Meeus/Jones/Butcher - calendário Gregoriano
  const y = Number(year);
  const a = y % 19;
  const b = Math.floor(y / 100);
  const c = y % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31); // 3=Mar, 4=Apr
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(y, month - 1, day));
}

function normalizeMunicipioUf(value) {
  return normalizeText(String(value || '').replace(/['`]/g, ''));
}

function getDefaultMunicipalHolidaysForYear(municipioUf, year) {
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1970 || y > 2100) {
    return [];
  }
  const key = normalizeMunicipioUf(municipioUf);
  if (!key) {
    return [];
  }

  // Santa Bárbara d'Oeste - SP: 04/12 (Dia da Padroeira Santa Bárbara e fundação do município)
  if (key === 'SANTA BARBARA D OESTE/SP' || key === 'SANTA BARBARA D OESTE SP') {
    return [`${y}-12-04`];
  }
  return [];
}

function getBrazilHolidaySetUtc(year) {
  const y = Number(year);
  const set = new Set([
    `${y}-01-01`, // Confraternização Universal
    `${y}-04-21`, // Tiradentes
    `${y}-05-01`, // Dia do Trabalho
    `${y}-09-07`, // Independência
    `${y}-10-12`, // Nossa Senhora Aparecida
    `${y}-11-02`, // Finados
    `${y}-11-15`, // Proclamação da República
    `${y}-12-25`, // Natal
    `${y}-07-09` // SP: Revolução Constitucionalista (Estado de SP)
  ]);

  const easter = computeEasterSundayUtc(y);
  // Feriados móveis mais comuns (negócio): Carnaval (seg/ter), Paixão de Cristo, Corpus Christi
  set.add(ymdFromDateUtc(addDaysUtc(easter, -48))); // Carnaval (segunda)
  set.add(ymdFromDateUtc(addDaysUtc(easter, -47))); // Carnaval (terça)
  set.add(ymdFromDateUtc(addDaysUtc(easter, -2))); // Paixão de Cristo (sexta-feira santa)
  set.add(ymdFromDateUtc(addDaysUtc(easter, 60))); // Corpus Christi
  return set;
}

function computeBusinessDaysInMonth(month, config) {
  const safe = String(month || '').trim();
  const match = safe.match(/^(\d{4})-(\d{2})$/);
  if (!match) {
    return 0;
  }
  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1; // 0-based
  if (!Number.isFinite(year) || !Number.isFinite(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    return 0;
  }

  const first = new Date(Date.UTC(year, monthIndex, 1));
  const last = new Date(Date.UTC(year, monthIndex + 1, 0));
  let count = 0;
  for (let d = 1; d <= last.getUTCDate(); d += 1) {
    const dt = new Date(Date.UTC(year, monthIndex, d));
    const dow = dt.getUTCDay(); // 0=Dom, 6=Sáb
    if (dow >= 1 && dow <= 5) {
      count += 1;
    }
  }

  const shouldDiscount = config ? config.descontar_feriados !== false : true;
  if (!shouldDiscount) {
    return count;
  }

  const holidaySet = getBrazilHolidaySetUtc(year);
  getDefaultMunicipalHolidaysForYear(config && config.municipio_uf ? config.municipio_uf : '', year)
    .forEach((ymd) => holidaySet.add(String(ymd || '').trim()));
  const extras = (config && Array.isArray(config.feriados_extras)) ? config.feriados_extras : [];
  extras.forEach((ymd) => holidaySet.add(String(ymd || '').trim()));

  let discount = 0;
  holidaySet.forEach((ymd) => {
    if (!String(ymd).startsWith(`${safe}-`)) {
      return;
    }
    const [_, mm, dd] = String(ymd).split('-');
    const day = Number(dd);
    if (!Number.isFinite(day) || day < 1 || day > 31) {
      return;
    }
    const dt = new Date(Date.UTC(year, monthIndex, day));
    const dow = dt.getUTCDay();
    if (dow >= 1 && dow <= 5) {
      discount += 1;
    }
  });

  return Math.max(0, count - discount);
}

function computeRhKpisFromSnapshot(ctx, snapshot, filters) {
  const config = snapshot && snapshot.config && typeof snapshot.config === 'object' ? snapshot.config : {};
  const month = String(filters.month || '').trim();
  const prevMonth = computePreviousMonth(month);

  const rows = listRhHorasExtrasFromSnapshot(ctx, snapshot, filters);
  const colaboradores = listRhColaboradoresFromSnapshot(ctx, snapshot);
  const setores = listRhSetoresFromSnapshot(ctx, snapshot);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorById = new Map(setores.map((s) => [String(s.id), s]));

  const totalHoras = rows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  const diasSet = new Set(rows.map((r) => String(r.data || '').trim()).filter(Boolean));
  const diasComLancamento = diasSet.size;
  const diasTrabalhadosNoMes = computeBusinessDaysInMonth(month, config);
  const mediaDiaria = diasTrabalhadosNoMes ? totalHoras / diasTrabalhadosNoMes : 0;
  const colaboradoresNoPeriodo = new Set(rows.map((r) => String(r.colaborador_id))).size;

  const horasPorSetor = new Map();
  const horasPorColab = new Map();
  rows.forEach((r) => {
    const collab = collabById.get(String(r.colaborador_id));
    if (!collab) {
      return;
    }
    const setorId = String(collab.setor_id || '').trim();
    const setorName = setorId && setorById.get(setorId) ? String(setorById.get(setorId).nome_setor || '').trim() : 'SEM SETOR';
    horasPorSetor.set(setorName, (horasPorSetor.get(setorName) || 0) + (Number(r.quantidade_horas) || 0));
    const collabName = String(collab.nome || '').trim();
    horasPorColab.set(collabName, (horasPorColab.get(collabName) || 0) + (Number(r.quantidade_horas) || 0));
  });

  const porSetor = Array.from(horasPorSetor.entries())
    .map(([setor, horas]) => ({ setor, horas }))
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const rankingColaboradores = Array.from(horasPorColab.entries())
    .map(([colaborador, horas]) => {
      const margin = getRhMarginForHours(config, horas);
      return { colaborador, horas, margem_percent: margin.percent, margem_faixa: margin.faixa };
    })
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const faixaCount = new Map();
  let margemPonderada = 0;
  rankingColaboradores.forEach((row) => {
    const faixa = String(row.margem_faixa || '').trim() || '—';
    faixaCount.set(faixa, (faixaCount.get(faixa) || 0) + 1);
    margemPonderada += (Number(row.horas || 0) * Number(row.margem_percent || 0));
  });
  const margemMediaPonderada = totalHoras ? (margemPonderada / totalHoras) : 0;
  const colaboradoresPorFaixa = Array.from(faixaCount.entries()).map(([faixa, colaboradoresCount]) => ({ faixa, colaboradores: colaboradoresCount }));

  const limitePadrao = Number(config.limite_mensal_padrao || 40);
  const alertas = [];
  const byColabId = new Map();
  rows.forEach((r) => {
    const key = String(r.colaborador_id);
    byColabId.set(key, (byColabId.get(key) || 0) + (Number(r.quantidade_horas) || 0));
  });
  byColabId.forEach((horas, colabId) => {
    const collab = collabById.get(String(colabId));
    const override = collab ? Number(String(collab.limite_mensal || '').replace(',', '.')) : NaN;
    const limite = Number.isFinite(override) && override > 0 ? override : limitePadrao;
    if (Number.isFinite(limite) && limite > 0 && horas > limite) {
      alertas.push({
        colaborador_id: colabId,
        colaborador: collab ? String(collab.nome || '').trim() : colabId,
        horas,
        limite
      });
    }
  });

  let totalPrev = 0;
  if (prevMonth) {
    const prevRows = listRhHorasExtrasFromSnapshot(ctx, snapshot, { ...filters, month: prevMonth });
    totalPrev = prevRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  }
  const diff = totalHoras - totalPrev;
  const pct = totalPrev ? (diff / totalPrev) * 100 : (totalHoras ? 100 : 0);
  const comparacaoTexto = prevMonth
    ? `Vs ${prevMonth}: ${diff >= 0 ? '+' : ''}${diff.toFixed(2)} h (${pct.toFixed(1)}%)`
    : '';

  const evolucaoMensal = [];
  let cursor = month;
  for (let i = 0; i < 6; i += 1) {
    if (!cursor) {
      break;
    }
    const cursorRows = listRhHorasExtrasFromSnapshot(ctx, snapshot, { ...filters, month: cursor });
    const cursorTotal = cursorRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
    evolucaoMensal.push({ month: cursor, totalHoras: cursorTotal });
    cursor = computePreviousMonth(cursor);
  }
  evolucaoMensal.reverse();

  return {
    month,
    totalHoras,
    mediaDiaria,
    diasComLancamento,
    diasTrabalhadosNoMes,
    colaboradoresNoPeriodo,
    totalRegistros: rows.length,
    porSetor,
    rankingColaboradores,
    colaboradoresPorFaixa,
    margemMediaPonderada,
    alertas,
    limiteTexto: `Limite padrão: ${limitePadrao}h/mês`,
    comparacao: { prevMonth, totalPrev, diff, pct, texto: comparacaoTexto },
    evolucaoMensal,
    config
  };
}

function computeRhAtrasosKpisFromSnapshot(ctx, snapshot, filters) {
  const config = snapshot && snapshot.config && typeof snapshot.config === 'object' ? snapshot.config : {};
  const month = String(filters.month || '').trim();
  const prevMonth = computePreviousMonth(month);

  const rows = listRhAtrasosFromSnapshot(ctx, snapshot, filters);
  const colaboradores = listRhColaboradoresFromSnapshot(ctx, snapshot);
  const setores = listRhSetoresFromSnapshot(ctx, snapshot);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorById = new Map(setores.map((s) => [String(s.id), s]));

  const totalHoras = rows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  const diasSet = new Set(rows.map((r) => String(r.data || '').trim()).filter(Boolean));
  const diasComLancamento = diasSet.size;
  const diasTrabalhadosNoMes = computeBusinessDaysInMonth(month, config);
  const mediaDiaria = diasTrabalhadosNoMes ? totalHoras / diasTrabalhadosNoMes : 0;
  const colaboradoresNoPeriodo = new Set(rows.map((r) => String(r.colaborador_id))).size;

  const horasPorSetor = new Map();
  const horasPorColab = new Map();
  rows.forEach((r) => {
    const collab = collabById.get(String(r.colaborador_id));
    if (!collab) {
      return;
    }
    const setorId = String(collab.setor_id || '').trim();
    const setorName = setorId && setorById.get(setorId) ? String(setorById.get(setorId).nome_setor || '').trim() : 'SEM SETOR';
    horasPorSetor.set(setorName, (horasPorSetor.get(setorName) || 0) + (Number(r.quantidade_horas) || 0));
    const collabName = String(collab.nome || '').trim();
    horasPorColab.set(collabName, (horasPorColab.get(collabName) || 0) + (Number(r.quantidade_horas) || 0));
  });

  const porSetor = Array.from(horasPorSetor.entries())
    .map(([setor, horas]) => ({ setor, horas }))
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const rankingColaboradores = Array.from(horasPorColab.entries())
    .map(([colaborador, horas]) => ({ colaborador, horas }))
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const alertas = [];
  const byColabId = new Map();
  rows.forEach((r) => {
    const key = String(r.colaborador_id);
    byColabId.set(key, (byColabId.get(key) || 0) + (Number(r.quantidade_horas) || 0));
  });
  byColabId.forEach((horas, colabId) => {
    const collab = collabById.get(String(colabId));
    const name = collab ? String(collab.nome || '').trim() : String(colabId);
    if (horas >= 1) {
      alertas.push({ colaborador: name, horas });
    }
  });
  alertas.sort((a, b) => Number(b.horas) - Number(a.horas));

  let totalPrev = 0;
  if (prevMonth) {
    const prevRows = listRhAtrasosFromSnapshot(ctx, snapshot, { ...filters, month: prevMonth });
    totalPrev = prevRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  }
  const diff = totalHoras - totalPrev;
  const pct = totalPrev ? (diff / totalPrev) * 100 : (totalHoras ? 100 : 0);
  const comparacaoTexto = prevMonth
    ? `Vs ${prevMonth}: ${diff >= 0 ? '+' : ''}${diff.toFixed(2)} h (${pct.toFixed(1)}%)`
    : '';

  const evolucaoMensal = [];
  let cursor = month;
  for (let i = 0; i < 6; i += 1) {
    if (!cursor) {
      break;
    }
    const cursorRows = listRhAtrasosFromSnapshot(ctx, snapshot, { ...filters, month: cursor });
    const cursorTotal = cursorRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
    evolucaoMensal.push({ month: cursor, totalHoras: cursorTotal });
    cursor = computePreviousMonth(cursor);
  }
  evolucaoMensal.reverse();

  return {
    month,
    totalHoras,
    mediaDiaria,
    diasComLancamento,
    diasTrabalhadosNoMes,
    colaboradoresNoPeriodo,
    totalRegistros: rows.length,
    porSetor,
    rankingColaboradores,
    alertas,
    comparacao: { prevMonth, totalPrev, diff, pct, texto: comparacaoTexto },
    evolucaoMensal,
    config
  };
}

function computeRhKpis(ctx, workbook, filters) {
  const config = getRhConfig(workbook);
  const month = String(filters.month || '').trim();
  const prevMonth = computePreviousMonth(month);

  const rows = listRhHorasExtras(ctx, workbook, filters);
  const colaboradores = listRhColaboradores(ctx, workbook);
  const setores = listRhSetores(ctx, workbook);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
  const setorById = new Map(setores.map((s) => [String(s.id), s]));

  const totalHoras = rows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  const diasSet = new Set(rows.map((r) => String(r.data || '').trim()).filter(Boolean));
  const diasComLancamento = diasSet.size;
  const diasTrabalhadosNoMes = computeBusinessDaysInMonth(month, config);
  const mediaDiaria = diasTrabalhadosNoMes ? totalHoras / diasTrabalhadosNoMes : 0;
  const colaboradoresNoPeriodo = new Set(rows.map((r) => String(r.colaborador_id))).size;

  const horasPorSetor = new Map();
  const horasPorColab = new Map();
  rows.forEach((r) => {
    const collab = collabById.get(String(r.colaborador_id));
    if (!collab) {
      return;
    }
    const setorId = String(collab.setor_id || '').trim();
    const setorName = setorId && setorById.get(setorId) ? String(setorById.get(setorId).nome_setor || '').trim() : 'SEM SETOR';
    horasPorSetor.set(setorName, (horasPorSetor.get(setorName) || 0) + (Number(r.quantidade_horas) || 0));
    const collabName = String(collab.nome || '').trim();
    horasPorColab.set(collabName, (horasPorColab.get(collabName) || 0) + (Number(r.quantidade_horas) || 0));
  });

  const porSetor = Array.from(horasPorSetor.entries())
    .map(([setor, horas]) => ({ setor, horas }))
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const rankingColaboradores = Array.from(horasPorColab.entries())
    .map(([colaborador, horas]) => {
      const margin = getRhMarginForHours(config, horas);
      return { colaborador, horas, margem_percent: margin.percent, margem_faixa: margin.faixa };
    })
    .sort((a, b) => Number(b.horas) - Number(a.horas));

  const faixaCount = new Map();
  let margemPonderada = 0;
  rankingColaboradores.forEach((row) => {
    const faixa = String(row.margem_faixa || '').trim() || '—';
    faixaCount.set(faixa, (faixaCount.get(faixa) || 0) + 1);
    margemPonderada += (Number(row.horas || 0) * Number(row.margem_percent || 0));
  });
  const margemMediaPonderada = totalHoras ? (margemPonderada / totalHoras) : 0;
  const colaboradoresPorFaixa = Array.from(faixaCount.entries()).map(([faixa, colaboradores]) => ({ faixa, colaboradores }));

  const limitePadrao = Number(config.limite_mensal_padrao || 40);
  const alertas = [];
  const byColabId = new Map();
  rows.forEach((r) => {
    const key = String(r.colaborador_id);
    byColabId.set(key, (byColabId.get(key) || 0) + (Number(r.quantidade_horas) || 0));
  });
  byColabId.forEach((horas, colabId) => {
    const collab = collabById.get(String(colabId));
    const override = collab ? Number(String(collab.limite_mensal || '').replace(',', '.')) : NaN;
    const limite = Number.isFinite(override) && override > 0 ? override : limitePadrao;
    if (Number.isFinite(limite) && limite > 0 && horas > limite) {
      alertas.push({
        colaborador_id: colabId,
        colaborador: collab ? String(collab.nome || '').trim() : colabId,
        horas,
        limite
      });
    }
  });

  let totalPrev = 0;
  if (prevMonth) {
    const prevRows = listRhHorasExtras(ctx, workbook, { ...filters, month: prevMonth });
    totalPrev = prevRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
  }
  const diff = totalHoras - totalPrev;
  const pct = totalPrev ? (diff / totalPrev) * 100 : (totalHoras ? 100 : 0);
  const comparacaoTexto = prevMonth
    ? `Vs ${prevMonth}: ${diff >= 0 ? '+' : ''}${diff.toFixed(2)} h (${pct.toFixed(1)}%)`
    : '';

  const evolucaoMensal = [];
  let cursor = month;
  for (let i = 0; i < 6; i += 1) {
    if (!cursor) {
      break;
    }
    const cursorRows = listRhHorasExtras(ctx, workbook, { ...filters, month: cursor });
    const cursorTotal = cursorRows.reduce((sum, r) => sum + (Number(r.quantidade_horas) || 0), 0);
    evolucaoMensal.push({ month: cursor, totalHoras: cursorTotal });
    cursor = computePreviousMonth(cursor);
  }
  evolucaoMensal.reverse();

  return {
    month,
    totalHoras,
    mediaDiaria,
    diasComLancamento,
    diasTrabalhadosNoMes,
    colaboradoresNoPeriodo,
    totalRegistros: rows.length,
    porSetor,
    rankingColaboradores,
    colaboradoresPorFaixa,
    margemMediaPonderada,
    alertas,
    limiteTexto: `Limite padrão: ${limitePadrao}h/mês`,
    comparacao: { prevMonth, totalPrev, diff, pct, texto: comparacaoTexto },
    evolucaoMensal,
    config
  };
}

function appendRhAudit(workbook, { username, action, entity, entityId, before, after }) {
  const aoa = readSheetAoa(workbook, 'AUDIT_LOG');
  const next = Array.isArray(aoa) && aoa.length ? [...aoa] : [[
    'ID',
    'DATA_HORA',
    'USERNAME',
    'ACTION',
    'ENTITY',
    'ENTITY_ID',
    'BEFORE_JSON',
    'AFTER_JSON'
  ]];
  next.push([
    String(generateRhId()),
    new Date().toISOString(),
    normalizeUsername(username),
    String(action || '').trim(),
    String(entity || '').trim(),
    String(entityId || '').trim(),
    before ? JSON.stringify(before) : '',
    after ? JSON.stringify(after) : ''
  ]);
  writeSheetAoa(workbook, 'AUDIT_LOG', next);
}

function getHeaderFromAoa(aoa, defaultHeader) {
  const headerRow = Array.isArray(aoa && aoa[0]) ? aoa[0] : null;
  if (headerRow && headerRow.length) {
    return headerRow.map((h) => String(h || '').trim().toUpperCase());
  }
  return (defaultHeader || []).map((h) => String(h || '').trim().toUpperCase());
}

function objectsToAoa(headers, objects) {
  const safeHeaders = (headers || []).map((h) => String(h || '').trim().toUpperCase()).filter(Boolean);
  const rows = Array.isArray(objects) ? objects : [];
  const aoa = [safeHeaders];
  rows.forEach((obj) => {
    const row = safeHeaders.map((h) => (obj && Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ''));
    aoa.push(row);
  });
  return aoa;
}

function upsertRhSetor({ username, id, nome_setor, gestor_responsavel, centro_custo }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'SETORES');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'NOME_SETOR',
    'GESTOR_RESPONSAVEL',
    'CENTRO_CUSTO',
    'CREATED_AT',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const safeId = String(id || '').trim();
  const safeNome = String(nome_setor || '').trim();
  if (!safeNome) {
    throw new Error('Nome do setor obrigatório.');
  }
  const safeGestor = String(gestor_responsavel || '').trim();
  const safeCc = String(centro_custo || '').trim();

  let before = null;
  let target = null;
  if (safeId) {
    target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
    if (target) {
      before = { ...target };
    }
  }
  if (!target) {
    target = {
      ID: safeId || String(generateRhId()),
      CREATED_AT: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    };
    rows.push(target);
  }

  target.NOME_SETOR = safeNome;
  target.GESTOR_RESPONSAVEL = safeGestor;
  target.CENTRO_CUSTO = safeCc;
  target.UPDATED_AT = now;
  if (String(target.DELETED_AT || '').trim()) {
    target.DELETED_AT = '';
  }

  writeSheetAoa(workbook, 'SETORES', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: before ? 'SETORES_UPDATE' : 'SETORES_CREATE',
    entity: 'SETORES',
    entityId: target.ID,
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);
  return { setor: { id: String(target.ID), nome_setor: safeNome, gestor_responsavel: safeGestor, centro_custo: safeCc } };
}

function deleteRhSetor({ username, id }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'SETORES');
  const header = getHeaderFromAoa(aoa, ['ID', 'NOME_SETOR', 'GESTOR_RESPONSAVEL', 'CENTRO_CUSTO', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const safeId = String(id || '').trim();
  if (!safeId) {
    throw new Error('ID do setor obrigatório.');
  }
  const target = rows.find((r) => String(r.ID || '').trim() === safeId);
  if (!target) {
    throw new Error('Setor não encontrado.');
  }
  const before = { ...target };
  target.DELETED_AT = now;
  target.UPDATED_AT = now;
  writeSheetAoa(workbook, 'SETORES', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: 'SETORES_DELETE', entity: 'SETORES', entityId: safeId, before, after: target });
  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function upsertRhColaborador({ username, id, nome, matricula, setor_id, cargo, status, data_admissao, limite_mensal }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'COLABORADORES');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'NOME',
    'MATRICULA',
    'SETOR_ID',
    'CARGO',
    'STATUS',
    'DATA_ADMISSAO',
    'LIMITE_MENSAL',
    'CREATED_AT',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const safeId = String(id || '').trim();
  const safeNome = String(nome || '').trim();
  if (!safeNome) {
    throw new Error('Nome do colaborador obrigatório.');
  }
  const safeMat = String(matricula || '').trim();
  const safeSetor = String(setor_id || '').trim();
  const safeCargo = String(cargo || '').trim();
  const safeStatus = String(status || '').trim().toLowerCase() === 'inativo' ? 'inativo' : 'ativo';
  const safeAd = String(data_admissao || '').trim();
  const safeLim = String(limite_mensal || '').trim();

  let before = null;
  let target = null;
  if (safeId) {
    target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
    if (target) {
      before = { ...target };
    }
  }
  if (!target) {
    target = {
      ID: safeId || String(generateRhId()),
      CREATED_AT: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    };
    rows.push(target);
  }

  target.NOME = safeNome;
  target.MATRICULA = safeMat;
  target.SETOR_ID = safeSetor;
  target.CARGO = safeCargo;
  target.STATUS = safeStatus;
  target.DATA_ADMISSAO = safeAd;
  target.LIMITE_MENSAL = safeLim;
  target.UPDATED_AT = now;
  if (String(target.DELETED_AT || '').trim()) {
    target.DELETED_AT = '';
  }

  writeSheetAoa(workbook, 'COLABORADORES', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: before ? 'COLABORADORES_UPDATE' : 'COLABORADORES_CREATE',
    entity: 'COLABORADORES',
    entityId: target.ID,
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);
  return { colaborador: { id: String(target.ID), nome: safeNome } };
}

function toggleRhColaboradorStatus({ username, id }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'COLABORADORES');
  const header = getHeaderFromAoa(aoa, ['ID', 'NOME', 'MATRICULA', 'SETOR_ID', 'CARGO', 'STATUS', 'DATA_ADMISSAO', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => String(r.ID || '').trim() === safeId);
  if (!target) {
    throw new Error('Colaborador não encontrado.');
  }
  const before = { ...target };
  const current = String(target.STATUS || '').trim().toLowerCase();
  target.STATUS = current === 'inativo' ? 'ativo' : 'inativo';
  target.UPDATED_AT = now;
  writeSheetAoa(workbook, 'COLABORADORES', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: 'COLABORADORES_TOGGLE', entity: 'COLABORADORES', entityId: safeId, before, after: target });
  writeWorkbookFile(workbook, filePath);
  return { ok: true, status: target.STATUS };
}

function deleteRhColaborador({ username, id }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'COLABORADORES');
  const header = getHeaderFromAoa(aoa, ['ID', 'NOME', 'MATRICULA', 'SETOR_ID', 'CARGO', 'STATUS', 'DATA_ADMISSAO', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => String(r.ID || '').trim() === safeId);
  if (!target) {
    throw new Error('Colaborador não encontrado.');
  }
  const before = { ...target };
  target.DELETED_AT = now;
  target.UPDATED_AT = now;
  writeSheetAoa(workbook, 'COLABORADORES', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: 'COLABORADORES_DELETE', entity: 'COLABORADORES', entityId: safeId, before, after: target });
  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function upsertRhHoraExtra({ username, id, colaborador_id, data, quantidade_horas, tipo_hora, observacao, justificativa }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const config = getRhConfig(workbook);
  const now = new Date().toISOString();

  const safeColabId = String(colaborador_id || '').trim();
  const safeData = String(data || '').trim();
  const safeTipo = String(tipo_hora || '').trim() || '50%';
  const safeObs = String(observacao || '').trim();
  const safeJust = String(justificativa || '').trim();
  const safeQtd = Number(quantidade_horas || 0);
  if (!safeColabId) {
    throw new Error('Colaborador obrigatório.');
  }
  if (!safeData || !/^\d{4}-\d{2}-\d{2}$/.test(safeData)) {
    throw new Error('Data inválida.');
  }
  if (!Number.isFinite(safeQtd) || safeQtd <= 0) {
    throw new Error('Quantidade de horas inválida.');
  }
  const limitJust = Number(config.justificativa_acima_horas || 2);
  if (Number.isFinite(limitJust) && safeQtd > limitJust && !safeJust) {
    throw new Error(`Justificativa obrigatória acima de ${limitJust}h.`);
  }

  const aoa = readSheetAoa(workbook, 'HORAS_EXTRAS');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'TIPO_HORA',
    'OBSERVACAO',
    'JUSTIFICATIVA',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const existingByKey = rows.find((r) => (
    !String(r.DELETED_AT || '').trim()
    && String(r.COLABORADOR_ID || '').trim() === safeColabId
    && String(r.DATA || '').trim() === safeData
    && String(r.TIPO_HORA || '').trim() === safeTipo
  )) || null;

  const safeId = String(id || '').trim();
  let before = null;
  let target = null;
  if (safeId) {
    target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
    if (target) {
      before = { ...target };
    }
  }
  if (!safeId && existingByKey) {
    target = existingByKey;
    before = { ...existingByKey };
  }
  if (safeId && existingByKey && String(existingByKey.ID || '').trim() !== safeId) {
    throw new Error('Já existe um lançamento de H.E para este colaborador/data/tipo.');
  }
  if (!target) {
    target = {
      ID: safeId || String(generateRhId()),
      CRIADO_POR: ctx.username,
      DATA_REGISTRO: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    };
    rows.push(target);
  }

  target.COLABORADOR_ID = safeColabId;
  target.DATA = safeData;
  target.QUANTIDADE_HORAS = safeQtd;
  target.TIPO_HORA = safeTipo;
  target.OBSERVACAO = safeObs;
  target.JUSTIFICATIVA = safeJust;
  target.UPDATED_AT = now;
  if (String(target.DELETED_AT || '').trim()) {
    target.DELETED_AT = '';
  }

  writeSheetAoa(workbook, 'HORAS_EXTRAS', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: before ? 'HORAS_EXTRAS_UPDATE' : 'HORAS_EXTRAS_CREATE',
    entity: 'HORAS_EXTRAS',
    entityId: target.ID,
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);
  return { ok: true, id: String(target.ID) };
}

function upsertRhAtraso({ username, id, colaborador_id, data, quantidade_horas, observacao }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();

  const safeColabId = String(colaborador_id || '').trim();
  const safeData = String(data || '').trim();
  const safeObs = String(observacao || '').trim();
  const safeQtd = Number(quantidade_horas || 0);
  if (!safeColabId) {
    throw new Error('Colaborador obrigatório.');
  }
  if (!safeData || !/^\d{4}-\d{2}-\d{2}$/.test(safeData)) {
    throw new Error('Data inválida.');
  }
  if (!Number.isFinite(safeQtd) || safeQtd <= 0) {
    throw new Error('Quantidade de horas inválida.');
  }

  const aoa = readSheetAoa(workbook, 'ATRASOS');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'OBSERVACAO',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const existingByKey = rows.find((r) => (
    !String(r.DELETED_AT || '').trim()
    && String(r.COLABORADOR_ID || '').trim() === safeColabId
    && String(r.DATA || '').trim() === safeData
  )) || null;

  const safeId = String(id || '').trim();
  let before = null;
  let target = null;
  if (safeId) {
    target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
    if (target) {
      before = { ...target };
    }
  }
  if (!safeId && existingByKey) {
    target = existingByKey;
    before = { ...existingByKey };
  }
  if (safeId && existingByKey && String(existingByKey.ID || '').trim() !== safeId) {
    throw new Error('Já existe um atraso lançado para este colaborador/data.');
  }
  if (!target) {
    target = {
      ID: safeId || String(generateRhId()),
      CRIADO_POR: ctx.username,
      DATA_REGISTRO: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    };
    rows.push(target);
  }

  target.COLABORADOR_ID = safeColabId;
  target.DATA = safeData;
  target.QUANTIDADE_HORAS = safeQtd;
  target.OBSERVACAO = safeObs;
  target.UPDATED_AT = now;
  if (String(target.DELETED_AT || '').trim()) {
    target.DELETED_AT = '';
  }

  writeSheetAoa(workbook, 'ATRASOS', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: before ? 'ATRASOS_UPDATE' : 'ATRASOS_CREATE',
    entity: 'ATRASOS',
    entityId: target.ID,
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);
  return { ok: true, id: String(target.ID) };
}

function parseTimeToMinutes(value) {
  const text = String(value || '').trim();
  if (!text) {
    return null;
  }
  const m = text.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) {
    return null;
  }
  const hh = Number(m[1]);
  const mm = Number(m[2]);
  if (!Number.isFinite(hh) || !Number.isFinite(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) {
    return null;
  }
  return (hh * 60) + mm;
}

function formatMinutesToHoursDecimal(totalMinutes) {
  const mins = Number(totalMinutes);
  if (!Number.isFinite(mins) || mins <= 0) {
    return 0;
  }
  return Math.round((mins / 60) * 100) / 100;
}

function formatMinutesToHoursHm(totalMinutes) {
  const mins = Number(totalMinutes);
  if (!Number.isFinite(mins) || mins <= 0) {
    return '';
  }
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}:${String(m).padStart(2, '0')}`;
}

function computeHeTotalMinutes(hora_inicio, hora_fim) {
  const start = parseTimeToMinutes(hora_inicio);
  const end = parseTimeToMinutes(hora_fim);
  if (start == null || end == null) {
    return null;
  }
  if (end <= start) {
    return null;
  }
  return end - start;
}

function computeHeTotalHours(hora_inicio, hora_fim) {
  const totalMinutes = computeHeTotalMinutes(hora_inicio, hora_fim);
  if (totalMinutes == null) {
    return null;
  }
  return formatMinutesToHoursDecimal(totalMinutes);
}

function formatDateToYmdLocal(date) {
  const dt = date instanceof Date ? date : new Date();
  if (Number.isNaN(dt.getTime())) {
    return '';
  }
  const yyyy = String(dt.getFullYear()).padStart(4, '0');
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function getRhSolicitacoesHeSheet(workbook) {
  const aoa = readSheetAoa(workbook, 'SOLICITACOES_HE');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'NUMERO_DOCUMENTO',
    'DATA_SOLICITACAO',
    'DATA_HE',
    'COLABORADOR_ID',
    'NOME_COLABORADOR',
    'SETOR',
    'FINALIDADE',
    'HORA_INICIO',
    'HORA_FIM',
    'TOTAL_HORAS',
    'SOLICITANTE',
    'DATA_CRIACAO',
    'ULTIMA_ATUALIZACAO',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  return { aoa, header, rows };
}

function nextRhSolicitacaoHeId(rows) {
  const list = Array.isArray(rows) ? rows : [];
  let max = 0;
  list.forEach((r) => {
    const raw = String(r.ID ?? '').trim();
    const n = Number(raw);
    if (Number.isFinite(n) && n > max) {
      max = n;
    }
  });
  return max + 1;
}

function nextRhSolicitacaoHeNumeroDocumento(rows, year) {
  const yyyy = Number(year);
  if (!Number.isFinite(yyyy) || yyyy < 2000 || yyyy > 2100) {
    throw new Error('Ano inválido para numeração.');
  }
  const prefix = `HE-${yyyy}-`;
  let maxSeq = 0;
  (Array.isArray(rows) ? rows : []).forEach((r) => {
    const doc = String(r.NUMERO_DOCUMENTO || '').trim();
    if (!doc.startsWith(prefix)) {
      return;
    }
    const seq = Number(doc.slice(prefix.length));
    if (Number.isFinite(seq) && seq > maxSeq) {
      maxSeq = seq;
    }
  });
  const next = maxSeq + 1;
  return `${prefix}${String(next).padStart(4, '0')}`;
}

function canEditRhSolicitacaoRow(ctx, row) {
  if (ctx && ctx.canEdit) {
    return true;
  }
  const owner = normalizeUsername(row && row.SOLICITANTE ? row.SOLICITANTE : '');
  return owner && ctx && normalizeUsername(ctx.username) === owner;
}

function upsertRhSolicitacaoHe(payload) {
  const ctx = ensureRhAccess(String(payload && payload.username || '').trim(), { requireEdit: false });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const todayYmd = formatDateToYmdLocal(new Date());

  const safeId = String(payload && payload.id ? payload.id : '').trim();
  const data_he = String(payload && payload.data_he ? payload.data_he : '').trim();
  const colaborador_id = String(payload && payload.colaborador_id ? payload.colaborador_id : '').trim();
  const nome_colaborador = String(payload && payload.nome_colaborador ? payload.nome_colaborador : '').trim();
  const setor = String(payload && payload.setor ? payload.setor : '').trim();
  const finalidade = String(payload && payload.finalidade ? payload.finalidade : '').trim();
  const hora_inicio = String(payload && payload.hora_inicio ? payload.hora_inicio : '').trim();
  const hora_fim = String(payload && payload.hora_fim ? payload.hora_fim : '').trim();
  const solicitante = normalizeUsername(ctx.username);

  if (!data_he || !/^\d{4}-\d{2}-\d{2}$/.test(data_he)) {
    throw new Error('Informe a data da H.E.');
  }
  if (!colaborador_id) {
    throw new Error('Selecione o colaborador.');
  }
  if (!nome_colaborador) {
    throw new Error('Nome do colaborador obrigatório.');
  }
  if (!setor) {
    throw new Error('Setor obrigatório.');
  }
  if (!finalidade) {
    throw new Error('Informe a finalidade.');
  }
  const total_horas = computeHeTotalHours(hora_inicio, hora_fim);
  if (total_horas == null || !Number.isFinite(total_horas) || total_horas <= 0) {
    throw new Error('Horário inválido (início/fim).');
  }

  const { header, rows } = getRhSolicitacoesHeSheet(workbook);

  // Anti-duplicidade: mesmo colaborador + data + inicio + fim
  const dup = rows.find((r) => (
    !String(r.DELETED_AT || '').trim()
    && String(r.COLABORADOR_ID || '').trim() === colaborador_id
    && String(r.DATA_HE || '').trim() === data_he
    && String(r.HORA_INICIO || '').trim() === hora_inicio
    && String(r.HORA_FIM || '').trim() === hora_fim
  )) || null;
  if (!safeId && dup) {
    throw new Error('Já existe uma solicitação com o mesmo colaborador/data/horário.');
  }
  if (safeId && dup && String(dup.ID || '').trim() !== safeId) {
    throw new Error('Já existe outra solicitação com o mesmo colaborador/data/horário.');
  }

  let target = null;
  let before = null;
  if (safeId) {
    target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
    if (!target) {
      throw new Error('Solicitação não encontrada.');
    }
    if (!canEditRhSolicitacaoRow(ctx, target)) {
      throw new Error('Você não tem permissão para editar esta solicitação.');
    }
    before = { ...target };
  }

  const isCreate = !target;
  if (!target) {
    const nextId = nextRhSolicitacaoHeId(rows);
    const year = Number(String(todayYmd).slice(0, 4));
    const numero_documento = nextRhSolicitacaoHeNumeroDocumento(rows, year);
    target = {
      ID: String(nextId),
      NUMERO_DOCUMENTO: numero_documento,
      DATA_SOLICITACAO: todayYmd,
      SOLICITANTE: solicitante,
      DATA_CRIACAO: now,
      ULTIMA_ATUALIZACAO: now,
      DELETED_AT: ''
    };
    rows.push(target);
  }

  target.DATA_HE = data_he;
  target.COLABORADOR_ID = colaborador_id;
  target.NOME_COLABORADOR = nome_colaborador;
  target.SETOR = setor;
  target.FINALIDADE = finalidade;
  target.HORA_INICIO = hora_inicio;
  target.HORA_FIM = hora_fim;
  target.TOTAL_HORAS = Number(total_horas);
  if (!target.SOLICITANTE) {
    target.SOLICITANTE = solicitante;
  }
  if (!target.DATA_CRIACAO) {
    target.DATA_CRIACAO = now;
  }
  target.ULTIMA_ATUALIZACAO = now;
  if (String(target.DELETED_AT || '').trim()) {
    target.DELETED_AT = '';
  }

  writeSheetAoa(workbook, 'SOLICITACOES_HE', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: isCreate ? 'SOLICITACOES_HE_CREATE' : 'SOLICITACOES_HE_UPDATE',
    entity: 'SOLICITACOES_HE',
    entityId: String(target.ID),
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);

  return {
    ok: true,
    row: {
      id: String(target.ID),
      numero_documento: String(target.NUMERO_DOCUMENTO || '').trim(),
      data_solicitacao: String(target.DATA_SOLICITACAO || '').trim(),
      data_he: String(target.DATA_HE || '').trim(),
      colaborador_id: String(target.COLABORADOR_ID || '').trim(),
      nome_colaborador: String(target.NOME_COLABORADOR || '').trim(),
      setor: String(target.SETOR || '').trim(),
      finalidade: String(target.FINALIDADE || '').trim(),
      hora_inicio: String(target.HORA_INICIO || '').trim(),
      hora_fim: String(target.HORA_FIM || '').trim(),
      total_horas: Number(target.TOTAL_HORAS || 0),
      solicitante: String(target.SOLICITANTE || '').trim(),
      data_criacao: String(target.DATA_CRIACAO || '').trim(),
      ultima_atualizacao: String(target.ULTIMA_ATUALIZACAO || '').trim()
    }
  };
}

function formatYmdToPtBr(ymd) {
  const text = String(ymd || '').trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) {
    return text;
  }
  return `${m[3]}/${m[2]}/${m[1]}`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getRhSolicitacaoHePdfDir() {
  const fiscalRoot = getFiscalRoot();
  const dir = path.join(fiscalRoot, 'PartsSeals', 'Solicitação de Horas Extras');
  ensureDirSync(dir);
  try {
    fs.accessSync(dir, fs.constants.W_OK);
  } catch (error) {
    throw new Error(`Sem permissão de escrita em: ${dir}`);
  }
  return dir;
}

function getRhSolicitacaoHePdfPath(numeroDocumento) {
  const safe = String(numeroDocumento || '').trim();
  if (!safe) {
    throw new Error('Número do documento inválido.');
  }
  const dir = getRhSolicitacaoHePdfDir();
  return path.join(dir, `SOLICITACAO_${safe}.pdf`);
}

function getPartsSealsLogoDataUrl() {
  try {
    const logoPath = path.join(app.getAppPath(), 'img', 'LogoParts.png');
    if (!fs.existsSync(logoPath)) {
      return '';
    }
    const buf = fs.readFileSync(logoPath);
    return `data:image/png;base64,${buf.toString('base64')}`;
  } catch (error) {
    return '';
  }
}

function buildRhSolicitacaoHePdfHtml(row) {
  const logoUrl = getPartsSealsLogoDataUrl();
  const numero = escapeHtml(row.numero_documento);
  const dataSolic = escapeHtml(formatYmdToPtBr(row.data_solicitacao));
  const dataHe = escapeHtml(formatYmdToPtBr(row.data_he));
  const solicitante = escapeHtml(row.solicitante);
  const colaborador = escapeHtml(row.nome_colaborador);
  const setor = escapeHtml(row.setor);
  const inicio = escapeHtml(row.hora_inicio);
  const fim = escapeHtml(row.hora_fim);
  const totalMinutes = computeHeTotalMinutes(row.hora_inicio, row.hora_fim);
  const totalHm = totalMinutes != null ? formatMinutesToHoursHm(totalMinutes) : '';
  const totalFallback = formatMinutesToHoursHm(Math.round(Number(row.total_horas || 0) * 60));
  const total = escapeHtml(totalHm || totalFallback || '');
  const finalidade = escapeHtml(row.finalidade).replace(/\n/g, '<br />');

  return `<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Solicitação de Horas Extras</title>
    <style>
      :root { --text:#111827; --muted:#6b7280; --border:#e5e7eb; --accent:#7f1d1d; }
      * { box-sizing: border-box; }
      body { font-family: Arial, Helvetica, sans-serif; color: var(--text); margin: 0; padding: 36px 40px; }
      .top { display:flex; align-items:center; justify-content:space-between; gap:16px; }
      .logo { height: 46px; object-fit: contain; }
      .title { text-align:center; flex: 1; }
      .title h1 { margin: 0; font-size: 18px; letter-spacing: 0.6px; }
      .title h2 { margin: 6px 0 0; font-size: 12px; color: var(--muted); font-weight: 600; }
      .pill { border: 1px solid var(--border); border-left: 5px solid var(--accent); padding: 10px 12px; border-radius: 12px; }
      .meta { display:grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 18px; }
      .meta .k { color: var(--muted); font-size: 11px; font-weight: 700; letter-spacing: 0.2px; }
      .meta .v { margin-top: 2px; font-size: 12px; font-weight: 700; }
      .section { margin-top: 16px; border: 1px solid var(--border); border-radius: 14px; overflow: hidden; }
      .section .head { background: #f9fafb; padding: 10px 12px; border-bottom: 1px solid var(--border); font-size: 12px; font-weight: 800; }
      .section .body { padding: 12px; display:grid; gap: 10px; }
      .row { display:grid; grid-template-columns: 160px 1fr; gap: 10px; }
      .row .label { color: var(--muted); font-size: 12px; font-weight: 700; }
      .row .value { font-size: 12px; font-weight: 700; }
      .finalidade { padding: 10px 12px; border: 1px dashed var(--border); border-radius: 12px; background:#ffffff; line-height: 1.35; }
      .sign { margin-top: 22px; display:grid; gap: 14px; }
      .line { display:grid; grid-template-columns: 1fr; gap: 6px; }
      .line .lbl { color: var(--muted); font-size: 12px; font-weight: 700; }
      .line .sig { border-bottom: 1px solid #111827; height: 18px; }
      .footerNote { margin-top: 12px; color: var(--muted); font-size: 10px; }
    </style>
  </head>
  <body>
    <div class="top">
      <div style="width: 160px;">
        ${logoUrl ? `<img class="logo" src="${logoUrl}" alt="PartsSeals" />` : ''}
      </div>
      <div class="title">
        <h1>SOLICITAÇÃO DE HORAS EXTRAS</h1>
        <h2>${numero}</h2>
      </div>
      <div style="width: 160px;"></div>
    </div>

    <div class="meta">
      <div class="pill">
        <div class="k">NÚMERO DO DOCUMENTO</div>
        <div class="v">${numero}</div>
      </div>
      <div class="pill">
        <div class="k">DATA DA SOLICITAÇÃO</div>
        <div class="v">${dataSolic}</div>
      </div>
    </div>

    <div class="section">
      <div class="head">INFORMAÇÕES DA SOLICITAÇÃO</div>
      <div class="body">
        <div class="row"><div class="label">Solicitante</div><div class="value">${solicitante}</div></div>
        <div class="row"><div class="label">Colaborador</div><div class="value">${colaborador}</div></div>
        <div class="row"><div class="label">Setor</div><div class="value">${setor}</div></div>
        <div class="row"><div class="label">Data da H.E</div><div class="value">${dataHe}</div></div>
        <div class="row"><div class="label">Horário de Início</div><div class="value">${inicio}</div></div>
        <div class="row"><div class="label">Horário de Término</div><div class="value">${fim}</div></div>
        <div class="row"><div class="label">Total de Horas</div><div class="value">${total}</div></div>
        <div class="row"><div class="label">Finalidade</div><div class="value"></div></div>
        <div class="finalidade">${finalidade || '—'}</div>
      </div>
    </div>

    <div class="sign">
      <div class="row"><div class="label">Data</div><div class="value">____ / ____ / ______</div></div>
      <div class="line"><div class="lbl">Ciência do Colaborador</div><div class="sig"></div></div>
      <div class="line"><div class="lbl">Aprovação Diretoria</div><div class="sig"></div></div>
    </div>

    <div class="footerNote">Documento gerado automaticamente pelo sistema.</div>
  </body>
</html>`;
}

async function generateRhSolicitacaoHePdf(row) {
  const outPath = getRhSolicitacaoHePdfPath(row.numero_documento);
  const html = buildRhSolicitacaoHePdfHtml(row);
  const win = new BrowserWindow({
    show: false,
    width: 900,
    height: 1100,
    webPreferences: {
      sandbox: true,
      contextIsolation: true
    }
  });

  try {
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    const pdfBuffer = await win.webContents.printToPDF({
      printBackground: true,
      pageSize: 'A4',
      marginsType: 0
    });
    fs.writeFileSync(outPath, pdfBuffer);
    return outPath;
  } finally {
    try {
      win.close();
    } catch (error) {
      // ignore
    }
  }
}

function listRhSolicitacoesHe({ username, filters }) {
  const ctx = ensureRhAccess(String(username || '').trim(), { requireEdit: false });
  const { workbook } = readRhDb();
  const { rows } = getRhSolicitacoesHeSheet(workbook);

  const safe = (val) => String(val || '').trim();
  const safeLower = (val) => safe(val).toLowerCase();
  const canEdit = !!ctx.canEdit;

  const f = filters && typeof filters === 'object' ? filters : {};
  const filtroColabId = safe(f.colaborador_id || f.colaboradorId);
  const filtroSetor = safeLower(f.setor);
  const filtroSolicitante = safeLower(f.solicitante);
  const filtroDoc = safeLower(f.numero_documento || f.numeroDocumento);
  const filtroIni = safe(f.data_ini || f.dataIni);
  const filtroFim = safe(f.data_fim || f.dataFim);

  const visible = rows
    .filter((r) => !safe(r.DELETED_AT))
    .filter((r) => canEdit || normalizeUsername(r.SOLICITANTE) === normalizeUsername(ctx.username))
    .filter((r) => !filtroColabId || safe(r.COLABORADOR_ID) === filtroColabId)
    .filter((r) => !filtroSetor || safeLower(r.SETOR).includes(filtroSetor))
    .filter((r) => !filtroSolicitante || safeLower(r.SOLICITANTE).includes(filtroSolicitante))
    .filter((r) => !filtroDoc || safeLower(r.NUMERO_DOCUMENTO).includes(filtroDoc))
    .filter((r) => {
      const d = safe(r.DATA_HE);
      if (filtroIni && d < filtroIni) return false;
      if (filtroFim && d > filtroFim) return false;
      return true;
    })
    .map((r) => ({
      id: safe(r.ID),
      numero_documento: safe(r.NUMERO_DOCUMENTO),
      data_solicitacao: safe(r.DATA_SOLICITACAO),
      data_he: safe(r.DATA_HE),
      colaborador_id: safe(r.COLABORADOR_ID),
      nome_colaborador: safe(r.NOME_COLABORADOR),
      setor: safe(r.SETOR),
      finalidade: safe(r.FINALIDADE),
      hora_inicio: safe(r.HORA_INICIO),
      hora_fim: safe(r.HORA_FIM),
      total_horas: Number(r.TOTAL_HORAS || 0),
      solicitante: safe(r.SOLICITANTE),
      data_criacao: safe(r.DATA_CRIACAO),
      ultima_atualizacao: safe(r.ULTIMA_ATUALIZACAO)
    }))
    .sort((a, b) => String(b.data_solicitacao).localeCompare(String(a.data_solicitacao)) || String(b.numero_documento).localeCompare(String(a.numero_documento)));

  return { ok: true, canEdit, rows: visible.slice(0, 1000) };
}

function deleteRhSolicitacaoHe({ username, id }) {
  const ctx = ensureRhAccess(String(username || '').trim(), { requireEdit: false });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();

  const { header, rows } = getRhSolicitacoesHeSheet(workbook);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => String(r.ID || '').trim() === safeId) || null;
  if (!target) {
    throw new Error('Solicitação não encontrada.');
  }
  if (!canEditRhSolicitacaoRow(ctx, target)) {
    throw new Error('Você não tem permissão para excluir esta solicitação.');
  }

  const before = { ...target };
  target.DELETED_AT = now;
  target.ULTIMA_ATUALIZACAO = now;
  writeSheetAoa(workbook, 'SOLICITACOES_HE', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: 'SOLICITACOES_HE_DELETE',
    entity: 'SOLICITACOES_HE',
    entityId: safeId,
    before,
    after: target
  });
  writeWorkbookFile(workbook, filePath);

  const numero = String(target.NUMERO_DOCUMENTO || '').trim();
  if (numero) {
    const pdfPath = getRhSolicitacaoHePdfPath(numero);
    try {
      if (fs.existsSync(pdfPath)) {
        fs.unlinkSync(pdfPath);
      }
    } catch (error) {
      // ignore (não bloqueia exclusão do registro)
    }
  }

  return { ok: true };
}

async function regenRhSolicitacaoHePdf({ username, id }) {
  const ctx = ensureRhAccess(String(username || '').trim(), { requireEdit: false });
  const { workbook } = readRhDb();
  const { rows } = getRhSolicitacoesHeSheet(workbook);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => !String(r.DELETED_AT || '').trim() && String(r.ID || '').trim() === safeId) || null;
  if (!target) {
    throw new Error('Solicitação não encontrada.');
  }
  if (!canEditRhSolicitacaoRow(ctx, target)) {
    throw new Error('Você não tem permissão para gerar o PDF desta solicitação.');
  }

  const row = {
    id: String(target.ID || '').trim(),
    numero_documento: String(target.NUMERO_DOCUMENTO || '').trim(),
    data_solicitacao: String(target.DATA_SOLICITACAO || '').trim(),
    data_he: String(target.DATA_HE || '').trim(),
    colaborador_id: String(target.COLABORADOR_ID || '').trim(),
    nome_colaborador: String(target.NOME_COLABORADOR || '').trim(),
    setor: String(target.SETOR || '').trim(),
    finalidade: String(target.FINALIDADE || '').trim(),
    hora_inicio: String(target.HORA_INICIO || '').trim(),
    hora_fim: String(target.HORA_FIM || '').trim(),
    total_horas: Number(target.TOTAL_HORAS || 0),
    solicitante: String(target.SOLICITANTE || '').trim(),
    data_criacao: String(target.DATA_CRIACAO || '').trim(),
    ultima_atualizacao: String(target.ULTIMA_ATUALIZACAO || '').trim()
  };

  const filePath = await generateRhSolicitacaoHePdf(row);
  return { ok: true, filePath };
}

async function openRhSolicitacaoHePdf({ username, id }) {
  const ctx = ensureRhAccess(String(username || '').trim(), { requireEdit: false });
  const { workbook } = readRhDb();
  const { rows } = getRhSolicitacoesHeSheet(workbook);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => !String(r.DELETED_AT || '').trim() && String(r.ID || '').trim() === safeId) || null;
  if (!target) {
    throw new Error('Solicitação não encontrada.');
  }
  if (!canEditRhSolicitacaoRow(ctx, target)) {
    throw new Error('Você não tem permissão para visualizar esta solicitação.');
  }
  const numero = String(target.NUMERO_DOCUMENTO || '').trim();
  const pdfPath = getRhSolicitacaoHePdfPath(numero);
  if (!fs.existsSync(pdfPath)) {
    throw new Error('PDF não encontrado. Gere novamente.');
  }
  const opened = await shell.openPath(pdfPath);
  if (opened) {
    throw new Error(opened);
  }
  return { ok: true, filePath: pdfPath };
}

function importRhPontoInterno({ username, date, workbookPath }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const safeDate = String(date || '').trim();
  if (!safeDate || !/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
    throw new Error('Data inválida. Use YYYY-MM-DD.');
  }

  const safePath = String(workbookPath || '').trim();
  if (!safePath) {
    throw new Error('Selecione a planilha base.');
  }
  if (!fs.existsSync(safePath)) {
    throw new Error(`Planilha base não encontrada: ${safePath}`);
  }

  const { workbook: baseWb } = readWorkbookCached(safePath);
  const baseSheets = Array.isArray(baseWb.SheetNames) ? baseWb.SheetNames : [];
  if (!baseSheets.length) {
    throw new Error('Planilha base sem abas válidas.');
  }

  const { workbook: dbWb, filePath } = readRhDb();
  const colaboradores = listRhColaboradores({ ...ctx, allowedSectorId: '' }, dbWb);
  const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));

  const aoa = readSheetAoa(dbWb, 'HORAS_EXTRAS');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'TIPO_HORA',
    'OBSERVACAO',
    'JUSTIFICATIVA',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const atrasosAoa = readSheetAoa(dbWb, 'ATRASOS');
  const atrasosHeader = getHeaderFromAoa(atrasosAoa, [
    'ID',
    'COLABORADOR_ID',
    'DATA',
    'QUANTIDADE_HORAS',
    'OBSERVACAO',
    'CRIADO_POR',
    'DATA_REGISTRO',
    'UPDATED_AT',
    'DELETED_AT'
  ]);
  const atrasoRows = mapAoaToObjects([atrasosHeader, ...(atrasosAoa.slice(1) || [])]);

  const tipoHora = 'ponto_interno';
  const byKey = new Map();
  rows.forEach((r) => {
    if (String(r.DELETED_AT || '').trim()) {
      return;
    }
    const key = `${String(r.COLABORADOR_ID || '').trim()}||${String(r.DATA || '').trim()}||${String(r.TIPO_HORA || '').trim()}`;
    if (!key.startsWith('||')) {
      byKey.set(key, r);
    }
  });

  const atrasosByKey = new Map();
  atrasoRows.forEach((r) => {
    if (String(r.DELETED_AT || '').trim()) {
      return;
    }
    const key = `${String(r.COLABORADOR_ID || '').trim()}||${String(r.DATA || '').trim()}`;
    if (!key.startsWith('||')) {
      atrasosByKey.set(key, r);
    }
  });

  const now = new Date().toISOString();
  const imported = [];
  const updated = [];
  const skipped = [];
  const unmatchedSheets = [];
  const importedAtrasos = [];
  const updatedAtrasos = [];
  const skippedAtrasos = [];
  const unmatchedSheetsAtrasos = [];

  baseSheets.forEach((sheetName) => {
    const ws = baseWb.Sheets[sheetName];
    if (!ws) {
      return;
    }

    const { match } = matchColaboradorBySheetName(colaboradores, sheetName);
    const colaboradorId = (match && match.id && collabById.has(String(match.id))) ? String(match.id) : '';

    const heCell = ws.S2 || null;
    const horas = parseHeHoursFromCell(heCell);
    if (!Number.isFinite(horas) || horas <= 0) {
      skipped.push({ sheetName, horas });
    } else if (!colaboradorId) {
      unmatchedSheets.push({ sheetName, horas });
    } else {
      const key = `${colaboradorId}||${safeDate}||${tipoHora}`;
      const existing = byKey.get(key) || null;
      const obs = `Importado (Ponto interno) • ${path.basename(safePath)}`;
      if (existing) {
        existing.QUANTIDADE_HORAS = Number(horas);
        existing.OBSERVACAO = obs;
        existing.UPDATED_AT = now;
        updated.push({ sheetName, colaboradorId, horas });
      } else {
        const row = {
          ID: String(generateRhId()),
          COLABORADOR_ID: colaboradorId,
          DATA: safeDate,
          QUANTIDADE_HORAS: Number(horas),
          TIPO_HORA: tipoHora,
          OBSERVACAO: obs,
          JUSTIFICATIVA: '',
          CRIADO_POR: ctx.username,
          DATA_REGISTRO: now,
          UPDATED_AT: now,
          DELETED_AT: ''
        };
        rows.push(row);
        byKey.set(key, row);
        imported.push({ sheetName, colaboradorId, horas });
      }
    }

    const atrasoCell = ws.R2 || null;
    const atrasoHoras = parseRhAtrasoHoursFromCell(atrasoCell);
    if (!Number.isFinite(atrasoHoras) || atrasoHoras <= 0) {
      skippedAtrasos.push({ sheetName, atrasoHoras });
      return;
    }
    if (!colaboradorId) {
      unmatchedSheetsAtrasos.push({ sheetName, atrasoHoras });
      return;
    }

    const atrasoKey = `${colaboradorId}||${safeDate}`;
    const existingAtraso = atrasosByKey.get(atrasoKey) || null;
    const obsAtraso = `Importado (Atrasos) • ${path.basename(safePath)}`;
    if (existingAtraso) {
      existingAtraso.QUANTIDADE_HORAS = Number(atrasoHoras);
      existingAtraso.OBSERVACAO = obsAtraso;
      existingAtraso.UPDATED_AT = now;
      updatedAtrasos.push({ sheetName, colaboradorId, atrasoHoras });
    } else {
      const row = {
        ID: String(generateRhId()),
        COLABORADOR_ID: colaboradorId,
        DATA: safeDate,
        QUANTIDADE_HORAS: Number(atrasoHoras),
        OBSERVACAO: obsAtraso,
        CRIADO_POR: ctx.username,
        DATA_REGISTRO: now,
        UPDATED_AT: now,
        DELETED_AT: ''
      };
      atrasoRows.push(row);
      atrasosByKey.set(atrasoKey, row);
      importedAtrasos.push({ sheetName, colaboradorId, atrasoHoras });
    }
  });

  if (!imported.length && !updated.length && !importedAtrasos.length && !updatedAtrasos.length) {
    return {
      ok: true,
      imported: 0,
      updated: 0,
      skipped: skipped.length,
      unmatched: unmatchedSheets.length,
      importedAtrasos: 0,
      updatedAtrasos: 0,
      skippedAtrasos: skippedAtrasos.length,
      unmatchedAtrasos: unmatchedSheetsAtrasos.length,
      details: {
        unmatchedSheets: unmatchedSheets.slice(0, 20),
        unmatchedSheetsAtrasos: unmatchedSheetsAtrasos.slice(0, 20)
      }
    };
  }

  if (imported.length || updated.length) {
    writeSheetAoa(dbWb, 'HORAS_EXTRAS', objectsToAoa(header, rows));
  }
  if (importedAtrasos.length || updatedAtrasos.length) {
    writeSheetAoa(dbWb, 'ATRASOS', objectsToAoa(atrasosHeader, atrasoRows));
  }
  appendRhAudit(dbWb, {
    username: ctx.username,
    action: 'IMPORT_PONTO_INTERNO',
    entity: 'HORAS_EXTRAS',
    entityId: safeDate,
    before: null,
    after: {
      date: safeDate,
      sourceFile: safePath,
      imported: imported.length,
      updated: updated.length,
      skipped: skipped.length,
      unmatched: unmatchedSheets.length,
      importedAtrasos: importedAtrasos.length,
      updatedAtrasos: updatedAtrasos.length,
      skippedAtrasos: skippedAtrasos.length,
      unmatchedAtrasos: unmatchedSheetsAtrasos.length,
      unmatchedSheets: unmatchedSheets.slice(0, 30).map((x) => x.sheetName),
      unmatchedSheetsAtrasos: unmatchedSheetsAtrasos.slice(0, 30).map((x) => x.sheetName)
    }
  });
  writeWorkbookFile(dbWb, filePath);

  return {
    ok: true,
    imported: imported.length,
    updated: updated.length,
    skipped: skipped.length,
    unmatched: unmatchedSheets.length,
    importedAtrasos: importedAtrasos.length,
    updatedAtrasos: updatedAtrasos.length,
    skippedAtrasos: skippedAtrasos.length,
    unmatchedAtrasos: unmatchedSheetsAtrasos.length,
    details: {
      unmatchedSheets: unmatchedSheets.slice(0, 20),
      unmatchedSheetsAtrasos: unmatchedSheetsAtrasos.slice(0, 20)
    }
  };
}

function deleteRhHoraExtra({ username, id }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'HORAS_EXTRAS');
  const header = getHeaderFromAoa(aoa, ['ID', 'COLABORADOR_ID', 'DATA', 'QUANTIDADE_HORAS', 'TIPO_HORA', 'OBSERVACAO', 'JUSTIFICATIVA', 'CRIADO_POR', 'DATA_REGISTRO', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const safeId = String(id || '').trim();
  const target = rows.find((r) => String(r.ID || '').trim() === safeId);
  if (!target) {
    throw new Error('Lançamento não encontrado.');
  }
  const before = { ...target };
  target.DELETED_AT = now;
  target.UPDATED_AT = now;
  writeSheetAoa(workbook, 'HORAS_EXTRAS', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: 'HORAS_EXTRAS_DELETE', entity: 'HORAS_EXTRAS', entityId: safeId, before, after: target });
  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function upsertRhUserRule({ username, target, role, setor_id }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const safeTarget = normalizeUsername(target);
  if (!safeTarget) {
    throw new Error('Usuário alvo obrigatório.');
  }
  const safeRole = normalizeRhRole(role);
  const safeSetor = String(setor_id || '').trim();
  if (safeRole === 'GESTOR' && !safeSetor) {
    throw new Error('Gestor precisa de setor.');
  }

  const aoa = readSheetAoa(workbook, 'USERS_RH');
  const header = getHeaderFromAoa(aoa, ['USERNAME', 'ROLE', 'SETOR_ID', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const existing = rows.find((r) => normalizeUsername(r.USERNAME) === safeTarget) || null;
  const before = existing ? { ...existing } : null;
  const targetRow = existing || { USERNAME: safeTarget };
  targetRow.ROLE = safeRole;
  targetRow.SETOR_ID = safeSetor;
  targetRow.UPDATED_AT = now;
  targetRow.DELETED_AT = '';
  if (!existing) {
    rows.push(targetRow);
  }

  writeSheetAoa(workbook, 'USERS_RH', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: before ? 'USERS_RH_UPDATE' : 'USERS_RH_CREATE', entity: 'USERS_RH', entityId: safeTarget, before, after: targetRow });
  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function deleteRhUserRule({ username, target }) {
  const ctx = ensureRhAccess(username, { requireEdit: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const safeTarget = normalizeUsername(target);
  if (!safeTarget) {
    throw new Error('Usuário alvo obrigatório.');
  }
  const aoa = readSheetAoa(workbook, 'USERS_RH');
  const header = getHeaderFromAoa(aoa, ['USERNAME', 'ROLE', 'SETOR_ID', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const targetRow = rows.find((r) => normalizeUsername(r.USERNAME) === safeTarget);
  if (!targetRow) {
    throw new Error('Usuário RH não encontrado.');
  }
  const before = { ...targetRow };
  targetRow.DELETED_AT = now;
  targetRow.UPDATED_AT = now;
  writeSheetAoa(workbook, 'USERS_RH', objectsToAoa(header, rows));
  appendRhAudit(workbook, { username: ctx.username, action: 'USERS_RH_DELETE', entity: 'USERS_RH', entityId: safeTarget, before, after: targetRow });
  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function normalizeRequestStatus(value) {
  const text = String(value || '').trim().toUpperCase();
  if (text === 'APROVADO' || text === 'NEGADO') {
    return text;
  }
  return 'PENDENTE';
}

function listRhAccessRequests(workbook) {
  const objs = readSheetObjects(workbook, 'ACCESS_REQUESTS');
  return objs.map((row) => ({
    id: String(row.ID || '').trim(),
    username: normalizeUsername(row.USERNAME),
    requested_role: normalizeRhRole(row.REQUESTED_ROLE),
    requested_setor_id: String(row.REQUESTED_SETOR_ID || '').trim(),
    requested_at: String(row.REQUESTED_AT || '').trim(),
    status: normalizeRequestStatus(row.STATUS),
    decided_at: String(row.DECIDED_AT || '').trim(),
    decided_by: normalizeUsername(row.DECIDED_BY)
  })).filter((r) => r.id && r.username && r.requested_at);
}

function createRhAccessRequest({ username, requested_role, requested_setor_id }) {
  const ctx = ensureRhAccess(username, { requireEdit: false, allowUnconfigured: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();

  if (ctx.canEdit) {
    throw new Error('Usuário RH (edição) não precisa solicitar acesso.');
  }
  const role = normalizeRhRole(requested_role);
  if (role !== 'GESTOR') {
    throw new Error('Solicitação inválida (role).');
  }
  const setorId = String(requested_setor_id || '').trim();
  if (!setorId) {
    throw new Error('Selecione um setor para solicitar.');
  }

  const setores = listRhSetores({ ...ctx, allowedSectorId: '' }, workbook);
  if (!setores.some((s) => String(s.id) === setorId)) {
    throw new Error('Setor solicitado não encontrado.');
  }

  const requests = listRhAccessRequests(workbook);
  const pending = requests.find((r) => r.username === ctx.username && r.status === 'PENDENTE') || null;
  if (pending) {
    throw new Error('Já existe uma solicitação pendente para este usuário.');
  }

  const aoa = readSheetAoa(workbook, 'ACCESS_REQUESTS');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'USERNAME',
    'REQUESTED_ROLE',
    'REQUESTED_SETOR_ID',
    'REQUESTED_AT',
    'STATUS',
    'DECIDED_AT',
    'DECIDED_BY'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const row = {
    ID: String(generateRhId()),
    USERNAME: ctx.username,
    REQUESTED_ROLE: role,
    REQUESTED_SETOR_ID: setorId,
    REQUESTED_AT: now,
    STATUS: 'PENDENTE',
    DECIDED_AT: '',
    DECIDED_BY: ''
  };
  rows.push(row);
  writeSheetAoa(workbook, 'ACCESS_REQUESTS', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: 'ACCESS_REQUEST_CREATE',
    entity: 'ACCESS_REQUESTS',
    entityId: row.ID,
    before: null,
    after: row
  });
  writeWorkbookFile(workbook, filePath);
  return { ok: true, requestId: row.ID };
}

function decideRhAccessRequest({ username, requestId, decision }) {
  const ctx = ensureRhAccess(username, { requireEdit: true, allowUnconfigured: true });
  const { workbook, filePath } = readRhDb();
  const now = new Date().toISOString();
  const safeDecision = String(decision || '').trim().toUpperCase();
  if (safeDecision !== 'APROVADO' && safeDecision !== 'NEGADO') {
    throw new Error('Decisão inválida.');
  }
  const safeId = String(requestId || '').trim();
  if (!safeId) {
    throw new Error('ID da solicitação obrigatório.');
  }

  const aoa = readSheetAoa(workbook, 'ACCESS_REQUESTS');
  const header = getHeaderFromAoa(aoa, [
    'ID',
    'USERNAME',
    'REQUESTED_ROLE',
    'REQUESTED_SETOR_ID',
    'REQUESTED_AT',
    'STATUS',
    'DECIDED_AT',
    'DECIDED_BY'
  ]);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);
  const target = rows.find((r) => String(r.ID || '').trim() === safeId);
  if (!target) {
    throw new Error('Solicitação não encontrada.');
  }
  const before = { ...target };
  const status = normalizeRequestStatus(target.STATUS);
  if (status !== 'PENDENTE') {
    throw new Error('Solicitação já decidida.');
  }

  target.STATUS = safeDecision;
  target.DECIDED_AT = now;
  target.DECIDED_BY = ctx.username;

  writeSheetAoa(workbook, 'ACCESS_REQUESTS', objectsToAoa(header, rows));
  appendRhAudit(workbook, {
    username: ctx.username,
    action: 'ACCESS_REQUEST_DECIDE',
    entity: 'ACCESS_REQUESTS',
    entityId: safeId,
    before,
    after: target
  });

  if (safeDecision === 'APROVADO') {
    upsertRhUserRule({
      username: ctx.username,
      target: String(target.USERNAME || '').trim(),
      role: String(target.REQUESTED_ROLE || 'GESTOR').trim(),
      setor_id: String(target.REQUESTED_SETOR_ID || '').trim()
    });
  }

  writeWorkbookFile(workbook, filePath);
  return { ok: true };
}

function getFiscalRoot() {
  const settings = loadSettings();
  let fiscalRoot = String(settings.fiscalRoot || '').trim();
  if (!fiscalRoot) {
    throw new Error('Configure o Caminho Fiscal (R:) nas Configuracoes.');
  }
  // Em Windows, "R:" (sem barra) é relativo ao diretório atual do drive.
  // Normaliza para "R:\\".
  if (/^[a-zA-Z]:$/.test(fiscalRoot)) {
    fiscalRoot = `${fiscalRoot}\\`;
  }
  return fiscalRoot;
}

function resolveFiscalPath(relativeParts) {
  const fiscalRoot = getFiscalRoot();
  return path.join(fiscalRoot, ...relativeParts);
}

function parsePtBrDateToYmd(text) {
  const value = String(text || '').trim();
  if (!value) {
    return '';
  }
  const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!match) {
    return '';
  }
  const dd = String(match[1]).padStart(2, '0');
  const mm = String(match[2]).padStart(2, '0');
  let yyyy = Number(match[3]);
  if (match[3].length === 2) {
    yyyy = yyyy <= 50 ? 2000 + yyyy : 1900 + yyyy;
  }
  if (!Number.isFinite(yyyy) || yyyy < 1900 || yyyy > 2100) {
    return '';
  }
  return `${String(yyyy).padStart(4, '0')}-${mm}-${dd}`;
}

function parseRncSecondsValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Se veio como Date, usa apenas a parte de tempo.
    return (value.getHours() * 3600) + (value.getMinutes() * 60) + value.getSeconds();
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    // Excel pode armazenar duração como fração do dia.
    if (value > 0 && value < 1) {
      return Math.round(value * 86400);
    }
    // Caso já esteja em segundos.
    if (value >= 1) {
      return Math.round(value);
    }
    return 0;
  }

  const text = String(value || '').trim();
  if (!text) {
    return 0;
  }

  // HH:MM:SS ou MM:SS
  const parts = text.split(':').map((p) => p.trim()).filter(Boolean);
  if (parts.length === 3) {
    const h = parseNumericValue(parts[0]);
    const m = parseNumericValue(parts[1]);
    const s = parseNumericValue(parts[2]);
    return Math.max(0, Math.round((h * 3600) + (m * 60) + s));
  }
  if (parts.length === 2) {
    const m = parseNumericValue(parts[0]);
    const s = parseNumericValue(parts[1]);
    return Math.max(0, Math.round((m * 60) + s));
  }

  return Math.max(0, Math.round(parseNumericValue(text)));
}

function resolveRncWorkbookPath() {
  const fiscalRoot = getFiscalRoot();
  const fileName = 'Relatório de Não Conformidade BANCO DE DADOS.xlsm';
  const normalizedRoot = normalizeText(path.basename(fiscalRoot));

  const candidates = [
    resolveFiscalPath(FISCAL_RNC_RELATIVE),
    path.join(fiscalRoot, 'PartsSeals', 'Rnc Interna', fileName),
    path.join(fiscalRoot, 'Rnc Interna', fileName)
  ];

  // Quando o usuário configura direto em R:\PartsSeals, evita duplicar "PartsSeals".
  if (normalizedRoot === 'PARTSSEALS') {
    candidates.push(path.join(fiscalRoot, 'Rnc Interna', fileName));
  }

  // Quando o usuário configura direto em R:\Rede, tenta incluir PartsSeals.
  if (normalizedRoot === 'REDE') {
    candidates.push(path.join(fiscalRoot, 'PartsSeals', 'Rnc Interna', fileName));
  }

  for (const p of candidates) {
    if (p && fs.existsSync(p)) {
      return p;
    }
  }

  // Fallback: procura por diretório e arquivo pelo nome normalizado.
  const searchBases = [fiscalRoot, path.dirname(fiscalRoot)];
  for (const base of searchBases) {
    if (!base || !fs.existsSync(base)) {
      continue;
    }

    const partsSealsDir = findChildDirByNormalizedName(base, 'PartsSeals') || base;
    const rncDir = findChildDirByNormalizedName(partsSealsDir, 'Rnc Interna') || findChildDirByNormalizedName(base, 'Rnc Interna');
    if (!rncDir || !fs.existsSync(rncDir)) {
      continue;
    }

    const files = fs.readdirSync(rncDir, { withFileTypes: true }).filter((d) => d.isFile());
    const target = normalizeText('Relatório de Não Conformidade BANCO DE DADOS');
    const exact = files.find((f) => {
      if (path.extname(f.name).toLowerCase() !== '.xlsm') {
        return false;
      }
      return normalizeText(path.parse(f.name).name) === target;
    });
    if (exact) {
      return path.join(rncDir, exact.name);
    }

    const fuzzy = files.find((f) => {
      if (path.extname(f.name).toLowerCase() !== '.xlsm') {
        return false;
      }
      const baseName = normalizeText(path.parse(f.name).name);
      return baseName.includes('NAO CONFORMIDADE')
        && baseName.includes('BANCO')
        && baseName.includes('DADOS');
    });
    if (fuzzy) {
      return path.join(rncDir, fuzzy.name);
    }
  }

  return null;
}

function resolvePcpDashboardDailyRootPath() {
  const settings = loadSettings();
  const customPath = String(settings.pcpDashboardDailyRoot || '').trim();
  if (customPath) {
    if (!fs.existsSync(customPath)) {
      throw new Error(`Pasta do Dashboard PCP nao encontrada: ${customPath}`);
    }
    return customPath;
  }

  const fiscalRoot = getFiscalRoot();
  const normalizedRoot = normalizeText(path.basename(fiscalRoot));

  const candidatePaths = [
    path.join('R:\\', ...PCP_DASHBOARD_DAILY_RELATIVE),
    path.join(fiscalRoot, ...PCP_DASHBOARD_DAILY_RELATIVE),
    path.join(fiscalRoot, 'PartsSeals', 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'),
    path.join(fiscalRoot, 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'),
    path.join('R:\\', ...PCP_DASHBOARD_DAILY_RELATIVE_LEGACY),
    path.join(fiscalRoot, ...PCP_DASHBOARD_DAILY_RELATIVE_LEGACY),
    path.join(fiscalRoot, 'PartsSeals', 'Programação de Produção Diaria'),
    path.join(fiscalRoot, 'Programação de Produção Diaria')
  ];
  if (normalizedRoot === 'REDE') {
    candidatePaths.push(path.join(fiscalRoot, 'PartsSeals', 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'));
    candidatePaths.push(path.join(fiscalRoot, 'PartsSeals', 'Programação de Produção Diaria'));
  }
  if (normalizedRoot === 'PARTSSEALS') {
    candidatePaths.push(path.join(fiscalRoot, 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'));
    candidatePaths.push(path.join(fiscalRoot, 'Programação de Produção Diaria'));
  }

  for (const candidate of candidatePaths) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  const relativeOptions = [
    PCP_DASHBOARD_DAILY_RELATIVE,
    PCP_DASHBOARD_DAILY_RELATIVE_LEGACY,
    ['PartsSeals', 'PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'],
    ['PROGRAMAÇÃO DE PRODUÇÃO-DIARIA'],
    ['PartsSeals', 'Programação de Produção Diaria'],
    ['Programação de Produção Diaria']
  ];
  const searchBases = [fiscalRoot, path.dirname(fiscalRoot)];

  for (const basePath of searchBases) {
    if (!basePath || !fs.existsSync(basePath)) {
      continue;
    }
    for (const parts of relativeOptions) {
      let currentPath = basePath;
      let ok = true;
      for (const part of parts) {
        const nextPath = findChildDirByNormalizedName(currentPath, part);
        if (!nextPath) {
          ok = false;
          break;
        }
        currentPath = nextPath;
      }
      if (ok && fs.existsSync(currentPath)) {
        return currentPath;
      }
    }
  }

  throw new Error(`Nao encontrei a pasta do Dashboard PCP. Verifique o Caminho Fiscal (R:) ou configure "pcpDashboardDailyRoot".`);
}

function extractReqBase(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }
  const match = text.match(/^(\d+)\s*(?:\/.*)?$/);
  return match ? match[1] : text.split('/')[0];
}

function formatExcelDatePtBr(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString('pt-BR');
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      const dt = new Date(parsed.y, parsed.m - 1, parsed.d);
      return dt.toLocaleDateString('pt-BR');
    }
  }

  const text = String(value).trim();
  if (!text) {
    return '';
  }
  const asDate = new Date(text);
  if (!Number.isNaN(asDate.getTime())) {
    return asDate.toLocaleDateString('pt-BR');
  }
  return text;
}

function loadSettings() {
  const file = SETTINGS_FILE();
  if (!fs.existsSync(file)) {
    return { ...DEFAULT_SETTINGS };
  }

  try {
    const data = JSON.parse(fs.readFileSync(file, 'utf8'));
    return {
      ...DEFAULT_SETTINGS,
      ...data
    };
  } catch (error) {
    return { ...DEFAULT_SETTINGS };
  }
}

function saveSettings(nextSettings) {
  const file = SETTINGS_FILE();
  const current = loadSettings();
  const merged = {
    ...current,
    ...nextSettings
  };

  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(merged, null, 2), 'utf8');
  return merged;
}

function normalizeTrackingCode(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9]/g, '');
}

function looksLikeCorreiosTrackingCode(code) {
  if (/^[A-Z]{2}\d{9}[A-Z]{2}$/.test(code)) {
    return true;
  }
  return /^[A-Z0-9]{8,32}$/.test(code);
}

function findFiscalInfoByTrackingCode(code) {
  const normalizedCode = normalizeTrackingCode(code);
  if (!normalizedCode) {
    return null;
  }

  try {
    const { rows } = readBancoPedidosRowsLite();
    const matches = rows.filter((row) => normalizeTrackingCode(row.rastreio) === normalizedCode);
    if (!matches.length) {
      return null;
    }

    const nfSet = new Set(matches.map((m) => String(m.nf || '').trim()).filter(Boolean));
    const clienteSet = new Set(matches.map((m) => String(m.cliente || '').trim()).filter(Boolean));
    const pedidoSet = new Set(matches.map((m) => String(m.pedido || '').trim()).filter(Boolean));

    const nfs = Array.from(nfSet);
    const clientes = Array.from(clienteSet);
    const pedidos = Array.from(pedidoSet);
    nfs.sort((a, b) => String(b).localeCompare(String(a), 'pt-BR', { numeric: true }));
    clientes.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
    pedidos.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));

    return {
      nf: nfs[0] || '',
      cliente: clientes[0] || '',
      pedidos: pedidos.slice(0, 8),
      nfs,
      clientes
    };
  } catch (error) {
    return null;
  }
}

async function httpsJson(url, options = {}) {
  if (typeof fetch === 'function') {
    try {
      const response = await fetch(url, options);
      const text = await response.text();
      let data = null;
      try {
        data = text ? JSON.parse(text) : null;
      } catch (error) {
        data = null;
      }
      return { ok: response.ok, status: response.status, data, text };
    } catch (error) {
      // fallback for environments where undici/fetch fails (proxy/DNS/TLS)
    }
  }

  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const requestOptions = {
      method: options.method || 'GET',
      hostname: parsed.hostname,
      path: `${parsed.pathname}${parsed.search}`,
      headers: options.headers || {}
    };

    const req = https.request(requestOptions, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        let data = null;
        try {
          data = text ? JSON.parse(text) : null;
        } catch (error) {
          data = null;
        }
        resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, data, text });
      });
    });

    req.on('error', reject);
    req.end();
  });
}

async function trackSeurastreio(code, apiKey) {
  const url = `https://seurastreio.com.br/api/public/rastreio/${encodeURIComponent(code)}`;
  const headers = {
    Accept: 'application/json',
    Authorization: `Bearer ${apiKey}`
  };

  const result = await httpsJson(url, { method: 'GET', headers });
  if (result.status === 404) {
    return { valid: false, code, provider: 'seurastreio' };
  }
  if (!result.ok) {
    const detail = result.data && (result.data.message || result.data.error || result.data.detail);
    throw new Error(`Erro ao consultar rastreio (status ${result.status}). ${detail || ''}`.trim());
  }

  const rastreio = result.data && (result.data.rastreio || result.data);
  const last = rastreio && (rastreio.eventoMaisRecente || rastreio.ultimoEvento || rastreio.lastEvent);
  const status = (last && (last.descricao || last.status || last.evento)) || (rastreio && rastreio.status) || '';
  const local = (last && (last.local || last.unidade || last.cidade)) || '';
  const destino = (last && (last.destino || last.destinatario || last.para || last.enderecoDestino)) || (rastreio && rastreio.destino) || '';
  const updatedAt = (last && (last.data || last.dataHora || last.data_hora)) || '';

  return {
    valid: true,
    code,
    provider: 'seurastreio',
    status: String(status || '').trim() || 'Sem status',
    local: String(local || '').trim(),
    destino: String(destino || '').trim(),
    updatedAt: String(updatedAt || '').trim()
  };
}

function toPtBrMonthName(date) {
  return date.toLocaleString('pt-BR', { month: 'long' }).toUpperCase();
}

function findChildDirByNormalizedName(parent, expectedName) {
  if (!fs.existsSync(parent)) {
    return null;
  }

  const target = normalizeText(expectedName);
  const dirs = fs.readdirSync(parent, { withFileTypes: true }).filter((d) => d.isDirectory());
  const found = dirs.find((d) => normalizeText(d.name) === target);
  return found ? path.join(parent, found.name) : null;
}

function resolveBaseListPath(rootFolder, listTypeConfig) {
  const direct = path.join(rootFolder, ...listTypeConfig.relativePath);
  if (fs.existsSync(direct)) {
    return direct;
  }

  let currentPath = rootFolder;
  for (const part of listTypeConfig.relativePath) {
    const nextPath = findChildDirByNormalizedName(currentPath, part);
    if (!nextPath) {
      return null;
    }
    currentPath = nextPath;
  }

  return currentPath;
}

function resolveMonthFolder(yearFolder, monthName) {
  const direct = path.join(yearFolder, monthName);
  if (fs.existsSync(direct)) {
    return direct;
  }

  const normalized = findChildDirByNormalizedName(yearFolder, monthName);
  return normalized;
}

function resolveListFile(monthFolder, filePrefix, day, month) {
  const dd = String(day).padStart(2, '0');
  const mm = String(month).padStart(2, '0');
  const expectedCore = `${normalizeText(filePrefix)} ${dd}-${mm}`;

  const files = fs.readdirSync(monthFolder, { withFileTypes: true }).filter((d) => d.isFile());
  const allowedExt = new Set(['.xlsx', '.xls', '.xlsm', '.xlsb']);

  const exact = files.find((f) => {
    const ext = path.extname(f.name).toLowerCase();
    if (!allowedExt.has(ext)) {
      return false;
    }

    const base = normalizeText(path.parse(f.name).name);
    return base === expectedCore;
  });

  if (exact) {
    return path.join(monthFolder, exact.name);
  }

  const fuzzy = files.find((f) => {
    const ext = path.extname(f.name).toLowerCase();
    if (!allowedExt.has(ext)) {
      return false;
    }

    const base = normalizeText(path.parse(f.name).name);
    return base.includes(`${dd}-${mm}`) && base.includes(normalizeText(filePrefix));
  });

  return fuzzy ? path.join(monthFolder, fuzzy.name) : null;
}

function getCellDisplayValue(worksheet, row, col) {
  const cell = getCell(worksheet, row, col);
  if (!cell) {
    return '';
  }

  if (cell.w !== undefined && cell.w !== null) {
    return String(cell.w).trim();
  }

  if (cell.v === undefined || cell.v === null) {
    return '';
  }

  return String(cell.v).trim();
}

function normalizeColorHex(value) {
  const onlyHex = String(value || '').replace(/[^0-9a-f]/gi, '').toUpperCase();
  if (onlyHex.length === 8) {
    return onlyHex.slice(2);
  }
  if (onlyHex.length === 6) {
    return onlyHex;
  }
  return '';
}

function extractCellFillHex(cell) {
  if (!cell || !cell.s) {
    return '';
  }

  const style = cell.s;
  const directFg = normalizeColorHex(style.fgColor && style.fgColor.rgb);
  if (directFg) {
    return directFg;
  }

  const directBg = normalizeColorHex(style.bgColor && style.bgColor.rgb);
  if (directBg) {
    return directBg;
  }

  const fillFg = normalizeColorHex(style.fill && style.fill.fgColor && style.fill.fgColor.rgb);
  if (fillFg) {
    return fillFg;
  }

  const fillBg = normalizeColorHex(style.fill && style.fill.bgColor && style.fill.bgColor.rgb);
  if (fillBg) {
    return fillBg;
  }

  return '';
}

function hasStandardHeader(sheet) {
  const col0 = normalizeText(getCellDisplayValue(sheet, 6, 0));
  const col1 = normalizeText(getCellDisplayValue(sheet, 6, 1));
  const col2 = normalizeText(getCellDisplayValue(sheet, 6, 2));
  const col9 = normalizeText(getCellDisplayValue(sheet, 6, 9));

  return col0 === 'LINHA' && col1 === 'QTD' && col2.includes('ITENS') && col9 === 'OBS';
}

function hasMoldagemHeader(sheet) {
  const col0 = normalizeText(getCellDisplayValue(sheet, 7, 0));
  const col1 = normalizeText(getCellDisplayValue(sheet, 7, 1));
  const col2 = normalizeText(getCellDisplayValue(sheet, 7, 2));
  const col13 = normalizeText(getCellDisplayValue(sheet, 7, 13));

  return col0 === 'COD' && col1 === 'QTD' && col2.includes('DESCRICAO') && col13 === 'STATUS';
}

function findWorksheetWithHeader(workbook, listType, config) {
  if (listType === 'moldagem') {
    const preferred = workbook.Sheets[config.preferredSheetName];
    if (preferred && hasMoldagemHeader(preferred)) {
      return { sheetName: config.preferredSheetName, sheet: preferred };
    }
  }

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matches = listType === 'moldagem' ? hasMoldagemHeader(sheet) : hasStandardHeader(sheet);
    if (matches) {
      return { sheetName, sheet };
    }
  }

  return null;
}

function parseStandardTableRows(sheet, listType, columns, machineValue) {
  const rows = [];
  let rowIndex = 7;
  let consecutiveEmptyLinha = 0;

  while (true) {
    const values = [];
    for (let col = 0; col < columns.length; col += 1) {
      values.push(getCellDisplayValue(sheet, rowIndex, col));
    }

    const linhaValue = values[0];
    if (linhaValue === '') {
      consecutiveEmptyLinha += 1;
      if (consecutiveEmptyLinha >= 2) {
        break;
      }
      rowIndex += 1;
      continue;
    }

    consecutiveEmptyLinha = 0;

    const row = {};
    columns.forEach((key, index) => {
      row[key] = values[index];
    });

    const qtdCell = getCell(sheet, rowIndex, 1);
    const qtdRaw = qtdCell && qtdCell.v !== undefined && qtdCell.v !== null ? String(qtdCell.v).trim() : '';
    row.__qtdIsTextNumeric = !!qtdCell
      && qtdCell.t === 's'
      && /^\s*'?\d+(?:[.,]\d+)?\s*$/.test(qtdRaw);

    row.__isQtdDivider = normalizeText(row.QTD) === 'QTD';

    if (listType === 'projeto') {
      const qtdFillHex = extractCellFillHex(qtdCell);
      const normalizedObs = normalizeText(row.OBS);
      const isProjetoLiberadoObs = normalizedObs.includes('PROJETO LIBERADO') || normalizedObs === 'PRONTO';
      row.__highlight = PROJECT_READY_QTD_COLORS.has(qtdFillHex) || isProjetoLiberadoObs;
      if (row.__highlight && !row.__isQtdDivider) {
        row.OBS = 'PROJETO LIBERADO';
      }
    } else {
      row.__highlight = READY_OBS_VALUES.has(normalizeText(row.OBS));
    }

    const shouldIncludeRow = !machineValue
      || row.__isQtdDivider
      || normalizeText(row['MAQ./OPER.']) === normalizeText(machineValue);

    if (shouldIncludeRow) {
      rows.push(row);
    }

    rowIndex += 1;
  }

  return rows;
}

function parseMoldagemRows(sheet) {
  const rows = [];
  let rowIndex = 8;
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  const maxRow = range ? range.e.r + 10 : 50000;

  while (rowIndex <= maxRow) {
    const codValue = normalizeText(getCellDisplayValue(sheet, rowIndex, 0));
    if (codValue === 'LIMITE') {
      break;
    }

    const values = [];
    for (let col = 0; col < MOLDAGEM_COLUMNS.length; col += 1) {
      values.push(getCellDisplayValue(sheet, rowIndex, col));
    }

    const isEmptyRow = values.every((value) => normalizeText(value) === '');
    if (isEmptyRow) {
      rowIndex += 1;
      continue;
    }

    const row = {};
    MOLDAGEM_COLUMNS.forEach((key, index) => {
      row[key] = values[index];
    });
    row.MOLDES = String(row.MOLDES || '').replace(/Ã/g, 'Ø');
    row.EIXO = String(row.EIXO || '').replace(/Ã/g, 'Ø');

    row.__isQtdDivider = false;
    const normalizedStatus = normalizeText(row.STATUS);
    if (normalizedStatus === MOLDAGEM_STATUS_ESTOQUE) {
      row.__rowClass = 'status-estoque';
    } else if (normalizedStatus === MOLDAGEM_STATUS_PRENSADO) {
      row.__rowClass = 'status-prensado';
    }

    rows.push(row);
    rowIndex += 1;
  }

  return rows;
}

function parseNumericValue(value) {
  const raw = String(value || '').trim().replace(/^'+/, '');
  if (!raw) {
    return 0;
  }

  let normalized = raw.replace(/\s+/g, '');
  const hasComma = normalized.includes(',');
  const hasDot = normalized.includes('.');

  if (hasComma && hasDot) {
    normalized = normalized.replace(/\./g, '').replace(',', '.');
  } else if (hasComma) {
    normalized = normalized.replace(',', '.');
  }

  const num = Number(normalized);
  if (!Number.isFinite(num)) {
    return 0;
  }

  return num;
}

function parseProducedPiecesFromObs(obsValue, plannedQtd) {
  const normalizedObs = normalizeText(obsValue);
  if (normalizedObs === 'PRONTO') {
    return plannedQtd;
  }

  const partialMatch = normalizedObs.match(/PARCIAL\s*[-:]\s*(\d+(?:[.,]\d+)?)/);
  if (!partialMatch) {
    return 0;
  }

  const partialQtd = parseNumericValue(partialMatch[1]);
  if (!Number.isFinite(partialQtd) || partialQtd <= 0) {
    return 0;
  }

  return partialQtd;
}

function isReadyPieceQtdValue(qtdValue) {
  const raw = String(qtdValue || '').trim();
  return /^'\s*\d/.test(raw);
}

function buildDuplicatedReqSet(rows, reqKey) {
  const reqCounter = new Map();
  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    const normalizedReq = normalizeText(row[reqKey]);
    if (!normalizedReq) {
      return;
    }

    reqCounter.set(normalizedReq, (reqCounter.get(normalizedReq) || 0) + 1);
  });

  return new Set(
    Array.from(reqCounter.entries())
      .filter(([, count]) => count > 1)
      .map(([req]) => req)
  );
}

function isCriticalByText(descriptionValue, obsValue) {
  const normalizedDescription = normalizeText(descriptionValue);
  const normalizedObs = normalizeText(obsValue);
  const hasCriticalInDescription = normalizedDescription.includes('C/ COR') || normalizedDescription.includes('RANHURA');
  const hasCriticalInObs = normalizedObs.includes('CORTE') || normalizedObs.includes('RANHURA');
  return hasCriticalInDescription || hasCriticalInObs;
}

function calculateMachineKpis(rows, machineValue) {
  const normalizedMachine = normalizeText(machineValue);
  let totalOrdens = 0;
  let ordensFeitas = 0;
  let totalPecasPlanejadas = 0;
  let totalPecasProduzidas = 0;

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    if (normalizeText(row['MAQ./OPER.']) !== normalizedMachine) {
      return;
    }

    totalOrdens += 1;
    const qtd = parseNumericValue(row.QTD);
    totalPecasPlanejadas += qtd;

    if (normalizeText(row.OBS) === 'PRONTO') {
      ordensFeitas += 1;
    }

    totalPecasProduzidas += parseProducedPiecesFromObs(row.OBS, qtd);
  });

  const eficienciaOrdens = totalOrdens > 0 ? (ordensFeitas / totalOrdens) * 100 : 0;
  const eficienciaPecas = totalPecasPlanejadas > 0 ? (totalPecasProduzidas / totalPecasPlanejadas) * 100 : 0;

  return {
    totalOrdens,
    ordensFeitas,
    totalPecasPlanejadas,
    totalPecasProduzidas,
    eficienciaOrdens,
    eficienciaPecas
  };
}

function calculateMasterKpis(rows) {
  let totalOrdens = 0;
  let ordensFeitas = 0;
  let ordensProntas = 0;
  let totalPecasPlanejadas = 0;
  let totalPecasFeitas = 0;
  let totalPecasProntas = 0;
  let ordensTeflon = 0;
  let totalPecasTeflon = 0;
  let ordensCriticas = 0;
  let pecasCriticas = 0;
  const duplicatedReqs = buildDuplicatedReqSet(rows, 'N REQ');

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    totalOrdens += 1;
    const qtd = parseNumericValue(row.QTD);
    totalPecasPlanejadas += qtd;
    if (row.__isTeflon) {
      ordensTeflon += 1;
      totalPecasTeflon += qtd;
    }

    const hasDuplicatedReq = normalizeText(row['N REQ']) !== '' && duplicatedReqs.has(normalizeText(row['N REQ']));
    const isCritical = isCriticalByText(row['DESCRIÇÃO'], row.OBS) || hasDuplicatedReq;
    if (isCritical) {
      ordensCriticas += 1;
      pecasCriticas += qtd;
    }

    const hasMachineDown = normalizeText(row['BAIXA MAQUINA']) !== '';
    const isReady = normalizeText(row.OBS) === 'PRONTO';

    if (hasMachineDown) {
      ordensFeitas += 1;
      totalPecasFeitas += qtd;
    }

    if (isReady) {
      ordensProntas += 1;
      totalPecasProntas += qtd;
    }
  });

  const eficienciaOrdensFeitas = totalOrdens > 0 ? (ordensFeitas / totalOrdens) * 100 : 0;
  const eficienciaPecasFeitas = totalPecasPlanejadas > 0 ? (totalPecasFeitas / totalPecasPlanejadas) * 100 : 0;
  const eficienciaOrdens = totalOrdens > 0 ? (ordensProntas / totalOrdens) * 100 : 0;
  const eficienciaPecas = totalPecasPlanejadas > 0 ? (totalPecasProntas / totalPecasPlanejadas) * 100 : 0;
  const percentualTeflonOrdens = totalOrdens > 0 ? (ordensTeflon / totalOrdens) * 100 : 0;
  const percentualTeflonPecas = totalPecasPlanejadas > 0 ? (totalPecasTeflon / totalPecasPlanejadas) * 100 : 0;
  const percentualOrdensCriticas = totalOrdens > 0 ? (ordensCriticas / totalOrdens) * 100 : 0;
  const percentualPecasCriticas = totalPecasPlanejadas > 0 ? (pecasCriticas / totalPecasPlanejadas) * 100 : 0;

  return {
    mode: 'mestra',
    totalOrdens,
    ordensFeitas,
    ordensProntas,
    totalPecasPlanejadas,
    totalPecasFeitas,
    totalPecasProntas,
    ordensTeflon,
    totalPecasTeflon,
    ordensCriticas,
    pecasCriticas,
    percentualOrdensCriticas,
    percentualPecasCriticas,
    eficienciaOrdensFeitas,
    eficienciaPecasFeitas,
    eficienciaOrdens,
    eficienciaPecas,
    percentualTeflonOrdens,
    percentualTeflonPecas
  };
}

function calculateProjectKpis(rows) {
  let totalProgramas = 0;
  let totalProgramasNovos = 0;
  let programasOtimizados = 0;
  let itensProjetoLiberado = 0;

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    totalProgramas += 1;

    const normalizedPrograma = normalizeText(row.PROGRAMA);
    const isProgramaNovo = PROJECT_NEW_PROGRAM_VALUES.has(normalizedPrograma);
    if (isProgramaNovo) {
      totalProgramasNovos += 1;
    } else {
      programasOtimizados += 1;
    }

    const normalizedObs = normalizeText(row.OBS);
    const isProjetoLiberado = normalizedObs.includes('PROJETO LIBERADO') || normalizedObs === 'PRONTO';
    if (isProjetoLiberado) {
      itensProjetoLiberado += 1;
    }
  });

  const itensFaltandoLiberar = Math.max(0, totalProgramas - itensProjetoLiberado);
  const eficienciaLiberacao = totalProgramas > 0 ? (itensProjetoLiberado / totalProgramas) * 100 : 0;

  return {
    mode: 'projeto',
    totalProgramas,
    totalProgramasNovos,
    programasOtimizados,
    itensProjetoLiberado,
    itensFaltandoLiberar,
    eficienciaLiberacao
  };
}

function calculateAcabamentoKpis(rows) {
  let totalOrdens = 0;
  let ordensProntas = 0;
  let totalPecas = 0;
  let pecasFeitas = 0;
  let ordensCriticas = 0;
  let pecasCriticas = 0;

  const isCriticalOrder = (row) => {
    const normalizedReq = normalizeText(row['N REQ']);
    const hasDuplicatedReq = normalizedReq !== '' && duplicatedReqs.has(normalizedReq);
    return isCriticalByText(row['DESCRIÇÃO'], row.OBS) || hasDuplicatedReq;
  };

  const detectOperatorFromObs = (obsValue) => {
    const normalizedObs = normalizeText(obsValue);
    if (normalizedObs.includes('GUSTAVO')) {
      return 'GUSTAVO';
    }

    for (const operator of ACABAMENTO_OPERATORS) {
      if (normalizedObs === operator || normalizedObs.includes(operator)) {
        return operator;
      }
    }
    return null;
  };

  const operatorStats = new Map();
  ACABAMENTO_OPERATORS.forEach((operator) => {
    operatorStats.set(operator, {
      operador: operator,
      ordens: 0,
      pecas: 0,
      ordensCriticas: 0,
      pecasCriticas: 0
    });
  });

  const duplicatedReqs = buildDuplicatedReqSet(rows, 'N REQ');

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    totalOrdens += 1;
    const qtd = parseNumericValue(row.QTD);
    totalPecas += qtd;
    const isCritical = isCriticalOrder(row);
    if (isCritical) {
      ordensCriticas += 1;
      pecasCriticas += qtd;
    }

    const operator = detectOperatorFromObs(row.OBS);
    if (operator && operatorStats.has(operator)) {
      const stats = operatorStats.get(operator);
      stats.ordens += 1;
      stats.pecas += qtd;
      if (isCritical) {
        stats.ordensCriticas += 1;
        stats.pecasCriticas += qtd;
      }
    }

    const readyObs = isReadyObsValue(row.OBS);
    if (readyObs) {
      ordensProntas += 1;
      pecasFeitas += qtd;
      return;
    }

    pecasFeitas += parseProducedPiecesFromObs(row.OBS, qtd);
  });

  const eficienciaOrdens = totalOrdens > 0 ? (ordensProntas / totalOrdens) * 100 : 0;
  const eficienciaPecas = totalPecas > 0 ? (pecasFeitas / totalPecas) * 100 : 0;
  const percentualOrdensCriticas = totalOrdens > 0 ? (ordensCriticas / totalOrdens) * 100 : 0;
  const percentualPecasCriticas = totalPecas > 0 ? (pecasCriticas / totalPecas) * 100 : 0;

  const operadores = Array.from(operatorStats.values())
    .filter((item) => item.ordens > 0 || item.pecas > 0)
    .map((item) => ({
      ...item,
      percentualOrdens: totalOrdens > 0 ? (item.ordens / totalOrdens) * 100 : 0,
      percentualPecas: totalPecas > 0 ? (item.pecas / totalPecas) * 100 : 0
    }));

  return {
    mode: 'acabamento',
    totalOrdens,
    ordensProntas,
    totalPecas,
    pecasFeitas,
    ordensCriticas,
    pecasCriticas,
    percentualOrdensCriticas,
    percentualPecasCriticas,
    operadores,
    eficienciaOrdens,
    eficienciaPecas
  };
}

function calculateMoldagemKpis(rows) {
  let pesoProcessadoTotal = 0;
  let totalOrdensProcessadas = 0;
  let totalBuchasProcessadas = 0;
  const operadorStats = new Map();

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    const isPrensado = normalizeText(row.STATUS) === MOLDAGEM_STATUS_PRENSADO;
    if (!isPrensado) {
      return;
    }

    totalOrdensProcessadas += 1;
    const pesoTotal = parseNumericValue(row['PESO TOTAL']);
    const totalBuchas = parseNumericValue(row['TOTAL BUCHAS']);
    pesoProcessadoTotal += pesoTotal;
    totalBuchasProcessadas += totalBuchas;

    const operador = normalizeText(row.OPERADOR) || 'SEM OPERADOR';
    const currentStats = operadorStats.get(operador) || {
      operador,
      ordens: 0,
      buchas: 0,
      peso: 0
    };
    currentStats.ordens += 1;
    currentStats.buchas += totalBuchas;
    currentStats.peso += pesoTotal;
    operadorStats.set(operador, currentStats);
  });

  const operadores = Array.from(operadorStats.values())
    .map((item) => ({
      ...item,
      percentualOrdens: totalOrdensProcessadas > 0 ? (item.ordens / totalOrdensProcessadas) * 100 : 0,
      percentualBuchas: totalBuchasProcessadas > 0 ? (item.buchas / totalBuchasProcessadas) * 100 : 0,
      percentualPeso: pesoProcessadoTotal > 0 ? (item.peso / pesoProcessadoTotal) * 100 : 0
    }))
    .sort((a, b) => {
      if (b.ordens !== a.ordens) {
        return b.ordens - a.ordens;
      }
      if (b.buchas !== a.buchas) {
        return b.buchas - a.buchas;
      }
      return b.peso - a.peso;
    });

  return {
    mode: 'moldagem',
    pesoProcessadoTotal,
    totalOrdensProcessadas,
    totalBuchasProcessadas,
    operadores
  };
}

function isReadyObsValue(obsValue) {
  return READY_OBS_VALUES.has(normalizeText(obsValue));
}

function buildProjetoReadyIndex(rows) {
  return buildReadyIndex(rows || [], (row) => {
    return !!row.__highlight || normalizeText(row.OBS) === 'PROJETO LIBERADO' || isReadyObsValue(row.OBS);
  });
}

function buildMachineReadyByKey(machineData) {
  const machineReadyByKey = new Map();
  machineData.forEach((data, index) => {
    if (!data) {
      return;
    }

    const machineType = MACHINE_LIST_TYPES[index];
    const machineLabel = LIST_TYPES[machineType].label.toUpperCase();
    data.rows.forEach((row) => {
      if (row.__isQtdDivider) {
        return;
      }

      const plannedQtd = parseNumericValue(row.QTD);
      const producedQtd = parseProducedPiecesFromObs(row.OBS, plannedQtd);
      const hasMachineDown = isReadyObsValue(row.OBS)
        || (plannedQtd > 0 && producedQtd >= plannedQtd);

      if (!hasMachineDown) {
        return;
      }

      const key = getMasterRowKey(row);
      if (!machineReadyByKey.has(key)) {
        machineReadyByKey.set(key, machineLabel);
      }
    });
  });

  return machineReadyByKey;
}

function getMasterRowKey(row) {
  const nPc = normalizeText(row['N PC']);
  const nReq = normalizeText(row['N REQ']);
  const descricao = normalizeText(row['DESCRIÇÃO']);
  const qtd = parseNumericValue(row.QTD);
  const linha = normalizeText(row.LINHA);

  if (nPc || nReq) {
    return `PC:${nPc}|REQ:${nReq}`;
  }
  if (descricao) {
    return `DESC:${descricao}|QTD:${qtd}`;
  }
  return `LINHA:${linha}|QTD:${qtd}`;
}

function buildReadyIndex(rows, isReadyRowFn) {
  const index = new Set();
  rows.forEach((row) => {
    if (row.__isQtdDivider || !isReadyRowFn(row)) {
      return;
    }
    index.add(getMasterRowKey(row));
  });
  return index;
}

function loadListDataByType(selectedDate, listType, settings, options = {}) {
  const config = LIST_TYPES[listType];
  const allowMissing = !!options.allowMissing;

  if (!config || !config.relativePath) {
    if (allowMissing) {
      return null;
    }
    throw new Error('Tipo de listagem invalido.');
  }

  const baseListPath = resolveBaseListPath(settings.rootFolder, config);
  if (!baseListPath) {
    if (allowMissing) {
      return null;
    }
    throw new Error(`Nao encontrei a pasta ${config.relativePath.join('/')} dentro da pasta raiz.`);
  }

  const yearFolder = path.join(baseListPath, String(selectedDate.getFullYear()));
  if (!fs.existsSync(yearFolder)) {
    if (allowMissing) {
      return null;
    }
    throw new Error(`Nao encontrei a pasta do ano ${selectedDate.getFullYear()}.`);
  }

  const monthName = toPtBrMonthName(selectedDate);
  const monthFolder = resolveMonthFolder(yearFolder, monthName);
  if (!monthFolder) {
    if (allowMissing) {
      return null;
    }
    throw new Error(`Nao encontrei a pasta do mes ${monthName}.`);
  }

  const filePath = resolveListFile(monthFolder, config.filePrefix, selectedDate.getDate(), selectedDate.getMonth() + 1);
  if (!filePath) {
    if (allowMissing) {
      return null;
    }
    throw new Error('Nao encontrei o arquivo da lista para a data selecionada.');
  }

  const workbook = XLSX.readFile(filePath, { cellDates: true, cellStyles: true });
  const sheetInfo = findWorksheetWithHeader(workbook, listType, config);
  if (!sheetInfo) {
    if (allowMissing) {
      return null;
    }
    if (listType === 'moldagem') {
      throw new Error('Nao foi encontrada a aba Lista com cabecalho esperado na linha 8.');
    }
    throw new Error('Nao foi encontrada uma planilha com o cabecalho esperado na linha 7.');
  }

  const rows = listType === 'moldagem'
    ? parseMoldagemRows(sheetInfo.sheet)
    : parseStandardTableRows(sheetInfo.sheet, listType, config.columns, config.machineValue);

  let kpis = null;
  if (config.machineValue) {
    kpis = calculateMachineKpis(rows, config.machineValue);
  } else if (listType === 'acabamento') {
    kpis = calculateAcabamentoKpis(rows);
  } else if (listType === 'projeto') {
    kpis = calculateProjectKpis(rows);
  } else if (listType === 'moldagem') {
    kpis = calculateMoldagemKpis(rows);
  }

  return {
    listType,
    filePath,
    sheetName: sheetInfo.sheetName,
    columns: config.columns,
    rows,
    kpis
  };
}

function loadMasterListData(selectedDate, settings) {
  const acabamentoData = loadListDataByType(selectedDate, 'acabamento', settings);
  const projetoData = loadListDataByType(selectedDate, 'projeto', settings, { allowMissing: true });
  const machineData = MACHINE_LIST_TYPES.map((type) => loadListDataByType(selectedDate, type, settings, { allowMissing: true }));

  const projetoReady = buildProjetoReadyIndex(projetoData ? projetoData.rows : []);

  const acabamentoReady = buildReadyIndex(acabamentoData.rows, (row) => isReadyObsValue(row.OBS));

  const machineReadyByKey = buildMachineReadyByKey(machineData);

  const rows = acabamentoData.rows.map((row) => {
      if (row.__isQtdDivider) {
        return {
          LINHA: row.LINHA || '',
          QTD: row.QTD || '',
          'DESCRIÇÃO': row['DESCRIÇÃO'] || '',
          'N PC': row['N PC'] || '',
          'N REQ': row['N REQ'] || '',
          CLIENTE: row.CLIENTE || '',
          'BAIXA PROJETO': '',
          'BAIXA MAQUINA': '',
          'BAIXA ACABAMENTO': '',
          OBS: '',
          __isTeflon: false,
          __buchaValue: '',
          __isQtdDivider: true
        };
      }

      const masterRow = {
        LINHA: row.LINHA || '',
        QTD: row.QTD || '',
        'DESCRIÇÃO': row['DESCRIÇÃO'] || '',
        'N PC': row['N PC'] || '',
        'N REQ': row['N REQ'] || '',
        CLIENTE: row.CLIENTE || '',
        'BAIXA PROJETO': '',
        'BAIXA MAQUINA': '',
        'BAIXA ACABAMENTO': '',
        OBS: '',
        __isTeflon: parseNumericValue(row.BUCHA) >= 1,
        __buchaValue: row.BUCHA || ''
      };

      const key = getMasterRowKey(row);
      const cellClasses = {};
      const hasProjeto = projetoReady.has(key);
      const machineLabel = machineReadyByKey.get(key);
      const hasMachine = !!machineLabel;
      const isReadyPiece = !!row.__qtdIsTextNumeric || isReadyPieceQtdValue(row.QTD);
      const hasMachineStatus = hasMachine || isReadyPiece;
      const hasAcabamento = acabamentoReady.has(key);

      if (hasProjeto) {
        masterRow['BAIXA PROJETO'] = 'PROJETO FEITO';
        cellClasses['BAIXA PROJETO'] = 'highlight-cell';
      }

      if (machineLabel) {
        masterRow['BAIXA MAQUINA'] = `BAIXADO ${machineLabel}`;
        cellClasses['BAIXA MAQUINA'] = 'highlight-cell';
      } else if (isReadyPiece) {
        masterRow['BAIXA MAQUINA'] = 'PEÇA PRONTA';
        cellClasses['BAIXA MAQUINA'] = 'highlight-cell';
      }

      if (hasAcabamento) {
        masterRow['BAIXA ACABAMENTO'] = 'BAIXADO ACAB.';
        cellClasses['BAIXA ACABAMENTO'] = 'highlight-cell';
        masterRow.OBS = 'PRONTO';
        masterRow.__highlight = true;
      }

      if (hasMachine && !hasProjeto && !cellClasses['BAIXA PROJETO']) {
        cellClasses['BAIXA PROJETO'] = 'warning-cell';
      }

      if (hasAcabamento && !hasMachineStatus && !cellClasses['BAIXA MAQUINA']) {
        cellClasses['BAIXA MAQUINA'] = 'warning-cell';
      }

      if (hasAcabamento && !hasProjeto && !cellClasses['BAIXA PROJETO']) {
        cellClasses['BAIXA PROJETO'] = 'warning-cell';
      }

      if (Object.keys(cellClasses).length > 0) {
        masterRow.__cellClasses = cellClasses;
      }

      return masterRow;
    });

  const foundMachineNames = machineData
    .filter((data) => !!data)
    .map((data) => data.listType.toUpperCase().replace('FAGOR', 'FAGOR ').replace('MCS', 'MCS ').replace(/\s+/g, ' ').trim());

  const machineInfo = foundMachineNames.length > 0 ? foundMachineNames.join(', ') : 'nenhuma';

  return {
    listType: 'mestra',
    filePath: `${acabamentoData.filePath} | Cruzamento: PROJETO + ${machineInfo} + ACABAMENTO`,
    sheetName: acabamentoData.sheetName,
    columns: MESTRA_COLUMNS,
    rows,
    kpis: calculateMasterKpis(rows)
  };
}

function loadListData({ date, listType }) {
  if (!date) {
    throw new Error('Selecione uma data.');
  }

  const settings = loadSettings();
  if (!settings.rootFolder) {
    throw new Error('Configure a pasta raiz na aba Configuracao.');
  }

  if (!LIST_TYPES[listType]) {
    throw new Error('Tipo de listagem invalido.');
  }

  if (!fs.existsSync(settings.rootFolder)) {
    throw new Error('A pasta raiz configurada nao existe mais.');
  }

  const selectedDate = new Date(`${date}T00:00:00`);
  if (Number.isNaN(selectedDate.getTime())) {
    throw new Error('Data invalida.');
  }

  if (listType === 'mestra') {
    return loadMasterListData(selectedDate, settings);
  }

  return loadListDataByType(selectedDate, listType, settings);
}

function getFreightGroupName(value) {
  const raw = String(value || '').trim();
  return raw ? raw.toUpperCase() : 'SEM ENTREGA';
}

function getFreightClientName(value) {
  const raw = String(value || '').trim();
  return raw ? raw.toUpperCase() : 'SEM CLIENTE';
}

function getFreightOrderId(row, fallbackIndex) {
  const nPc = normalizeText(row['N PC']);
  const nReq = normalizeText(row['N REQ']);
  if (nPc) {
    return `PC:${nPc}`;
  }
  if (nReq) {
    return `REQ:${nReq}`;
  }

  const linha = normalizeText(row.LINHA);
  const descricao = normalizeText(row['DESCRIÇÃO']);
  const qtd = parseNumericValue(row.QTD);
  return `LINHA:${linha}|DESC:${descricao}|QTD:${qtd}|IDX:${fallbackIndex}`;
}

function getFreightOrderLabel(row, fallbackIndex) {
  const nPc = String(row['N PC'] || '').trim().toUpperCase();
  if (nPc) {
    return nPc;
  }

  const nReq = String(row['N REQ'] || '').trim().toUpperCase();
  if (nReq) {
    return `REQ ${nReq}`;
  }

  const linha = String(row.LINHA || '').trim();
  return linha ? `Linha ${linha}` : `Linha ${fallbackIndex + 1}`;
}

function getFreightItemBaseKey(row) {
  const nPc = normalizeText(row['N PC']);
  const nReq = normalizeText(row['N REQ']);
  const descricao = normalizeText(row['DESCRIÇÃO']);
  const qtd = parseNumericValue(row.QTD);
  const cliente = normalizeText(row.CLIENTE);

  if (nPc || nReq) {
    return `PC:${nPc}|REQ:${nReq}|DESC:${descricao}|QTD:${qtd}|CLI:${cliente}`;
  }
  if (descricao) {
    return `DESC:${descricao}|QTD:${qtd}|CLI:${cliente}`;
  }
  return `QTD:${qtd}|CLI:${cliente}`;
}

function getFreightItemLineKey(row) {
  const linha = normalizeText(row.LINHA);
  return `LINHA:${linha}|${getFreightItemBaseKey(row)}`;
}

function buildReadyMatcherForFreight(rows, isReadyRowFn) {
  const lineKeys = new Set();
  const totalBaseCounts = new Map();
  const readyBaseKeys = new Set();
  (rows || []).forEach((row) => {
    if (!row || row.__isQtdDivider) {
      return;
    }

    const baseKey = getFreightItemBaseKey(row);
    totalBaseCounts.set(baseKey, (totalBaseCounts.get(baseKey) || 0) + 1);

    if (!isReadyRowFn(row)) {
      return;
    }

    const lineKey = getFreightItemLineKey(row);
    lineKeys.add(lineKey);
    readyBaseKeys.add(baseKey);
  });

  return {
    has(row) {
      const lineKey = getFreightItemLineKey(row);
      if (lineKeys.has(lineKey)) {
        return true;
      }

      const baseKey = getFreightItemBaseKey(row);
      // Fallback sem linha so quando existe apenas 1 linha total desse item.
      return readyBaseKeys.has(baseKey) && (totalBaseCounts.get(baseKey) || 0) === 1;
    }
  };
}

function buildMachineReadyByKeyForFreight(machineData) {
  const lineKeyToInfo = new Map();
  const baseKeyToInfos = new Map();
  const totalBaseCounts = new Map();
  (machineData || []).forEach((data, index) => {
    if (!data) {
      return;
    }

    const machineType = MACHINE_LIST_TYPES[index];
    const machineLabel = LIST_TYPES[machineType].label.toUpperCase();
    (data.rows || []).forEach((row) => {
      if (!row || row.__isQtdDivider) {
        return;
      }

      const baseKey = getFreightItemBaseKey(row);
      totalBaseCounts.set(baseKey, (totalBaseCounts.get(baseKey) || 0) + 1);

      const plannedQtd = parseNumericValue(row.QTD);
      const producedQtd = parseProducedPiecesFromObs(row.OBS, plannedQtd);
      const hasMachineDown = isReadyObsValue(row.OBS)
        || (plannedQtd > 0 && producedQtd >= plannedQtd);
      if (!hasMachineDown) {
        return;
      }

      const lineKey = getFreightItemLineKey(row);
      const info = {
        machineLabel,
        obsValue: String(row.OBS || '').trim()
      };

      if (!lineKeyToInfo.has(lineKey)) {
        lineKeyToInfo.set(lineKey, info);
      }

      if (!baseKeyToInfos.has(baseKey)) {
        baseKeyToInfos.set(baseKey, new Map());
      }
      const byMachine = baseKeyToInfos.get(baseKey);
      if (!byMachine.has(machineLabel)) {
        byMachine.set(machineLabel, info);
      }
    });
  });

  return {
    getInfo(row) {
      const lineKey = getFreightItemLineKey(row);
      if (lineKeyToInfo.has(lineKey)) {
        return lineKeyToInfo.get(lineKey);
      }

      const baseKey = getFreightItemBaseKey(row);
      const infosByMachine = baseKeyToInfos.get(baseKey);
      if (infosByMachine && infosByMachine.size === 1 && (totalBaseCounts.get(baseKey) || 0) === 1) {
        return Array.from(infosByMachine.values())[0];
      }
      return null;
    },
    has(row) {
      return !!this.getInfo(row);
    }
  };
}

function buildFreightSummary(sourceRows, acabamentoRows, projetoRows, machineData) {
  const acabamentoReady = buildReadyMatcherForFreight(acabamentoRows, (row) => isReadyObsValue(row.OBS));
  const projetoReady = buildReadyMatcherForFreight(
    projetoRows || [],
    (row) => !!row.__highlight || normalizeText(row.OBS) === 'PROJETO LIBERADO' || isReadyObsValue(row.OBS)
  );
  const machineReadyByKey = buildMachineReadyByKeyForFreight(machineData || []);
  const freightMap = new Map();

  sourceRows.forEach((row, index) => {
    if (row.__isQtdDivider) {
      return;
    }

    const freightName = getFreightGroupName(row.ENTREGA);
    const clientName = getFreightClientName(row.CLIENTE);
    const freightId = normalizeText(freightName);
    const clientId = normalizeText(clientName);
    const orderId = getFreightOrderId(row, index);
    const orderLabel = getFreightOrderLabel(row, index);
    const hasAcabamentoDown = acabamentoReady.has(row);
    const machineInfo = machineReadyByKey.getInfo(row);
    const machineLabel = machineInfo ? machineInfo.machineLabel : null;
    const hasMachineDown = !!machineLabel;
    const hasProjetoDown = projetoReady.has(row);
    const itemStatus = hasAcabamentoDown
      ? 'acabamento'
      : hasMachineDown
        ? 'maquina'
        : hasProjetoDown
          ? 'projeto'
          : 'pendente';

    if (!freightMap.has(freightId)) {
      freightMap.set(freightId, {
        id: freightId,
        name: freightName,
        clients: new Map()
      });
    }

    const freight = freightMap.get(freightId);
    if (!freight.clients.has(clientId)) {
      freight.clients.set(clientId, {
        id: clientId,
        name: clientName,
        orders: new Map()
      });
    }

    const client = freight.clients.get(clientId);
    if (!client.orders.has(orderId)) {
      client.orders.set(orderId, {
        id: orderId,
        orderLabel,
        items: [],
        totalItems: 0,
        doneItems: 0,
        pendingItems: 0
      });
    }

    const order = client.orders.get(orderId);
    order.items.push({
      lineLabel: String(row.LINHA || '').trim(),
      clientName,
      orderLabel,
      reqNumber: String(row['N REQ'] || '').trim(),
      description: String(row['DESCRIÇÃO'] || '').trim(),
      qtd: String(row.QTD || '').trim(),
      deliveryType: freightName,
      statusText: String(row.OBS || '').trim(),
      statusValue: machineInfo ? String(machineInfo.obsValue || '').trim() : String(row.OBS || '').trim(),
      isDone: hasAcabamentoDown,
      statusCode: itemStatus
    });
    order.totalItems += 1;
    if (hasAcabamentoDown) {
      order.doneItems += 1;
    } else {
      order.pendingItems += 1;
    }
  });

  const freightRows = Array.from(freightMap.values()).map((freight) => {
    const clients = Array.from(freight.clients.values()).map((client) => {
      const orders = Array.from(client.orders.values());
      orders.sort((a, b) => a.orderLabel.localeCompare(b.orderLabel, 'pt-BR'));
      orders.forEach((order) => {
        order.items.sort((a, b) => a.lineLabel.localeCompare(b.lineLabel, 'pt-BR'));
      });

      const totalOrders = orders.length;
      const doneOrders = orders.filter((order) => order.pendingItems === 0).length;
      const pendingOrders = totalOrders - doneOrders;

      return {
        id: client.id,
        name: client.name,
        totalOrders,
        doneOrders,
        pendingOrders,
        orders
      };
    });

    clients.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
    const totalOrders = clients.reduce((acc, client) => acc + client.totalOrders, 0);
    const doneOrders = clients.reduce((acc, client) => acc + client.doneOrders, 0);
    const pendingOrders = totalOrders - doneOrders;

    return {
      id: freight.id,
      name: freight.name,
      totalOrders,
      doneOrders,
      pendingOrders,
      clients
    };
  });

  freightRows.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
  return freightRows;
}

function loadFreightSummary({ date, listType }) {
  if (!date) {
    throw new Error('Selecione uma data para abrir o resumo de fretes.');
  }

  const settings = loadSettings();
  if (!settings.rootFolder) {
    throw new Error('Configure a pasta raiz na aba Configuracao.');
  }
  if (!fs.existsSync(settings.rootFolder)) {
    throw new Error('A pasta raiz configurada nao existe mais.');
  }
  if (!LIST_TYPES[listType]) {
    throw new Error('Tipo de listagem invalido para resumo de fretes.');
  }

  const selectedDate = new Date(`${date}T00:00:00`);
  if (Number.isNaN(selectedDate.getTime())) {
    throw new Error('Data invalida.');
  }

  if (listType === 'moldagem' || listType === 'mestra') {
    throw new Error('Resumo de fretes disponivel para listas com coluna ENTREGA (ex: Acabamento/Projeto/Maquinas).');
  }

  const sourceData = loadListDataByType(selectedDate, listType, settings);
  if (!sourceData.columns.includes('ENTREGA')) {
    throw new Error('A lista selecionada nao possui coluna ENTREGA.');
  }

  const acabamentoData = listType === 'acabamento'
    ? sourceData
    : loadListDataByType(selectedDate, 'acabamento', settings, { allowMissing: true });
  const projetoData = loadListDataByType(selectedDate, 'projeto', settings, { allowMissing: true });
  const machineData = MACHINE_LIST_TYPES.map((type) => loadListDataByType(selectedDate, type, settings, { allowMissing: true }));

  if (!acabamentoData) {
    throw new Error('Nao encontrei a lista de acabamento para validar o status pronto/faltando.');
  }

  return {
    date,
    listType,
    listLabel: LIST_TYPES[listType].label,
    freights: buildFreightSummary(
      sourceData.rows,
      acabamentoData.rows,
      projetoData ? projetoData.rows : [],
      machineData
    )
  };
}

function getTodayYmdLocal() {
  const now = new Date();
  const offsetMs = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offsetMs).toISOString().slice(0, 10);
}

function loadPcpEfficiencySnapshot({ date, listType }) {
  const settings = loadSettings();
  if (!settings.rootFolder) {
    throw new Error('Configure a pasta raiz na aba Configuracao.');
  }
  if (!fs.existsSync(settings.rootFolder)) {
    throw new Error('A pasta raiz configurada nao existe mais.');
  }

  const safeDate = date || getTodayYmdLocal();
  const selectedDate = new Date(`${safeDate}T00:00:00`);
  if (Number.isNaN(selectedDate.getTime())) {
    throw new Error('Data invalida.');
  }

  const selectedType = listType || 'all';
  const validTypes = new Set(['all', ...MACHINE_LIST_TYPES]);
  if (!validTypes.has(selectedType)) {
    throw new Error('Listagem de eficiencia invalida.');
  }

  const typesToLoad = selectedType === 'all' ? MACHINE_LIST_TYPES : [selectedType];
  const opBased = new Set(['fagor1', 'mcs1']);
  const machineItems = typesToLoad.map((type) => {
    const config = LIST_TYPES[type];
    const data = loadListDataByType(selectedDate, type, settings, { allowMissing: true });

    if (!data) {
      return {
        listType: type,
        machineName: config.label,
        basis: opBased.has(type) ? 'ordens' : 'pecas',
        totalOrdens: 0,
        ordensFeitas: 0,
        totalPecasPlanejadas: 0,
        totalPecasProduzidas: 0,
        missing: true
      };
    }

    const kpis = data.kpis || calculateMachineKpis(data.rows, config.machineValue);
    return {
      listType: type,
      machineName: config.label,
      basis: opBased.has(type) ? 'ordens' : 'pecas',
      totalOrdens: kpis.totalOrdens || 0,
      ordensFeitas: kpis.ordensFeitas || 0,
      totalPecasPlanejadas: kpis.totalPecasPlanejadas || 0,
      totalPecasProduzidas: kpis.totalPecasProduzidas || 0,
      missing: false
    };
  });

  const totalOrdens = machineItems.reduce((acc, item) => acc + item.totalOrdens, 0);
  const ordensFeitas = machineItems.reduce((acc, item) => acc + item.ordensFeitas, 0);
  const totalPecasPlanejadas = machineItems.reduce((acc, item) => acc + item.totalPecasPlanejadas, 0);
  const totalPecasProduzidas = machineItems.reduce((acc, item) => acc + item.totalPecasProduzidas, 0);

  return {
    date: safeDate,
    emittedAtIso: new Date().toISOString(),
    listType: selectedType,
    listLabel: selectedType === 'all' ? 'Geral (todas máquinas)' : LIST_TYPES[selectedType].label,
    machines: machineItems,
    overall: {
      totalOrdens,
      ordensFeitas,
      totalPecasPlanejadas,
      totalPecasProduzidas
    }
  };
}

function parseDateToDdMm(dateValue) {
  return parseDashboardDate(dateValue).ddMm;
}

function parseDashboardDate(dateValue) {
  const dt = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(dt.getTime())) {
    throw new Error('Data invalida para Dashboard PCP.');
  }
  const dd = String(dt.getDate()).padStart(2, '0');
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  return { dt, ddMm: `${dd}-${mm}` };
}

function findDailyDashboardWorkbookPath(dateValue) {
  const { dt, ddMm } = parseDashboardDate(dateValue);
  const basePath = resolvePcpDashboardDailyRootPath();

  const yearName = String(dt.getFullYear());
  const yearFolder = fs.existsSync(path.join(basePath, yearName))
    ? path.join(basePath, yearName)
    : findChildDirByNormalizedName(basePath, yearName);

  const monthFolder = yearFolder ? resolveMonthFolder(yearFolder, toPtBrMonthName(dt)) : null;

  const searchPaths = [];
  if (monthFolder && fs.existsSync(monthFolder)) {
    searchPaths.push(monthFolder);
  }
  searchPaths.push(basePath);

  const target = normalizeText(ddMm);
  for (const searchPath of searchPaths) {
    if (!searchPath || !fs.existsSync(searchPath)) {
      continue;
    }

    const files = fs.readdirSync(searchPath, { withFileTypes: true })
      .filter((entry) => entry.isFile())
      .map((entry) => entry.name)
      .filter((name) => /\.xlsm$/i.test(name));

    const candidates = files.filter((name) => normalizeText(name).includes(target));
    if (!candidates.length) {
      continue;
    }

    const sorted = candidates.sort((a, b) => a.localeCompare(b, 'pt-BR'));
    return path.join(searchPath, sorted[0]);
  }

  throw new Error(`Nao encontrei planilha de usinagem para ${ddMm} em ${searchPaths.join(' | ')}.`);
}

function detectDashboardHeaderMap(headerRow) {
  const map = {};
  (headerRow || []).forEach((value, idx) => {
    const key = normalizeText(value);
    const compact = key.replace(/\s+/g, '');
    if (!key) {
      return;
    }
    if (key === 'LINHA' || key === 'LIN' || compact === 'LINHA' || compact === 'LIN') map.linha = idx;
    if (key === 'QTD') map.qtd = idx;
    if (key === 'N PC' || key === 'PC' || compact === 'NPC') map.npc = idx;
    if (key === 'N REQ' || key === 'REQ' || compact === 'NREQ') map.nreq = idx;
    if (key === 'CLIENTE') map.cliente = idx;
    if (compact === 'CLIENTE') map.cliente = idx;
    if (key === 'BUCHA') map.bucha = idx;
    if (key === 'OBS') map.obs = idx;
    if (key === 'ACABAMENTO' || compact === 'ACABAMENTO') map.acabamento = idx;
    if ((key.includes('MAQ') && key.includes('OPER')) || compact.includes('MAQOPER')) map.maqoper = idx;
    if (key.includes('DESCRICAO') || compact.includes('ITENSATRASADOSLISTAANTERIOR')) map.descricao = idx;
    if (key === 'PROGRAMA' || compact === 'PROGRAMA') map.programa = idx;
    if (key === 'ENTREGA' || compact === 'ENTREGA') map.entrega = idx;
  });

  const mandatory = ['qtd', 'npc', 'nreq', 'cliente', 'obs'];
  const ok = mandatory.every((field) => map[field] !== undefined);
  return ok ? map : null;
}

function isDashboardQtdCellValid(value) {
  const text = String(value || '').trim();
  if (!text) {
    return false;
  }
  return normalizeText(text) !== 'QTD';
}

function loadDashboardPlanilhaRows(workbook) {
  let selectedSheetName = workbook.SheetNames.find((name) => {
    const n = normalizeText(name);
    const compact = n.replace(/\s+/g, '');
    return n === 'TABELA7' || n === 'TABELA 7' || n === 'PLANILHA 1' || compact === 'PLANILHA1';
  }) || workbook.SheetNames[0];

  const readAoa = (sheetName) => {
    const ws = workbook.Sheets[sheetName];
    if (!ws) {
      return null;
    }
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  };

  let aoa = readAoa(selectedSheetName);
  if (!aoa) {
    throw new Error('Nao foi possivel abrir a planilha diaria para o Dashboard PCP.');
  }

  // Detecta automaticamente a aba correta pela assinatura da Tabela7 (linha 14, B=QTD e J=OBS).
  const detectSheetBySignature = () => {
    for (const name of workbook.SheetNames) {
      const data = readAoa(name);
      if (!data || data.length < 14) {
        continue;
      }
      const row14 = data[13] || [];
      const colB = normalizeText(row14[1] || '');
      const colJ = normalizeText(row14[9] || '');
      if (colB === 'QTD' && colJ === 'OBS') {
        return { name, data };
      }
    }
    return null;
  };

  const signature = detectSheetBySignature();
  if (signature) {
    selectedSheetName = signature.name;
    aoa = signature.data;
  }

  const detectHeaderInAoa = (data) => {
    if (!Array.isArray(data) || !data.length) {
      return { headerIndex: -1, headerMap: null };
    }

    // Prioridade: Tabela7 fixa na linha 14 (index 13), colunas A:J.
    const fixedHeaderIndex = 13;
    const fixedHeaderRow = data[fixedHeaderIndex] || [];
    const fixedMap = detectDashboardHeaderMap(fixedHeaderRow.slice(0, 10));
    const fixedColB = normalizeText(fixedHeaderRow[1] || '');
    const fixedColJ = normalizeText(fixedHeaderRow[9] || '');
    const fixedMatches = (fixedColB === 'QTD' && fixedColJ === 'OBS')
      || (fixedMap && fixedMap.qtd === 1 && fixedMap.obs === 9 && fixedMap.npc !== undefined && fixedMap.nreq !== undefined);
    if (fixedMatches) {
      return {
        headerIndex: fixedHeaderIndex,
        headerMap: {
          linha: 0,
          qtd: 1,
          descricao: 2,
          npc: 3,
          nreq: 4,
          cliente: 5,
          programa: 6,
          bucha: 7,
          entrega: 8,
          obs: 9,
          maqoper: 10,
          acabamento: 12
        }
      };
    }

    // Fallback: detecção dinâmica por cabeçalho.
    for (let i = 0; i < Math.min(data.length, 120); i += 1) {
      const rowSlice = (data[i] || []).slice(0, 30);
      const map = detectDashboardHeaderMap(rowSlice);
      const rowNorm = rowSlice.map((cell) => normalizeText(cell));
      const compact = rowNorm.map((txt) => txt.replace(/\s+/g, ''));
      const hasQtd = rowNorm.includes('QTD');
      const hasObs = rowNorm.includes('OBS');
      const hasLin = rowNorm.includes('LINHA') || rowNorm.includes('LIN') || compact.includes('LINHA') || compact.includes('LIN');
      const hasReq = rowNorm.includes('N REQ') || compact.includes('NREQ');
      const hasPc = rowNorm.includes('N PC') || compact.includes('NPC');
      if (map && hasQtd && hasObs && hasLin && hasReq && hasPc) {
        return { headerIndex: i, headerMap: map };
      }
    }

    return { headerIndex: -1, headerMap: null };
  };

  let { headerIndex, headerMap } = detectHeaderInAoa(aoa);

  // Se não encontrou na aba inicial, varre todas as abas e escolhe a primeira com cabeçalho válido.
  if (headerIndex < 0 || !headerMap) {
    const preferredOrder = [
      ...workbook.SheetNames.filter((name) => {
        const n = normalizeText(name);
        const compact = n.replace(/\s+/g, '');
        return n === 'PLANILHA 1' || compact === 'PLANILHA1';
      }),
      ...workbook.SheetNames.filter((name) => {
        const n = normalizeText(name);
        return n !== 'PLANILHA 1' && n !== 'PLANILHA1';
      })
    ];

    for (const sheetName of preferredOrder) {
      const candidateAoa = readAoa(sheetName);
      const detected = detectHeaderInAoa(candidateAoa);
      if (detected.headerIndex >= 0 && detected.headerMap) {
        selectedSheetName = sheetName;
        aoa = candidateAoa;
        headerIndex = detected.headerIndex;
        headerMap = detected.headerMap;
        break;
      }
    }
  }

  if (headerIndex < 0 || !headerMap) {
    throw new Error(`Nao encontrei cabecalho esperado da Tabela7. Aba lida: ${selectedSheetName}`);
  }

  const rows = [];
  for (let i = headerIndex + 1; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const linha = headerMap.linha !== undefined ? String(row[headerMap.linha] || '').trim() : String(i - headerIndex).trim();
    const qtdRaw = row[headerMap.qtd];
    const descricao = String(row[headerMap.descricao] || '').trim();
    const nPc = String(row[headerMap.npc] || '').trim();
    const nReq = String(row[headerMap.nreq] || '').trim();
    const cliente = String(row[headerMap.cliente] || '').trim();
    const programa = String(row[headerMap.programa] || '').trim();
    const entrega = String(row[headerMap.entrega] || '').trim();
    const bucha = String(row[headerMap.bucha] || '').trim();
    const obs = String(row[headerMap.obs] || '').trim();
    const maqOper = String(
      (headerMap.maqoper !== undefined ? row[headerMap.maqoper] : row[10]) || ''
    ).trim();
    const acabamento = String(
      (headerMap.acabamento !== undefined ? row[headerMap.acabamento] : row[12]) || ''
    ).trim();

    const qtdCurrent = String(qtdRaw || '').trim();
    const qtdNext1 = String(((aoa[i + 1] || [])[headerMap.qtd]) || '').trim();
    const qtdNext2 = String(((aoa[i + 2] || [])[headerMap.qtd]) || '').trim();
    // Fim da Tabela7: QTD atual vazio e as proximas duas linhas de QTD vazias.
    if (!qtdCurrent && !qtdNext1 && !qtdNext2) {
      break;
    }

    const hasData = [linha, descricao, nPc, nReq, cliente, obs, maqOper, acabamento, qtdCurrent]
      .some((v) => String(v || '').trim() !== '');
    if (!hasData) {
      continue;
    }

    rows.push({
      LINHA: linha,
      QTD: qtdRaw,
      'DESCRIÇÃO': descricao,
      'N PC': nPc,
      'N REQ': nReq,
      CLIENTE: cliente,
      PROGRAMA: programa,
      ENTREGA: entrega,
      BUCHA: bucha,
      'MAQ./OPER.': maqOper,
      OBS: obs,
      ACABAMENTO: acabamento
    });
  }

  return { rows, sheetName: selectedSheetName };
}

function parseDashboardCncMetrics(rows) {
  const machines = ['FAGOR 1', 'FAGOR 2', 'MCS 1', 'MCS 2', 'MCS 3'];
  const dataMap = new Map(
    machines.map((name) => [name, {
      machine: name,
      opsPlanned: 0,
      opsDone: 0,
      piecesPlanned: 0,
      piecesDone: 0,
      rnc: 0,
      lostTime: 0
    }])
  );

  const normalizeMachine = (value) => {
    const text = normalizeText(value).replace(/\s+/g, '');
    if (text.includes('FAGOR1')) return 'FAGOR 1';
    if (text.includes('FAGOR2')) return 'FAGOR 2';
    if (text.includes('MCS1')) return 'MCS 1';
    if (text.includes('MCS2')) return 'MCS 2';
    if (text.includes('MCS3')) return 'MCS 3';
    return null;
  };

  rows.forEach((row) => {
    if (!isDashboardQtdCellValid(row.QTD)) {
      return;
    }
    const machine = normalizeMachine(row['MAQ./OPER.']);
    if (!machine || !dataMap.has(machine)) {
      return;
    }
    const qtd = parseNumericValue(row.QTD);
    const ready = normalizeText(row.OBS) === 'PRONTO';
    const item = dataMap.get(machine);
    item.opsPlanned += 1;
    item.piecesPlanned += qtd;
    if (ready) {
      item.opsDone += 1;
      item.piecesDone += qtd;
    }
  });

  const list = Array.from(dataMap.values()).map((item) => ({
    ...item,
    efficiencyOps: item.opsPlanned > 0 ? (item.opsDone / item.opsPlanned) * 100 : 0
  }));

  const totalPiecesDone = list.reduce((acc, item) => acc + item.piecesDone, 0);
  list.forEach((item) => {
    item.contributionPercent = totalPiecesDone > 0 ? (item.piecesDone / totalPiecesDone) * 100 : 0;
  });
  return list;
}

function detectAcabamentoOperator(acabamentoValue) {
  const normalizedObs = normalizeText(acabamentoValue);
  if (normalizedObs.includes('GUSTAVO')) {
    return 'GUSTAVO';
  }
  for (const operator of ACABAMENTO_OPERATORS) {
    if (normalizedObs === operator || normalizedObs.includes(operator)) {
      return operator;
    }
  }
  return null;
}

function parseDashboardMetrics(rows) {
  let pecasPlanejadas = 0;
  let pecasFeitas = 0;
  let opPlanejadas = 0;
  let opFeitas = 0;
  let opTeflon = 0;
  let pecasTeflon = 0;
  const opsByCompany = new Map();
  const acabamentoByOperator = new Map();

  rows.forEach((row) => {
    if (row.__isQtdDivider) {
      return;
    }

    if (!isDashboardQtdCellValid(row.QTD)) {
      return;
    }

    opPlanejadas += 1;
    const qtd = parseNumericValue(row.QTD);
    pecasPlanejadas += qtd;

    const cliente = String(row.CLIENTE || '').trim().toUpperCase() || 'SEM CLIENTE';
    opsByCompany.set(cliente, (opsByCompany.get(cliente) || 0) + 1);

    const ready = normalizeText(row.OBS) === 'PRONTO';
    if (ready) {
      pecasFeitas += qtd;
      opFeitas += 1;
    }

    if (parseNumericValue(row.BUCHA) >= 1) {
      opTeflon += 1;
      pecasTeflon += qtd;
    }

    const operator = detectAcabamentoOperator(row.ACABAMENTO);
    if (operator) {
      if (!acabamentoByOperator.has(operator)) {
        acabamentoByOperator.set(operator, { operador: operator, ops: 0, pecas: 0 });
      }
      const opData = acabamentoByOperator.get(operator);
      opData.ops += 1;
      opData.pecas += qtd;
    }
  });

  const eficOps = opPlanejadas > 0 ? (opFeitas / opPlanejadas) * 100 : 0;
  const teflonOpsPercent = opPlanejadas > 0 ? (opTeflon / opPlanejadas) * 100 : 0;
  const teflonPecasPercent = pecasPlanejadas > 0 ? (pecasTeflon / pecasPlanejadas) * 100 : 0;

  return {
    kpis: {
      pecasFeitas,
      pecasPlanejadas,
      opFeitas,
      opPlanejadas,
      eficOps,
      rncQtd: 0,
      rncCost: 0,
      totalCost: 0,
      opGeradas: 0,
      teflonOpsPercent,
      teflonPecasPercent
    },
    opsByCompany: Array.from(opsByCompany.entries())
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 16),
    acabamentoByOperator: Array.from(acabamentoByOperator.values())
      .sort((a, b) => b.ops - a.ops)
  };
}

function parseDashboardMaterials(workbook) {
  const sheetName = workbook.SheetNames.find((name) => normalizeText(name) === 'MOLDAGEM');
  if (!sheetName) {
    return [];
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return [];
  }

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  const materialsAlias = new Map([
    ['PURO', 'Puro'],
    ['BRONZE FREE', 'Bronze Free'],
    ['GRAFITE', 'Grafite'],
    ['T-46', 'T-46'],
    ['CARBONO', 'Carbono'],
    ['MOLIB', 'Molibdenio'],
    ['MOLIBDENIO', 'Molibdenio'],
    ['FV AZUL', 'Fibra de vidro azul'],
    ['FIBRA DE VIDRO AZUL', 'Fibra de vidro azul'],
    ['BRONZE LOW', 'Bronze Low'],
    ['FV BRANCA', 'Fibra de vidro branca'],
    ['FIBRA DE VIDRO BRANCA', 'Fibra de vidro branca'],
    ['FV AMARELA', 'Fibra de vidro amarela'],
    ['FIBRA DE VIDRO AMARELA', 'Fibra de vidro amarela'],
    ['FV PRETA', 'Fibra de vidro preta'],
    ['FIBRA DE VIDRO PRETA', 'Fibra de vidro preta']
  ]);

  const materialHeaderIndex = aoa.findIndex((row) => normalizeText((row && row[0]) || '') === 'MATERIAL');
  if (materialHeaderIndex < 0) {
    return [];
  }
  // Layout fixo informado:
  // A=Material, C=Estoque, D=Buchas, E=Utilizado, F=Refugo, G=OP
  const COL = {
    material: 0,
    estoque: 2,
    buchas: 3,
    kgUsed: 4,
    refugo: 5,
    ops: 6
  };

  const materials = [];
  for (let i = materialHeaderIndex + 1; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const rawName = String(row[COL.material] || '').trim();
    const normalizedName = normalizeText(rawName);
    if (!rawName) {
      if (materials.length > 0) {
        break;
      }
      continue;
    }
    const mappedName = materialsAlias.get(normalizedName);
    if (!mappedName) {
      if (materials.length > 0) {
        break;
      }
      continue;
    }

    materials.push({
      name: mappedName,
      estoque: parseNumericValue(row[COL.estoque]),
      buchas: parseNumericValue(row[COL.buchas]),
      kgUsed: parseNumericValue(row[COL.kgUsed]),
      refugoKg: parseNumericValue(row[COL.refugo]),
      ops: parseNumericValue(row[COL.ops])
    });
  }
  return materials;
}

function loadRncSnapshotForDate(dateValue) {
  const safeDate = String(dateValue || '').trim();
  if (!safeDate) {
    return {
      kpis: { rncQtd: 0, rncCost: 0, totalCost: 0 },
      usinagem: { machines: [] },
      report: { kpis: {}, charts: {} },
      debug: { sourceFile: null, sheetName: null }
    };
  }

  let workbookPath = null;
  let fiscalRoot = '';
  let expected = '';
  let fiscalRootExists = false;
  let expectedExists = false;
  try {
    fiscalRoot = getFiscalRoot();
    expected = resolveFiscalPath(FISCAL_RNC_RELATIVE);
    fiscalRootExists = !!(fiscalRoot && fs.existsSync(fiscalRoot));
    expectedExists = !!(expected && fs.existsSync(expected));
    workbookPath = resolveRncWorkbookPath();
  } catch (error) {
    return {
      kpis: { rncQtd: 0, rncCost: 0, totalCost: 0 },
      usinagem: { machines: [] },
      debug: {
        error: (error && error.message) || String(error || ''),
        sourceFile: null,
        sheetName: null,
        fiscalRoot,
        fiscalRootExists
      }
    };
  }
  if (!workbookPath || !fs.existsSync(workbookPath)) {
    return {
      kpis: { rncQtd: 0, rncCost: 0, totalCost: 0 },
      usinagem: { machines: [] },
      report: { kpis: {}, charts: {} },
      debug: { sourceFile: workbookPath || null, sheetName: null, fiscalRoot, expected, fiscalRootExists, expectedExists }
    };
  }

  let workbook = null;
  try {
    workbook = readDashboardWorkbookSafe(workbookPath);
  } catch (error) {
    const msg = (error && error.message) || '';
    const code = (error && error.code) || 'UNKNOWN';
    throw new Error(`Nao consegui ler a planilha de RNC (arquivo em uso/bloqueado). Codigo: ${code}. ${msg}`.trim());
  }
  const sheetName = workbook.SheetNames.find((name) => {
    const normalized = normalizeText(name).replace(/\s+/g, '');
    return normalized === 'RNC' || normalized.startsWith('RNC');
  });
  if (!sheetName || !workbook.Sheets[sheetName]) {
    return {
      kpis: { rncQtd: 0, rncCost: 0, totalCost: 0 },
      usinagem: { machines: [] },
      report: { kpis: {}, charts: {} },
      debug: { sourceFile: workbookPath, sheetName: sheetName || null }
    };
  }

  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

  // Layout fixo informado (Tabela1 iniciando na coluna B, cabeçalho na linha 5):
  // B=Nº RNC, ... M=Custo apurado, P=Data, Q=Segs Perdidos, R=Setor, T=Maquina
  const COL = {
    nrnc: 1, // B
    qtd: 6, // G
    descricao: 7, // H
    custo: 12, // M
    data: 15, // P
    segs: 16, // Q
    setor: 17, // R
    maquina: 19 // T
    ,
    material: 20 // U
  };
  const headerIndex = 4; // linha 5 (0-based)
  const startRow = headerIndex + 1; // linha 6 (0-based)
  let rncQtd = 0;
  let rncCost = 0;

  const machineStats = new Map();
  const selectedYmd = safeDate;
  let usinagemPiecesRefugadas = 0;
  let usinagemSegsPerdidos = 0;
  let usinagemCostApurado = 0;
  let totalSegsPerdidos = 0;

  const sectorStats = new Map();
  const reasonStats = new Map();
  const materialStats = new Map();
  const machineTimeStatsAll = new Map();
  const machineTimeStatsUsinagem = new Map();

  const toYmdFromCell = (cell) => {
    if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
      const yyyy = cell.getFullYear();
      const mm = String(cell.getMonth() + 1).padStart(2, '0');
      const dd = String(cell.getDate()).padStart(2, '0');
      return `${String(yyyy).padStart(4, '0')}-${mm}-${dd}`;
    }
    const raw = String(cell || '').trim();
    return parsePtBrDateToYmd(raw) || parsePtBrDateToYmd(formatExcelDatePtBr(cell));
  };

  const normalizeSector = (value) => {
    const raw = String(value || '').trim();
    const n = normalizeText(raw);
    if (!n) {
      return { key: 'SEM SETOR', label: 'SEM SETOR' };
    }
    if (n.includes('USINAGEM')) return { key: 'USINAGEM', label: 'USINAGEM' };
    if (n.includes('ACABAMENTO')) return { key: 'ACABAMENTO', label: 'ACABAMENTO' };
    if (n.includes('MOLDAGEM')) return { key: 'MOLDAGEM', label: 'MOLDAGEM' };
    if (n.includes('PROJETO')) return { key: 'PROJETO', label: 'PROJETO' };
    if (n.includes('PCP')) return { key: 'PCP', label: 'PCP' };
    if (n.includes('FORNECEDOR')) return { key: 'FORNECEDOR', label: 'FORNECEDOR' };
    return { key: n, label: raw.toUpperCase() };
  };

  for (let r = startRow; r <= range.e.r; r += 1) {
    const nrncCell = getCell(ws, r, COL.nrnc);
    const dateCell = getCell(ws, r, COL.data);
    const nrncValue = (nrncCell && nrncCell.v) != null ? String(nrncCell.v).trim() : '';
    const rawDate = dateCell ? dateCell.v : '';

    if (!nrncValue && !rawDate) {
      continue;
    }

    const rowYmd = toYmdFromCell(rawDate);
    if (!rowYmd || rowYmd !== selectedYmd) {
      continue;
    }

    rncQtd += 1;
    const costCell = getCell(ws, r, COL.custo);
    const cost = costCell ? parseNumericValue(costCell.v) : 0;
    rncCost += cost;

    const setorCell = getCell(ws, r, COL.setor);
    const sectorInfo = normalizeSector(setorCell ? setorCell.v : '');
    if (!sectorStats.has(sectorInfo.key)) {
      sectorStats.set(sectorInfo.key, { sector: sectorInfo.label, count: 0, cost: 0, segs: 0 });
    }
    const sectorEntry = sectorStats.get(sectorInfo.key);
    sectorEntry.count += 1;
    sectorEntry.cost += cost;

    const qtdCell = getCell(ws, r, COL.qtd);
    const qtd = qtdCell ? parseNumericValue(qtdCell.v) : 0;

    const machineCell = getCell(ws, r, COL.maquina);
    const machineRaw = machineCell ? String(machineCell.v || '').trim() : '';
    const machine = machineRaw || 'SEM MÁQUINA';
    const segsCell = getCell(ws, r, COL.segs);
    let segs = segsCell ? parseRncSecondsValue(segsCell.v) : 0;
    if (!segs) {
      if (qtd > 0) {
        segs = Math.round(qtd * 55);
      }
    }

    totalSegsPerdidos += segs;
    sectorEntry.segs += segs;

    if (!machineTimeStatsAll.has(machine)) {
      machineTimeStatsAll.set(machine, { machine, segs: 0 });
    }
    machineTimeStatsAll.get(machine).segs += segs;

    const descCell = getCell(ws, r, COL.descricao);
    const descRaw = descCell ? String(descCell.v || '').trim() : '';
    const descKey = normalizeText(descRaw);
    if (descKey) {
      const existing = reasonStats.get(descKey) || { name: descRaw, count: 0, cost: 0 };
      existing.count += 1;
      existing.cost += cost;
      reasonStats.set(descKey, existing);
    }

    const materialCell = getCell(ws, r, COL.material);
    const materialRaw = materialCell ? String(materialCell.v || '').trim() : '';
    const materialKey = normalizeText(materialRaw);
    if (materialKey) {
      const existing = materialStats.get(materialKey) || { name: materialRaw, count: 0 };
      existing.count += 1;
      materialStats.set(materialKey, existing);
    }

    if (sectorInfo.key !== 'USINAGEM') {
      continue;
    }

    if (!machineTimeStatsUsinagem.has(machine)) {
      machineTimeStatsUsinagem.set(machine, { machine, segs: 0 });
    }
    machineTimeStatsUsinagem.get(machine).segs += segs;

    if (!machineStats.has(machine)) {
      machineStats.set(machine, { machine, rnc: 0, segs: 0, cost: 0, pieces: 0 });
    }
    const entry = machineStats.get(machine);
    entry.rnc += 1;
    entry.segs += segs;
    entry.cost += cost;
    entry.pieces += qtd;

    usinagemPiecesRefugadas += qtd;
    usinagemSegsPerdidos += segs;
    usinagemCostApurado += cost;
  }

  const machines = Array.from(machineStats.values())
    .filter((m) => m.rnc > 0 || m.segs > 0 || m.cost > 0)
    .sort((a, b) => {
      if (b.rnc !== a.rnc) return b.rnc - a.rnc;
      if (b.segs !== a.segs) return b.segs - a.segs;
      return b.cost - a.cost;
    });

  const sectors = Array.from(sectorStats.values())
    .filter((s) => s.count > 0)
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return b.cost - a.cost;
    });

  const reasonsByCount = Array.from(reasonStats.values())
    .filter((r) => r.count > 0)
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return b.cost - a.cost;
    })
    .slice(0, 24);

  const reasonsByCost = Array.from(reasonStats.values())
    .filter((r) => r.cost > 0)
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 24);

  const materialsByCount = Array.from(materialStats.values())
    .filter((m) => m.count > 0)
    .sort((a, b) => b.count - a.count)
    .slice(0, 24);

  const machineTimeUsinagem = Array.from(machineTimeStatsUsinagem.values())
    .filter((m) => m.segs > 0)
    .sort((a, b) => b.segs - a.segs)
    .slice(0, 24);

  const totalSectorsCount = sectors.reduce((acc, s) => acc + Number(s.count || 0), 0);
  const sectorPie = sectors.map((s) => ({
    sector: s.sector,
    count: s.count,
    pct: totalSectorsCount > 0 ? (Number(s.count || 0) / totalSectorsCount) * 100 : 0
  }));

  const topSectorByCount = sectors.length ? sectors[0] : null;
  const topSectorByCost = sectors.slice().sort((a, b) => b.cost - a.cost)[0] || null;
  const topReason = reasonsByCount.length ? reasonsByCount[0] : null;

  return {
    kpis: {
      rncQtd,
      rncCost,
      totalCost: rncCost
    },
    usinagem: {
      machines,
      totals: {
        piecesRefugadas: usinagemPiecesRefugadas,
        segsPerdidos: usinagemSegsPerdidos,
        costApurado: usinagemCostApurado
      }
    },
    report: {
      kpis: {
        rncQtd,
        rncCost,
        avgCost: rncQtd > 0 ? rncCost / rncQtd : 0,
        segsPerdidos: totalSegsPerdidos,
        topSectorByCount: topSectorByCount ? topSectorByCount.sector : '-',
        topSectorByCost: topSectorByCost ? topSectorByCost.sector : '-',
        topReason: topReason ? topReason.name : '-'
      },
      charts: {
        sectorPie,
        sectorCounts: sectors.map((s) => ({ name: s.sector, value: s.count })),
        reasonsCount: reasonsByCount.map((r) => ({ name: r.name, value: r.count })),
        reasonsCost: reasonsByCost.map((r) => ({ name: r.name, value: r.cost })),
        machineTime: machineTimeUsinagem.map((m) => ({ name: m.machine, value: m.segs })),
        materialsCount: materialsByCount.map((m) => ({ name: m.name, value: m.count }))
      }
    },
    debug: {
      sourceFile: workbookPath,
      sheetName,
      headerIndex,
      col: COL
    }
  };
}

function findAoaCellByText(aoa, targetText) {
  const target = normalizeText(targetText);
  for (let r = 0; r < aoa.length; r += 1) {
    const row = aoa[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      if (normalizeText(row[c]) === target) {
        return { row: r, col: c };
      }
    }
  }
  return null;
}

function parseMoldagemPressesAndOperators(workbook) {
  const sheetName = workbook.SheetNames.find((name) => {
    const n = normalizeText(name).replace(/\s+/g, '');
    return n === 'TABELA9' || n === 'MOLDAGEM';
  });
  if (!sheetName) {
    return { presses: [], operators: [] };
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return { presses: [], operators: [] };
  }

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });

  const consolidatePresses1to3 = (inputPresses) => {
    const stats = new Map([
      [1, { press: 1, name: 'Prensa 1', ops: 0, buchas: 0 }],
      [2, { press: 2, name: 'Prensa 2', ops: 0, buchas: 0 }],
      [3, { press: 3, name: 'Prensa 3', ops: 0, buchas: 0 }]
    ]);

    (inputPresses || []).forEach((item) => {
      const pressNum = parseNumericValue(item && (item.press != null ? item.press : item.name));
      if (![1, 2, 3].includes(pressNum)) {
        return;
      }
      const entry = stats.get(pressNum);
      entry.ops += parseNumericValue(item && item.ops);
      entry.buchas += parseNumericValue(item && item.buchas);
    });

    return Array.from(stats.values());
  };

  const detectMoldagemPressTable = () => {
    // Preferência: Tabela9 fixa (cabeçalho na linha 15 / index 14).
    const sheetNorm = normalizeText(sheetName).replace(/\s+/g, '');
    if (sheetNorm === 'TABELA9' && aoa.length > 15) {
      const headerIndex = 14;
      const headerRow = aoa[headerIndex] || [];
      const normalized = headerRow.map((cell) => normalizeText(cell));
      const compact = normalized.map((txt) => txt.replace(/\s+/g, ''));

      let statusCol = normalized.findIndex((v) => v === 'STATUS');
      if (statusCol < 0) {
        statusCol = 13; // coluna N (fallback informado)
      }

      let pressColResolved = compact.findIndex((v) => v === 'NPRENSA' || v === 'NOPRENSA');
      if (pressColResolved < 0) {
        pressColResolved = normalized.findIndex((v) => v === 'PRENSA' || v.startsWith('PRENSA'));
      }
      if (pressColResolved < 0) {
        pressColResolved = normalized.findIndex((v) => v.startsWith('N') && v.includes('PRENSA'));
      }
      if (pressColResolved < 0) {
        pressColResolved = 14; // coluna O (fallback informado)
      }

      let totalBuchasColResolved = compact.findIndex((v) => v === 'TOTALBUCHAS');
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = normalized.findIndex((v) => v.includes('TOTAL') && v.includes('BUCHAS'));
      }
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = compact.findIndex((v) => v === 'BUCHAPRENSADAS');
      }
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = normalized.findIndex((v) => v.includes('BUCHA') && v.includes('PRENSAD'));
      }

      if (statusCol >= 0 && pressColResolved >= 0 && totalBuchasColResolved >= 0) {
        return { headerRow: headerIndex, statusCol, pressCol: pressColResolved, totalBuchasCol: totalBuchasColResolved };
      }
    }

    const scanLimit = Math.min(aoa.length, 200);
    for (let r = 0; r < scanLimit; r += 1) {
      const row = aoa[r] || [];
      if (!row.length) {
        continue;
      }

      const normalized = row.map((cell) => normalizeText(cell));
      const compact = normalized.map((txt) => txt.replace(/\s+/g, ''));

      const statusCol = normalized.findIndex((v) => v === 'STATUS');
      if (statusCol < 0) {
        continue;
      }

      let pressColResolved = compact.findIndex((v) => v === 'NPRENSA' || v === 'NOPRENSA');
      if (pressColResolved < 0) {
        pressColResolved = normalized.findIndex((v) => v === 'PRENSA' || v.startsWith('PRENSA'));
      }
      if (pressColResolved < 0) {
        pressColResolved = normalized.findIndex((v) => v.startsWith('N') && v.includes('PRENSA'));
      }

      let totalBuchasColResolved = compact.findIndex((v) => v === 'TOTALBUCHAS');
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = normalized.findIndex((v) => v.includes('TOTAL') && v.includes('BUCHAS'));
      }
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = compact.findIndex((v) => v === 'BUCHAPRENSADAS');
      }
      if (totalBuchasColResolved < 0) {
        totalBuchasColResolved = normalized.findIndex((v) => v.includes('BUCHA') && v.includes('PRENSAD'));
      }

      if (pressColResolved < 0 || totalBuchasColResolved < 0) {
        continue;
      }

      return { headerRow: r, statusCol, pressCol: pressColResolved, totalBuchasCol: totalBuchasColResolved };
    }
    return null;
  };

  const presses = [];
  const pressTable = detectMoldagemPressTable();
  if (pressTable) {
    const stats = new Map([
      [1, { press: 1, name: 'Prensa 1', ops: 0, buchas: 0 }],
      [2, { press: 2, name: 'Prensa 2', ops: 0, buchas: 0 }],
      [3, { press: 3, name: 'Prensa 3', ops: 0, buchas: 0 }]
    ]);

    let seenAny = false;
    for (let i = pressTable.headerRow + 1; i < aoa.length; i += 1) {
      const row = aoa[i] || [];
      const statusValue = String(row[pressTable.statusCol] || '').trim();
      const pressValue = row[pressTable.pressCol];
      const buchasValue = row[pressTable.totalBuchasCol];
      const hasAnyData = statusValue || String(pressValue || '').trim() || String(buchasValue || '').trim();
      if (!hasAnyData) {
        if (seenAny) {
          break;
        }
        continue;
      }
      seenAny = true;

      const statusNorm = normalizeText(statusValue);
      if (!statusNorm.includes(MOLDAGEM_STATUS_PRENSADO)) {
        continue;
      }

      const pressNum = parseNumericValue(pressValue);
      if (![1, 2, 3].includes(pressNum)) {
        continue;
      }

      const item = stats.get(pressNum);
      item.ops += 1;
      item.buchas += parseNumericValue(buchasValue);
    }

    presses.push(...Array.from(stats.values()));
  } else {
    const pressHeader = findAoaCellByText(aoa, 'PRENSA');
    if (pressHeader) {
      for (let i = pressHeader.row + 1; i < Math.min(aoa.length, pressHeader.row + 20); i += 1) {
        const row = aoa[i] || [];
        const pressRaw = row[pressHeader.col];
        const pressNum = parseNumericValue(pressRaw);
        const nextOps = parseNumericValue(row[pressHeader.col + 1]);
        const nextBuchas = parseNumericValue(row[pressHeader.col + 2]);
        const hasAnyData = String(pressRaw || '').trim() || nextOps || nextBuchas;
        if (!hasAnyData && presses.length) {
          break;
        }
        if (!pressNum || pressNum < 1) {
          continue;
        }

        presses.push({
          name: `Prensa ${String(pressNum).replace(/\.0+$/, '')}`,
          press: pressNum,
          ops: nextOps,
          buchas: nextBuchas
        });
      }
    }
  }

  const operators = [];
  const operatorsHeader = findAoaCellByText(aoa, 'OPERADORES');
  if (operatorsHeader) {
    for (let i = operatorsHeader.row + 1; i < Math.min(aoa.length, operatorsHeader.row + 20); i += 1) {
      const row = aoa[i] || [];
      const name = String(row[operatorsHeader.col] || '').trim();
      const normalizedName = normalizeText(name);
      const ops = parseNumericValue(row[operatorsHeader.col + 1]);
      const buchas = parseNumericValue(row[operatorsHeader.col + 2]);
      const peso = parseNumericValue(row[operatorsHeader.col + 3]);
      const hasAnyData = name || ops || buchas || peso;
      if (!hasAnyData && operators.length) {
        break;
      }
      if (!name || normalizedName === 'OPERADORES' || normalizedName === 'OPERADOR') {
        continue;
      }
      operators.push({
        name,
        ops,
        buchas,
        peso
      });
      if (operators.length >= 8) {
        break;
      }
    }
  }

  return {
    presses: consolidatePresses1to3(presses).sort((a, b) => a.press - b.press),
    operators
  };
}

function loadPcpMoldagemSnapshot({ date }) {
  const safeDate = date || getTodayYmdLocal();
  const workbookPath = findDailyDashboardWorkbookPath(safeDate);
  const workbook = readDashboardWorkbookSafe(workbookPath);
  const materials = parseDashboardMaterials(workbook);
  const extras = parseMoldagemPressesAndOperators(workbook);

  const kgProcessados = materials.reduce((acc, item) => acc + Number(item.kgUsed || 0), 0);
  const buchasMoldadas = materials.reduce((acc, item) => acc + Number(item.buchas || 0), 0);
  const kgRefugo = materials.reduce((acc, item) => acc + Number(item.refugoKg || 0), 0);

  const materialMaiorSaida = materials
    .filter((item) => Number(item.kgUsed || 0) > 0)
    .reduce((best, item) => {
      if (!best || Number(item.kgUsed || 0) > Number(best.kgUsed || 0)) {
        return item;
      }
      return best;
    }, null);

  const materialMaisBuchas = materials
    .filter((item) => Number(item.buchas || 0) > 0)
    .reduce((best, item) => {
      if (!best || Number(item.buchas || 0) > Number(best.buchas || 0)) {
        return item;
      }
      return best;
    }, null);

  return {
    date: safeDate,
    sourceFile: workbookPath,
    title: `MOLDAGEM - ${safeDate}`,
    kpis: {
      eficMoldagem: 100,
      kgProcessados,
      buchasMoldadas,
      materialMaiorSaida: materialMaiorSaida ? materialMaiorSaida.name : '-',
      materialMaisBuchas: materialMaisBuchas ? materialMaisBuchas.name : '-',
      kgRefugo,
      custoRefugo: 0
    },
    charts: {
      materials,
      presses: extras.presses,
      operators: extras.operators
    }
  };
}

function countOpsGeneratedByDate(dateValue) {
  const selected = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(selected.getTime())) {
    return 0;
  }

  try {
    const bancoOrdensPath = resolveFiscalPath(FISCAL_BANCO_ORDENS_RELATIVE);
    if (!fs.existsSync(bancoOrdensPath)) {
      return 0;
    }

    const { workbook } = readWorkbookCached(bancoOrdensPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
      return 0;
    }

    const aoa = sheetToAoaFast(sheet);
    if (!aoa.length) {
      return 0;
    }

    let headerIndex = -1;
    let entryIdx = -1;
    const scanLimit = Math.min(aoa.length, 30);
    for (let i = 0; i < scanLimit; i += 1) {
      const headers = (aoa[i] || []).map((cell) => normalizeText(cell));
      const idx = headers.findIndex((h) => h.includes('ENTRADA') && h.includes('PCP'));
      if (idx >= 0) {
        headerIndex = i;
        entryIdx = idx;
        break;
      }
    }

    if (entryIdx < 0) {
      return 0;
    }

    let count = 0;
    for (let i = headerIndex + 1; i < aoa.length; i += 1) {
      const cell = aoa[i][entryIdx];
      if (!cell) {
        continue;
      }

      let rowDate = null;
      if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
        rowDate = cell;
      } else if (typeof cell === 'number') {
        const parsed = XLSX.SSF.parse_date_code(cell);
        if (parsed && Number.isFinite(parsed.y) && Number.isFinite(parsed.m) && Number.isFinite(parsed.d)) {
          rowDate = new Date(parsed.y, parsed.m - 1, parsed.d);
        }
      } else {
        const text = String(cell || '').trim();
        if (text) {
          const candidate = new Date(text);
          if (!Number.isNaN(candidate.getTime())) {
            rowDate = candidate;
          } else {
            const brMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            if (brMatch) {
              rowDate = new Date(Number(brMatch[3]), Number(brMatch[2]) - 1, Number(brMatch[1]));
            }
          }
        }
      }

      if (!rowDate || Number.isNaN(rowDate.getTime())) {
        continue;
      }
      if (
        rowDate.getFullYear() === selected.getFullYear()
        && rowDate.getMonth() === selected.getMonth()
        && rowDate.getDate() === selected.getDate()
      ) {
        count += 1;
      }
    }
    return count;
  } catch (error) {
    return 0;
  }
}

function readDashboardWorkbookSafe(workbookPath) {
  const options = { cellDates: true, cellStyles: true };
  try {
    return XLSX.readFile(workbookPath, options);
  } catch (firstError) {
    try {
      const raw = fs.readFileSync(workbookPath);
      return XLSX.read(raw, { type: 'buffer', ...options });
    } catch (secondError) {
      let tempPath = '';
      try {
        tempPath = path.join(app.getPath('temp'), `pcp-dashboard-${Date.now()}.xlsm`);
        fs.copyFileSync(workbookPath, tempPath);
        const rawTemp = fs.readFileSync(tempPath);
        return XLSX.read(rawTemp, { type: 'buffer', ...options });
      } catch (thirdError) {
        const code = (thirdError && thirdError.code) || (secondError && secondError.code) || (firstError && firstError.code) || 'UNKNOWN';
        throw new Error(`Nao consegui ler a planilha do Dashboard (arquivo em uso/bloqueado). Codigo: ${code}`);
      } finally {
        if (tempPath) {
          try {
            if (fs.existsSync(tempPath)) {
              fs.unlinkSync(tempPath);
            }
          } catch (_) {
            // Ignora erro ao limpar arquivo temporario.
          }
        }
      }
    }
  }
}

function sumPlannedPiecesFromTable7(workbook, selectedSheetName = '') {
  const preferredSheet = workbook.SheetNames.find((name) => {
    const n = normalizeText(name).replace(/\s+/g, '');
    return n === 'PLANILHA1';
  });
  const sheetName = preferredSheet || selectedSheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return {
      total: 0,
      sheetName,
      firstDataRow: 15,
      lastDataRow: null,
      rowCount: 0,
      reason: 'sheet_not_found'
    };
  }

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  if (!Array.isArray(aoa) || aoa.length < 15) {
    return {
      total: 0,
      sheetName,
      firstDataRow: 15,
      lastDataRow: null,
      rowCount: 0,
      reason: 'sheet_too_short'
    };
  }

  // Regra fixa informada:
  // linha 14 (index 13) = cabeçalho da Tabela7, coluna B (index 1) = QTD.
  const headerRowIndex = 13;
  const qtdCol = 1;
  const headerText = normalizeText((aoa[headerRowIndex] || [])[qtdCol] || '');
  if (headerText !== 'QTD') {
    return {
      total: 0,
      sheetName,
      firstDataRow: 15,
      lastDataRow: null,
      rowCount: 0,
      reason: `header_not_qtd:${headerText || 'EMPTY'}`
    };
  }

  const firstDataRow = headerRowIndex + 1; // linha 15
  let total = 0;
  let lastDataRow = null;
  for (let i = firstDataRow; i < aoa.length; i += 1) {
    const qtdCurrent = String(((aoa[i] || [])[qtdCol]) || '').trim();
    const qtdNext1 = String((((aoa[i + 1] || [])[qtdCol])) || '').trim();
    const qtdNext2 = String((((aoa[i + 2] || [])[qtdCol])) || '').trim();

    total += parseNumericValue((aoa[i] || [])[qtdCol]);

    // Regra solicitada: a linha atual é a última quando as 2 abaixo estiverem vazias.
    if (!qtdNext1 && !qtdNext2) {
      lastDataRow = i + 1; // 1-based
      break;
    }
  }

  if (lastDataRow == null) {
    lastDataRow = aoa.length;
  }

  return {
    total,
    sheetName,
    firstDataRow: firstDataRow + 1, // 1-based
    lastDataRow,
    rowCount: Math.max(0, lastDataRow - (firstDataRow + 1) + 1),
    reason: 'ok'
  };
}

function loadPcpDashboardSnapshot({ date }) {
  const safeDate = date || getTodayYmdLocal();
  const workbookPath = findDailyDashboardWorkbookPath(safeDate);
  const workbook = readDashboardWorkbookSafe(workbookPath);

  const mainRowsData = loadDashboardPlanilhaRows(workbook);
  const metrics = parseDashboardMetrics(mainRowsData.rows);
  const cnc = parseDashboardCncMetrics(mainRowsData.rows);
  let rnc = null;
  try {
    rnc = loadRncSnapshotForDate(safeDate);
  } catch (error) {
    rnc = {
      kpis: { rncQtd: 0, rncCost: 0, totalCost: 0 },
      usinagem: { machines: [] },
      debug: { error: (error && error.message) || String(error || ''), sourceFile: null, sheetName: null }
    };
  }
  const plannedPiecesByRule = sumPlannedPiecesFromTable7(workbook, mainRowsData.sheetName);
  if (plannedPiecesByRule.total > 0) {
    metrics.kpis.pecasPlanejadas = plannedPiecesByRule.total;
  }
  const materials = parseDashboardMaterials(workbook);
  const opGeradas = countOpsGeneratedByDate(safeDate);
  metrics.kpis.opGeradas = opGeradas;
  metrics.kpis.rncQtd = rnc.kpis.rncQtd;
  metrics.kpis.rncCost = rnc.kpis.rncCost;
  metrics.kpis.totalCost = rnc.kpis.totalCost;

  return {
    date: safeDate,
    sourceFile: workbookPath,
    sourceSheet: mainRowsData.sheetName,
    title: `Dashboard Diário PCP - ${safeDate}`,
    kpis: metrics.kpis,
    charts: {
      opsByCompany: metrics.opsByCompany,
      piecesDaily: {
        feitas: metrics.kpis.pecasFeitas,
        planejadas: metrics.kpis.pecasPlanejadas
      },
      opsDaily: {
        feitas: metrics.kpis.opFeitas,
        planejadas: metrics.kpis.opPlanejadas
      },
      cnc,
      rncUsinagem: rnc.usinagem,
      rncReport: rnc.report,
      acabamentoByOperator: metrics.acabamentoByOperator,
      materials
    },
    debug: {
      plannedPiecesByRule,
      rnc,
      runtime: {
        mainFile: __filename,
        appVersion: typeof app.getVersion === 'function' ? app.getVersion() : ''
      }
    }
  };
}

async function exportPcpDashboardPdf(webContents, payload) {
  if (!webContents) {
    throw new Error('Nao foi possivel acessar a janela para exportacao do dashboard.');
  }
  const win = BrowserWindow.fromWebContents(webContents);

  const suggestedName = String((payload && payload.suggestedName) || '').trim() || 'pcp-dashboard.pdf';
  const defaultPath = path.join(app.getPath('downloads'), suggestedName);
  const saveResult = await dialog.showSaveDialog({
    title: String((payload && payload.dialogTitle) || '').trim() || 'Salvar dashboard em PDF',
    defaultPath,
    filters: [{ name: 'Documento PDF', extensions: ['pdf'] }]
  });

  if (saveResult.canceled || !saveResult.filePath) {
    return { canceled: true };
  }

  const previousBounds = win ? win.getContentBounds() : null;
  const previousZoomFactor = typeof webContents.getZoomFactor === 'function' ? webContents.getZoomFactor() : 1;
  const exportWidth = Number((payload && payload.exportWidth) || 1122);
  const exportHeight = Number((payload && payload.exportHeight) || 794);

  try {
    if (win && Number.isFinite(exportWidth) && Number.isFinite(exportHeight)) {
      win.setContentSize(Math.max(1122, Math.floor(exportWidth)), Math.max(794, Math.floor(exportHeight)));
      await new Promise((resolve) => setTimeout(resolve, 120));
    }

    if (typeof webContents.setZoomFactor === 'function') {
      webContents.setZoomFactor(1);
    }
    await webContents.executeJavaScript('window.scrollTo(0, 0);', true).catch(() => {});
    await new Promise((resolve) => setTimeout(resolve, 90));

    const pdfBuffer = await webContents.printToPDF({
      printBackground: true,
      landscape: true,
      pageSize: 'A4',
      preferCSSPageSize: true,
      margins: { top: 0, bottom: 0, left: 0, right: 0 }
    });

    fs.writeFileSync(saveResult.filePath, pdfBuffer);
    return { canceled: false, filePath: saveResult.filePath };
  } finally {
    if (typeof webContents.setZoomFactor === 'function') {
      webContents.setZoomFactor(previousZoomFactor || 1);
    }
    if (win && previousBounds) {
      win.setContentBounds(previousBounds);
    }
  }
}

async function exportPcpEfficiencyImage(webContents, payload) {
  if (!webContents) {
    throw new Error('Nao foi possivel acessar a janela para exportacao.');
  }
  const win = BrowserWindow.fromWebContents(webContents);
  if (!win) {
    throw new Error('Janela principal nao encontrada para exportacao.');
  }

  const rawRect = payload && payload.rect ? payload.rect : {};
  const selector = String((payload && payload.selector) || '').trim();
  const x = Math.max(0, Math.floor(Number(rawRect.x || 0)));
  const y = Math.max(0, Math.floor(Number(rawRect.y || 0)));
  const width = Math.max(1, Math.floor(Number(rawRect.width || 0)));
  const height = Math.max(1, Math.floor(Number(rawRect.height || 0)));
  if (width <= 1 || height <= 1) {
    throw new Error('Area de exportacao invalida.');
  }

  const oldContentBounds = win.getContentBounds();
  const autoFit = !!(payload && payload.autoFit);

  const getRectFromSelector = async () => {
    if (!selector) {
      return null;
    }

    return webContents.executeJavaScript(
      `(() => {
        const el = document.querySelector(${JSON.stringify(selector)});
        if (!el) return null;
        const rect = el.getBoundingClientRect();
        const fullWidth = Math.max(Math.ceil(rect.width), Math.ceil(el.scrollWidth || 0), Math.ceil(el.offsetWidth || 0));
        const fullHeight = Math.max(Math.ceil(rect.height), Math.ceil(el.scrollHeight || 0), Math.ceil(el.offsetHeight || 0));
        return {
          x: Math.max(0, Math.floor(rect.left + window.scrollX)),
          y: Math.max(0, Math.floor(rect.top + window.scrollY)),
          width: Math.max(1, fullWidth),
          height: Math.max(1, fullHeight)
        };
      })();`,
      true
    );
  };

  try {
    if (autoFit) {
      const targetWidth = Math.max(
        oldContentBounds.width,
        Math.ceil(Number((payload && payload.targetWidth) || 0)),
        x + width + 24
      );
      const targetHeight = Math.max(
        oldContentBounds.height,
        Math.ceil(Number((payload && payload.targetHeight) || 0)),
        y + height + 24
      );

      // Expande temporariamente a area de render para capturar o painel inteiro.
      win.setContentSize(targetWidth, targetHeight);
      await new Promise((resolve) => setTimeout(resolve, 120));
    }

    let finalRect = { x, y, width, height };
    const firstRect = await getRectFromSelector();
    if (firstRect && Number.isFinite(firstRect.width) && Number.isFinite(firstRect.height)) {
      const needWidth = Math.max(oldContentBounds.width, firstRect.x + firstRect.width + 340);
      const needHeight = Math.max(oldContentBounds.height, firstRect.y + firstRect.height + 320);
      if (needWidth > win.getContentBounds().width || needHeight > win.getContentBounds().height) {
        win.setContentSize(Math.ceil(needWidth), Math.ceil(needHeight));
        await new Promise((resolve) => setTimeout(resolve, 120));
      }

      const secondRect = await getRectFromSelector();
      const sourceRect = secondRect || firstRect;
      const padLeft = 20;
      const padTop = 20;
      const padRight = 220;
      const padBottom = 220;
      finalRect = {
        x: Math.max(0, Math.floor((sourceRect.x || 0) - padLeft)),
        y: Math.max(0, Math.floor((sourceRect.y || 0) - padTop)),
        width: Math.max(1, Math.floor((sourceRect.width || 0) + padLeft + padRight)),
        height: Math.max(1, Math.floor((sourceRect.height || 0) + padTop + padBottom))
      };
    }

    let pngBuffer = null;
    const debuggerClient = webContents.debugger;
    const alreadyAttached = debuggerClient.isAttached();
    try {
      if (!alreadyAttached) {
        debuggerClient.attach('1.3');
      }
      await debuggerClient.sendCommand('Page.enable');
      const shot = await debuggerClient.sendCommand('Page.captureScreenshot', {
        format: 'png',
        captureBeyondViewport: true,
        fromSurface: true,
        clip: {
          x: Number(finalRect.x || 0),
          y: Number(finalRect.y || 0),
          width: Number(finalRect.width || 1),
          height: Number(finalRect.height || 1),
          scale: 1
        }
      });
      if (shot && shot.data) {
        pngBuffer = Buffer.from(shot.data, 'base64');
      }
    } catch (captureError) {
      const fallback = await webContents.capturePage(finalRect);
      pngBuffer = fallback.toPNG();
    } finally {
      if (!alreadyAttached && debuggerClient.isAttached()) {
        debuggerClient.detach();
      }
    }

    if (!pngBuffer || !pngBuffer.length) {
      throw new Error('Falha ao gerar imagem de exportacao.');
    }

    const suggestedName = String((payload && payload.suggestedName) || '').trim() || 'pcp-eficiencia.png';
    const defaultPath = path.join(app.getPath('downloads'), suggestedName);
    const saveResult = await dialog.showSaveDialog({
      title: 'Salvar imagem de eficiencia',
      defaultPath,
      filters: [{ name: 'Imagem PNG', extensions: ['png'] }]
    });

    if (saveResult.canceled || !saveResult.filePath) {
      return { canceled: true };
    }

    fs.writeFileSync(saveResult.filePath, pngBuffer);
    return { canceled: false, filePath: saveResult.filePath };
  } finally {
    if (autoFit) {
      win.setContentSize(oldContentBounds.width, oldContentBounds.height);
    }
  }
}

function createWindow() {
  const iconCandidates = [
    path.join(__dirname, '..', 'img', 'KuruJossIcon.ico'),
    path.join(__dirname, '..', 'img', 'KuruJossIcon.png')
  ];
  const iconPath = iconCandidates.find((candidate) => fs.existsSync(candidate)) || null;

  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 980,
    minHeight: 680,
    ...(iconPath ? { icon: iconPath } : {}),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.maximize();
}

let FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE = {
  filePath: '',
  mtimeMs: -1,
  pedidoToReqs: null,
  reqBaseToPedidos: null
};

let FISCAL_BANCO_PEDIDOS_ROWS_CACHE = { filePath: '', mtimeMs: -1, result: null };
let FISCAL_BANCO_PEDIDOS_LITE_CACHE = { filePath: '', mtimeMs: -1, rows: null, looksLikeHeader: false };
let RH_SNAPSHOT_CACHE = { filePath: '', mtimeMs: -1, snapshot: null };

function getFiscalBancoOrdensIndex() {
  const bancoOrdensPath = resolveFiscalPath(FISCAL_BANCO_ORDENS_RELATIVE);
  if (!fs.existsSync(bancoOrdensPath)) {
    throw new Error(`BancoDeOrdens nao encontrado: ${bancoOrdensPath}`);
  }
  const mtimeMs = getFileMtimeMs(bancoOrdensPath);
  if (
    FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.pedidoToReqs
    && FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.reqBaseToPedidos
    && FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.filePath === bancoOrdensPath
    && FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.mtimeMs === mtimeMs
  ) {
    return {
      pedidoToReqs: FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.pedidoToReqs,
      reqBaseToPedidos: FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE.reqBaseToPedidos
    };
  }

  const diskCache = readFastCache('fiscal-banco-ordens-index', bancoOrdensPath, mtimeMs);
  if (diskCache && Array.isArray(diskCache.pedidoToReqs) && Array.isArray(diskCache.reqBaseToPedidos)) {
    const pedidoToReqs = new Map(
      diskCache.pedidoToReqs.map(([pedido, reqs]) => [String(pedido), new Set((reqs || []).map((x) => String(x)))])
    );
    const reqBaseToPedidos = new Map(
      diskCache.reqBaseToPedidos.map(([reqBase, pedidos]) => [String(reqBase), new Set((pedidos || []).map((x) => String(x)))])
    );
    FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE = {
      filePath: bancoOrdensPath,
      mtimeMs,
      pedidoToReqs,
      reqBaseToPedidos
    };
    return { pedidoToReqs, reqBaseToPedidos };
  }

  const { workbook } = readWorkbookCached(bancoOrdensPath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Planilha BancoDeOrdens sem abas validas.');
  }

  const rows = sheetToAoaFast(worksheet);
  const pedidoToReqs = new Map();
  const reqBaseToPedidos = new Map();
  rows.forEach((row) => {
    const pedido = String(row[3] || '').trim();
    if (!pedido) {
      return;
    }
    const req = String(row[4] || '').trim();
    const reqBase = extractReqBase(req);
    const reqKey = String(req || reqBase || '').trim();
    if (!reqKey) {
      return;
    }
    if (!pedidoToReqs.has(pedido)) {
      pedidoToReqs.set(pedido, new Set());
    }
    pedidoToReqs.get(pedido).add(reqKey);
    if (reqBase) {
      if (!reqBaseToPedidos.has(reqBase)) {
        reqBaseToPedidos.set(reqBase, new Set());
      }
      reqBaseToPedidos.get(reqBase).add(pedido);
    }
  });

  FISCAL_BANCO_ORDENS_REQ_INDEX_CACHE = {
    filePath: bancoOrdensPath,
    mtimeMs,
    pedidoToReqs,
    reqBaseToPedidos
  };
  writeFastCache('fiscal-banco-ordens-index', bancoOrdensPath, mtimeMs, {
    pedidoToReqs: Array.from(pedidoToReqs.entries()).map(([pedido, set]) => [pedido, Array.from(set)]),
    reqBaseToPedidos: Array.from(reqBaseToPedidos.entries()).map(([reqBase, set]) => [reqBase, Array.from(set)])
  });
  return { pedidoToReqs, reqBaseToPedidos };
}

function getFiscalPedidoToReqsIndex() {
  return getFiscalBancoOrdensIndex().pedidoToReqs;
}

function getFiscalReqBaseToPedidosIndex() {
  return getFiscalBancoOrdensIndex().reqBaseToPedidos;
}

function loadFiscalItemsFromBancoOrdens(identifiers) {
  const list = Array.isArray(identifiers) ? identifiers : [];
  const normalized = list.map((value) => String(value || '').trim()).filter(Boolean);
  if (!normalized.length) {
    return [];
  }

  const tokenSet = new Set(normalized);
  const bancoOrdensPath = resolveFiscalPath(FISCAL_BANCO_ORDENS_RELATIVE);
  if (!fs.existsSync(bancoOrdensPath)) {
    throw new Error(`BancoDeOrdens nao encontrado: ${bancoOrdensPath}`);
  }

  const { workbook } = readWorkbookCached(bancoOrdensPath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Planilha BancoDeOrdens sem abas validas.');
  }

  const rows = sheetToAoaFast(worksheet);
  const items = [];

  rows.forEach((row) => {
    const pedido = String(row[3] || '').trim();
    const req = String(row[4] || '').trim();
    const reqBase = extractReqBase(req);
    if (!pedido && !reqBase) {
      return;
    }

    const matchesPedido = pedido && tokenSet.has(pedido);
    const matchesReq = reqBase && tokenSet.has(reqBase);
    const matchesReqFull = req && tokenSet.has(req);
    if (!matchesPedido && !matchesReq && !matchesReqFull) {
      return;
    }

    items.push({
      qtd: row[1] ?? '',
      produto: String(row[2] || '').trim(),
      pedido,
      req: req || reqBase,
      dataEntrada: formatExcelDatePtBr(row[13])
    });
  });

  return items;
}

function appendRowsToBancoPedidos(rowsToAppend) {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const { workbook } = readWorkbookCached(bancoPedidosPath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Banco Pedidos sem abas validas.');
  }

  const existing = sheetToAoaFast(worksheet);
  let lastUsedRow = 0;
  existing.forEach((row, index) => {
    const hasValue = Array.isArray(row) && row.some((cell) => String(cell || '').trim() !== '');
    if (hasValue) {
      lastUsedRow = index + 1;
    }
  });

  const startRow = lastUsedRow + 1;
  XLSX.utils.sheet_add_aoa(worksheet, rowsToAppend, { origin: `A${startRow}` });
  writeWorkbookFile(workbook, bancoPedidosPath);
}

function parseBancoPedidosRowsFromAoa(aoa) {
  const safeAoa = Array.isArray(aoa) ? aoa : [];
  const headerRow = safeAoa[0] || [];
  const looksLikeHeader = normalizeText(headerRow[0]) === 'QNTD' && normalizeText(headerRow[3]) === 'CLIENTE';
  const startIndex = looksLikeHeader ? 1 : 0;

  const rows = [];
  for (let i = startIndex; i < safeAoa.length; i += 1) {
    const row = safeAoa[i] || [];
    const hasValue = row.some((cell) => String(cell || '').trim() !== '');
    if (!hasValue) {
      continue;
    }

    rows.push({
      rowIndex: i, // 0-based in AOA (and worksheet)
      qtd: row[0] ?? '',
      produto: row[1] ?? '',
      pedido: row[2] ?? '',
      cliente: row[3] ?? '',
      status: row[4] ?? '',
      dataEntrada: row[5] ?? '',
      dataDespache: row[6] ?? '',
      nf: row[7] ?? '',
      rastreio: row[8] ?? '',
      dataFaturada: row[9] ?? '',
      baixadoPor: row[10] ?? '',
      horario: row[11] ?? ''
    });
  }

  return { rows, looksLikeHeader };
}

function readBancoPedidosRowsLite() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const mtimeMs = getFileMtimeMs(bancoPedidosPath);
  if (
    Array.isArray(FISCAL_BANCO_PEDIDOS_LITE_CACHE.rows)
    && FISCAL_BANCO_PEDIDOS_LITE_CACHE.filePath === bancoPedidosPath
    && FISCAL_BANCO_PEDIDOS_LITE_CACHE.mtimeMs === mtimeMs
  ) {
    return {
      bancoPedidosPath,
      mtimeMs,
      rows: FISCAL_BANCO_PEDIDOS_LITE_CACHE.rows,
      looksLikeHeader: !!FISCAL_BANCO_PEDIDOS_LITE_CACHE.looksLikeHeader
    };
  }

  const diskCache = readFastCache('fiscal-banco-pedidos-rows', bancoPedidosPath, mtimeMs);
  if (diskCache && Array.isArray(diskCache.rows)) {
    const rows = diskCache.rows;
    const looksLikeHeader = !!diskCache.looksLikeHeader;
    FISCAL_BANCO_PEDIDOS_LITE_CACHE = { filePath: bancoPedidosPath, mtimeMs, rows, looksLikeHeader };
    return { bancoPedidosPath, mtimeMs, rows, looksLikeHeader };
  }

  const { workbook } = readWorkbookCached(bancoPedidosPath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Banco Pedidos sem abas validas.');
  }

  const aoa = sheetToAoaFast(worksheet);
  const parsed = parseBancoPedidosRowsFromAoa(aoa);
  FISCAL_BANCO_PEDIDOS_LITE_CACHE = {
    filePath: bancoPedidosPath,
    mtimeMs,
    rows: parsed.rows,
    looksLikeHeader: parsed.looksLikeHeader
  };
  writeFastCache('fiscal-banco-pedidos-rows', bancoPedidosPath, mtimeMs, {
    looksLikeHeader: parsed.looksLikeHeader,
    rows: parsed.rows
  });
  return { bancoPedidosPath, mtimeMs, rows: parsed.rows, looksLikeHeader: parsed.looksLikeHeader };
}

function readBancoPedidosRows() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const { workbook, mtimeMs } = readWorkbookCached(bancoPedidosPath);
  if (
    FISCAL_BANCO_PEDIDOS_ROWS_CACHE.result
    && FISCAL_BANCO_PEDIDOS_ROWS_CACHE.filePath === bancoPedidosPath
    && FISCAL_BANCO_PEDIDOS_ROWS_CACHE.mtimeMs === mtimeMs
  ) {
    return FISCAL_BANCO_PEDIDOS_ROWS_CACHE.result;
  }
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Banco Pedidos sem abas validas.');
  }

  const aoa = sheetToAoaFast(worksheet);
  const parsed = parseBancoPedidosRowsFromAoa(aoa);

  const result = {
    workbook,
    sheetName,
    worksheet,
    rows: parsed.rows,
    bancoPedidosPath,
    looksLikeHeader: parsed.looksLikeHeader
  };
  FISCAL_BANCO_PEDIDOS_ROWS_CACHE = { filePath: bancoPedidosPath, mtimeMs, result };
  FISCAL_BANCO_PEDIDOS_LITE_CACHE = {
    filePath: bancoPedidosPath,
    mtimeMs,
    rows: parsed.rows,
    looksLikeHeader: parsed.looksLikeHeader
  };
  writeFastCache('fiscal-banco-pedidos-rows', bancoPedidosPath, mtimeMs, {
    looksLikeHeader: parsed.looksLikeHeader,
    rows: parsed.rows
  });
  return result;
}

function listNfsFromBancoPedidos() {
  const { rows } = readBancoPedidosRowsLite();
  const map = new Map();

  rows.forEach((row) => {
    const nf = String(row.nf || '').trim();
    if (!nf) {
      return;
    }

    if (!map.has(nf)) {
      map.set(nf, {
        nf,
        cliente: String(row.cliente || '').trim(),
        status: String(row.status || '').trim(),
        dataFaturada: String(row.dataFaturada || '').trim(),
        rastreio: String(row.rastreio || '').trim(),
        dataDespache: String(row.dataDespache || '').trim(),
        pedidos: new Set(),
        itens: 0
      });
    }

    const entry = map.get(nf);
    entry.itens += 1;
    const pedido = String(row.pedido || '').trim();
    if (pedido) {
      entry.pedidos.add(pedido);
    }

    const status = String(row.status || '').trim();
    if (status) {
      entry.status = status;
    }
    const dataFaturada = String(row.dataFaturada || '').trim();
    if (dataFaturada) {
      entry.dataFaturada = dataFaturada;
    }
    const rastreio = String(row.rastreio || '').trim();
    if (rastreio) {
      entry.rastreio = rastreio;
    }
    const dataDespache = String(row.dataDespache || '').trim();
    if (dataDespache) {
      entry.dataDespache = dataDespache;
    }
    const cliente = String(row.cliente || '').trim();
    if (cliente && !entry.cliente) {
      entry.cliente = cliente;
    }
  });

  let pedidoToReqs = null;
  try {
    pedidoToReqs = getFiscalPedidoToReqsIndex();
  } catch (error) {
    pedidoToReqs = null;
  }

  const list = Array.from(map.values()).map((entry) => {
    const pedidos = Array.from(entry.pedidos || []);
    pedidos.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
    const pedidosLabel = pedidos.length > 10
      ? `${pedidos.slice(0, 10).join(' | ')} | +${pedidos.length - 10}`
      : pedidos.join(' | ');

    let reqsLabel = '';
    if (pedidoToReqs) {
      const reqSet = new Set();
      pedidos.forEach((pedido) => {
        const reqs = pedidoToReqs.get(String(pedido));
        if (!reqs) {
          return;
        }
        reqs.forEach((req) => reqSet.add(String(req)));
      });
      const reqs = Array.from(reqSet).map((req) => String(req || '').trim()).filter(Boolean);
      reqs.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
      reqsLabel = reqs.length > 10
        ? `${reqs.slice(0, 10).join(' | ')} | +${reqs.length - 10}`
        : reqs.join(' | ');
    }

    return {
      nf: entry.nf,
      cliente: entry.cliente,
      status: entry.status,
      dataFaturada: entry.dataFaturada,
      rastreio: entry.rastreio,
      dataDespache: entry.dataDespache,
      itens: entry.itens,
      pedidosLabel,
      reqsLabel
    };
  });
  list.sort((a, b) => String(b.nf).localeCompare(String(a.nf), 'pt-BR', { numeric: true }));
  return list;
}

function getNfItemsFromBancoPedidos(nf) {
  const target = String(nf || '').trim();
  if (!target) {
    return [];
  }

  const { rows } = readBancoPedidosRowsLite();
  return rows
    .filter((row) => String(row.nf || '').trim() === target)
    .map((row) => ({
      qtd: row.qtd ?? '',
      produto: row.produto ?? '',
      pedido: row.pedido ?? '',
      cliente: row.cliente ?? '',
      status: row.status ?? '',
      dataEntrada: row.dataEntrada ?? '',
      dataDespache: row.dataDespache ?? '',
      nf: row.nf ?? '',
      rastreio: row.rastreio ?? '',
      dataFaturada: row.dataFaturada ?? '',
      baixadoPor: row.baixadoPor ?? '',
      horario: row.horario ?? ''
    }));
}

function updateNfInBancoPedidos({ nf, status, rastreio, dataDespache, editorUser, editorAt }) {
  const target = String(nf || '').trim();
  if (!target) {
    throw new Error('NF invalida.');
  }
  const editor = String(editorUser || '').trim();
  if (!editor) {
    throw new Error('Usuario do editor obrigatorio.');
  }

  const dt = editorAt instanceof Date ? editorAt : new Date(editorAt);
  if (Number.isNaN(dt.getTime())) {
    throw new Error('Data/hora do editor invalida.');
  }
  const horario = dt.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

  const { workbook, sheetName, worksheet, bancoPedidosPath } = readBancoPedidosRows();
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');

  let updated = 0;
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    const nfCell = getCell(worksheet, r, 7);
    const nfValue = String((nfCell && nfCell.v) || '').trim();
    if (!nfValue || nfValue !== target) {
      continue;
    }

    if (status != null) {
      setCell(worksheet, r, 4, status);
    }
    if (dataDespache != null) {
      setCell(worksheet, r, 6, dataDespache);
    }
    if (rastreio != null) {
      setCell(worksheet, r, 8, rastreio);
    }

    setCell(worksheet, r, 10, editor);
    setCell(worksheet, r, 11, horario);
    updated += 1;
  }

  if (!updated) {
    throw new Error(`NF nao encontrada no banco de pedidos: ${target}`);
  }

  workbook.Sheets[sheetName] = worksheet;
  writeWorkbookFile(workbook, bancoPedidosPath);
  return updated;
}

function normalizeId(value) {
  return String(value == null ? '' : value).trim();
}

function buildBancoPedidosIndexes(rows) {
  const nfSet = new Set();
  const pedidoSet = new Set();
  const itemKeySet = new Set();

  rows.forEach((row) => {
    const nf = normalizeId(row.nf);
    if (nf) {
      nfSet.add(nf);
    }

    const pedido = normalizeId(row.pedido);
    if (pedido) {
      pedidoSet.add(pedido);
    }

    const qtd = normalizeId(row.qtd);
    const produto = normalizeId(row.produto);
    if (pedido && (qtd || produto)) {
      itemKeySet.add(`${pedido}||${qtd}||${produto}`);
    }
  });

  return { nfSet, pedidoSet, itemKeySet };
}

function validateFiscalCadastro({ nf, itemsToInsert }) {
  const { rows } = readBancoPedidosRowsLite();
  const { nfSet, pedidoSet, itemKeySet } = buildBancoPedidosIndexes(rows);

  const normalizedNf = normalizeId(nf);
  if (normalizedNf && nfSet.has(normalizedNf)) {
    throw new Error(`NF ${normalizedNf} ja existe no banco de pedidos.`);
  }

  const pedidosDuplicados = new Set();
  const itensDuplicados = [];

  (itemsToInsert || []).forEach((item) => {
    const pedido = normalizeId(item.pedido);
    const qtd = normalizeId(item.qtd);
    const produto = normalizeId(item.produto);

    if (pedido && pedidoSet.has(pedido)) {
      pedidosDuplicados.add(pedido);
    }

    if (pedido) {
      const key = `${pedido}||${qtd}||${produto}`;
      if (itemKeySet.has(key)) {
        itensDuplicados.push({ pedido, qtd, produto });
      }
    }
  });

  if (pedidosDuplicados.size) {
    const list = Array.from(pedidosDuplicados).slice(0, 10).join(', ');
    const suffix = pedidosDuplicados.size > 10 ? ` (+${pedidosDuplicados.size - 10})` : '';
    throw new Error(`Ja existem pedidos cadastrados no banco: ${list}${suffix}.`);
  }

  if (itensDuplicados.length) {
    const preview = itensDuplicados
      .slice(0, 5)
      .map((x) => `${x.pedido} - ${x.qtd} - ${x.produto}`)
      .join(' | ');
    const suffix = itensDuplicados.length > 5 ? ` (+${itensDuplicados.length - 5})` : '';
    throw new Error(`Ja existem itens cadastrados (duplicados): ${preview}${suffix}.`);
  }
}

function getOrCreateSheet(workbook, name) {
  if (workbook.Sheets[name]) {
    return workbook.Sheets[name];
  }
  const ws = XLSX.utils.aoa_to_sheet([]);
  workbook.SheetNames.push(name);
  workbook.Sheets[name] = ws;
  return ws;
}

function appendHistoryRow(workbook, row) {
  const sheetName = 'FISCAL_HISTORICO';
  const ws = getOrCreateSheet(workbook, sheetName);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });

  if (!aoa.length) {
    aoa.push(['DATA', 'HORA', 'ACAO', 'NF', 'CLIENTE', 'PEDIDOS', 'ITENS', 'USUARIO', 'MOTIVO']);
  }

  aoa.push(row);
  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
}

function appendDeletedDetailRows(workbook, deletedRows, meta) {
  const sheetName = 'FISCAL_NF_APAGADAS';
  const ws = getOrCreateSheet(workbook, sheetName);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });

  if (!aoa.length) {
    aoa.push([
      'DATA',
      'HORA',
      'USUARIO',
      'MOTIVO',
      'QNTD',
      'PRODUTO',
      'PEDIDO',
      'CLIENTE',
      'STATUS',
      'DATA ENTRADA',
      'DATA DESPACHE',
      'NF',
      'RASTREIO',
      'DATA FATURADA',
      'BAIXADO POR',
      'HORARIO'
    ]);
  }

  deletedRows.forEach((row) => {
    aoa.push([
      meta.data,
      meta.hora,
      meta.usuario,
      meta.motivo,
      row[0] ?? '',
      row[1] ?? '',
      row[2] ?? '',
      row[3] ?? '',
      row[4] ?? '',
      row[5] ?? '',
      row[6] ?? '',
      row[7] ?? '',
      row[8] ?? '',
      row[9] ?? '',
      row[10] ?? '',
      row[11] ?? ''
    ]);
  });

  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
}

function deleteNfFromBancoPedidos({ nf, reason, editorUser, editorAt }) {
  const target = String(nf || '').trim();
  if (!target) {
    throw new Error('NF invalida.');
  }
  const motivo = String(reason || '').trim();
  if (!motivo) {
    throw new Error('Motivo obrigatorio para apagar a NF.');
  }
  const editor = String(editorUser || '').trim();
  if (!editor) {
    throw new Error('Usuario do editor obrigatorio.');
  }

  const dt = editorAt instanceof Date ? editorAt : new Date(editorAt);
  if (Number.isNaN(dt.getTime())) {
    throw new Error('Data/hora do editor invalida.');
  }
  const data = dt.toLocaleDateString('pt-BR');
  const hora = dt.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

  const { workbook, sheetName, worksheet, bancoPedidosPath } = readBancoPedidosRows();
  const aoa = sheetToAoaFast(worksheet);

  const headerRow = aoa[0] || [];
  const hasHeader = normalizeText(headerRow[0]) === 'QNTD' && normalizeText(headerRow[3]) === 'CLIENTE';
  const startIndex = hasHeader ? 1 : 0;

  const kept = hasHeader ? [aoa[0]] : [];
  const deleted = [];

  for (let i = startIndex; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const nfValue = String(row[7] || '').trim();
    if (nfValue === target) {
      deleted.push(row);
    } else {
      kept.push(row);
    }
  }

  if (!deleted.length) {
    throw new Error(`NF nao encontrada no banco de pedidos: ${target}`);
  }

  const pedidos = Array.from(new Set(deleted.map((row) => String(row[2] || '').trim()).filter(Boolean)));
  pedidos.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
  const cliente = String((deleted.find((row) => String(row[3] || '').trim()) || [])[3] || '').trim();

  appendHistoryRow(workbook, [data, hora, 'APAGAR_NF', target, cliente, pedidos.join(' | '), deleted.length, editor, motivo]);
  appendDeletedDetailRows(workbook, deleted, { data, hora, usuario: editor, motivo });

  const nextWs = XLSX.utils.aoa_to_sheet(kept);
  workbook.Sheets[sheetName] = nextWs;
  writeWorkbookFile(workbook, bancoPedidosPath);

  return deleted.length;
}

function listDeletedNfHistory() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const { workbook } = readWorkbookCached(bancoPedidosPath);
  const ws = workbook.Sheets['FISCAL_HISTORICO'];
  if (!ws) {
    return [];
  }

  const aoa = sheetToAoaFast(ws);
  if (!aoa.length) {
    return [];
  }

  const startIndex = 1;
  const rows = [];
  for (let i = startIndex; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const acao = String(row[2] || '').trim();
    if (acao !== 'APAGAR_NF') {
      continue;
    }

    rows.push({
      data: String(row[0] || '').trim(),
      hora: String(row[1] || '').trim(),
      nf: String(row[3] || '').trim(),
      usuario: String(row[7] || '').trim(),
      motivo: String(row[8] || '').trim()
    });
  }

  rows.reverse();
  return rows;
}

function findNfsByPedidoOrReq(identifier) {
  const token = String(identifier || '').trim();
  if (!token) {
    return [];
  }

  const { rows } = readBancoPedidosRowsLite();
  const pedidoToNfs = new Map();
  rows.forEach((row) => {
    const pedido = String(row.pedido || '').trim();
    const nf = String(row.nf || '').trim();
    if (!pedido || !nf) {
      return;
    }
    if (!pedidoToNfs.has(pedido)) {
      pedidoToNfs.set(pedido, new Set());
    }
    pedidoToNfs.get(pedido).add(nf);
  });

  if (pedidoToNfs.has(token)) {
    return Array.from(pedidoToNfs.get(token));
  }

  const reqBase = extractReqBase(token);
  if (!reqBase) {
    return [];
  }

  let pedidos = [];
  try {
    const reqBaseToPedidos = getFiscalReqBaseToPedidosIndex();
    const pedidoSet = reqBaseToPedidos.get(reqBase);
    pedidos = pedidoSet ? Array.from(pedidoSet) : [];
  } catch (error) {
    pedidos = [];
  }

  const nfSet = new Set();
  pedidos.forEach((pedido) => {
    const set = pedidoToNfs.get(pedido);
    if (!set) {
      return;
    }
    set.forEach((nf) => nfSet.add(nf));
  });

  return Array.from(nfSet);
}

app.whenReady().then(() => {
  ensureDefaultAdminUser();
  ipcMain.handle('settings:get', () => loadSettings());

  ipcMain.handle('settings:save', (_, nextSettings) => {
    return saveSettings(nextSettings || {});
  });

  ipcMain.handle('settings:selectRootFolder', async () => {
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory']
    });

    if (result.canceled || !result.filePaths || !result.filePaths[0]) {
      return null;
    }

    const selected = result.filePaths[0];
    saveSettings({ rootFolder: selected });
    return selected;
  });

  ipcMain.handle('settings:selectFiscalRootFolder', async () => {
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory']
    });

    if (result.canceled || !result.filePaths || !result.filePaths[0]) {
      return null;
    }

    const selected = result.filePaths[0];
    saveSettings({ fiscalRoot: selected });
    return selected;
  });

  ipcMain.handle('users:list', () => {
    ensureDefaultAdminUser();
    return { users: listUsers() };
  });

  ipcMain.handle('users:create', (_, payload) => {
    if (!payload || payload.adminPassword !== ADMIN_SETTINGS_PASSWORD) {
      throw new Error('Senha administrativa invalida para alterar logins.');
    }
    const username = String((payload && payload.username) || '').trim();
    const password = String((payload && payload.password) || '').trim();
    const permissions = payload && payload.permissions && typeof payload.permissions === 'object'
      ? payload.permissions
      : {};
    const created = createUser({ username, password, permissions });
    return { user: created };
  });

  ipcMain.handle('users:update', (_, payload) => {
    if (!payload || payload.adminPassword !== ADMIN_SETTINGS_PASSWORD) {
      throw new Error('Senha administrativa invalida para alterar logins.');
    }
    const username = String((payload && payload.username) || '').trim();
    const permissions = payload && payload.permissions && typeof payload.permissions === 'object'
      ? payload.permissions
      : {};
    const updated = updateUserPermissions({ username, permissions });
    return { user: updated };
  });

  ipcMain.handle('users:setPassword', (_, payload) => {
    if (!payload || payload.adminPassword !== ADMIN_SETTINGS_PASSWORD) {
      throw new Error('Senha administrativa invalida para alterar logins.');
    }
    const username = String((payload && payload.username) || '').trim();
    const password = String((payload && payload.password) || '').trim();
    const result = setUserPassword({ username, password });
    return { user: result };
  });

  ipcMain.handle('users:delete', (_, payload) => {
    if (!payload || payload.adminPassword !== ADMIN_SETTINGS_PASSWORD) {
      throw new Error('Senha administrativa invalida para alterar logins.');
    }
    const username = String((payload && payload.username) || '').trim();
    deleteUser({ username });
    return { ok: true };
  });

  ipcMain.handle('auth:verify', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const password = String((payload && payload.password) || '').trim();
    const requireFiscal = !!(payload && payload.requireFiscal);
    ensureDefaultAdminUser();
    return verifyLogin({ username, password, requireFiscal });
  });

  ipcMain.handle('list:load', (_, payload) => {
    return loadListData(payload || {});
  });

  ipcMain.handle('freight:summary', (_, payload) => {
    return loadFreightSummary(payload || {});
  });

  ipcMain.handle('pcp:efficiency', (_, payload) => {
    return loadPcpEfficiencySnapshot(payload || {});
  });

  ipcMain.handle('pcp:efficiency:export-image', async (event, payload) => {
    return exportPcpEfficiencyImage(event.sender, payload || {});
  });

  ipcMain.handle('pcp:dashboard', (_, payload) => {
    return loadPcpDashboardSnapshot(payload || {});
  });

  ipcMain.handle('pcp:moldagem', (_, payload) => {
    return loadPcpMoldagemSnapshot(payload || {});
  });

  ipcMain.handle('pcp:dashboard:export-pdf', async (event, payload) => {
    return exportPcpDashboardPdf(event.sender, payload || {});
  });

  ipcMain.handle('rh:context', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    return {
      ok: true,
      username: ctx.username,
      role: ctx.role,
      allowedSectorId: ctx.allowedSectorId,
      canEdit: ctx.canEdit,
      dbFilePath: ctx.dbFilePath,
      config: ctx.config
    };
  });

  ipcMain.handle('rh:setores:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    return { setores: listRhSetoresFromSnapshot(ctx, snapshot) };
  });

  ipcMain.handle('rh:setores:upsert', (_, payload) => {
    return upsertRhSetor({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim(),
      nome_setor: String((payload && payload.nome_setor) || '').trim(),
      gestor_responsavel: String((payload && payload.gestor_responsavel) || '').trim(),
      centro_custo: String((payload && payload.centro_custo) || '').trim()
    });
  });

  ipcMain.handle('rh:setores:delete', (_, payload) => {
    return deleteRhSetor({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:colaboradores:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    return { colaboradores: listRhColaboradoresFromSnapshot(ctx, snapshot) };
  });

  ipcMain.handle('rh:colaboradores:upsert', (_, payload) => {
    return upsertRhColaborador({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim(),
      nome: String((payload && payload.nome) || '').trim(),
      matricula: String((payload && payload.matricula) || '').trim(),
      setor_id: String((payload && payload.setor_id) || '').trim(),
      cargo: String((payload && payload.cargo) || '').trim(),
      status: String((payload && payload.status) || '').trim(),
      data_admissao: String((payload && payload.data_admissao) || '').trim(),
      limite_mensal: String((payload && payload.limite_mensal) || '').trim()
    });
  });

  ipcMain.handle('rh:colaboradores:toggleStatus', (_, payload) => {
    return toggleRhColaboradorStatus({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:colaboradores:delete', (_, payload) => {
    return deleteRhColaborador({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:he:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    const month = String((payload && payload.month) || '').trim();
    const setorId = String((payload && payload.setorId) || '').trim();
    const colaboradorId = String((payload && payload.colaboradorId) || '').trim();
    const tipoHora = String((payload && payload.tipoHora) || '').trim();
    const rows = listRhHorasExtrasFromSnapshot(ctx, snapshot, { month, setorId, colaboradorId, tipoHora });
    return { rows };
  });

  ipcMain.handle('rh:atrasos:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    const month = String((payload && payload.month) || '').trim();
    const setorId = String((payload && payload.setorId) || '').trim();
    const colaboradorId = String((payload && payload.colaboradorId) || '').trim();
    const rows = listRhAtrasosFromSnapshot(ctx, snapshot, { month, setorId, colaboradorId });
    return { rows };
  });

  ipcMain.handle('rh:he:upsert', (_, payload) => {
    return upsertRhHoraExtra({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim(),
      colaborador_id: String((payload && payload.colaborador_id) || '').trim(),
      data: String((payload && payload.data) || '').trim(),
      quantidade_horas: Number((payload && payload.quantidade_horas) || 0),
      tipo_hora: String((payload && payload.tipo_hora) || '').trim(),
      observacao: String((payload && payload.observacao) || '').trim(),
      justificativa: String((payload && payload.justificativa) || '').trim()
    });
  });

  ipcMain.handle('rh:atrasos:upsert', (_, payload) => {
    return upsertRhAtraso({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim(),
      colaborador_id: String((payload && payload.colaborador_id) || '').trim(),
      data: String((payload && payload.data) || '').trim(),
      quantidade_horas: Number((payload && payload.quantidade_horas) || 0),
      observacao: String((payload && payload.observacao) || '').trim()
    });
  });

  ipcMain.handle('rh:solicitacoes:he:upsert', async (_, payload) => {
    const result = upsertRhSolicitacaoHe({
      ...(payload || {}),
      username: String((payload && payload.username) || '').trim()
    });

    try {
      const pdfFilePath = await generateRhSolicitacaoHePdf(result.row);
      return { ...result, pdfFilePath };
    } catch (error) {
      // Mantém o registro salvo, mas devolve o erro para o usuário poder corrigir (permissão/pasta) e gerar novamente.
      throw new Error(`Solicitação salva, mas não consegui gerar o PDF: ${error.message || error}`);
    }
  });

  ipcMain.handle('rh:solicitacoes:he:list', (_, payload) => {
    return listRhSolicitacoesHe({
      username: String((payload && payload.username) || '').trim(),
      filters: payload && payload.filters ? payload.filters : (payload || {})
    });
  });

  ipcMain.handle('rh:solicitacoes:he:delete', (_, payload) => {
    return deleteRhSolicitacaoHe({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:solicitacoes:he:pdf:regen', async (_, payload) => {
    return regenRhSolicitacaoHePdf({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:solicitacoes:he:pdf:open', async (_, payload) => {
    return openRhSolicitacaoHePdf({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:he:delete', (_, payload) => {
    return deleteRhHoraExtra({
      username: String((payload && payload.username) || '').trim(),
      id: String((payload && payload.id) || '').trim()
    });
  });

  ipcMain.handle('rh:kpis', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    const month = String((payload && payload.month) || '').trim();
    const setorId = String((payload && payload.setorId) || '').trim();
    const colaboradorId = String((payload && payload.colaboradorId) || '').trim();
    const tipoHora = String((payload && payload.tipoHora) || '').trim();
    return { kpis: computeRhKpisFromSnapshot(ctx, snapshot, { month, setorId, colaboradorId, tipoHora }) };
  });

  ipcMain.handle('rh:atrasos:kpis', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    const month = String((payload && payload.month) || '').trim();
    const setorId = String((payload && payload.setorId) || '').trim();
    const colaboradorId = String((payload && payload.colaboradorId) || '').trim();
    return { kpis: computeRhAtrasosKpisFromSnapshot(ctx, snapshot, { month, setorId, colaboradorId }) };
  });

  ipcMain.handle('rh:users:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { workbook } = readRhDb();
    const users = listRhUsers(workbook);
    const visible = ctx.canEdit ? users : users.filter((u) => normalizeUsername(u.username) === normalizeUsername(ctx.username));
    return { users: visible };
  });

  ipcMain.handle('rh:users:upsert', (_, payload) => {
    return upsertRhUserRule({
      username: String((payload && payload.username) || '').trim(),
      target: String((payload && payload.target) || '').trim(),
      role: String((payload && payload.role) || '').trim(),
      setor_id: String((payload && payload.setor_id) || '').trim()
    });
  });

  ipcMain.handle('rh:users:delete', (_, payload) => {
    return deleteRhUserRule({
      username: String((payload && payload.username) || '').trim(),
      target: String((payload && payload.target) || '').trim()
    });
  });

  ipcMain.handle('rh:config:get', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    return { config: snapshot && snapshot.config ? snapshot.config : {} };
  });

  ipcMain.handle('rh:config:set', (_, payload) => {
    return setRhConfig({
      username: String((payload && payload.username) || '').trim(),
      config: payload && payload.config ? payload.config : {}
    });
  });

  ipcMain.handle('rh:audit:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const limit = Math.max(1, Math.min(500, Number((payload && payload.limit) || 200)));
    const { workbook } = readRhDb();
    const objs = readSheetObjects(workbook, 'AUDIT_LOG');
    const rows = objs
      .map((r) => ({
        id: String(r.ID || '').trim(),
        data_hora: String(r.DATA_HORA || '').trim(),
        username: String(r.USERNAME || '').trim(),
        action: String(r.ACTION || '').trim(),
        entity: String(r.ENTITY || '').trim(),
        entity_id: String(r.ENTITY_ID || '').trim()
      }))
      .filter((r) => r.data_hora)
      .sort((a, b) => String(b.data_hora).localeCompare(String(a.data_hora)));

    const visible = ctx.canEdit ? rows : rows.slice(0, limit);
    return { rows: visible.slice(0, limit) };
  });

  ipcMain.handle('rh:export:excel', async (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false });
    const { snapshot } = readRhSnapshot();
    const config = snapshot && snapshot.config ? snapshot.config : {};
    const month = String((payload && payload.month) || '').trim();
    const setorId = String((payload && payload.setorId) || '').trim();
    const colaboradorId = String((payload && payload.colaboradorId) || '').trim();
    const tipoHora = String((payload && payload.tipoHora) || '').trim();
    const suggestedName = String((payload && payload.suggestedName) || '').trim() || `rh-he-${month}.xlsx`;

    const rows = listRhHorasExtrasFromSnapshot(ctx, snapshot, { month, setorId, colaboradorId, tipoHora });
    const colaboradores = listRhColaboradoresFromSnapshot(ctx, snapshot);
    const setores = listRhSetoresFromSnapshot(ctx, snapshot);
    const collabById = new Map(colaboradores.map((c) => [String(c.id), c]));
    const setorById = new Map(setores.map((s) => [String(s.id), s]));
    const totalsByColabId = new Map();
    rows.forEach((r) => {
      const key = String(r.colaborador_id || '').trim();
      if (!key) {
        return;
      }
      totalsByColabId.set(key, (totalsByColabId.get(key) || 0) + (Number(r.quantidade_horas) || 0));
    });

    const exportAoa = [[
      'DATA',
      'COLABORADOR',
      'SETOR',
      'HORAS',
      'MARGEM_FAIXA',
      'MARGEM_PERCENT',
      'TIPO',
      'OBSERVACAO',
      'JUSTIFICATIVA',
      'CRIADO_POR',
      'DATA_REGISTRO'
    ]];
    rows.forEach((r) => {
      const collab = collabById.get(String(r.colaborador_id));
      const setorIdForCollab = collab ? String(collab.setor_id || '').trim() : '';
      const setor = setorIdForCollab && setorById.get(setorIdForCollab) ? String(setorById.get(setorIdForCollab).nome_setor || '').trim() : '';
      const monthTotal = totalsByColabId.get(String(r.colaborador_id || '').trim()) || 0;
      const margin = getRhMarginForHours(config, monthTotal);
      exportAoa.push([
        String(r.data || '').trim(),
        collab ? String(collab.nome || '').trim() : '',
        setor,
        Number(r.quantidade_horas || 0),
        String(margin.faixa || '').trim(),
        Number(margin.percent || 0),
        String(r.tipo_hora || '').trim(),
        String(r.observacao || '').trim(),
        String(r.justificativa || '').trim(),
        String(r.criado_por || '').trim(),
        String(r.data_registro || '').trim()
      ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(exportAoa);
    XLSX.utils.book_append_sheet(wb, ws, 'HORAS_EXTRAS');

    const defaultPath = path.join(app.getPath('downloads'), suggestedName);
    const saveResult = await dialog.showSaveDialog({
      title: 'Salvar exportação RH (Excel)',
      defaultPath,
      filters: [{ name: 'Planilha Excel', extensions: ['xlsx'] }]
    });
    if (saveResult.canceled || !saveResult.filePath) {
      return { canceled: true };
    }
    writeWorkbookFile(wb, saveResult.filePath);
    return { canceled: false, filePath: saveResult.filePath };
  });

  ipcMain.handle('rh:import:interno', (_, payload) => {
    return importRhPontoInterno({
      username: String((payload && payload.username) || '').trim(),
      date: String((payload && payload.date) || '').trim(),
      workbookPath: String((payload && payload.workbookPath) || '').trim()
    });
  });

  ipcMain.handle('rh:import:pickWorkbook', async () => {
    const result = await dialog.showOpenDialog({
      title: 'Selecionar planilha base (RH)',
      properties: ['openFile'],
      filters: [{ name: 'Planilha Excel', extensions: ['xlsx', 'xlsm', 'xls'] }]
    });

    if (result.canceled || !result.filePaths || !result.filePaths[0]) {
      return null;
    }

    return result.filePaths[0];
  });

  ipcMain.handle('rh:export:pdf', async (event, payload) => {
    // Reaproveita o exportador de PDF já existente (printToPDF).
    return exportPcpDashboardPdf(event.sender, {
      ...(payload || {}),
      dialogTitle: 'Salvar exportação RH em PDF'
    });
  });

  ipcMain.handle('rh:access:request', (_, payload) => {
    return createRhAccessRequest({
      username: String((payload && payload.username) || '').trim(),
      requested_role: String((payload && payload.requested_role) || '').trim(),
      requested_setor_id: String((payload && payload.requested_setor_id) || '').trim()
    });
  });

  ipcMain.handle('rh:access:requests:list', (_, payload) => {
    const username = String((payload && payload.username) || '').trim();
    const ctx = ensureRhAccess(username, { requireEdit: false, allowUnconfigured: true });
    const { workbook } = readRhDb();
    const all = listRhAccessRequests(workbook)
      .sort((a, b) => String(b.requested_at).localeCompare(String(a.requested_at)));

    const visible = ctx.canEdit ? all : all.filter((r) => r.username === ctx.username);
    return { rows: visible.slice(0, 200) };
  });

  ipcMain.handle('rh:access:requests:decide', (_, payload) => {
    return decideRhAccessRequest({
      username: String((payload && payload.username) || '').trim(),
      requestId: String((payload && payload.requestId) || '').trim(),
      decision: String((payload && payload.decision) || '').trim()
    });
  });

  ipcMain.handle('fiscal:list-nfs', () => {
    const list = listNfsFromBancoPedidos();
    return { nfs: list };
  });

  ipcMain.handle('fiscal:get-nf', (_, payload) => {
    const nf = String((payload && payload.nf) || '').trim();
    const items = getNfItemsFromBancoPedidos(nf);
    return { items };
  });

  ipcMain.handle('fiscal:preview-items', (_, payload) => {
    const identifiers = payload && payload.identifiers ? payload.identifiers : [];
    const items = loadFiscalItemsFromBancoOrdens(identifiers);
    return { items };
  });

  ipcMain.handle('fiscal:register-nf', (_, payload) => {
    const nf = String((payload && payload.nf) || '').trim();
    const client = String((payload && payload.client) || '').trim();
    const user = String((payload && payload.user) || '').trim();
    const faturadaAt = String((payload && payload.faturadaAt) || '').trim();
    const identifiers = payload && payload.identifiers ? payload.identifiers : [];

    if (!nf || !client || !user) {
      throw new Error('NF, Cliente e Usuario sao obrigatorios.');
    }

    const faturadaDate = new Date(faturadaAt);
    if (!faturadaAt || Number.isNaN(faturadaDate.getTime())) {
      throw new Error('Data/Horario faturado invalido.');
    }

    const items = loadFiscalItemsFromBancoOrdens(identifiers);
    if (!items.length) {
      throw new Error('Nenhum item encontrado para gravar no Banco de Pedidos.');
    }

    validateFiscalCadastro({
      nf,
      itemsToInsert: items
    });

    const dataFaturada = faturadaDate.toLocaleDateString('pt-BR');
    const horario = faturadaDate.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    const rowsToAppend = items.map((item) => ([
      item.qtd ?? '',
      item.produto ?? '',
      item.pedido ?? '',
      client,
      'Faturado',
      item.dataEntrada ?? '',
      '',
      nf,
      '',
      dataFaturada,
      user,
      horario
    ]));

    appendRowsToBancoPedidos(rowsToAppend);

    const inserted = items.map((item) => ({
      qtd: item.qtd ?? '',
      produto: item.produto ?? '',
      pedido: item.pedido ?? '',
      cliente: client,
      nf,
      status: 'Faturado',
      dataFaturada,
      horario
    }));

    return { inserted };
  });

  ipcMain.handle('fiscal:update-nf', (_, payload) => {
    const nf = String((payload && payload.nf) || '').trim();
    const status = payload && payload.status != null ? String(payload.status).trim() : null;
    const rastreio = payload && payload.rastreio != null ? String(payload.rastreio).trim() : null;
    const dataDespache = payload && payload.dataDespache != null ? String(payload.dataDespache).trim() : null;
    const editorUser = String((payload && payload.editorUser) || '').trim();
    const editorAt = (payload && payload.editorAt) || new Date();

    const updated = updateNfInBancoPedidos({
      nf,
      status,
      rastreio,
      dataDespache,
      editorUser,
      editorAt
    });

    return { updated };
  });

  ipcMain.handle('fiscal:delete-nf', (_, payload) => {
    const nf = String((payload && payload.nf) || '').trim();
    const reason = String((payload && payload.reason) || '').trim();
    const editorUser = String((payload && payload.editorUser) || '').trim();
    const editorAt = (payload && payload.editorAt) || new Date();

    const deletedRows = deleteNfFromBancoPedidos({
      nf,
      reason,
      editorUser,
      editorAt
    });

    return { deletedRows };
  });

  ipcMain.handle('fiscal:history', () => {
    const rows = listDeletedNfHistory();
    return { rows };
  });

  ipcMain.handle('fiscal:find-nf', (_, payload) => {
    const identifier = String((payload && payload.identifier) || '').trim();
    const nfs = findNfsByPedidoOrReq(identifier);
    return { nfs };
  });

  ipcMain.handle('fiscal:tracking:lookup', async (_, payload) => {
    const code = normalizeTrackingCode(payload && payload.code);
    if (!code || !looksLikeCorreiosTrackingCode(code)) {
      throw new Error('Codigo de rastreio invalido. Ex: AA123456789BR');
    }

    const settings = loadSettings();
    const apiKey = String(settings.trackingApiKey || '').trim().replace(/^Bearer\s+/i, '');
    if (!apiKey) {
      throw new Error('Configure a chave da API de rastreio na aba Configuracao (SeuRastreio).');
    }

    const tracking = await trackSeurastreio(code, apiKey);
    return {
      ...tracking,
      fiscal: findFiscalInfoByTrackingCode(code)
    };
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
