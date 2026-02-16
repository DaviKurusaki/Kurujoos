const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const SETTINGS_FILE = () => path.join(app.getPath('userData'), 'settings.json');
const DEFAULT_SETTINGS = {
  rootFolder: '',
  fiscalRoot: '',
  theme: 'light',
  lastLoginUsername: ''
};
const ADMIN_SETTINGS_PASSWORD = '0604';

function readUsersFile() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const workbook = XLSX.readFile(bancoPedidosPath, { cellDates: true });
  const sheetName = 'USERS';
  if (!workbook.Sheets[sheetName]) {
    const ws = XLSX.utils.aoa_to_sheet([
      ['USERNAME', 'SALT', 'ITERATIONS', 'HASH', 'CAN_FISCAL', 'CAN_PCP', 'CREATED_AT', 'PERMISSIONS']
    ]);
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = ws;
  }

  const ws = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
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

  XLSX.writeFile(workbook, bancoPedidosPath);
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

const FISCAL_BANK_PEDIDOS_RELATIVE = ['Arquivos KuruJoos', 'Financeiro', 'banco Pedidos.xlsx'];
const FISCAL_BANCO_ORDENS_RELATIVE = ['PartsSeals', 'Banco de Ordens', 'BancoDeOrdens.xlsm'];
const PCP_DASHBOARD_DAILY_RELATIVE = ['Rede', 'PartsSeals', 'Programação de Produção Diaria'];

function getFiscalRoot() {
  const settings = loadSettings();
  const fiscalRoot = String(settings.fiscalRoot || '').trim();
  if (!fiscalRoot) {
    throw new Error('Configure o Caminho Fiscal (R:) nas Configuracoes.');
  }
  return fiscalRoot;
}

function resolveFiscalPath(relativeParts) {
  const fiscalRoot = getFiscalRoot();
  return path.join(fiscalRoot, ...relativeParts);
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
    path.join(fiscalRoot, 'PartsSeals', 'Programação de Produção Diaria'),
    path.join(fiscalRoot, 'Programação de Produção Diaria')
  ];
  if (normalizedRoot === 'REDE') {
    candidatePaths.push(path.join(fiscalRoot, 'PartsSeals', 'Programação de Produção Diaria'));
  }
  if (normalizedRoot === 'PARTSSEALS') {
    candidatePaths.push(path.join(fiscalRoot, 'Programação de Produção Diaria'));
  }

  for (const candidate of candidatePaths) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }

  const relativeOptions = [
    PCP_DASHBOARD_DAILY_RELATIVE,
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
  const dt = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(dt.getTime())) {
    throw new Error('Data invalida para Dashboard PCP.');
  }
  const dd = String(dt.getDate()).padStart(2, '0');
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  return `${dd}-${mm}`;
}

function findDailyDashboardWorkbookPath(dateValue) {
  const ddMm = parseDateToDdMm(dateValue);
  const basePath = resolvePcpDashboardDailyRootPath();
  const files = fs.readdirSync(basePath, { withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((name) => /\.xlsm$/i.test(name));

  const target = normalizeText(ddMm);
  const candidates = files.filter((name) => normalizeText(name).includes(target));
  if (!candidates.length) {
    throw new Error(`Nao encontrei planilha de usinagem para ${ddMm} em ${basePath}.`);
  }

  const sorted = candidates.sort((a, b) => a.localeCompare(b, 'pt-BR'));
  return path.join(basePath, sorted[0]);
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
  const sheetName = workbook.SheetNames.find((name) => normalizeText(name) === 'MOLDAGEM');
  if (!sheetName) {
    return { presses: [], operators: [] };
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return { presses: [], operators: [] };
  }

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });

  const presses = [];
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
    presses: presses.sort((a, b) => a.press - b.press),
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

    const workbook = XLSX.readFile(bancoOrdensPath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
      return 0;
    }

    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
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
  const plannedPiecesByRule = sumPlannedPiecesFromTable7(workbook, mainRowsData.sheetName);
  if (plannedPiecesByRule.total > 0) {
    metrics.kpis.pecasPlanejadas = plannedPiecesByRule.total;
  }
  const materials = parseDashboardMaterials(workbook);
  const opGeradas = countOpsGeneratedByDate(safeDate);
  metrics.kpis.opGeradas = opGeradas;

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
      acabamentoByOperator: metrics.acabamentoByOperator,
      materials
    },
    debug: {
      plannedPiecesByRule
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
    title: 'Salvar dashboard em PDF',
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

  const workbook = XLSX.readFile(bancoOrdensPath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Planilha BancoDeOrdens sem abas validas.');
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
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

  const workbook = XLSX.readFile(bancoPedidosPath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Banco Pedidos sem abas validas.');
  }

  const existing = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
  let lastUsedRow = 0;
  existing.forEach((row, index) => {
    const hasValue = Array.isArray(row) && row.some((cell) => String(cell || '').trim() !== '');
    if (hasValue) {
      lastUsedRow = index + 1;
    }
  });

  const startRow = lastUsedRow + 1;
  XLSX.utils.sheet_add_aoa(worksheet, rowsToAppend, { origin: `A${startRow}` });
  XLSX.writeFile(workbook, bancoPedidosPath);
}

function readBancoPedidosRows() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const workbook = XLSX.readFile(bancoPedidosPath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Banco Pedidos sem abas validas.');
  }

  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
  const headerRow = aoa[0] || [];
  const looksLikeHeader = normalizeText(headerRow[0]) === 'QNTD' && normalizeText(headerRow[3]) === 'CLIENTE';
  const startIndex = looksLikeHeader ? 1 : 0;

  const rows = [];
  for (let i = startIndex; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
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

  return { workbook, sheetName, worksheet, rows, bancoPedidosPath, looksLikeHeader };
}

function listNfsFromBancoPedidos() {
  const { rows } = readBancoPedidosRows();
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

  const list = Array.from(map.values()).map((entry) => {
    const pedidos = Array.from(entry.pedidos || []);
    pedidos.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
    const pedidosLabel = pedidos.length > 10
      ? `${pedidos.slice(0, 10).join(' | ')} | +${pedidos.length - 10}`
      : pedidos.join(' | ');

    let reqsLabel = '';
    try {
      const ordensItems = loadFiscalItemsFromBancoOrdens(pedidos);
      const reqs = Array.from(new Set(ordensItems.map((item) => String(item.req || '').trim()).filter(Boolean)));
      reqs.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
      reqsLabel = reqs.length > 10
        ? `${reqs.slice(0, 10).join(' | ')} | +${reqs.length - 10}`
        : reqs.join(' | ');
    } catch (error) {
      reqsLabel = '';
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

  const { rows } = readBancoPedidosRows();
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
  XLSX.writeFile(workbook, bancoPedidosPath);
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
  const { rows } = readBancoPedidosRows();
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
  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });

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
  XLSX.writeFile(workbook, bancoPedidosPath);

  return deleted.length;
}

function listDeletedNfHistory() {
  const bancoPedidosPath = resolveFiscalPath(FISCAL_BANK_PEDIDOS_RELATIVE);
  if (!fs.existsSync(bancoPedidosPath)) {
    throw new Error(`Banco Pedidos nao encontrado: ${bancoPedidosPath}`);
  }

  const workbook = XLSX.readFile(bancoPedidosPath, { cellDates: true });
  const ws = workbook.Sheets['FISCAL_HISTORICO'];
  if (!ws) {
    return [];
  }

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
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

  const { rows } = readBancoPedidosRows();
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
    const bancoOrdensPath = resolveFiscalPath(FISCAL_BANCO_ORDENS_RELATIVE);
    if (!fs.existsSync(bancoOrdensPath)) {
      return [];
    }

    const workbook = XLSX.readFile(bancoOrdensPath, { cellDates: true });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!worksheet) {
      return [];
    }

    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
    const pedidoSet = new Set();
    aoa.forEach((row) => {
      const pedido = String(row[3] || '').trim();
      const req = String(row[4] || '').trim();
      const base = extractReqBase(req);
      if (pedido && base === reqBase) {
        pedidoSet.add(pedido);
      }
    });
    pedidos = Array.from(pedidoSet);
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
