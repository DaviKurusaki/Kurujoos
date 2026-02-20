/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DEFAULT_DB_PATH = 'R:\\Arquivos KuruJoos\\RH\\GestaoHorasExtras.xlsx';

function ensureDirSync(folderPath) {
  if (!folderPath) return;
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath, { recursive: true });
  }
}

function generateId() {
  return Date.now() * 1000 + Math.floor(Math.random() * 1000);
}

function readSheetAoa(workbook, sheetName) {
  const ws = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
}

function writeSheetAoa(workbook, sheetName, aoa) {
  const next = Array.isArray(aoa) ? aoa : [];
  if (!workbook.Sheets[sheetName]) {
    workbook.SheetNames.push(sheetName);
  }
  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(next);
}

function getHeaderFromAoa(aoa, defaultHeader) {
  const headerRow = Array.isArray(aoa && aoa[0]) ? aoa[0] : null;
  if (headerRow && headerRow.length) {
    return headerRow.map((h) => String(h || '').trim().toUpperCase());
  }
  return (defaultHeader || []).map((h) => String(h || '').trim().toUpperCase());
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
      if (!key) return;
      obj[key] = row[idx] !== undefined ? row[idx] : '';
    });
    result.push(obj);
  }
  return result;
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

function normalizeKey(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}

function createWorkbookIfMissing(filePath) {
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
  XLSX.writeFile(wb, filePath);
  return wb;
}

function ensureSheetsAndColumns(workbook) {
  const required = {
    SETORES: ['ID', 'NOME_SETOR', 'GESTOR_RESPONSAVEL', 'CENTRO_CUSTO', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT'],
    COLABORADORES: ['ID', 'NOME', 'MATRICULA', 'SETOR_ID', 'CARGO', 'STATUS', 'DATA_ADMISSAO', 'LIMITE_MENSAL', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']
  };
  Object.entries(required).forEach(([sheetName, headers]) => {
    if (!workbook.Sheets[sheetName]) {
      writeSheetAoa(workbook, sheetName, [headers]);
      return;
    }
    const aoa = readSheetAoa(workbook, sheetName);
    const existingHeader = Array.isArray(aoa[0]) ? aoa[0].map((h) => String(h || '').trim().toUpperCase()) : [];
    const missing = headers.map((h) => String(h || '').trim().toUpperCase()).filter((h) => h && !existingHeader.includes(h));
    if (!missing.length) return;
    const nextHeader = [...existingHeader, ...missing];
    const next = [nextHeader];
    for (let i = 1; i < aoa.length; i += 1) {
      const row = Array.isArray(aoa[i]) ? [...aoa[i]] : [];
      while (row.length < nextHeader.length) row.push('');
      next.push(row);
    }
    writeSheetAoa(workbook, sheetName, next);
  });
}

function upsertSetores(workbook, setores) {
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'SETORES');
  const header = getHeaderFromAoa(aoa, ['ID', 'NOME_SETOR', 'GESTOR_RESPONSAVEL', 'CENTRO_CUSTO', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const byName = new Map();
  rows.forEach((r) => {
    if (String(r.DELETED_AT || '').trim()) return;
    const name = normalizeKey(r.NOME_SETOR);
    if (name) byName.set(name, r);
  });

  const created = [];
  setores.forEach((nomeSetor) => {
    const key = normalizeKey(nomeSetor);
    if (!key) return;
    if (byName.has(key)) return;
    const row = {
      ID: String(generateId()),
      NOME_SETOR: String(nomeSetor).trim(),
      GESTOR_RESPONSAVEL: '',
      CENTRO_CUSTO: '',
      CREATED_AT: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    };
    rows.push(row);
    byName.set(key, row);
    created.push(nomeSetor);
  });

  writeSheetAoa(workbook, 'SETORES', objectsToAoa(header, rows));
  return { byName, created };
}

function upsertColaboradores(workbook, collabs, setorByName) {
  const now = new Date().toISOString();
  const aoa = readSheetAoa(workbook, 'COLABORADORES');
  const header = getHeaderFromAoa(aoa, ['ID', 'NOME', 'MATRICULA', 'SETOR_ID', 'CARGO', 'STATUS', 'DATA_ADMISSAO', 'LIMITE_MENSAL', 'CREATED_AT', 'UPDATED_AT', 'DELETED_AT']);
  const rows = mapAoaToObjects([header, ...(aoa.slice(1) || [])]);

  const byName = new Map();
  rows.forEach((r) => {
    if (String(r.DELETED_AT || '').trim()) return;
    const key = normalizeKey(r.NOME);
    if (key) byName.set(key, r);
  });

  const created = [];
  const updated = [];
  const skipped = [];

  collabs.forEach(({ setor, nome }) => {
    const nomeKey = normalizeKey(nome);
    const setorKey = normalizeKey(setor === 'ADM' ? 'Administrativo' : setor);
    const setorRow = setorByName.get(setorKey);
    if (!setorRow) {
      skipped.push({ nome, reason: `Setor não encontrado: ${setor}` });
      return;
    }
    const setorId = String(setorRow.ID || '').trim();
    if (!setorId) {
      skipped.push({ nome, reason: `Setor sem ID: ${setor}` });
      return;
    }

    const existing = byName.get(nomeKey) || null;
    if (existing) {
      const previousSetorId = String(existing.SETOR_ID || '').trim();
      existing.NOME = String(nome).trim();
      existing.SETOR_ID = setorId;
      existing.STATUS = 'ativo';
      existing.UPDATED_AT = now;
      if (String(existing.DELETED_AT || '').trim()) existing.DELETED_AT = '';
      if (previousSetorId !== setorId) {
        updated.push({ nome, from: previousSetorId, to: setorId });
      }
      return;
    }

    rows.push({
      ID: String(generateId()),
      NOME: String(nome).trim(),
      MATRICULA: '',
      SETOR_ID: setorId,
      CARGO: '',
      STATUS: 'ativo',
      DATA_ADMISSAO: '',
      LIMITE_MENSAL: '',
      CREATED_AT: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    });
    created.push(nome);
  });

  writeSheetAoa(workbook, 'COLABORADORES', objectsToAoa(header, rows));
  return { created, updated, skipped };
}

function main() {
  const dbPath = process.argv[2] || DEFAULT_DB_PATH;
  console.log(`RH DB: ${dbPath}`);

  const workbook = fs.existsSync(dbPath)
    ? XLSX.readFile(dbPath, { cellDates: true })
    : createWorkbookIfMissing(dbPath);

  ensureSheetsAndColumns(workbook);

  const setores = [
    'Administrativo',
    'Usinagem',
    'Moldagem',
    'Projeto',
    'Acabamento',
    'Logistica',
    'PCP'
  ];

  const colaboradores = [
    { setor: 'Usinagem', nome: 'André' },
    { setor: 'Moldagem', nome: 'Andressa Larissa' },
    { setor: 'Acabamento', nome: 'Darci Lopes' },
    { setor: 'PCP', nome: 'Davi Augusto' },
    { setor: 'Usinagem', nome: 'Diogo Pereira' },
    { setor: 'Usinagem', nome: 'Edson Ricardo' },
    { setor: 'Acabamento', nome: 'Gustavo Scaffi' },
    { setor: 'Usinagem', nome: 'João Guilherme' },
    { setor: 'Usinagem', nome: 'João Pedro' },
    { setor: 'PCP', nome: 'José Augusto' },
    { setor: 'Usinagem', nome: 'Kevin Rogério' },
    { setor: 'Usinagem', nome: 'Leonardo Martins' },
    { setor: 'ADM', nome: 'Mayara Tofaneli' },
    { setor: 'Usinagem', nome: 'Murilo Fantacussi' },
    { setor: 'Usinagem', nome: 'Murilo Lucato' },
    { setor: 'Acabamento', nome: 'Natalia Rebeca' },
    { setor: 'Acabamento', nome: 'Paloma Lima' },
    { setor: 'Usinagem', nome: 'Renan Moraes' },
    { setor: 'Moldagem', nome: 'Suélen Santos' },
    { setor: 'Usinagem', nome: 'Tiago Andrade' }
  ];

  const { byName: setorByName, created: setoresCriados } = upsertSetores(workbook, setores);
  const { created, updated, skipped } = upsertColaboradores(workbook, colaboradores, setorByName);

  XLSX.writeFile(workbook, dbPath);

  console.log(`Setores criados: ${setoresCriados.length}`);
  if (setoresCriados.length) console.log(`- ${setoresCriados.join('\n- ')}`);
  console.log(`Colaboradores criados: ${created.length}`);
  console.log(`Colaboradores atualizados (mudança de setor): ${updated.length}`);
  if (skipped.length) {
    console.log(`Colaboradores ignorados: ${skipped.length}`);
    skipped.forEach((s) => console.log(`- ${s.nome}: ${s.reason}`));
  }
  console.log('OK.');
}

main();
