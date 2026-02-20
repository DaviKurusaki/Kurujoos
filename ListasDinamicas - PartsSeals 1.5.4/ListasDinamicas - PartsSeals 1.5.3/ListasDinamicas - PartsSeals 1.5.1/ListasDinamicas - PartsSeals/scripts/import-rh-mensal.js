/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DEFAULT_DB_PATH = 'R:\\Arquivos KuruJoos\\RH\\GestaoHorasExtras.xlsx';

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
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function parseHmsToHours(hms) {
  const text = String(hms || '').trim();
  const match = text.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!match) return 0;
  const hh = Number(match[1]);
  const mm = Number(match[2]);
  const ss = Number(match[3]);
  if (![hh, mm, ss].every((n) => Number.isFinite(n) && n >= 0)) return 0;
  return hh + mm / 60 + ss / 3600;
}

function getLastDayOfMonthYmd(month) {
  const safe = String(month || '').trim();
  const match = safe.match(/^(\d{4})-(\d{2})$/);
  if (!match) return '';
  const yyyy = Number(match[1]);
  const mm = Number(match[2]);
  if (!Number.isFinite(yyyy) || !Number.isFinite(mm) || mm < 1 || mm > 12) return '';
  const dt = new Date(Date.UTC(yyyy, mm, 0)); // day 0 of next month
  const dd = String(dt.getUTCDate()).padStart(2, '0');
  return `${match[1]}-${match[2]}-${dd}`;
}

function main() {
  const month = process.argv[2] || '2026-01';
  const dbPath = process.argv[3] || DEFAULT_DB_PATH;
  const ymd = getLastDayOfMonthYmd(month);
  if (!ymd) {
    console.error('Uso: node scripts/import-rh-mensal.js YYYY-MM [caminho.xlsx]');
    process.exit(1);
  }

  if (!fs.existsSync(dbPath)) {
    console.error(`Arquivo não encontrado: ${dbPath}`);
    process.exit(1);
  }

  const wb = XLSX.readFile(dbPath, { cellDates: true });
  const collabAoa = readSheetAoa(wb, 'COLABORADORES');
  const collabObjs = mapAoaToObjects(collabAoa).filter((r) => !String(r.DELETED_AT || '').trim());
  const collabByName = new Map();
  collabObjs.forEach((c) => {
    const key = normalizeKey(c.NOME);
    if (key) collabByName.set(key, c);
  });

  const heAoa = readSheetAoa(wb, 'HORAS_EXTRAS');
  const heHeader = getHeaderFromAoa(heAoa, [
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
  const heRows = mapAoaToObjects([heHeader, ...(heAoa.slice(1) || [])]);

  const tag = `IMPORT_MENSAL:${month}`;
  const now = new Date().toISOString();

  const items = [
    { nome: 'André', hms: '05:45:00' },
    { nome: 'Andressa Larissa', hms: '03:32:00' },
    { nome: 'Darci Lopes', hms: '10:51:00' },
    { nome: 'Davi Augusto', hms: '22:14:00' },
    { nome: 'Diogo Pereira', hms: '06:18:00' },
    { nome: 'Edson Ricardo', hms: '02:57:00' },
    { nome: 'Gustavo Scaffi', hms: '20:32:00' },
    { nome: 'João Guilherme', hms: '13:09:00' },
    { nome: 'João Pedro', hms: '00:00:00' },
    { nome: 'José Augusto', hms: '20:52:00' },
    { nome: 'Kevin Rogério', hms: '14:47:00' },
    { nome: 'Leonardo Martins', hms: '06:20:00' },
    { nome: 'Mayara Tofaneli', hms: '01:50:00' },
    { nome: 'Murilo Fantacussi', hms: '00:00:00' },
    { nome: 'Murilo Lucato', hms: '12:16:00' },
    { nome: 'Natalia Rebeca', hms: '09:59:00' },
    { nome: 'Paloma Lima', hms: '11:16:00' },
    { nome: 'Renan Moraes', hms: '09:06:00' },
    { nome: 'Suélen Santos', hms: '00:00:00' },
    { nome: 'Tiago Andrade', hms: '00:40:00' }
  ];

  let created = 0;
  let updated = 0;
  const missing = [];
  const skippedZero = [];

  items.forEach((item) => {
    const hours = parseHmsToHours(item.hms);
    if (!hours) {
      skippedZero.push(item.nome);
      return;
    }
    const key = normalizeKey(item.nome);
    const collab = collabByName.get(key) || null;
    if (!collab) {
      missing.push(item.nome);
      return;
    }
    const colaboradorId = String(collab.ID || '').trim();
    if (!colaboradorId) {
      missing.push(item.nome);
      return;
    }

    const existing = heRows.find((r) =>
      !String(r.DELETED_AT || '').trim()
      && String(r.COLABORADOR_ID || '').trim() === colaboradorId
      && String(r.DATA || '').trim() === ymd
      && String(r.OBSERVACAO || '').trim() === tag
    ) || null;

    if (existing) {
      existing.QUANTIDADE_HORAS = hours;
      existing.TIPO_HORA = 'outro';
      existing.JUSTIFICATIVA = `Importação consolidada (${month}).`;
      existing.UPDATED_AT = now;
      updated += 1;
      return;
    }

    heRows.push({
      ID: String(generateId()),
      COLABORADOR_ID: colaboradorId,
      DATA: ymd,
      QUANTIDADE_HORAS: hours,
      TIPO_HORA: 'outro',
      OBSERVACAO: tag,
      JUSTIFICATIVA: `Importação consolidada (${month}).`,
      CRIADO_POR: 'IMPORT',
      DATA_REGISTRO: now,
      UPDATED_AT: now,
      DELETED_AT: ''
    });
    created += 1;
  });

  writeSheetAoa(wb, 'HORAS_EXTRAS', objectsToAoa(heHeader, heRows));
  XLSX.writeFile(wb, dbPath);

  console.log(`Mês: ${month} | Data usada: ${ymd}`);
  console.log(`Criados: ${created} | Atualizados: ${updated}`);
  console.log(`Zeros ignorados: ${skippedZero.length}`);
  if (missing.length) {
    console.log(`Não encontrados em COLABORADORES: ${missing.length}`);
    missing.forEach((n) => console.log(`- ${n}`));
  }
  console.log('OK.');
}

main();

