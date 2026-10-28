import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

/* ==== src/constants.js ==== */
const JD = {
  green: '#367C2B',
  greenDark: '#1F4D19',
  greenPale: '#E8F5E0',
  yellow: '#FFDE00',
};

const P = {
  bg: '#F2F1ED',
  card: '#FFFFFF',
  dark: '#111111',
  mid: '#444444',
  muted: '#777777',
  border: '#DDD9D0',
  borderDark: '#BBB7AE',
  A: '#059669', Ab: '#ECFDF5', Ab2: '#D1FAE5', Aborder: '#6EE7B7',
  B: '#DC2626', Bb: '#FEF2F2', Bb2: '#FEE2E2', Bborder: '#FCA5A5',
  C: '#D97706', Cb: '#FFFBEB', Cb2: '#FEF3C7', Cborder: '#FCD34D',
  warn: '#F59E0B', warnBg: '#FFFBEB',
  info: '#2563EB', infoBg: '#EFF6FF',
};

const CC = {
  A: { bg: P.Ab, border: P.Aborder, activeBg: P.Ab2, color: P.A, label: 'Cenário A — Localizado' },
  B: { bg: P.Bb, border: P.Bborder, activeBg: P.Bb2, color: P.B, label: 'Cenário B — Não localizado' },
  C: { bg: P.Cb, border: P.Cborder, activeBg: P.Cb2, color: P.C, label: 'Cenário C — Não cadastrado' },
};

const AUDIT_COLS = [
  'ID_AUDITORIA',
  'LINHA_ORIGEM',
  'TIPO_REGISTRO',
  'RESPONSÁVEL ENVIO',
  'CENÁRIO',
  'OBSERVAÇÃO AUDITORIA',
  'ARQUIVO FOTO',
];

const FINAL_COLS = [
  'STATUS AUDITORIA',
  'CENÁRIO',
  'OBSERVAÇÃO AUDITORIA',
  'ARQUIVO FOTO',
  'RESPONSÁVEL ENVIO',
  'DATA IMPORTAÇÃO',
  'ORIGEM CRUZAMENTO',
];

const FIELD_ALIASES = {
  responsavel: ['RESPONSÁVEL', 'RESPONSAVEL', 'RESPONSÁVEL AUDITORIA', 'RESPONSAVEL AUDITORIA'],
  filial: ['FILIAL', 'UNIDADE', 'LOJA'],
  centroCusto: ['CENTRO DE CUSTO', 'CENTRO CUSTO', 'CC', 'C CUSTO'],
  area: ['ÁREA', 'AREA', 'DEPARTAMENTO', 'SETOR', 'LOCALIZAÇÃO', 'LOCALIZACAO'],
  codigo: ['CÓDIGO', 'CODIGO', 'CODIGO DO BEM', 'CÓDIGO DO BEM', 'COD. BEM', 'COD BEM'],
  descricao: ['DESCRIÇÃO', 'DESCRICAO', 'DESCRIÇÃO DO BEM', 'DESCRICAO DO BEM', 'BEM'],
  plaqueta: ['PLAQUETA', 'Nº PLAQUETA', 'N° PLAQUETA', 'NUM PLAQUETA', 'NÚMERO PLAQUETA', 'NUMERO PLAQUETA', 'PATRIMÔNIO', 'PATRIMONIO'],
};

const FALLBACK_INDEX = {
  responsavel: 0,
  filial: 1,
  centroCusto: 2,
  area: 3,
  codigo: 4,
  descricao: 6,
  plaqueta: 7,
};

/* ==== src/utils/text.js ==== */
const clean = (value) => String(value ?? '').trim();

const normalizeText = (value) => clean(value)
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .toUpperCase();

const safeName = (value, fallback = 'arquivo') => {
  const base = clean(value) || fallback;
  return base
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || fallback;
};

const todayBR = () => new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');

const getExtension = (fileName, fallback = 'jpg') => {
  const ext = clean(fileName).split('.').pop();
  return ext && ext !== fileName ? ext.toLowerCase() : fallback;
};

/* ==== src/utils/headers.js ==== */

const normalizeHeader = normalizeText;

const findHeaderIndex = (headers, names, fallbackIndex = -1) => {
  const normalizedNames = names.map(normalizeHeader);
  const idx = (headers || []).findIndex((h) => normalizedNames.includes(normalizeHeader(h)));
  return idx >= 0 ? idx : fallbackIndex;
};

const fieldIndex = (headers, field) => findHeaderIndex(headers, FIELD_ALIASES[field] || [field], FALLBACK_INDEX[field] ?? -1);

const getByField = (headers, row, field) => {
  const idx = fieldIndex(headers, field);
  return idx >= 0 ? row?.[idx] ?? '' : '';
};

const setByField = (headers, row, field, value) => {
  const idx = fieldIndex(headers, field);
  if (idx >= 0) row[idx] = value;
};

const getByHeader = (headers, row, names, fallbackIndex = -1) => {
  const idx = findHeaderIndex(headers, names, fallbackIndex);
  return idx >= 0 ? row?.[idx] ?? '' : '';
};

const setByHeader = (headers, row, names, value, fallbackIndex = -1) => {
  const idx = findHeaderIndex(headers, names, fallbackIndex);
  if (idx >= 0) row[idx] = value;
};

const rowToMap = (headers, row) => {
  const map = new Map();
  (headers || []).forEach((h, i) => {
    const key = normalizeHeader(h);
    if (key && !map.has(key)) map.set(key, row?.[i] ?? '');
  });
  return map;
};

const valueByNormalizedHeader = (headers, row, headerName) => {
  const key = normalizeHeader(headerName);
  const idx = (headers || []).findIndex((h) => normalizeHeader(h) === key);
  return idx >= 0 ? row?.[idx] ?? '' : '';
};

const hasRealValue = (row) => (row || []).some((cell) => clean(cell) !== '');

/* ==== src/utils/storage.js ==== */

const localStorageKey = (name) => `audit_v4_${safeName(name, 'responsavel')}`;

const importStorageKey = 'audit_v4_import_status';

/* ==== src/utils/excel.js ==== */

  const isDateHeader = (header) => {
    const h = normalizeHeader(header);

    return (
      h.includes('DATA') ||
      h.includes('DT ') ||
      h.includes('AQUISICAO') ||
      h.includes('AQUISIÇÃO')
    );
  };

  const isMoneyHeader = (header) => {
  const h = normalizeHeader(header);

  const camposQueNaoSaoValores = [
      'CENTRO DE CUSTO',
      'CENTRO CUSTO',
      'C CUSTO',
      'CC',
      'CODIGO',
      'CODIGO DO BEM',
      'COD BEM',
      'PLAQUETA',
      'PATRIMONIO',
      'FILIAL',
      'UNIDADE',
      'LOJA',
    ];

    if (camposQueNaoSaoValores.some((campo) => h === campo || h.includes(campo))) {
      return false;
    }

    return (
      h.includes('VALOR') ||
      h.includes('VLR') ||
      h.includes('PRECO') ||
      h.includes('SALDO') ||
      h.includes('TOTAL')
    );
  };

  const excelSerialToDate = (serial) => {
    const num = Number(serial);

    if (!Number.isFinite(num)) return clean(serial);

    const utcDays = Math.floor(num - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);

    const day = String(dateInfo.getUTCDate()).padStart(2, '0');
    const month = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
    const year = dateInfo.getUTCFullYear();

    return `${day}/${month}/${year}`;
  };

  const formatDateBR = (value) => {
    const text = clean(value);

    if (!text) return '';

    // Exemplo: 44345
    if (/^\d+(\.\d+)?$/.test(text)) {
      const num = Number(text);

      if (num > 20000 && num < 80000) {
        return excelSerialToDate(num);
      }

      return text;
    }

    // Exemplo americano: 5/29/2021 ou 5/29/21
    const usDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);

    if (usDate) {
      const first = Number(usDate[1]);
      const second = Number(usDate[2]);
      let year = usDate[3];

      if (year.length === 2) {
        year = `20${year}`;
      }

      // Se o segundo número for maior que 12, é claramente formato americano M/D/Y
      if (second > 12 && first >= 1 && first <= 12) {
        return `${String(second).padStart(2, '0')}/${String(first).padStart(2, '0')}/${year}`;
      }

      // Se o primeiro número for maior que 12, já está no formato brasileiro D/M/Y
      if (first > 12 && second >= 1 && second <= 12) {
        return `${String(first).padStart(2, '0')}/${String(second).padStart(2, '0')}/${year}`;
      }

      // Em caso ambíguo, mantém como brasileiro com dois dígitos
      return `${String(first).padStart(2, '0')}/${String(second).padStart(2, '0')}/${year}`;
    }

    // Exemplo: 2021-05-29
    const isoDate = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

    if (isoDate) {
      const year = isoDate[1];
      const month = isoDate[2];
      const day = isoDate[3];

      return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
    }

    return text;
  };

  const formatMoneyBR = (value) => {
    const text = clean(value);

    if (!text) return '';

    // Evita alterar textos não numéricos
    if (!/^[\d.,\-R$\s]+$/.test(text)) return text;

    let normalized = text
      .replace(/R\$/gi, '')
      .replace(/\$/g, '')
      .replace(/\s/g, '');

    // Caso já venha como 4.500,00
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(normalized)) {
      const num = Number(normalized.replace(/\./g, '').replace(',', '.'));

      if (Number.isFinite(num)) {
        return num.toLocaleString('pt-BR', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      }
    }

    // Caso venha como 4,500.00
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(normalized)) {
      const num = Number(normalized.replace(/,/g, ''));

      if (Number.isFinite(num)) {
        return num.toLocaleString('pt-BR', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      }
    }

    // Caso venha como 4500.00
    if (/^-?\d+(\.\d+)?$/.test(normalized)) {
      const num = Number(normalized);

      if (Number.isFinite(num)) {
        return num.toLocaleString('pt-BR', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      }
    }

    // Caso venha como 4500,00 ou 650,00
    if (/^-?\d+(,\d+)?$/.test(normalized)) {
      const num = Number(normalized.replace(',', '.'));

      if (Number.isFinite(num)) {
        return num.toLocaleString('pt-BR', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      }
    }

    return text;
  };

  const formatCellByHeader = (header, value) => {
    if (isDateHeader(header)) {
      return formatDateBR(value);
    }

    if (isMoneyHeader(header)) {
      return formatMoneyBR(value);
    }

    return value;
  };

  const parseWorkbookArrayBuffer = (arrayBuffer) => {
    const wb = XLSX.read(arrayBuffer, {
      type: 'array',
      cellDates: false,
      cellNF: true,
      cellText: false,
    });

    const ws = wb.Sheets[wb.SheetNames[0]];

    const raw = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: '',
      raw: false,
      dateNF: 'dd/mm/yyyy',
    });

    const headers = (raw[0] || []).map((h) => String(h ?? '').trim());

    const rows = raw
      .slice(1)
      .filter(hasRealValue)
      .map((row) => headers.map((header, index) => (
        formatCellByHeader(header, row[index] ?? '')
      )));

    return {
      wb,
      headers,
      rows,
      sheetName: wb.SheetNames[0],
    };
  };

const makeWorkbookArray = (sheets) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows, widths }) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = widths || (rows[0] || []).map(() => ({ wch: 22 }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
};

const downloadBlob = (blob, fileName) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
};

/* ==== src/utils/auditPackage.js ==== */

const photoFileName = ({ plaqueta, scenario, index, originalName, prefix = 'item' }) => {
  const base = safeName(plaqueta || `${prefix}_${index + 1}`, `${prefix}_${index + 1}`);
  return `${base}_${scenario}.${getExtension(originalName)}`;
};

const base64Payload = (dataUrl) => String(dataUrl || '').split(',')[1] || '';

const validateAudit = ({ myItems, responses, extraItems, headers }) => {
  const pendencias = [];

  myItems.forEach(({ row, idx }) => {
    const r = responses[idx] || {};
    const plaqueta = clean(getByField(headers, row, 'plaqueta')) || `Linha ${idx + 2}`;
    const descricao = clean(getByField(headers, row, 'descricao')) || 'Sem descrição';

    if (!r.cenario) pendencias.push(`${plaqueta} - ${descricao}: cenário não marcado`);
    if (r.cenario === 'A' && !r.foto) pendencias.push(`${plaqueta} - ${descricao}: Cenário A precisa de foto`);
    if (r.cenario === 'B' && !clean(r.obs)) pendencias.push(`${plaqueta} - ${descricao}: Cenário B precisa de observação`);
  });

  extraItems.forEach((item, i) => {
    if (!clean(item.desc)) pendencias.push(`Cenário C - Item #${i + 1}: descrição obrigatória`);
    if (!item.foto) pendencias.push(`Cenário C - Item #${i + 1}: foto obrigatória`);
  });

  return pendencias;
};

const montarLinhaExtra = ({ headers, item, selectedResp, filial }) => {
  const row = new Array(headers.length).fill('');
  setByField(headers, row, 'responsavel', selectedResp);
  setByField(headers, row, 'filial', filial);
  setByField(headers, row, 'descricao', item.desc || '');
  setByField(headers, row, 'plaqueta', item.plaqueta || '');
  return row;
};

const buildAuditRowsAndPhotos = ({ headers, myItems, responses, extraItems, selectedResp }) => {
  const filial = myItems[0]?.row ? clean(getByField(headers, myItems[0].row, 'filial')) : '';
  const auditHeaders = [...headers, ...AUDIT_COLS];
  const photos = [];

  const linhasBase = myItems.map(({ row, idx }, i) => {
    const r = responses[idx] || {};
    const plaqueta = clean(getByField(headers, row, 'plaqueta'));
    let fotoPath = '';

    if (r.foto) {
      const fileName = photoFileName({ plaqueta, scenario: r.cenario, index: i, originalName: r.fotoName, prefix: `linha_${idx + 2}` });
      fotoPath = `Fotos/${fileName}`;
      photos.push({ path: fotoPath, dataUrl: r.foto, fileName });
    }

    return [
      ...headers.map((_, col) => row[col] ?? ''),
      idx,
      idx + 2,
      'BASE',
      selectedResp,
      r.cenario || '',
      r.obs || '',
      fotoPath,
    ];
  });

  const linhasExtras = extraItems.map((item, i) => {
    const plaqueta = clean(item.plaqueta);
    const baseRow = montarLinhaExtra({ headers, item, selectedResp, filial });
    let fotoPath = '';

    if (item.foto) {
      const fileName = photoFileName({ plaqueta, scenario: 'C', index: i, originalName: item.fotoName, prefix: 'novo' });
      fotoPath = `Fotos/${fileName}`;
      photos.push({ path: fotoPath, dataUrl: item.foto, fileName });
    }

    return [
      ...baseRow,
      '',
      '',
      'EXTRA',
      selectedResp,
      'C',
      item.obs || '',
      fotoPath,
    ];
  });

  return {
    rows: [auditHeaders, ...linhasBase, ...linhasExtras],
    photos,
  };
};

const buildAuditZipBlob = async ({ headers, myItems, responses, extraItems, selectedResp }) => {
  const { rows, photos } = buildAuditRowsAndPhotos({ headers, myItems, responses, extraItems, selectedResp });
  const excelArray = makeWorkbookArray([{ name: 'Auditoria', rows }]);
  const zip = new JSZip();
  const safeResp = safeName(selectedResp, 'responsavel');

  zip.file('auditoria.xlsx', excelArray);
  zip.file('manifest.json', JSON.stringify({ responsavel: selectedResp, geradoEm: new Date().toISOString(), versao: '2.0' }, null, 2));

  photos.forEach((photo) => {
    zip.file(photo.path, base64Payload(photo.dataUrl), { base64: true });
  });

  const blob = await zip.generateAsync({ type: 'blob' });
  const fileName = `Auditoria_${safeResp}_${todayBR()}.zip`;
  return { blob, fileName, excelArray, rows, photos };
};

/* ==== src/utils/finalPackage.js ==== */

const normKey = (v) => normalizeText(v);
const baseNameFromPath = (path) => clean(path).split('/').pop() || clean(path).split('\\').pop();

const readAuditZip = async (file) => {
  const zip = await JSZip.loadAsync(file);
  const entries = Object.values(zip.files);
  const excelEntry = entries.find((entry) => !entry.dir && /(^|\/)auditoria\.xlsx$/i.test(entry.name))
    || entries.find((entry) => !entry.dir && /\.xlsx$/i.test(entry.name));

  if (!excelEntry) throw new Error('ZIP sem arquivo auditoria.xlsx.');

  const excelBuffer = await excelEntry.async('arraybuffer');
  const { headers, rows } = parseWorkbookArrayBuffer(excelBuffer);

  let manifest = null;
  const manifestEntry = entries.find((entry) => !entry.dir && /manifest\.json$/i.test(entry.name));
  if (manifestEntry) {
    try { manifest = JSON.parse(await manifestEntry.async('string')); } catch { manifest = null; }
  }

  const respIdx = findHeaderIndex(headers, ['RESPONSÁVEL ENVIO', 'RESPONSAVEL ENVIO'], -1);
  const responsavel = clean(manifest?.responsavel)
    || clean(respIdx >= 0 ? rows.find((row) => clean(row[respIdx]))?.[respIdx] : '')
    || clean(getByField(headers, rows[0] || [], 'responsavel'))
    || file.name.replace(/\.zip$/i, '');

  const cenarioIdx = findHeaderIndex(headers, ['CENÁRIO', 'CENARIO'], -1);
  const tipoIdx = findHeaderIndex(headers, ['TIPO_REGISTRO'], -1);
  const countBase = rows.filter((row) => clean(tipoIdx >= 0 ? row[tipoIdx] : '').toUpperCase() !== 'EXTRA').length;
  const countExtra = rows.filter((row) => clean(tipoIdx >= 0 ? row[tipoIdx] : '').toUpperCase() === 'EXTRA' || clean(cenarioIdx >= 0 ? row[cenarioIdx] : '').toUpperCase() === 'C').length;

  const photos = [];
  for (const entry of entries) {
    if (entry.dir) continue;
    if (!/^fotos\//i.test(entry.name)) continue;
    const data = await entry.async('uint8array');
    photos.push({ originalPath: entry.name, fileName: baseNameFromPath(entry.name), data });
  }

  return {
    id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
    fileName: file.name,
    responsavel,
    headers,
    rows,
    photos,
    importedAt: new Date().toISOString(),
    countBase,
    countExtra,
  };
};

const makePhotoLookup = (pkg) => {
  const map = new Map();
  pkg.photos.forEach((photo) => {
    map.set(normKey(photo.originalPath), photo);
    map.set(normKey(photo.fileName), photo);
    map.set(normKey(`Fotos/${photo.fileName}`), photo);
  });
  return map;
};

const findPhoto = (lookup, photoPath) => {
  const raw = clean(photoPath);
  if (!raw || raw === '—') return null;
  return lookup.get(normKey(raw)) || lookup.get(normKey(baseNameFromPath(raw))) || null;
};

const buildFinalZipBlob = async ({ baseHeaders, baseRows, responsaveis, imports, manualReceived }) => {
  const finalZip = new JSZip();
  const dataImportacao = new Date().toLocaleString('pt-BR');
  const finalHeaders = [...baseHeaders, ...FINAL_COLS];
  const statusOffset = baseHeaders.length;
  const finalRows = baseRows.map((row) => [
    ...baseHeaders.map((_, i) => row[i] ?? ''),
    'PENDENTE', '', '', '', '', '', '',
  ]);

  const plaquetaToIndices = new Map();
  baseRows.forEach((row, idx) => {
    const plaqueta = clean(getByField(baseHeaders, row, 'plaqueta'));
    if (!plaqueta) return;
    const key = normKey(plaqueta);
    if (!plaquetaToIndices.has(key)) plaquetaToIndices.set(key, []);
    plaquetaToIndices.get(key).push(idx);
  });

  const conflicts = [['MOTIVO', 'RESPONSÁVEL', 'PLAQUETA', 'ID_AUDITORIA', 'CENÁRIO', 'ARQUIVO ORIGEM', 'DETALHE']];
  const nonRegistered = [['RESPONSÁVEL ENVIO', 'DATA IMPORTAÇÃO', ...baseHeaders, ...AUDIT_COLS, 'ARQUIVO FOTO FINAL']];
  const importedByResp = new Map();

  for (const pkg of imports) {
    const safeResp = safeName(pkg.responsavel, 'responsavel');
    const photoLookup = makePhotoLookup(pkg);
    importedByResp.set(normKey(pkg.responsavel), pkg);

    for (const photo of pkg.photos) {
      finalZip.file(`Fotos/${safeResp}/${photo.fileName}`, photo.data);
    }

    const idIdx = findHeaderIndex(pkg.headers, ['ID_AUDITORIA'], -1);
    const tipoIdx = findHeaderIndex(pkg.headers, ['TIPO_REGISTRO'], -1);
    const respIdx = findHeaderIndex(pkg.headers, ['RESPONSÁVEL ENVIO', 'RESPONSAVEL ENVIO'], -1);
    const cenarioIdx = findHeaderIndex(pkg.headers, ['CENÁRIO', 'CENARIO'], -1);
    const obsIdx = findHeaderIndex(pkg.headers, ['OBSERVAÇÃO AUDITORIA', 'OBSERVACAO AUDITORIA'], -1);
    const fotoIdx = findHeaderIndex(pkg.headers, ['ARQUIVO FOTO'], -1);

    for (const auditRow of pkg.rows) {
      const tipo = clean(tipoIdx >= 0 ? auditRow[tipoIdx] : '').toUpperCase();
      const cenario = clean(cenarioIdx >= 0 ? auditRow[cenarioIdx] : '').toUpperCase();
      const responsavelEnvio = clean(respIdx >= 0 ? auditRow[respIdx] : '') || pkg.responsavel;
      const plaqueta = clean(getByField(pkg.headers, auditRow, 'plaqueta'));
      const idAuditoria = clean(idIdx >= 0 ? auditRow[idIdx] : '');
      const obs = clean(obsIdx >= 0 ? auditRow[obsIdx] : '');
      const fotoPathOriginal = clean(fotoIdx >= 0 ? auditRow[fotoIdx] : '');
      const photo = findPhoto(photoLookup, fotoPathOriginal);
      const finalPhotoPath = photo ? `Fotos/${safeName(responsavelEnvio, 'responsavel')}/${photo.fileName}` : '';

      if (tipo === 'EXTRA' || cenario === 'C') {
        nonRegistered.push([
          responsavelEnvio,
          dataImportacao,
          ...baseHeaders.map((_, i) => auditRow[i] ?? ''),
          ...AUDIT_COLS.map((col) => getByHeader(pkg.headers, auditRow, [col], -1)),
          finalPhotoPath,
        ]);
        continue;
      }

      let targetIndex = -1;
      let origin = '';

      if (plaqueta) {
        const matches = plaquetaToIndices.get(normKey(plaqueta)) || [];
        if (matches.length === 1) {
          targetIndex = matches[0];
          origin = 'PLAQUETA';
        } else if (matches.length > 1) {
          conflicts.push(['PLAQUETA DUPLICADA NA BASE', responsavelEnvio, plaqueta, idAuditoria, cenario, pkg.fileName, `A plaqueta aparece ${matches.length} vezes na base.`]);
          continue;
        } else {
          conflicts.push(['PLAQUETA NÃO ENCONTRADA NA BASE', responsavelEnvio, plaqueta, idAuditoria, cenario, pkg.fileName, 'Item enviado com plaqueta inexistente na base principal.']);
          continue;
        }
      } else if (idAuditoria !== '' && !Number.isNaN(Number(idAuditoria)) && baseRows[Number(idAuditoria)]) {
        targetIndex = Number(idAuditoria);
        origin = 'ID_AUDITORIA';
      } else {
        conflicts.push(['SEM PLAQUETA E SEM ID VÁLIDO', responsavelEnvio, plaqueta, idAuditoria, cenario, pkg.fileName, 'Não foi possível cruzar automaticamente.']);
        continue;
      }

      if (targetIndex < 0) continue;
      if (finalRows[targetIndex][statusOffset] === 'RECEBIDO') {
        conflicts.push(['ITEM IMPORTADO MAIS DE UMA VEZ', responsavelEnvio, plaqueta, idAuditoria, cenario, pkg.fileName, `Já havia auditoria para a linha ${targetIndex + 2}.`]);
        continue;
      }

      finalRows[targetIndex][statusOffset] = 'RECEBIDO';
      finalRows[targetIndex][statusOffset + 1] = cenario;
      finalRows[targetIndex][statusOffset + 2] = obs;
      finalRows[targetIndex][statusOffset + 3] = finalPhotoPath;
      finalRows[targetIndex][statusOffset + 4] = responsavelEnvio;
      finalRows[targetIndex][statusOffset + 5] = dataImportacao;
      finalRows[targetIndex][statusOffset + 6] = origin;
    }
  }

  const summaryHeaders = ['RESPONSÁVEL', 'STATUS', 'ITENS BASE IMPORTADOS', 'ITENS CENÁRIO C', 'ARQUIVO IMPORTADO', 'DATA IMPORTAÇÃO'];
  const summaryRows = (responsaveis || []).map((resp) => {
    const pkg = importedByResp.get(normKey(resp));
    const manual = manualReceived?.has?.(resp);
    return [
      resp,
      pkg ? 'IMPORTADO' : manual ? 'MARCADO MANUALMENTE' : 'PENDENTE',
      pkg?.countBase || 0,
      pkg?.countExtra || 0,
      pkg?.fileName || '',
      pkg ? new Date(pkg.importedAt).toLocaleString('pt-BR') : '',
    ];
  });

  const excelArray = makeWorkbookArray([
    { name: 'Auditoria Final', rows: [finalHeaders, ...finalRows] },
    { name: 'Bens não cadastrados', rows: nonRegistered },
    { name: 'Conflitos', rows: conflicts },
    { name: 'Resumo', rows: [summaryHeaders, ...summaryRows] },
  ]);

  finalZip.file('Excel_Final.xlsx', excelArray);

  const blob = await finalZip.generateAsync({ type: 'blob' });
  return { blob, fileName: `Auditoria_Final_${todayBR()}.zip`, conflictsCount: Math.max(0, conflicts.length - 1) };
};

/* ==== src/components/TopBrand.jsx ==== */

function TopBrand({ title = 'Auditoria Patrimonial', subtitle = 'Ferramenta de inventário' }) {
  return (
    <div style={{ background: JD.green, color: '#fff', borderRadius: 12, padding: '16px 20px', marginBottom: 24, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
      <div>
        <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '.08em', opacity: .8, marginBottom: 4 }}>{title.toUpperCase()}</div>
        <div style={{ fontWeight: 900, fontSize: 18 }}>{subtitle}</div>
      </div>
      <div style={{ background: JD.yellow, color: JD.green, fontWeight: 900, fontSize: 11, padding: '7px 14px', borderRadius: 7, letterSpacing: '.06em', whiteSpace: 'nowrap' }}>CIARAMA</div>
    </div>
  );
}

/* ==== src/components/PhotoSlot.jsx ==== */

function PhotoSlot({ foto, fotoName, color, mandatory, onChange }) {
  const [processing, setProcessing] = useState(false);
  const missing = mandatory && !foto;

  const handleFile = async (file) => {
    if (!file || processing) return;

    try {
      setProcessing(true);
      await onChange(file);
    } catch (error) {
      console.error('Erro ao aplicar foto:', error);
      alert('Não foi possível aplicar a foto. Tente novamente.');
    } finally {
      setProcessing(false);
    }
  };

  if (!foto) {
    return (
      <label style={{
        display: 'inline-flex', alignItems: 'center', gap: 7, padding: '8px 14px',
        border: `2px ${missing ? 'dashed' : 'solid'} ${missing ? P.B : P.border}`,
        borderRadius: 9, background: missing ? '#FFF5F5' : '#fff', cursor: processing ? 'wait' : 'pointer',
        fontSize: 13, fontWeight: 600, color: missing ? P.B : P.muted, opacity: processing ? .7 : 1,
      }}>
        {processing ? 'Processando foto...' : missing ? 'Foto obrigatória' : 'Adicionar foto'}
        <input
          type="file"
          accept="image/*"
          capture="environment"
          disabled={processing}
          style={{ display: 'none' }}
          onChange={async (e) => {
            const f = e.target.files?.[0];
            e.target.value = '';
            if (f) await handleFile(f);
          }}
        />
      </label>
    );
  }

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
      <img src={foto} alt="Evidência" style={{ height: 72, borderRadius: 8, border: `1.5px solid ${color}`, objectFit: 'cover' }} />
      <div>
        <div style={{ fontSize: 12, fontWeight: 600, color, marginBottom: 4 }}>{fotoName}</div>
        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, border: `1px solid ${P.border}`, borderRadius: 6, padding: '3px 10px', fontSize: 12, color: P.muted, cursor: processing ? 'wait' : 'pointer', opacity: processing ? .7 : 1 }}>
          {processing ? 'Processando...' : 'Trocar'}
          <input
            type="file"
            accept="image/*"
            capture="environment"
            disabled={processing}
            style={{ display: 'none' }}
            onChange={async (e) => {
              const f = e.target.files?.[0];
              e.target.value = '';
              if (f) await handleFile(f);
            }}
          />
        </label>
      </div>
    </div>
  );
}

/* ==== src/pages/UploadPage.jsx ==== */

function UploadPage({ onFile, autoStatus }) {
  const fileRef = useRef(null);
  const [dragging, setDragging] = useState(false);

  return (
    <div style={{ background: P.bg, minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '2rem' }}>
      <div style={{ width: '100%', maxWidth: 520 }}>
        <TopBrand subtitle="Ferramenta de inventário" />

        <div
          onClick={() => fileRef.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => { e.preventDefault(); setDragging(false); onFile(e.dataTransfer.files?.[0]); }}
          style={{
            border: `2px dashed ${dragging ? JD.green : P.borderDark}`,
            borderRadius: 14,
            padding: '2.5rem 2rem',
            textAlign: 'center',
            background: dragging ? JD.greenPale : P.card,
            cursor: 'pointer',
            transition: 'all .2s',
            marginBottom: 16,
          }}
        >
          <div style={{ fontSize: 38, marginBottom: 10 }}>Arquivo Excel</div>
          <p style={{ fontWeight: 700, fontSize: 15, margin: '0 0 5px' }}>Arraste ou clique para selecionar</p>
          <p style={{ color: P.muted, fontSize: 13, margin: 0 }}>Relatorio_Patrimonial.xlsx</p>
        </div>

        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={(e) => onFile(e.target.files?.[0])} />

        <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 16, fontSize: 13, color: P.mid, lineHeight: 1.5 }}>
          <b>Carregamento automático:</b> o sistema tenta abrir a planilha da pasta <code>public</code>. Use o nome <code>Relatorio_Patrimonial.xlsx</code>.
          {autoStatus && <div style={{ marginTop: 8, color: P.muted }}>{autoStatus}</div>}
        </div>
      </div>
    </div>
  );
}

/* ==== src/pages/SelectPage.jsx ==== */

function SelectPage({ headers, allData, responsaveis, onSelect, onImportMode, onReloadBase }) {
  // Um cartão por filial de cada responsável.
  const resumo = useMemo(() => {
    return responsaveis.flatMap((resp) => {
      const porFilial = new Map();
      allData.forEach((row) => {
        if (clean(getByField(headers, row, 'responsavel')) !== resp) return;
        const filial = clean(getByField(headers, row, 'filial'));
        porFilial.set(filial, (porFilial.get(filial) || 0) + 1);
      });
      return [...porFilial.entries()]
        .sort(([a], [b]) => a.localeCompare(b, 'pt-BR'))
        .map(([filial, total]) => ({ resp, filial, total }));
    });
  }, [headers, allData, responsaveis]);

  return (
    <div style={{ background: P.bg, minHeight: '100vh', padding: '2rem 1.5rem' }}>
      <div style={{ maxWidth: 860, margin: '0 auto' }}>
        <TopBrand subtitle="Selecione o responsável" />

        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, marginBottom: 18, flexWrap: 'wrap' }}>
          <button onClick={onReloadBase} style={{ background: '#fff', border: `1px solid ${P.border}`, borderRadius: 10, padding: '10px 14px', fontWeight: 700, cursor: 'pointer' }}>Trocar planilha base</button>
          <button onClick={onImportMode} style={{ background: JD.green, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 18px', fontWeight: 800, cursor: 'pointer' }}>Importar Arquivo</button>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(240px,1fr))', gap: 12 }}>
          {resumo.map((item) => (
            <button key={`${item.resp}|${item.filial}`} onClick={() => onSelect(item.resp, item.filial)} style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: '16px 18px', textAlign: 'left', cursor: 'pointer' }}>
              <div style={{ fontWeight: 900, fontSize: 15, marginBottom: 4 }}>{item.resp}</div>
              <div style={{ color: P.muted, fontSize: 12 }}>{item.filial || 'Filial não informada'}</div>
              <div style={{ marginTop: 10, fontSize: 13, fontWeight: 800, color: JD.green }}>{item.total} itens</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ==== src/pages/AuditPage.jsx ==== */

const compressImage = (file, maxSize = 900, quality = 0.65) => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);

    img.onload = () => {
      try {
        let width = img.width;
        let height = img.height;

        if (width > height && width > maxSize) {
          height = Math.round((height * maxSize) / width);
          width = maxSize;
        } else if (height > maxSize) {
          width = Math.round((width * maxSize) / height);
          height = maxSize;
        }

        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;

        const ctx = canvas.getContext('2d');
        if (!ctx) {
          URL.revokeObjectURL(objectUrl);
          reject(new Error('Não foi possível preparar a imagem.'));
          return;
        }

        ctx.drawImage(img, 0, 0, width, height);

        canvas.toBlob(
          (blob) => {
            URL.revokeObjectURL(objectUrl);

            if (!blob) {
              reject(new Error('Não foi possível compactar a imagem.'));
              return;
            }

            const reader = new FileReader();

            reader.onloadend = () => {
              const fileName = file.name
                ? file.name.replace(/\.[^.]+$/, '.jpg')
                : `foto_${Date.now()}.jpg`;

              resolve({
                dataUrl: reader.result,
                fileName,
              });
            };

            reader.onerror = () => reject(new Error('Erro ao converter a imagem.'));
            reader.readAsDataURL(blob);
          },
          'image/jpeg',
          quality
        );
      } catch (error) {
        URL.revokeObjectURL(objectUrl);
        reject(error);
      }
    };

    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error('Erro ao carregar a imagem.'));
    };

    img.src = objectUrl;
  });
};

const readPhoto = async (file, cb) => {
  if (!file) return;

  try {
    const { dataUrl, fileName } = await compressImage(file);
    cb(dataUrl, fileName);
  } catch (error) {
    console.error('Erro ao processar foto:', error);
    alert('Não foi possível carregar a foto. Tente tirar novamente ou escolher uma imagem menor.');
  }
};

function AuditPage({ headers, allData, selectedResp, initialFilial, responses, setResponses, extraItems, setExtraItems, savedAt, onBack }) {
  const [search, setSearch] = useState('');
  const [areaFilter, setAreaFilter] = useState('');
  const [filialFilter, setFilialFilter] = useState(initialFilial || '');
  const [onlyUnmarked, setOnlyUnmarked] = useState(false);
  const [exporting, setExporting] = useState(false);

  const myItems = useMemo(() => allData.map((row, idx) => ({ row, idx }))
    .filter(({ row }) => clean(getByField(headers, row, 'responsavel')) === selectedResp), [headers, allData, selectedResp]);

  const areas = useMemo(() => [...new Set(myItems.map(({ row }) => clean(getByField(headers, row, 'area'))).filter(Boolean))].sort(), [headers, myItems]);

  const filiais = useMemo(() => [...new Set(myItems.map(({ row }) => clean(getByField(headers, row, 'filial'))).filter(Boolean))].sort(), [headers, myItems]);

  const filtered = useMemo(() => {
    const s = search.toLowerCase();
    return myItems.filter(({ row, idx }) => {
      const plaqueta = clean(getByField(headers, row, 'plaqueta')).toLowerCase();
      const descricao = clean(getByField(headers, row, 'descricao')).toLowerCase();
      const area = clean(getByField(headers, row, 'area'));
      const filial = clean(getByField(headers, row, 'filial'));
      return (!s || plaqueta.includes(s) || descricao.includes(s))
        && (!areaFilter || area === areaFilter)
        && (!filialFilter || filial === filialFilter)
        && (!onlyUnmarked || !responses[idx]?.cenario);
    });
  }, [headers, myItems, search, areaFilter, filialFilter, onlyUnmarked, responses]);

  // Agrupa os itens visíveis por filial, só para exibição.
  const gruposFilial = useMemo(() => {
    const map = new Map();
    filtered.forEach((item) => {
      const filial = clean(getByField(headers, item.row, 'filial')) || 'Filial não informada';
      if (!map.has(filial)) map.set(filial, []);
      map.get(filial).push(item);
    });
    return [...map.entries()].sort(([a], [b]) => a.localeCompare(b, 'pt-BR'));
  }, [headers, filtered]);

  const stats = useMemo(() => {
    const A = myItems.filter(({ idx }) => responses[idx]?.cenario === 'A').length;
    const B = myItems.filter(({ idx }) => responses[idx]?.cenario === 'B').length;
    const notMarked = myItems.length - A - B;
    const extra = extraItems.length;
    return { total: myItems.length, A, B, extra, done: A + B, pct: myItems.length ? Math.round(((A + B) / myItems.length) * 100) : 0, notMarked };
  }, [myItems, responses, extraItems]);

  const setR = (idx, key, value) => setResponses((prev) => ({ ...prev, [idx]: { ...prev[idx], [key]: value } }));
  const addExtra = () => setExtraItems((prev) => [...prev, { id: Date.now(), desc: '', plaqueta: '', obs: '', foto: null, fotoName: '' }]);
  const upExtra = (id, key, value) => setExtraItems((prev) => prev.map((item) => item.id === id ? { ...item, [key]: value } : item));

  const salvarFotoItem = async (idx, file) => {
    await readPhoto(file, (dataUrl, name) => {
      setResponses((prev) => ({
        ...prev,
        [idx]: {
          ...prev[idx],
          foto: dataUrl,
          fotoName: name,
        },
      }));
    });
  };

  const salvarFotoExtra = async (id, file) => {
    await readPhoto(file, (dataUrl, name) => {
      setExtraItems((prev) => prev.map((item) => (
        item.id === id
          ? {
              ...item,
              foto: dataUrl,
              fotoName: name,
            }
          : item
      )));
    });
  };

  const validar = (acao) => {
    const pendencias = validateAudit({ myItems, responses, extraItems, headers });
    if (!pendencias.length) return true;
    const lista = pendencias.slice(0, 12).map((p) => `- ${p}`).join('\n');
    const resto = pendencias.length > 12 ? `\n\nE mais ${pendencias.length - 12} pendência(s).` : '';
    alert(`Não é possível ${acao}.\n\nExistem ${pendencias.length} pendência(s):\n\n${lista}${resto}`);
    return false;
  };

  const gerarZip = async () => buildAuditZipBlob({ headers, myItems, responses, extraItems, selectedResp });

  const baixarZip = async () => {
    if (!validar('baixar o ZIP')) return;
    setExporting(true);
    try {
      const { blob, fileName } = await gerarZip();
      downloadBlob(blob, fileName);
    } catch (error) {
      alert(`Erro ao gerar ZIP: ${error.message}`);
    } finally {
      setExporting(false);
    }
  };

  const enviarAuditoria = async () => {
    if (!validar('enviar a auditoria')) return;
    setExporting(true);
    try {
      const { blob, fileName } = await gerarZip();
      const response = await fetch('/api/enviar-auditoria', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/zip',
          'x-file-name': fileName,
          'x-responsavel': selectedResp,
        },
        body: blob,
      });

      let data = null;
      try { data = await response.json(); } catch { data = null; }
      if (!response.ok) throw new Error(data?.message || data?.error || `Erro HTTP ${response.status}`);
      alert('Auditoria enviada com sucesso.');
    } catch (error) {
      alert(`Não foi possível enviar a auditoria: ${error.message}`);
    } finally {
      setExporting(false);
    }
  };

  const filial = filialFilter || clean(getByField(headers, myItems[0]?.row || [], 'filial'));

  return (
    <div style={{ background: P.bg, minHeight: '100vh' }}>
      <div style={{ position: 'sticky', top: 0, zIndex: 10, background: P.dark, color: '#fff', padding: '10px 20px' }}>
        <div style={{ maxWidth: 960, margin: '0 auto' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, marginBottom: 6 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <button onClick={onBack} style={{ background: 'none', border: 'none', color: '#aaa', cursor: 'pointer', fontSize: 13, padding: 0 }}>Voltar</button>
              <div><span style={{ fontWeight: 800, fontSize: 15 }}>{selectedResp}</span><span style={{ color: '#aaa', fontSize: 12, marginLeft: 10 }}>{filial}</span></div>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              {savedAt && <span style={{ fontSize: 11, color: '#4ADE80' }}>Salvo {savedAt.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}</span>}
              <button onClick={baixarZip} disabled={exporting} style={{ background: 'transparent', color: '#ccc', border: '1px solid #555', borderRadius: 7, padding: '5px 10px', fontSize: 12, cursor: 'pointer' }}>Baixar ZIP</button>
              <button onClick={enviarAuditoria} disabled={exporting} style={{ background: exporting ? '#555' : '#22C55E', color: '#fff', border: 'none', borderRadius: 7, padding: '6px 14px', fontWeight: 800, fontSize: 13, cursor: 'pointer' }}>{exporting ? 'Processando...' : 'Enviar Auditoria'}</button>
            </div>
          </div>
          <div style={{ height: 4, background: '#333', borderRadius: 2 }}><div style={{ height: '100%', width: `${stats.pct}%`, background: stats.pct === 100 ? '#22C55E' : '#3B82F6', borderRadius: 2, transition: 'width .4s' }} /></div>
          <div style={{ display: 'flex', gap: 14, marginTop: 5, fontSize: 11, flexWrap: 'wrap' }}>
            <span style={{ color: '#4ADE80' }}>A: {stats.A}</span>
            <span style={{ color: '#F87171' }}>B: {stats.B}</span>
            <span style={{ color: '#FCD34D' }}>C: {stats.extra}</span>
            <span style={{ color: '#aaa' }}>{stats.done}/{stats.total}</span>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 960, margin: '0 auto', padding: '18px 16px' }}>
        <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 14, display: 'grid', gridTemplateColumns: '1fr 190px 190px 150px', gap: 10, marginBottom: 14 }}>
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar por plaqueta ou descrição" style={{ border: `1px solid ${P.border}`, borderRadius: 8, padding: '9px 10px' }} />
          <select value={areaFilter} onChange={(e) => setAreaFilter(e.target.value)} style={{ border: `1px solid ${P.border}`, borderRadius: 8, padding: '9px 10px' }}>
            <option value="">Todas as áreas</option>
            {areas.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
          <select value={filialFilter} onChange={(e) => setFilialFilter(e.target.value)} style={{ border: `1px solid ${P.border}`, borderRadius: 8, padding: '9px 10px' }}>
            <option value="">Todas as filiais</option>
            {filiais.map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13 }}><input type="checkbox" checked={onlyUnmarked} onChange={(e) => setOnlyUnmarked(e.target.checked)} /> Não marcados</label>
        </div>

        {gruposFilial.map(([nomeFilial, itens]) => {
          const feitos = itens.filter(({ idx }) => responses[idx]?.cenario).length;
          return (
          <div key={nomeFilial} style={{ marginBottom: 22 }}>
            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: 10, borderBottom: `2px solid ${JD.green}`, padding: '4px 2px 6px', marginBottom: 12 }}>
              <div style={{ fontWeight: 900, fontSize: 16, color: JD.greenDark }}>{nomeFilial}</div>
              <div style={{ fontSize: 12, color: P.muted, fontWeight: 700 }}>{feitos}/{itens.length} marcados</div>
            </div>
        {itens.map(({ row, idx }) => {
          const r = responses[idx] || {};
          const plaqueta = clean(getByField(headers, row, 'plaqueta')) || 'Sem plaqueta';
          const descricao = clean(getByField(headers, row, 'descricao')) || 'Sem descrição';
          const area = clean(getByField(headers, row, 'area')) || clean(getByField(headers, row, 'centroCusto'));
          const filial = clean(getByField(headers,row, 'Filial')) || 'Filial não informada'
          return (
            <div key={idx} style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, marginBottom: 12, overflow: 'hidden' }}>
              <div style={{ padding: '14px 16px', borderBottom: `1px solid ${P.border}` }}>
                <div style={{ fontSize: 12, color: P.muted, marginBottom: 4 }}>Linha {idx + 2} · Filial:<b>{filial}</b> · Plaqueta: <b>{plaqueta}</b> · {area}</div>
                <div style={{ fontWeight: 900, fontSize: 15 }}>{descricao}</div>
              </div>
              <div style={{ padding: 14, display: 'grid', gap: 10 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2,minmax(0,1fr))', gap: 10 }}>
                  {['A', 'B'].map((c) => (
                    <button key={c} onClick={() => setR(idx, 'cenario', c)} style={{ background: r.cenario === c ? CC[c].activeBg : CC[c].bg, border: `1.5px solid ${r.cenario === c ? CC[c].color : CC[c].border}`, borderRadius: 9, padding: '10px 12px', cursor: 'pointer', color: CC[c].color, fontWeight: 900 }}>{CC[c].label}</button>
                  ))}
                </div>

                {r.cenario === 'A' && (
                  
                // <PhotoSlot foto={r.foto} fotoName={r.fotoName} color={P.A} mandatory onChange={(file) => salvarFotoItem(idx, file)} />
                  <div style={{ display: 'grid', gap: 8 }}>
                      <PhotoSlot foto={r.foto} fotoName={r.fotoName} color={P.A} mandatory onChange={(file) => salvarFotoItem(idx, file)} />
                      <textarea value={r.obs || ''} onChange={(e) => setR(idx, 'obs', e.target.value)} placeholder="Observação opcional" style={{ minHeight: 70, resize: 'vertical', border: `1px solid ${P.border}`, borderRadius: 8, padding: 10 }} />
                  </div>
                )}

                {r.cenario === 'B' && (
                  <div style={{ display: 'grid', gap: 8 }}>
                    <textarea value={r.obs || ''} onChange={(e) => setR(idx, 'obs', e.target.value)} placeholder="Observação obrigatória: informe onde o bem pode estar ou a justificativa." style={{ minHeight: 70, resize: 'vertical', border: `1px solid ${!clean(r.obs) ? P.B : P.border}`, borderRadius: 8, padding: 10 }} />
                    <PhotoSlot foto={r.foto} fotoName={r.fotoName} color={P.B} onChange={(file) => salvarFotoItem(idx, file)} />
                  </div>
                )}
              </div>
            </div>
          );
        })}
          </div>
          );
        })}

        <div style={{ background: '#FFF8E0', border: `1.5px solid ${P.Cborder}`, borderRadius: 12, padding: 16, marginTop: 18 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 12 }}>
            <div><b>Cenário C — Bens não cadastrados</b><div style={{ color: P.muted, fontSize: 12 }}>Use para bens encontrados fisicamente que não aparecem na base.</div></div>
            <button onClick={addExtra} style={{ background: P.C, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 12px', fontWeight: 800, cursor: 'pointer' }}>Adicionar item</button>
          </div>

          {extraItems.map((item, i) => (
            <div key={item.id} style={{ background: '#fff', border: `1px solid ${P.Cborder}`, borderRadius: 10, padding: 12, marginBottom: 10 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 160px', gap: 10, marginBottom: 10 }}>
                <input value={item.desc} onChange={(e) => upExtra(item.id, 'desc', e.target.value)} placeholder="Descrição obrigatória" style={{ border: `1px solid ${!clean(item.desc) ? P.B : P.border}`, borderRadius: 8, padding: 10 }} />
                <input value={item.plaqueta} onChange={(e) => upExtra(item.id, 'plaqueta', e.target.value)} placeholder="Plaqueta opcional" style={{ border: `1px solid ${P.border}`, borderRadius: 8, padding: 10 }} />
              </div>
              <textarea value={item.obs} onChange={(e) => upExtra(item.id, 'obs', e.target.value)} placeholder="Observação opcional" style={{ width: '100%', minHeight: 55, border: `1px solid ${P.border}`, borderRadius: 8, padding: 10, marginBottom: 10 }} />
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                <PhotoSlot foto={item.foto} fotoName={item.fotoName} color={P.C} mandatory onChange={(file) => salvarFotoExtra(item.id, file)} />
                <button onClick={() => setExtraItems((prev) => prev.filter((x) => x.id !== item.id))} style={{ background: 'none', border: 'none', color: P.B, fontWeight: 800, cursor: 'pointer' }}>Remover</button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ==== src/pages/ImportPage.jsx ==== */

function ImportPage({ baseHeaders, baseRows, responsaveis, onBack }) {
  const inputRef = useRef(null);
  const [imports, setImports] = useState([]);
  const [manualReceived, setManualReceived] = useState(new Set());
  const [busy, setBusy] = useState(false);
  const [lastMessage, setLastMessage] = useState('');

  const importedRespSet = useMemo(() => new Set(imports.map((pkg) => normalizeText(pkg.responsavel))), [imports]);

  const statuses = useMemo(() => responsaveis.map((resp) => {
    const imported = importedRespSet.has(normalizeText(resp));
    const manual = manualReceived.has(resp);
    return { resp, status: imported ? 'IMPORTADO' : manual ? 'MARCADO MANUALMENTE' : 'PENDENTE' };
  }), [responsaveis, importedRespSet, manualReceived]);

  const addFiles = async (files) => {
    const list = Array.from(files || []).filter((file) => /\.zip$/i.test(file.name));
    if (!list.length) return;

    setBusy(true);
    setLastMessage('');
    const novos = [];
    const erros = [];

    for (const file of list) {
      try {
        const pkg = await readAuditZip(file);
        novos.push(pkg);
      } catch (error) {
        erros.push(`${file.name}: ${error.message}`);
      }
    }

    if (novos.length) setImports((prev) => [...prev, ...novos]);
    setLastMessage(`${novos.length} arquivo(s) importado(s).${erros.length ? ` Erros: ${erros.join(' | ')}` : ''}`);
    setBusy(false);
  };

  const toggleManual = (resp) => {
    setManualReceived((prev) => {
      const next = new Set(prev);
      if (next.has(resp)) next.delete(resp);
      else next.add(resp);
      return next;
    });
  };

  const gerarFinal = async () => {
    if (!imports.length) {
      alert('Importe pelo menos um ZIP antes de gerar o Excel final.');
      return;
    }
    setBusy(true);
    try {
      const { blob, fileName, conflictsCount } = await buildFinalZipBlob({
        baseHeaders,
        baseRows,
        responsaveis,
        imports,
        manualReceived,
      });
      downloadBlob(blob, fileName);
      setLastMessage(`Arquivo final gerado. Conflitos encontrados: ${conflictsCount}.`);
    } catch (error) {
      alert(`Erro ao gerar arquivo final: ${error.message}`);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ background: P.bg, minHeight: '100vh', padding: '2rem 1.5rem' }}>
      <div style={{ maxWidth: 980, margin: '0 auto' }}>
        <button onClick={onBack} style={{ background: 'none', border: 'none', color: P.muted, cursor: 'pointer', fontSize: 13, marginBottom: 20, padding: 0 }}>Voltar</button>
        <TopBrand title="Importação de Auditorias" subtitle="Importe os ZIPs recebidos por e-mail" />

        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0,1.1fr) minmax(300px,.9fr)', gap: 16, alignItems: 'start' }}>
          <div>
            <div
              onClick={() => inputRef.current?.click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); addFiles(e.dataTransfer.files); }}
              style={{ border: `2px dashed ${P.borderDark}`, borderRadius: 12, padding: '2rem', textAlign: 'center', background: P.card, cursor: 'pointer', marginBottom: 14 }}
            >
              <div style={{ fontSize: 28, fontWeight: 900, color: JD.green, marginBottom: 8 }}>Importar Arquivo</div>
              <p style={{ fontWeight: 700, fontSize: 14, margin: '0 0 4px' }}>Arraste os ZIPs recebidos ou clique</p>
              <p style={{ color: P.muted, fontSize: 12, margin: 0 }}>Pode selecionar um ou vários arquivos .zip</p>
              <input ref={inputRef} type="file" accept=".zip" multiple style={{ display: 'none' }} onChange={(e) => addFiles(e.target.files)} />
            </div>

            {lastMessage && <div style={{ background: P.infoBg, color: P.info, border: `1px solid #BFDBFE`, borderRadius: 10, padding: 12, fontSize: 13, marginBottom: 12 }}>{lastMessage}</div>}

            {imports.map((pkg, i) => (
              <div key={pkg.id} style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 10, padding: '12px 14px', display: 'flex', alignItems: 'center', gap: 12, marginBottom: 8 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 900 }}>{pkg.responsavel}</div>
                  <div style={{ fontSize: 12, color: P.muted }}>{pkg.fileName} · {pkg.countBase} itens base · {pkg.countExtra} bens não cadastrados</div>
                </div>
                <button onClick={() => setImports((prev) => prev.filter((_, idx) => idx !== i))} style={{ background: 'none', border: 'none', color: P.B, cursor: 'pointer', fontWeight: 800 }}>Remover</button>
              </div>
            ))}

            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, marginTop: 12 }}>
              <div>
                <div style={{ fontWeight: 900 }}>{imports.length} ZIP(s) importado(s)</div>
                <div style={{ fontSize: 12, color: P.muted }}>O botão gera um pacote final com Excel_Final.xlsx e pasta Fotos por responsável.</div>
              </div>
              <button onClick={gerarFinal} disabled={busy || !imports.length} style={{ background: JD.green, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 18px', fontWeight: 900, cursor: 'pointer', whiteSpace: 'nowrap' }}>{busy ? 'Processando...' : 'Gerar Excel Final'}</button>
            </div>
          </div>

          <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <div>
                <div style={{ fontWeight: 900 }}>Controle de recebimento</div>
                <div style={{ fontSize: 12, color: P.muted }}>Marque manualmente quando precisar.</div>
              </div>
              <div style={{ color: JD.green, fontWeight: 900 }}>{statuses.filter((s) => s.status !== 'PENDENTE').length}/{statuses.length}</div>
            </div>

            <div style={{ display: 'grid', gap: 7, maxHeight: 560, overflow: 'auto' }}>
              {statuses.map((s) => (
                <div key={s.resp} style={{ display: 'flex', alignItems: 'center', gap: 8, border: `1px solid ${P.border}`, borderRadius: 9, padding: '8px 10px', background: s.status === 'PENDENTE' ? '#fff' : JD.greenPale }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 800, fontSize: 13 }}>{s.resp}</div>
                    <div style={{ fontSize: 11, color: s.status === 'PENDENTE' ? P.muted : JD.green }}>{s.status}</div>
                  </div>
                  {!importedRespSet.has(normalizeText(s.resp)) && (
                    <button onClick={() => toggleManual(s.resp)} style={{ background: manualReceived.has(s.resp) ? P.B : JD.green, color: '#fff', border: 'none', borderRadius: 7, padding: '6px 8px', fontWeight: 800, fontSize: 11, cursor: 'pointer' }}>{manualReceived.has(s.resp) ? 'Desmarcar' : 'Marcar'}</button>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>

        <div style={{ background: P.warnBg, border: `1px solid ${P.warn}`, borderRadius: 12, padding: 14, marginTop: 16, color: P.mid, fontSize: 13, lineHeight: 1.5 }}>
          O Excel final mantém os dados da base principal e adiciona as colunas de auditoria. Itens sem plaqueta são cruzados pelo ID_AUDITORIA. Itens do Cenário C vão para a aba “Bens não cadastrados”.
        </div>
      </div>
    </div>
  );
}

/* ==== src/App.jsx ==== */

const AUTO_EXCEL_CANDIDATES = [
  '/Relatorio_Patrimonial.xlsx',
  '/Relatório_Patrimonial.xlsx',
  '/relatorio-patrimonial.xlsx',
  '/auditoria-patrimonial.xlsx',
  '/Relatório-Patrimonial.xlsx',
];

function App() {
  const [stage, setStage] = useState('upload');
  const [headers, setHeaders] = useState([]);
  const [allData, setAllData] = useState([]);
  const [selectedResp, setSelectedResp] = useState('');
  const [selectedFilial, setSelectedFilial] = useState('');
  const [responses, setResponses] = useState({});
  const [extraItems, setExtraItems] = useState([]);
  const [savedAt, setSavedAt] = useState(null);
  const [autoStatus, setAutoStatus] = useState('Tentando carregar planilha automática...');

  const responsaveis = useMemo(() => {
    const list = allData.map((row) => clean(getByField(headers, row, 'responsavel'))).filter(Boolean);
    return [...new Set(list)].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }, [headers, allData]);

  const applyWorkbook = (arrayBuffer) => {
    const parsed = parseWorkbookArrayBuffer(arrayBuffer);
    if (!parsed.headers.length || !parsed.rows.length) throw new Error('Planilha sem cabeçalho ou sem dados.');
    setHeaders(parsed.headers);
    setAllData(parsed.rows);
    setStage('select');
  };

  const parseFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        applyWorkbook(e.target.result);
      } catch (error) {
        alert(`Erro ao ler Excel: ${error.message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  useEffect(() => {
    let cancelled = false;

    const loadAutoExcel = async () => {
      for (const path of AUTO_EXCEL_CANDIDATES) {
        try {
          const response = await fetch(path, { cache: 'no-store' });
          if (!response.ok) continue;
          const buffer = await response.arrayBuffer();
          if (cancelled) return;
          applyWorkbook(buffer);
          setAutoStatus(`Planilha carregada automaticamente: ${path}`);
          return;
        } catch {
          // tenta próximo nome
        }
      }
      if (!cancelled) setAutoStatus('Nenhuma planilha automática foi encontrada em public. Selecione manualmente.');
    };

    loadAutoExcel();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
  if (!selectedResp || stage !== 'audit') return;

  try {
    localStorage.setItem(
      localStorageKey(selectedResp),
      JSON.stringify({
        responses,
        extraItems,
        savedAt: Date.now()
      })
    );

      setSavedAt(new Date());
    } catch (error) {
      console.error('Erro ao salvar progresso:', error);

      if (error?.name === 'QuotaExceededError') {
        alert(
          'O navegador não conseguiu salvar automaticamente porque as fotos ficaram muito pesadas. ' +
          'Continue a auditoria e gere o ZIP ao finalizar, ou reduza a quantidade/tamanho das fotos.'
        );
      }
    }
  }, [responses, extraItems, selectedResp, stage]);

  const selectResp = (resp, filial = '') => {
    setSelectedResp(resp);
    setSelectedFilial(filial);
    const saved = localStorage.getItem(localStorageKey(resp));
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setResponses(parsed.responses || {});
        setExtraItems(parsed.extraItems || []);
        setSavedAt(parsed.savedAt ? new Date(parsed.savedAt) : null);
      } catch {
        setResponses({});
        setExtraItems([]);
        setSavedAt(null);
      }
    } else {
      setResponses({});
      setExtraItems([]);
      setSavedAt(null);
    }
    setStage('audit');
  };

  if (stage === 'upload') {
    return <UploadPage onFile={parseFile} autoStatus={autoStatus} />;
  }

  if (stage === 'import') {
    return <ImportPage baseHeaders={headers} baseRows={allData} responsaveis={responsaveis} onBack={() => setStage('select')} />;
  }

  if (stage === 'audit') {
    return (
      <AuditPage
        headers={headers}
        allData={allData}
        selectedResp={selectedResp}
        initialFilial={selectedFilial}
        responses={responses}
        setResponses={setResponses}
        extraItems={extraItems}
        setExtraItems={setExtraItems}
        savedAt={savedAt}
        onBack={() => setStage('select')}
      />
    );
  }

  return (
    <SelectPage
      headers={headers}
      allData={allData}
      responsaveis={responsaveis}
      onSelect={selectResp}
      onImportMode={() => setStage('import')}
      onReloadBase={() => setStage('upload')}
    />
  );
}

export default App;
