const SHEET_NAME = 'PRODUTOS';
const SPREADSHEET_ID = '1EBM7_ac7XW0DI5IXg8p_VdiFZxjBshEvG3nFQeqmj-c';
const HEADER_ROW = 2;

function doPost(e) {
  const raw = e.parameter && e.parameter.payload
    ? e.parameter.payload
    : (e.postData && e.postData.contents ? e.postData.contents : '{}');
  const data = JSON.parse(raw);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return json({ error: 'Aba PRODUTOS não encontrada.' });

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  const normalizedHeaders = headers.map(normalize);
  // Na planilha atual o cabeçalho está na linha 2 e o produto fica na coluna B.
  // O fallback precisa ser 2 para nunca gravar o nome na coluna Nº (A).
  const nameColumn = 2;
  const values = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = HEADER_ROW; i < values.length; i++) {
    if (normalize(values[i][nameColumn - 1]) === normalize(data.nome)) {
      row = i + 1;
      break;
    }
  }
  const wasExisting = row >= 0;
  if (data.acao === 'apagar') {
    if (row < 0) return json({ ok: false, action: 'not_found', product: data.nome });
    sheet.deleteRow(row);
    SpreadsheetApp.flush();
    return json({ ok: true, action: 'deleted', row: row, product: data.nome });
  }
  if (row < 0) {
    // Usa exclusivamente a coluna B: a primeira célula vazia após o cabeçalho.
    row = HEADER_ROW + 1;
    while (row <= sheet.getMaxRows() && normalize(sheet.getRange(row, nameColumn).getDisplayValue())) row++;
    if (row > sheet.getMaxRows()) sheet.insertRowAfter(sheet.getMaxRows());
  }

  const currentRow = row <= sheet.getLastRow()
    ? sheet.getRange(row, 1, 1, lastColumn).getValues()[0]
    : [];

  const output = headers.map((header, index) => {
    const value = valueForHeader(normalize(header), data);
    return value === undefined ? (currentRow[index] || '') : value;
  });
  sheet.getRange(row, 1, 1, output.length).setValues([output]);
  return json({ action: wasExisting ? 'updated' : 'added', row: row });
}

function doGet() {
  return json({ ok: true, message: 'Integração da calculadora ativa.' });
}

function findColumn(headers, names) {
  for (const name of names) {
    const index = headers.indexOf(name);
    if (index >= 0) return index + 1;
  }
  return 0;
}

function valueForHeader(header, data) {
  const aliases = {
    nome: 'nome', produto: 'nome', nomeproduto: 'nome',
    gramas: 'gramas', peso: 'gramas',
    precorolo: 'precoRolo', precofilamento: 'precoRolo',
    horas: 'horas', minutos: 'minutos', tempohoras: 'tempoHoras',
    maodeobra: 'maoDeObra', maquina: 'maquinaPorHora', manutencao: 'maquinaPorHora',
    custoerro: 'custoErroPercentual', erro: 'custoErroPercentual',
    embalagem: 'embalagem', outroscustos: 'outrosCustos', frete: 'frete',
    custopeca: 'custoPeca', precosugerido: 'precoSugerido',
    valorporhora: 'valorPorHora', lucro: 'lucroDesejado', canal: 'canal'
  };
  const key = aliases[header];
  return key ? (data[key] ?? '') : undefined;
}

function normalize(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
