const SHEET_NAME = 'PRODUTOS';
const SPREADSHEET_ID = '1EBM7_ac7XW0DI5IXg8p_VdiFZxjBshEvG3nFQeqmj-c';

function doPost(e) {
  const raw = e.parameter && e.parameter.payload
    ? e.parameter.payload
    : (e.postData && e.postData.contents ? e.postData.contents : '{}');
  const data = JSON.parse(raw);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return json({ error: 'Aba PRODUTOS não encontrada.' });

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const normalizedHeaders = headers.map(normalize);
  const nameColumn = findColumn(normalizedHeaders, ['nome', 'produto', 'nomeproduto']) || 1;
  const values = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (normalize(values[i][nameColumn - 1]) === normalize(data.nome)) {
      row = i + 1;
      break;
    }
  }
  if (row < 0) row = Math.max(sheet.getLastRow() + 1, 2);

  const output = headers.map(header => valueForHeader(normalize(header), data));
  sheet.getRange(row, 1, 1, output.length).setValues([output]);
  return json({ action: row <= values.length ? 'updated' : 'added', row: row });
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
  return key ? (data[key] ?? '') : '';
}

function normalize(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
