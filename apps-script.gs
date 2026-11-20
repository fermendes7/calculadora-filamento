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

  let lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  const hasLinkColumn = headers.some(header => ['link', 'url', 'linkproduto'].includes(normalize(header)));
  if (data.link && !hasLinkColumn) {
    lastColumn += 1;
    sheet.getRange(HEADER_ROW, lastColumn).setValue('LINK');
    headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  }
  if (headers.length >= 7 && normalize(headers[6]) === 'custo') {
    sheet.getRange(HEADER_ROW, 7).setValue('Custo de erro');
    headers[6] = 'Custo de erro';
  }
  if (headers.length >= 11 && ['montagem', 'custosadicionais'].includes(normalize(headers[10]))) {
    sheet.getRange(HEADER_ROW, 11).setValue('Outros custos totais');
    headers[10] = 'Outros custos totais';
  }
  const normalizedHeaders = headers.map(normalize);
  // Na planilha atual o cabeçalho está na linha 2 e o produto fica na coluna B.
  // O fallback precisa ser 2 para nunca gravar o nome na coluna Nº (A).
  const nameColumn = 2;
  const values = sheet.getDataRange().getDisplayValues();
  const formulas = sheet.getDataRange().getFormulas();
  let row = -1;
  for (let i = HEADER_ROW; i < values.length; i++) {
    const displayName = normalize(values[i][nameColumn - 1]);
    const formulaName = normalize(formulas[i][nameColumn - 1]);
    if (displayName === normalize(data.nome) || formulaName.includes(normalize(data.nome))) {
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
    const value = valueForHeader(normalize(header), data, index);
    return value === undefined ? (currentRow[index] || '') : value;
  });
  // A coluna A é a numeração correspondente à posição do produto na coluna B.
  if (output.length >= 1) output[0] = row - HEADER_ROW;
  sheet.getRange(row, 1, 1, output.length).setValues([output]);
  sheet.getRange(row, 1, 1, output.length).setFontWeight('bold');
  if (data.link) {
    const safeUrl = String(data.link).replace(/"/g, '""');
    const safeName = String(data.nome).replace(/"/g, '""');
    sheet.getRange(row, nameColumn).setFormula(`=HYPERLINK("${safeUrl}";"${safeName}")`);
  }
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

function valueForHeader(header, data, index) {
  const aliases = {
    nome: 'nome', produto: 'nome', nomeproduto: 'nome',
    link: 'link', url: 'link', linkproduto: 'link',
    gramas: 'gramas', peso: 'gramas', objeto: 'objetos', objetos: 'objetos',
    preco: 'precoSugerido', precorolo: 'precoRolo', precofilamento: 'precoRolo',
    escala: undefined, esc: undefined,
    tempo: 'tempo', horas: 'horas', minutos: 'minutos', tempohoras: 'tempoHoras',
    custo: 'custoPeca', custoobjeto: 'custoPeca', custodeerro: 'erroValor',
    maodeobra: 'maoDeObra', maquina: 'maquinaTotal', manutencao: 'maquinaTotal',
    montagem: 'outrosCustos', energia: 'energia',
    embalagem: 'embalagem', embal: 'embalagem', outroscustos: 'outrosCustos', outrocustostotais: 'outrosCustos', frete: 'frete',
    custopeca: 'custoPeca', custototal: 'custoTotal', precosugerido: 'precoSugerido',
    erro: 'custoErroPercentual', custoerro: 'custoErroPercentual',
    valorporhora: 'valorPorHora', lucro: 'lucroDesejado', canal: 'canal',
    valor: 'precoSugerido', percentualvenda: 'lucroDesejado'
  };
  const key = aliases[header];
  if (key) return data[key] ?? '';
  // Colunas repetidas da planilha antiga: preserva os valores existentes
  // quando não há um cabeçalho específico para o campo.
  if (header === 'preco' && index === 2) return data.precoSugerido ?? '';
  return undefined;
}

function normalize(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
