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
  if (data.acao === 'reordenar') return reorderColumns_(sheet);
  if (data.acao === 'corrigirErros') return corrigirErros_(sheet);

  let lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  // O nome fica sempre em B. Se a coluna de link ainda não existir, crie-a
  // explicitamente para evitar que o URL seja perdido ao montar a linha.
  let linkColumn = headers.findIndex(h => normalize(h) === 'link') + 1;
  if (!linkColumn) {
    linkColumn = lastColumn + 1;
    sheet.getRange(HEADER_ROW, linkColumn).setValue('LINK');
    lastColumn = linkColumn;
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
  ensureHeader_(sheet, 'NOME DO ARQUIVO');
  lastColumn = Math.max(sheet.getLastColumn(), 1);
  headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  linkColumn = headers.findIndex(h => normalize(h) === 'link') + 1;
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
  const savedHeaders = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const objectColumn = savedHeaders.findIndex(h => normalize(h) === 'objeto') + 1;
  const costColumn = savedHeaders.findIndex(h => normalize(h) === 'custo') + 1;
  if (objectColumn) sheet.getRange(row, objectColumn).setNumberFormat('0');
  if (costColumn) {
    const costCell = sheet.getRange(row, costColumn).getA1Notation();
    for (const rule of [{ name: 'promocao', factor: 1.5 }, { name: 'parcerias', factor: 1.3 }]) {
      const channelColumn = savedHeaders.findIndex(h => normalize(h) === rule.name) + 1;
      if (channelColumn) sheet.getRange(row, channelColumn).setFormula(`=${costCell}*${rule.factor}`);
    }
  }
  sheet.getRange(row, 1, 1, output.length).setFontWeight('bold');
  if (data.link) {
    const url = String(data.link).trim();
    const safeName = String(data.nome).trim();
    // Grava o URL também na coluna LINK e o nome como rich text em B.
    // Isso torna o salvamento idempotente: tanto produto novo quanto existente
    // recebem/atualizam o link sempre.
    sheet.getRange(row, linkColumn).setValue(url);
    sheet.getRange(row, nameColumn).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(safeName).setLinkUrl(url).build()
    );
  }
  if (!data.link && linkColumn) sheet.getRange(row, linkColumn).clearContent();
  SpreadsheetApp.flush();
  return json({ action: wasExisting ? 'updated' : 'added', row: row });
}

function doGet() {
  return json({ ok: true, message: 'Integração da calculadora ativa.' });
}

function corrigirErros_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getDisplayValues()[0];
  const costNames = ['custo', 'valor', 'total', 'frete', 'custodeerro', 'outroscustostotais'];
  const costColumns = headers.map((h, i) => costNames.includes(normalize(h)) ? i + 1 : 0).filter(Boolean);
  const linkColumn = headers.findIndex(h => normalize(h) === 'link') + 1;
  if (lastRow > HEADER_ROW && costColumns.length) {
    for (const col of costColumns) {
      const range = sheet.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1);
      const values = range.getDisplayValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]).includes('#REF!') || String(values[i][0]).includes('#ERROR!')) {
          range.getCell(i + 1, 1).clearContent();
        }
      }
    }
  }
  if (linkColumn && lastRow > HEADER_ROW) {
    const range = sheet.getRange(HEADER_ROW + 1, linkColumn, lastRow - HEADER_ROW, 1);
    const values = range.getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      const url = String(values[i][0] || '').trim();
      if (/^https?:\/\//i.test(url)) {
        range.getCell(i + 1, 1).setRichTextValue(
          SpreadsheetApp.newRichTextValue().setText('Abrir link').setLinkUrl(url).build()
        );
      }
    }
    sheet.setColumnWidth(linkColumn, 95);
  }
  SpreadsheetApp.flush();
  return json({ ok: true, action: 'errors_fixed', links_shortened: Boolean(linkColumn) });
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
    nomedoarquivo: 'nomeArquivo', arquivo: 'nomeArquivo', file: 'nomeArquivo', filename: 'nomeArquivo',
    link: 'link', url: 'link', linkproduto: 'link',
    gramas: 'gramas', peso: 'gramas', objeto: 'objetos', objetos: 'objetos',
    preco: 'precoSugerido', precorolo: 'precoRolo', precofilamento: 'precoRolo',
    escala: 'escala', esc: 'escala',
    tempo: 'tempo', horas: 'horas', minutos: 'minutos', tempohoras: 'tempoHoras',
    custo: 'custoPeca', custoobjeto: 'custoPeca',
    // A coluna "Custo de erro" da planilha representa a taxa percentual (ex.: 20%).
    // O valor calculado em reais fica na coluna "Erro".
    custodeerro: 'custoErroPercentual',
    maodeobra: 'maoDeObra', maoobra: 'maoDeObra', maquina: 'maquinaTotal', manutencao: 'maquinaTotal',
    montagem: 'outrosCustos', energia: 'energia',
    embalagem: 'embalagem', embal: 'embalagem', outroscustos: 'outrosCustos', outroscustostotais: 'outrosCustos', outrocustostotais: 'outrosCustos', frete: 'frete',
    custopeca: 'custoPeca', custototal: 'custoTotal', precosugerido: 'precoSugerido',
    erro: 'erroValor', custoerro: 'custoErroPercentual',
    valorporhora: 'valorPorHora', lucro: 'lucroDesejado', canal: 'canal',
    valor: 'precoSugerido', percentualvenda: 'lucroDesejado', venda: 'lucroDesejado', desconto: 'discountPercent', total: 'precoSugerido'
  };
  const key = aliases[header];
  if (key) return data[key] ?? '';
  // Colunas repetidas da planilha antiga: preserva os valores existentes
  // quando não há um cabeçalho específico para o campo.
  if (header === 'preco' && index === 2) return data.precoSugerido ?? '';
  return undefined;
}

function ensurePercentualVendaAfterCusto_(sheet) {
  const last = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(HEADER_ROW, 1, 1, last).getDisplayValues()[0];
  const custo = headers.findIndex(h => normalize(h) === 'custo') + 1;
  const percentual = headers.findIndex(h => ['percentualvenda','venda'].includes(normalize(h))) + 1;
  if (!custo || !percentual || percentual === custo + 1) return;
  sheet.moveColumns(sheet.getRange(1, percentual, sheet.getMaxRows(), 1), custo + 1);
}

function ensureHeader_(sheet, label) {
  const last = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(HEADER_ROW, 1, 1, last).getDisplayValues()[0];
  if (headers.some(h => normalize(h) === normalize(label))) return;
  sheet.getRange(HEADER_ROW, last + 1).setValue(label);
}

function normalize(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function reorderColumns_(sheet) {
  // Remove a coluna antiga "ERRO". O percentual já fica em "Custo de erro"
  // e o valor em reais é somado ao custo final, sem uma segunda coluna.
  // A faixa verde da linha 1 tem mesclagens antigas; desfaça somente essa
  // linha para permitir mover as colunas sem alterar os dados das linhas.
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  let last = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(HEADER_ROW, 1, 1, last).getDisplayValues()[0];
  for (let i = headers.length - 1; i >= 1; i--) {
    if (['erro', 'total'].includes(normalize(headers[i]))) sheet.deleteColumn(i + 1);
  }

  const order = ['produto','preco','gramas','tempo','escala','objeto','maodeobra','maoobra','maquina','energia','embal','outroscustostotais','custodeerro','custo','venda','percentualvenda','valor','desconto','frete','parcerias'];
  let target = 2;
  for (const wanted of order) {
    last = Math.max(sheet.getLastColumn(), 1);
    headers = sheet.getRange(HEADER_ROW, 1, 1, last).getDisplayValues()[0];
    let from = 0;
    for (let i = 1; i < headers.length; i++) {
      if (normalize(headers[i]) === wanted || (wanted === 'venda' && normalize(headers[i]) === 'percentualvenda')) { from = i + 1; break; }
    }
    if (!from) continue;
    if (from !== target) sheet.moveColumns(sheet.getRange(1, from, sheet.getMaxRows(), 1), target);
    target++;
  }
  // Nome do arquivo e link ficam sempre nas duas últimas colunas.
  for (const wanted of ['nomedoarquivo', 'link']) {
    last = Math.max(sheet.getLastColumn(), 1);
    headers = sheet.getRange(HEADER_ROW, 1, 1, last).getDisplayValues()[0];
    const from = headers.findIndex(h => normalize(h) === wanted) + 1;
    if (from && from !== last) sheet.moveColumns(sheet.getRange(1, from, sheet.getMaxRows(), 1), last);
  }
  SpreadsheetApp.flush();
  return json({ ok: true, action: 'reordered', lastColumn: sheet.getLastColumn() });
}
