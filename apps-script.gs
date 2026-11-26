const SHEET_NAME = 'PRODUTOS';
const SPREADSHEET_ID = '1EBM7_ac7XW0DI5IXg8p_VdiFZxjBshEvG3nFQeqmj-c';
const HEADER_ROW = 2;

// Recalcula os preços derivados quando qualquer custo é alterado diretamente
// na aba PRODUTOS.
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  const firstRow = Math.max(e.range.getRow(), HEADER_ROW + 1);
  const lastRow = e.range.getLastRow();
  if (firstRow <= lastRow) applyResponsiveFormulas_(sheet, firstRow, lastRow);
}

function configurarGatilho() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'onEdit')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();
  return 'Gatilho instalado';
}

function doPost(e) {
  const raw = e.parameter && e.parameter.payload
    ? e.parameter.payload
    : (e.postData && e.postData.contents ? e.postData.contents : '{}');
  const data = JSON.parse(raw);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return json({ error: 'Aba PRODUTOS não encontrada.' });
  if (data.acao === 'reordenar') return reorderColumns_(sheet);
  if (data.acao === 'normalizarEstrutura') return normalizarEstrutura_(sheet);
  if (data.acao === 'corrigirErros') return corrigirErros_(sheet);
  if (data.acao === 'recalcular') {
    applyResponsiveFormulas_(sheet);
    SpreadsheetApp.flush();
    return json({ ok: true, action: 'recalculated' });
  }

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
  // Padroniza o produto na planilha: a coluna B sempre recebe maiúsculas.
  const normalizedProductName = String(data.nome || '').trim().toUpperCase();
  const normalizedData = Object.assign({}, data, { nome: normalizedProductName });
  const values = sheet.getDataRange().getDisplayValues();
  const formulas = sheet.getDataRange().getFormulas();
  let row = -1;
  for (let i = HEADER_ROW; i < values.length; i++) {
    const displayName = normalize(values[i][nameColumn - 1]);
    const formulaName = normalize(formulas[i][nameColumn - 1]);
    if (displayName === normalize(normalizedProductName) || formulaName.includes(normalize(normalizedProductName))) {
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
    const value = valueForHeader(normalize(header), normalizedData, index);
    return value === undefined ? (currentRow[index] || '') : value;
  });
  // A coluna A é a numeração correspondente à posição do produto na coluna B.
  if (output.length >= 1) output[0] = row - HEADER_ROW;
  sheet.getRange(row, 1, 1, output.length).setValues([output]);
  const savedHeaders = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const objectColumn = savedHeaders.findIndex(h => normalize(h) === 'objeto') + 1;
  const costColumn = savedHeaders.findIndex(h => normalize(h) === 'custo') + 1;
  if (objectColumn) sheet.getRange(row, objectColumn).setNumberFormat('0');
  const energyColumn = savedHeaders.findIndex(h => normalize(h) === 'energia') + 1;
  if (energyColumn) sheet.getRange(row, energyColumn).setNumberFormat('R$ #,##0.0000');
  if (costColumn) {
    const rawCost = sheet.getRange(row, costColumn).getValue();
    const numericCost = typeof rawCost === 'number' ? rawCost : Number(String(rawCost).replace(',', '.'));
    // Parcerias e Promoção usam o custo final com custos adicionais,
    // igual à calculadora: custo da peça + embalagem + outros custos + erro.
    const totalCost = Number(data.custoTotal);
    const channelBase = Number.isFinite(totalCost) && totalCost !== 0 ? totalCost : numericCost;
    for (const rule of [{ name: 'promocao', factor: 1.5 }, { name: 'parcerias', factor: 1.3 }]) {
      const channelColumn = savedHeaders.findIndex(h => normalize(h) === rule.name) + 1;
      if (channelColumn) {
        if (Number.isFinite(channelBase) && channelBase !== 0) {
          sheet.getRange(row, channelColumn).setValue(channelBase * rule.factor);
        } else {
          sheet.getRange(row, channelColumn).clearContent();
        }
      }
    }
  }
  sheet.getRange(row, 1, 1, output.length).setFontWeight('bold');
  if (data.link) {
    const url = String(data.link).trim();
    const safeName = normalizedProductName;
    // Grava o URL também na coluna LINK e o nome como rich text em B.
    // Isso torna o salvamento idempotente: tanto produto novo quanto existente
    // recebem/atualizam o link sempre.
    sheet.getRange(row, linkColumn).setValue(url);
    sheet.getRange(row, nameColumn).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(safeName).setLinkUrl(url).build()
    );
  }
  applyResponsiveFormulas_(sheet, row);
  if (!data.link && linkColumn) sheet.getRange(row, linkColumn).clearContent();
  SpreadsheetApp.flush();
  return json({ action: wasExisting ? 'updated' : 'added', row: row });
}

function doGet() {
  return json({ ok: true, message: 'Integração da calculadora ativa.' });
}

function normalizarEstrutura_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return json({ ok: false, error: 'Sem colunas.' });
  const all = sheet.getRange(HEADER_ROW, 1, Math.max(1, lastRow - HEADER_ROW + 1), lastColumn).getValues();
  const headers = all[0].map(h => String(h || '').trim());
  const normalized = headers.map(normalize);
  const wanted = [
    ['produto'], ['preco'], ['parcerias','parceria'], ['promocao'], ['lucroporhora','valorporhora','rhora'],
    ['gramas','grama','peso'], ['tempo'], ['escala','esc','esca'], ['objeto','objetos'],
    ['maodeobra','maoobra'], ['maquina','manutencao'], ['energia'], ['embalagem','embal'],
    ['outroscustostotais','outroscustos','outros'], ['custodeerro','erro'], ['custo','custofinal'],
    ['qtdmesa','qtd'], ['valor','precosugerido','custosadd'],
    ['desconto'], ['frete'], ['nomedoarquivo','nomearquivo','arquivo'], ['link','url']
  ];
  const used = new Set();
  const order = [];
  const canonical = ['PRODUTO','PREÇO','PARCERIAS','PROMOÇÃO','LUCRO POR HORA','GRAMAS','TEMPO','ESCA.','OBJETO','MÃO DE OBRA','MÁQUINA','ENERGIA','EMBALAGEM','OUTROS CUSTOS TOTAIS','CUSTO DE ERRO','CUSTO','QTD MESA','CUSTOS + ADD','DESCONTO','FRETE','NOME DO ARQUIVO','LINK'];
  const outHeaders = [];
  const sourceIndexes = [];
  wanted.forEach((names, pos) => {
    const idx = normalized.findIndex((h, i) => !used.has(i) && names.includes(h));
    used.add(idx);
    sourceIndexes.push(idx >= 0 ? idx : -1);
    outHeaders.push(canonical[pos]);
  });
  const rows = all.slice(1).map(row => sourceIndexes.map(i => i >= 0 ? row[i] : ''));
  const width = outHeaders.length;
  sheet.getRange(HEADER_ROW, 1, Math.max(1, lastRow - HEADER_ROW + 1), width).clearContent();
  sheet.getRange(HEADER_ROW, 1, 1, width).setValues([outHeaders]);
  if (rows.length) sheet.getRange(HEADER_ROW + 1, 1, rows.length, width).setValues(rows);
  const dataStart = HEADER_ROW + 1;
  const dataRows = Math.max(1, lastRow - HEADER_ROW);
  const fmt = (label, pattern) => { const col = outHeaders.findIndex(h => normalize(h) === normalize(label)) + 1; if (col) sheet.getRange(dataStart, col, dataRows, 1).setNumberFormat(pattern); };
  ['PREÇO','PARCERIAS','PROMOÇÃO','LUCRO POR HORA','MÃO DE OBRA','MÁQUINA','ENERGIA','EMBALAGEM','OUTROS CUSTOS TOTAIS','CUSTO DE ERRO','CUSTO','VALOR','FRETE'].forEach(h => fmt(h, 'R$ #,##0.00'));
  ['GRAMAS','OBJETO','QTD MESA'].forEach(h => fmt(h, '0.00'));
  fmt('TEMPO', '[h]:mm:ss');
  ['ESCA.','% VENDA','DESCONTO'].forEach(h => fmt(h, '0%'));
  const productCol = outHeaders.findIndex(h => normalize(h) === 'produto') + 1;
  if (productCol && rows.length) {
    const productValues = sheet.getRange(HEADER_ROW + 1, productCol, rows.length, 1).getValues();
    sheet.getRange(HEADER_ROW + 1, productCol, rows.length, 1).setValues(productValues.map(r => [String(r[0] || '').toUpperCase()]));
  }
  applyResponsiveFormulas_(sheet, HEADER_ROW + 1, Math.max(HEADER_ROW + 1, lastRow));
  SpreadsheetApp.flush();
  return json({ ok: true, action: 'normalized', columns: outHeaders });
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
    parcerias: 'parcerias', parceria: 'parcerias',
    promocao: 'promocao', promocao: 'promocao',
    nomedoarquivo: 'nomeArquivo', arquivo: 'nomeArquivo', file: 'nomeArquivo', filename: 'nomeArquivo',
    link: 'link', url: 'link', linkproduto: 'link',
    gramas: 'gramas', peso: 'gramas', objeto: 'objetos', objetos: 'objetos', qtd: 'objetos', qtdmesa: 'objetos',
    preco: 'precoSugerido', precorolo: 'precoRolo', precofilamento: 'precoRolo',
    escala: 'escala', esc: 'escala', esca: 'escala', escalatempo: 'escala',
    tempo: 'tempo', horas: 'horas', minutos: 'minutos', tempohoras: 'tempoHoras',
    custo: 'custoTotal', custoobjeto: 'custoPeca',
    // A coluna "Custo de erro" da planilha representa a taxa percentual (ex.: 20%).
    // O valor calculado em reais fica na coluna "Erro".
    // A coluna exibida como ERRO % recebe a taxa, não o valor em reais.
    custodeerro: 'custoErroPercentual', erropercentual: 'custoErroPercentual',
    maodeobra: 'maoDeObra', maoobra: 'maoDeObra',
    // A coluna MÁQUINA deve receber o total calculado pela calculadora
    // (Máquina/manutenção), e não apenas a tarifa por hora informada.
    maquina: 'maquinaTotal', manutencao: 'maquinaTotal',
    montagem: 'outrosCustos', energia: 'energia',
    embalagem: 'embalagem', embal: 'embalagem', outroscustos: 'outrosCustos', outroscustostotais: 'outrosCustos', outrocustostotais: 'outrosCustos', frete: 'frete',
    custopeca: 'custoPeca', custototal: 'custoTotal', precosugerido: 'precoSugerido',
    // O cabeçalho ERRO % recebe a taxa percentual da calculadora.
    // Nunca usar erroValor aqui, pois ele é o custo em reais e viraria 56%.
    erro: 'custoErroPercentual', erropercentual: 'custoErroPercentual', custoerro: 'custoErroPercentual',
    valorporhora: 'valorPorHora', lucroporhora: 'valorPorHora', lucroporh: 'valorPorHora', rhora: 'valorPorHora',
    valorhora: 'valorPorHora', lucro: 'lucroDesejado', canal: 'canal',
    valor: 'precoSugerido', custosadd: 'custoTotal', percentualvenda: 'lucroDesejado', venda: 'lucroDesejado', desconto: 'discountPercent', total: 'precoSugerido'
  };
  const key = aliases[header];
  if (key) {
    // Percentuais são armazenados como fração para a formatação da planilha:
    // 20 -> 0,20 (20%) e 150 -> 1,50 (150%).
    if (key === 'custoErroPercentual' || key === 'lucroDesejado') {
      const percent = Number(data[key]);
      return Number.isFinite(percent) ? percent / 100 : '';
    }
    if (key === 'escala') {
      const escala = Number(data[key]);
      return Number.isFinite(escala) ? escala / 100 : '';
    }
    return data[key] ?? '';
  }
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

function columnLetter_(column) {
  let result = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    column = Math.floor((column - 1) / 26);
  }
  return result;
}

function applyResponsiveFormulas_(sheet, onlyRow, onlyLastRow) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getDisplayValues()[0];
  const col = name => headers.findIndex(h => normalize(h) === name) + 1;
  const custo = col('custo');
  const parceria = col('parcerias');
  const promocao = col('promocao');
  const escala = col('escala') || col('escalaproduto');
  const erro = col('custodeerro') || col('erro');
  const venda = 0;
  const preco = col('preco');
  const tempo = col('tempo');
  const hora = col('lucroporhora') || col('valorporhora') || col('rhora');
  const first = onlyRow || HEADER_ROW + 1;
  const last = onlyLastRow || onlyRow || sheet.getLastRow();
  for (let row = first; row <= last; row++) {
    if (!String(sheet.getRange(row, 2).getDisplayValue()).trim()) continue;
    if (escala) {
      let v = Number(sheet.getRange(row, escala).getValue());
      if (!Number.isFinite(v) || v < 0.5) v = 1;
      else if (v > 1) v /= 100;
      sheet.getRange(row, escala).setValue(v).setNumberFormat('0%');
    }
    if (erro) {
      let v = Number(sheet.getRange(row, erro).getValue());
      // Preserve o valor recebido da calculadora. Só aplica 20% quando a
      // célula está realmente vazia ou inválida.
      if (!Number.isFinite(v) || v < 0) v = 0.2;
      else if (v > 1) v /= 100;
      sheet.getRange(row, erro).setValue(v).setNumberFormat('0%');
    }
    if (venda) {
      let v = Number(sheet.getRange(row, venda).getValue());
      if (!Number.isFinite(v) || v <= 0) v = 1.5;
      else if (v > 3) v /= 100;
      sheet.getRange(row, venda).setValue(v).setNumberFormat('0%');
    }
    if (custo) {
      const base = `${columnLetter_(custo)}${row}`;
      if (parceria) sheet.getRange(row, parceria).setFormula(`=${base}*13/10`);
      if (promocao) sheet.getRange(row, promocao).setFormula(`=${base}*15/10`);
    }
    if (preco && tempo && hora) {
      sheet.getRange(row, hora).setFormula(`=IFERROR(${columnLetter_(preco)}${row}/(${columnLetter_(tempo)}${row}*24),0)`);
    }
  }
  SpreadsheetApp.flush();
}

function reorderColumns_(sheet) {
  // Remove somente a coluna antiga "ESCA.". A nova coluna "ESCALA"
  // criada pelo usuário deve permanecer e ficar na coluna F.
  let headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  // Remove a coluna antiga "VENDA". A coluna correta permanece como
  // "% VENDA" / "PERCENTUAL VENDA".
  for (let i = headers.length - 1; i >= 0; i--) {
    if (['venda', 'percentualvenda'].includes(normalize(headers[i]))) sheet.deleteColumn(i + 1);
  }
  headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (let i = headers.length - 1; i >= 0; i--) {
    if (normalize(headers[i]) === 'esca') sheet.deleteColumn(i + 1);
  }
  headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let escala = headers.findIndex(h => normalize(h) === 'escala') + 1;
  if (!escala) {
    // A coluna F já foi criada manualmente pelo usuário; apenas corrige o
    // cabeçalho sem mover os valores que estão nessa coluna.
    sheet.getRange(HEADER_ROW, 6).setValue('ESCALA');
    escala = 6;
  } else if (escala !== 6) {
    sheet.moveColumns(sheet.getRange(1, escala, sheet.getMaxRows(), 1), 6);
  }
  ensureHeader_(sheet, 'QTD MESA');
  ensureHeader_(sheet, 'LUCRO POR HORA');
  // Ordem principal da planilha. A coluna A continua sendo a numeração.
  // Parcerias e Promoção ficam junto do Preço; os campos de cadastro ficam
  // depois de Escala, antes de Objeto.
  const desiredOrder = [
    'produto', 'preco', 'parcerias', 'promocao', 'gramas', 'tempo',
    'escala', 'qtdmesa', 'lucroporhora', 'objeto', 'maodeobra', 'maquina',
    'energia', 'embalagem', 'outroscustostotais', 'custodeerro', 'custo',
    'custosadd', 'frete', 'vendamkt', 'demanda',
    'nomedoarquivo', 'link'
  ];
  let target = 2;
  for (const wanted of desiredOrder) {
    headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const aliases = wanted === 'qtdmesa' ? ['qtdmesa']
      : wanted === 'lucroporhora' ? ['lucroporhora', 'valorporhora']
      : [wanted];
    let from = 0;
    for (const alias of aliases) {
      const found = headers.findIndex((h, i) => i >= 1 && normalize(h) === alias) + 1;
      if (found) { from = found; break; }
    }
    if (!from || from === target) { if (from) target++; continue; }
    sheet.moveColumns(sheet.getRange(1, from, sheet.getMaxRows(), 1), target);
    target++;
  }
  headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const valueColumn = headers.findIndex(h => ['valor', 'custosadd'].includes(normalize(h))) + 1;
  if (valueColumn) sheet.getRange(HEADER_ROW, valueColumn).setValue('CUSTOS + ADD');
  const custo = headers.findIndex(h => normalize(h) === 'custo') + 1;
  const parceria = headers.findIndex(h => normalize(h) === 'parcerias') + 1;
  const promocao = headers.findIndex(h => normalize(h) === 'promocao') + 1;
  const energyColumn = headers.findIndex(h => normalize(h) === 'energia') + 1;
  if (energyColumn && sheet.getLastRow() > HEADER_ROW) {
    sheet.getRange(HEADER_ROW + 1, energyColumn, sheet.getLastRow() - HEADER_ROW, 1).setNumberFormat('R$ #,##0.0000');
  }
  applyResponsiveFormulas_(sheet);
  SpreadsheetApp.flush();
  return json({ ok: true, action: 'reordered', escala: 'F', base: 'custo final com custos adicionais', parcerias: 'base + 30%', promocao: 'base + 50%' });
  /*
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
  */
}
