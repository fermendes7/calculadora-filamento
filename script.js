const price = document.querySelector('#price');
const usedWeight = document.querySelector('#usedWeight');
const cost = document.querySelector('#cost');
const unitCost = document.querySelector('#unitCost');
const formula = document.querySelector('#formula');
const printHours = document.querySelector('#printHours');
const printMinutes = document.querySelector('#printMinutes');
const objectCount = document.querySelector('#objectCount');
const laborCost = document.querySelector('#laborCost');
const breakdownFilament = document.querySelector('#breakdownFilament');
const breakdownEnergy = document.querySelector('#breakdownEnergy');
const breakdownLabor = document.querySelector('#breakdownLabor');
const pieceMachine = document.querySelector('#pieceMachine');
const pieceTotal = document.querySelector('#pieceTotal');
const pieceFormula = document.querySelector('#pieceFormula');
const packagingCost = document.querySelector('#packagingCost');
const otherCosts = document.querySelector('#otherCosts');
const machineCostPerHour = document.querySelector('#machineCostPerHour');
const errorCost = document.querySelector('#errorCost');
const extraPackaging = document.querySelector('#extraPackaging');
const extraLabor = document.querySelector('#extraLabor');
const extraOtherCosts = document.querySelector('#extraOtherCosts');
const extraError = document.querySelector('#extraError');
const finalPrice = document.querySelector('#finalPrice');
const extraFormula = document.querySelector('#extraFormula');
const priceType = document.querySelector('#priceType');
const pricingMode = document.querySelector('#pricingMode');
const desiredProfit = document.querySelector('#desiredProfit');
const desiredProfitSuffix = document.querySelector('#desiredProfitSuffix');
const profitPercent = document.querySelector('#profitPercent');
const profitBaseCost = document.querySelector('#profitBaseCost');
const profitValue = document.querySelector('#profitValue');
const salePriceInformed = document.querySelector('#salePriceInformed');
const profitPerPiece = document.querySelector('#profitPerPiece');
const partnershipPrice = document.querySelector('#partnershipPrice');
const promotionPrice = document.querySelector('#promotionPrice');
const salePriceInput = document.querySelector('#salePriceInput');
const salePrice = document.querySelector('#salePrice');
const hourlyValue = document.querySelector('#hourlyValue');
const productName = document.querySelector('#productName');
const productLink = document.querySelector('#productLink');
const fileName = document.querySelector('#fileName');
const scaleInfo = document.querySelector('#scaleInfo');
const saveObjectCount = document.querySelector('#saveObjectCount');
const saveHourlyValue = document.querySelector('#saveHourlyValue');
const saveStatus = document.querySelector('#saveStatus');
const saveProductButton = document.querySelector('#saveProduct');
const deleteProductButton = document.querySelector('#deleteProduct');

// Cole aqui a URL do Web App do Google Apps Script ligado à planilha.
const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbyYjHsA_-XGa1XqjnKubwGLcMuhUJsuIEkFP6s5dlh3XP64TbVHYYULUrJ25LOxIXSo/exec';
const profitFormula = document.querySelector('#profitFormula');
const shippingCost = document.querySelector('#shippingCost');
const fixedFee = document.querySelector('#fixedFee');
const discountPercent = document.querySelector('#discountPercent');

const money = value => value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const moneyPrecise = value => `R$ ${Number(value || 0).toLocaleString('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })}`;
const number = value => value.toLocaleString('pt-BR', { maximumFractionDigits: 1 });

function calculate() {
  const paid = Number(price.value) || 0;
  const total = 1000;
  const grams = Number(usedWeight.value) || 0;
  const perGram = total > 0 ? paid / total : 0;
  cost.textContent = money(perGram * grams);
  unitCost.textContent = `${money(perGram)}/g`;
  formula.textContent = `${money(paid)} ÷ ${number(total)} g × ${number(grams)} g`;
  calculatePiece();
}

function calculatePiece() {
  const quantity = Math.max(1, Math.floor(Number(objectCount.value) || 1));
  const printerPower = 0.14;
  const energyRate = 0.85;
  const paid = Number(price.value) || 0;
  const total = 1000;
  const grams = Number(usedWeight.value) || 0;
  const filament = total > 0 ? (paid / total) * grams * quantity : 0;
  const power = printerPower;
  const rate = energyRate;
  const hours = Number(printHours.value) || 0;
  const minutes = Math.min(59, Math.max(0, Number(printMinutes.value) || 0));
  const labor = (Number(laborCost.value) || 0) * quantity;
  const totalHours = hours + minutes / 60;
  const energy = power * totalHours * rate * quantity;
  const machine = (Number(machineCostPerHour.value) || 0) * totalHours * quantity;
  breakdownFilament.textContent = money(filament);
  breakdownEnergy.textContent = moneyPrecise(energy);
  breakdownLabor.textContent = money(labor);
  pieceMachine.textContent = money(machine);
  pieceTotal.textContent = money(filament + energy + labor + machine);
  pieceFormula.textContent = `${money(filament)} + (${number(power)} kW × ${number(hours)} h ${number(minutes)} min × ${money(rate)}) + ${money(labor)} + (${number(hours)} h ${number(minutes)} min × ${money(Number(machineCostPerHour.value) || 0)} máquina)`;
}

function calculateExtra() {
  const quantity = Math.max(1, Math.floor(Number(objectCount.value) || 1));
  const paid = Number(price.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const filament = (paid / 1000) * grams * quantity;
  const power = 0.14;
  const rate = 0.85;
  const hours = Number(printHours.value) || 0;
  const minutes = Math.min(59, Math.max(0, Number(printMinutes.value) || 0));
  const labor = (Number(laborCost.value) || 0) * quantity;
  const pieceTotalValue = filament + (power * (hours + minutes / 60) * rate * quantity) + labor;
  const packaging = (Number(packagingCost.value) || 0) * quantity;
  const others = (Number(otherCosts.value) || 0) * quantity;
  const machine = Number(machineCostPerHour.value) * (hours + minutes / 60) * quantity || 0;
  const errorPercent = Math.min(100, Math.max(0, Number(errorCost.value) || 0));
  const totalPieceWithMachine = pieceTotalValue + machine;
  const error = totalPieceWithMachine * (errorPercent / 100);
  extraPackaging.textContent = money(packaging);
  extraLabor.textContent = money(labor);
  extraOtherCosts.textContent = money(others);
  extraError.textContent = money(error);
  extraError.previousElementSibling.textContent = `Custo de erro (${number(errorPercent)}%)`;
  finalPrice.textContent = money(pieceTotalValue + packaging + others + machine + error);
  extraFormula.textContent = `${money(pieceTotalValue)} (inclui ${money(labor)} de mão de obra) + ${money(packaging)} + ${money(others)} + ${money(machine)} + ${number(errorPercent)}% de erro (${money(error)})`;
}

function calculateProfit() {
  const quantity = Math.max(1, Math.floor(Number(objectCount.value) || 1));
  const paid = Number(price.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const printTimeHours = (Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60);
  const pieceCost = (((paid / 1000) * grams) + (0.14 * printTimeHours * 0.85) + (Number(laborCost.value) || 0)) * quantity;
  const packaging = (Number(packagingCost.value) || 0) * quantity;
  const others = (Number(otherCosts.value) || 0) * quantity;
  const machine = Number(machineCostPerHour.value) * ((Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60)) * quantity || 0;
  const errorPercent = Math.min(100, Math.max(0, Number(errorCost.value) || 0));
  const pieceTotalValue = pieceCost;
  const totalPieceWithMachine = pieceTotalValue + machine;
  const error = totalPieceWithMachine * (errorPercent / 100);
  const shipping = Number(shippingCost.value) || 0;
  const fixed = Number(fixedFee.value) || 0;
  const baseCost = pieceCost + packaging + others + machine + error + shipping;
  const percent = Number(profitPercent.value) || 0;
  const discount = Number(discountPercent.value) || 0;
  const feeRate = Math.min(99.99, Math.max(0, percent)) / 100;
  const discountRate = Math.min(99.99, Math.max(0, discount)) / 100;
  const requestedProfit = Math.max(0, Number(desiredProfit.value) || 0);
  const profitRate = pricingMode.value === 'margin' ? requestedProfit / 100 : 0;
  const divisor = Math.max(0.0001, 1 - feeRate - discountRate);
  const sale = pricingMode.value === 'margin'
    ? ((baseCost + fixed) * (1 + profitRate)) / divisor
    : (baseCost + fixed + requestedProfit) / Math.max(0.0001, 1 - feeRate - discountRate);
  const fee = sale * feeRate + sale * discountRate + fixed;
  const informedSale = Number(salePriceInput.value) || 0;
  const informedFee = informedSale * feeRate + informedSale * discountRate + fixed;
  const actualProfit = informedSale - baseCost - informedFee;
  const totalPrintHours = (Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60);
  profitBaseCost.textContent = money(baseCost);
  profitValue.textContent = money(fee);
  salePriceInformed.textContent = money(informedSale);
  profitPerPiece.textContent = money(actualProfit);
  // Usa o custo final com custos adicionais como base dos canais.
  // Ex.: R$ 12,70 -> Parcerias R$ 16,51 e Promoção R$ 19,05.
  partnershipPrice.textContent = money(baseCost * 1.3);
  promotionPrice.textContent = money(baseCost * 1.5);
  salePrice.textContent = money(sale);
  hourlyValue.textContent = money(totalPrintHours > 0 ? sale / totalPrintHours : 0);
  desiredProfitSuffix.textContent = pricingMode.value === 'margin' ? '%' : 'R$';
  profitFormula.textContent = pricingMode.value === 'margin'
    ? `(${money(baseCost)} + ${money(fixed)}) × (1 + ${number(requestedProfit)}% lucro) ÷ (1 − ${number(percent)}% comissão − ${number(discount)}% desconto) = ${money(sale)}`
    : `(${money(baseCost)} + ${money(fixed)} + ${money(requestedProfit)} de lucro) ÷ (1 − ${number(percent)}% comissão − ${number(discount)}% desconto) = ${money(sale)}`;
}

function updateAll() {
  calculate();
  calculateExtra();
  calculateProfit();
  if (saveObjectCount && objectCount) saveObjectCount.value = objectCount.value || 1;
  if (saveHourlyValue) {
    const value = Number(hourlyValue.textContent.replace(/[^0-9,-]/g, '').replace('.', '').replace(',', '.')) || 0;
    saveHourlyValue.value = value.toFixed(2);
  }
}

saveProductButton.addEventListener('click', async () => {
  const name = productName.value.trim();
  if (!name) {
    saveStatus.textContent = 'Informe o nome do produto.';
    productName.focus();
    return;
  }
  if (!SHEETS_ENDPOINT) {
    saveStatus.textContent = 'A conexão com a planilha ainda precisa ser configurada no Apps Script.';
    return;
  }
  updateAll();
  const totalHours = (Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60);
  const quantity = Math.max(1, Math.floor(Number(objectCount.value) || 1));
  // Salvar exatamente os valores exibidos na calculadora. A calibragem não é
  // acrescentada apenas no salvamento, pois isso fazia a planilha divergir da tela.
  const adjustedTotalHours = totalHours;
  const filamentValue = (Number(price.value) || 0) / 1000 * (Number(usedWeight.value) || 0) * quantity;
  const energyValue = 0.14 * adjustedTotalHours * 0.85 * quantity;
  const machineValue = (Number(machineCostPerHour.value) || 0) * adjustedTotalHours * quantity;
  const laborValue = (Number(laborCost.value) || 0) * quantity;
  const pieceBaseValue = filamentValue + energyValue + laborValue + machineValue;
  const errorPercentValue = Math.min(100, Math.max(0, Number(errorCost.value) || 0));
  const errorValue = pieceBaseValue * errorPercentValue / 100;
  const packagingValue = (Number(packagingCost.value) || 0) * quantity;
  const otherCostsValue = (Number(otherCosts.value) || 0) * quantity;
  const additionalValue = packagingValue + otherCostsValue + errorValue;
  const finalCostValue = pieceBaseValue + additionalValue;
  const payload = {
    nome: name,
    link: productLink.value.trim(),
    nomeArquivo: fileName.value.trim(),
    // A planilha formata ESCALA como percentual; envie 100% como 1,
    // 65% como 0,65, evitando 10000% ou 6500%.
    escala: Number(scaleInfo.value.replace('%', '').replace(',', '.')) || 0,
    gramas: (Number(usedWeight.value) || 0) * quantity,
    precoRolo: Number(price.value) || 0,
    horas: Number(printHours.value) || 0,
    minutos: Number(printMinutes.value) || 0,
    tempoHoras: adjustedTotalHours,
    tempo: `${String(Math.floor(adjustedTotalHours)).padStart(2, '0')}:${String(Math.round((adjustedTotalHours % 1) * 60)).padStart(2, '0')}:00`,
    objetos: Number(saveObjectCount.value) || quantity,
    tempoPreparacaoMinutos: 0,
    energia: energyValue,
    filamento: filamentValue,
    maquinaTotal: machineValue,
    erroValor: errorValue,
    custosAdicionais: additionalValue,
    custoBase: pieceBaseValue,
    maoDeObra: laborValue,
    maquinaPorHora: Number(machineCostPerHour.value) || 0,
    custoErroPercentual: Number(errorCost.value) || 0,
    embalagem: packagingValue,
    outrosCustos: otherCostsValue,
    frete: Number(shippingCost.value) || 0,
    custoPeca: pieceBaseValue,
    custoTotal: finalCostValue,
    precoSugerido: Number(salePrice.textContent.replace(/[^0-9,-]/g, '').replace('.', '').replace(',', '.')) || 0,
    valorPorHora: Number(hourlyValue.textContent.replace(/[^0-9,-]/g, '').replace('.', '').replace(',', '.')) || 0,
    lucroDesejado: Number(desiredProfit.value) || 0,
    discountPercent: Number(discountPercent.value) || 0,
    canal: priceType.value
  };
  saveProductButton.disabled = true;
  saveStatus.textContent = 'Salvando...';
  try {
    await fetch(SHEETS_ENDPOINT, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: JSON.stringify(payload)
    });
    saveStatus.textContent = 'Dados enviados. Confira a aba PRODUTOS na planilha.';
  } catch (error) {
    saveStatus.textContent = 'Não foi possível salvar. Verifique a conexão com o Apps Script.';
  } finally {
    saveProductButton.disabled = false;
  }
});

deleteProductButton.addEventListener('click', async () => {
  const name = productName.value.trim();
  if (!name) {
    saveStatus.textContent = 'Informe o nome exato do produto para apagar.';
    productName.focus();
    return;
  }
  if (!confirm(`Apagar o produto "${name}" da planilha?`)) return;
  deleteProductButton.disabled = true;
  saveStatus.textContent = 'Apagando...';
  try {
    await fetch(SHEETS_ENDPOINT, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: JSON.stringify({ acao: 'apagar', nome: name })
    });
    saveStatus.textContent = 'Solicitação de exclusão enviada. Confira a aba PRODUTOS.';
  } catch (error) {
    saveStatus.textContent = 'Não foi possível apagar. Verifique a conexão com o Apps Script.';
  } finally {
    deleteProductButton.disabled = false;
  }
});

const allInputs = document.querySelectorAll('input');
allInputs.forEach(input => input.addEventListener('input', () => {
  updateAll();
}));

priceType.addEventListener('change', () => {
  const selected = priceType.options[priceType.selectedIndex];
  profitPercent.value = selected.dataset.rate;
  fixedFee.value = selected.dataset.fixed;
  discountPercent.value = selected.dataset.discount;
  updateAll();
});
pricingMode.addEventListener('change', updateAll);
updateAll();






