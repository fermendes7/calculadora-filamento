const price = document.querySelector('#price');
const usedWeight = document.querySelector('#usedWeight');
const cost = document.querySelector('#cost');
const unitCost = document.querySelector('#unitCost');
const formula = document.querySelector('#formula');
const printHours = document.querySelector('#printHours');
const printMinutes = document.querySelector('#printMinutes');
const laborCost = document.querySelector('#laborCost');
const breakdownFilament = document.querySelector('#breakdownFilament');
const breakdownEnergy = document.querySelector('#breakdownEnergy');
const breakdownLabor = document.querySelector('#breakdownLabor');
const pieceMachine = document.querySelector('#pieceMachine');
const pieceTotal = document.querySelector('#pieceTotal');
const pieceFormula = document.querySelector('#pieceFormula');
const calculateFilamentButton = document.querySelector('#calculateFilament');
const calculatePieceButton = document.querySelector('#calculatePiece');
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
const calculateExtraButton = document.querySelector('#calculateExtra');
const priceType = document.querySelector('#priceType');
const pricingMode = document.querySelector('#pricingMode');
const desiredProfit = document.querySelector('#desiredProfit');
const desiredProfitSuffix = document.querySelector('#desiredProfitSuffix');
const profitPercent = document.querySelector('#profitPercent');
const profitBaseCost = document.querySelector('#profitBaseCost');
const profitValue = document.querySelector('#profitValue');
const salePriceInformed = document.querySelector('#salePriceInformed');
const profitPerPiece = document.querySelector('#profitPerPiece');
const salePriceInput = document.querySelector('#salePriceInput');
const salePrice = document.querySelector('#salePrice');
const profitFormula = document.querySelector('#profitFormula');
const calculateProfitButton = document.querySelector('#calculateProfit');
const shippingCost = document.querySelector('#shippingCost');
const fixedFee = document.querySelector('#fixedFee');
const discountPercent = document.querySelector('#discountPercent');

const money = value => value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
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
  const printerPower = 0.14;
  const energyRate = 0.85;
  const paid = Number(price.value) || 0;
  const total = 1000;
  const grams = Number(usedWeight.value) || 0;
  const filament = total > 0 ? (paid / total) * grams : 0;
  const power = printerPower;
  const rate = energyRate;
  const hours = Number(printHours.value) || 0;
  const minutes = Math.min(59, Math.max(0, Number(printMinutes.value) || 0));
  const labor = Number(laborCost.value) || 0;
  const totalHours = hours + minutes / 60;
  const energy = power * totalHours * rate;
  const machine = (Number(machineCostPerHour.value) || 0) * totalHours;
  breakdownFilament.textContent = money(filament);
  breakdownEnergy.textContent = money(energy);
  breakdownLabor.textContent = money(labor);
  pieceMachine.textContent = money(machineCostPerHour.value * totalHours || 0);
  pieceTotal.textContent = money(filament + energy + labor + machine);
  pieceFormula.textContent = `${money(filament)} + (${number(power)} kW × ${number(hours)} h ${number(minutes)} min × ${money(rate)}) + ${money(labor)} + (${number(hours)} h ${number(minutes)} min × ${money(Number(machineCostPerHour.value) || 0)} máquina)`;
}

function calculateExtra() {
  const paid = Number(price.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const filament = (paid / 1000) * grams;
  const power = 0.14;
  const rate = 0.85;
  const hours = Number(printHours.value) || 0;
  const minutes = Math.min(59, Math.max(0, Number(printMinutes.value) || 0));
  const labor = Number(laborCost.value) || 0;
  const pieceTotalValue = filament + (power * (hours + minutes / 60) * rate) + labor;
  const packaging = Number(packagingCost.value) || 0;
  const others = Number(otherCosts.value) || 0;
  const machine = Number(machineCostPerHour.value) * (hours + minutes / 60) || 0;
  const errorPercent = Math.min(100, Math.max(0, Number(errorCost.value) || 0));
  const error = pieceTotalValue * (errorPercent / 100);
  extraPackaging.textContent = money(packaging);
  extraLabor.textContent = money(labor);
  extraOtherCosts.textContent = money(others);
  extraError.textContent = money(error);
  extraError.previousElementSibling.textContent = `Custo de erro (${number(errorPercent)}%)`;
  finalPrice.textContent = money(pieceTotalValue + packaging + others + machine + error);
  extraFormula.textContent = `${money(pieceTotalValue)} (inclui ${money(labor)} de mão de obra) + ${money(packaging)} + ${money(others)} + ${money(machine)} + ${number(errorPercent)}% de erro (${money(error)})`;
}

function calculateProfit() {
  const paid = Number(price.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const pieceCost = (paid / 1000) * grams + (0.14 * ((Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60)) * 0.85) + (Number(laborCost.value) || 0);
  const packaging = Number(packagingCost.value) || 0;
  const others = Number(otherCosts.value) || 0;
  const machine = Number(machineCostPerHour.value) * ((Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60)) || 0;
  const errorPercent = Math.min(100, Math.max(0, Number(errorCost.value) || 0));
  const pieceTotalValue = pieceCost;
  const error = pieceTotalValue * (errorPercent / 100);
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
  profitBaseCost.textContent = money(baseCost);
  profitValue.textContent = money(fee);
  salePriceInformed.textContent = money(informedSale);
  profitPerPiece.textContent = money(actualProfit);
  salePrice.textContent = money(sale);
  desiredProfitSuffix.textContent = pricingMode.value === 'margin' ? '%' : 'R$';
  profitFormula.textContent = pricingMode.value === 'margin'
    ? `(${money(baseCost)} + ${money(fixed)}) × (1 + ${number(requestedProfit)}% lucro) ÷ (1 − ${number(percent)}% comissão − ${number(discount)}% desconto) = ${money(sale)}`
    : `(${money(baseCost)} + ${money(fixed)} + ${money(requestedProfit)} de lucro) ÷ (1 − ${number(percent)}% comissão − ${number(discount)}% desconto) = ${money(sale)}`;
}

function updateAll() {
  calculate();
  calculateExtra();
  calculateProfit();
}

calculateFilamentButton.addEventListener('click', updateAll);
calculatePieceButton.addEventListener('click', updateAll);
calculateExtraButton.addEventListener('click', updateAll);
calculateProfitButton.addEventListener('click', updateAll);

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
