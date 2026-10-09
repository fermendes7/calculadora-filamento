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
const pieceTotal = document.querySelector('#pieceTotal');
const pieceFormula = document.querySelector('#pieceFormula');
const calculateFilamentButton = document.querySelector('#calculateFilament');
const calculatePieceButton = document.querySelector('#calculatePiece');
const packagingCost = document.querySelector('#packagingCost');
const otherCosts = document.querySelector('#otherCosts');
const machineCostPerHour = document.querySelector('#machineCostPerHour');
const extraPackaging = document.querySelector('#extraPackaging');
const extraOtherCosts = document.querySelector('#extraOtherCosts');
const extraMachine = document.querySelector('#extraMachine');
const finalPrice = document.querySelector('#finalPrice');
const extraFormula = document.querySelector('#extraFormula');
const calculateExtraButton = document.querySelector('#calculateExtra');
const priceType = document.querySelector('#priceType');
const profitPercent = document.querySelector('#profitPercent');
const profitBaseCost = document.querySelector('#profitBaseCost');
const profitValue = document.querySelector('#profitValue');
const salePrice = document.querySelector('#salePrice');
const profitFormula = document.querySelector('#profitFormula');
const calculateProfitButton = document.querySelector('#calculateProfit');

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

calculateFilamentButton.addEventListener('click', calculate);

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
  breakdownFilament.textContent = money(filament);
  breakdownEnergy.textContent = money(energy);
  breakdownLabor.textContent = money(labor);
  pieceTotal.textContent = money(filament + energy + labor);
  pieceFormula.textContent = `${money(filament)} + (${number(power)} kW × ${number(hours)} h ${number(minutes)} min × ${money(rate)}) + ${money(labor)}`;
}

calculatePieceButton.addEventListener('click', calculatePiece);

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
  extraPackaging.textContent = money(packaging);
  extraOtherCosts.textContent = money(others);
  extraMachine.textContent = money(machine);
  finalPrice.textContent = money(pieceTotalValue + packaging + others + machine);
  extraFormula.textContent = `${money(pieceTotalValue)} + ${money(packaging)} + ${money(others)} + ${money(machine)}`;
}

calculateExtraButton.addEventListener('click', calculateExtra);

function calculateProfit() {
  const paid = Number(price.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const pieceCost = (paid / 1000) * grams + (0.14 * ((Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60)) * 0.85) + (Number(laborCost.value) || 0);
  const packaging = Number(packagingCost.value) || 0;
  const others = Number(otherCosts.value) || 0;
  const machine = Number(machineCostPerHour.value) * ((Number(printHours.value) || 0) + (Math.min(59, Math.max(0, Number(printMinutes.value) || 0)) / 60)) || 0;
  const baseCost = pieceCost + packaging + others + machine;
  const percent = Number(profitPercent.value) || 0;
  const profit = baseCost * (percent / 100);
  profitBaseCost.textContent = money(baseCost);
  profitValue.textContent = money(profit);
  salePrice.textContent = money(baseCost + profit);
  profitFormula.textContent = `${money(baseCost)} + ${number(percent)}% de lucro = ${money(baseCost + profit)}`;
}

priceType.addEventListener('change', () => {
  const selected = priceType.options[priceType.selectedIndex];
  profitPercent.value = selected.dataset.profit;
});
calculateProfitButton.addEventListener('click', calculateProfit);
calculate();
calculateExtra();
calculateProfit();
