const price = document.querySelector('#price');
const rollWeight = document.querySelector('#rollWeight');
const usedWeight = document.querySelector('#usedWeight');
const cost = document.querySelector('#cost');
const unitCost = document.querySelector('#unitCost');
const formula = document.querySelector('#formula');
const pieceFilament = document.querySelector('#pieceFilament');
const printerPower = document.querySelector('#printerPower');
const energyRate = document.querySelector('#energyRate');
const printHours = document.querySelector('#printHours');
const pieceGrams = document.querySelector('#pieceGrams');
const laborCost = document.querySelector('#laborCost');
const breakdownFilament = document.querySelector('#breakdownFilament');
const breakdownEnergy = document.querySelector('#breakdownEnergy');
const breakdownLabor = document.querySelector('#breakdownLabor');
const pieceTotal = document.querySelector('#pieceTotal');
const pieceFormula = document.querySelector('#pieceFormula');

const money = value => value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const number = value => value.toLocaleString('pt-BR', { maximumFractionDigits: 1 });

function calculate() {
  const paid = Number(price.value) || 0;
  const total = Number(rollWeight.value) || 0;
  const grams = Number(usedWeight.value) || 0;
  const perGram = total > 0 ? paid / total : 0;
  cost.textContent = money(perGram * grams);
  unitCost.textContent = `${money(perGram)}/g`;
  formula.textContent = `${money(paid)} ÷ ${number(total)} g × ${number(grams)} g`;
  pieceFilament.value = (perGram * grams).toFixed(2);
  pieceGrams.value = grams;
  calculatePiece();
}

[price, rollWeight, usedWeight].forEach(input => input.addEventListener('input', calculate));

function calculatePiece() {
  const filament = Number(pieceFilament.value) || 0;
  const power = Number(printerPower.value) || 0;
  const rate = Number(energyRate.value) || 0;
  const hours = Number(printHours.value) || 0;
  const labor = Number(laborCost.value) || 0;
  const energy = power * hours * rate;
  breakdownFilament.textContent = money(filament);
  breakdownEnergy.textContent = money(energy);
  breakdownLabor.textContent = money(labor);
  pieceTotal.textContent = money(filament + energy + labor);
  pieceFormula.textContent = `${money(filament)} + (${number(power)} kW × ${number(hours)} h × ${money(rate)}) + ${money(labor)}`;
}

[pieceFilament, printerPower, energyRate, printHours, pieceGrams, laborCost].forEach(input => input.addEventListener('input', calculatePiece));
calculate();
