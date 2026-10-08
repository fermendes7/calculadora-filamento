const price = document.querySelector('#price');
const usedWeight = document.querySelector('#usedWeight');
const cost = document.querySelector('#cost');
const unitCost = document.querySelector('#unitCost');
const formula = document.querySelector('#formula');
const printerPower = document.querySelector('#printerPower');
const energyRate = document.querySelector('#energyRate');
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
  const paid = Number(price.value) || 0;
  const total = 1000;
  const grams = Number(usedWeight.value) || 0;
  const filament = total > 0 ? (paid / total) * grams : 0;
  const power = Number(printerPower.value) || 0;
  const rate = Number(energyRate.value) || 0;
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
calculate();
