const price = document.querySelector('#price');
const rollWeight = document.querySelector('#rollWeight');
const usedWeight = document.querySelector('#usedWeight');
const cost = document.querySelector('#cost');
const unitCost = document.querySelector('#unitCost');
const formula = document.querySelector('#formula');

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
}

[price, rollWeight, usedWeight].forEach(input => input.addEventListener('input', calculate));
calculate();
