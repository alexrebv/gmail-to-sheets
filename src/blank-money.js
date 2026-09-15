// Цена и сумма в сводном бланке заказа.
//
// Бланк один и тот же в трёх местах: его собирает gmail-to-sheets при разборе
// почты (orderExcel.js), приложение при повторной выгрузке (export-tel.js) и
// корректировка (export-corr.js). Раскладка обязана совпадать - человек
// сверяет файлы глазами, - поэтому колонка цены и формула суммы живут здесь,
// а не переписываются в каждом месте по-своему.
//
// Колонки: 1 Название | 2 Артикул | 3 Фасовка | 4 Цена | 5.. объекты | Итого
//
// Сумма считается формулой, а не числом: в бланке правят количества руками, и
// сумма должна пересчитываться сама. SUMPRODUCT берёт столбец цены и столбец
// объекта - пустая клетка идёт за ноль.

const PRICE_COL = 4;              // куда встаёт цена
const FIRST_OBJ_COL = 5;          // с какой колонки начинаются объекты
const PRICE_HEAD = 'Цена, р.';
const SUM_HEAD = 'Сумма заказа, р.';
const MONEY_FMT = '# ##0.00';
const PRICE_WIDTH = 11;

function colLetter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

/** Сколько всего колонок в бланке при таком числе объектов. */
const totalCols = objects => FIRST_OBJ_COL - 1 + objects + 1;

/** Буква первой и последней колонки объектов. */
const objRange = objects => ({
  first: colLetter(FIRST_OBJ_COL),
  last: colLetter(FIRST_OBJ_COL + objects - 1),
});

/**
 * Формула суммы по одной колонке: цена на количество, сложенное по всем
 * строкам позиций. Цена закреплена ($D), колонка объекта - нет: формулу
 * протягивают вправо.
 */
function sumFormula(col, firstRow, lastRow) {
  const p = colLetter(PRICE_COL);
  return `SUMPRODUCT($${p}$${firstRow}:$${p}$${lastRow},${col}${firstRow}:${col}${lastRow})`;
}

/**
 * Строка «Сумма заказа, р.»: по колонке на объект плюс итог справа.
 * Итог складывает сами суммы объектов - так он сходится с ними в копейку,
 * а не считается второй раз по колонке «Итого».
 */
function sumRowValues(objects, firstRow, lastRow, sumRowN) {
  const vals = [SUM_HEAD, null, null, null];
  for (let i = 0; i < objects; i++) {
    vals.push({ formula: sumFormula(colLetter(FIRST_OBJ_COL + i), firstRow, lastRow) });
  }
  const { first, last } = objRange(objects);
  vals.push({ formula: `SUM(${first}${sumRowN}:${last}${sumRowN})` });
  return vals;
}

module.exports = { PRICE_COL, FIRST_OBJ_COL, PRICE_HEAD, SUM_HEAD, MONEY_FMT, PRICE_WIDTH,
                   colLetter, totalCols, objRange, sumFormula, sumRowValues };
