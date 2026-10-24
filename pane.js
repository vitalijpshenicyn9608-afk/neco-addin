/* NECO Test add-in: диагностика (чтение) + запись одной ячейки с проверкой. */
'use strict';
const el = id => document.getElementById(id);
const showErr = m => { el('error').textContent = m || ''; };
const fmt = row => row.map(v => (v === '' || v === null ? '·' : String(v))).join(' | ');

async function diagnose() {
  showErr('');
  el('diagOut').textContent = 'Читаю...';
  let url = '';
  try { url = String(Office.context.document.url || ''); } catch (e) { url = ''; }
  await Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('name');
    ws.protection.load('protected');
    const used = ws.getUsedRange(true);
    used.load(['address', 'rowIndex', 'rowCount']);
    await ctx.sync();

    const lastRow = used.rowIndex + used.rowCount;          // 1-based номер последней строки
    const head = ws.getRange('A1:M1');
    head.load('values');
    const from = Math.max(2, lastRow - 2);
    const tail = ws.getRange('A' + from + ':M' + lastRow);
    tail.load(['values', 'numberFormat']);
    await ctx.sync();

    const lines = [];
    lines.push('Адрес книги из Excel: ' + (url || '(не предоставлен)'));
    lines.push('Лист: ' + ws.name);
    lines.push('Защита листа: ' + (ws.protection.protected ? 'ВКЛЮЧЕНА' : 'нет'));
    lines.push('Используемый диапазон: ' + used.address);
    lines.push('');
    lines.push('Строка 1 (A:M):');
    lines.push(fmt(head.values[0]));
    lines.push('');
    lines.push('Последние строки (' + from + '–' + lastRow + '), значения:');
    tail.values.forEach((r, i) => lines.push((from + i) + ': ' + fmt(r)));
    lines.push('');
    lines.push('Форматы ячеек последней строки:');
    lines.push(fmt(tail.numberFormat[tail.numberFormat.length - 1]));
    el('diagOut').textContent = lines.join('\n');
  });
}

async function testWrite() {
  showErr('');
  if (el('confirm').value.trim() !== 'ТЕСТ') throw Error('Напишите ТЕСТ перед записью.');
  await Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('name');
    ws.protection.load('protected');
    const sel = ctx.workbook.getSelectedRange();
    sel.load(['address', 'cellCount', 'values']);
    await ctx.sync();

    if (ws.protection.protected) throw Error('Лист защищён: запись отменена.');
    if (sel.cellCount !== 1) throw Error('Выберите ровно одну ячейку.');
    const cur = sel.values[0][0];
    if (cur !== '' && cur !== null) throw Error('Выбранная ячейка не пустая: запись отменена.');

    sel.values = [['NECO-TEST']];
    await ctx.sync();

    sel.load('values');
    await ctx.sync();
    const ok = sel.values[0][0] === 'NECO-TEST';
    el('writeOut').textContent = ok
      ? 'Записано и перечитано: ' + sel.address + '. Теперь удалите эту ячейку вручную и проверьте, что у другого сотрудника она тоже появилась.'
      : 'Проверка повторного чтения НЕ пройдена: ' + sel.address;
    el('confirm').value = '';
  });
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    showErr('Откройте надстройку внутри Excel.');
    return;
  }
  el('diag').onclick = () => diagnose().catch(e => { el('diagOut').textContent = ''; showErr(e.message); });
  el('write').onclick = () => testWrite().catch(e => showErr(e.message));
});
