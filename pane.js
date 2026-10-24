/* NECO Booking add-in. Запись только в пустые строки, с повторной проверкой и перечитыванием. */
'use strict';

// ---- Настройки ----------------------------------------------------------
// Часть имени файла, в который разрешена запись. Для тестов — клон «ПРОБА».
// Когда решите работать с настоящей книгой, замените на её имя (одной строкой) и перезалейте pane.js.
const EXPECTED_FILE_PART = 'ПРОБА';
// Сколько последних строк листа читать для проверок (дубли за 5 дней и поиск свободной строки).
const WINDOW_ROWS = 2000;
// Копировать оформление строки целиком с последней занятой строки (границы и т.п.).
// По умолчанию ВЫКЛЮЧЕНО: сначала проверьте на клоне, что условное форматирование столбца C не страдает.
const COPY_ROW_FORMATS = false;
// ------------------------------------------------------------------------

const P = NecoPlanner;
const el = id => document.getElementById(id);
let payload = null;
let shownPlanKey = '';
let busy = false;

const showErr = m => { el('error').textContent = m || ''; };
function planKey(pl) { return pl.results.map(r => r.kind + ':' + (r.row || '-') + ':' + P.key(r.item)).join(';'); }
function labelPlan(pl) {
  const lines = pl.results.map(r => {
    const it = r.item, base = it.booking + ' · ' + it.port + ' · ' + it.quantity + ' × ' + it.type;
    return r.kind === 'duplicate' ? base + ': ПРОПУЩЕН (' + r.why + ')' : base + ': строка ' + r.row;
  });
  lines.push('', 'Новых: ' + pl.added + '; дублей: ' + pl.duplicates);
  return lines.join('\n');
}
function docUrl() { try { return decodeURIComponent(String(Office.context.document.url || '')); } catch (e) { return String(Office.context.document.url || ''); } }

function checkEnv() {
  const url = docUrl();
  if (!url) throw Error('Excel не сообщил адрес книги: запись заблокирована.');
  if (!url.toUpperCase().includes(EXPECTED_FILE_PART.toUpperCase())) {
    throw Error('Открыта не та книга. Ожидается файл, в имени которого есть «' + EXPECTED_FILE_PART + '». Сейчас: ' + url);
  }
}

// Читает целевой лист: защита, заголовки, окно последних строк.
async function readSheet(ctx, name) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
  ws.load('isNullObject,name');
  await ctx.sync();
  if (ws.isNullObject) throw Error('В книге нет листа «' + name + '».');
  ws.protection.load('protected');
  const used = ws.getUsedRange(true);
  used.load('rowIndex,rowCount');
  const head = ws.getRange('A1:M1');
  head.load('values');
  await ctx.sync();

  const h = head.values[0].map(v => P.norm(v));
  const okHead = h[0].includes('ДАТА ЗАЯВКИ') && h[2].includes('НОМЕР БУКИНГА') && h[3].includes('КОЛИЧЕСТВО') &&
                 h[4].includes('ТИП') && h[6].includes('РЕЛИЗ') && h[10] === 'POL';
  if (!okHead) throw Error('Заголовки листа не совпадают со схемой A:M. Запись заблокирована.');
  if (ws.protection.protected) throw Error('Лист «' + name + '» защищён. Запись заблокирована.');

  const last = used.rowIndex + used.rowCount;
  if (last < 3) throw Error('На листе почти нет данных: запись заблокирована.');
  const first = Math.max(2, last - WINDOW_ROWS + 1);
  const win = ws.getRange('A' + first + ':M' + last);
  win.load('values');
  await ctx.sync();
  return { ws, first, last, values: win.values };
}

async function buildPlan() {
  checkEnv();
  return Excel.run(async ctx => {
    const st = await readSheet(ctx, payload.sheet);
    return P.plan(st.values, st.first, payload, new Date(), new Date());
  });
}

async function check() {
  showErr(''); el('result').textContent = ''; el('write').disabled = true; el('recheck').disabled = !payload;
  if (!payload) return;
  const pl = await buildPlan();
  shownPlanKey = planKey(pl);
  el('plan').textContent = 'Лист: ' + payload.sheet + '\n' + labelPlan(pl);
  el('write').disabled = pl.added === 0;
  if (pl.added === 0) el('result').textContent = 'Все строки уже есть в таблице. Записывать нечего.';
}

async function onPaste() {
  showErr(''); payload = null; shownPlanKey = '';
  el('write').disabled = true; el('recheck').disabled = true; el('result').textContent = '';
  const text = el('paste').value.trim();
  if (!text) { el('plan').textContent = 'Заявка не вставлена.'; return; }
  try {
    payload = JSON.parse(text);
    P.validatePayload(payload, new Date());
  } catch (e) {
    payload = null;
    el('plan').textContent = '';
    showErr(e instanceof SyntaxError ? 'В буфере не заявка NECO. Нажмите в программе «Копировать для Excel Online».' : e.message);
    return;
  }
  el('plan').textContent = 'Проверяю таблицу...';
  try { await check(); } catch (e) { el('plan').textContent = ''; showErr(e.message); }
}

async function write() {
  if (busy) return;
  busy = true; el('write').disabled = true; showErr('');
  try {
    if (!payload || !shownPlanKey) throw Error('Сначала вставьте заявку и дождитесь проверки.');
    checkEnv();
    const today = new Date();
    const serial = P.excelSerial(today);
    let report = null;

    await Excel.run(async ctx => {
      // Читаем заново прямо перед записью: за это время таблицу мог изменить коллега.
      const st = await readSheet(ctx, payload.sheet);
      const pl = P.plan(st.values, st.first, payload, today, today);
      if (planKey(pl) !== shownPlanKey) {
        shownPlanKey = planKey(pl);
        el('plan').textContent = 'Лист: ' + payload.sheet + '\n' + labelPlan(pl);
        throw Error('Таблица изменилась с момента проверки. Ничего не записано: проверьте новый план и нажмите «Записать» ещё раз.');
      }
      if (!pl.added) throw Error('Все строки уже есть в таблице. Записывать нечего.');

      const ws = st.ws;
      // Образцы оформления с реальных строк.
      const aTpl = ws.getRange('A' + pl.lastData);
      aTpl.load('numberFormat');
      let gTpl = null;
      if (pl.gTemplate) {
        gTpl = ws.getRange('G' + pl.gTemplate);
        gTpl.format.fill.load('color');
        gTpl.format.load('horizontalAlignment');
      }
      await ctx.sync();
      const aFmt = aTpl.numberFormat[0][0];
      const gFill = gTpl ? gTpl.format.fill.color : null;
      const gAlign = gTpl ? gTpl.format.horizontalAlignment : null;

      ws.activate();
      for (const r of pl.results) {
        if (r.kind !== 'new') continue;
        const n = r.row, it = r.item;
        if (COPY_ROW_FORMATS) {
          ws.getRange('A' + n + ':M' + n).copyFrom('A' + pl.lastData + ':M' + pl.lastData, Excel.RangeCopyType.formats);
        }
        const a = ws.getRange('A' + n);
        a.values = [[serial]];
        if (aFmt) a.numberFormat = [[aFmt]];
        ws.getRange('C' + n + ':F' + n).values = [[P.norm(it.booking), Number(it.quantity), P.norm(it.type), it.comment ? String(it.comment) : '']];
        const g = ws.getRange('G' + n);
        g.values = [['pending']];
        if (gFill && gFill !== '#FFFFFF') g.format.fill.color = gFill;
        if (gAlign && gAlign !== 'General') g.format.horizontalAlignment = gAlign;
        if (it.voyage) ws.getRange('H' + n).values = [[String(it.voyage)]];
        ws.getRange('K' + n).values = [[String(it.port).trim()]];
      }
      await ctx.sync();

      // Перечитываем записанное и сверяем.
      const checks = [];
      for (const r of pl.results) {
        if (r.kind !== 'new') continue;
        const rg = ws.getRange('A' + r.row + ':M' + r.row);
        rg.load('values');
        checks.push({ r, rg });
      }
      await ctx.sync();
      const problems = [];
      for (const { r, rg } of checks) {
        const v = rg.values[0], it = r.item;
        const ok = Number(v[0]) === serial && P.norm(v[2]) === P.norm(it.booking) && Number(v[3]) === Number(it.quantity) &&
                   P.norm(v[4]) === P.norm(it.type) && P.norm(v[6]) === 'PENDING' && P.norm(v[10]) === P.norm(it.port);
        if (!ok) problems.push(r.row);
      }
      if (problems.length) throw Error('Повторное чтение не совпало на строках: ' + problems.join(', ') + '. Проверьте их вручную.');
      report = labelPlan(pl);
    });

    el('result').className = 'ok';
    el('result').textContent = 'Записано и перечитано.\n' + report + '\n\nПроверьте в заголовке Excel, что стоит «Сохранено».';
    el('paste').value = ''; payload = null; shownPlanKey = '';
    el('recheck').disabled = true;
  } catch (e) {
    el('result').className = '';
    showErr(e.message);
  } finally {
    busy = false;
    el('write').disabled = !(payload && shownPlanKey);
  }
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) { showErr('Откройте надстройку внутри Excel.'); return; }
  el('env').textContent = 'Книга: ' + (docUrl() || '(адрес не получен)');
  const paste = el('paste');
  paste.addEventListener('input', () => onPaste());
  el('recheck').onclick = () => check().catch(e => showErr(e.message));
  el('write').onclick = () => write();
});
