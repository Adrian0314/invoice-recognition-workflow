/* ==========================================================================
   发票识别工作流 · 在线版
   在浏览器里用 Pyodide 运行仓库中的引擎本体：
     ../optimized_python/idle_invoice_bot.py  +  invoice_templates.json
   识别结论与本地 `python optimized_python/idle_invoice_bot.py -i ... -o ...` 一致。
   ========================================================================== */
'use strict';

const PYODIDE_VERSION = 'v0.28.3';
const PYODIDE_URL = `https://cdn.jsdelivr.net/pyodide/${PYODIDE_VERSION}/full/`;

/* 引擎与模板：直接从仓库里的源文件取，不做副本，永远与主实现一致 */
const ENGINE_URL = new URL('../optimized_python/idle_invoice_bot.py', location.href).href;
const TEMPLATE_URL = new URL('../optimized_python/invoice_templates.json', location.href).href;

/* 示例数据（与本地验证基线用的同一批样本，共 24 份）
   刻意写成显式路径数组而不是拼出来的：.github/scripts/assert_site.py 会逐个核对
   文件是否存在，改动 samples 目录时能立刻在 CI 上发现。 */
const DEMO_MANIFEST = [
  'samples/01_数电普票_技术服务费.pdf',
  'samples/02_数电专票_咨询服务费.pdf',
  'samples/03_电子普票_老版_办公用品.pdf',
  'samples/04_数电普票_与01同号_重复.pdf',
  'samples/05_数电普票_勾稽不平.pdf',
  'samples/06_数电专票_税号异常.pdf',
  'samples/07_一页两张发票.pdf',
  'samples/08_扫描件_无文本层.pdf',
  'samples/09_图片发票.png',
  'samples/10_同一文件的副本.pdf',
  'samples/11_同事又转发了一次.pdf',
  'samples_intl/T1-01_专业发票_软件开发服务.pdf',
  'samples_intl/T1-02_专业发票_设备采购.pdf',
  'samples_intl/T2-03_代理发票_代理服务费.pdf',
  'samples_intl/T2-04_代理发票_佣金结算.pdf',
  'samples_intl/T3-05_估算发票_项目预估价.pdf',
  'samples_intl/T3-06_估算发票_季度预算估算.pdf',
  'samples_intl/T4-07_条款发票_年度维保.pdf',
  'samples_intl/T4-08_条款发票_物流服务.pdf',
  'samples_intl/T5-09_简式发票_办公用品.pdf',
  'samples_intl/T5-10_简式发票_翻译服务.pdf',
  'samples_intl/T6-11_演讲者发票_主题演讲.pdf',
  'samples_intl/T6-12_演讲者发票_培训课程.pdf',
  'samples_mock/虚拟IT公司样票_10张.pdf',
];

/* ------------------------------------------------------------------ 状态 */
const S = {
  pyodide: null,
  ready: false,
  queue: [],          // {name, size, bytes:Uint8Array}
  lastResult: null,
  running: false,
  filter: null,       // 按处理状态筛选
};

const $ = id => document.getElementById(id);
const el = (tag, cls, txt) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (txt != null) n.textContent = txt;
  return n;
};

/* ------------------------------------------------------------------ 日志 */
function log(line) {
  const box = $('log');
  if (!box) return;
  box.textContent += (box.textContent ? '\n' : '') + line;
  box.scrollTop = box.scrollHeight;
}
function setBoot(title, detail, pct) {
  $('boot-title').textContent = title;
  if (detail != null) $('boot-detail').textContent = detail;
  $('boot-bar').style.width = `${pct == null ? 0 : pct}%`;
  if (pct >= 100) $('boot-card').style.display = 'none';
}

/* ================================================================== 启动 */
async function boot() {
  try {
    setBoot('正在下载运行时（Pyodide）…', '首次约 11.6 MB，之后走浏览器缓存', 8);
    if (typeof loadPyodide !== 'function') {
      throw new Error('未能加载 Pyodide（CDN 不可达？）请检查网络后刷新页面。');
    }
    const py = await loadPyodide({
      indexURL: PYODIDE_URL,
      stdout: t => log(t),
      stderr: t => log(t),
    });
    S.pyodide = py;

    setBoot('正在读取识别引擎…', '从仓库的 optimized_python/ 取出引擎与版式模板', 55);
    const [engineCode, tplText] = await Promise.all([
      fetch(ENGINE_URL).then(r => {
        if (!r.ok) throw new Error(`取引擎失败：HTTP ${r.status}（${ENGINE_URL}）`);
        return r.text();
      }),
      fetch(TEMPLATE_URL).then(r => (r.ok ? r.text() : '[]')),
    ]);

    setBoot('正在挂载引擎…', '把引擎写进浏览器内存文件系统并导入', 78);
    py.FS.mkdirTree('/app/optimized_python');
    py.FS.writeFile('/app/optimized_python/idle_invoice_bot.py', engineCode);
    py.FS.writeFile('/app/optimized_python/invoice_templates.json', tplText);
    py.FS.mkdirTree('/work/in');
    py.FS.mkdirTree('/work/out');

    py.runPython(DRIVER);
    const ver = py.runPython('bot.__version__');

    S.ready = true;
    setBoot('引擎就绪', `引擎版本 v${ver}（与仓库源文件同一份）`, 100);
    log(`引擎已加载：idle_invoice_bot v${ver}`);
    log(`引擎来源：${ENGINE_URL}`);
    ['btn-run', 'btn-demo'].forEach(id => { $(id).disabled = false; });
    refreshQueue();
  } catch (err) {
    setBoot('启动失败', String(err && err.message || err), 0);
    $('boot-bar').style.background = 'var(--err)';
    log('启动失败：' + (err && err.stack || err));
  }
}

/* ============================================================ Python 驱动 */
const DRIVER = String.raw`
import importlib.util as _ilu, io as _io, json as _json, sys as _sys, shutil as _sh, traceback as _tb
from pathlib import Path as _P

_ENG = _P("/app/optimized_python/idle_invoice_bot.py")
_spec = _ilu.spec_from_file_location("invoice_bot", str(_ENG))
bot = _ilu.module_from_spec(_spec)
_sys.modules["invoice_bot"] = bot
_spec.loader.exec_module(bot)

HINT_BROWSER = "（在线版无 OCR 引擎，此项需在本地运行识别；见 README）"


def _clear(d):
    p = _P(d)
    if p.exists():
        for q in p.rglob("*"):
            try:
                q.unlink() if q.is_file() else q.rmdir()
            except OSError:
                pass
    p.mkdir(parents=True, exist_ok=True)


def web_run():
    """处理 /work/in 下的文件，返回 JSON 字符串。"""
    cfg = dict(bot.DEFAULT_CONFIG)
    cfg["input_dir"] = "/work/in"
    cfg["output_dir"] = "/work/out"
    cfg["recursive"] = True
    cfg["ocr_enabled"] = "off"          # 浏览器内没有 OCR 推理引擎
    cfg["llm_fallback"] = {"enabled": False}
    _clear(cfg["output_dir"])

    logs = []
    old = _sys.stdout
    _sys.stdout = _io.StringIO()
    try:
        res = bot.run(cfg, log=logs.append, verbose=False)
        stray = _sys.stdout.getvalue()
        if stray.strip():
            logs.extend(stray.strip().splitlines())
    except BaseException as exc:
        _sys.stdout = old
        return _json.dumps({"ok": False,
                            "error": "%s: %s" % (type(exc).__name__, exc),
                            "trace": _tb.format_exc(), "log": logs},
                           ensure_ascii=False)
    finally:
        _sys.stdout = old

    out = _P(cfg["output_dir"])

    def rd(name):
        p = out / name
        return p.read_text(encoding="utf-8") if p.exists() else ""

    try:
        rows = _json.loads(rd(cfg["json_name"]) or "[]")
    except Exception:
        rows = []

    # 给"需要 OCR 但浏览器做不到"的行补一句人话解释
    for r in rows:
        note = str(r.get("异常/提示说明") or "")
        if ("OCR" in note) or ("Pillow" in note) or ("文本层" in note):
            if HINT_BROWSER not in note:
                r["异常/提示说明"] = (note + " " + HINT_BROWSER).strip()

    summ = res.get("summary") or {}
    return _json.dumps({
        "ok": not res.get("failed"),
        "reason": res.get("reason"),
        "version": bot.__version__,
        "stats": res.get("stats") or {},
        "elapsed": summ.get("elapsed"),
        "extractor_mix": summ.get("extractor_mix"),
        "files_count": summ.get("files"),
        "rows": rows,
        "files": {
            "csv": rd(cfg["csv_name"]),
            "issue_csv": rd(cfg["issue_csv_name"]),
            "review_csv": rd(cfg["review_csv_name"]),
            "json": rd(cfg["json_name"]),
            "report_html": rd(cfg["report_name"]),
            "audit_jsonl": rd(cfg["audit_name"]),
        },
        "log": logs,
    }, ensure_ascii=False)


web_run._pyodide_ok = True
`;

/* ================================================================== 队列 */
function refreshQueue() {
  const list = $('filelist');
  list.innerHTML = '';
  S.queue.forEach((f, i) => {
    const li = el('li');
    li.appendChild(el('span', null, f.name));
    const right = el('span');
    right.appendChild(el('span', 'size', fmtSize(f.size)));
    const rm = el('button', 'rm', '×');
    rm.title = '移除';
    rm.onclick = e => { e.stopPropagation(); S.queue.splice(i, 1); refreshQueue(); };
    right.appendChild(rm);
    li.appendChild(right);
    list.appendChild(li);
  });
  list.hidden = S.queue.length === 0;
  $('queue-info').textContent = S.queue.length ? `待识别 ${S.queue.length} 个文件` : '';
  $('btn-clear').disabled = S.queue.length === 0 || S.running;
  $('btn-run').disabled = !S.ready || S.queue.length === 0 || S.running;
  $('btn-demo').disabled = !S.ready || S.running;
}

function fmtSize(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1048576) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1048576).toFixed(2)} MB`;
}

function addFiles(items) {
  for (const it of items) {
    if (S.queue.some(q => q.name === it.name && q.size === it.size)) continue;
    S.queue.push(it);
  }
  refreshQueue();
}

async function handleFiles(fileList) {
  const allow = /\.(pdf|png|jpe?g|bmp|tiff?|webp)$/i;
  const items = [];
  for (const f of fileList) {
    if (!allow.test(f.name)) { log(`跳过不支持的文件：${f.name}`); continue; }
    items.push({ name: f.name, size: f.size, bytes: new Uint8Array(await f.arrayBuffer()) });
  }
  addFiles(items);
}

async function loadDemo() {
  $('btn-demo').disabled = true;
  log('正在从仓库读取示例样本…');
  const items = [];
  for (const rel of DEMO_MANIFEST) {
    const url = new URL('../' + rel, location.href).href;
    try {
      const r = await fetch(url);
      if (!r.ok) { log(`示例缺失（HTTP ${r.status}）：${rel}`); continue; }
      const b = new Uint8Array(await r.arrayBuffer());
      items.push({ name: rel.split('/').pop(), size: b.length, bytes: b });
    } catch (e) {
      log(`示例读取失败：${rel} —— ${e}`);
    }
  }
  addFiles(items);
  log(`已载入示例 ${items.length} 份`);
  $('btn-demo').disabled = false;
}

/* ================================================================== 运行 */
const STATUS_CLASS = {
  '正常': 'ok', '存在提示': 'warn', '异常': 'err',
  '待复核': 'info', '重复未入账': 'info', '待人工录入': '',
};

async function run() {
  if (!S.ready || !S.queue.length || S.running) return;
  S.running = true;
  refreshQueue();
  $('log-card').hidden = false;
  log('');
  log(`=== 开始识别 ${S.queue.length} 个文件 ===`);

  // 写入内存文件系统
  const py = S.pyodide;
  try {
    py.FS.mkdirTree('/work/in');
    for (const q of S.queue) py.FS.writeFile('/work/in/' + q.name, q.bytes);
  } catch (e) {
    log('写入文件失败：' + e);
    S.running = false; refreshQueue(); return;
  }

  const t0 = performance.now();
  let data;
  try {
    const raw = py.runPython('web_run()');
    data = JSON.parse(raw);
  } catch (e) {
    log('引擎执行异常：' + (e && e.stack || e));
    S.running = false; refreshQueue(); return;
  }
  const wall = (performance.now() - t0) / 1000;

  (data.log || []).forEach(log);

  if (!data.ok && data.error) {
    log('识别失败：' + data.error);
    if (data.trace) log(data.trace);
    S.running = false; refreshQueue(); return;
  }

  S.lastResult = data;
  renderResult(data, wall);
  S.running = false;
  refreshQueue();
  log(`=== 完成，用时 ${wall.toFixed(2)}s ===`);
}

/* ================================================================ 渲染 */
const STAT_DEFS = [
  ['台账行数', 'rows', '', null],
  ['正常', 'ok', 'ok', '正常'],
  ['存在提示', 'warn', 'warn', '存在提示'],
  ['异常', 'error', 'err', '异常'],
  ['待复核', 'review', 'info', '待复核'],
  ['重复拦截', 'duplicate', 'info', '重复未入账'],
  ['待人工录入', 'ocr_todo', 'warn', '待人工录入'],
  ['合计金额', 'sum_amount', 'money', null],
  ['合计税额', 'sum_tax', 'money', null],
  ['价税合计', 'sum_total', 'money', null],
];

function renderResult(data, wall) {
  const st = data.stats || {};
  S.filter = null;
  $('result-card').hidden = false;
  $('result-meta').textContent =
    `引擎 v${data.version}　·　${data.files_count} 个文件　·　${(data.elapsed ?? wall).toFixed(2)}s`
    + (data.extractor_mix ? `　·　提取器 ${data.extractor_mix}` : '');

  const box = $('stats');
  box.innerHTML = '';
  for (const [label, key, cls, status] of STAT_DEFS) {
    const isMoney = cls === 'money';
    const d = el('div', 'stat ' + (isMoney ? 'money' : cls) + (status ? ' click' : ''));
    if (status) d.title = `只看「${status}」的行`;
    d.appendChild(el('div', 'k', label));
    d.appendChild(el('div', 'v', isMoney ? money(st[key]) : String(st[key] ?? 0)));
    if (status) {
      d.dataset.status = status;
      d.addEventListener('click', () => {
        S.filter = (S.filter === status) ? null : status;
        renderTable();
      });
    }
    box.appendChild(d);
  }
  renderTable();
  $('result-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderTable() {
  const all = (S.lastResult && S.lastResult.rows) || [];
  const shown = S.filter ? all.filter(r => String(r['处理状态']) === S.filter) : all;

  document.querySelectorAll('.stat.click').forEach(d =>
    d.classList.toggle('active', d.dataset.status === S.filter));

  const tb = $('tbl').querySelector('tbody');
  tb.innerHTML = '';

  if (!shown.length) {
    const tr = el('tr', 'emptyrow');
    const td = el('td', null, S.filter ? `没有「${S.filter}」的行` : '没有数据');
    td.colSpan = 8;
    tr.appendChild(td);
    tb.appendChild(tr);
    return;
  }

  shown.forEach(r => {
    const idx = all.indexOf(r) + 1;
    const stt = String(r['处理状态'] ?? '');

    const tr = el('tr', 'main');
    tr.appendChild(el('td', null, String(idx)));
    tr.appendChild(el('td', null, String(r['源文件名'] ?? '')));
    const stTd = el('td');
    stTd.appendChild(el('span', 'pill ' + (STATUS_CLASS[stt] ?? ''), stt));
    tr.appendChild(stTd);
    tr.appendChild(el('td', null, dash(r['发票号码'])));
    tr.appendChild(el('td', null, dash(r['开票日期'])));
    tr.appendChild(el('td', null, dash(r['销售方名称'])));
    tr.appendChild(el('td', 'num', dash(r['价税合计'])));
    tr.appendChild(el('td', 'num', dash(r['识别置信度'])));
    tb.appendChild(tr);

    const sub = el('tr', 'sub');
    const td = el('td');
    td.colSpan = 8;
    const kv = (k, v) => {
      const s = el('span', 'kv');
      s.appendChild(el('span', null, k + ' '));
      s.appendChild(el('b', null, dash(v)));
      return s;
    };
    td.appendChild(kv('类型', r['发票类型']));
    td.appendChild(kv('金额', r['金额(不含税)']));
    td.appendChild(kv('税额', r['税额']));
    td.appendChild(kv('税率', r['税率/征收率']));
    td.appendChild(kv('购买方', r['购买方名称']));
    const note = dash(r['异常/提示说明']);
    if (note !== '—') td.appendChild(el('span', 'note', note));
    sub.appendChild(td);
    tb.appendChild(sub);
  });
}

const dash = v => (v === '' || v == null) ? '—' : String(v);
const money = v => (v == null) ? '—'
  : Number(v).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ================================================================ 导出 */
const EXPORT_META = {
  csv: ['发票明细.csv', 'text/csv;charset=utf-8'],
  issue_csv: ['异常记录.csv', 'text/csv;charset=utf-8'],
  review_csv: ['待复核.csv', 'text/csv;charset=utf-8'],
  json: ['发票明细.json', 'application/json;charset=utf-8'],
  audit_jsonl: ['审计轨迹.jsonl', 'application/jsonl;charset=utf-8'],
  report_html: ['处理报告.html', 'text/html;charset=utf-8'],
};

function exportFile(key) {
  const res = S.lastResult;
  if (!res) return;
  const content = (res.files || {})[key];
  if (!content) { alert('这一项没有内容（可能是本批次没有异常 / 待复核记录）'); return; }
  const [name, mime] = EXPORT_META[key];
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  const blob = new Blob([content], { type: mime });
  const a = el('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${stamp}_${name}`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

function printReport() {
  const res = S.lastResult;
  if (!res) return;
  const html = (res.files || {}).report_html;
  if (html) {
    const w = window.open('', '_blank');
    if (!w) { alert('浏览器拦截了新窗口，请允许弹出窗口后重试'); return; }
    w.document.write(html);
    w.document.close();
    return;
  }
  window.print();
}

/* ================================================================== 绑定 */
function bind() {
  const drop = $('drop');
  const input = $('file');

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => { handleFiles(input.files); input.value = ''; });

  ['dragenter', 'dragover'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => {
    if (e.dataTransfer && e.dataTransfer.files) handleFiles(e.dataTransfer.files);
  });
  // 整页拖放也接住
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => {
    e.preventDefault();
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
      handleFiles(e.dataTransfer.files);
    }
  });

  $('btn-run').addEventListener('click', run);
  $('btn-demo').addEventListener('click', loadDemo);
  $('btn-clear').addEventListener('click', () => {
    S.queue = []; S.lastResult = null; S.filter = null;
    $('result-card').hidden = true; $('log').textContent = '';
    $('log-card').hidden = true;
    refreshQueue();
  });
  $('btn-print').addEventListener('click', printReport);
  document.querySelectorAll('[data-export]').forEach(b =>
    b.addEventListener('click', () => exportFile(b.dataset.export)));
}

/* ============================================================== 自检模式 */
async function selftest() {
  const out = el('pre');
  out.id = 'selftest-out';
  out.style.display = 'none';
  document.body.appendChild(out);
  const put = o => { out.textContent = JSON.stringify(o); };
  try {
    await loadDemo();
    await run();
    const r = S.lastResult || {};
    const st = r.stats || {};
    put({
      ok: !!r.ok, version: r.version,
      rows: (r.rows || []).length,
      ok_rows: st.ok, warn: st.warn, error: st.error, review: st.review,
      duplicate: st.duplicate, ocr_todo: st.ocr_todo,
      sum_total: st.sum_total,
      extractor_mix: r.extractor_mix,
      files: Object.keys(r.files || {}),
      statuses: (r.rows || []).reduce((a, x) => {
        const k = String(x['处理状态']); a[k] = (a[k] || 0) + 1; return a;
      }, {}),
    });
  } catch (e) {
    put({ ok: false, error: String(e && e.stack || e) });
  }
  document.title = 'SELFTEST_DONE';
}

/* ================================================================== 入口 */
bind();
boot().then(() => {
  if (new URLSearchParams(location.search).get('selftest') === '1') selftest();
});
window.__APP = S;
