// Servidor do inventário: guarda contagem em data.json, vários PCs podem bipar ao mesmo tempo.
const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
const ExcelJS = require('exceljs');
const QRCode = require('qrcode');
const PDFDocument = require('pdfkit');

const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data.json');

// ---------- "banco" em arquivo ----------
let db = { itens: {}, log: [] };
try { db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8')); } catch (e) {}
function salvar() {
  fs.writeFileSync(DB_FILE + '.tmp', JSON.stringify(db));
  fs.renameSync(DB_FILE + '.tmp', DB_FILE);
}
const norm = s => String(s || '').trim().toUpperCase();
function achar(codigo) {
  const k = norm(codigo);
  return Object.values(db.itens).find(i => norm(i.codigo) === k);
}

// ---------- CSV do Protheus ----------
function num(v) {
  v = String(v || '0').trim();
  if (v.includes(',')) v = v.replace(/\./g, '').replace(',', '.'); // 1.234,56 -> 1234.56
  return parseFloat(v) || 0;
}
function splitLinha(l, sep) {
  const out = []; let cur = '', aspas = false;
  for (const ch of l) {
    if (ch === '"') aspas = !aspas;
    else if (ch === sep && !aspas) { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out.map(s => s.trim());
}
function parseCSV(txt) {
  const linhas = txt.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim());
  if (!linhas.length) return [];
  const sep = linhas[0].includes(';') ? ';' : ',';
  const cab = splitLinha(linhas[0], sep).map(s => s.toLowerCase());
  const achaCol = nomes => cab.findIndex(c => nomes.includes(c));
  const iC = achaCol(['codigo', 'código', 'b1_cod', 'cod']);
  const iD = achaCol(['descricao', 'descrição', 'b1_desc', 'desc']);
  const iS = achaCol(['saldo', 'b2_qatu', 'qtd']);
  if (iC < 0) throw new Error('Coluna "codigo" (ou B1_COD) não encontrada no cabeçalho.');
  return linhas.slice(1).map(l => {
    const p = splitLinha(l, sep);
    return { codigo: p[iC], descricao: iD >= 0 ? p[iD] : '', saldo: iS >= 0 ? num(p[iS]) : 0 };
  }).filter(i => i.codigo);
}

// ---------- API ----------
const app = express();
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/estado', (req, res) => {
  const itens = Object.values(db.itens);
  const contados = itens.filter(i => i.contado !== null).length;
  const diverg = itens.filter(i => i.contado !== null && i.contado !== i.saldo).length;
  res.json({ itens, resumo: { total: itens.length, contados, diverg }, ultimo: db.log[db.log.length - 1] || null });
});

app.post('/api/importar', (req, res) => {
  try {
    const lista = parseCSV(req.body.csv || '');
    const novo = {};
    lista.forEach(i => {
      const antigo = db.itens[i.codigo];
      novo[i.codigo] = { ...i, contado: antigo ? antigo.contado : null };
    });
    db.itens = novo; salvar();
    res.json({ importados: lista.length });
  } catch (e) { res.status(400).json({ erro: e.message }); }
});

function registrar(req, res, modo) {
  const it = achar(req.body.codigo);
  if (!it) return res.status(404).json({ erro: 'Código não encontrado: ' + req.body.codigo });
  const q = modo === 'somar' ? Number(req.body.qtd ?? 1) : Number(req.body.qtd);
  if (isNaN(q) || q < 0) return res.status(400).json({ erro: 'Quantidade inválida.' });
  const antes = it.contado;
  it.contado = modo === 'somar' ? (it.contado || 0) + q : q;
  db.log.push({ codigo: it.codigo, antes, depois: it.contado, modo, por: req.body.contador || '', quando: new Date().toISOString() });
  salvar();
  res.json({ item: it });
}
app.post('/api/bipar', (req, res) => registrar(req, res, 'somar'));
app.post('/api/definir', (req, res) => registrar(req, res, 'definir'));

app.post('/api/desfazer', (req, res) => {
  const ult = db.log.pop();
  if (!ult) return res.status(400).json({ erro: 'Nada para desfazer.' });
  const it = db.itens[ult.codigo] || achar(ult.codigo);
  if (it) it.contado = ult.antes;
  salvar();
  res.json({ desfeito: ult });
});

app.post('/api/zerar', (req, res) => {
  Object.values(db.itens).forEach(i => { i.contado = null; });
  db.log = []; salvar();
  res.json({ ok: true });
});

app.get('/api/exportar.xlsx', async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Inventario');
  ws.columns = [
    { header: 'Código', key: 'codigo', width: 18 },
    { header: 'Descrição', key: 'descricao', width: 42 },
    { header: 'Saldo sistema', key: 'saldo', width: 15 },
    { header: 'Contado', key: 'contado', width: 12 },
    { header: 'Diferença', key: 'dif', width: 12 },
    { header: 'Status', key: 'status', width: 16 },
  ];
  ws.getRow(1).font = { bold: true };
  Object.values(db.itens).forEach(i => {
    const contado = i.contado !== null;
    const dif = contado ? i.contado - i.saldo : null;
    const row = ws.addRow({
      codigo: i.codigo, descricao: i.descricao, saldo: i.saldo,
      contado: contado ? i.contado : '', dif: contado ? dif : '',
      status: !contado ? 'Não contado' : dif === 0 ? 'OK' : 'Divergente',
    });
    const cor = !contado ? 'FFE5E5E5' : dif === 0 ? null : 'FFF8C9C4';
    if (cor) row.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor } }; });
  });
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const wl = wb.addWorksheet('Log');
  wl.columns = [
    { header: 'Quando', key: 'q', width: 22 }, { header: 'Contador', key: 'p', width: 18 },
    { header: 'Código', key: 'c', width: 18 }, { header: 'Operação', key: 'm', width: 12 },
    { header: 'Antes', key: 'a', width: 10 }, { header: 'Depois', key: 'd', width: 10 },
  ];
  wl.getRow(1).font = { bold: true };
  db.log.forEach(l => wl.addRow({ q: new Date(l.quando).toLocaleString('pt-BR'), p: l.por, c: l.codigo, m: l.modo, a: l.antes ?? '', d: l.depois }));

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="inventario_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
});

app.get('/api/etiquetas.pdf', async (req, res) => {
  const itens = Object.values(db.itens);
  if (!itens.length) return res.status(400).send('Importe os itens antes de gerar etiquetas.');
  const doc = new PDFDocument({ size: 'A4', margin: 0 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="etiquetas.pdf"');
  doc.pipe(res);
  const COLS = 3, LINHAS = 8, W = doc.page.width / COLS, H = doc.page.height / LINHAS;
  for (let n = 0; n < itens.length; n++) {
    const pos = n % (COLS * LINHAS);
    if (n && pos === 0) doc.addPage();
    const x = (pos % COLS) * W, y = Math.floor(pos / COLS) * H;
    const lado = Math.min(H - 16, W * 0.45);
    const png = await QRCode.toBuffer(itens[n].codigo, { margin: 1, width: 300 });
    doc.rect(x + 3, y + 3, W - 6, H - 6).lineWidth(0.3).stroke('#999');
    doc.image(png, x + 8, y + (H - lado) / 2, { width: lado });
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10).text(itens[n].codigo, x + lado + 14, y + H / 2 - 12, { width: W - lado - 20 });
    doc.font('Helvetica').fontSize(7).text(itens[n].descricao || '', x + lado + 14, y + H / 2 + 2, { width: W - lado - 20, height: 24 });
  }
  doc.end();
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Inventário rodando em http://localhost:${PORT}`);
  Object.values(os.networkInterfaces()).flat().filter(n => n.family === 'IPv4' && !n.internal)
    .forEach(n => console.log(`Na rede da empresa: http://${n.address}:${PORT}`));
});