const ABA = 'Base';
const LIMITE = 400; // acima disso (dias de atraso) = Prejuízo
const META = ['Data de entrada', 'Entrou como', 'Classificação', 'Última atualização', 'Saiu de base em'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Controle de Fichas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Usa a planilha onde o script foi criado; se o script for avulso, cria e guarda uma planilha própria.
function planilha_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('PLANILHA');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create('Controle de Fichas - Base');
  p.setProperty('PLANILHA', ss.getId());
  return ss;
}

function aba_() {
  const ss = planilha_();
  return ss.getSheetByName(ABA) || ss.insertSheet(ABA);
}

function urlPlanilha() {
  return planilha_().getUrl();
}

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
}

function num_(v) {
  let s = String(v == null ? '' : v).replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// ===== BASE DE CONTRATOS =====
// Guarda TODOS os contratos (1 linha por contrato). Contratos já existentes são atualizados,
// novos são adicionados e, se marcarSaidas = true, os que sumiram do arquivo ganham "Saiu de base em".
function importar(dataIso, linhas, marcarSaidas) {
  if (!dataIso || !linhas || linhas.length < 2) throw new Error('Informe a data e a base com cabeçalho.');
  const cab = linhas[0].map(String);
  const iCon = cab.findIndex(c => norm_(c).indexOf('CONTRATO') === 0);
  if (iCon < 0) throw new Error('Coluna "Contrato" não encontrada no cabeçalho.');

  const sh = aba_();
  if (sh.getLastRow() === 0) {
    const h = META.concat(cab);
    if (h.length > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), h.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#e8eaf6');
    sh.setFrozenRows(1);
  }
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const orig = hdr.slice(META.length);
  const mapa = orig.map(h => cab.findIndex(c => norm_(c) === norm_(h)));
  const jCon = orig.findIndex(h => norm_(h).indexOf('CONTRATO') === 0);
  const jSit = orig.findIndex(h => norm_(h).indexOf('SITUA') === 0);
  const jDia = orig.findIndex(h => norm_(h).indexOf('DIAS') === 0);
  if (jCon < 0) throw new Error('A aba Base não tem coluna de contrato.');

  const n = sh.getLastRow() - 1;
  const dados = n > 0 ? sh.getRange(2, 1, n, hdr.length).getValues() : [];
  const pos = new Map();
  dados.forEach((r, i) => pos.set(String(r[META.length + jCon]).trim(), i));

  const vistos = new Set(), adds = [];
  let atualizados = 0, fichas = 0, voltaram = 0, saiu = 0;
  linhas.slice(1).forEach(r => {
    const con = String(r[iCon] == null ? '' : r[iCon]).trim();
    if (!con || vistos.has(con)) return;
    vistos.add(con);
    const o = mapa.map(j => (j < 0 || r[j] == null) ? '' : String(r[j]));
    const cls = jDia >= 0 && num_(o[jDia]) > LIMITE ? 'Prejuízo' : 'Ativo';
    if (pos.has(con)) {
      const l = dados[pos.get(con)];
      if (String(l[4])) voltaram++;
      o.forEach((v, k) => { l[META.length + k] = v; });
      l[2] = cls; l[3] = dataIso; l[4] = '';
      atualizados++;
    } else {
      const ficha = jSit >= 0 && norm_(o[jSit]).indexOf('FICHA NOVA') >= 0;
      if (ficha) fichas++;
      adds.push([dataIso, ficha ? 'Ficha nova' : 'Base', cls, dataIso, ''].concat(o));
    }
  });

  if (marcarSaidas) {
    dados.forEach(l => {
      const c = String(l[META.length + jCon]).trim();
      if (!vistos.has(c) && !String(l[4])) { l[4] = dataIso; saiu++; }
    });
  }
  if (dados.length) sh.getRange(2, 1, dados.length, hdr.length).setValues(dados);
  if (adds.length) {
    const ini = sh.getLastRow() + 1, falta = ini + adds.length - 1 - sh.getMaxRows();
    if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta);
    const rg = sh.getRange(ini, 1, adds.length, hdr.length);
    rg.setNumberFormat('@');
    rg.setValues(adds);
  }
  return { novos: adds.length, fichas: fichas, atualizados: atualizados, voltaram: voltaram, saiu: saiu, total: vistos.size };
}

function listar() {
  const sh = aba_();
  if (sh.getLastRow() < 2) return { cab: [], linhas: [] };
  const tz = Session.getScriptTimeZone();
  const d = sh.getDataRange().getValues();
  return {
    cab: d[0].map(String),
    linhas: d.slice(1).map(r => r.map(c => c instanceof Date ? Utilities.formatDate(c, tz, 'yyyy-MM-dd') : String(c)))
  };
}

// ===== ARQUIVOS (Google Drive) =====
function pasta_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('PASTA');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Controle de Fichas - Arquivos');
  p.setProperty('PASTA', f.getId());
  return f;
}

function meta_(f) {
  return {
    id: f.getId(), nome: f.getName(), tipo: f.getMimeType(), tam: f.getSize(),
    data: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    url: f.getUrl()
  };
}

function enviarArquivo(nome, mime, b64) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', nome);
  return meta_(pasta_().createFile(blob));
}

function listarArquivos() {
  const it = pasta_().getFiles(), o = [];
  while (it.hasNext()) o.push(meta_(it.next()));
  return o.sort((a, b) => b.data.localeCompare(a.data));
}

function obterImagem(id) {
  const b = DriveApp.getFileById(id).getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

function excluirArquivo(id) {
  DriveApp.getFileById(id).setTrashed(true);
  return true;
}

// ===== BE (BI) =====
const ABA_BI = 'BI';
// colunas do BI que o painel usa (sem acento/underscore/espaço, em maiúsculas); CPF/CNPJ entram se existirem
const BI_USO = ['COOPERATIVA','DSPESSOA','NRCC','NRCONTRATOCYBER','FLGPREJUIZO','DSGARANTIA','ACORDO','CNTRSTATUS','FASEDECOBRANCA',
  'PAGAMENTOTOTAL','RECUPERACAOEFETIVA','VLCONTRATADO','VLPARCELA','QTDIASATRASO','DTULTVNCTONAOPAGO','FAIXAOPERACAO',
  'MAIORFAIXAOPERACAODACONTA','CARTEIRAATRASOATUAL','CARTEIRADEVEDORATUAL','CARTEIRAPREJUIZO','ASSESSORIA','ASSESSORIALEGAL',
  'RISCOATUAL','FLGJUDCLIENTE','NRACORDO','DTACORDO','VLACORDO','QTPARCELASACORDO','QTPARCELASPAGASACORDO','SEGMENTOCARTEIRA'];

function chave_(s) { return norm_(s).replace(/[^A-Z0-9]/g, ''); }

// Recebe o BI em partes. Parte 0 limpa a aba BI; as demais acrescentam. NÃO remove duplicidades por contrato.
function importarBI(parte, cab, linhas, dataIso) {
  const ss = planilha_();
  const sh = ss.getSheetByName(ABA_BI) || ss.insertSheet(ABA_BI);
  const h = ['Importado em'].concat(cab.map(String));
  if (parte === 0) {
    sh.clear();
    if (h.length > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), h.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#e8eaf6');
    sh.setFrozenRows(1);
  }
  if (!linhas || !linhas.length) return 0;
  const dados = linhas.map(r => [dataIso].concat(h.slice(1).map((_, i) => r[i] == null ? '' : String(r[i]))));
  const ini = sh.getLastRow() + 1, falta = ini + dados.length - 1 - sh.getMaxRows();
  if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta);
  const rg = sh.getRange(ini, 1, dados.length, h.length);
  rg.setNumberFormat('@');
  rg.setValues(dados);
  return dados.length;
}

// Devolve só as colunas que o painel usa (o BI inteiro fica guardado na aba BI).
function listarBI() {
  const sh = planilha_().getSheetByName(ABA_BI);
  if (!sh || sh.getLastRow() < 2) return { cab: [], linhas: [], importadoEm: '' };
  const d = sh.getDataRange().getValues();
  const h = d[0].map(String), uso = new Set(BI_USO), idx = [];
  h.forEach((c, i) => {
    if (i === 0) return;
    const k = chave_(c);
    if (uso.has(k) || k.indexOf('CPF') === 0 || k.indexOf('CNPJ') === 0) idx.push(i);
  });
  return { cab: idx.map(i => h[i]), linhas: d.slice(1).map(r => idx.map(i => String(r[i]))), importadoEm: String(d[1][0]) };
}
