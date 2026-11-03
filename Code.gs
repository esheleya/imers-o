/**
 * IMERSÃO MÉTODO SER+I — Aplicativo de pré-inscrição.
 * Projeto associado a uma Planilha Google (Extensões > Apps Script).
 * Direitos e responsabilidades sobre os dados: organização do evento.
 *
 * V5: checkout único pela SumUp, no valor de R$ 347,00.
 * IMPORTANTE: este link leva à operadora; o app não processa cartões
 * nem comprova aprovação. Conferência manual no painel SumUp.
 */
const SERI = Object.freeze({
  ABA: 'Inscrições',
  ABA_PAGAMENTOS: 'Pagamentos Confirmados',
  VALOR: 347,
  LIMITE_DE_PARTICIPANTES: 40,
  LINK_SUMUP: 'https://pay.sumup.com/b2c/QI0DDZQA', // Link público informado pelo organizador (Pix ou cartão)
  ID_PLANILHA: ''   // Deixe vazio se abrir Apps Script a partir da Planilha.
});

const CABECALHOS = [
  'Protocolo', 'Enviado em', 'Nome completo', 'CPF', 'Nascimento',
  'WhatsApp', 'E-mail', 'Cidade', 'Profissão / atuação', 'Empresa',
  'Como conheceu', 'Forma de pagamento', 'Aceite do evento',
  'Ciência sobre privacidade', 'Valor (R$)', 'Situação do pagamento',
  'Situação da vaga', 'Observações'
];

const CABECALHOS_PAGAMENTOS = [
  'Protocolo', 'Inscrito em', 'Nome completo', 'WhatsApp', 'E-mail',
  'Cidade', 'Forma de pagamento', 'Valor (R$)', 'Pagamento',
  'Vaga', 'Observações'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Inscrição Oficial — Imersão Método SER+I')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Cria um menu para a organização atualizar a lista de pagamentos. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SER+I — Organização')
    .addItem('Atualizar pagamentos confirmados', 'atualizarPagamentosConfirmados')
    .addItem('Preparar planilhas', 'prepararPlanilha')
    .addToUi();
}

/** Execute uma vez no editor para criar a aba e conceder permissões. */
function prepararPlanilha() {
  const sheet = pegarAba_();
  formatarAbaInscricoes_(sheet);
  const pagamentos = pegarAbaPagamentos_();
  pagamentos.setFrozenRows(1);
  pagamentos.getRange(1, 1, 1, CABECALHOS_PAGAMENTOS.length)
    .setBackground('#0A1D37').setFontColor('#F3D28A')
    .setFontWeight('bold').setWrap(true);
  pagamentos.setColumnWidth(3, 260);
  pagamentos.setColumnWidth(5, 240);
  pagamentos.getRange('D:D').setNumberFormat('@');
  pagamentos.getRange('H:H').setNumberFormat('R$ #,##0.00');
  atualizarPagamentosConfirmados();
  return 'Planilhas de inscrições e pagamentos prontas.';
}

function formatarAbaInscricoes_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, CABECALHOS.length)
    .setBackground('#0A1D37').setFontColor('#F3D28A')
    .setFontWeight('bold').setWrap(true);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(7, 240);
  sheet.setColumnWidth(18, 260);
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('F:F').setNumberFormat('@');
  sheet.getRange('O:O').setNumberFormat('R$ #,##0.00');
}

/** Recria o resumo com quem está marcado como pagamento aprovado. */
function atualizarPagamentosConfirmados() {
  const origem = pegarAba_();
  const destino = pegarAbaPagamentos_();
  const ultimaLinha = origem.getLastRow();
  const dados = ultimaLinha > 1
    ? origem.getRange(2, 1, ultimaLinha - 1, CABECALHOS.length).getValues()
    : [];
  const aprovados = dados
    .filter(row => normalizar_(row[15]) === 'aprovado')
    .map(row => [row[0], row[1], row[2], row[5], row[6], row[7], row[11], row[14], row[15], row[16], row[17]]);

  const linhasAntigas = destino.getLastRow();
  if (linhasAntigas > 1) {
    destino.getRange(2, 1, linhasAntigas - 1, CABECALHOS_PAGAMENTOS.length).clearContent();
  }
  if (aprovados.length) {
    destino.getRange(2, 1, aprovados.length, CABECALHOS_PAGAMENTOS.length).setValues(aprovados);
  }
  destino.getRange('H:H').setNumberFormat('R$ #,##0.00');
  return aprovados.length + ' pagamento(s) confirmado(s).';
}

/** Atualiza o resumo quando a organização altera o status de pagamento. */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SERI.ABA) return;
  const primeiraColuna = e.range.getColumn();
  const ultimaColuna = primeiraColuna + e.range.getNumColumns() - 1;
  if (primeiraColuna <= 16 && ultimaColuna >= 16) atualizarPagamentosConfirmados();
}

function submitRegistration(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Dados inválidos.');
  if (String(payload.website || '').trim()) throw new Error('Envio inválido.');
  const p = validar_(payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = pegarAba_();
    const last = sheet.getLastRow();
    if (last > 1) {
      // Colunas D, P e Q: CPF, situação do pagamento e situação da vaga.
      const rows = sheet.getRange(2, 1, last - 1, CABECALHOS.length).getDisplayValues();
      const confirmados = rows.filter(row =>
        normalizar_(row[15]) === 'aprovado' ||
        normalizar_(row[16]) === 'confirmada'
      ).length;
      if (confirmados >= SERI.LIMITE_DE_PARTICIPANTES) {
        throw new Error('As vagas já foram preenchidas. Consulte a organização.');
      }
      const existente = rows.find(row => digitos_(row[3]) === p.cpf);
      if (existente) {
        // Resposta genérica para não divulgar a situação de dados pessoais.
        return respostaSucesso_(null);
      }
    }

    const protocolo = 'SERI-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([
      protocolo, new Date(), celula_(p.nome), "'" + p.cpf,
      p.nascimento, "'" + p.whatsapp, celula_(p.email), celula_(p.cidade),
      celula_(p.profissao), celula_(p.empresa), celula_(p.origem), p.pagamento,
      'Sim', 'Sim', SERI.VALOR, 'Pendente', 'Pré-inscrição', ''
    ]);
    SpreadsheetApp.flush();
    return respostaSucesso_(protocolo);
  } finally {
    lock.releaseLock();
  }
}

function respostaSucesso_(protocolo) {
  return {
    ok: true,
    protocolo: protocolo,
    message: 'Sua pré-inscrição foi registrada. A vaga será confirmada após a aprovação do pagamento.',
    paymentLink: linkSeguro_(SERI.LINK_SUMUP)
  };
}

function pegarAba_() {
  const ss = SERI.ID_PLANILHA
    ? SpreadsheetApp.openById(SERI.ID_PLANILHA)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('A planilha não foi encontrada. Abra o projeto em Extensões > Apps Script dentro de uma Planilha Google ou defina ID_PLANILHA.');
  let sheet = ss.getSheetByName(SERI.ABA);
  if (!sheet) sheet = ss.insertSheet(SERI.ABA);
  if (sheet.getLastRow() === 0) sheet.appendRow(CABECALHOS);
  return sheet;
}

function pegarAbaPagamentos_() {
  const ss = SERI.ID_PLANILHA
    ? SpreadsheetApp.openById(SERI.ID_PLANILHA)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('A planilha não foi encontrada.');
  let sheet = ss.getSheetByName(SERI.ABA_PAGAMENTOS);
  if (!sheet) sheet = ss.insertSheet(SERI.ABA_PAGAMENTOS);
  if (sheet.getLastRow() === 0) sheet.appendRow(CABECALHOS_PAGAMENTOS);
  return sheet;
}

function validar_(p) {
  const nome = texto_(p.nome, 120);
  const cpf = digitos_(p.cpf);
  const nascimento = texto_(p.nascimento, 10);
  const whatsapp = digitos_(p.whatsapp);
  const email = texto_(p.email, 160).toLowerCase();
  const cidade = texto_(p.cidade, 120);
  const profissao = texto_(p.profissao, 160);
  const empresa = texto_(p.empresa, 160);
  const origem = texto_(p.origem, 160);
  const pagamento = texto_(p.pagamento, 30);
  const aceite = p.aceiteEvento === true;
  const privacidade = p.cienciaPrivacidade === true;
  if (nome.split(/\s+/).filter(Boolean).length < 2) throw new Error('Informe seu nome completo.');
  if (!cpfValido_(cpf)) throw new Error('CPF inválido. Verifique os números.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(nascimento)) throw new Error('Informe a data de nascimento.');
  const d = new Date(nascimento + 'T12:00:00Z');
  if (Number.isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== nascimento || d > new Date() || d.getUTCFullYear() < 1900) throw new Error('Data de nascimento inválida.');
  if (![10,11].includes(whatsapp.length)) throw new Error('Informe um WhatsApp com DDD.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Informe um e-mail válido.');
  if (!cidade || !profissao || !origem) throw new Error('Preencha todos os campos obrigatórios.');
  if (!['Pix', 'Cartão de crédito'].includes(pagamento)) throw new Error('Escolha Pix ou cartão de crédito.');
  if (!aceite || !privacidade) throw new Error('Leia e aceite as informações do evento e o aviso de privacidade.');
  return { nome, cpf, nascimento, whatsapp, email, cidade, profissao, empresa, origem, pagamento };
}

function texto_(v, max) {
  return String(v || '').replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
}
function digitos_(v) { return String(v || '').replace(/\D/g, ''); }
function normalizar_(v) { return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase(); }
function celula_(v) {
  // Evita que dados digitados se tornem fórmulas na Planilha.
  return /^[=+@\-]/.test(v) ? "'" + v : v;
}
function linkSeguro_(url) {
  const u = String(url || '').trim();
  return /^https:\/\/pay\.sumup\.com\/b2c\/[a-z\d]+(?:[?#].*)?$/i.test(u) ? u : '';
}
function cpfValido_(cpf) {
  if (!/^\d{11}$/.test(cpf) || /^(\d)\1{10}$/.test(cpf)) return false;
  for (let tamanho = 9; tamanho < 11; tamanho++) {
    let soma = 0;
    for (let i = 0; i < tamanho; i++) soma += Number(cpf[i]) * (tamanho + 1 - i);
    const digito = (soma * 10) % 11;
    if ((digito === 10 ? 0 : digito) !== Number(cpf[tamanho])) return false;
  }
  return true;
}
