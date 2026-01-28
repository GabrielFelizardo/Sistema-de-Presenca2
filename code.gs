// ========================================
// 🇨🇭 SISTEMA RSVP v17.1 - BACKEND SAAS
// ========================================

const CONFIG = {
  NOME_PLANILHA_PREFIXO: 'Meus Eventos',
  NOME_ABA_BANCO: 'Banco_Nomes',
  VERSAO: '17.1'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gestor RSVP v17.1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- SESSÃO & LOGIN ---

function obterSessao() {
  const props = PropertiesService.getUserProperties();
  const email = props.getProperty('email_usuario');
  const planilhaId = props.getProperty('planilha_id');
  
  if (email && planilhaId) {
    return { autenticado: true, email: email, planilhaId: planilhaId };
  }
  return { autenticado: false };
}

function fazerLogin(email) {
  if (!email || !email.includes('@')) throw new Error("E-mail inválido");
  
  const emailLimpo = email.trim().toLowerCase();
  
  // 1. Busca ou Cria a Planilha do Usuário
  const planilhaId = obterOuCriarPlanilha(emailLimpo);
  
  // 2. Salva Sessão no Cache do Script (UserProperties)
  const props = PropertiesService.getUserProperties();
  props.setProperty('email_usuario', emailLimpo);
  props.setProperty('planilha_id', planilhaId);
  
  return { sucesso: true, email: emailLimpo };
}

function sair() {
  PropertiesService.getUserProperties().deleteAllProperties();
  return { sucesso: true };
}

// --- GERENCIADOR DE PLANILHAS ---

function obterOuCriarPlanilha(email) {
  // Tenta achar planilha pelo nome no Drive
  const nomeArquivo = `${CONFIG.NOME_PLANILHA_PREFIXO} - ${email}`;
  const arquivos = DriveApp.getFilesByName(nomeArquivo);
  
  if (arquivos.hasNext()) {
    return arquivos.next().getId();
  }
  
  // Se não existe, cria nova
  const novaSS = SpreadsheetApp.create(nomeArquivo);
  const id = novaSS.getId();
  
  // Configura aba inicial
  const aba = novaSS.getSheets()[0];
  aba.setName("Exemplo");
  aba.appendRow(["Nome", "Telefone", "Email", "Presença"]);
  aba.getRange("A1:D1").setFontWeight("bold");
  
  return id;
}

// --- FUNÇÕES DO SISTEMA (Eventos, Convidados, etc) ---

function listarEventos() {
  const sessao = obterSessao();
  if (!sessao.autenticado) throw new Error("Não autenticado");
  
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const abas = ss.getSheets();
  const eventos = [];
  
  abas.forEach(aba => {
    const nome = aba.getName();
    if (nome === "Exemplo" || nome === CONFIG.NOME_ABA_BANCO) return;
    
    // Estatísticas rápidas
    const dados = aba.getDataRange().getValues();
    let total = Math.max(0, dados.length - 1);
    let confirmados = 0;
    
    if (total > 0) {
      const idxPresenca = dados[0].indexOf("Presença");
      if (idxPresenca > -1) {
        for (let i = 1; i < dados.length; i++) {
          if (dados[i][idxPresenca] === "Confirmado") confirmados++;
        }
      }
    }
    
    eventos.push({ nome: nome, total: total, confirmados: confirmados });
  });
  
  return { sucesso: true, eventos: eventos };
}

function criarEvento(nome, colunas) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  
  if (ss.getSheetByName(nome)) throw new Error("Evento já existe");
  
  const aba = ss.insertSheet(nome);
  const headers = ["Nome", ...colunas, "Presença"]; // Garante colunas base
  
  aba.appendRow(headers);
  aba.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
  aba.setFrozenRows(1);
  
  return { sucesso: true };
}

function obterConvidados(nomeEvento) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = ss.getSheetByName(nomeEvento);
  
  const dados = aba.getDataRange().getValues();
  if (dados.length < 1) return { sucesso: true, colunas: [], convidados: [] };
  
  const headers = dados[0];
  const convidados = [];
  
  for (let i = 1; i < dados.length; i++) {
    let conv = { _linha: i + 1 };
    headers.forEach((h, idx) => conv[h] = dados[i][idx]);
    convidados.push(conv);
  }
  
  return { sucesso: true, colunas: headers, convidados: convidados };
}

function adicionarConvidado(nomeEvento, dados) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = ss.getSheetByName(nomeEvento);
  
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const novaLinha = headers.map(h => dados[h] || "");
  
  aba.appendRow(novaLinha);
  return { sucesso: true };
}

function togglePresenca(nomeEvento, linha) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = ss.getSheetByName(nomeEvento);
  
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const idx = headers.indexOf("Presença");
  if (idx === -1) throw new Error("Coluna Presença não encontrada");
  
  const celula = aba.getRange(linha, idx + 1);
  const val = celula.getValue();
  const novoVal = val === "Confirmado" ? "" : "Confirmado";
  
  celula.setValue(novoVal);
  return { sucesso: true, novoStatus: novoVal };
}

function excluirConvidado(nomeEvento, linha) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = ss.getSheetByName(nomeEvento);
  aba.deleteRow(linha);
  return { sucesso: true };
}

function excluirEvento(nomeEvento) {
  const sessao = obterSessao();
  const ss = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = ss.getSheetByName(nomeEvento);
  ss.deleteSheet(aba);
  return { sucesso: true };
}
