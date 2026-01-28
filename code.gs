// ========================================
// 🇨🇭 SISTEMA RSVP v17.0 - SWISS DESIGN
// ========================================

const CONFIG = {
  NOME_PLANILHA_PREFIXO: 'Meus Eventos',
  NOME_ABA_BANCO: 'Banco_Nomes',
  VERSAO: '17.0'
};

// ========================================
// 🔐 AUTENTICAÇÃO E INICIALIZAÇÃO
// ========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema RSVP v17.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obterSessao() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { autenticado: false };
    
    // Tenta encontrar a planilha deste usuário
    const planilhaId = obterOuCriarPlanilhaUsuario(email);
    
    return {
      autenticado: true,
      email: email,
      planilhaId: planilhaId
    };
  } catch (erro) {
    Logger.log(`❌ Erro ao obter sessão: ${erro.message}`);
    return { autenticado: false };
  }
}

function obterOuCriarPlanilhaUsuario(email) {
  const props = PropertiesService.getUserProperties();
  let planilhaId = props.getProperty('planilha_id_' + email);
  
  // Verifica se a planilha ainda existe
  if (planilhaId) {
    try {
      SpreadsheetApp.openById(planilhaId);
      return planilhaId;
    } catch (e) {
      Logger.log("Planilha antiga não encontrada, criando nova...");
    }
  }
  
  // Tenta buscar no Drive pelo nome
  const nomePlanilha = `${CONFIG.NOME_PLANILHA_PREFIXO} - ${email}`;
  const arquivos = DriveApp.getFilesByName(nomePlanilha);
  
  if (arquivos.hasNext()) {
    planilhaId = arquivos.next().getId();
  } else {
    // Cria nova planilha
    const novaPlanilha = SpreadsheetApp.create(nomePlanilha);
    planilhaId = novaPlanilha.getId();
    configurarPlanilhaInicial(novaPlanilha);
  }
  
  props.setProperty('planilha_id_' + email, planilhaId);
  return planilhaId;
}

function configurarPlanilhaInicial(planilha) {
  // Remove aba padrão
  const abas = planilha.getSheets();
  if (abas.length > 0 && abas[0].getName().includes("Página")) {
     abas[0].setName("Exemplo");
  }
  
  // Cria Banco de Nomes se não existir
  if (!planilha.getSheetByName(CONFIG.NOME_ABA_BANCO)) {
    const abaBanco = planilha.insertSheet(CONFIG.NOME_ABA_BANCO);
    abaBanco.getRange('A1:C1').setValues([['Nome', 'Email', 'Telefone']]);
    abaBanco.getRange('A1:C1').setFontWeight('bold');
    abaBanco.setFrozenRows(1);
  }
}

// ========================================
// 📊 GERENCIAMENTO DE EVENTOS
// ========================================

function listarEventos() {
  try {
    const sessao = obterSessao();
    if (!sessao.autenticado) throw new Error('Usuário não autenticado');
    
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const abas = planilha.getSheets();
    const eventos = [];

    for (let aba of abas) {
      const nomeAba = aba.getName();
      if (nomeAba === CONFIG.NOME_ABA_BANCO || nomeAba === "Exemplo") continue;
      
      const dados = aba.getDataRange().getValues();
      const totalConvidados = Math.max(0, dados.length - 1);
      
      let confirmados = 0;
      if (dados.length > 1) {
        const headers = dados[0];
        const colPresenca = headers.indexOf('Presença');
        if (colPresenca !== -1) {
          for (let i = 1; i < dados.length; i++) {
            if (dados[i][colPresenca] === 'Confirmado') confirmados++;
          }
        }
      }
      
      eventos.push({
        nome: nomeAba,
        totalConvidados: totalConvidados,
        confirmados: confirmados,
        pendentes: totalConvidados - confirmados
      });
    }
    
    return { sucesso: true, eventos: eventos };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function criarEvento(nomeEvento, colunas) {
  try {
    const sessao = obterSessao();
    if (!sessao.autenticado) throw new Error('Sessão expirada');
    
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    if (planilha.getSheetByName(nomeEvento)) throw new Error('Já existe um evento com este nome');
    
    const novaAba = planilha.insertSheet(nomeEvento);
    // Adiciona colunas obrigatórias
    const colunasFinais = ['Nome', ...colunas, 'Presença'];
    
    novaAba.getRange(1, 1, 1, colunasFinais.length).setValues([colunasFinais]);
    novaAba.getRange(1, 1, 1, colunasFinais.length).setFontWeight('bold').setBackground('#f3f4f6');
    novaAba.setFrozenRows(1);
    
    return { sucesso: true, mensagem: 'Evento criado com sucesso' };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function excluirEvento(nomeEvento) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    if (aba) {
      planilha.deleteSheet(aba);
      return { sucesso: true };
    }
    throw new Error('Evento não encontrado');
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

// ========================================
// 👥 GERENCIAMENTO DE CONVIDADOS
// ========================================

function obterConvidados(nomeEvento) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    
    if (!aba) throw new Error('Evento não encontrado');
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: true, convidados: [], colunas: dados.length > 0 ? dados[0] : [] };
    
    const headers = dados[0];
    const convidados = [];
    
    for (let i = 1; i < dados.length; i++) {
      const convidado = {};
      for (let j = 0; j < headers.length; j++) {
        convidado[headers[j]] = dados[i][j];
      }
      convidado._linha = i + 1;
      convidados.push(convidado);
    }
    
    return { sucesso: true, convidados: convidados, colunas: headers };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function adicionarConvidado(nomeEvento, dadosConvidado) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const novaLinha = [];
    
    for (let header of headers) {
      novaLinha.push(dadosConvidado[header] || '');
    }
    
    aba.appendRow(novaLinha);
    return { sucesso: true, mensagem: 'Adicionado' };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function editarConvidado(nomeEvento, linha, dadosConvidado) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const novaLinha = [];
    
    for (let header of headers) {
      novaLinha.push(dadosConvidado[header] || '');
    }
    
    aba.getRange(linha, 1, 1, novaLinha.length).setValues([novaLinha]);
    return { sucesso: true };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function excluirConvidado(nomeEvento, linha) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    aba.deleteRow(linha);
    return { sucesso: true };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function togglePresenca(nomeEvento, linha) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const colPresenca = headers.indexOf('Presença') + 1;
    
    if (colPresenca === 0) throw new Error('Coluna Presença não encontrada');
    
    const celula = aba.getRange(linha, colPresenca);
    const valorAtual = celula.getValue();
    const novoValor = (valorAtual === 'Confirmado') ? '' : 'Confirmado';
    
    celula.setValue(novoValor);
    
    return { sucesso: true, novoStatus: novoValor };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

// ========================================
// 🏦 BANCO DE NOMES
// ========================================

function obterBancoNomes() {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(CONFIG.NOME_ABA_BANCO);
    if (!aba) return { sucesso: true, contatos: [] };
    
    const dados = aba.getDataRange().getValues();
    const contatos = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) {
        contatos.push({ nome: dados[i][0], email: dados[i][1] || '', telefone: dados[i][2] || '' });
      }
    }
    return { sucesso: true, contatos: contatos };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function salvarNoBanco(nome, email, telefone) {
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(CONFIG.NOME_ABA_BANCO);
    
    // Verifica duplicidade simples
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === nome) return { sucesso: true, mensagem: 'Já existe' };
    }
    
    aba.appendRow([nome, email, telefone]);
    return { sucesso: true };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message };
  }
}

function exportarParaExcel(nomeEvento) {
  const sessao = obterSessao();
  const planilha = SpreadsheetApp.openById(sessao.planilhaId);
  const aba = planilha.getSheetByName(nomeEvento);
  const url = `${planilha.getUrl()}/export?format=xlsx&gid=${aba.getSheetId()}`;
  return { sucesso: true, url: url };
}

function importarConvidadosExcel(nomeEvento, dadosJson) {
  // A lógica de importação do Excel agora vem tratada do front-end
  // Reutiliza a função de adicionarConvidado em loop
  try {
    const sessao = obterSessao();
    const planilha = SpreadsheetApp.openById(sessao.planilhaId);
    const aba = planilha.getSheetByName(nomeEvento);
    
    // Mapeia headers
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const novasLinhas = [];
    
    dadosJson.forEach(pessoa => {
      let linha = [];
      headers.forEach(h => linha.push(pessoa[h] || ''));
      novasLinhas.push(linha);
    });
    
    if(novasLinhas.length > 0) {
      aba.getRange(aba.getLastRow()+1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    }
    
    return { sucesso: true, total: novasLinhas.length };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}
