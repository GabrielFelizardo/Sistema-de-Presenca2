// =======================================================================
// ARQUIVO: CODE.GS (BACKEND v5 - COMPLETO COM FORMULÁRIO PÚBLICO)
// =======================================================================

function doGet(e) {
  // ROTEAMENTO: Decide qual HTML mostrar baseado na URL
  // Ex: .../exec?page=form -> Mostra Formulário
  // Ex: .../exec -> Mostra Painel Admin
  
  if (e.parameter.page === 'form') {
    return HtmlService.createTemplateFromFile('form')
        .evaluate()
        .setTitle('Confirmar Presença')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('painel')
      .evaluate()
      .setTitle('Gestor de Eventos - Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ROTEADOR API (POST)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const dados = JSON.parse(e.postData.contents);
    const acao = dados.acao;
    let resposta = {};

    // --- AÇÕES DO ADMIN ---
    if (acao === 'listarEventos') resposta = listarEventos();
    else if (acao === 'criarNovoEvento') resposta = criarNovoEvento(dados.nome, dados.template, dados.dadosImportados);
    else if (acao === 'obterDadosEvento') resposta = obterDadosEvento(dados.nomeEvento);
    else if (acao === 'adicionarConvidado') resposta = adicionarConvidado(dados.nomeEvento, dados.arrayDados);
    else if (acao === 'importarListaConvidados') resposta = importarListaInteligente(dados.nomeEvento, dados.matrizDados, dados.temCabecalho);
    
    // --- AÇÕES DO CONVIDADO (NOVO) ---
    else if (acao === 'buscarConvidado') resposta = buscarConvidado(dados.nomeEvento, dados.nomeBusca);
    else if (acao === 'salvarResposta') resposta = salvarResposta(dados.nomeEvento, dados.linha, dados.respostas);
    
    else resposta = { erro: "Ação desconhecida" };

    return ContentService.createTextOutput(JSON.stringify(resposta)).setMimeType(ContentService.MimeType.JSON);

  } catch (erro) {
    return ContentService.createTextOutput(JSON.stringify({ erro: erro.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =======================================================================
// LÓGICA DE NEGÓCIO
// =======================================================================

function getSpreadsheetId() {
  const props = PropertiesService.getUserProperties();
  let id = props.getProperty('ID_MINHA_PLANILHA');
  if (!id) {
    const ss = SpreadsheetApp.create("Meus Eventos (Sistema RSVP)");
    id = ss.getId();
    props.setProperty('ID_MINHA_PLANILHA', id);
    const sheet = ss.getSheets()[0];
    sheet.setName("Exemplo");
    sheet.appendRow(["Nome", "Telefone", "Status"]);
  }
  return id;
}

function getSpreadsheet() { return SpreadsheetApp.openById(getSpreadsheetId()); }

// --- FUNÇÕES ADMIN ---
function listarEventos() {
  const ss = getSpreadsheet();
  return ss.getSheets().map(s => ({ nome: s.getName(), id: s.getSheetId() }));
}

function criarNovoEvento(nome, tipo, dadosRaw) {
  const ss = getSpreadsheet();
  if (ss.getSheetByName(nome)) throw new Error("Evento já existe!");
  
  const sheet = ss.insertSheet(nome);
  let colunas = [];
  let dadosParaInserir = [];

  if (tipo === 'Importar' && dadosRaw) {
    const linhas = dadosRaw.trim().split('\n');
    const matriz = linhas.map(l => l.split('\t'));
    colunas = matriz[0];
    if (matriz.length > 1) dadosParaInserir = matriz.slice(1);
    if (!colunas.includes("Status")) colunas.push("Status");
  } else {
    if (tipo === 'Casamento') colunas = ["Nome", "Mesa", "Restrição", "Status", "Acompanhantes"];
    else if (tipo === 'Corporativo') colunas = ["Nome", "Empresa", "Cargo", "Status", "Email"];
    else if (tipo === 'Churrasco') colunas = ["Nome", "O que leva", "Status"];
    else colunas = ["Nome", "Telefone", "Status", "Email"];
  }

  sheet.appendRow(colunas);
  sheet.getRange(1, 1, 1, colunas.length).setFontWeight("bold").setBackground("#f3f4f6");

  if (dadosParaInserir.length > 0) {
    const indexStatus = colunas.indexOf("Status");
    const dadosFinais = dadosParaInserir.map(linha => {
      let novaLinha = new Array(colunas.length).fill("");
      linha.forEach((d, i) => { if (i < novaLinha.length) novaLinha[i] = d; });
      if (novaLinha[indexStatus] === "") novaLinha[indexStatus] = "Pendente";
      return novaLinha;
    });
    sheet.getRange(2, 1, dadosFinais.length, dadosFinais[0].length).setValues(dadosFinais);
  }
  return { sucesso: true, nome: nome };
}

function obterDadosEvento(nome) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(nome);
  if (!sheet) throw new Error("Aba não encontrada");
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return { headers: [], rows: [] };
  return { headers: data[0], rows: data.slice(1) };
}

function adicionarConvidado(nomeEvento, dados) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeEvento);
  sheet.appendRow(dados);
  return { sucesso: true };
}

function importarListaInteligente(nomeEvento, matrizDados, temCabecalho) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  
  let colunasAtuais = sheet.getDataRange().getValues()[0] || [];
  let dadosParaInserir = matrizDados;
  
  if (temCabecalho) {
    const novosCabecalhos = matrizDados[0];
    dadosParaInserir = matrizDados.slice(1);
    novosCabecalhos.forEach((novoHeader) => {
      if (!colunasAtuais.includes(novoHeader)) {
        const novaColIndex = colunasAtuais.length + 1;
        sheet.getRange(1, novaColIndex).setValue(novoHeader).setFontWeight("bold");
        colunasAtuais.push(novoHeader);
      }
    });
  }

  if (!colunasAtuais.includes("Status")) {
    const novaColIndex = colunasAtuais.length + 1;
    sheet.getRange(1, novaColIndex).setValue("Status").setFontWeight("bold");
    colunasAtuais.push("Status");
  }

  const indiceStatus = colunasAtuais.indexOf("Status");
  const matrizFinal = dadosParaInserir.map(linhaImportada => {
    let linhaNova = new Array(colunasAtuais.length).fill(""); 
    if (temCabecalho) {
      matrizDados[0].forEach((header, indexOrigem) => {
        const indexDestino = colunasAtuais.indexOf(header);
        if (indexDestino > -1) linhaNova[indexDestino] = linhaImportada[indexOrigem];
      });
    } else {
      linhaImportada.forEach((dado, index) => { if (index < linhaNova.length) linhaNova[index] = dado; });
    }
    if (linhaNova[indiceStatus] === "") linhaNova[indiceStatus] = "Pendente";
    return linhaNova;
  });

  if (matrizFinal.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
  }
  return { sucesso: true, qtd: matrizFinal.length };
}

// --- FUNÇÕES DO CONVIDADO (RSVP) ---

function buscarConvidado(nomeEvento, nomeBusca) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado (verifique o link).");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const linhas = data.slice(1);
  const nomeLimpo = nomeBusca.toLowerCase().trim();

  // Procura o nome (busca simples por "contém")
  // Retorna a primeira ocorrência encontrada
  for (let i = 0; i < linhas.length; i++) {
    const row = linhas[i];
    // Assume que o Nome é a primeira coluna (índice 0)
    if (row[0].toString().toLowerCase().includes(nomeLimpo)) {
      return {
        sucesso: true,
        encontrado: true,
        nomeCompleto: row[0],
        linha: i + 2, // +2 porque pula cabeçalho e array começa em 0
        colunas: headers,
        dadosAtuais: row
      };
    }
  }

  return { sucesso: true, encontrado: false };
}

function salvarResposta(nomeEvento, linha, respostas) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeEvento);
  
  // Atualiza as colunas necessárias
  // respostas é um objeto: { "Status": "Confirmado", "Alergia": "Gluten" }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (const [coluna, valor] of Object.entries(respostas)) {
    const colIndex = headers.indexOf(coluna);
    if (colIndex > -1) {
      sheet.getRange(linha, colIndex + 1).setValue(valor);
    }
  }
  
  return { sucesso: true };
}
