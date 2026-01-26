// =======================================================================
// ARQUIVO: CODE.GS v15.1 FINAL - SEM BUGS
// Sistema RSVP - Templates + Banco de Nomes + Migração
// =======================================================================

function doGet(e) {
  return ContentService.createTextOutput("Sistema v15.1 Online! Backend operante.")
      .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("Sem dados.");
    const dados = JSON.parse(e.postData.contents);
    const acao = dados.acao;
    let resposta = {};

    // LOG para debug de email
    if (acao === 'enviarEmails') {
      Logger.log('=== ENVIO DE EMAILS ===');
      Logger.log('Evento: ' + dados.nomeEvento);
      Logger.log('Selecionados: ' + dados.indices.length);
    }

    // --- ROTEADOR ---
    if (acao === 'listarEventos') resposta = listarEventos();
    else if (acao === 'criarNovoEvento') resposta = criarNovoEvento(dados.nome, dados.template, dados.dadosImportados);
    else if (acao === 'obterDadosEvento') resposta = obterDadosEvento(dados.nomeEvento);
    else if (acao === 'adicionarConvidado') resposta = adicionarConvidado(dados.nomeEvento, dados.arrayDados);
    else if (acao === 'importarListaConvidados') resposta = importarListaInteligente(dados.nomeEvento, dados.matrizDados, dados.temCabecalho);
    else if (acao === 'buscarConvidado') resposta = buscarConvidado(dados.nomeEvento, dados.nomeBusca);
    else if (acao === 'salvarResposta') resposta = salvarResposta(dados.nomeEvento, dados.linha, dados.respostas);
    else if (acao === 'atualizarConvidado') resposta = atualizarConvidado(dados.nomeEvento, dados.linha, dados.novosDados);
    else if (acao === 'excluirConvidado') resposta = excluirConvidado(dados.nomeEvento, dados.linha);
    else if (acao === 'renomearEvento') resposta = renomearEvento(dados.nomeAntigo, dados.nomeNovo);
    else if (acao === 'adicionarColuna') resposta = adicionarColuna(dados.nomeEvento, dados.novaColuna);
    else if (acao === 'removerColuna') resposta = removerColuna(dados.nomeEvento, dados.nomeColuna);
    else if (acao === 'enviarEmails') resposta = enviarEmails(dados.nomeEvento, dados.indices, dados.assunto, dados.mensagem, dados.linkBase);
    
    // --- TEMPLATES (v15.0) ---
    else if (acao === 'salvarTemplate') resposta = salvarTemplate(dados.nomeTemplate, dados.colunas, dados.emailAssunto, dados.emailMensagem);
    else if (acao === 'listarTemplates') resposta = listarTemplates();
    else if (acao === 'excluirTemplate') resposta = excluirTemplate(dados.nomeTemplate);
    else if (acao === 'criarEventoDeTemplate') resposta = criarEventoDeTemplate(dados.nomeEvento, dados.nomeTemplate, dados.usarEmail);
    
    // --- BANCO DE NOMES (v15.0) ---
    else if (acao === 'buscarNoBanco') resposta = buscarNoBanco(dados.termo);
    else if (acao === 'adicionarAoBanco') resposta = adicionarAoBancoNomes(dados.nome, dados.email, dados.telefone, dados.eventoAtual);
    else if (acao === 'listarBancoNomes') resposta = listarBancoNomes();
    else if (acao === 'editarNoBanco') resposta = editarNoBanco(dados.nomeAntigo, dados.nomeNovo, dados.email, dados.telefone, dados.propagarEventos);
    else if (acao === 'excluirDoBanco') resposta = excluirDoBanco(dados.nome);
    else if (acao === 'verificarAtualizacaoDados') resposta = verificarAtualizacaoDados(dados.nome, dados.email, dados.telefone);
    else if (acao === 'atualizarDadosBanco') resposta = atualizarDadosBanco(dados.nome, dados.email, dados.telefone);
    else if (acao === 'excluirEvento') resposta = excluirEvento(dados.nomeEvento);
    else if (acao === 'migrarEventosParaBanco') resposta = migrarEventosParaBanco();
    
    else resposta = { erro: "Ação desconhecida: " + acao };

    return ContentService.createTextOutput(JSON.stringify(resposta)).setMimeType(ContentService.MimeType.JSON);

  } catch (erro) {
    Logger.log('ERRO GERAL: ' + erro.toString());
    return ContentService.createTextOutput(JSON.stringify({ erro: erro.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =======================================================================
// UTILITÁRIOS
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
    sheet.appendRow(["Nome", "Telefone", "Email", "Status"]);
  }
  return id;
}

function getSpreadsheet() { return SpreadsheetApp.openById(getSpreadsheetId()); }

function getSheetByNameSafe(ss, nome) {
  if (!nome || nome === 'undefined' || nome === 'null' || nome === '') {
    Logger.log('ERRO: Nome de aba inválido: "' + nome + '"');
    return null;
  }
  const nomeLimpo = nome.toString().trim();
  return ss.getSheetByName(nomeLimpo);
}

// =======================================================================
// LÓGICA DE NEGÓCIO
// =======================================================================

function listarEventos() {
  const ss = getSpreadsheet();
  const abasSistema = ['Dashboard', 'Templates', 'Banco_Nomes', 'Exemplo'];
  
  return ss.getSheets()
    .filter(s => !abasSistema.includes(s.getName()))
    .map(s => ({ nome: s.getName(), id: s.getSheetId() }));
}

function criarNovoEvento(nome, tipo, dadosRaw) {
  const ss = getSpreadsheet();
  const nomeLimpo = nome.trim();
  if (ss.getSheetByName(nomeLimpo)) throw new Error("Evento já existe!");
  const sheet = ss.insertSheet(nomeLimpo);
  
  let colunas = [];
  let dadosParaInserir = [];

  if (tipo === 'Importar' && dadosRaw) {
    const linhas = dadosRaw.trim().split('\n');
    const matriz = linhas.map(l => l.split('\t'));
    colunas = matriz[0];
    if (matriz.length > 1) dadosParaInserir = matriz.slice(1);
    if (!colunas.includes("Status")) colunas.push("Status");
    if (!colunas.includes("Email") && !colunas.includes("E-mail")) colunas.push("Email");
  } else {
    if (tipo === 'Basico') {
      colunas = ["Nome", "Email", "Telefone", "Confirmado"];
    } else if (tipo === 'Casamento') {
      colunas = ["Nome", "Email", "Telefone", "Confirmado", "Acompanhantes", "Mesa", "Restrição Alimentar", "Mensagem"];
    } else if (tipo === 'Corporativo') {
      colunas = ["Nome", "Email", "Telefone", "Empresa", "Cargo", "Confirmado", "Workshop Escolhido"];
    } else if (tipo === 'Infantil') {
      colunas = ["Nome da Criança", "Idade", "Nome do Responsável", "Email", "Telefone", "Confirmado", "Alergias", "Mensagem"];
    } else if (tipo === 'Formatura') {
      colunas = ["Nome", "Email", "Telefone", "Curso", "Turma", "Confirmado", "Qtd Convites", "Mesa"];
    } else if (tipo === 'Workshop') {
      colunas = ["Nome", "Email", "Telefone", "Confirmado", "Nível de Experiência", "Tópicos de Interesse", "Precisa Material"];
    } else if (tipo === 'Jantar') {
      colunas = ["Nome", "Email", "Telefone", "Confirmado", "Num. de Pessoas", "Horário Preferido", "Restrições Alimentares"];
    } else {
      colunas = ["Nome", "Email", "Telefone", "Confirmado"];
    }
  }

  sheet.appendRow(colunas);
  sheet.getRange(1, 1, 1, colunas.length).setFontWeight("bold").setBackground("#f3f4f6");

  if (dadosParaInserir.length > 0) {
    const indexStatus = colunas.indexOf("Status");
    const dadosFinais = dadosParaInserir.map(linha => {
      let novaLinha = new Array(colunas.length).fill("");
      linha.forEach((d, i) => { if (i < novaLinha.length) novaLinha[i] = d; });
      if (indexStatus > -1 && novaLinha[indexStatus] === "") novaLinha[indexStatus] = "Pendente";
      return novaLinha;
    });
    sheet.getRange(2, 1, dadosFinais.length, dadosFinais[0].length).setValues(dadosFinais);
  }
  return { sucesso: true, nome: nomeLimpo };
}

function obterDadosEvento(nome) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nome);
  if (!sheet) throw new Error("Evento não encontrado: " + nome);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return { headers: [], rows: [] };
  return { headers: data[0], rows: data.slice(1) };
}

function renomearEvento(nomeAntigo, nomeNovo) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeAntigo);
  if (!sheet) throw new Error("Evento original não encontrado.");
  if (getSheetByNameSafe(ss, nomeNovo)) throw new Error("Já existe um evento com o novo nome.");
  sheet.setName(nomeNovo.trim());
  return { sucesso: true };
}

function adicionarColuna(nomeEvento, novaColuna) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");
  
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.includes(novaColuna)) throw new Error("Coluna já existe.");
  
  sheet.getRange(1, lastCol + 1).setValue(novaColuna).setFontWeight("bold").setBackground("#f3f4f6");
  return { sucesso: true };
}

function removerColuna(nomeEvento, nomeColuna) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(nomeColuna);
  
  if (index === -1) throw new Error("Coluna não encontrada.");
  sheet.deleteColumn(index + 1);
  return { sucesso: true };
}

function adicionarConvidado(nomeEvento, dados) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");
  sheet.appendRow(dados);
  return { sucesso: true };
}

function atualizarConvidado(nomeEvento, linhaReal, novosDados) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");
  sheet.getRange(linhaReal, 1, 1, novosDados.length).setValues([novosDados]);
  return { sucesso: true };
}

function excluirConvidado(nomeEvento, linhaReal) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");
  sheet.deleteRow(linhaReal);
  return { sucesso: true };
}

function importarListaInteligente(nomeEvento, matrizDados, temCabecalho) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
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
    if (indiceStatus > -1 && linhaNova[indiceStatus] === "") linhaNova[indiceStatus] = "Pendente";
    return linhaNova;
  });

  if (matrizFinal.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
  }
  return { sucesso: true, qtd: matrizFinal.length };
}

// =======================================================================
// 📧 EMAIL COM VALIDAÇÃO INTELIGENTE
// =======================================================================

function enviarEmails(nomeEvento, indices, assunto, mensagem, linkBase) {
  Logger.log('=== INÍCIO ENVIO DE EMAILS ===');
  Logger.log('Evento: "' + nomeEvento + '" (tipo: ' + typeof nomeEvento + ')');
  Logger.log('Índices: ' + JSON.stringify(indices));
  
  if (!nomeEvento || nomeEvento === 'undefined' || nomeEvento === 'null' || nomeEvento.trim() === '') {
    Logger.log('❌ ERRO: Nome do evento inválido');
    throw new Error("Nome do evento inválido. Valor recebido: '" + nomeEvento + "'");
  }

  if (!indices || !Array.isArray(indices) || indices.length === 0) {
    throw new Error("Nenhum convidado selecionado para envio.");
  }

  if (indices.length > 50) {
    throw new Error("Limite de segurança: máximo de 50 envios por vez. Você selecionou " + indices.length + ".");
  }

  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  
  if (!sheet) {
    const abasDisponiveis = ss.getSheets().map(s => s.getName()).join(', ');
    Logger.log('❌ ERRO: Aba não encontrada');
    Logger.log('Abas disponíveis: ' + abasDisponiveis);
    throw new Error(`Evento "${nomeEvento}" não encontrado. Abas disponíveis: ${abasDisponiveis}`);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  Logger.log('✓ Headers encontrados: ' + headers.join(', '));

  const indexEmail = headers.findIndex(h => {
    const limpo = h.toString().trim().toLowerCase();
    return limpo === 'email' || limpo === 'e-mail';
  });
  
  if (indexEmail === -1) {
    throw new Error("Coluna 'Email' não encontrada. Adicione uma coluna chamada 'Email' ou 'E-mail' na planilha.");
  }

  const indexNome = headers.findIndex(h => h.toString().trim().toLowerCase() === 'nome');
  
  Logger.log(`✓ Coluna Email: índice ${indexEmail}`);
  Logger.log(`✓ Coluna Nome: índice ${indexNome}`);

  let enviados = 0;
  let erros = 0;
  let detalhesErros = [];

  indices.forEach((index, i) => {
    try {
      Logger.log(`\n📧 Processando ${i+1}/${indices.length}: índice ${index}`);
      
      if (index < 0 || index >= rows.length) {
        Logger.log(`⚠️ AVISO: Índice ${index} fora do range (0-${rows.length-1})`);
        erros++;
        detalhesErros.push(`Índice ${index} inválido (fora do range)`);
        return;
      }
      
      const row = rows[index];
      const email = row[indexEmail] ? row[indexEmail].toString().trim() : '';
      const nome = indexNome > -1 ? row[indexNome] : "Convidado";
      
      Logger.log(`  Nome: ${nome}`);
      Logger.log(`  Email: ${email}`);
      
      if (!email || !email.includes("@")) {
        Logger.log(`⚠️ AVISO: Email inválido ou vazio`);
        erros++;
        detalhesErros.push(`${nome}: email inválido ou vazio`);
        return;
      }
      
      const linkPersonalizado = `${linkBase}?evento=${encodeURIComponent(nomeEvento.trim())}`;
      
      let corpo = mensagem
        .replace(/{Nome}/g, nome)
        .replace(/{Link}/g, linkPersonalizado);
      
      let assuntoFinal = assunto
        .replace(/{Nome}/g, nome)
        .replace(/{Evento}/g, nomeEvento);
      
      MailApp.sendEmail({
        to: email,
        subject: assuntoFinal,
        body: corpo
      });
      
      Logger.log(`✅ Email enviado com sucesso para ${email}`);
      enviados++;
      
      if (i < indices.length - 1) {
        Utilities.sleep(1000);
      }
      
    } catch (e) {
      Logger.log(`❌ ERRO ao enviar para índice ${index}: ${e.toString()}`);
      erros++;
      const nomeErro = indexNome > -1 && rows[index] ? rows[index][indexNome] : `Índice ${index}`;
      detalhesErros.push(`${nomeErro}: ${e.message}`);
    }
  });
  
  Logger.log(`\n=== RESUMO ===`);
  Logger.log(`✅ Enviados: ${enviados}`);
  Logger.log(`❌ Erros: ${erros}`);
  if (detalhesErros.length > 0) {
    Logger.log(`Detalhes dos erros:`);
    detalhesErros.forEach(det => Logger.log(`  - ${det}`));
  }
  
  return { 
    sucesso: true, 
    enviados: enviados, 
    erros: erros,
    detalhes: detalhesErros.length > 0 ? detalhesErros : null
  };
}

// =======================================================================
// CONVIDADO
// =======================================================================

function buscarConvidado(nomeEvento, nomeBusca) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sucesso: true, encontrado: false };

  const headers = data[0];
  const linhas = data.slice(1);
  const nomeLimpo = nomeBusca.toLowerCase().trim();

  for (let i = 0; i < linhas.length; i++) {
    const row = linhas[i];
    const indexNome = headers.findIndex(h => h.toString().trim().toLowerCase() === 'nome');
    const valNome = indexNome > -1 ? row[indexNome] : row[0];
    
    if (valNome.toString().toLowerCase().includes(nomeLimpo)) {
      return {
        sucesso: true,
        encontrado: true,
        nomeCompleto: valNome,
        linha: i + 2,
        colunas: headers,
        dadosAtuais: row
      };
    }
  }
  return { sucesso: true, encontrado: false };
}

function salvarResposta(nomeEvento, linha, respostas) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (const [coluna, valor] of Object.entries(respostas)) {
    const colIndex = headers.indexOf(coluna);
    if (colIndex > -1) sheet.getRange(linha, colIndex + 1).setValue(valor);
  }
  return { sucesso: true };
}

// =======================================================================
// SISTEMA DE TEMPLATES (v15.0)
// =======================================================================

function inicializarAbaTemplates() {
  const ss = getSpreadsheet();
  let sheetTemplates = ss.getSheetByName("Templates");
  
  if (!sheetTemplates) {
    sheetTemplates = ss.insertSheet("Templates");
    sheetTemplates.appendRow(["Nome Template", "Colunas (JSON)", "Email Assunto", "Email Mensagem", "Data Criação", "Vezes Usado"]);
    sheetTemplates.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#dbeafe");
    sheetTemplates.setFrozenRows(1);
  }
  
  return sheetTemplates;
}

function salvarTemplate(nomeTemplate, colunas, emailAssunto, emailMensagem) {
  const sheetTemplates = inicializarAbaTemplates();
  const nomeLimpo = nomeTemplate.trim();
  
  if (!nomeLimpo) throw new Error("Nome do template não pode estar vazio.");
  
  const data = sheetTemplates.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeLimpo) {
      throw new Error("Já existe um template com este nome.");
    }
  }
  
  const colunasJSON = JSON.stringify(colunas);
  const dataAtual = new Date().toLocaleDateString('pt-BR');
  
  sheetTemplates.appendRow([
    nomeLimpo,
    colunasJSON,
    emailAssunto || "",
    emailMensagem || "",
    dataAtual,
    0
  ]);
  
  return { sucesso: true, nome: nomeLimpo };
}

function listarTemplates() {
  const sheetTemplates = inicializarAbaTemplates();
  const data = sheetTemplates.getDataRange().getValues();
  
  if (data.length < 2) return { templates: [] };
  
  const templates = [];
  for (let i = 1; i < data.length; i++) {
    templates.push({
      nome: data[i][0],
      colunas: JSON.parse(data[i][1]),
      emailAssunto: data[i][2],
      emailMensagem: data[i][3],
      dataCriacao: data[i][4],
      vezesUsado: data[i][5]
    });
  }
  
  return { templates: templates };
}

function excluirTemplate(nomeTemplate) {
  const sheetTemplates = inicializarAbaTemplates();
  const data = sheetTemplates.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeTemplate) {
      sheetTemplates.deleteRow(i + 1);
      return { sucesso: true };
    }
  }
  
  throw new Error("Template não encontrado.");
}

function criarEventoDeTemplate(nomeEvento, nomeTemplate, usarEmail) {
  const sheetTemplates = inicializarAbaTemplates();
  const data = sheetTemplates.getDataRange().getValues();
  
  let templateEncontrado = null;
  let linhaTemplate = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeTemplate) {
      templateEncontrado = {
        colunas: JSON.parse(data[i][1]),
        emailAssunto: data[i][2],
        emailMensagem: data[i][3]
      };
      linhaTemplate = i + 1;
      break;
    }
  }
  
  if (!templateEncontrado) throw new Error("Template não encontrado.");
  
  const ss = getSpreadsheet();
  const nomeLimpo = nomeEvento.trim();
  if (ss.getSheetByName(nomeLimpo)) throw new Error("Evento já existe!");
  
  const sheet = ss.insertSheet(nomeLimpo);
  const colunas = templateEncontrado.colunas;
  
  sheet.appendRow(colunas);
  sheet.getRange(1, 1, 1, colunas.length).setFontWeight("bold").setBackground("#f3f4f6");
  
  const vezesUsado = data[linhaTemplate - 1][5] || 0;
  sheetTemplates.getRange(linhaTemplate, 6).setValue(vezesUsado + 1);
  
  const resultado = { 
    sucesso: true, 
    nome: nomeLimpo,
    colunas: colunas
  };
  
  if (usarEmail) {
    resultado.emailAssunto = templateEncontrado.emailAssunto;
    resultado.emailMensagem = templateEncontrado.emailMensagem;
  }
  
  return resultado;
}

// =======================================================================
// BANCO DE NOMES (v15.0)
// =======================================================================

function inicializarBancoNomes() {
  const ss = getSpreadsheet();
  let sheetBanco = ss.getSheetByName("Banco_Nomes");
  
  if (!sheetBanco) {
    sheetBanco = ss.insertSheet("Banco_Nomes");
    sheetBanco.appendRow(["Nome Completo", "Email", "Telefone", "Eventos Participou", "Último Evento", "Data Última Participação"]);
    sheetBanco.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#dcfce7");
    sheetBanco.setFrozenRows(1);
    sheetBanco.setColumnWidth(4, 300);
  }
  
  return sheetBanco;
}

function buscarNoBanco(termo) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  if (data.length < 2) return { encontrado: false };
  
  const termoLimpo = termo.toLowerCase().trim();
  const resultados = [];
  
  for (let i = 1; i < data.length; i++) {
    const nome = data[i][0].toString().toLowerCase();
    
    if (nome.includes(termoLimpo)) {
      const eventosArray = data[i][3] ? data[i][3].split(';').filter(e => e.trim()) : [];
      
      resultados.push({
        nome: data[i][0],
        email: data[i][1],
        telefone: data[i][2],
        eventosParticipou: eventosArray,
        ultimoEvento: data[i][4] || "Nenhum",
        dataUltimaParticipacao: data[i][5] || "N/A",
        totalEventos: eventosArray.length
      });
    }
  }
  
  if (resultados.length === 0) {
    return { encontrado: false };
  }
  
  resultados.sort((a, b) => {
    if (!a.dataUltimaParticipacao || a.dataUltimaParticipacao === "N/A") return 1;
    if (!b.dataUltimaParticipacao || b.dataUltimaParticipacao === "N/A") return -1;
    return new Date(b.dataUltimaParticipacao) - new Date(a.dataUltimaParticipacao);
  });
  
  return { 
    encontrado: true, 
    resultados: resultados 
  };
}

function adicionarAoBancoNomes(nome, email, telefone, eventoAtual) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  const nomeLimpo = nome.trim();
  const emailLimpo = email ? email.trim() : "";
  const telefoneLimpo = telefone ? telefone.trim() : "";
  const dataAtual = new Date().toLocaleDateString('pt-BR');
  
  let pessoaExiste = false;
  let linhaPessoa = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nomeLimpo.toLowerCase()) {
      pessoaExiste = true;
      linhaPessoa = i + 1;
      break;
    }
  }
  
  if (pessoaExiste) {
    const eventosAtuais = data[linhaPessoa - 1][3] || "";
    const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
    
    if (!eventosArray.includes(eventoAtual)) {
      eventosArray.push(eventoAtual);
    }
    
    const eventosNovos = eventosArray.join(';');
    
    sheetBanco.getRange(linhaPessoa, 2).setValue(emailLimpo || data[linhaPessoa - 1][1]);
    sheetBanco.getRange(linhaPessoa, 3).setValue(telefoneLimpo || data[linhaPessoa - 1][2]);
    sheetBanco.getRange(linhaPessoa, 4).setValue(eventosNovos);
    sheetBanco.getRange(linhaPessoa, 5).setValue(eventoAtual);
    sheetBanco.getRange(linhaPessoa, 6).setValue(dataAtual);
    
    return { sucesso: true, atualizado: true };
    
  } else {
    sheetBanco.appendRow([
      nomeLimpo,
      emailLimpo,
      telefoneLimpo,
      eventoAtual,
      eventoAtual,
      dataAtual
    ]);
    
    return { sucesso: true, atualizado: false };
  }
}

function listarBancoNomes() {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  if (data.length < 2) return { nomes: [] };
  
  const nomes = [];
  for (let i = 1; i < data.length; i++) {
    const eventosArray = data[i][3] ? data[i][3].split(';').filter(e => e.trim()) : [];
    
    nomes.push({
      nome: data[i][0],
      email: data[i][1] || "",
      telefone: data[i][2] || "",
      eventosParticipou: eventosArray,
      ultimoEvento: data[i][4] || "Nenhum",
      dataUltimaParticipacao: data[i][5] || "N/A",
      totalEventos: eventosArray.length,
      linha: i + 1
    });
  }
  
  return { nomes: nomes };
}

function editarNoBanco(nomeAntigo, nomeNovo, email, telefone, propagarEventos) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  let linhaPessoa = -1;
  let eventosParticipou = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nomeAntigo.toLowerCase()) {
      linhaPessoa = i + 1;
      eventosParticipou = data[i][3] ? data[i][3].split(';').filter(e => e.trim()) : [];
      break;
    }
  }
  
  if (linhaPessoa === -1) throw new Error("Nome não encontrado no banco.");
  
  sheetBanco.getRange(linhaPessoa, 1).setValue(nomeNovo.trim());
  sheetBanco.getRange(linhaPessoa, 2).setValue(email || "");
  sheetBanco.getRange(linhaPessoa, 3).setValue(telefone || "");
  
  if (propagarEventos && nomeAntigo !== nomeNovo) {
    const ss = getSpreadsheet();
    let eventosAtualizados = 0;
    
    eventosParticipou.forEach(nomeEvento => {
      const sheet = getSheetByNameSafe(ss, nomeEvento);
      if (!sheet) return;
      
      const dataEvento = sheet.getDataRange().getValues();
      const headers = dataEvento[0];
      const indexNome = headers.findIndex(h => h.toString().toLowerCase().includes('nome'));
      
      if (indexNome === -1) return;
      
      for (let i = 1; i < dataEvento.length; i++) {
        if (dataEvento[i][indexNome].toString().toLowerCase() === nomeAntigo.toLowerCase()) {
          sheet.getRange(i + 1, indexNome + 1).setValue(nomeNovo);
          eventosAtualizados++;
        }
      }
    });
    
    return { 
      sucesso: true, 
      propagado: true, 
      eventosAtualizados: eventosAtualizados 
    };
  }
  
  return { sucesso: true, propagado: false };
}

function excluirDoBanco(nome) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nome.toLowerCase()) {
      sheetBanco.deleteRow(i + 1);
      return { sucesso: true };
    }
  }
  
  throw new Error("Nome não encontrado.");
}

function verificarAtualizacaoDados(nome, email, telefone) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nome.toLowerCase()) {
      const emailBanco = data[i][1] ? data[i][1].toString().trim() : "";
      const telefoneBanco = data[i][2] ? data[i][2].toString().trim() : "";
      const emailNovo = email ? email.toString().trim() : "";
      const telefoneNovo = telefone ? telefone.toString().trim() : "";
      
      const temEmailNovo = emailNovo && emailNovo !== "" && emailNovo !== emailBanco;
      const temTelefoneNovo = telefoneNovo && telefoneNovo !== "" && telefoneNovo !== telefoneBanco;
      
      if (temEmailNovo || temTelefoneNovo) {
        return {
          precisaAtualizar: true,
          nome: data[i][0],
          emailAtual: emailBanco,
          emailNovo: temEmailNovo ? emailNovo : null,
          telefoneAtual: telefoneBanco,
          telefoneNovo: temTelefoneNovo ? telefoneNovo : null,
          linha: i + 1
        };
      }
      
      return { precisaAtualizar: false };
    }
  }
  
  return { precisaAtualizar: false, novoNome: true };
}

function atualizarDadosBanco(nome, email, telefone) {
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nome.toLowerCase()) {
      const linhaPessoa = i + 1;
      
      if (email && email.trim() !== "") {
        sheetBanco.getRange(linhaPessoa, 2).setValue(email.trim());
      }
      
      if (telefone && telefone.trim() !== "") {
        sheetBanco.getRange(linhaPessoa, 3).setValue(telefone.trim());
      }
      
      return { sucesso: true };
    }
  }
  
  throw new Error("Nome não encontrado no banco.");
}

function excluirEvento(nomeEvento) {
  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  
  if (!sheet) throw new Error("Evento não encontrado.");
  
  const sheetBanco = inicializarBancoNomes();
  const data = sheetBanco.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const eventosAtuais = data[i][3] || "";
    const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
    
    const eventosAtualizados = eventosArray.filter(e => e !== nomeEvento);
    
    if (eventosAtualizados.length !== eventosArray.length) {
      const novosEventos = eventosAtualizados.join(';');
      sheetBanco.getRange(i + 1, 4).setValue(novosEventos);
      
      if (data[i][4] === nomeEvento) {
        const novoUltimo = eventosAtualizados.length > 0 ? eventosAtualizados[eventosAtualizados.length - 1] : "Nenhum";
        sheetBanco.getRange(i + 1, 5).setValue(novoUltimo);
      }
    }
  }
  
  ss.deleteSheet(sheet);
  
  return { sucesso: true, mensagem: "Evento excluído. Nomes mantidos no banco." };
}

// =======================================================================
// MIGRAÇÃO: POPULAR BANCO COM EVENTOS ANTIGOS (v15.1)
// =======================================================================

function migrarEventosParaBanco() {
  const ss = getSpreadsheet();
  const sheetBanco = inicializarBancoNomes();
  const abasSistema = ['Dashboard', 'Templates', 'Banco_Nomes', 'Exemplo'];
  
  let totalAdicionados = 0;
  let totalAtualizados = 0;
  
  const eventos = ss.getSheets()
    .filter(s => !abasSistema.includes(s.getName()))
    .map(s => s.getName());
  
  eventos.forEach(nomeEvento => {
    const sheet = ss.getSheetByName(nomeEvento);
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) return;
    
    const headers = data[0];
    const indexNome = headers.indexOf('Nome') >= 0 ? headers.indexOf('Nome') : 
                      headers.indexOf('Nome Completo') >= 0 ? headers.indexOf('Nome Completo') : 0;
    const indexEmail = headers.findIndex(h => {
      const limpo = h.toString().trim().toLowerCase();
      return limpo === 'email' || limpo === 'e-mail';
    });
    const indexTelefone = headers.indexOf('Telefone');
    
    for (let i = 1; i < data.length; i++) {
      const nome = data[i][indexNome];
      if (!nome || nome.toString().trim() === '') continue;
      
      const email = indexEmail >= 0 ? data[i][indexEmail] : '';
      const telefone = indexTelefone >= 0 ? data[i][indexTelefone] : '';
      
      try {
        const resultado = adicionarAoBancoNomes(
          nome.toString().trim(),
          email ? email.toString().trim() : '',
          telefone ? telefone.toString().trim() : '',
          nomeEvento
        );
        
        if (resultado.atualizado) {
          totalAtualizados++;
        } else {
          totalAdicionados++;
        }
      } catch (e) {
        Logger.log(`Erro ao migrar ${nome} do evento ${nomeEvento}: ${e}`);
      }
    }
  });
  
  return {
    sucesso: true,
    adicionados: totalAdicionados,
    atualizados: totalAtualizados,
    total: totalAdicionados + totalAtualizados,
    eventos: eventos.length
  };
}
