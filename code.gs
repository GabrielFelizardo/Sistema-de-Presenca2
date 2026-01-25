// =======================================================================
// ARQUIVO: CODE.GS (BACKEND v14.5 - EMAIL INTELIGENTE)
// =======================================================================

function doGet(e) {
  return ContentService.createTextOutput("Sistema v14.5 Online! Backend operante.")
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
  return ss.getSheets().map(s => ({ nome: s.getName(), id: s.getSheetId() }));
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
    if (tipo === 'Casamento') colunas = ["Nome", "Mesa", "Restrição", "Status", "Email", "Acompanhantes"];
    else if (tipo === 'Corporativo') colunas = ["Nome", "Empresa", "Cargo", "Status", "Email"];
    else if (tipo === 'Churrasco') colunas = ["Nome", "O que leva", "Status", "Email"];
    else colunas = ["Nome", "Telefone", "Email", "Status"];
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
  
  // ✅ VALIDAÇÃO 1: Nome do evento
  if (!nomeEvento || nomeEvento === 'undefined' || nomeEvento === 'null' || nomeEvento.trim() === '') {
    Logger.log('❌ ERRO: Nome do evento inválido');
    throw new Error("Nome do evento inválido. Valor recebido: '" + nomeEvento + "'");
  }

  // ✅ VALIDAÇÃO 2: Índices
  if (!indices || !Array.isArray(indices) || indices.length === 0) {
    throw new Error("Nenhum convidado selecionado para envio.");
  }

  // ✅ VALIDAÇÃO 3: Limite de segurança
  if (indices.length > 50) {
    throw new Error("Limite de segurança: máximo de 50 envios por vez. Você selecionou " + indices.length + ".");
  }

  // ✅ VALIDAÇÃO 4: Busca a planilha
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

  // ✅ VALIDAÇÃO 5: Verifica se existe coluna Email
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

  // 📊 PROCESSAMENTO
  let enviados = 0;
  let erros = 0;
  let detalhesErros = [];

  indices.forEach((index, i) => {
    try {
      Logger.log(`\n📧 Processando ${i+1}/${indices.length}: índice ${index}`);
      
      // Valida índice
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
      
      // Valida email
      if (!email || !email.includes("@")) {
        Logger.log(`⚠️ AVISO: Email inválido ou vazio`);
        erros++;
        detalhesErros.push(`${nome}: email inválido ou vazio`);
        return;
      }
      
      // Monta o link personalizado
      const linkPersonalizado = `${linkBase}?evento=${encodeURIComponent(nomeEvento.trim())}`;
      
      // Substitui variáveis
      let corpo = mensagem
        .replace(/{Nome}/g, nome)
        .replace(/{Link}/g, linkPersonalizado);
      
      let assuntoFinal = assunto
        .replace(/{Nome}/g, nome)
        .replace(/{Evento}/g, nomeEvento);
      
      // Envia email
      MailApp.sendEmail({
        to: email,
        subject: assuntoFinal,
        body: corpo
      });
      
      Logger.log(`✅ Email enviado com sucesso para ${email}`);
      enviados++;
      
      // Delay para não sobrecarregar (1 email por segundo)
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
