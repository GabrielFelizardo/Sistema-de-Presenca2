// =======================================================================
// ARQUIVO: CODE.GS (BACKEND v14.3 - CORREÇÃO UNDEFINED)
// =======================================================================

function doGet(e) {
  return ContentService.createTextOutput("Sistema v14.3 Online! Backend operante.")
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

// SEGURANÇA: Função para pegar aba garantindo que o nome é válido
function getSheetByNameSafe(ss, nome) {
  if (!nome || nome === 'undefined' || nome === 'null') return null;
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

// --- ESTRUTURA ---
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

// --- EDIÇÃO ---
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
    if (indexStatus > -1 && linhaNova[indexStatus] === "") linhaNova[indexStatus] = "Pendente";
    return linhaNova;
  });

  if (matrizFinal.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
  }
  return { sucesso: true, qtd: matrizFinal.length };
}

// --- EMAIL BLINDADO ---
function enviarEmails(nomeEvento, indices, assunto, mensagem, linkBase) {
  if (!nomeEvento || nomeEvento === 'undefined') throw new Error("Nome do evento inválido ou não selecionado.");

  const ss = getSpreadsheet();
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  
  if (!sheet) throw new Error(`Não foi possível encontrar a aba '${nomeEvento}'. Verifique espaços ou caracteres especiais.`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const indexEmail = headers.findIndex(h => {
    const limpo = h.toString().trim().toLowerCase();
    return limpo === 'email' || limpo === 'e-mail';
  });
  
  const indexNome = headers.findIndex(h => h.toString().trim().toLowerCase() === 'nome');
  
  if (indexEmail === -1) throw new Error("Coluna 'Email' não encontrada.");
  if (indices.length > 50) throw new Error("Limite de segurança: 50 envios por vez.");

  let enviados = 0;
  let erros = 0;

  indices.forEach(index => {
    try {
      const row = rows[index];
      const email = row[indexEmail];
      const nome = indexNome > -1 ? row[indexNome] : "Convidado";
      
      if (email && email.toString().includes("@")) {
        const linkPersonalizado = `${linkBase}?evento=${encodeURIComponent(nomeEvento.trim())}`;
        
        let corpo = mensagem
          .replace(/{Nome}/g, nome)
          .replace(/{Link}/g, linkPersonalizado);
          
        MailApp.sendEmail({
          to: email,
          subject: assunto.replace(/{Nome}/g, nome),
          body: corpo
        });
        enviados++;
      } else {
        erros++;
      }
    } catch (e) {
      console.error(e);
      erros++;
    }
  });
  
  return { sucesso: true, enviados: enviados, erros: erros };
}

// --- CONVIDADO ---
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
