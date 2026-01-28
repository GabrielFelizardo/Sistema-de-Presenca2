// =======================================================================
// CODE.GS v16.5.1 - CORRIGIDO - SISTEMA COM APROVAÇÃO MANUAL
// Bug fix: TypeError data[i][3].split is not a function
// =======================================================================

// ⚠️ CONFIGURAÇÕES:
const ID_PLANILHA_CONTROLE = ""; // ← Deixe vazio no primeiro deploy
const EMAIL_ADMIN = "gfelizardo14@gmail.com"; // ← SEU EMAIL PARA NOTIFICAÇÕES

// ========== FUNÇÃO AUXILIAR - GARANTIR STRING ==========
function garantirString(valor) {
  if (valor === null || valor === undefined) {
    return "";
  }
  return valor.toString();
}

function doGet(e) {
  return ContentService.createTextOutput("Sistema RSVP v16.5.1 Online!")
      .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Sem dados na requisição.");
    }

    const dados = JSON.parse(e.postData.contents);
    const acao = dados.acao;
    let resposta = {};

    Logger.log('📥 Ação: ' + acao);

    // ROTEADOR
    if (acao === 'autenticar') {
      resposta = autenticarUsuario(dados.email);
    }
    else if (acao === 'listarEventos') {
      resposta = listarEventos(dados.email);
    }
    else if (acao === 'criarNovoEvento') {
      resposta = criarNovoEvento(dados.email, dados.nome, dados.template, dados.dadosImportados);
    }
    else if (acao === 'obterDadosEvento') {
      resposta = obterDadosEvento(dados.email, dados.nomeEvento);
    }
    else if (acao === 'adicionarConvidado') {
      resposta = adicionarConvidado(dados.email, dados.nomeEvento, dados.arrayDados);
    }
    else if (acao === 'importarListaConvidados') {
      resposta = importarListaInteligente(dados.email, dados.nomeEvento, dados.matrizDados, dados.temCabecalho);
    }
    else if (acao === 'buscarConvidado') {
      resposta = buscarConvidado(dados.email, dados.nomeEvento, dados.nomeBusca);
    }
    else if (acao === 'salvarResposta') {
      resposta = salvarResposta(dados.email, dados.nomeEvento, dados.linha, dados.respostas);
    }
    else if (acao === 'atualizarConvidado') {
      resposta = atualizarConvidado(dados.email, dados.nomeEvento, dados.linha, dados.novosDados);
    }
    else if (acao === 'excluirConvidado') {
      resposta = excluirConvidado(dados.email, dados.nomeEvento, dados.linha);
    }
    else if (acao === 'renomearEvento') {
      resposta = renomearEvento(dados.email, dados.nomeAntigo, dados.nomeNovo);
    }
    else if (acao === 'adicionarColuna') {
      resposta = adicionarColuna(dados.email, dados.nomeEvento, dados.novaColuna);
    }
    else if (acao === 'removerColuna') {
      resposta = removerColuna(dados.email, dados.nomeEvento, dados.nomeColuna);
    }
    else if (acao === 'enviarEmails') {
      resposta = enviarEmails(dados.email, dados.nomeEvento, dados.indices, dados.assunto, dados.mensagem, dados.linkBase);
    }
    else if (acao === 'excluirEvento') {
      resposta = excluirEvento(dados.email, dados.nomeEvento);
    }
    else if (acao === 'salvarTemplate') {
      resposta = salvarTemplate(dados.email, dados.nomeTemplate, dados.colunas, dados.emailAssunto, dados.emailMensagem);
    }
    else if (acao === 'listarTemplates') {
      resposta = listarTemplates(dados.email);
    }
    else if (acao === 'excluirTemplate') {
      resposta = excluirTemplate(dados.email, dados.nomeTemplate);
    }
    else if (acao === 'criarEventoDeTemplate') {
      resposta = criarEventoDeTemplate(dados.email, dados.nomeEvento, dados.nomeTemplate, dados.usarEmail);
    }
    else if (acao === 'buscarNoBanco') {
      resposta = buscarNoBanco(dados.email, dados.termo);
    }
    else if (acao === 'adicionarAoBanco') {
      resposta = adicionarAoBancoNomes(dados.email, dados.nome, dados.emailContato, dados.telefone, dados.eventoAtual);
    }
    else if (acao === 'listarBancoNomes') {
      resposta = listarBancoNomes(dados.email);
    }
    else if (acao === 'editarNoBanco') {
      resposta = editarNoBanco(dados.email, dados.nomeAntigo, dados.nomeNovo, dados.emailContato, dados.telefone, dados.propagarEventos);
    }
    else if (acao === 'excluirDoBanco') {
      resposta = excluirDoBanco(dados.email, dados.nome);
    }
    else if (acao === 'migrarEventosParaBanco') {
      resposta = migrarEventosParaBanco(dados.email);
    }
    else {
      resposta = { erro: "Ação desconhecida: " + acao };
    }

    return ContentService.createTextOutput(JSON.stringify(resposta))
        .setMimeType(ContentService.MimeType.JSON);

  } catch (erro) {
    Logger.log('❌ ERRO: ' + erro.toString());
    return ContentService.createTextOutput(JSON.stringify({ 
      erro: erro.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =======================================================================
// AUTENTICAÇÃO v16.5 - COM SISTEMA DE APROVAÇÃO
// =======================================================================

function autenticarUsuario(email) {
  if (!email || !email.includes('@')) {
    throw new Error("Email inválido.");
  }

  const emailLimpo = email.toLowerCase().trim();
  Logger.log('🔐 Autenticando: ' + emailLimpo);
  
  // ✅ PASSO 1: Criar ou abrir planilha de controle
  let ssControle;
  
  if (!ID_PLANILHA_CONTROLE || ID_PLANILHA_CONTROLE === "") {
    // 🆕 PRIMEIRA VEZ - Criar planilha de controle
    Logger.log('🆕 Criando planilha de controle...');
    ssControle = SpreadsheetApp.create("Controle_Usuarios_RSVP");
    const idControle = ssControle.getId();
    
    Logger.log('✅ Planilha de controle criada!');
    Logger.log('📋 ID: ' + idControle);
    Logger.log('');
    Logger.log('⚠️ IMPORTANTE - COPIE E COLE NO CÓDIGO:');
    Logger.log('const ID_PLANILHA_CONTROLE = "' + idControle + '";');
    Logger.log('');
    
    // Configurar aba com coluna Status
    const sheetUsuarios = ssControle.getSheets()[0];
    sheetUsuarios.setName("Usuarios");
    sheetUsuarios.appendRow(["Email", "Status", "ID_Planilha", "Data_Cadastro", "Data_Aprovacao", "Ultimo_Acesso"]);
    sheetUsuarios.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#667eea").setFontColor("white");
    sheetUsuarios.setFrozenRows(1);
    sheetUsuarios.setColumnWidth(1, 250);
    sheetUsuarios.setColumnWidth(2, 100);
    sheetUsuarios.setColumnWidth(3, 300);
    
  } else {
    // ✅ JÁ EXISTE - Abrir planilha de controle
    ssControle = SpreadsheetApp.openById(ID_PLANILHA_CONTROLE);
  }
  
  let sheetUsuarios = ssControle.getSheetByName("Usuarios");
  if (!sheetUsuarios) {
    sheetUsuarios = ssControle.getSheets()[0];
    sheetUsuarios.setName("Usuarios");
    if (sheetUsuarios.getLastRow() === 0) {
      sheetUsuarios.appendRow(["Email", "Status", "ID_Planilha", "Data_Cadastro", "Data_Aprovacao", "Ultimo_Acesso"]);
      sheetUsuarios.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#667eea").setFontColor("white");
      sheetUsuarios.setFrozenRows(1);
    }
  }

  // ✅ PASSO 2: Verificar se usuário já existe
  const data = sheetUsuarios.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && garantirString(data[i][0]).toLowerCase() === emailLimpo) {
      const status = garantirString(data[i][1]);
      const planilhaId = garantirString(data[i][2]);
      
      Logger.log('👤 Usuário encontrado - Status: ' + status);
      
      // ⚠️ VERIFICAR STATUS
      if (status === 'Pendente') {
        Logger.log('🟡 Acesso pendente de aprovação');
        return {
          sucesso: false,
          status: 'pendente',
          mensagem: 'Seu acesso está aguardando aprovação do administrador.'
        };
      }
      
      if (status === 'Rejeitado') {
        Logger.log('❌ Acesso rejeitado');
        return {
          sucesso: false,
          status: 'rejeitado',
          mensagem: 'Seu acesso foi negado. Entre em contato com o administrador.'
        };
      }
      
      if (status === 'Aprovado') {
        // ✅ APROVADO - Verificar se já tem planilha
        if (!planilhaId || planilhaId === '-' || planilhaId === '') {
          // Criar planilha agora
          Logger.log('✅ Aprovado mas sem planilha - criando agora...');
          const novaPlanilha = SpreadsheetApp.create("Meus Eventos (Sistema RSVP)");
          const novoId = novaPlanilha.getId();
          
          configurarPlanilhaNova(novaPlanilha);
          
          // Adicionar usuário como editor
          novaPlanilha.addEditor(emailLimpo);
          Logger.log('✅ Usuário adicionado como editor');
          
          // Atualizar planilha de controle
          sheetUsuarios.getRange(i + 1, 3).setValue(novoId);
          sheetUsuarios.getRange(i + 1, 5).setValue(new Date());
          sheetUsuarios.getRange(i + 1, 6).setValue(new Date());
          
          // Enviar email para admin transferir propriedade
          enviarEmailTransferencia(emailLimpo, novoId);
          
          Logger.log('✅ Planilha criada: ' + novoId);
          
          return {
            sucesso: true,
            novoUsuario: true,
            planilhaId: novoId,
            email: emailLimpo,
            status: 'aprovado'
          };
        }
        
        // Já tem planilha - atualizar último acesso
        const agora = new Date();
        sheetUsuarios.getRange(i + 1, 6).setValue(agora);
        
        Logger.log('✅ Usuário aprovado - Planilha: ' + planilhaId);
        
        return {
          sucesso: true,
          novoUsuario: false,
          planilhaId: planilhaId,
          email: emailLimpo,
          status: 'aprovado'
        };
      }
    }
  }
  
  // 🆕 PASSO 3: NOVO USUÁRIO - Registrar como PENDENTE
  Logger.log('🆕 Novo usuário - registrando como PENDENTE');
  
  const agora = new Date();
  sheetUsuarios.appendRow([
    emailLimpo,
    'Pendente',
    '-',
    agora,
    '-',
    agora
  ]);
  
  // Colorir linha como amarelo (pendente)
  const ultimaLinha = sheetUsuarios.getLastRow();
  sheetUsuarios.getRange(ultimaLinha, 1, 1, 6).setBackground("#fff3cd");
  
  Logger.log('✅ Usuário registrado como PENDENTE');
  
  // 📧 ENVIAR EMAIL PARA ADMIN
  enviarEmailNovoUsuario(emailLimpo);
  
  return {
    sucesso: false,
    status: 'pendente',
    mensagem: 'Solicitação enviada! Aguarde a aprovação do administrador. Você receberá um email quando for aprovado.',
    email: emailLimpo
  };
}

function enviarEmailNovoUsuario(emailUsuario) {
  if (!EMAIL_ADMIN || EMAIL_ADMIN === "seu.email@gmail.com") {
    Logger.log('⚠️ EMAIL_ADMIN não configurado - pulando notificação');
    return;
  }
  
  try {
    const urlPlanilha = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_CONTROLE}/edit`;
    
    const assunto = "🔔 Nova solicitação de acesso - Sistema RSVP";
    
    const corpo = `
Olá,

Um novo usuário solicitou acesso ao Sistema RSVP:

👤 Email: ${emailUsuario}
📅 Data: ${new Date().toLocaleString('pt-BR')}

Para aprovar ou rejeitar o acesso:

1. Abra a Planilha de Controle:
   ${urlPlanilha}

2. Localize a linha do usuário

3. Altere a coluna "Status":
   - "Aprovado" → Usuário poderá fazer login
   - "Rejeitado" → Usuário será bloqueado

Após aprovar, o usuário fará login novamente e você receberá outro email com o link da planilha dele para transferir a propriedade.

---
Sistema RSVP v16.5.1
Notificação Automática
    `.trim();
    
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: assunto,
      body: corpo
    });
    
    Logger.log('📧 Email de notificação enviado para admin');
    
  } catch (erro) {
    Logger.log('⚠️ Erro ao enviar email: ' + erro);
  }
}

function enviarEmailTransferencia(emailUsuario, planilhaId) {
  if (!EMAIL_ADMIN || EMAIL_ADMIN === "seu.email@gmail.com") {
    Logger.log('⚠️ EMAIL_ADMIN não configurado - pulando notificação');
    return;
  }
  
  try {
    const urlPlanilha = `https://docs.google.com/spreadsheets/d/${planilhaId}/edit`;
    
    const assunto = "✅ Planilha criada - Transferir propriedade para " + emailUsuario;
    
    const corpo = `
Olá,

A planilha foi criada com sucesso para o usuário aprovado:

👤 Usuário: ${emailUsuario}
📊 Planilha: Meus Eventos (Sistema RSVP)
🔗 Link direto: ${urlPlanilha}

O usuário já foi adicionado como EDITOR e pode usar o sistema normalmente.

Para transferir a PROPRIEDADE da planilha:

1. Abra a planilha (clique no link acima)

2. Clique no botão "Compartilhar" (canto superior direito)

3. Ao lado do email "${emailUsuario}", clique no ícone ⋮ (três pontinhos)

4. Selecione "Transferir propriedade"

5. Confirme a transferência

Após a transferência:
- O usuário será o dono da planilha
- A planilha aparecerá no Drive dele
- Você perderá acesso à planilha (correto!)

Observação: Se você tiver conta Gmail gratuita, pode ser que a opção "Transferir propriedade" não apareça. Neste caso, deixe o usuário como Editor - ele terá controle total mesmo assim.

---
Sistema RSVP v16.5.1
Notificação Automática
    `.trim();
    
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: assunto,
      body: corpo
    });
    
    Logger.log('📧 Email de transferência enviado para admin');
    
  } catch (erro) {
    Logger.log('⚠️ Erro ao enviar email: ' + erro);
  }
}

function configurarPlanilhaNova(ss) {
  // Aba Exemplo
  const sheetExemplo = ss.getSheets()[0];
  sheetExemplo.setName("Exemplo");
  sheetExemplo.clear();
  sheetExemplo.appendRow(["Nome", "Email", "Telefone", "Status"]);
  sheetExemplo.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f4f6");
  sheetExemplo.setColumnWidth(1, 200);
  sheetExemplo.setColumnWidth(2, 250);
  
  // Aba Templates
  const sheetTemplates = ss.insertSheet("Templates");
  sheetTemplates.appendRow(["Nome Template", "Colunas (JSON)", "Email Assunto", "Email Mensagem", "Data Criação", "Vezes Usado"]);
  sheetTemplates.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#dbeafe");
  sheetTemplates.setFrozenRows(1);
  sheetTemplates.setColumnWidth(2, 300);
  
  // Aba Banco de Nomes
  const sheetBanco = ss.insertSheet("Banco_Nomes");
  sheetBanco.appendRow(["Nome Completo", "Email", "Telefone", "Eventos Participou", "Último Evento", "Data Última Participação"]);
  sheetBanco.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#dcfce7");
  sheetBanco.setFrozenRows(1);
  sheetBanco.setColumnWidth(1, 200);
  sheetBanco.setColumnWidth(4, 300);
  
  Logger.log('✅ Planilha configurada com 3 abas');
}

function getPlanilhaUsuario(email) {
  if (!email) {
    throw new Error("Email não fornecido.");
  }
  
  if (!ID_PLANILHA_CONTROLE || ID_PLANILHA_CONTROLE === "") {
    throw new Error("Sistema não configurado. Faça login primeiro para criar a planilha de controle.");
  }
  
  const emailLimpo = garantirString(email).toLowerCase().trim();
  
  const ssControle = SpreadsheetApp.openById(ID_PLANILHA_CONTROLE);
  const sheetUsuarios = ssControle.getSheetByName("Usuarios");
  
  if (!sheetUsuarios) {
    throw new Error("Sistema não inicializado.");
  }
  
  const data = sheetUsuarios.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && garantirString(data[i][0]).toLowerCase() === emailLimpo) {
      const status = garantirString(data[i][1]);
      const planilhaId = garantirString(data[i][2]);
      
      // Verificar status
      if (status !== 'Aprovado') {
        throw new Error("Acesso não autorizado.");
      }
      
      if (!planilhaId || planilhaId === '-' || planilhaId === '') {
        throw new Error("Planilha não configurada. Faça login novamente.");
      }
      
      Logger.log('📊 Abrindo planilha: ' + planilhaId);
      return SpreadsheetApp.openById(planilhaId);
    }
  }
  
  throw new Error("Usuário não encontrado. Faça login novamente.");
}

function getSheetByNameSafe(ss, nome) {
  if (!nome || nome === 'undefined' || nome === 'null' || nome === '') {
    Logger.log('⚠️ Nome de aba inválido');
    return null;
  }
  return ss.getSheetByName(garantirString(nome).trim());
}

// =======================================================================
// FUNÇÕES DE EVENTOS (Com correções de .split())
// =======================================================================

function listarEventos(email) {
  const ss = getPlanilhaUsuario(email);
  const abasSistema = ['Dashboard', 'Templates', 'Banco_Nomes', 'Exemplo'];
  return ss.getSheets()
    .filter(s => !abasSistema.includes(s.getName()))
    .map(s => ({ nome: s.getName(), id: s.getSheetId() }));
}

function criarNovoEvento(email, nome, tipo, dadosRaw) {
  const ss = getPlanilhaUsuario(email);
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
    if (tipo === 'Basico') colunas = ["Nome", "Email", "Telefone", "Confirmado"];
    else if (tipo === 'Casamento') colunas = ["Nome", "Email", "Telefone", "Confirmado", "Acompanhantes", "Mesa", "Restrição Alimentar", "Mensagem"];
    else if (tipo === 'Corporativo') colunas = ["Nome", "Email", "Telefone", "Empresa", "Cargo", "Confirmado", "Workshop Escolhido"];
    else if (tipo === 'Infantil') colunas = ["Nome da Criança", "Idade", "Nome do Responsável", "Email", "Telefone", "Confirmado", "Alergias", "Mensagem"];
    else if (tipo === 'Formatura') colunas = ["Nome", "Email", "Telefone", "Curso", "Turma", "Confirmado", "Qtd Convites", "Mesa"];
    else if (tipo === 'Workshop') colunas = ["Nome", "Email", "Telefone", "Confirmado", "Nível de Experiência", "Tópicos de Interesse", "Precisa Material"];
    else if (tipo === 'Jantar') colunas = ["Nome", "Email", "Telefone", "Confirmado", "Num. de Pessoas", "Horário Preferido", "Restrições Alimentares"];
    else colunas = ["Nome", "Email", "Telefone", "Confirmado"];
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

function obterDadosEvento(email, nome) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nome);
  if (!sheet) throw new Error("Evento não encontrado");
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return { headers: [], rows: [] };
  return { headers: data[0], rows: data.slice(1) };
}

function renomearEvento(email, nomeAntigo, nomeNovo) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeAntigo);
  if (!sheet) throw new Error("Evento não encontrado");
  if (getSheetByNameSafe(ss, nomeNovo)) throw new Error("Nome já existe");
  sheet.setName(nomeNovo.trim());
  return { sucesso: true };
}

function adicionarColuna(email, nomeEvento, novaColuna) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.includes(novaColuna)) throw new Error("Coluna já existe");
  sheet.getRange(1, lastCol + 1).setValue(novaColuna).setFontWeight("bold").setBackground("#f3f4f6");
  return { sucesso: true };
}

function removerColuna(email, nomeEvento, nomeColuna) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(nomeColuna);
  if (index === -1) throw new Error("Coluna não encontrada");
  sheet.deleteColumn(index + 1);
  return { sucesso: true };
}

function adicionarConvidado(email, nomeEvento, dados) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  sheet.appendRow(dados);
  return { sucesso: true };
}

function atualizarConvidado(email, nomeEvento, linhaReal, novosDados) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  sheet.getRange(linhaReal, 1, 1, novosDados.length).setValues([novosDados]);
  return { sucesso: true };
}

function excluirConvidado(email, nomeEvento, linhaReal) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  sheet.deleteRow(linhaReal);
  return { sucesso: true };
}

function importarListaInteligente(email, nomeEvento, matrizDados, temCabecalho) {
  const ss = getPlanilhaUsuario(email);
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

function enviarEmails(email, nomeEvento, indices, assunto, mensagem, linkBase) {
  if (!indices || indices.length === 0) throw new Error("Nenhum convidado selecionado");
  if (indices.length > 50) throw new Error("Limite: 50 por vez");
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const indexEmail = headers.findIndex(h => {
    if (!h) return false;
    const limpo = garantirString(h).trim().toLowerCase();
    return limpo === 'email' || limpo === 'e-mail';
  });
  if (indexEmail === -1) throw new Error("Coluna Email não encontrada");
  const indexNome = headers.findIndex(h => h && garantirString(h).trim().toLowerCase() === 'nome');
  let enviados = 0, erros = 0, detalhesErros = [];
  indices.forEach((index, i) => {
    try {
      if (index < 0 || index >= rows.length) {
        erros++;
        detalhesErros.push(`Índice ${index} inválido`);
        return;
      }
      const row = rows[index];
      const emailDestino = row[indexEmail] ? garantirString(row[indexEmail]).trim() : '';
      const nome = indexNome > -1 ? row[indexNome] : "Convidado";
      if (!emailDestino || !emailDestino.includes("@")) {
        erros++;
        detalhesErros.push(`${nome}: email inválido`);
        return;
      }
      const linkPersonalizado = `${linkBase}?evento=${encodeURIComponent(nomeEvento.trim())}`;
      let corpo = mensagem.replace(/{Nome}/g, nome).replace(/{Link}/g, linkPersonalizado);
      let assuntoFinal = assunto.replace(/{Nome}/g, nome).replace(/{Evento}/g, nomeEvento);
      MailApp.sendEmail({ to: emailDestino, subject: assuntoFinal, body: corpo });
      enviados++;
      if (i < indices.length - 1) Utilities.sleep(1000);
    } catch (e) {
      erros++;
      const nomeErro = indexNome > -1 && rows[index] ? rows[index][indexNome] : `Índice ${index}`;
      detalhesErros.push(`${nomeErro}: ${e.message}`);
    }
  });
  return { sucesso: true, enviados: enviados, erros: erros, detalhes: detalhesErros.length > 0 ? detalhesErros : null };
}

function buscarConvidado(email, nomeEvento, nomeBusca) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sucesso: true, encontrado: false };
  const headers = data[0];
  const linhas = data.slice(1);
  const nomeLimpo = nomeBusca.toLowerCase().trim();
  for (let i = 0; i < linhas.length; i++) {
    const row = linhas[i];
    const indexNome = headers.findIndex(h => h && garantirString(h).trim().toLowerCase() === 'nome');
    const valNome = indexNome > -1 ? row[indexNome] : row[0];
    if (valNome && garantirString(valNome).toLowerCase().includes(nomeLimpo)) {
      return { sucesso: true, encontrado: true, nomeCompleto: valNome, linha: i + 2, colunas: headers, dadosAtuais: row };
    }
  }
  return { sucesso: true, encontrado: false };
}

function salvarResposta(email, nomeEvento, linha, respostas) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (const [coluna, valor] of Object.entries(respostas)) {
    const colIndex = headers.indexOf(coluna);
    if (colIndex > -1) sheet.getRange(linha, colIndex + 1).setValue(valor);
  }
  return { sucesso: true };
}

function excluirEvento(email, nomeEvento) {
  const ss = getPlanilhaUsuario(email);
  const sheet = getSheetByNameSafe(ss, nomeEvento);
  if (!sheet) throw new Error("Evento não encontrado");
  const sheetBanco = ss.getSheetByName("Banco_Nomes");
  if (sheetBanco) {
    const data = sheetBanco.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // ✅ CORREÇÃO: Garantir string antes de split
      const eventosAtuais = garantirString(data[i][3]);
      const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
      const eventosAtualizados = eventosArray.filter(e => e !== nomeEvento);
      if (eventosAtualizados.length !== eventosArray.length) {
        sheetBanco.getRange(i + 1, 4).setValue(eventosAtualizados.join(';'));
        if (data[i][4] === nomeEvento) {
          const novoUltimo = eventosAtualizados.length > 0 ? eventosAtualizados[eventosAtualizados.length - 1] : "Nenhum";
          sheetBanco.getRange(i + 1, 5).setValue(novoUltimo);
        }
      }
    }
  }
  ss.deleteSheet(sheet);
  return { sucesso: true };
}

// =======================================================================
// TEMPLATES E BANCO DE NOMES (COM CORREÇÕES)
// =======================================================================

function inicializarAbaTemplates(email) {
  const ss = getPlanilhaUsuario(email);
  let sheetTemplates = ss.getSheetByName("Templates");
  if (!sheetTemplates) {
    sheetTemplates = ss.insertSheet("Templates");
    sheetTemplates.appendRow(["Nome Template", "Colunas (JSON)", "Email Assunto", "Email Mensagem", "Data Criação", "Vezes Usado"]);
    sheetTemplates.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#dbeafe");
    sheetTemplates.setFrozenRows(1);
  }
  return sheetTemplates;
}

function salvarTemplate(email, nomeTemplate, colunas, emailAssunto, emailMensagem) {
  const sheetTemplates = inicializarAbaTemplates(email);
  const nomeLimpo = nomeTemplate.trim();
  if (!nomeLimpo) throw new Error("Nome vazio");
  const data = sheetTemplates.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeLimpo) throw new Error("Template já existe");
  }
  sheetTemplates.appendRow([nomeLimpo, JSON.stringify(colunas), emailAssunto || "", emailMensagem || "", new Date().toLocaleDateString('pt-BR'), 0]);
  return { sucesso: true, nome: nomeLimpo };
}

function listarTemplates(email) {
  const sheetTemplates = inicializarAbaTemplates(email);
  const data = sheetTemplates.getDataRange().getValues();
  if (data.length < 2) return { templates: [] };
  const templates = [];
  for (let i = 1; i < data.length; i++) {
    templates.push({ nome: data[i][0], colunas: JSON.parse(data[i][1]), emailAssunto: data[i][2], emailMensagem: data[i][3], dataCriacao: data[i][4], vezesUsado: data[i][5] });
  }
  return { templates: templates };
}

function excluirTemplate(email, nomeTemplate) {
  const sheetTemplates = inicializarAbaTemplates(email);
  const data = sheetTemplates.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeTemplate) {
      sheetTemplates.deleteRow(i + 1);
      return { sucesso: true };
    }
  }
  throw new Error("Template não encontrado");
}

function criarEventoDeTemplate(email, nomeEvento, nomeTemplate, usarEmail) {
  const sheetTemplates = inicializarAbaTemplates(email);
  const data = sheetTemplates.getDataRange().getValues();
  let templateEncontrado = null, linhaTemplate = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === nomeTemplate) {
      templateEncontrado = { colunas: JSON.parse(data[i][1]), emailAssunto: data[i][2], emailMensagem: data[i][3] };
      linhaTemplate = i + 1;
      break;
    }
  }
  if (!templateEncontrado) throw new Error("Template não encontrado");
  const ss = getPlanilhaUsuario(email);
  const nomeLimpo = nomeEvento.trim();
  if (ss.getSheetByName(nomeLimpo)) throw new Error("Evento já existe");
  const sheet = ss.insertSheet(nomeLimpo);
  sheet.appendRow(templateEncontrado.colunas);
  sheet.getRange(1, 1, 1, templateEncontrado.colunas.length).setFontWeight("bold").setBackground("#f3f4f6");
  sheetTemplates.getRange(linhaTemplate, 6).setValue((data[linhaTemplate - 1][5] || 0) + 1);
  const resultado = { sucesso: true, nome: nomeLimpo, colunas: templateEncontrado.colunas };
  if (usarEmail) {
    resultado.emailAssunto = templateEncontrado.emailAssunto;
    resultado.emailMensagem = templateEncontrado.emailMensagem;
  }
  return resultado;
}

function inicializarBancoNomes(email) {
  const ss = getPlanilhaUsuario(email);
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

function buscarNoBanco(email, termo) {
  const sheetBanco = inicializarBancoNomes(email);
  const data = sheetBanco.getDataRange().getValues();
  if (data.length < 2) return { encontrado: false };
  const termoLimpo = termo.toLowerCase().trim();
  const resultados = [];
  for (let i = 1; i < data.length; i++) {
    const nome = data[i][0] ? garantirString(data[i][0]).toLowerCase() : "";
    if (nome.includes(termoLimpo)) {
      // ✅ CORREÇÃO: Garantir string antes de split
      const eventosAtuais = garantirString(data[i][3]);
      const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
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
  if (resultados.length === 0) return { encontrado: false };
  return { encontrado: true, resultados: resultados };
}

function adicionarAoBancoNomes(email, nome, emailContato, telefone, eventoAtual) {
  const sheetBanco = inicializarBancoNomes(email);
  const data = sheetBanco.getDataRange().getValues();
  const nomeLimpo = nome.trim();
  const dataAtual = new Date().toLocaleDateString('pt-BR');
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && garantirString(data[i][0]).toLowerCase() === nomeLimpo.toLowerCase()) {
      // ✅ CORREÇÃO: Garantir string antes de split
      const eventosAtuais = garantirString(data[i][3]);
      const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
      if (!eventosArray.includes(eventoAtual)) eventosArray.push(eventoAtual);
      sheetBanco.getRange(i + 1, 2).setValue(emailContato || data[i][1]);
      sheetBanco.getRange(i + 1, 3).setValue(telefone || data[i][2]);
      sheetBanco.getRange(i + 1, 4).setValue(eventosArray.join(';'));
      sheetBanco.getRange(i + 1, 5).setValue(eventoAtual);
      sheetBanco.getRange(i + 1, 6).setValue(dataAtual);
      return { sucesso: true, atualizado: true };
    }
  }
  sheetBanco.appendRow([nomeLimpo, emailContato || "", telefone || "", eventoAtual, eventoAtual, dataAtual]);
  return { sucesso: true, atualizado: false };
}

function listarBancoNomes(email) {
  const sheetBanco = inicializarBancoNomes(email);
  const data = sheetBanco.getDataRange().getValues();
  if (data.length < 2) return { nomes: [] };
  const nomes = [];
  for (let i = 1; i < data.length; i++) {
    // ✅ CORREÇÃO: Garantir string antes de split
    const eventosAtuais = garantirString(data[i][3]);
    const eventosArray = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
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

function editarNoBanco(email, nomeAntigo, nomeNovo, emailContato, telefone, propagarEventos) {
  const sheetBanco = inicializarBancoNomes(email);
  const data = sheetBanco.getDataRange().getValues();
  let linhaPessoa = -1, eventosParticipou = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && garantirString(data[i][0]).toLowerCase() === nomeAntigo.toLowerCase()) {
      linhaPessoa = i + 1;
      // ✅ CORREÇÃO: Garantir string antes de split
      const eventosAtuais = garantirString(data[i][3]);
      eventosParticipou = eventosAtuais ? eventosAtuais.split(';').filter(e => e.trim()) : [];
      break;
    }
  }
  if (linhaPessoa === -1) throw new Error("Nome não encontrado");
  sheetBanco.getRange(linhaPessoa, 1).setValue(nomeNovo.trim());
  sheetBanco.getRange(linhaPessoa, 2).setValue(emailContato || "");
  sheetBanco.getRange(linhaPessoa, 3).setValue(telefone || "");
  if (propagarEventos && nomeAntigo !== nomeNovo) {
    const ss = getPlanilhaUsuario(email);
    let eventosAtualizados = 0;
    eventosParticipou.forEach(nomeEvento => {
      const sheet = getSheetByNameSafe(ss, nomeEvento);
      if (!sheet) return;
      const dataEvento = sheet.getDataRange().getValues();
      const headers = dataEvento[0];
      const indexNome = headers.findIndex(h => h && garantirString(h).toLowerCase().includes('nome'));
      if (indexNome === -1) return;
      for (let i = 1; i < dataEvento.length; i++) {
        if (dataEvento[i][indexNome] && garantirString(dataEvento[i][indexNome]).toLowerCase() === nomeAntigo.toLowerCase()) {
          sheet.getRange(i + 1, indexNome + 1).setValue(nomeNovo);
          eventosAtualizados++;
        }
      }
    });
    return { sucesso: true, propagado: true, eventosAtualizados: eventosAtualizados };
  }
  return { sucesso: true, propagado: false };
}

function excluirDoBanco(email, nome) {
  const sheetBanco = inicializarBancoNomes(email);
  const data = sheetBanco.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && garantirString(data[i][0]).toLowerCase() === nome.toLowerCase()) {
      sheetBanco.deleteRow(i + 1);
      return { sucesso: true };
    }
  }
  throw new Error("Nome não encontrado");
}

function migrarEventosParaBanco(email) {
  const ss = getPlanilhaUsuario(email);
  const sheetBanco = inicializarBancoNomes(email);
  const abasSistema = ['Dashboard', 'Templates', 'Banco_Nomes', 'Exemplo'];
  let totalAdicionados = 0, totalAtualizados = 0;
  const eventos = ss.getSheets().filter(s => !abasSistema.includes(s.getName())).map(s => s.getName());
  eventos.forEach(nomeEvento => {
    const sheet = ss.getSheetByName(nomeEvento);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const indexNome = headers.indexOf('Nome') >= 0 ? headers.indexOf('Nome') : 0;
    const indexEmail = headers.findIndex(h => h && garantirString(h).trim().toLowerCase().match(/^e?-?mail$/));
    const indexTelefone = headers.indexOf('Telefone');
    for (let i = 1; i < data.length; i++) {
      const nome = data[i][indexNome];
      if (!nome || !garantirString(nome).trim()) continue;
      try {
        const resultado = adicionarAoBancoNomes(
          email, 
          garantirString(nome).trim(), 
          indexEmail >= 0 ? garantirString(data[i][indexEmail] || "").trim() : '', 
          indexTelefone >= 0 ? garantirString(data[i][indexTelefone] || "").trim() : '', 
          nomeEvento
        );
        if (resultado.atualizado) totalAtualizados++; else totalAdicionados++;
      } catch (e) { }
    }
  });
  return { sucesso: true, adicionados: totalAdicionados, atualizados: totalAtualizados, total: totalAdicionados + totalAtualizados, eventos: eventos.length };
}
