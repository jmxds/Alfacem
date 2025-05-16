function processarCriacaoUsuarios4() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const abaCriacao = spreadsheet.getSheetByName('Criacao Aluno');
  
  if (!abaCriacao) {
    Logger.log('A aba "Criacao Aluno" não foi encontrada.'); 
    return;
  }
  
  const emailCriador = Session.getActiveUser().getEmail();
  const dados = abaCriacao.getDataRange().getValues();

  const artigos = [
    'de', 'da', 'do', 'dos', 'das', 
    'a', 'o', 'as', 'os', 
    'e', 'em', 'no', 'na', 'nos', 'nas'
  ];

  for (let i = 1; i < dados.length; i++) {
    const nomeCompleto = dados[i][0];
    
    if (nomeCompleto && dados[i][4] !== "Criado") {
      const nomeNormalizado = normalizarNome(nomeCompleto);
      const nomePartes = nomeNormalizado.trim().split(' ').filter(parte => !artigos.includes(parte.toLowerCase()));
      let email = '';
      let mensagemEmail = '';
      let emailCriado = false;
      let ultimoNomePadrao = nomePartes[nomePartes.length - 1];

      // Verifica se já existe um usuário com o mesmo nome
      const usuariosExistentes = AdminDirectory.Users.list({
        domain: 'aluno.alfacem.com.br',
        query: `name='${nomeCompleto}'`
      }).users;

      if (usuariosExistentes && usuariosExistentes.length > 0) {
        abaCriacao.getRange(i + 1, 5).setValue("Falha");
        abaCriacao.getRange(i + 1, 8).setValue(`Já existe um usuário com o nome ${nomeCompleto}`);
        continue;
      }

      // Tenta criar email com diferentes combinações de sobrenomes
      for (let j = nomePartes.length - 1; j > 0; j--) {
        email = `${nomePartes[0]}.${nomePartes[j]}@aluno.alfacem.com.br`.toLowerCase();
        
        if (!verificarEmailExistente(email)) {
          emailCriado = true;
          mensagemEmail = nomePartes[j] !== ultimoNomePadrao ? `Email criado com sobrenome diferente (${nomePartes[j]})` : '';
          break;
        }
      }

      // Se não conseguiu criar email com sobrenomes completos, tenta com iniciais
      if (!emailCriado) {
        for (let j = 1; j < nomePartes.length - 1; j++) {
          const inicialSobrenome = nomePartes[j].charAt(0);
          email = `${nomePartes[0]}.${inicialSobrenome}.${ultimoNomePadrao}@aluno.alfacem.com.br`.toLowerCase();
          
          if (!verificarEmailExistente(email)) {
            emailCriado = true;
            mensagemEmail = `Email criado com inicial de sobrenome (${nomePartes[j]})`;
            break;
          }
        }
      }

      if (!emailCriado) {
        abaCriacao.getRange(i + 1, 5).setValue("Falha");
        abaCriacao.getRange(i + 1, 8).setValue("Não foi possível criar um email único");
        continue;
      }
      
      if (!isValidEmail(email)) {
        abaCriacao.getRange(i + 1, 5).setValue("Falha");
        abaCriacao.getRange(i + 1, 8).setValue("Email inválido");
        continue;
      }

      const senhaPadrao = "Alfacem2025";
      const message = criarUsuario4(nomeCompleto, email, senhaPadrao);
      
      if (message) {
        abaCriacao.getRange(i + 1, 5).setValue("Falha");
        abaCriacao.getRange(i + 1, 8).setValue(message);
      } else {
        const dataAtual = new Date();
        const dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

        abaCriacao.getRange(i + 1, 2).setValue(nomePartes[0]);
        abaCriacao.getRange(i + 1, 3).setValue(ultimoNomePadrao);
        abaCriacao.getRange(i + 1, 4).setValue(email);
        abaCriacao.getRange(i + 1, 5).setValue("Criado");
        abaCriacao.getRange(i + 1, 6).setValue(dataFormatada);
        abaCriacao.getRange(i + 1, 7).setValue(emailCriador);
        
        if (mensagemEmail) {
          abaCriacao.getRange(i + 1, 8).setValue(mensagemEmail);
        }
      }
      
      SpreadsheetApp.flush();
    }
  }
}

// As outras funções (criarUsuario4, verificarEmailExistente, isValidEmail, normalizarNome) permanecem as mesmas

function criarUsuario4(nomeCompleto, email, senhaPadrao) {
  try {
    const user = {
      primaryEmail: email,
      name: {
        fullName: nomeCompleto,
        givenName: nomeCompleto.split(' ')[0],
        familyName: nomeCompleto.split(' ').slice(1).join(' ')
      },
      password: senhaPadrao,
      changePasswordAtNextLogin: true
    };
    AdminDirectory.Users.insert(user);
    return null;
  } catch (e) {
    Logger.log(e);
    return e.message;
  }
}

function verificarEmailExistente(email) {
  try {
    try {
      AdminDirectory.Users.get(email);
      return true; // Email já existe
    } catch (error) {
      // Se der erro de não encontrado, significa que o email não existe
      return false;
    }
  } catch (error) {
    // Erro inesperado
    Logger.log(`Erro ao verificar email ${email}: ${error}`);
    return true; // Considera como existente em caso de erro
  }
}

function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// Função para normalizar o nome removendo acentos e caracteres especiais
function normalizarNome(nome) {
  return nome
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s]/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}
