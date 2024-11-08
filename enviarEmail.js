function enviarEmails() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = planilha.getSheetByName('ENVIODEBOLETOS'); // Nome da planilha
  var range = sheet.getDataRange();
  var dados = range.getValues();

  var date = new Date();
  var mesAno = `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
  
  // ID da pasta no Google Drive onde os arquivos estão localizados
  var pastaId = '1TSf-r248xIVbd-vnYa-0brwR65WECS1-'; // Substitua pelo ID da sua pasta
  var pastaPrincipal = DriveApp.getFolderById(pastaId); // Acessa a pasta pelo ID
  var subpastas = pastaPrincipal.getFoldersByName(mesAno);

  if (!subpastas.hasNext()) {
    Logger.log('Subpasta não encontrada.');
    return; // Encerra o script se a subpasta não for encontrada
  }

  var subpasta = subpastas.next();

  // Faz uma busca nos valores que estão na planilha atual e guarda em variáveis
  for (var i = 1; i < dados.length; i++) { // Inicia em 1 para ignorar a linha de cabeçalho
    var codigonome = dados[i][0];
    var unidade = dados[i][1]; // Coluna UNIDADE
    var nome = dados[i][2]; // Coluna Nome
    var empreendimento = dados[i][3]; // Coluna Empreendimento
    var email = dados[i][4]; // Coluna E-mail
    var empresa = dados[i][5]; // Coluna Empresa
    var apartamento = dados[i][6]; // Coluna Apartamento
    var bairro = dados[i][7]; // Coluna Bairro
    
    // Buscar todos os arquivos na subpasta
    var arquivos = subpasta.getFiles(); 
    var arquivoEncontrado = null;

    // Carregar a imagem da assinatura do Google Drive
    var assinaturaImagemUrl = 'https://ci3.googleusercontent.com/mail-sig/AIorK4yZio8GhRhWJ1eI4nR0KLKYzThI3dYO-9yP3szPQ_pj2wogFKduojh09wbl3SnLJPOwv4_5Y-5lqt3z';
    
    // Procurar pelo arquivo cujo nome começa com as 6 primeiras letras da coluna 'UNIDADE' na planilha
    while (arquivos.hasNext()) {
      var arquivo = arquivos.next();
      if (arquivo.getName().substring(0, 6) === codigonome.substring(0, 6)) {
        arquivoEncontrado = arquivo;
        break; // Encerra a busca ao encontrar o arquivo correspondente
      }
    }

    // Texto para enviar no corpo do e-mail
    if (arquivoEncontrado) {
      var assunto = `BOLETO: ${empresa} | ${empreendimento}, apartamento n° ${apartamento} | ${mesAno}`;
      var mensagem = `
        <p>Prezado(a) <strong>${nome}</strong>, tudo incrível? Espero que este e-mail lhe encontre bem.</p>
        <p>Segue o seu boleto referente ao empreendimento <strong>${empreendimento}</strong>, Apartamento n° <strong>${apartamento}</strong>, localizado em ${bairro}.</p>
        <p>Em caso de dúvidas ou informações adicionais, não hesite em nos contatar.</p>
        <p>Se você já efetuou o pagamento, por gentileza desconsidere este e-mail.</p>
        <p>Atenciosamente,</p>
        <p><img src="${assinaturaImagemUrl}" alt="Assinatura"></p>
      `;
      
      MailApp.sendEmail({
        to: email,
        subject: assunto,
        htmlBody: mensagem,
        attachments: [arquivoEncontrado.getAs(MimeType.PDF)] // Supondo que o arquivo é um PDF
      });
      
      Logger.log('E-mail enviado para: ' + email);
    } else {
      Logger.log('Arquivo não encontrado para a UNIDADE: ' + unidade);
    }
  }
}
