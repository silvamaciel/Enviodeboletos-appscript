# Script de Envio de Boletos por E-mail

Este script em Google Apps Script permite o envio automatizado de boletos por e-mail para destinatários listados em uma planilha do Google Sheets. Ele busca os arquivos de boletos em uma pasta do Google Drive e os envia para cada destinatário de acordo com as informações da planilha.

## Funcionalidade do Script

1. Acessa uma planilha do Google Sheets chamada **"ENVIODEBOLETOS"**.
2. Coleta os dados de cada linha da planilha, como nome, unidade, e-mail, empreendimento, e outros detalhes.
3. Verifica a presença de uma subpasta no Google Drive com o nome do mês e ano atuais.
4. Dentro da subpasta, encontra o arquivo correspondente à unidade especificada.
5. Envia um e-mail para cada destinatário com o boleto anexado, seguindo o modelo de mensagem especificado.

## Estrutura da Planilha

A planilha deve conter as seguintes colunas (na ordem e estrutura esperada):

1. **Código Nome**: Identificação do cliente.
2. **Unidade**: Unidade do cliente.
3. **Nome**: Nome do destinatário.
4. **Empreendimento**: Nome do empreendimento.
5. **E-mail**: Endereço de e-mail do destinatário.
6. **Empresa**: Nome da empresa.
7. **Apartamento**: Número do apartamento.
8. **Bairro**: Bairro onde o empreendimento está localizado.

A primeira linha da planilha deve ser o cabeçalho, e os dados começam a partir da segunda linha.

## Configuração do Google Drive

1. Crie uma pasta principal no Google Drive para armazenar os arquivos dos boletos.
2. Dentro desta pasta, crie subpastas nomeadas no formato `MM/YYYY`, correspondendo ao mês e ano dos boletos a serem enviados.
3. Substitua o valor de `pastaId` no código com o ID da sua pasta principal do Google Drive.

## Configuração do Script

1. Abra o **Google Apps Script** a partir da planilha do Google Sheets onde o script será executado.
2. Cole o código no editor de script.
3. No campo `pastaId`, substitua pelo ID da sua pasta principal do Google Drive onde os arquivos estão armazenados.
4. Atualize a URL da imagem da assinatura se desejar incluir uma diferente.

## Exemplo de Uso

1. Preencha a planilha com os dados dos destinatários e informações dos boletos.
2. Execute o script clicando em **Executar > Enviar E-mails** no Google Apps Script.
3. O script enviará e-mails para cada destinatário com o boleto referente ao empreendimento, desde que o arquivo do boleto correspondente seja encontrado na subpasta.

## Estrutura do Código

- `SpreadsheetApp.getActiveSpreadsheet()` e `getSheetByName`: Acessa a planilha ativa e a aba com o nome específico.
- `DriveApp.getFolderById(pastaId)`: Acessa a pasta principal no Google Drive pelo ID.
- Laço `for`: Itera sobre cada linha de dados na planilha.
- `MailApp.sendEmail`: Envia o e-mail com o arquivo de boleto anexado.

## Observações

- **Subpasta do Mês e Ano**: Certifique-se de que o nome da subpasta esteja no formato `MM/YYYY` e que corresponda ao mês e ano atuais.
- **Formato do Arquivo**: O script assume que o arquivo é um PDF e o envia como tal.
- **Erros e Logs**: O script registra no console do Google Apps Script as informações sobre cada e-mail enviado e notifica se algum arquivo não foi encontrado.
