# Projeto de Envio Automático de Boletos por E-mail

Este projeto utiliza o Google Apps Script para realizar o envio automatizado de boletos por e-mail. Com base nos dados de uma planilha do Google Sheets, o script busca arquivos de boletos em uma pasta do Google Drive e os envia para os destinatários indicados na planilha.

## Funcionalidades

- **Busca dados na planilha do Google Sheets** chamada **"ENVIODEBOLETOS"**.
- Verifica se há uma subpasta específica no Google Drive, baseada no **mês e ano atuais**, onde os arquivos de boletos estão armazenados.
- Procura o arquivo correspondente à unidade do destinatário e envia por e-mail com um modelo de mensagem específico.
- Registra os e-mails enviados e notifica caso algum arquivo não seja encontrado.

---

## Configuração do Projeto

### Passo 1: Criação da Planilha no Google Sheets

1. **Acesse o Google Sheets**: Abra [Google Sheets](https://sheets.google.com).
2. **Crie uma nova planilha**: Clique em **"Em branco"** para criar uma nova planilha.
3. **Renomeie a planilha**: Nomeie a planilha como **"Envio de Boletos"** ou outro nome que preferir.
4. **Adicione as Colunas**: Na primeira linha, insira as seguintes colunas exatamente nesta ordem:

   - A1: `Código Nome`
   - B1: `Unidade`
   - C1: `Nome`
   - D1: `Empreendimento`
   - E1: `E-mail`
   - F1: `Empresa`
   - G1: `Apartamento`
   - H1: `Bairro`

   A planilha deve ter uma estrutura como esta:

   | Código Nome | Unidade | Nome | Empreendimento | E-mail             | Empresa   | Apartamento | Bairro  |
   |-------------|---------|------|----------------|--------------------|-----------|-------------|---------|
   | 001234      | UN-001  | João | Empreendimento X | email@exemplo.com | Empresa A | 101         | Centro  |

5. **Renomeie a aba**: Clique na aba padrão "Planilha1" e renomeie para **"ENVIODEBOLETOS"** (o script busca a aba por este nome).

---

### Passo 2: Configuração do Google Apps Script

1. **Abrindo o Google Apps Script**:
   - Na planilha criada, vá até **"Extensões"** no menu superior e selecione **"Apps Script"**. Isso abrirá o editor do Google Apps Script em uma nova aba.
   
2. **Inserindo o Código**:
   - Apague qualquer conteúdo padrão que esteja no editor.
   - Cole o código do script completo na área de código.

3. **Configuração do ID da Pasta do Google Drive**:
   - Encontre a linha `var pastaId = '1TSf-r248xIVbd-vnYa-0brwR65WECS1-';`.
   - Substitua `'1TSf-r248xIVbd-vnYa-0brwR65WECS1-'` pelo **ID da sua pasta principal** do Google Drive onde os arquivos de boletos estão armazenados. 
   - Para obter o ID da pasta:
     - Acesse a pasta desejada no Google Drive.
     - O ID aparece na URL, logo após `https://drive.google.com/drive/folders/`.

4. **Configuração da URL da Assinatura** (opcional):
   - Se desejar, substitua a URL de `assinaturaImagemUrl` para uma imagem de assinatura personalizada.

---

### Passo 3: Autorizar o Script

1. **Executar o Script pela Primeira Vez**:
   - No editor do Apps Script, clique em **"Executar"** (ícone de triângulo) para rodar o script.
   - O Google solicitará que você **autorize o script** para acessar o Google Sheets e o Google Drive.
   - Clique em **"Revisar permissões"**, escolha sua conta do Google e clique em **"Permitir"** para conceder as permissões necessárias.

---

### Passo 4: Estrutura de Pastas no Google Drive

1. **Criação da Pasta Principal**:
   - No Google Drive, crie uma pasta principal onde serão armazenados os boletos. Nomeie essa pasta como preferir.

2. **Subpastas Mensais**:
   - Dentro da pasta principal, crie subpastas no formato `MM/YYYY` (mês/ano), como `01/2024`, `02/2024`, etc.
   - Cada subpasta deve conter os arquivos de boletos referentes ao mês e ano especificados. O script procura o boleto correspondente à unidade do cliente usando as primeiras seis letras do valor da coluna `Unidade`.

---

### Passo 5: Executar o Script e Enviar E-mails

1. **Preencher a Planilha**:
   - Adicione os dados dos destinatários na planilha, incluindo as informações da unidade, nome, e-mail e outros.

2. **Executar o Envio**:
   - No editor do Google Apps Script, clique em **"Executar"** para rodar o script.
   - O script enviará automaticamente um e-mail para cada destinatário, incluindo o boleto correspondente como anexo PDF, caso o arquivo seja encontrado na subpasta especificada.

---

## Estrutura do Código

- `SpreadsheetApp.getActiveSpreadsheet()` e `getSheetByName`: Acessa a planilha ativa e a aba com o nome específico.
- `DriveApp.getFolderById(pastaId)`: Acessa a pasta principal no Google Drive pelo ID.
- Laço `for`: Itera sobre cada linha de dados na planilha.
- `MailApp.sendEmail`: Envia o e-mail com o boleto anexado, utilizando as informações do destinatário e o conteúdo do corpo da mensagem.

## Observações Finais

- **Formato das Subpastas**: Certifique-se de que as subpastas no Google Drive estejam no formato `MM/YYYY`.
- **Formato do Arquivo**: O script considera que os arquivos de boletos são PDFs e os envia como tal.
- **Erros e Logs**: O script registra as atividades e erros no console do Google Apps Script, indicando se algum e-mail foi enviado ou se algum arquivo não foi encontrado para determinada unidade.
