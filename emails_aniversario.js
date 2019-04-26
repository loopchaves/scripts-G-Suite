//-----------------------------------------------------------------------------------------------------
// Envia emails
function enviaEmails()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var row = 1; row < data.length; row++)
  {
    if (data[row][3] === "Agendado")
    {
      var drafts = GmailApp.getDraftMessages();
      if (drafts.length > 0) // Se tiver algum rascunho continua verificação
      {
        var messageHtml;
        var testDate = new Date(); // Para salvar a data anterior
        var anexos;
        var cont = 0; // Contador de rascunhos com o título "Feliz Aniversário"
        
        // Inicia no primeiro rascunho e vai de um por um
        for (var i = 0; i < drafts.length; i++)
        {
          // Verifica se tem algum racunho com assunto "Feliz Aniversário"
          if (drafts[i].getSubject() === "Feliz Aniversário")
          {
            // Se tiver, vai comparar a data com a anterior e se for maior salva
            if (cont === 0) // Primeiro
            {
              messageHtml = drafts[i].getBody();
              testDate = drafts[i].getDate();
              anexos = drafts[i].getAttachments();
              cont++;
            }
            else if (drafts[i].getDate() > testDate) // Se tiver mais de um email, testa se a data é maior que a anterior
            {
              messageHtml = drafts[i].getBody();
              testDate = drafts[i].getDate();
              anexos = drafts[i].getAttachments();
              cont++;
            }
          }
        }
        
        // Aqui que acontece a mágica
        var anexo;
        for (var k in anexos)
        {
          anexo = anexos[k];
        }
        var blobImage = anexo.copyBlob().setName("imagem");
        
        // Pega redimensionamento da imagem
        var width = messageHtml.match(/width="[^"]+"/g);
        var height = messageHtml.match(/height="[^"]+"/g);
        
        // Substitui as tags na mensagem
        var nome = data[row][0];
        messageHtml = messageHtml.replace(/NOME/g, nome);
        messageHtml = messageHtml.replace(/<img[^>]+>/g, "<img src='cid:imagem' " + width + " " + height + ">");
      } else {
        // Se não tiver nenhum rascunho finaliza
        messageHtml = 0;
        sheet.getRange(row, 4).setValue("Erro 01");
        sheet.getRange(row, 5).setValue(messageHtml);
        sheet.getRange(row, 6).setValue(blobImage);
      }
      //--------------------------------------------------------------------
      // Envia o modelo do rascunho para o email da planilha
      var schedule = data[row][2];
      var time = new Date().getTime();
      if ((schedule != "") && (schedule.getTime() <= time) && (messageHtml != 0))
      {
        var para = data[row][1];
        MailApp.sendEmail({
          to: para,
          subject: "Feliz aniversário",
          htmlBody: messageHtml,
          inlineImages: 
          {
            imagem: blobImage
          }
        });
        sheet.getRange("D" + (row + 1)).setValue("Enviado");
      } else {
        // Se deu erro ao enviar email
        sheet.getRange(row, 4).setValue("Erro 02");
        sheet.getRange(row, 5).setValue(para);
        sheet.getRange(row, 6).setValue(messageHtml);
        sheet.getRange(row, 7).setValue(blobImage);
      }
    }
  }
} // FIM
//-----------------------------------------------------------------------------------------------------
// Cria o agendamento dos emails baseado na planilha
function agendar() 
{
  // Apaga os Triggers enviaEmails
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++)
  {
    if (triggers[i].getHandlerFunction() === "enviaEmails")
    {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Cria novamente os Triggers
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var time = new Date().getTime();
  var code = [];
  for (var row in data)
  {
    if (row != 0)
    {
      var schedule = data[row][2];
      if (schedule !== "")
      {
        if (schedule.getTime() > time)
        {
          ScriptApp.newTrigger("enviaEmails")
          .timeBased()
          .at(schedule)
          .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
          .create();
          code.push("Agendado");
        } else {
          code.push("A data está no passado");
        }
      } else {
        code.push("Não agendado");
      }
    }
  }
  // Adiciona o code na planilha
  for (var i = 0; i < code.length; i++)
  {
    sheet.getRange("D" + (i + 2)).setValue(code[i]);
  }
} // FIM
