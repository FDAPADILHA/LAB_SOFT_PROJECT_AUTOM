function doGet(e) {
  if ( e.parameter.email ){
    var page = HtmlService.createHtmlOutputFromFile("verificar cupom").setTitle("Validar cupom")
    return page.evaluate()
  }
  return HtmlService.createHtmlOutputFromFile("verificar cupom").setTitle("Validar Cupom")
}

function changeStatusToUsedVoucher(cupom, email) {

  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main")
  var values = mainSheet.getDataRange().getValues();
  for (row = 1; row < values.length; row++) {
    if (cupom == values[row][10] && email == values[row][1]) {
      
      if (values[row][12] == false) {
        mainSheet.getRange(row + 1, 13).setValue(true)
        console.log('foiqqw;')
        try {
          MailApp.sendEmail({
            to: 'autom.adossikacakes@gmail.com',
            subject: 'Validado com Sucesso',
            htmlBody: "<h2>Validado com sucesso!</h2><br><br>Nome: " + values[row][8] + "<br>E-mail:" + values[row][1],
          });
        }
        catch {
          Logger.log('Error to send Voucher');
        }
      }
      else {
        try {
          MailApp.sendEmail({
            to: 'autom.adossikacakes@gmail.com',
            subject: 'Não validado',
            htmlBody: "<h2>Não validado!</h2><br><br>Nome: " + values[row][8] + "<br>E-mail:" + values[row][1],
          });
        }
        catch {
          Logger.log('Error to send Voucher');
        }
      }

    }

  }

  return HtmlService.createHtmlOutputFromFile("verificar cupom").setTitle("Validar Cupom");
}

function sendVoucherToUser() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main")
  var values = mainSheet.getDataRange().getValues()

  for (row = 1; row < values.length; row++) {

    var file = DriveApp.getFileById(values[row][13])

    if (values[row][11] == false) {
      try {
        MailApp.sendEmail({
          to: values[row][1],
          subject: 'Bem-vindo ao Clube Adossia',
          body: '<h2>Obrigado por se cadastrar' + values[row][9] + '</h2><br><br>Acompanhe-nos para receber promoções semanalmente.<br>Seu cupom está em anexo e para usá-lo, basta enviar para o <a href="instagram.com/adossika_cakes">Instagram Adossika</a><br>Mas lembre-se que ele só poderá ser usado uma vez.<br><br>Atenciosamente,<br>--<br><i>Adossika Cakes</i>',
          attach: [file.getAs(MimeType.PDF)]
        });

        mainSheet.getRange(row + 1, 12).setValue(true);
      }
      catch {
        Logger.log('Error to send Voucher');
      }
    }
  }
}