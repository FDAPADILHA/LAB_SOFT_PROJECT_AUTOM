function enviarPropagandaSemanal() {
var values = sheet.getDataRange().getValues();

for (row = 1; row < values.length; row++) {

      mailApp.sendEmail
        ({
          to: values[row][1],
          subject: "Propaganda da semana",
          htmlBody: "Bom dia"+values[row][9]+", <br><br> Conheça nossas novidades da semana. <br><br>Atenciosamente,<br><br><i>Adossika Cakes</i>",
          attachments: files.getFilesByName("propaganda.jpg").next(),
        });
  }
}