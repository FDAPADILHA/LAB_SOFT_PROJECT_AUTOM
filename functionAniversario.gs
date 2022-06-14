const mailApp = MailApp;
const app = SpreadsheetApp;
const files = DriveApp;
const spreadsheet = app.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName("Main");

var aniversariantes = new Date();
var dia = aniversariantes.getDate();
var mes = aniversariantes.getMonth();

function enviaCartaoAniversario() {
  var values = sheet.getDataRange().getValues();

  for (row = 1; row < values.length; row++) {

    diaMesAniversario = getDiaMesAniversario(values[row][4])

    if ( aniversario(diaMesAniversario[0], diaMesAniversario[1])) {
      var file = files.getFileById(values[row][18]); // OBTEM O ARQUIVO PDF

      mailApp.sendEmail
        ({
          to: values[row][1],
          subject: "Feliz Aniversário",
          htmlBody: "Bom dia, <br><br> Hoje é um dia mais do que especial. Nós da Adossika Cakes agradecemos sua atenção e lhe desejamos um feliz aniversário! Que você aproveite este dia tão especial <br><br>Atenciosamente,<br><br><i>Adossika Cakes</i>",
          attachments: [file.getAs(MimeType.PDF)],
        });

    }
  }
}


//====================================================================================================
// VERIFICA SE EH ANIVERSÁRIO
function aniversario(diaAniversario, mesAniversario) {
  if (diaAniversario == dia && mesAniversario - 1 == mes) {
    return true;
  }
  else {
    return false;
  }
}

function getDiaMesAniversario(dataCompleta) {
  const dataAniversario = new Date(dataCompleta)
  diaMesAniversario = [dataAniversario.getDay(), dataAniversario.getMonth()]
  return diaMesAniversario
}