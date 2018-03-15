// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "SEND";

function sendEmails() 
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var uacmLogo = UrlFetchApp.fetch("http://www.analuisafontela.com/wp-content/uploads/2018/02/JxTCeRsn.jpg").getBlob().setName("uacmLogo");
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 5)// Los número son las columnas que acepta
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var col=data[0];
  var salto=col[4];
  for (var i = 0; i < data.length; ++i) {
    var col = data[i];
    var emailAddress = col[0];  // Primera columna 
    var message = "Hola "+col[2]+","+salto+salto+
        "Por fin ha llegado la semana del URJC TECHNOLOGY FEST y estamos esperándote con todo preparado. Queremos recordarte que los bloques horarios en que te has inscrito son "+col[1]+
        "."+salto+salto+ "Para el registro te pedimos que estés un poco antes del comienzo de cada ponencia, cuando llegues y para agilizar el proceso, te encontraras con letreros con las distintas letras, sitúate en la fila que corresponda con la inicial de tu primer apellido."+salto+salto+
        "Un saludo,"+salto+salto+"La organización";       // Second column
    
    var emailSent = col[3]; // Cuarta columna     
    if (emailSent != EMAIL_SENT) 
    {  // Prevents sending duplicates
      var subject = "UACM (URJC Technology Fest)";
      MailApp.sendEmail(emailAddress, subject, message,
        { htmlBody: 
         "<img src='cid:uacmLogo'>UNIÓN DE ALUMNOS DEL CAMPUS DE MÓSTOLES<br><br>"+
         "Hola "+col[2]+",<br><br>"+
         "Por fin ha llegado la semana del URJC TECHNOLOGY FEST y estamos esperándote con todo preparado. Queremos recordarte que los bloques horarios en que te has inscrito son "+col[1]+
         ".<br><br>Para el registro te pedimos que estés un poco antes del comienzo de cada ponencia, cuando llegues y para agilizar el proceso, te encontraras con letreros con las distintas letras, sitúate en la fila que corresponda con la inicial de tu primer apellido.<br><br>"+
         "Un saludo,<br><br>La organización",
       inlineImages: 
          { uacmLogo: uacmLogo
          }
        }                 
      );
      sheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();  
    }
  }
}