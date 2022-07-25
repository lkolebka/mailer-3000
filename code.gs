// works on Google Sheet

function SendMail(){
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const ws = ss.getSheetByName("Mailer 3000");
 const name = ws.getRange("E4").getValue();
 
  // met la date et l'heure au format texte
  var txt1 = ws.getRange("C8");
  var txt2 = txt1.setNumberFormat('@');
  
  var txt3 = ws.getRange('C10');
  var txt4 = txt3.setNumberFormat('@');
     
 
  var date = 
   ws.getRange("C8").getValue(); // date
 var hour = 
  ws.getRange("C10").getValue(); // heure


 const mailAdress = ws.getRange("I4").getValue(); //mail value 

 const htmlTemplate = HtmlService.createTemplateFromFile("email") //replaces the HTML values with the values from the code file
 htmlTemplate.name = name; 
 htmlTemplate.date = date;
 htmlTemplate.hour = hour; 
 
 const htmlForEmail = htmlTemplate.evaluate().getContent() //Fills the HTML with the new values
 console.log(htmlForEmail);
 
 GmailApp.sendEmail(
          mailAdress, //destinatire
          "Entretien d'embauche - "+ name, // sujet
          "Merci d'ouvrir cet email avec une boite mail prenant en charge le HTML", //in case of error with mail client
 
         { htmlBody: htmlForEmail } //corps en HTML 
 
);

    //resets the format to date mode in order to use the calendar
     var txt5 = txt1.setNumberFormat('dd-MM-yyyy');
   
    
   // +1 coutner 
   var range = ws.getRange("E8"); 
   var value = range.getValue();
   range.setValue(value + 1);


  //Notification toast
  var message = 'Le mail a été envoyé à '+ name
  var title = 'Mail envoyé ✅';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}



 




  
 
