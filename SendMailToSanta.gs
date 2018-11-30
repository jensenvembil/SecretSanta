function InformSecretSanta(){
  console.info("InformSecretSanta----> START");
   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]); 
  var sheet = spreadsheet.getActiveSheet();
   var lastRow = sheet.getLastRow();
	  var startRow = 2;	
	 
	  for (var i =startRow ; i <= lastRow; i++) {
       
		var SecretSantaEmailId =sheet.getRange(i, 8).getValue();
        var childName = sheet.getRange(i, 4).getValue();
      var mailContent = sheet.getRange(i, 9).getValue();   
        var SecretSantaName = sheet.getRange(i, 7).getValue();   
       
       var sendMailStatus = sendMail(SecretSantaEmailId,mailContent,childName,SecretSantaName);
        if(sendMailStatus){
        sheet.getRange(i, 10).setValue(true);
        }
        else{
        sheet.getRange(i, 10).setValue(false);
        }
	  }
  SpreadsheetApp.flush();
  /*Setting Email IDs start*/
   console.info("InformSecretSanta----> END");
	}
  
function sendMail(SecretSantaEmailId,mailContent,childName,SecretSantaName){
  var status = false;
console.info("sendMail----> START");

  var _to = "email@domain.com";
  
  /*Setting Email IDs end*/ 
   var subjectLine ="[STC Secret Santa] Mail from "+childName;
		  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
   console.info("Email Quota Remaining "+emailQuotaRemaining);
		  if(emailQuotaRemaining >20 )
		  {
			MailApp.sendEmail({
			   to:_to,
			 // bcc:_bcc,
			  subject:subjectLine,
			  htmlBody: mailContent, 
			});
            console.info("sendMail----> Mail Sent to:"+SecretSantaEmailId);
            status = true;
          } 
		  else
		  {
			 MailApp.sendEmail({
			  to:'email@domain.com',
			  subject: "Email Quata Exhaust--> Birthday Wishes!!!",
			  htmlBody:"<h1> Email Quota <b>"+ emailQuotaRemaining +"</b> remaining</h1>"
					  
			});
            console.info("Mail not sent as Quota Expired , Remaining Quota:"+emailQuotaRemaining);
		  }  
	   console.info("sendMail----> END");
  return status;
}