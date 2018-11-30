
function AssignSecretSanta() {
  //code to assign secret santa
   var peopleList=[]; 
  
   //Seting variables to get spreadsheet
		
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]); // getting details of Image Links
  var sheet = spreadsheet.getActiveSheet();
   var lastRow = sheet.getLastRow();
	  var startRow = 2;	
	 
	  for (var i =startRow ; i <= lastRow; i++) {
        var email=sheet.getRange(i, 2).getValue();
		var FullName=sheet.getRange(i, 4).getValue();
        var wishList= sheet.getRange(i, 5).isBlank()?'anything as you wish':sheet.getRange(i, 5).getValue();
        var number= sheet.getRange(i, 3).getValue();	  
        
        peopleList.push({fullName: FullName,emailId:email, secretSanta:"", secretSantaEmailId:"", wishList:wishList, number:number});  
	  }
  SpreadsheetApp.flush();
 var iteration= 0;
  peopleList = GetNewSantaList(peopleList);
  
  while(checkIfAnyUserandSecretSantaIsTheSame(peopleList)){   
    console.info("No of Try:"+iteration++);
      peopleList = GetNewSantaList(peopleList);    
  } 
  
   WriteNamesOfSecretSanta(peopleList);
}

function GetNewSantaList(peopleList){
   console.info("GetNewSantaList--->START");
  var santaList = peopleList.slice();  
  
  //select a random record  
  peopleList.sort(function() { return 0.5 - Math.random();});  
        
  peopleList.forEach(function(item){ 
    console.info("Selected Memeber:"+ item.fullName+"Email ID:"+item.emailId);
    santaList.sort(function() { return 0.5 - Math.random();});   
	while(santaList.length > 0){
      var santaData = santaList[0].emailId == item.emailId? santaList.pop():santaList.shift();
      console.info("Selected Santa:"+ santaData.fullName+ "Email ID :"+ santaData.emailId);
		 if(santaData.emailId != item.emailId){
            console.info("Name: "+item.fullName +" Secret Santa: "+item.secretSanta+"is not equal hence assigingig secret santa");
			 item.secretSanta = santaData.fullName;    
           item.secretSantaEmailId = santaData.emailId;
           Logger.log("Name: "+item.fullName +" Secret Santa: "+item.secretSanta);    
           console.info("Name: "+item.fullName +" Secret Santa: "+item.secretSanta);
           // Logger.log("santaList empty")
           break;
		 }
      else{
        console.info("Count of members if same --->"+santaData.length+"Selected Member:"+ item.fullName +"and Secret Santa:"+ santaData.fullName+"are the same.");
        if(santaList.length ==0){
          //this means  there is only one element in the list as well as the secret santa list hence will never proceed further 
          //hence to be re considered for re-shuffelling
          item.secretSanta = santaData.fullName;    
           item.secretSantaEmailId = santaData.emailId;
          console.info("on length 0 there is only one element is present. Name: "+item.fullName +" Secret Santa: "+santaData.fullName);
           break;
        }
        else{
          console.info("Pushing santa back to list as the Selected Member:"+ item.fullName +"and Secret Santa:"+ santaData.fullName+"are the same.");
           santaList.push(santaData);
        }
      }
      console.info("count of members in santa List--->"+santaData.length);
	}
  });
   console.info("GetNewSantaList--->END");
  return peopleList;
}

function checkIfAnyUserandSecretSantaIsTheSame(peopleList){
console.info("checkIfAnyUserandSecretSantaIsTheSame--->START");
  var isSame = false;
 
  peopleList.forEach(function(item){
    
    if(item.emailId == item.secretSantaEmailId){
      Logger.log("Duplicate found User: "+ item.emailId+" Santa: "+item.secretSantaEmailId);
      isSame = true;  
      return isSame;      
    }
  });
  console.info("checkIfAnyUserandSecretSantaIsTheSame  isSame--->"+isSame);
  console.info("checkIfAnyUserandSecretSantaIsTheSame--->END");
  return isSame;
}

function WriteNamesOfSecretSanta(peopleList){
  console.info("WriteNamesOfSecretSanta--->START");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]); // getting details of Image Links
   var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var startRow = 2;	 
  if(peopleList.length >0){
    for (var i =startRow ; i <= lastRow; i++) {
      var email=sheet.getRange(i, 2).getValue();
      var FullName=sheet.getRange(i, 4).getValue();		
      peopleList.forEach(function(item){
        if(item.emailId==email){          
        sheet.getRange(i, 7).setValue(item.secretSanta);
          sheet.getRange(i, 8).setValue(item.secretSantaEmailId);   
          var mailContent= getMailContent(item);
          sheet.getRange(i, 9).setValue(mailContent);  
        }
      });      
    }
  }
  SpreadsheetApp.flush();
  console.info("Santa has been assigned!");
Logger.log("Santa has been assigned!");
  console.info("WriteNamesOfSecretSanta--->END");
}

function getMailContent(people){ 
  console.info("GetMailContent--->START");
  //random quality of this person
  var quality = getRandomQuality();
var html='';
  html+='<div style="font-family: comic sans ms, sans-serif;">';
  html+='<p><strong><span style="font-size: 16pt; color: #800000;"><span>Dear Santa</span>,</span></strong><br />&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;';
  html+='<span>My name is <span style="color: #339966;"><strong>'+people.fullName+'</strong> </span>and my email address is <span style="color: #339966;"><strong>'+people.emailId+'</strong></span>.';
  html+='This year i have been <strong><span style="color: #339966;"><em>'+quality+'</em> </span></strong>.</span><br />';
  html+='<span >&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Santa, If you are not too busy here is my wish list. ';
  html+='I would Love to have <span style="color: #339966;"><strong>'+people.wishList+'</strong></span> for this christmas.</span></p>';
  html+='<p><span >Thank you Santa, for all that you do!</span><br /><span >Please pray for me and I hope that someone is good to you too.</span></p>';
  html+='<p><span >With Love,</span><br />';
  html+='<span style="color: #993300;"><strong>'+people.fullName+'</strong></span><br /><span style="color: #008000; ">'+people.number+'</span><br /><span style="color: #008000; ">'+people.emailId+'</span></p>'
  html+='<p><strong>PS:&nbsp;</strong><span >This is </span><span >an auto-generated <strong>Secret mail</strong></span><span >, please <strong>do not </strong></span><strong><span >reveal yourself</span></strong><span > until the day of <span style="text-decoration: underline;"><strong>Christmas</strong> </span>or <span style="text-decoration: underline;"><strong>with a small gift</strong></span></span><strong><span >.</span></strong></p>';
  html+='</div>';
  console.info("GetMailContent--->END");
return html;
}

function getRandomQuality(){
var quality=["good all the time","good some of the time","naughty (but nice)"];
   quality.sort(function() { return 0.5 - Math.random();});  
  return quality[0];
}
