
/*
Traverses folders in drive and accesses the target folder specified 
before creating the correct file structure beneath it 
*/

function activeUser(e){
  // FORM RESPONSES FOR FILE CREATION AND OVERVIEW DATA//
  var responses = e.response.getItemResponses();
  var formData = Array(6);
  for(var i = 0; i < responses.length; i++) {
    var response = responses[i].getResponse();
    formData[i] = response;
  }
  // Only trigger once
  var email = Session.getActiveUser().getEmail();
  if(email){
    createFolderStructure(formData); //CREATE FOLDER TREE
  }
}

function createFolderStructure(formData){ 
  
  // Target folder
  var ESMC = formData[0];

  var folders = DriveApp.getFolders(); 
  while (folders.hasNext()){
    var folder = folders.next(); 
    if(folder.getName()== '1 - Engines'){
      
      // Iterate counter for naming convention
      var counter = 0; 
      var subFolders = folder.getFolders(); 
      while(subFolders.hasNext()){
        var subFolder = subFolders.next(); 
        counter++;
      }
      
      var folderName;
      var folderNumber = 0;
      
      
      // Naming convention check
      if(ESMC == '1 - Engines') {
        folderNumber = 1001 + counter;
        folderName = folderNumber.toString() + "-ENG";
      } else if(ESMC == '2 - Systems'){
        folderNumber = 2001 + counter;
        folderName = folderNumber.toString() + "-SYS";
      } else if(ESMC == '3 - Military'){
        folderNumber = 3001 + counter;
        folderName = folderNumber.toString() + "-MIL";
      } else if(ESMC == '4 - Commercial'){
        folderNumber = 4001 + counter;
        folderName = folderNumber.toString() + "-COM";
      }
      
      //Creates initial folder in target folder
      var clientFolder = folder.createFolder(folderName); //Top level-project name
      
      //Creates first tier in initial folder
      var firstTier = DriveApp.getFolders(); 
      while(firstTier.hasNext()){
        var firstFolder = firstTier.next(); 
        if(firstFolder.getName() == clientFolder.getName()){
          //FIRST TIER SUBFOLDER CREATION//
          var Steve = firstFolder.createFolder('STEVE');
          Steve.createFolder(1001+counter+"-ENG-00"); //Creates Master Assembly file in STEVE
          var Documents = firstFolder.createFolder('Documents');
          //POPULATING OVERVIEW// 
          editOverview(formData, folderNumber, folderName, clientFolder);
          break;
        }
      }
        
      //Creates second tier in first tier folder
      var secondTier = firstFolder.getFolders(); 
      while(secondTier.hasNext()){
        var secondFolder = secondTier.next(); 
        if(secondFolder.getName() == Documents.getName()){
          //SECOND TIER OF DOCUMENTS// 
          var Meetings = secondFolder.createFolder('Meetings');
          var Reports = secondFolder.createFolder('Reports');
          var Diagrams = secondFolder.createFolder('Diagrams');
          var BOMS = secondFolder.createFolder('BOMS');
          var Finance = secondFolder.createFolder('Finance');
          var Management = secondFolder.createFolder('Management'); 
        }
      }
      
      
      
       //Creates third tier in second tier folder
      var thirdTier = secondFolder.getFolders();
      while(thirdTier.hasNext()){
        var thirdFolder = thirdTier.next(); 
        if(thirdFolder.getName() == Finance.getName()){
          //THIRD TIER OF FINANCE//
          var Invoices = thirdFolder.createFolder('Invoices'); 
          var Receipts = thirdFolder.createFolder('Receipts'); 
          var Quotes = thirdFolder.createFolder('Quotes'); 
          break;
        }
      }
      // Final break to exit first while-loop
      break;
    }
  }
}

function editOverview(array, folderNumber, folderName, destination){
  var template = DocumentApp.openByUrl('https://docs.google.com/document/d/1NNmE_zE-QQ502E0akp_J0qHjs5qTSPneXqLDRjWXzEc/edit');
  var overview = DriveApp.getFileById(template.getId()); 
  var overviewName = folderNumber + "-PRJ-Overview";
  var copy = overview.makeCopy(overviewName, destination ); //CREATES OVERVIEW TITLE//
  
  var docId = DocumentApp.openById(copy.getId()); 
  
  var dateCreated = copy.getDateCreated();
  var body = docId.getBody(); 
  
  body.replaceText("Doc Temp", folderName + "-Overview");
  body.replaceText("Project Temp", array[1]);
  body.replaceText("PM Name", array[2]);
  body.replaceText("Date Temp", dateCreated);
  body.replaceText("EST Complete Date", array[3]);
  body.replaceText("Background Text", array[4]); 
  body.replaceText("Objectives List", array[5]); 
  
  // Edit the header
  var header = docId.getHeader();
  header.replaceText("ARC ID TEMP", folderName + "-Overview" );
}



