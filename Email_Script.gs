/* This script is used for retrieving all the data from the sheet and sending there topic along with the batch number and teammates to all the students
@author : Nihal M Pise
@version: 1.0 
@Date   : 29/04/2021
*/

// The getDataOne() is used to fatch the names,indexvalues of different variables
function getDataOne() {
 // Declaration Of Variable and Objects
 let theCurrentlyOpenSheet=SpreadsheetApp.getActive() //fetching the currently active/opened Google Sheet 
 let allTheValuesOfNameColumn=theCurrentlyOpenSheet.getRange('C7:D69').getValues(); // Fatch the names with in the spified range
 let batchingTheList=0,batchNo=0,indexvar=0; 
 let listOfStudents=[];
 let topicIndex=0,dateIndex=0,minIndexValue=0,maxIndexvalue=4;
  //Fetching the values of Name Column
  for(let name= 0; name<allTheValuesOfNameColumn.length;name++)
      {   batchingTheList++; 
            listOfStudents[indexvar] =allTheValuesOfNameColumn[name][0];
            indexvar++; 
          if(batchingTheList==4||name == allTheValuesOfNameColumn.length-1)
          { batchNo++;           
            batchingTheList=0;
            getDataTwo(listOfStudents,batchNo,topicIndex,dateIndex,minIndexValue,maxIndexvalue);//Calling getDataTwo Function 
            minIndexValue=maxIndexvalue;
            maxIndexvalue+=4;
            if (maxIndexvalue == 64)
            {
              maxIndexvalue-=1; //Last batch has only 3 members 
            }
            indexvar=0;
            listOfStudents=[0,0,0,null]
            topicIndex+=4;
            dateIndex+=4;
          }
         
      }      
}

// This Function is Used to sending the email as well as fetching the Data from the active sheet  
function getDataTwo(studentName,batchNo,topicIndex,dateIndex,minIndexValue,maxIndexvalue)
{
    // Declaration Of Variable and Objects 
    let theCurrentlyOpenSheet=SpreadsheetApp.getActive() 
    let allTheValuesOfDateColumn=theCurrentlyOpenSheet.getRange('I7:J69').getValues();
    let allTheValuesOfTopicColumn=theCurrentlyOpenSheet.getRange('F7:H69').getValues();
    let allTheValuesOfUsnColumn=theCurrentlyOpenSheet.getRange('E7:E69').getValues();
    let usersEmail=[];
    let subject="Software Testing Batch List",body;

    for(minIndexValue;minIndexValue<maxIndexvalue;minIndexValue++){
        //Creation of Email id with Usn 
        usersEmail[minIndexValue]=allTheValuesOfUsnColumn[minIndexValue][0]+"@students.git.edu";
        //Defining Body of Email 
        body="Students Name :\t"+studentName+"\n\n"+"Batch No :\t"+batchNo+"\n\n"+"Topic Name :"+allTheValuesOfTopicColumn[topicIndex][0] +"\n\n"+"Date :"+allTheValuesOfDateColumn[dateIndex][0];
        //Sending Email
        MailApp.sendEmail(usersEmail[minIndexValue],subject,body)
   }
   Utilities.sleep(100000); // Email Delay
}
 