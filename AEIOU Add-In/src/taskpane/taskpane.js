/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import {date,parseTime,time,formatDate} from "./dateTimeList.js"



Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {

  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result)  {
      document.getElementById("item-subject").innerHTML = "Not a meeting!";
        var input = result.value;
        console.log(input);

        // remove all punctuation
        var replaced = input.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()""]/g," ");
        console.log(replaced);
        // changes the consecutive empty spaces into only 1 empty space
        var replaced1 = replaced.replace(/\s{2,}/g,' ');
        console.log(replaced1);
        
        // replace empty space with +
        var replaced2 = replaced1.slice(0, -1).split(' ').join('+') + replaced1.slice(-1);
        console.log(replaced2);
        
        // heroku website
        var predicted = 'https://outlook-addin-aeiou.herokuapp.com/predict?textbody=' + replaced2.toLowerCase() ;
        console.log(predicted) ;
        
        //get predicted value
        var xmlHttp = new XMLHttpRequest();
        xmlHttp.open( "GET", predicted, false ); // false for synchronous request
        xmlHttp.send( null );
        var returned_prediction = xmlHttp.responseText;
        console.log(`this is the predicted return:`);
        //Initialize variables for extraction of date/time
        let listOfDates = date()
        let listOfTime = time()
        let startDate;
        let startTime;
        let endTime;
        let startTimeHour;
        let startTimeMin;
        let endTimeHour;
        let endTimeMin;
        //Extract date
        console.log(input.toLowerCase())
        for(let i=0;i<listOfDates.length;i++){
          if(input.toLowerCase().includes(listOfDates[i].toLowerCase())){
            startDate = listOfDates[i].toLowerCase()
            break
          }
        }
        if (!startDate.includes("-")&&!startDate.includes("/")){
          startDate = startDate.replace("th","")
          startDate = startDate.replace("1st","")
          startDate = startDate.replace("3rd","")
          startDate = startDate.replace("2nd","")
          startDate+=" 2022"
        } else{
          startDate = formatDate(startDate)
        }
        console.log(startDate)
        console.log(input.toLowerCase().includes("8 aug") + "hello")
        let position = 0;
        //Extract Time
        for(let i=0;i<listOfTime.length;i++){
          if(input.toLowerCase().includes(listOfTime[i].toLowerCase())&&position===0){
            startTime = listOfTime[i].toLowerCase()
            position+=1
            console.log(startTime,position)
          } else if(input.toLowerCase().includes(listOfTime[i].toLowerCase())&&position===1){
            endTime = listOfTime[i].toLowerCase()
            break
          }
        }

        console.log(startTime)
        console.log(endTime)
        if(!endTime){
          endTime = startTime
        }
        console.log("Here")
        let parseStartTime = parseTime(startTime)
        let parseEndTime = parseTime(endTime)
        startTimeHour = parseStartTime.hour 
        startTimeMin = parseStartTime.min 
        endTimeHour = parseEndTime.hour 
        endTimeMin = parseEndTime.min 

        //If endTime not found, assume +1 Hour
        if(startTimeHour===endTimeHour&&startTimeMin===endTimeMin){
          endTimeHour+=1
          console.log(endTimeHour)
        }
        console.log("final" + startTimeHour)

        if (returned_prediction == "0") {
          document.getElementById("item-subject").innerHTML = "Not a meeting!"
        }
        else {
          document.getElementById("item-subject").innerHTML = "Meeting detected!"
          const meetingDate = new Date(startDate);
          const endDate = new Date(startDate);
          meetingDate.setHours(startTimeHour);
          meetingDate.setMinutes(startTimeMin);
          endDate.setHours(endTimeHour)
          endDate.setMinutes(endTimeMin)
          Office.context.mailbox.displayNewAppointmentForm({
                  start: meetingDate,
                  end: endDate,
                  location: "",
                  subject: Office.context.mailbox.item.subject,
                  body: result.value
        }
  )};

})}
        
      

        
        
        // Put in model to use the result.value
        
        //result = Result of running model
    //     const start = new Date();
    //     const end = new Date();
    //     end.setHours(start.getHours() + 1);

    //     //if (result){displayNewAppointmentForm}
    //     Office.context.mailbox.displayNewAppointmentForm({
    //       start: start,
    //       end: end,
    //       location: "",
    //       subject: "",
    //       body: result.value
    //     });
    //     //else{document.getElementById("item-subject").innerHTML = Not a meeting!}
  // });
  
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.body;
  //document.body.getElementsByClassName("foo")