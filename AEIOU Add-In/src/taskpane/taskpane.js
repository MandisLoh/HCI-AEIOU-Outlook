/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */




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
        document.getElementById("item-subject").innerHTML = "<b>Body:</b> <br/>" + result.value;
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
        console.log(returned_prediction);
        
        if (returned_prediction = {"Prediction":"Meeting"}) {
        const start = new Date();
        const end = new Date();
        end.setHours(start.getHours() + 1);
        Office.context.mailbox.displayNewAppointmentForm({
                start: start,
                end: end,
                location: "",
                subject: Office.context.mailbox.item.subject,
                body: result.value
              })}         

        else {document.getElementById("item-subject").innerHTML = "Not a meeting!"}
          ; }
  )}
        
      

        
        
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