/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});


  /**
   * Insert your Outlook code here
   * 
   */
  // const spawn = require('child_process').spawn;
  // console.log("input")
  // const script = spawn('python.exe', ['ml copy 2.py', input]);
  // console.log(input)
  
  // script.stdout.on('data', (data) => {
  //   // datatoSend = data.toString();
  //   console.log(`${data}`)
  // });
  // script.stderr.on('data', (data) => {
  //   // As said before, convert the Uint8Array to a readable string.
  //   console.error(`stderr: ${data}`);
  // });
  
  // script.on('close', (code) => {
  //   console.log("Process quit with code : " + code);
  // });
   /* global document, Office */
  // Get a reference to the current message
  // const item = Office.context.mailbox.item

  export async function run() {
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
        document.getElementById("item-subject").innerHTML = "<b>Body:</b> <br/>" + result.value;
        
        input = result.value
  
        
   
        // Put in model to use the result.value
        tf.loadLayersModel('model.json')
        .then(function (model) {
            window.model = model;
            console.log("window model")
        });
        
        //result = Result of running model
        const start = new Date();
        const end = new Date();
        end.setHours(start.getHours() + 1);

        //if (result){displayNewAppointmentForm}
        Office.context.mailbox.displayNewAppointmentForm({
          start: start,
          end: end,
          location: "",
          subject: "",
          body: result.value
        });
        //else{document.getElementById("item-subject").innerHTML = Not a meeting!}
    });
  
        }

    // function callml(input) {   
    //   const spawn = require('child_process').spawn;
    //   console.log("input")
    //   const script = spawn('python.exe', ['ml copy 2.py', input]);
    //   console.log(input)
  
    //   script.stdout.on('data', (data) => {
    //     // datatoSend = data.toString();
    //     console.log(`${data}`)
    //   });
    //   script.stderr.on('data', (data) => {
    //     // As said before, convert the Uint8Array to a readable string.
    //     console.error(`stderr: ${data}`);
    //   });
  
    //   script.on('close', (code) => {
    //     console.log("Process quit with code : " + code);
        // res.send(datatoSend);
      
        
      
    //   });
  
    //   script.on('close', (code) => {
    //     console.log("Process quit with code : " + code);
        // res.send(datatoSend);
      
        
      
          
          
          
    
        
    

  
        // Put in model to use the result.value
        

        
    //     const start = new Date();
    //     const end = new Date();
    //     end.setHours(start.getHours() + 1);

    //     Office.context.mailbox.displayNewAppointmentForm({
    //       requiredAttendees: ["bob@contoso.com"],
    //       optionalAttendees: ["sam@contoso.com"],
    //       start: start,
    //       end: end,
    //       location: "Home",
    //       subject: "meeting",
    //       resources: ["projector@contoso.com"],
    //       body: result.value
    //     });
    // });
    
    
    
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.body;
  // document.body.getElementsByClassName("foo");




