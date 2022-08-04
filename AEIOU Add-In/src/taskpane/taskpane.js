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
  /**
   * Insert your Outlook code here
   */
  // Get a reference to the current message
  // const item = Office.context.mailbox.item
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
        document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.value;
        
        // Put in model to use the result.value
        // var input = result.value

        // const spawn = require('child_process').spawn;
        // const script = spawn('python', ['./ml.py', input.toString()]);

        // var datatoSend 

        // script.stdout.on('data', (data) => {
        //   datatoSend = data.toString();
        // });


        // // one of the ways to pass parameters into the py file if the other one doesnt work
        // // var data = result.value
        // // script.stdin.write(data);
        // // // End data write
        // // script.stdin.end();

        // script.on('exit', (code) => {
        //   console.log("Process quit with code : " + code);
        //   res.send(datatoSend);
        // });

        //
        const start = new Date();
        const end = new Date();
        end.setHours(start.getHours() + 1);

        Office.context.mailbox.displayNewAppointmentForm({
          requiredAttendees: ["bob@contoso.com"],
          optionalAttendees: ["sam@contoso.com"],
          start: start,
          end: end,
          location: "Home",
          subject: "meeting",
          resources: ["projector@contoso.com"],
          body: result.value
        });
    });
  
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.body;
  // document.body.getElementsByClassName("foo");
}



