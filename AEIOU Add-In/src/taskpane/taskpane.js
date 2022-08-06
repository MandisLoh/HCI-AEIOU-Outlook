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
        
        //result = Result of running model
        const start = new Date();
        const end = new Date();
        end.setHours(start.getHours() + 1);

        //if (result){displayNewAppointmentForm}
        Office.context.mailbox.displayNewAppointmentForm({
          start: start,
          end: end,
          location: "Home",
          subject: "meeting",
          body: result.value
        });
        //else{document.getElementById("item-subject").innerHTML = Not a meeting!}
    });
  
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.body;
  // document.body.getElementsByClassName("foo");
}