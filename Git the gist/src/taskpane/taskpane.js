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
    function callback(result) {
      //Do something with the result
      // const input = result.value;
      // console.log("Selected text:" + input)};
      
      // function realTimeDetection() {
        //input = Office.context.mailbox.item.body;
        Office.context.mailbox.item.body.getAsync(
          "text",
          { asyncContext: "This is passed to the callback" },
          function callback(result) {
            //Do something with the result
            const input = result.value;
            input.addEventListener("keydown", (e) => {
          //When user presses enter and event detected
          console.log(e.target.value);
          }, false);
            input = '123';
            element.dispatchEvent(new Event("keyup"));
          }
    )})}


