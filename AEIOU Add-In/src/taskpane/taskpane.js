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
                
        // const spawn = require('child_process').spawn;
        // const script = spawn('py', ['ml copy 2.py', result.value]);
        // console.log(result.value)

        // script.stdout.on('data', (data) => {
        //     // datatoSend = data.toString();
        //     console.log(`${data}`)
        // });
        // script.stderr.on('data', (data) => {
        //     // As said before, convert the Uint8Array to a readable string.
        //     console.error(`stderr: ${data}`);
        // });

        // script.on('close', (code) => {
        //     console.log("Process quit with code : " + code);
        //     // res.send(datatoSend);
        // });

        // Setting up tfjs with the model we downloaded
        tf.loadLayersModel('model.json')
        .then(function (model) {
            window.model = model;
            console.log("window model")
        });

        // Predict function
        let predicted = 0;
        var predict = function (input) {
        if (window.model) {
            window.model.predict([tf.tensor(input)
                .reshape([1, 28, 28, 1])])
                .array().then(function (scores) {
                    scores = scores[0];
                    predicted = scores
                        .indexOf(Math.max(...scores));
                    $('#number').html(predicted);
                    console.log(predicted)
                });
        } else {

            // The model takes a bit to load, 
            // if we are too fast, wait
            setTimeout(function () { predict(input) }, 50);
        }
        }

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
          location: "",
          subject: "",
          body: result.value
        });
        //else{document.getElementById("item-subject").innerHTML = Not a meeting!}
    });
  
  // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.body;
  // document.body.getElementsByClassName("foo");
}