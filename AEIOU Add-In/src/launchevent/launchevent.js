/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onMessageSendHandler(event) {
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: event },
      getBodyCallback
    );
  }
  
  function getBodyCallback(asyncResult){
    let event = asyncResult.asyncContext;
    let body = "";
    if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
      body = asyncResult.value;
    } else {
      let message = "Failed to get body text";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }
  
    let matches = hasMatches(body);
    if (matches) {
      Office.context.mailbox.item.getAttachmentsAsync(
        { asyncContext: event },
        getAttachmentsCallback);
    } else {
      event.completed({ allowEvent: true });
    }
  }
  
  function hasMatches(body) {
    if (body == null || body == "") {
      return false;
    }
  
    const arrayOfTerms = [
      "meet",
      "meeting",
      "meetup",
      "meet-up",
      "gather",
      "gathering",
      "interview",
      "interviewing",
      "book on",
      "booked on",
      "plans on",
      "busy on",
      "see you",
      "event",
      "schedule", 
      "Invite" ,
      "invited",  
      "Appointment", 
      "Assembly", 
      "Conference", 
      "Summit", 
      "Dialogue", 
       "Forum", 
       "Convocation", 
       "Get-together", 
       "Chat", 
       "Consultation"
    ];
    for (let index = 0; index < arrayOfTerms.length; index++) {
      const term = arrayOfTerms[index].trim();
      const regex = RegExp(term, 'i');
      if (regex.test(body)) {
        return true;
      }
    }
  
    return false;
  }
  
  function getAttachmentsCallback(asyncResult) {
    let event = asyncResult.asyncContext;
    if (asyncResult.value.length > 0) {
      for (let i = 0; i < asyncResult.value.length; i++) {
        if (asyncResult.value[i].isInline == false) {
          event.completed({ allowEvent: true });
          return;
        }
      }
  
      event.completed({ allowEvent: false, errorMessage: "Looks like you're sending a meeting invite. Have you saved to calendar?" });
    } else {
      event.completed({ allowEvent: false, errorMessage: "Looks like you're sending a meeting invite. Have you saved to calendar?" });
    }
  }
  
  // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// function onMessageComposeHandler(event) {
//     setSubject(event);
//   }
//   function onAppointmentComposeHandler(event) {
//     setSubject(event);
//   }
//   function setSubject(event) {
//     Office.context.mailbox.item.subject.setAsync(
//       "Set by an event-based add-in!",
//       {
//         "asyncContext": event
//       },
//       function (asyncResult) {
//         // Handle success or error.
//         if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//           console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
//         }
  
//         // Call event.completed() after all work is done.
//         asyncResult.asyncContext.completed();
//       });
//   }
  
//   // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
//   Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
//   Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);