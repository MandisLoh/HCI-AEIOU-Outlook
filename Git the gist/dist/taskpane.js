!function(){Office.onReady((function(e){e.host===Office.HostType.Outlook&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=run)})),$("#run-button").onclick("run");var e=new Date,o=new Date;o.setHours(e.getHours()+1),Office.context.mailbox.displayNewAppointmentForm({requiredAttendees:["bob@contoso.com"],optionalAttendees:["sam@contoso.com"],start:e,end:o,location:"Home",subject:"meeting",resources:["projector@contoso.com"],body:result.value})}();
//# sourceMappingURL=taskpane.js.map