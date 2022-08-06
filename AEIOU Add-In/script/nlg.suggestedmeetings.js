Type.registerNamespace('NLG.SuggestedMeetings');

/** @namespace NLG*/
var NLG = window['NLG'] || {};
NLG.SuggestedMeetings = NLG.SuggestedMeetings || {};

/**
 * @constructor
 * @extends {NLG.Common.SuggestionsUI}
 * @param {string} displayLanguage
 * @param {OSF.DDA.OutlookAppOm} appOm
 * @param {OSF.DDA.Settings} appSettings
 */
NLG.SuggestedMeetings.SuggestedMeetingsExtension = function(displayLanguage, appOm, appSettings) {
    this.$$d_findConflictWithCurrentItemCallback = Function.createDelegate(this, this.findConflictWithCurrentItemCallback);
    this.$$d__saveMeetingCallback$p$1 = Function.createDelegate(this, this._saveMeetingCallback$p$1);
    this.$$d__parseGetFolderResponseForFindItem$p$1 = Function.createDelegate(this, this._parseGetFolderResponseForFindItem$p$1);
    this.$$d__saveMeetingButtonClickHandler$p$1 = Function.createDelegate(this, this._saveMeetingButtonClickHandler$p$1);
    this.$$d__sendInvitationButtonClickHandler$p$1 = Function.createDelegate(this, this._sendInvitationButtonClickHandler$p$1);
    this.$$d__editDetailsClickHandler$p$1 = Function.createDelegate(this, this._editDetailsClickHandler$p$1);
    NLG.SuggestedMeetings.SuggestedMeetingsExtension['initializeBase'](this, [ displayLanguage, appOm, appSettings ]);
    this.outlookOm = appOm;
    this.perfTrackStart('Initialize');
    var /** {$h.ItemBase} */ thisItem = (this.outlookOm)['item'];
    var /** {Array} */ meetingsFromItem = this._getSelectedMeetingSuggestions$p$1(thisItem);
    this._meetings$p$1 = [];
    for (var /** {number} */ i = 0; i < meetingsFromItem['length'] && i < 99; i++) {
        if (this._meetingIsUnique$p$1(meetingsFromItem[i])) {
            this._meetings$p$1[this._meetings$p$1['length']] = new NLG.Common.SuggestedMeeting(meetingsFromItem[i], this._meetings$p$1['length']);
        }
    }
    if (this._meetings$p$1['length'] === 1) {
        this.set_shouldItemIconBeVisible(true);
    }
    this.start(this._meetings$p$1);
    this.perfTrackEnd('Initialize');
    this.logDatapoint({ 'NumberOfExtractedEntities': this._meetings$p$1['length'] });
}
NLG.SuggestedMeetings.SuggestedMeetingsExtension.$$cctor = function() {
    window['Apps']['Common']['ScriptHelper'].checkOfficeJsLoadedAndInitialize(NLG.SuggestedMeetings.SuggestedMeetingsExtension._initialize$p);
}
/**
 * @private
 * @function NLG.SuggestedMeetings.SuggestedMeetingsExtension._initialize$p
 */
NLG.SuggestedMeetings.SuggestedMeetingsExtension._initialize$p = function() {
    window['Microsoft']['Office']['WebExtension']['initialize'] = function(reason) {
        $(function() {
            var /** {NLG.SuggestedMeetings.SuggestedMeetingsExtension} */ thisExtension = new NLG.SuggestedMeetings.SuggestedMeetingsExtension(window['Office']['context']['displayLanguage'], window['Office']['context']['mailbox'], window['Office']['context']['roamingSettings']);
        });
    };
}
NLG.SuggestedMeetings.SuggestedMeetingsExtension.prototype = {
    /**
     * @private
     * @member {Array}
     */
    _meetings$p$1: null,
    /**
     * @private
     * @member {NLG.Common.SuggestedMeeting}
     */
    _currentMeeting$p$1: null,
    
    /**
     * @protected
     * @return {jQueryObject}
     * @override
     */
    loadIntroContent: function() {
        if (this._meetings$p$1['length'] > 1) {
            return ($('body').find('#introTemplate'))['tmpl'](this);
        }
        return null;
    },
    
    /**
     * @protected
     * @param {number} pageNumber
     * @param {number} pageSize
     * @return {jQueryObject}
     * @override
     */
    loadListViewContent: function(pageNumber, pageSize) {
        var /** {jQueryObject} */ listContainer = ($('body').find('#listControlTemplate'))['tmpl'](this._meetings$p$1[0]);
        var /** {number} */ pageStart = pageSize * pageNumber;
        var /** {number} */ pageEnd = Math['min'](pageStart + pageSize, this._meetings$p$1['length']);
        for (var /** {number} */ i = pageStart; i < pageEnd; i++) {
            var /** {jQueryObject} */ thisListItem = ($('body').find('#meetingListItemTemplate'))['tmpl'](this._meetings$p$1[i]);
            (listContainer).find('#itemList').append(thisListItem);
        }
        return listContainer;
    },
    
    /**
     * @protected
     * @param {number} index
     * @return {jQueryObject}
     * @override
     */
    loadItemFormContent: function(index) {
        var /** {jQueryObject} */ itemContainer = ($('body').find('#meetingItemFormTemplateDynamic'))['tmpl'](this._meetings$p$1[index]);
        return itemContainer;
    },
    
    /**
     * @protected
     * @param {jQueryObject} sender
     * @override
     */
    onItemLoad: function(sender) {
        if (sender) {
            var /** {jQueryObject} */ listItem = (sender)['closest']('li');
            var /** {number} */ itemIndex = window['parseInt'](listItem.attr('id'));
            this._currentMeeting$p$1 = this._meetings$p$1[itemIndex];
            $('#backBtn').attr('tabindex', '0');
        }
        else {
            $('#backBtn').hide();
            this._currentMeeting$p$1 = this._meetings$p$1[0];
        }
        $('#editMeetingDetailsButton').bind(this.userActionEvent, this.$$d__editDetailsClickHandler$p$1);
        $('#sendInvitationButton').bind(this.userActionEvent, this.$$d__sendInvitationButtonClickHandler$p$1);
        $('#saveMeetingButton').bind(this.userActionEvent, this.$$d__saveMeetingButtonClickHandler$p$1);
        $('#suggTaskbarBg').fadeTo(300, 0.9);
        $('#suggTaskbar').fadeTo(300, 1);
        this._getFolderCall$p$1(this.$$d__parseGetFolderResponseForFindItem$p$1);
    },
    
    /**
     * @protected
     * @param {jQueryObject} sender
     * @override
     */
    onItemDispose: function(sender) {
        $('#suggTaskbarBg').fadeOut(200);
        $('#suggTaskbar').fadeOut(200);
    },
    
    /**
     * @private
     * @param {$h.ItemBase} thisItem
     * @return {Array}
     */
    _getSelectedMeetingSuggestions$p$1: function(thisItem) {
        if (!!thisItem.getSelectedEntities && ((thisItem)['getSelectedEntities']())['meetingSuggestions']['length'] > 0) {
            var /** {$h.Entities} */ entities = (thisItem)['getSelectedEntities']();
            return (entities)['meetingSuggestions'];
        }
        return (thisItem)['getEntitiesByType'](window['Microsoft']['Office']['WebExtension']['MailboxEnums']['EntityType']['MeetingSuggestion']);
    },
    
    /**
     * @private
     * @param {$h.IMeetingSuggestion} meeting
     * @return {boolean}
     */
    _meetingIsUnique$p$1: function(meeting) {
        for (var /** {number} */ i = 0; i < this._meetings$p$1['length']; i++) {
            if ((this._meetings$p$1[i])['meetingText'] === meeting['meetingString']) {
                return false;
            }
        }
        return true;
    },
    
    /**
     * @private
     * @param {jQueryEvent} e
     */
    _editDetailsClickHandler$p$1: function(e) {
        this._updateLocationFromInput$p$1();
        var /** {Object} */ newAppointmentOptions = {};
        newAppointmentOptions['subject'] = this._currentMeeting$p$1['subject'];
        newAppointmentOptions['body'] = this._currentMeeting$p$1['meetingText'];
        newAppointmentOptions['location'] = this._currentMeeting$p$1['location'];
        newAppointmentOptions['start'] = this._currentMeeting$p$1['startTime'];
        newAppointmentOptions['end'] = this._currentMeeting$p$1['endTime'];
        if (this._currentMeeting$p$1['attendees']) {
            var /** {Array} */ requiredAttendees = [];
            for (var /** {number} */ i = 0; i < this._currentMeeting$p$1['attendees']['length']; i++) {
                if (this._validateEmailAddress$p$1((this._currentMeeting$p$1['attendees'][i])['emailAddress'])) {
                    requiredAttendees[i] = (this._currentMeeting$p$1['attendees'][i])['emailAddress'];
                }
            }
            newAppointmentOptions['requiredAttendees'] = requiredAttendees;
        }
        try {
            (this.outlookOm)['displayNewAppointmentForm'](newAppointmentOptions);
        }
        catch (ex) {
            this._displayFormError$p$1('There as an error saving this item.');
            this.logError('Unable to invoke new appointment form (' + ex['message'] + ')');
        }
        this.logDatapoint({ 'SaveToCalendarUserAction': 1 });
    },
    
    /**
     * @private
     * @param {string} emailAddress
     * @return {boolean}
     */
    _validateEmailAddress$p$1: function(emailAddress) {
        var /** {RegExp} */ emailPattern = new RegExp(/^\s*[\w\-\+_]+(\.[\w\-\+_]+)*\@[\w\-\+_]+\.[\w\-\+_]+(\.[\w\-\+_]+)*\s*$/);
        return emailPattern['test'](emailAddress);
    },
    
    /**
     * @private
     * @param {string} errorString
     */
    _displayFormError$p$1: function(errorString) {
        $('#notificationMessageText').text(errorString);
        $('#notificationMessage').fadeTo(300, 0.875);
        ($('#notificationMessage'))['delay'](1000).fadeOut(300);
    },
    
    /**
     * @private
     * @param {jQueryEvent} e
     */
    _sendInvitationButtonClickHandler$p$1: function(e) {
        this._saveMeetingToCalendar$p$1(true);
    },
    
    /**
     * @private
     * @param {jQueryEvent} e
     */
    _saveMeetingButtonClickHandler$p$1: function(e) {
        this._saveMeetingToCalendar$p$1(false);
    },
    
    /**
     * @private
     * @param {boolean} sendToAttendees
     */
    _saveMeetingToCalendar$p$1: function(sendToAttendees) {
        var /** {string} */ currentMeetingAttendees = '';
        var /** {boolean} */ hasAttendees = false;
        var /** {string} */ sendInvitationsMode = 'SendToNone';
        this._updateLocationFromInput$p$1();
        this.transitionFrame(new NLG.Common.Frame('Progress', ($('body').find('#progressScreenTemplate'))['tmpl'](null), null, null), 2);
        if (sendToAttendees) {
            sendInvitationsMode = 'SendToAllAndSaveCopy';
            currentMeetingAttendees = '<t:RequiredAttendees>';
            if (this._currentMeeting$p$1['attendees']) {
                for (var /** {number} */ i = 0; i < this._currentMeeting$p$1['attendees']['length']; i++) {
                    if (this._validateEmailAddress$p$1((this._currentMeeting$p$1['attendees'][i])['emailAddress'])) {
                        currentMeetingAttendees += '                  <t:Attendee>                     <t:Mailbox>                        <t:EmailAddress>' + (this._currentMeeting$p$1['attendees'][i])['emailAddress'] + '</t:EmailAddress>' + '                     </t:Mailbox>' + '                  </t:Attendee>';
                        hasAttendees = true;
                    }
                }
            }
            currentMeetingAttendees += '</t:RequiredAttendees>';
        }
        var /** {string} */ soapRequest = '<?xml version=\'1.0\' encoding=\'utf-8\'?><soap:Envelope xmlns:xsi=\'http://www.w3.org/2001/XMLSchema-instance\'                xmlns:m=\'http://schemas.microsoft.com/exchange/services/2006/messages\'                      xmlns:t=\'http://schemas.microsoft.com/exchange/services/2006/types\'                      xmlns:soap=\'http://schemas.xmlsoap.org/soap/envelope/\'>   <soap:Header>      <t:RequestServerVersion Version=\'Exchange2010\' />   </soap:Header>   <soap:Body>      <m:CreateItem SendMeetingInvitations=\'' + sendInvitationsMode + '\'>' + '         <m:Items>' + '            <t:CalendarItem>' + '               <t:Subject>' + this._currentMeeting$p$1['subject'] + '</t:Subject>' + '               <t:Body BodyType=\'Text\'>' + this._currentMeeting$p$1['meetingText'] + '</t:Body>' + '               <t:Start>' + this._currentMeeting$p$1['startTime'].format('yyyy-MM-ddTHH:mm:sszzz') + '</t:Start>' + '               <t:End>' + this._currentMeeting$p$1['endTime'].format('yyyy-MM-ddTHH:mm:sszzz') + '</t:End>' + '               <t:Location>' + this._currentMeeting$p$1['location'] + '</t:Location>' + ((hasAttendees) ? currentMeetingAttendees : '') + '            </t:CalendarItem>' + '         </m:Items>' + '      </m:CreateItem>' + '   </soap:Body>' + '</soap:Envelope>';
        try {
            this.outlookOm['makeEwsRequestAsync'](soapRequest, this.$$d__saveMeetingCallback$p$1);
        }
        catch (ex) {
            this._failedToScheduleAMeeting$p$1(ex['message']);
        }
        this.logDatapoint({ 'SaveToCalendarUserAction': (sendToAttendees) ? 0 : 2 });
    },
    
    /**
     * @private
     * @param {Function} successCallback
     */
    _getFolderCall$p$1: function(successCallback) {
        var /** {string} */ soapRequest = '<?xml version=\'1.0\' encoding=\'utf-8\'?>\n<soap:Envelope xmlns:xsi=\'http://www.w3.org/2001/XMLSchema-instance\' \n       xmlns:m=\'http://schemas.microsoft.com/exchange/services/2006/messages\' \n       xmlns:t=\'http://schemas.microsoft.com/exchange/services/2006/types\' \n       xmlns:soap=\'http://schemas.xmlsoap.org/soap/envelope/\'>\n  <soap:Header>\n    <t:RequestServerVersion Version=\'Exchange2013\' />\n  </soap:Header>\n  <soap:Body>\n    <m:GetFolder>\n      <m:FolderShape>\n        <t:BaseShape>IdOnly</t:BaseShape>\n      </m:FolderShape>\n      <m:FolderIds>\n        <t:DistinguishedFolderId Id=\'calendar\'/>\n      </m:FolderIds>\n    </m:GetFolder>\n  </soap:Body>\n</soap:Envelope>';
        try {
            this.outlookOm['makeEwsRequestAsync'](soapRequest, successCallback);
        }
        catch (e) {
            this._findConflictFailed$p$1(e['message']);
        }
    },
    
    /**
     * @private
     * @param {OSF.DDA.AsyncResult} result
     */
    _parseGetFolderResponseForFindItem$p$1: function(result) {
        var /** {string} */ folderId = '';
        var /** {string} */ changeKey = '';
        var /** {boolean} */ requestWasSuccessful = false;
        if (!result['error']) {
            var /** {Element} */ xmlResponse = NLG.Common.XmlDomDocument.fromString(result['value']);
            try {
                var /** {Element} */ changeKeyElem = (xmlResponse.getElementsByTagName('t:FolderId').length > 0) ? xmlResponse.getElementsByTagName('t:FolderId')[0] : xmlResponse.getElementsByTagName('FolderId')[0];
                folderId = changeKeyElem.attributes.getNamedItem('Id').value;
                changeKey = changeKeyElem.attributes.getNamedItem('ChangeKey').value;
                requestWasSuccessful = true;
            }
            catch ($$e_6) {
                requestWasSuccessful = false;
            }
        }
        if (requestWasSuccessful) {
            this._findConflictWithCurrentItemRequest$p$1(folderId, changeKey);
        }
        else {
            this._findConflictFailed$p$1('');
        }
    },
    
    /**
     * @private
     * @param {string} folderId
     * @param {string} changeKey
     */
    _findConflictWithCurrentItemRequest$p$1: function(folderId, changeKey) {
        var /** {string} */ soapRequest = '<?xml version=\'1.0\' encoding=\'utf-8\'?>\n<soap:Envelope xmlns:xsi=\'http://www.w3.org/2001/XMLSchema-instance\' \n       xmlns:m=\'http://schemas.microsoft.com/exchange/services/2006/messages\' \n       xmlns:t=\'http://schemas.microsoft.com/exchange/services/2006/types\' \n       xmlns:soap=\'http://schemas.xmlsoap.org/soap/envelope/\'>\n  <soap:Header>\n    <t:RequestServerVersion Version=\'Exchange2013\' />\n  </soap:Header>\n  <soap:Body>\n    <m:FindItem Traversal=\'Shallow\'>\n      <m:ItemShape>\n        <t:BaseShape>IdOnly</t:BaseShape>\n      </m:ItemShape>\n      <m:CalendarView MaxEntriesReturned=\'1\' StartDate=\'' + this._currentMeeting$p$1['startTime'].format('yyyy-MM-ddTHH:mm:sszzz') + '\' EndDate=\'' + this._currentMeeting$p$1['endTime'].format('yyyy-MM-ddTHH:mm:sszzz') + '\' />\n' + '      <m:ParentFolderIds>\n' + '        <t:FolderId Id=\'' + folderId + '\' ChangeKey=\'' + changeKey + '\' />\n' + '      </m:ParentFolderIds>\n' + '    </m:FindItem>\n' + '  </soap:Body>\n' + '</soap:Envelope>';
        try {
            this.outlookOm['makeEwsRequestAsync'](soapRequest, this.$$d_findConflictWithCurrentItemCallback);
        }
        catch (e) {
            this._findConflictFailed$p$1(e['message']);
        }
    },
    
    /**
     * @public
     * @param {OSF.DDA.AsyncResult} result
     */
    findConflictWithCurrentItemCallback: function(result) {
        var /** {boolean} */ conflict = true;
        var /** {boolean} */ requestWasSuccessful = false;
        if (!result['error']) {
            var /** {Element} */ xmlResponse = NLG.Common.XmlDomDocument.fromString(result['value']);
            try {
                var /** {Element} */ changeKeyElem = (xmlResponse.getElementsByTagName('m:RootFolder').length > 0) ? xmlResponse.getElementsByTagName('m:RootFolder')[0] : xmlResponse.getElementsByTagName('RootFolder')[0];
                conflict = changeKeyElem.attributes.getNamedItem('TotalItemsInView').value !== '0';
                requestWasSuccessful = true;
            }
            catch ($$e_5) {
                requestWasSuccessful = false;
            }
        }
        if (requestWasSuccessful) {
            this._currentMeeting$p$1['hasConflict'] = conflict;
            this._updateConflictText$p$1((conflict) ? window['_u']['ExtensionStrings']['l_SuggestedMeetingsHasConflict_Text'] : window['_u']['ExtensionStrings']['l_SuggestedMeetingsNoConflict_Text']);
        }
        else {
            this._findConflictFailed$p$1('');
        }
    },
    
    /**
     * @private
     * @param {OSF.DDA.AsyncResult} result
     */
    _saveMeetingCallback$p$1: function(result) {
        if (!result['error'] && result['value'].indexOf('NoError') >= 0) {
            this._currentMeeting$p$1['hasBeenScheduled'] = true;
            this.transitionFrame(new NLG.Common.Frame('Complete', ($('body').find('#meetingItemFormTemplateDynamic'))['tmpl'](this._currentMeeting$p$1), null, null), 2);
        }
        else {
            this._failedToScheduleAMeeting$p$1('');
        }
    },
    
    /**
     * @private
     * @param {string} e
     */
    _findConflictFailed$p$1: function(e) {
        this.logError('Failed to resolve a list of conflicts: ' + e);
        this._updateConflictText$p$1(window['_u']['ExtensionStrings']['l_SuggestedMeetingsScheduleUnavailable_Text']);
    },
    
    /**
     * @private
     * @param {string} text
     */
    _updateConflictText$p$1: function(text) {
        $('#conflictDiv').text(text);
    },
    
    /**
     * @private
     * @param {string} e
     */
    _failedToScheduleAMeeting$p$1: function(e) {
        this.logError('Failed to automatically schedule the meeting: ' + e);
        this._currentMeeting$p$1['canAutoSchedule'] = false;
        this.transitionFrame(new NLG.Common.Frame('Error', ($('body').find('#meetingItemFormTemplateDynamic'))['tmpl'](this._currentMeeting$p$1), null, null), 2);
    },
    
    /**
     * @private
     */
    _updateLocationFromInput$p$1: function() {
        this._currentMeeting$p$1['location'] = $('#whereDetails').val();
    },
    
    /**
     * @protected
     * @param {jQueryEvent} e
     * @override
     */
    accessibilityKeydownHandler: function(e) {
    }
}


/** @class */NLG.SuggestedMeetings.SuggestedMeetingsExtension.UserActionCTQValue = function() {}
NLG.SuggestedMeetings.SuggestedMeetingsExtension.UserActionCTQValue.prototype = {
    'sendInvitation': 0, 
    'editDetails': 1, 
    'addToCalendar': 2
}
NLG.SuggestedMeetings.SuggestedMeetingsExtension.UserActionCTQValue['registerEnum']('NLG.SuggestedMeetings.SuggestedMeetingsExtension.UserActionCTQValue', false);


NLG.SuggestedMeetings.SuggestedMeetingsExtension['registerClass']('NLG.SuggestedMeetings.SuggestedMeetingsExtension', NLG.Common.SuggestionsUI);
/**
 * @var {string}
 * @constant
 */
NLG.SuggestedMeetings.SuggestedMeetingsExtension.editTemplateName = '#meetingItemFormTemplateDynamic';
/**
 * @var {string}
 * @constant
 */
NLG.SuggestedMeetings.SuggestedMeetingsExtension.progressScreenTemplateName = '#progressScreenTemplate';
/**
 * @var {string}
 * @constant
 */
NLG.SuggestedMeetings.SuggestedMeetingsExtension.ewsDateTimeFormat = 'yyyy-MM-ddTHH:mm:sszzz';
NLG.SuggestedMeetings.SuggestedMeetingsExtension.$$cctor();
