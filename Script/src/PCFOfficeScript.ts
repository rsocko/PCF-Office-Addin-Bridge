/// <reference path="../node_modules/@types/office-js/index.d.ts" />

/* global Office */ // Required for Office.initialize = function (reason) { ... }
//older: document, window

//TODO: figure out the proper format for typescript documentation (methods, properties, classes, files, etc.) ?? jsdoc?

//TODO: split files into separate .ts files and use export/import (for cleaner code)

/*
    Bridge:
        - Properties: 
            - pcfControlSource - htmlWindow containing the PCF control (for cross-frame callback)
            - officeReady - boolean indicator if office has completed loading the add-in
            - pcfReady - boolean indicator that the pcf control in the powerApp has completed loading and ready to send/receive data
            - itemReady - abstract method/property representing the office object (Mail Item, Page, OneNote notebook/note, etc. is loaded and can be passed to PCF control)
        - Methods: 
            public:
                postMessageHandler(event) - event handler for browser cross-window/iframe calls for messages & data from PCF control to the Office add-in
                    - consider logic passed to child class in case the event is of a type (eg data passed from PCF >> Add-in implemented in the class - like to write to OneNote or mail message)
                sendMessageToIframe() - method to send cross-window/iframe message & data from office add-in to pcf control

            protected:
                loadOfficeObject() - abstract method to call the Office app-specific class' logic to setup office object (mail item, document, notebook, etc.) to pass to PCF control
                getOfficeObjectMessage() : IOfficeMessage - abstract method to retrieve the message object to send to PowerApps


*/

namespace PCFOfficeScript {

    interface IOfficeMessage{
        readonly messageType : string
        //TODO: change to enum for types - fixed choices?
    }

    class OutlookMailMessage implements IOfficeMessage{
        readonly messageType: string = "mailbox.item.message"
    }

    class NullMailItemMessage implements IOfficeMessage{
        readonly messageType: string = "mailbox.item.undefined";
    }

    interface IOfficeAppAdapter{
        isOfficeObjectReady() : boolean;
        getOfficeObjectMessage() : IOfficeMessage
        loadOfficeObject() : void
        //Event when Office Object properties are updated (from async calls to Office API - eg to load Body, etc.)
        onOfficeObjectUpdated?() : void;
    }

    //Interface that provides the ability to create a new instance of the Office App Adapter
    interface IOfficeAppAdapterFactory{
        new (officeContext : Office.Context) : IOfficeAppAdapter;
    }

    class OfficeBridge{

        private officeAppAdapter : IOfficeAppAdapter;

        //TODO: add an 'init' method (or constructor) - to set default properties, etc.

        public pcfControlSource = null; //TODO: can this be read only with private backing property?
        public officeReady = false; //TODO: consider making this read only with private backing property?
        public pcfReady = false; //TODO: consider making this read only with a private backing property?

        public setOfficeAdapter(officeAdapter : IOfficeAppAdapter) { 
            this.officeAppAdapter = officeAdapter;
            this.officeAppAdapter.onOfficeObjectUpdated = () => this.sendOfficeObjectIfReady();
        }

        //TODO: use strongly typed event class for PostMessageHandler
    
        //*****************************
        //EVENT HANDLER FOR CROSS-DOMAIN / IFRAME MESSAGING
        //*****************************
        public postMessageHandler( event : any ) {
             //Output Debug Information to Console
             console.log("PCFOffice: Add-in / Parent Window: Received message from PCF/iframe/PowerApp - via postMessageHandler");
             console.log("* Message:", event.data);
             console.log("* Origin:", event.origin);
             console.log("* Source:", event.source);
           
             // check request is from legitimate source and message is expected or not
             //if ( event.origin !== 'https://powerapps.com' ) { console.log("not matching right domain");return; }
         
             // check the message type received (and take appropriate action)
             if(event.data.messageType) {//if (event.data.messageType is defined)
                 if (event.data.messageType === "PCF.Ready") {
                     // Received message from PCF Component (in the iframe) that it is ready to receive messages
                     this.pcfReady = true;
                     this.pcfControlSource = event.source;
         
                     console.log("PCFOffice: PCF.Ready message received from PCF/iframe/PowerApp");
         
                     //TODO: need to convert Text to enum value
                     /*
                     //Get the Requested Body Format (if specified in the message)
                     if(event.data.bodyCoercionType) {
                         //TODO: add validation that it is HTML or TEXT
                         this.addinState.bodyCoercionType = event.data.bodyCoercionType;
         
                         console.log("PCFOffice: Body Format requested by PCF/iframe/PowerApp: " + this.addinState.bodyCoercionType);
                     }
                     */
         
                     // Send back the current selected Mail Item
                     if(event.source)
                     {
                         console.log("PCFOffice: Add-in / Parent Window: Loading Office Object - via office app-specific load-method");
                         if(this.officeAppAdapter){
                            this.officeAppAdapter.loadOfficeObject();

                            this.sendOfficeObjectIfReady(); 
                            //Return Mock Mail Item
                            //event.source.postMessage(getMockMessageObject(), "*");   
                         }
                     }
                 }
             }
        }

        private sendMessageToIframe(message : IOfficeMessage) {
            console.log("PCFOffice: Add-in / Parent Window: Sending message to PCF/iframe/PowerApp - via sendMessageToIframe");
        
            //TODO: #1 add error handling if called without Source specified/defined (must have received initial call from PCF control)
        
            //TODO: specify the origin of the add-in (or use wildcard)
        
            this.pcfControlSource.postMessage(message, "*");
        }

        public sendOfficeObjectIfReady() {
            //if the data is all ready - send to PCF component (in the iframe)
            if( this.officeReady && this.pcfReady && this.officeAppAdapter && this.officeAppAdapter.isOfficeObjectReady()) {
                console.log('sending message to iframe since officeReady and itemReady and pcfReady:' + this.pcfReady + ' officeReady: ' + this.officeReady + ' itemReady: ' + this.officeAppAdapter.isOfficeObjectReady());
                
                //Send the Mail Item to the PCF Component (in the iframe)
                this.sendMessageToIframe(this.officeAppAdapter.getOfficeObjectMessage());
            }
        }
    }

    class OutlookAppAdapter implements IOfficeAppAdapter{
        constructor (officeContext : Office.Context){
            //Setup Event Handler for ItemChanged Event (if the Add-in is pinnable and mail item changes)
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.ItemChanged, 
                (asyncResult : Office.AsyncResult<void>) => this.itemChanged(asyncResult)
            );
        }

        //Event when Office Object properties are updated (from async calls to Office API - eg to load Body, etc.)
        public onOfficeObjectUpdated?() : void;
        
        //bodyCoercionType: Office.CoercionType.Html, //"Html", //Office.CoercionType.Html
        private mailItem? : Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead;
        private mailItemBody? : string = undefined;
        private mailItemUniqueBody? : string = undefined;

        public isOfficeObjectReady() : boolean { 
            if(!this.mailItem) // mailItem is null - valid state (no mail item selected)
                return true;
            else if(this.mailItem && this.mailItemBody && this.mailItemUniqueBody) // mailItem and mailItemBody and mailItemUniqueBody are all NOT null - valid state
                return true;
            else // mailItem is not null and mailItemBody is null - invalid state
                return false;
        }

        private raiseOfficeObjectUpdated() : void {
              //If Event Handler is not defined - don't call it (return from this method)
              if (!this.onOfficeObjectUpdated) return
              //Call the Event Handler
              this.onOfficeObjectUpdated();
        }

        private getUniqueBody(accessToken : string) {
            // Get the item's REST ID.
            const itemId : string = this.getItemRestId();
          
            // Construct the REST URL to the current item.
            // Details for formatting the URL can be found at
            // https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
            const getMessageUrl = Office.context.mailbox.restUrl +
              '/v2.0/me/messages/' + itemId + '/?$select=uniqueBody';
        
              //TODO: what if the message isn't in the user's mailbox (eg shared mailbox or archive, etc.)
          
            //TODO: ensure jquery is loaded (or use alternative method for http/ajax call)
            $.ajax({
              url: getMessageUrl,
              dataType: 'json',
              headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done((item : any) => {
              // Message is passed in `item`.
              const uniqueBody = item.UniqueBody.Content;
              this.mailItemUniqueBody = uniqueBody; // Save the unique body to the add-in state.
        
              ///TODO: remove this console write (outputs sensitive data from email + lots of data?)
              console.log('PCFOffice: mailbox.item.uniqueBody: ' + uniqueBody);
              
              this.raiseOfficeObjectUpdated();

            }).fail((error) => {
              // TODO: Handle error.
              this.mailItemUniqueBody = undefined; // reset unique body in state
            });
        }

        private getCallbackTokenCallback(result : Office.AsyncResult<string>) : void {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const accessToken = result.value;

                console.log('PCFOffice: retrieved access token for Outlook REST API');
                console.log('PCFOffice: retrieving unique body of mail item from Outlook REST API');
            
                // Use the access token to retrieve the UniqueBody property.
                this.getUniqueBody(accessToken);

                this.raiseOfficeObjectUpdated();
            } else {
                // Handle the error.

                //TODO: consider other error handling

                //TODO: verify uniqueBody exists even in a message without replies

                this.mailItemUniqueBody = undefined; // reset unique body in state
            }
        }

        private getMailBodyCallback(result : Office.AsyncResult<string>) : void {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                //Save the Mail Item Body to the Add-in State
                this.mailItemBody = result.value; // Get body data
                
                //TODO: remove this console write (outputs sensitive data from email + lots of data?)
                console.log('PCFOffice: Successfully loaded mailbox.item.body: ' + result.value);

                //Now, get Unique Body of the Mail Item (from Outlook REST API):
                console.log('PCFOffice: Getting Unique Body of the Mail Item (from Outlook REST API)');

                //Get the access token to call the Outlook REST API
                Office.context.mailbox.getCallbackTokenAsync(
                    {isRest: true}, 
                    (result : Office.AsyncResult<string>) => this.getCallbackTokenCallback(result)
                );
            } else {
                this.mailItemBody = undefined; // reset body in state
                const err = result.error;
                console.error(err.name + ": " + err.message);
            }

            this.raiseOfficeObjectUpdated();
        }

        public loadOfficeObject(): void { 
            let mailItem = Office.context?.mailbox?.item;
        
            console.log('PCFOffice: mailbox.item: ' + mailItem);
        
            this.mailItemBody = undefined; // reset body in state
            this.mailItemUniqueBody = undefined; // reset unique body in state
        
            //Save the Mail Item to the Add-in State
            this.mailItem = mailItem;

            //mailItem could be null - if a message isn't selected in Outlook - this is a valid state
        
            if(mailItem) {
                //Get the Body of the Mail Item
                mailItem.body.getAsync(
                    Office.CoercionType.Html,//addinState.bodyCoercionType,
                    (result : Office.AsyncResult<string>) => this.getMailBodyCallback(result)
                );
            }
        }

        //TODO: check if mailBody or uniqueBody is NULL? does this fail?
        public getOfficeObjectMessage() : IOfficeMessage {
            if(this.mailItem) {
                let mailItemObject = {
                    messageType : "mailbox.item.message",
                    itemId : this.mailItem.itemId,
                    body : this.mailItemBody,
                    uniqueBody : this.mailItemUniqueBody,
                    subject : this.mailItem.subject,
                    normalizedSubject : this.mailItem.normalizedSubject,
                    from : this.mailItem.from,
                    sender : this.mailItem.sender,
                    to : this.mailItem.to,
                    cc : this.mailItem.cc,
                    attachments : this.mailItem.attachments,
                    dateTimeCreated : this.mailItem.dateTimeCreated,
                };
        
                return mailItemObject
            }
            else{
                let nullMailItemObject = {
                    messageType : "mailbox.item.undefined",
                    itemId : null
                };
        
                return nullMailItemObject;
            }
        }

        private getItemRestId() : string {
            if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
                // itemId is already REST-formatted.
                return Office.context.mailbox.item.itemId;
            } else {
                // Convert to an item ID for API v2.0.
                return Office.context.mailbox.convertToRestId(
                    Office.context.mailbox.item.itemId,
                    Office.MailboxEnums.RestVersion.v2_0
                );
            }
        }

        public itemChanged(eventArgs : any) {
            console.log("PCFOffice: Add-in / Parent Window: ItemChanged event received - via itemChanged");
        
            this.loadOfficeObject();
        }
    }

    //Create a new instance of the OfficeBridge class to facilitate communication between the Add-in and the PCF Component (in the iframe)
    let bridge = new OfficeBridge();
    //TODO: consider moving the event handling to the Bridge class itself (or a separate class) - to keep the Bridge class focused on the communication between the Add-in and the PCF Component (in the iframe)

    //CORRECT FOR 'this' REFERENCE in Event Handler
    //https://github.com/Microsoft/TypeScript/wiki/'this'-in-TypeScript
    //window.addEventListener('click', x.printThing, 10); // DANGER, method is not invoked where it is referenced
    //window.addEventListener('click', () => x.printThing(), 10); // SAFE, method is invoked in the same expression
    //  **needed to change the Event Handler to an arrow function to ensure the correct 'this' reference is used
    //  so went from addEventListener("message", this.postMessageHander, false) >>> addEventListener("message", (event: any => this.postMessageHandler(event), false)
        
    //  Add the Event Handler for Cross-Domain / IFrame / Parent Window Messaging
    if (window.addEventListener) {
        console.log("PCFOffice: Adding event listener for postMessage event to Add-in/Parent - to receive messages from PCF/iframe/PowerApp - via window.addEventListener");
        window.addEventListener("message", (event : any) => bridge.postMessageHandler(event), false);
    }
    else if(window.attachEvent) 
    {
        console.log("PCFOffice: Adding event listener for postMessage event to Add-in/Parent - to receive messages from PCF/iframe/PowerApp - via window.attachEvent");
        window.attachEvent("onmessage", (event : any) => bridge.postMessageHandler(event));
    }
    else
    {
        //TODO: need to raise actual exception / error if event handler cannot be added
        console.error("Could not add event handler for postMessage");
    }

    //TODO: consider moving this office logic into teh Bridge class (will need to pass several items into the constructor)
    Office.onReady((info) => {
      console.log("Office.onReady event received");
      console.log('host: ' + info.host);

      //use in debug mode since hostType is null outside of office
      //bridge.officeReady = true;

      //TODO: add other Office apps (OneNote, etc.)
      //TODO: consider moving this logic based on Office App into the constructor of the Adapter itself? or in the Bridge?

      if (info.host === Office.HostType.Outlook) {
        bridge.officeReady = true;

        let outlookConnector = new OutlookAppAdapter(Office.context);

        bridge.setOfficeAdapter(outlookConnector);

        //Setup Event Handler for ItemChanged Event (if the Add-in is pinnable and mail item changes)
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, (asyncResult : Office.AsyncResult<void>) => outlookConnector.itemChanged(asyncResult));

        if(bridge.pcfReady) {
            //Get the Mail Item
            outlookConnector.getOfficeObjectMessage();
            bridge.sendOfficeObjectIfReady();
        }
      }
    });
}


