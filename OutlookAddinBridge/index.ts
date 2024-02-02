import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { staticEmailAddressDetailsSchema, staticEmailAddressDetailsArraySchema, staticAttachmentsSchema } from "./staticSchema";

export class OutlookAddinBridge implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // Reference to ComponentFramework Context object
    private _context: ComponentFramework.Context<IInputs>;

    // PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;

    private _container : HTMLDivElement;
    private _content : HTMLDivElement;

    //Input/Configuration properties
    private _showOutputs : boolean;

    //Outlook properties
    private _mailItemId? : string;
    private _subject? : string;
    private _normalizedSubject? : string;
    private _body? : string;
    private _uniqueBody? : string;
    private _from? : any; //TODO: define the type for this (and other email address details - eg sender, to, cc)
    private _sender? : any;
    private _to? : any[];
    private _cc? : any[];
    private _attachments? : any[]; //TODO: define the type for this
    private _dateTimeCreated : Date;

    private sendMessageToParent(message : any) {
        console.log("PCFOffice: PCF Control: Sending message to Parent/Office Add-in - via sendMessageToParent. Message: " + message);
    
        if (window && window.parent && window.parent.parent) {
            window.parent.parent.postMessage(message, '*');
            //TODO: consider setting the sender origin to the current window origin (eg Powerapps.com)        
        }

        //TODO: log error / raise error if window.parent.parent is null - since cannot call postMessage on null
    }

    //*****************************
    //EVENT HANDLER FOR CROSS-DOMAN / IFRAME MESSAGING
    //*****************************
    private postMessageHandler( event : any ) {
        //ONLY log to console if this is a message from the Add-in/Office (eg messageType is defined)

        //TODO: include more definitive info (like a 'schema name like 'PCFOffice' or 'PCFOutlookAddinBridge') to ensure the message is from the correct source (eg Add-in/Office)
        //TODO: enforce version consistency between the Add-in/Office and the PCF Control (in the message object - include version)
    
        // check request is from legitimate source and message is expected or not
        //if ( event.origin !== 'https://powerapps.com' ) { console.log("not matching right domain");return; }

        // check the message type received (and take appropriate action)
        if(event.data.messageType) {//if (event.data.messageType is defined)
            //Output Debug Information to Console
            console.log("PCFOffice: PCF/iframe/PowerApp: Received message from Add-in/Office - via postMessageHandler");
            console.log("* Message:", event.data); //TODO: remove this line (since it could be sensitive data - eg the email message or recipients)
            console.log("* Origin:", event.origin);
            console.log("* Source:", event.source);

            switch(event.data.messageType) {
                case "mailbox.item.message":
                    // Received message from Addin/Office - with the current mailbox.item
                    console.log("PCFOffice: mailbox.item.message message received from Add-in/Office");

                    this._mailItemId = event.data?.itemId;
                    this._body = event.data?.body;
                    this._uniqueBody = event.data?.uniqueBody;
                    this._subject = event.data?.subject;
                    this._normalizedSubject = event.data?.normalizedSubject;
                    this._from = event.data?.from;
                    this._sender = event.data?.sender;
                    this._to = event.data?.to;
                    this._cc = event.data?.cc;
                    this._attachments = event.data?.attachments;
                    this._dateTimeCreated = new Date(event.data.dateTimeCreated);

                    console.log("PCFOffice: Subject Received from Add-in/Office: " + this._subject);

                    this._notifyOutputChanged();
                    this._context.factory.requestRender();
                    
                    break;

                case "mailbox.item.undefined":
                    // Received message from Addin/Office - with the current mailbox.item is null / not defined
                    console.log("PCFOffice: mailbox.item.undefined message received from Add-in/Office");

                    this._mailItemId = event.data?.itemId;
                    this._body = event.data?.body;
                    this._uniqueBody = event.data?.uniqueBody;
                    this._subject = event.data?.subject;
                    this._normalizedSubject = event.data?.normalizedSubject;
                    this._from = event.data?.from;
                    this._sender = event.data?.sender;
                    this._to = event.data?.to;
                    this._cc = event.data?.cc;
                    this._attachments = event.data?.attachments;
                    this._dateTimeCreated = new Date(event.data?.dateTimeCreated); //todo: make sure this works

                    this._notifyOutputChanged();
                    this._context.factory.requestRender();
                    
                    break;
            }
            
        }
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._content = document.createElement("div");
        this._container.setAttribute("style", "text-align: left;background-color: white");
        this._container.appendChild(this._content);
        
        this._showOutputs = this._context.parameters.showOutputs.raw || false;

        //this._domId = this.ID();

        //TODO: this is for testing - REMOVE
        /*
        this._from = {displayName: "John Doe", emailAddress: "john.doe@mail.com"};
        this._sender = {displayName: "John Doe", emailAddress: "john.doe@mail.com"};
        this._to = [{displayName: "Jane Doe", emailAddress: "jane@acme.com"}, {displayName: "Will Smithson", emailAddress: "will@jiggy.com"}];
        this._cc = [];
        this._attachments = [
            {
                attachmentType : "file",
                id : "AQMkADgzZDBlNGUxLTA2OGEtNDASD8391hYjRkLWNjAGI3NwEyOTasd0123AA2pUhyzeCSVHsjA4wl2PCWAHAEC3sR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAA4LcKn51qY0S6sNh0/GMGFA==",
                isInline : true,
                name : "image001.png",
                size : 3021
            },
            {
                attachmentType : "file",
                id : "AQMkADgzZDBlNGUxLAdEtNDY0Yy1hYjRkLWNjAGI3asd12emOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAESyIfd3OM535ASDAR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAAqMXTURHz0EGl1wUi4WUV+Q==",
                isInline : false,
                name : "Sample FileName.pdf",
                size : 521708
            }
        ];

        this._body = ""

        this._mailItemId = "8C7D3FA4-386D-5656-9D98-344B388C07CE";
        this._subject = "RE: Test Subject";
        this._normalizedSubject = "Test Subject";
        this._hasAttachments = true; //event.data?.hasAttachments; //TODO: make this false even if null
        this._attachmentCount = 2; //event.data?.attachmentCount; //TODO: make this zero if null
        this._dateTimeCreated = new Date(); //todo: make sure this works
        */
        //TODO: end REMOVE


        //Track container resize
        this._context.mode.trackContainerResize(false);

        //CORRECT FOR THIS REFERENCE
        //https://github.com/Microsoft/TypeScript/wiki/'this'-in-TypeScript
        //window.addEventListener('click', x.printThing, 10); // DANGER, method is not invoked where it is referenced
        //window.addEventListener('click', () => x.printThing(), 10); // SAFE, method is invoked in the same expression
        //  **needed to change the Event Handler to an arrow function to ensure the correct 'this' reference is used
        //  so went from addEventListener("message", this.postMessageHander, false) >>> addEventListener("message", (event: any => this.postMessageHandler(event), false)
        //  Add the Event Handler for Cross-Domain / IFrame / Parent Window Messaging
        if (window.addEventListener) {
            console.log("PCFOffice: Adding event listener for postMessage event to PCF/iframe/PowerApp - to receive messages from Add-in/Office - via window.addEventListener");
            window.addEventListener("message", (event : any) => this.postMessageHandler(event), false);
        }
        /* 
        TODO: need to fix since it throws an error building component - not defined on Window object in TypeScript error
        else if(window.attachEvent) 
        {
            console.log("PCFOffice: Adding event listener for postMessage event to PCF/iframe/PowerApp - to receive messages from Add-in/Office - via window.attachEvent");
            window.attachEvent("onmessage", this.postMessageHandler);
        }
        */
        else
        {
            console.error("Could not add event handler for postMessage");
        }

        //Notify the parent (Office Add-in) that the PCF Control is ready to receive messages
        this.sendMessageToParent({messageType: "PCF.Ready"});
    }

    /*
    ID = function () {
		// Math.random should be unique because of its seeding algorithm.
		// Convert it to base 36 (numbers + letters), and grab the first 9 characters
		// after the decimal.
		return '_' + Math.random().toString(36).substr(2, 9);


        //todo: replace substr function with current / non-deprecated method
	};*/

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        console.log('Called UpdateView');
        // Code to update control view
    
        // store the latest context
		this._context = context;
        
        this._showOutputs = this._context.parameters.showOutputs.raw || false;

        if(this._showOutputs) {
            //TODO: #3 add background color to the output div

            //TODO: add font color to the output div
            
            //TODO: #4 add properties (input) to the control to allow for the user to set the font color and background color

            var outputString = 
                "<span>Mail Item Id: " + this._mailItemId + "</span><br/>" +
                "<span>Subject: " + this._subject + "</span><br/>" + 
                "<span>Normalized Subject: " + this._normalizedSubject + "</span><br/>" +
                "<span>From: " + this._from?.displayName + " (" + this._from?.emailAddress + ")</span><br/>" +
                "<span>Sender: " + this._sender?.displayName + " (" + this._sender?.emailAddress + ")</span><br/>" +
                "<span>To: " + this._to?.map((item) => item?.displayName + " (" + item?.emailAddress + ")").join(", ") + "</span><br/>" +
                "<span>Cc: " + this._cc?.map((item) => item?.displayName + " (" + item?.emailAddress + ")").join(", ") + "</span><br/>" +
                "<span>Attachments: " + this._attachments?.map((item) => item?.name).join(", ") + "</span><br/>" +
                "<span>Date Created: " + this._dateTimeCreated + "</span><br/>" +
                "<span>Body: " + this._body + "</span><br/>" +
                "<span>Unique Body: " + this._uniqueBody + "</span><br/>";

            this._content.innerHTML = outputString;
        }
        else {
            this._content.innerHTML = "";
        }

        //TODO: is this actually needed - since typically calls from NotifyOutputChanges calls UpdateView
        this._notifyOutputChanged();
    }

    /**
     * It is called by the framework prior to a control init to get the output object(s) schema
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns an object schema based on nomenclature defined in manifest
     */
    public async getOutputSchema(context: ComponentFramework.Context<IInputs>): Promise<Record<string, unknown>> {
        return Promise.resolve({
            from: staticEmailAddressDetailsSchema,
            sender: staticEmailAddressDetailsSchema,
            to: staticEmailAddressDetailsArraySchema,
            cc: staticEmailAddressDetailsArraySchema,
            attachments: staticAttachmentsSchema,
        });
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            mailItemId: this._mailItemId,
            subject: this._subject,
            normalizedSubject: this._normalizedSubject,
            body: this._body,
            uniqueBody: this._uniqueBody,
            from: this._from,
            sender: this._sender,
            to: this._to,
            cc: this._cc,
            attachments: this._attachments,
            hasAttachments: this._attachments && this._attachments.length > 0,
            attachmentCount: this._attachments?.length,
            dateTimeCreated: this._dateTimeCreated
		} as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
        window.removeEventListener('message', (event : any) => this.postMessageHandler(event));

        //TODO : remove others resources / handlers
    }
}