//TODO : remove incompatible properties from schema and examples
//https://learn.microsoft.com/en-us/power-apps/developer/component-framework/reference/control/getoutputschema

export const staticEmailAddressDetailsSchema = {
    "$schema": "https://json-schema.org/draft/2019-09/schema",
    "$id": "http://example.com/example.json",
    "type": "object",
    "default": {},
    "title": "Email Address Details Schema",
    "required": [
        "displayName",
        "emailAddress"
    ],
    "properties": {
        "displayName": {
            "type": "string",
            "default": "",
            "title": "The displayName Schema",
            "examples": [
                "John Doe"
            ]
        },
        "emailAddress": {
            "type": "string",
            "default": "",
            "title": "The emailAddress Schema",
            "examples": [
                "john.doe@mail.com"
            ]
        }
    },
    "examples": [{
        "displayName": "John Doe",
        "emailAddress": "john.doe@mail.com"
    }]
};

export const staticEmailAddressDetailsArraySchema = {
    "$schema": "https://json-schema.org/draft/2019-09/schema",
    "$id": "http://example.com/example.json",
    "type": "array",
    "default": [],
    "title": "Email Address Details Array Schema",
    "items": {
        "type": "object",
        "title": "A Schema",
        "required": [
            "displayName",
            "emailAddress"
        ],
        "properties": {
            "displayName": {
                "type": "string",
                "title": "The displayName Schema",
                "examples": [
                    "Jane Doe",
                    "Will Smithson"
                ]
            },
            "emailAddress": {
                "type": "string",
                "title": "The emailAddress Schema",
                "examples": [
                    "jane@acme.com",
                    "will@jiggy.com"
                ]
            }
        },
        "examples": [{
            "displayName": "Jane Doe",
            "emailAddress": "jane@acme.com"
        },
        {
            "displayName": "Will Smithson",
            "emailAddress": "will@jiggy.com"
        }]
    },
    "examples": [
        [{
            "displayName": "Jane Doe",
            "emailAddress": "jane@acme.com"
        },
        {
            "displayName": "Will Smithson",
            "emailAddress": "will@jiggy.com"
        }]
    ]
};

export const staticAttachmentsSchema = {
    "$schema": "https://json-schema.org/draft/2019-09/schema",
    "$id": "http://example.com/example.json",
    "type": "array",
    "default": [],
    "title": "Attachments Schema",
    "items": {
        "type": "object",
        "title": "A Schema",
        "required": [
            "attachmentType",
            "id",
            "isInline",
            "name",
            "size"
        ],
        "properties": {
            "attachmentType": {
                "type": "string",
                "title": "The attachmentType Schema",
                "examples": [
                    "file"
                ]
            },
            "id": {
                "type": "string",
                "title": "The id Schema",
                "examples": [
                    "AQMkADgzZDBlNGUxLTA2OGEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAEC3sR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAA4LcKn51qY0S6sNh0/GMGFA==",
                    "AQMkADgzZDBlNGUxLAdEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAESyIfd3OM535ASDAR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAAqMXTURHz0EGl1wUi4WUV+Q=="
                ]
            },
            "isInline": {
                "type": "boolean",
                "title": "The isInline Schema",
                "examples": [
                    true,
                    false
                ]
            },
            "name": {
                "type": "string",
                "title": "The name Schema",
                "examples": [
                    "image001.png",
                    "Sample FileName.pdf"
                ]
            },
            "size": {
                "type": "integer",
                "title": "The size Schema",
                "examples": [
                    3021,
                    521708
                ]
            }
        },
        "examples": [{
            "attachmentType": "file",
            "id": "AQMkADgzZDBlNGUxLTA2OGEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAEC3sR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAA4LcKn51qY0S6sNh0/GMGFA==",
            "isInline": true,
            "name": "image001.png",
            "size": 3021
        },
        {
            "attachmentType": "file",
            "id": "AQMkADgzZDBlNGUxLAdEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAESyIfd3OM535ASDAR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAAqMXTURHz0EGl1wUi4WUV+Q==",
            "isInline": false,
            "name": "Sample FileName.pdf",
            "size": 521708
        }]
    },
    "examples": [
        [{
            "attachmentType": "file",
            "id": "AQMkADgzZDBlNGUxLTA2OGEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAEC3sR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAA4LcKn51qY0S6sNh0/GMGFA==",
            "isInline": true,
            "name": "image001.png",
            "size": 3021
        },
        {
            "attachmentType": "file",
            "id": "AQMkADgzZDBlNGUxLAdEtNDY0Yy1hYjRkLWNjAGI3NwEyOTBmOTAARgAAA2pUhyzeCSVHsjA4wl2PCWAHAESyIfd3OM535ASDAR7f0AAAMoAAAA+oHozAWwBEWCVg7GilxQ2QAFXCoCkgAAAAESABAAqMXTURHz0EGl1wUi4WUV+Q==",
            "isInline": false,
            "name": "Sample FileName.pdf",
            "size": 521708
        }]
    ]
};  