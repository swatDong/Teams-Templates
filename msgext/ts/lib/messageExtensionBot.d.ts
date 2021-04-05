import { ActivityHandler } from 'botbuilder';
export declare class MessageExtensionBot extends ActivityHandler {
    handleTeamsMessagingExtensionSubmitAction(context: any, action: any): {
        composeExtension: {
            type: string;
            attachmentLayout: string;
            attachments: {
                contentType: string;
                content: any;
                preview: import("botframework-schema").Attachment;
            }[];
        };
    };
    handleTeamsMessagingExtensionQuery(context: any, query: any): Promise<{
        composeExtension: {
            type: string;
            attachmentLayout: string;
            attachments: any[];
        };
    }>;
    handleTeamsMessagingExtensionSelectItem(context: any, obj: any): Promise<{
        composeExtension: {
            type: string;
            attachmentLayout: string;
            attachments: import("botframework-schema").Attachment[];
        };
    }>;
    handleTeamsAppBasedLinkQuery(context: any, query: any): {
        composeExtension: {
            attachmentLayout: string;
            type: string;
            attachments: import("botframework-schema").Attachment[];
        };
    };
}
