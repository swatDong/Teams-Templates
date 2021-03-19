// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { ActivityHandler, MessageFactory, CardFactory, TurnContext } from 'botbuilder';

export class EchoBot extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            const replyText = `Echo: ${ context.activity.text }`;
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    public async handleTeamsMessagingExtensionSubmitAction( context: TurnContext, action: any ): Promise<any> {
        switch ( action.commandId ) {
            case 'createCard':
            return createCardCommand( context, action );
            case 'shareMessage':
            return shareMessageCommand( context, action );
            default:
            throw new Error( 'NotImplemented' );
        }
    }

        // This handler is used for the processing of "composeExtension/queryLink" activities from Teams.
    // https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions#receive-requests-from-links-inserted-into-the-compose-message-box
    // By specifying domains under the messageHandlers section in the manifest, the bot can receive
    // events when a user enters in a domain in the compose box.
    handleTeamsAppBasedLinkQuery(context, query) {
        const attachment = CardFactory.thumbnailCard('Thumbnail Card',
            query.url,
            ['https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png']);

        const result = {
            attachmentLayout: 'list',
            type: 'result',
            attachments: [attachment]
        };

        const response = {
            composeExtension: result
        };
        return response;
    }

    async handleTeamsMessagingExtensionQuery(context, query) {
        // Note: The Teams manifest.json for this sample also inclues a Search Query, in order to enable installing from App Studio.
        // const searchQuery = query.parameters[0].value;
        const heroCard = CardFactory.heroCard('This is a Link Unfurling Sample',
            'This sample demonstrates how to handle link unfurling in Teams.  Please review the readme for more information.');
        heroCard.content.subtitle = 'It will unfurl links from *.BotFramework.com';
        const attachment = { ...heroCard, heroCard };

        switch (query.commandId) {
        case 'searchQuery':
            return {
                composeExtension: {
                    type: 'result',
                    attachmentLayout: 'list',
                    attachments: [
                        attachment
                    ]
                }
            };
        default:
            throw new Error('NotImplemented');
        }
    }
}

async function createCardCommand( context: TurnContext, action: any ): Promise<any> {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const heroCard = CardFactory.heroCard( data.title, data.text );
    heroCard.content.subtitle = data.subTitle;
    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };
  
    return {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'list',
        attachments: [
          attachment
        ]
      }
    };
}
  
async function shareMessageCommand( context: TurnContext, action: any ): Promise<any> {
    // The user has chosen to share a message by choosing the 'Share Message' context menu command.
    let userName = 'unknown';
    if ( action.messagePayload.from &&
      action.messagePayload.from.user &&
      action.messagePayload.from.user.displayName ) {
      userName = action.messagePayload.from.user.displayName;
    }
  
    // This Messaging Extension example allows the user to check a box to include an image with the
    // shared message.  This demonstrates sending custom parameters along with the message payload.
    let images = [];
    const includeImage = action.data.includeImage;
    if ( includeImage === true ) {
      images = [ 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU' ];
    }
    const heroCard = CardFactory.heroCard( `${ userName } originally sent this message:`,
      action.messagePayload.body.content,
      images );
  
    if ( action.messagePayload.attachments && action.messagePayload.attachments.length > 0 ) {
      // This sample does not add the MessagePayload Attachments.  This is left as an
      // exercise for the user.
      heroCard.content.subtitle = `(${ action.messagePayload.attachments.length } Attachments not included)`;
    }
  
    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };
  
    return {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'list',
        attachments: [
          attachment
        ]
      }
    };
}
