// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');

/*  
    Conversation Bot
    Teams bots are Microsoft Bot Framework bots.
    If a bot receives a message activity, the turn handler sees that incoming activity
    and sends it to the onMessage activity handler.
    Learn more: https://aka.ms/teams-bot-basics.

    NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
            registered with Bot Framework.
            Learn more: https://aka.ms/teams-register-bot. 
*/
class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();

        /* Registers an activity event handler for the message event, emitted for every incoming message activity. */
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            switch (context.activity.text.trim()) {
                case 'Hello':
                    await this.mentionActivityAsync(context);
                    break;
                default:
                    /* By default for unknown activity sent by user show a card with the available actions. */
                    const value = { count: 0 };
                    const card = CardFactory.heroCard(
                        'Lets talk...',
                        null,
                        [{
                            type: ActionTypes.MessageBack,
                            title: 'Say Hello',
                            value: value,
                            text: 'Hello'
                        }]);
                    await context.sendActivity({ attachments: [card] });
                    break;
            }
            await next();
        });
    }

    /* 
        Conversation Bot 
        Say hello and @ mention the current user.
    */
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${mention.text}`);
        replyActivity.entities = [mention];

        await context.sendActivity(replyActivity);
    }

    handleTeamsMessagingExtensionSubmitAction(context, action) {
        switch (action.commandId) {
            case 'createCard':
                return createCardCommand(context, action);
            case 'shareMessage':
                return shareMessageCommand(context, action);
            default:
                throw new Error('NotImplemented');
        }
    }

    async handleTeamsMessagingExtensionFetchTask(context, action) {
        try {
           const member = await this.getSingleMember(context);
           return {
                task: {
                    type: 'continue',
                    value: {
                        card: GetAdaptiveCardAttachment(),
                        height: 400,
                        title: 'Hello ' + member,
                        width: 300
                    },
                },
            };
        } catch (e) {
            if (e.code === 'BotNotInConversationRoster') {
                return {
                    task: {
                        type: 'continue',
                        value: {
                            card: GetJustInTimeCardAttachment(),
                            height: 400,
                            title: 'Adaptive Card - App Installation',
                            width: 300
                        },
                    },
                };
            }
            throw e;
        }
       
    }
    async getSingleMember(context) {
        try {
           const member = await TeamsInfo.getMember(
                context,
                context.activity.from.id
            );
           return member.name;
        } catch (e) {
            if (e.code === 'MemberNotFoundInConversation') {
                context.sendActivity(MessageFactory.text('Member not found.'));
                return e.code;
            } 
            throw e;        
        }   
    }

    async handleTeamsMessagingExtensionQuery(context, query) {
        const searchQuery = query.parameters[0].value;
        const response = await axios.get(`http://registry.npmjs.com/-/v1/search?${ querystring.stringify({ text: searchQuery, size: 8 }) }`);

        const attachments = [];
        response.data.objects.forEach(obj => {
            const heroCard = CardFactory.heroCard(obj.package.name);
            const preview = CardFactory.heroCard(obj.package.name);
            preview.content.tap = { type: 'invoke', value: { description: obj.package.description } };
            const attachment = { ...heroCard, preview };
            attachments.push(attachment);
        });

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: attachments
            }
        };
    }

    async handleTeamsMessagingExtensionSelectItem(context, obj) {
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [CardFactory.thumbnailCard(obj.description)]
            }
        };
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

module.exports.BotActivityHandler = BotActivityHandler;
