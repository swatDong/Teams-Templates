// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const axios = require('axios');
const querystring = require('querystring');
const { TeamsActivityHandler, CardFactory, TeamsInfo } = require('botbuilder');

class BotActivityHandler extends TeamsActivityHandler {

    // Action.
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


    // Search and Link Unfurling.

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