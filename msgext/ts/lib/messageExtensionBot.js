"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MessageExtensionBot = void 0;
const axios_1 = require("axios");
const querystring = require("querystring");
const botbuilder_1 = require("botbuilder");
class MessageExtensionBot extends botbuilder_1.ActivityHandler {
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
    // Search.
    handleTeamsMessagingExtensionQuery(context, query) {
        return __awaiter(this, void 0, void 0, function* () {
            const searchQuery = query.parameters[0].value;
            const response = yield axios_1.default.get(`http://registry.npmjs.com/-/v1/search?${querystring.stringify({ text: searchQuery, size: 8 })}`);
            const attachments = [];
            response.data.objects.forEach(obj => {
                const heroCard = botbuilder_1.CardFactory.heroCard(obj.package.name);
                const preview = botbuilder_1.CardFactory.heroCard(obj.package.name);
                preview.content.tap = { type: 'invoke', value: { description: obj.package.description } };
                const attachment = Object.assign(Object.assign({}, heroCard), { preview });
                attachments.push(attachment);
            });
            return {
                composeExtension: {
                    type: 'result',
                    attachmentLayout: 'list',
                    attachments: attachments
                }
            };
        });
    }
    handleTeamsMessagingExtensionSelectItem(context, obj) {
        return __awaiter(this, void 0, void 0, function* () {
            return {
                composeExtension: {
                    type: 'result',
                    attachmentLayout: 'list',
                    attachments: [botbuilder_1.CardFactory.thumbnailCard(obj.description)]
                }
            };
        });
    }
    // Link Unfurling.
    handleTeamsAppBasedLinkQuery(context, query) {
        const attachment = botbuilder_1.CardFactory.thumbnailCard('Thumbnail Card', query.url, [query.url]);
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
}
exports.MessageExtensionBot = MessageExtensionBot;
function createCardCommand(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const heroCard = botbuilder_1.CardFactory.heroCard(data.title, data.text);
    heroCard.content.subtitle = data.subTitle;
    const attachment = {
        contentType: heroCard.contentType,
        content: heroCard.content,
        preview: heroCard,
    };
    return {
        composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [attachment]
        },
    };
}
function shareMessageCommand(context, action) {
    // The user has chosen to share a message by choosing the 'Share Message' context menu command.
    let userName = 'unknown';
    if (action.messagePayload &&
        action.messagePayload.from &&
        action.messagePayload.from.user &&
        action.messagePayload.from.user.displayName) {
        userName = action.messagePayload.from.user.displayName;
    }
    // This Messaging Extension example allows the user to check a box to include an image with the
    // shared message.  This demonstrates sending custom parameters along with the message payload.
    let images = [];
    const includeImage = action.data.includeImage;
    if (includeImage === 'true') {
        images = [
            'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU',
        ];
    }
    const heroCard = botbuilder_1.CardFactory.heroCard(`${userName} originally sent this message:`, action.messagePayload.body.content, images);
    if (action.messagePayload &&
        action.messagePayload.attachment &&
        action.messagePayload.attachments.length > 0) {
        // This sample does not add the MessagePayload Attachments.  This is left as an
        // exercise for the user.
        heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
    }
    const attachment = {
        contentType: heroCard.contentType,
        content: heroCard.content,
        preview: heroCard
    };
    return {
        composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [attachment]
        },
    };
}
//# sourceMappingURL=messageExtensionBot.js.map