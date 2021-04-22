// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes } = require('botbuilder');
const { ComponentDialog } = require('botbuilder-dialogs');
const { MessageFactory, TurnContext } = require('botbuilder');

class RootDialog extends ComponentDialog {
    constructor(id, connectionName) {
        super(id);
        this.connectionName = connectionName;
    }

    async onBeginDialog(innerDc, options) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }

    async onContinueDialog(innerDc) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        const removedMentionText = TurnContext.removeRecipientMention(innerDc.context.activity, innerDc.context.activity.recipient.id);
        if (removedMentionText) {
            const text = removedMentionText.toLowerCase().replace(/\n|\r/g, '');    // Remove the line break           
            switch (text) {
                case 'logout': {
                    // The bot adapter encapsulates the authentication processes.
                    const botAdapter = innerDc.context.adapter;
                    await botAdapter.signOutUser(innerDc.context, this.connectionName);
                    await innerDc.context.sendActivity('You have been signed out.');
                    return await innerDc.cancelAllDialogs();
                }
                case 'login':
                    break;
                case 'intro': {
                    const introMessage = MessageFactory.text(`This Bot has implemented single sign-on (SSO) using Teams Account 
                            which user logged in Teams client, check TeamsFx authentication document here <link> 
                            and code in \`bot/dialogs/mainDialog.js\` to learn more about SSO.
                            Type \`log out\` to try log out the bot. And type \`log in\` to log in again. 
                            To learn more about building Bot using Microsoft Teams App Framework(TeamsFx), please refer to the [document](https://review.docs.microsoft.com/en-us/mods/build-your-first-app/build-bot?branch=main).`);
                    introMessage.textFormat = 'markdown';
                    await innerDc.context.sendActivity(introMessage);
                    return await innerDc.cancelAllDialogs();
                }
                default: {
                    await innerDc.context.sendActivity(`This is a hello world Bot built by Microsoft Teams App Framework(TeamsFx), 
                            which is designed only for illustration Bot purpose. This Bot by default will not handle any specific question or task. 
                            Please type \`intro\` to see the introduction card.`);
                    return await innerDc.cancelAllDialogs();
                }
            }

        }
    }
}

module.exports.LogoutDialog = RootDialog;
