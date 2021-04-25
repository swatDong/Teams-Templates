// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ComponentDialog } = require('botbuilder-dialogs');
const { TurnContext, ActionTypes, CardFactory } = require('botbuilder');

class RootDialog extends ComponentDialog {
    constructor(id) {
        super(id);
    }

    async onBeginDialog(innerDc, options) {
        const result = await this.triggerCommand(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }

    async onContinueDialog(innerDc) {
        return await super.onContinueDialog(innerDc);
    }

    async triggerCommand(innerDc) {
        const removedMentionText = TurnContext.removeRecipientMention(innerDc.context.activity, innerDc.context.activity.recipient.id);
        const text = removedMentionText?.toLowerCase().replace(/\n|\r/g, '');    // Remove the line break           
        switch (text) {
            case 'show':
                break;
            case 'intro': {
                const cardButtons = [{ type: ActionTypes.ImBack, title: 'Show Profile', value: 'show' }];
                const card = CardFactory.heroCard(
                    'Introduction',
                    null,
                    cardButtons,
                    {
                        text: `This Bot has implemented single sign-on (SSO) using Teams Account 
                            which user logged in Teams client, check <a href=\"placeholder\">TeamsFx authentication document</a> 
                            and code in <pre>bot/dialogs/mainDialog.js</pre> to learn more about SSO.
                            Type <strong>show</strong> or click the button below to show your profile by calling Microsoft Graph API with SSO.
                            To learn more about building Bot using Microsoft Teams App Framework(TeamsFx), please refer to the <a href=\"placeholder\">TeamsFx document</a> .`
                    });

                await innerDc.context.sendActivity({ attachments: [card] });
                return await innerDc.cancelAllDialogs();
            }
            default: {
                const cardButtons = [{ type: ActionTypes.ImBack, title: 'Show introduction card', value: 'intro' }];
                const card = CardFactory.heroCard(
                    '',
                    null,
                    cardButtons,
                    {
                        text: `This is a hello world Bot built by Microsoft Teams App Framework(TeamsFx), 
                            which is designed only for illustration Bot purpose. This Bot by default will not handle any specific question or task. 
                            Please type <strong>intro</strong> to see the introduction card.`
                    });
                await innerDc.context.sendActivity({ attachments: [card] });
                return await innerDc.cancelAllDialogs();
            }
        }
    }
}

module.exports.RootDialog = RootDialog;
