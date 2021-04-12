// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, ActionTypes } = require('botbuilder');
class DialogBot extends TeamsActivityHandler {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState, dialog) {
        super();
        if (!conversationState) {
            throw new Error('[DialogBot]: Missing parameter. conversationState is required');
        }
        if (!userState) {
            throw new Error('[DialogBot]: Missing parameter. userState is required');
        }
        if (!dialog) {
            throw new Error('[DialogBot]: Missing parameter. dialog is required');
        }

        this.conversationState = conversationState;
        this.userState = userState;
        this.dialog = dialog;
        this.dialogState = this.conversationState.createProperty('DialogState');

        this.onMessage(async (context, next) => {
            switch (context.activity.text) {
                case 'login':
                case 'logout': {
                    console.log('Running dialog with Message Activity.');

                    // Run the Dialog with the new message Activity.
                    await this.dialog.run(context, this.dialogState);

                    break;
                }
                default: {
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
            }

            await next();
        });
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
    }
}

module.exports.DialogBot = DialogBot;