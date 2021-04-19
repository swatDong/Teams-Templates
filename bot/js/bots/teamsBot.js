// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes, tokenExchangeOperationName } = require("botbuilder");
const { DialogBot } = require("./dialogBot");


class TeamsBot extends DialogBot {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);
        this.conversationState= conversationState;
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity('Welcome to TeamsBot. Type \'login\' to get logged in. Type \'logout\' to sign-out.');
                }
            }

            await next();
        });
    }

    async handleTeamsSigninVerifyState(context, query) {
        console.log('Running dialog with signin/verifystate from an Invoke Activity.');
        await this.dialog.run(context, this.dialogState);
    }

    async handleTeamsSigninTokenExchange(context, query) {
        await this.dialog.run(context, this.dialogState);
    }

    async onTokenResponseEvent(context) {
        console.log('Running dialog with Token Response Event Activity.');

        // Run the Dialog with the new Token Response Event Activity.
        await this.dialog.run(context, this.dialogState);
    }

    async onSignInInvoke(context) {
        if (
            context.activity &&
            context.activity.name === tokenExchangeOperationName
        ) {
            if (await this.shouldDedup(context)) {
                return;
            }
        }
        await this.dialog.run(context, this.dialogState);
    }

    // If a user is signed into multiple Teams clients, the Bot might receive a "signin/tokenExchange" from each client.
    // Each token exchange request for a specific user login will have an identical activity.value.Id.
    // Only one of these token exchange requests should be processed by the bot.  For a distributed bot in production,
    // this requires a distributed storage to ensure only one token exchange is processed.
    async shouldDedup(context) {
        const storeItem = { [context.activity.value.id]: "" };

        const state = this.conversationState.createProperty(
            context.activity.value.id
        );

        const value = await state.get(context);
        if (value) {
            return true;
        }
        await state.set(context, storeItem);
        await this.conversationState.saveChanges(context, false);

        return false;
    }
}

module.exports.TeamsBot = TeamsBot;