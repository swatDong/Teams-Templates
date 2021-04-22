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
    constructor(conversationState, userState, dialog, storage) {
        super(conversationState, userState, dialog);
        this.storage = storage;
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    const welcomeMsg = MessageFactory.text(`Congratulations! ${username}, your hello world Bot 
                    template is running. This bot will introduce you how to build bot using Microsoft Teams App Framework(TeamsFx). 
                    You can reply ‘intro’ to see the introduction card. TeamsFx helps you build Bot using [Bot Framework SDK](https://dev.botframework.com/)`);
                    welcomeMsg.textFormat = 'markdown';
                    await stepContext.context.sendActivity(welcomeMsg);
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
        const storeItem = {
            eTag: context.activity.value.id,
        };
        const storeItems = { [this.getStorageKey(context)]: storeItem };

        try {
            await this.storage.write(storeItems);
        } catch (err) {
            if (err instanceof Error && err.message.indexOf("eTag conflict")) {
                return true;
            }
            throw err;
        }
        return false;
    }

    getStorageKey(context) {
        if (!context || !context.activity || !context.activity.conversation) {
            throw new Error("Invalid context, can not get storage key!");
        }
        const activity = context.activity;
        const channelId = activity.channelId;
        const conversationId = activity.conversation.id;
        if (
            activity.type !== ActivityTypes.Invoke ||
            activity.name !== tokenExchangeOperationName
        ) {
            throw new Error(
                "TokenExchangeState can only be used with Invokes of signin/tokenExchange."
            );
        }
        const value = activity.value;
        if (!value || !value.id) {
            throw new Error(
                "Invalid signin/tokenExchange. Missing activity.value.id."
            );
        }
        return `${channelId}/${conversationId}/${value.id}`;
    }
}

module.exports.TeamsBot = TeamsBot;