// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { DialogSet, DialogTurnStatus, WaterfallDialog, ConfirmPrompt } = require('botbuilder-dialogs');
const { RootDialog } = require('./rootDialog');
const { tokenExchangeOperationName, ActivityTypes } = require("botbuilder");

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const TEAMS_SSO_PROMPT_ID = "TeamsFxSsoPrompt";

const { polyfills } = require('isomorphic-fetch');
const {
    createMicrosoftGraphClient,
    loadConfiguration,
    OnBehalfOfUserCredential,
    TeamsBotSsoPrompt
} = require("teamsdev-client");

class MainDialog extends RootDialog {
    constructor(dedupStorage) {
        super(MAIN_DIALOG, "");
        this.requiredScopes = ["User.Read"]; // hard code the scopes for demo purpose only
        loadConfiguration();
        this.addDialog(new TeamsBotSsoPrompt(TEAMS_SSO_PROMPT_ID, {
            scopes: this.requiredScopes,
            endOnInvalidMessage: true
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.ssoStep.bind(this),
            this.showInfoPrompt.bind(this),
            this.ensureOAuth.bind(this),
            this.showUserInfo.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
        this.dedupStorage = dedupStorage;
        this.dedupStorageKeys = [];
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} dialogContext
     * @param {*} accessor
     */
    async run(context, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async ssoStep(stepContext) {
        try {
            return await stepContext.beginDialog(TEAMS_SSO_PROMPT_ID);
        } catch (err) {
            console.error(err);
        }
    }

    async showInfoPrompt(stepContext) {
        // Get the token from the previous step.
        const tokenResponse = stepContext.result;
        if (tokenResponse) {
            return await stepContext.prompt(CONFIRM_PROMPT, 'You are now logged in. Would you like to view your token?');
        }
        await stepContext.context.sendActivity('Login was not successful please try again.');
        return await stepContext.endDialog();
    }

    async ensureOAuth(stepContext) {
        await stepContext.context.sendActivity('Thank you.');

        const result = stepContext.result;
        if (result) {
            // Call the prompt again because we need the token. The reasons for this are:
            // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
            // about refreshing it. We can always just call the prompt again to get the token.
            // 2. We never know how long it will take a user to respond. By the time the
            // user responds the token may have expired. The user would then be prompted to login again.
            //
            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed.
            return await stepContext.beginDialog(TEAMS_SSO_PROMPT_ID);
        }
        return await stepContext.endDialog();
    }


    async showUserInfo(stepContext) {
        // Get token response
        const tokenResponse = stepContext.result;

        if (tokenResponse) {
            await stepContext.context.sendActivity("Call Microsoft Graph on behalf of user...");

            // Call Microsoft Graph on behalf of user
            const oboCredential = new OnBehalfOfUserCredential(tokenResponse.ssoToken);
            const graphClient = createMicrosoftGraphClient(oboCredential, ["User.Read"]);
            const me = await graphClient.api("/me").get();
            if (me) {
                await stepContext.context.sendActivity(`You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${me.jobTitle}.`);

                // show user picture
                var photoBinary = await graphClient.api("/me/photo/$value").get();
                //const photoBuffer =await photoResponse.arrayBuffer();
                const buffer = Buffer.from(photoBinary);
                const imageUri = 'data:image/png;base64,' + buffer.toString('base64');
                const card = CardFactory.thumbnailCard("", CardFactory.images([imageUri]));
                await stepContext.context.sendActivity({ attachments: [card] });
            }
            else {
                await stepContext.context.sendActivity("Getting profile from Microsoft Graph failed! ");
            }

            return await stepContext.endDialog();
        }

        await stepContext.context.sendActivity("Login was not successful please try again.");
        return await stepContext.endDialog();
    }

    async onEndDialog(context, instance, reason) {
        const conversationId = context.activity.conversation.id;
        const currentDedupKeys = this.dedupStorageKeys.filter(key=>key.indexOf(conversationId) > 0);
        await this.dedupStorage.delete(currentDedupKeys);
        this.dedupStorageKeys = this.dedupStorageKeys.filter(key=>key.indexOf(conversationId) < 0);
    }

    // If a user is signed into multiple Teams clients, the Bot might receive a "signin/tokenExchange" from each client.
    // Each token exchange request for a specific user login will have an identical activity.value.Id.
    // Only one of these token exchange requests should be processed by the bot.  For a distributed bot in production,
    // this requires a distributed storage to ensure only one token exchange is processed.
    async shouldDedup(context) {
        const storeItem = {
            eTag: context.activity.value.id,
        };

        const key = this.getStorageKey(context);
        const storeItems = { [key]: storeItem };

        try {
            await this.dedupStorage.write(storeItems);
            this.dedupStorageKeys.push(key);
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

module.exports.MainDialog = MainDialog;
