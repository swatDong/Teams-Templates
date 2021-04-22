// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, DialogSet, DialogTurnStatus, WaterfallDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { ActivityTypes, tokenExchangeOperationName } = require("botbuilder");

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const TEAMS_SSO_PROMPT_ID = "ModsSsoPrompt";

const { polyfills } = require('isomorphic-fetch');
const {
    createMicrosoftGraphClient,
    loadConfiguration,
    OnBehalfOfUserCredential,
    TeamsBotSsoPrompt
} = require("teamsdev-client");
const { MemoryStorage } = require('botbuilder-core');

class MainDialog extends LogoutDialog {
    constructor(dedupStorage) {
        super(MAIN_DIALOG, process.env.connectionName);
        this.requiredScopes = ["User.Read"]; // hard code the scopes for demo purpose only
        loadConfiguration();
        this.addDialog(new TeamsBotSsoPrompt(TEAMS_SSO_PROMPT_ID, {
            scopes: this.requiredScopes,
            endOnInvalidMessage: true
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.callApi.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
        this.dedupStorage = dedupStorage;
        this.dedupStorageKeys = [];
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} dialogContext
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

    async promptStep(stepContext) {
        try {
            return await stepContext.beginDialog(TEAMS_SSO_PROMPT_ID);
        } catch (err) {
            console.error(err);
        }
    }


    async callApi(stepContext) {
        // Get token response
        const tokenResponse = stepContext.result;

        if (tokenResponse) {
            await stepContext.context.sendActivity("You are now logged in.");

            await stepContext.context.sendActivity("Call Microsoft Graph on behalf of user...");

            // Call Microsoft Graph on behalf of user
            const oboCredential = new OnBehalfOfUserCredential(tokenResponse.ssoToken);
            const graphClient = createMicrosoftGraphClient(oboCredential, ["User.Read"]);
            const me = await graphClient.api("/me").get();
            if (me) {
                await stepContext.context.sendActivity(`You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${me.jobTitle}.`);

                // show user picture
                //var photoBuffer = await graphClient.api("/me/photo/$value").get();
                // const photoBuffer =await photoResponse.arrayBuffer();
                // const photoData = photoResponse.data;
                // const imageUri = 'data:image/png;base64,' + photoData.toString('base64');
                //const card = CardFactory.thumbnailCard("", CardFactory.images([imageUri]));
                // await stepContext.context.sendActivity({ attachments: [card] });
            }
            else {
                await stepContext.context.sendActivity("Getting profile from Microsoft Graph failed! ");
            }

            // Call API hosted in Azure Functions on behalf of user
            // const apiConfig = getResourceConfiguration(ResourceType.API);

            // const url = apiConfig.endpoint.replace(/\/$/, "");
            // const response = await axios.default.get(url, {
            //     headers: {
            //         authorization: "Bearer " + tokenResponse.ssoToken
            //     }
            // });
            // await stepContext.context.sendActivity(
            //     "Call API hosted in Azure Functions on behalf of user. API endpoint: " + apiConfig.endpoint
            // );
            // await stepContext.context.sendActivity("Response.data: " + JSON.stringify(response.data));

            // console.log(response);
            return await stepContext.endDialog();
        }

        await stepContext.context.sendActivity("Login was not successful please try again.");
        return await stepContext.endDialog();
    }

    async onEndDialog(context, instance, reason) {
        await this.dedupStorage.delete(this.dedupStorageKeys);
        this.dedupStorageKeys = [];
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
