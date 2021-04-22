// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { DialogSet, DialogTurnStatus, WaterfallDialog } = require('botbuilder-dialogs');
const { RootDialog } = require('./rootDialog');

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

class MainDialog extends RootDialog {
    constructor() {
        super(MAIN_DIALOG, process.env.connectionName);
        this.requiredScopes = ["User.Read"]; // hard code the scopes for demo purpose only
        loadConfiguration();
        this.addDialog(new TeamsBotSsoPrompt(TEAMS_SSO_PROMPT_ID, {
            scopes: this.requiredScopes,
            endOnInvalidMessage: true
        }));

        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.ssoStep.bind(this),
            this.showUserInfo.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
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


    async showUserInfo(stepContext) {
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

            return await stepContext.endDialog();
        }

        await stepContext.context.sendActivity("Login was not successful please try again.");
        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
