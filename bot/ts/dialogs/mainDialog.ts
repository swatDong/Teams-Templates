import { ConfirmPrompt, DialogSet, DialogTurnStatus, WaterfallDialog } from "botbuilder-dialogs";
import { LogoutDialog } from "./logoutDialog";
import { TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  getResourceConfiguration,
  loadConfiguration,
  OnBehalfOfUserCredential,
  ResourceType,
  TeamsBotSsoPrompt,
  TeamsBotSsoPromptTokenResponse
} from "teamsdev-client";
import * as axios from "axios";
import "isomorphic-fetch";

const CONFIRM_PROMPT = "ConfirmPrompt";
const MAIN_DIALOG = "MainDialog";
const MAIN_WATERFALL_DIALOG = "MainWaterfallDialog";
const TEAMS_SSO_PROMPT_ID = "TeamsFxSsoPrompt";

export class MainDialog extends LogoutDialog {
  private requiredScopes: string[] = ["User.Read"]; // hard code the scopes for demo purpose only

  // Developer controlls the lifecycle of credential provider, as well as the cache in it.
  // In this sample the provider is shared in all conversations
  constructor() {
    super(MAIN_DIALOG, process.env.connectionName);
    loadConfiguration();
    this.addDialog(
      new TeamsBotSsoPrompt(TEAMS_SSO_PROMPT_ID, {
        scopes: this.requiredScopes,
        endOnInvalidMessage: true
      })
    );

    //this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
    this.addDialog(
      new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
        this.promptStep.bind(this),
        this.callApi.bind(this)
      ])
    );

    this.initialDialogId = MAIN_WATERFALL_DIALOG;
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   * If no dialog is active, it will start the default dialog.
   * @param {*} dialogContext
   */
  async run(context: TurnContext, accessor: any) {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);

    const dialogContext = await dialogSet.createContext(context);
    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    }
  }

  async promptStep(stepContext: any) {
    return await stepContext.beginDialog(TEAMS_SSO_PROMPT_ID);
  }

  async testBot(stepContext: any) {
    await stepContext.context.sendActivity("testbot");
    return await stepContext.endDialog();
  }

  // Following function shows how to call API with TeamsFx support (token credential, etc.)
  async callApi(stepContext: any) {
    // Get token response
    const tokenResponse = stepContext.result as TeamsBotSsoPromptTokenResponse;

    if (tokenResponse) {
      await stepContext.context.sendActivity("You are now logged in.");

      await stepContext.context.sendActivity("Call Microsoft Graph on behalf of user...");

      // Call Microsoft Graph on behalf of user
      const oboCredential = new OnBehalfOfUserCredential(tokenResponse.ssoToken);
      const graphClient = createMicrosoftGraphClient(oboCredential, ["User.Read"]);
      const me = await graphClient.api("/me").get();
      if (me) {
        await stepContext.context.sendActivity(`You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${me.jobTitle}; your photo is: `);

        // show user picture
        //var photoBuffer = await graphClient.api("/me/photo/$value").get();
        //const photoBuffer = await photoResponse.arrayBuffer();
        //const imageUri = 'data:image/png;base64,' + photoData.toString('base64');
        //const card = CardFactory.thumbnailCard("", CardFactory.images([imageUri]));
        //await stepContext.context.sendActivity({ attachments: [card] });
      }
      else {
        await stepContext.context.sendActivity("Getting profile from Microsoft Graph failed! ");
      }

      // // Call API hosted in Azure Functions on behalf of user
      // const apiConfig = getResourceConfiguration(ResourceType.API);

      // const url = apiConfig.endpoint.replace(/\/$/, "") + "api/httpTrigger1";
      // const response = await axios.default.get(url, {
      //   headers: {
      //     authorization: "Bearer " + tokenResponse.ssoToken
      //   }
      // });
      // await stepContext.context.sendActivity(
      //   "Call API hosted in Azure Functions on behalf of user. API endpoint: " + apiConfig.endpoint
      // );
      // await stepContext.context.sendActivity("Response.data: " + JSON.stringify(response.data));

      // console.log(response);
      return await stepContext.endDialog();
    }

    await stepContext.context.sendActivity("Login was not successful please try again.");
    return await stepContext.endDialog();
  }
}
