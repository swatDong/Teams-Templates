import { ConfirmPrompt, DialogSet, DialogTurnStatus, WaterfallDialog } from 'botbuilder-dialogs';
import { LogoutDialog } from './logoutDialog';
import { ModsSsoPrompt } from '../modsPrototype/modsSsoPrompt';
import { TurnContext } from 'botbuilder';
import { ModsOboCredentialProvider } from '../modsPrototype/modsOboCredentialProvider';

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const MODS_SSO_PROMPT = 'ModsSsoPrompt';

export class MainDialog extends LogoutDialog {
    private requiredScopes: string[] = ["User.Read"]; // hard code the scopes for demo purpose only

    // Developer controlls the lifecycle of credential provider, as well as the cache in it.
    // In this sample the provider is shared in all conversations
    private modsOboCredentialProvider: ModsOboCredentialProvider;

    constructor() {
        super(MAIN_DIALOG, process.env.connectionName);
        console.log('In MainDialog constructor\n');
        console.log(`TeamsAppId: ${process.env.TeamsAppId}\n`);
        console.log(`TeamsAppPassword: ${process.env.TeamsAppPassword}\n`);
        this.modsOboCredentialProvider = new ModsOboCredentialProvider(process.env.TeamsAppId!,
            process.env.TeamsAppPassword!,
            `https://login.microsoftonline.com/${process.env.TeamsAppTenant}`);

        this.addDialog(new ModsSsoPrompt(MODS_SSO_PROMPT,
            {
                credentialProvider: this.modsOboCredentialProvider,
                scopes: this.requiredScopes
            }
        ));

        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            this.displayTokenPhase1.bind(this),
            this.displayTokenPhase2.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} dialogContext
     */
    async run(context: TurnContext, accessor: any) {
        console.log('run\n');
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            console.log('beginDialog\n');
            await dialogContext.beginDialog(this.id);
        }
    }

    async promptStep(stepContext: any) {
        console.log('promptStep\n');
        return await stepContext.beginDialog(MODS_SSO_PROMPT);
    }

    async loginStep(stepContext: any) {
        console.log('loginStep\n');
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        const tokenResponse = stepContext.result;
        if (tokenResponse) {
            await stepContext.context.sendActivity('You are now logged in.');
            return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your token?');
        }
        await stepContext.context.sendActivity('Login was not successful please try again.');
        return await stepContext.endDialog();
    }

    async displayTokenPhase1(stepContext: any) {
        console.log('displayTokenPhase1\n');
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
            return await stepContext.beginDialog(MODS_SSO_PROMPT);
        }
        return await stepContext.endDialog();
    }

    async displayTokenPhase2(stepContext: any) {
        console.log('displayTokenPhase2\n');
        // Option 1 to use token from ModsSsoPrompt: read token set by ModsSsoPrompt from step context directly.
        // Sample code for option 1:
        const tokenResponse = stepContext.result;
        if (tokenResponse) {
            await stepContext.context.sendActivity(`Here is your token ${tokenResponse.token}`);
        }
        return await stepContext.endDialog();

        // Option 2 to use token from ModsSsoPrompt: use credential provider along with MODS SDK, ModsSsoPrompt
        // already cached the token in credential provider

        /* pseudocode for option 2, demo purpose only, more details pending discuss
           const modsSdk = new ModsSdk(stepContext.result.ssoToken, ...otherParameters); // ModsSsoPrompt also provides user SSO token in context
           const graphClient = modsSdk.getGraphClient(
               this.modsOboCredentialProvider); // the credential provider comes from this dialog
        */
    }
}