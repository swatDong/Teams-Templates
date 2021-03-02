import { ExtendedUserTokenProvider, TurnContext, TokenResponse, Activity, MessageFactory, InputHints, CardFactory, OAuthCard, ActionTypes, OAuthLoginTimeoutKey, ActivityTypes, verifyStateOperationName, tokenExchangeOperationName, StatusCodes } from 'botbuilder-core';
import { Dialog, DialogContext, DialogTurnResult, PromptOptions, PromptRecognizerResult } from 'botbuilder-dialogs';
import { ICredentialProvider, Credential } from './ICredentialProvider';
import { v4 as uuidv4 } from 'uuid';

/**
 * Response body returned for a token exchange invoke activity.
 */
class TokenExchangeInvokeResponse {
    id: string;
    connectionName: string;
    failureDetail: string;

    constructor(id: string, connectionName: string, failureDetail: string) {
        this.id = id;
        this.connectionName = connectionName;
        this.failureDetail = failureDetail;
    }
}

/**
 * Settings used to configure an `OAuthPrompt` instance.
 */
export interface SsoPromptSettings {
    scopes: string[];

    credentialProvider: ICredentialProvider;

    /**
     * (Optional) number of milliseconds the prompt will wait for the user to authenticate.
     * Defaults to a value `900,000` (15 minutes.)
     */
    timeout?: number;

    /**
     * (Optional) value indicating whether the OAuthPrompt should end upon
     * receiving an invalid message.  Generally the OAuthPrompt will ignore
     * incoming messages from the user during the auth flow, if they are not related to the
     * auth flow.  This flag enables ending the OAuthPrompt rather than
     * ignoring the user's message.  Typically, this flag will be set to 'true', but is 'false'
     * by default for backwards compatibility.
     */
    endOnInvalidMessage?: boolean;
}

export interface SsoPromptState {
    state: any;
    options: PromptOptions;
    expires: number; // Timestamp of when the prompt will timeout.
}

export class ModsSsoPrompt extends Dialog {
    private readonly PersistedCaller: string = 'botbuilder-dialogs.caller';
    private credentialProvider: ICredentialProvider;

    constructor(
        dialogId: string,
        private settings: SsoPromptSettings
    ) {
        super(dialogId);
        this.credentialProvider = settings.credentialProvider;
    }

    public async beginDialog(dc: DialogContext, options?: PromptOptions): Promise<DialogTurnResult> {
        if (!dc) {
            throw new Error("dialog context is undefined");
        }
        // Initialize prompt state
        const timeout: number = typeof this.settings.timeout === 'number' ? this.settings.timeout : 900000;
        const state: SsoPromptState = dc.activeDialog?.state as SsoPromptState;
        state.state = {};
        state.options = {};
        state.expires = new Date().getTime() + timeout;
        // state[this.PersistedCaller] = null; //Ignore skill bot case in demo

        const output = await this.credentialProvider.getUserToken(dc.context.activity.from.aadObjectId!, this.settings.scopes); // pass empty string as connection name

        if (output) {
            return await dc.endDialog(output);
        } else {
            await this.sendOAuthCardAsync(dc.context);
            return Dialog.EndOfTurn;
        }
    }

    public async continueDialog(dc: DialogContext): Promise<DialogTurnResult> {
        const state: SsoPromptState = dc.activeDialog?.state as SsoPromptState;
        const isMessage: boolean = dc.context.activity.type === ActivityTypes.Message;

        const recognized: PromptRecognizerResult<Credential> = await this.recognizeToken(dc);
        let isValid = false;
        if (recognized.succeeded) {
            isValid = true;
        }

        if (isValid) {
            return await dc.endDialog(recognized.value);
        }

        return Dialog.EndOfTurn;
    }

    private async recognizeToken(dc: DialogContext): Promise<PromptRecognizerResult<Credential>> {
        const context = dc.context;
        let token: Credential | undefined;

        if (this.isTokenExchangeRequestInvoke(context)) {
            const tokenExchangeResponse = await this.credentialProvider.exchangeToken(
                context.activity.value.token,
                this.settings.scopes
            );

            if (!tokenExchangeResponse || !tokenExchangeResponse.token) {
                await context.sendActivity(
                    this.getTokenExchangeInvokeResponse(
                        StatusCodes.PRECONDITION_FAILED,
                        'The bot is unable to exchange token. Ask for user consent.',
                        context.activity.value.id
                    )
                )
            } else {
                await context.sendActivity(
                    this.getTokenExchangeInvokeResponse(StatusCodes.OK, "", context.activity.value.id)
                );
                token = tokenExchangeResponse;
            }
        } else if (this.isTeamsVerificationInvoke(context)) {
            const code: any = context.activity.value.state;
            await this.sendOAuthCardAsync(dc.context);
            await context.sendActivity({ type: 'invokeResponse', value: { status: StatusCodes.OK } });
        }

        return token !== undefined ? { succeeded: true, value: token } : { succeeded: false };
    }

    private getTokenExchangeInvokeResponse(status: number, failureDetail: string, id?: string): Activity {
        const invokeResponse: Partial<Activity> = {
            type: 'invokeResponse',
            value: { status, body: new TokenExchangeInvokeResponse(id as string, "", failureDetail) },
        };
        return invokeResponse as Activity;
    }

    private isTeamsVerificationInvoke(context: TurnContext): boolean {
        const activity: Activity = context.activity;

        return activity.type === ActivityTypes.Invoke && activity.name === verifyStateOperationName;
    }

    private isTokenExchangeRequestInvoke(context: TurnContext): boolean {
        const activity: Activity = context.activity;

        return activity.type === ActivityTypes.Invoke && activity.name === tokenExchangeOperationName;
    }


    private async sendOAuthCardAsync(context: TurnContext): Promise<void> {
        const signInResource = {
            signInLink: `${process.env.BaseUrl}/public/auth-start.html?scope=${encodeURI(this.settings.scopes.join(" "))}&clientId=${process.env.TeamsAppId}&tenantId=${process.env.TeamsAppTenant}`,
            tokenExchangeResource: {
                id: uuidv4()
            }
        };
        const card = CardFactory.oauthCard("", "title", "text", signInResource.signInLink, signInResource.tokenExchangeResource);
        (card.content as OAuthCard).buttons[0].type = ActionTypes.Signin;
        const msg = MessageFactory.attachment(card);

        // Add the login timeout specified in OAuthPromptSettings to TurnState so it can be referenced if polling is needed
        if (!context.turnState.get(OAuthLoginTimeoutKey) && this.settings.timeout) {
            context.turnState.set(OAuthLoginTimeoutKey, this.settings.timeout);
        }

        await context.sendActivity(msg);
    }

}