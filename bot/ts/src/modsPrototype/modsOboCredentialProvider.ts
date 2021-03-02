import { ConfidentialClientApplication } from '@azure/msal-node';
import { ICredentialProvider, Credential } from './ICredentialProvider';


export class ModsOboCredentialProvider implements ICredentialProvider {
    private msalClient: ConfidentialClientApplication;

    constructor(clientId: string, clientSecret: string, authority: string) {
        this.msalClient = new ConfidentialClientApplication({
            auth: {
                clientId: clientId,
                authority: authority,
                clientSecret: clientSecret
            }
        });
    }

    public async getUserToken(aadObjectId: string, scopes: string[]): Promise<Credential | null> {
        const tokenCache = this.msalClient.getTokenCache();
        const account = await tokenCache.getAccountByLocalId(aadObjectId);
        if (account) {
            const result = await this.msalClient.acquireTokenSilent({
                account,
                scopes
            });
            if (result) {
                return {
                    token: result.accessToken,
                    expiration: result.expiresOn!.toISOString()
                }
            } else {
                return null;
            }
        }
        return null;
    }

    public async exchangeToken(ssoToken: string, scopes: string[]): Promise<Credential | null> {
        try {
            const result = await this.msalClient.acquireTokenOnBehalfOf({
                oboAssertion: ssoToken,
                scopes: scopes
            });
            if (result) {
                return {
                    token: result.accessToken,
                    expiration: result.expiresOn!.toISOString()
                };
            }
            return null;
        } catch (err) {
            return null;
        }
    }
}