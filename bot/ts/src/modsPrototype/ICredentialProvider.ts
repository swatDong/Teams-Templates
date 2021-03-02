export interface Credential {
    token: string;
    expiration: string;
}

export interface ICredentialProvider {
    getUserToken(aadObjectId: string, scopes: string[]): Promise<Credential | null>;
    exchangeToken(ssoToken: string, scopes: string[]): Promise<Credential | null>;
}