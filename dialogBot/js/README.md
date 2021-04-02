## Prerequisite

1. Install [ngrok](https://ngrok.com/).

## Steps

1. Create Teams AAD app on [Azure portal App Registration](https://ms.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
    1. Record client id. (referred to as `TEAMS-AAD-APP-CLIENT-ID`)
    1. Generate client secret and record it. (referred to as `TEAMS-AAD-APP-CLIENT-SECRET`)

1. Run ngrok with the following command to create a new tunnel. [REF](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-debug-channel-ngrok?view=azure-bot-service-4.0#run-ngrok)
    ```
    ngrok http -host-header=rewrite 3978
    ```
    Record the ngrok host: `xxx.ngrok.io` (referred to as `YOUR-NGROK-HOST`)

1. Create Bot AAD app.
    1. Record client id and tenant id. (referred to as `M365-AAD-APP-CLIENT-ID` and `M365-AAD-APP-TENANT-ID`)
    1. Generate client secret and record it. (referred to as `M365-AAD-APP-CLIENT-SECRET`)
    1. Set Authentication settings:
        - Redirect URIs: `https://<your-ngrok-host>/public/auth-end.html`
        - scopes: `api://botid-<TEAMS-AAD-APP-CLIENT-ID>/access_as_user`
        - authorized client applications:
			- 5e3ce6c0-2b1f-4285-8d4b-75ee78787346
			- 1fec8e78-bce4-4aaf-ab1b-5451cc387264

1. Install `mods-sdk` package from local path:
    1. In the root folder of mods-sdk, run the following command to pack the mods-sdk to a `microsoft-mods-sdk-x.x.x.tgz` package.
        ```
        npm pack
        ```
    1. In the root folder of bot test project, install mods-sdk by running:
        ```
        npm i <local-path-to-microsoft-mods-sdk-x.x.x.tgz>
        ```

1. In the root folder of bot test project, install other dependencies.
    ```
    npm i
    ```

1. Fill the environment variable:
    ```
    MicrosoftAppId=<TEAMS-AAD-APP-CLIENT-ID>
    MicrosoftAppPassword=<TEAMS-AAD-APP-CLIENT-SECRET>
    M365_CLIENT_ID=<M365-AAD-APP-CLIENT-ID>
    M365_CLIENT_SECRET=<M365-AAD-APP-CLIENT-SECRET>
    M365_TENANT_ID=<M365-AAD-APP-TENANT-ID>
    M365_AUTHORITY_HOST=https://login.microsoftonline.com
    INITIATE_LOGIN_ENDPOINT=<YOUR-NGROK-HOST>/public/auth-start.html
    ```
1. build the project.
    ```
    npm run build
    ```

1. Update [manifest.json](./teamsAppManifest/manifest.json) file:

    1. Replace all `<BOT_AAD_APP_ID>` to your `M365-AAD-APP-CLIENT-ID`.
    1. Replace all `<TEAMS_AAD_APP_ID>` to your `TEAMS-AAD-APP-CLIENT-ID`.
    1. Replace all `<YOUR_NGROK_HOST>` to your `YOUR-NGROK-HOST`.

1. Install bot app on teams.

    Zip files under *./teamsAppManifest* to a zip package and upload to teams through `upload a custom app`. Install the bot app on Teams.

1. Open bot project in VSCode window and press `F5` to start and debug the bot app. Enter `Hello` in the installed teams bot app conversation chat window.
