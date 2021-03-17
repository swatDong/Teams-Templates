/* This code sample provides a starter kit to implement server side logic for your Teams App in JavaScript with code
 * snippets, refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure
 * Functions developer guide.
 */

const ModsServerSdk = require('mods-server');

{{#apim}}
/**
 * swagger
 */
{{/apim}}
module.exports = async function (context, req, config) {
  context.log('JavaScript HTTP trigger function processed a request.');

  // Initialize response.
  const res = {
    status: 200,
    body: {}
  };

  // Put an echo into response body.
  res.body['receivedHTTPRequestBody'] = req.body || '';

  // Initialize MODS Server SDK with MODS configuration.
  let client;
  try {
    client = ModsServerSdk.MODS.getInstance(config);
  } catch(e) {
    context.log.error(e);
    res.status = 500;
    res.body['error'] = 'Fail to initialize MODS Server SDK.';
    return res;
  }

  // Query user's information from the authentication token.
  const currentUser = client.getUserInfo();
  if (currentUser && currentUser.displayName) {
    res.body['userInfoMessage'] = `User display name is ${currentUser.displayName}.`;
  } else {
    res.status = 400;
    res.body['userInfoMessage'] = 'Fail to get user display Name.';
    return res;
  }

  // Create a graph client to access user's Microsoft 365 data after user has consented.
  try {
    const graphClient = await client.getMicrosoftGraphClientWithUserIdentity(['.default']);
    const profile = await graphClient.api('/me').get();
    res.body['graphClientMessage'] = profile;
  } catch (e) {
    context.log.error(e);
    res.status = 500;
    res.body['graphClientMessage'] = 'Fail to get profile, maybe consent flow is required.';
    return res;
  }

  return res;
}