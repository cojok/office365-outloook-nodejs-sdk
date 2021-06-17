const msal = require('@azure/msal-node');
const axios = require('axios');
const express = require('express');

require('dotenv').config();


const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());

app.listen(port, () => {
  console.log('DOCS running', port);
});

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL Node configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/configuration.md
 */
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.AAD_ENDPOINT + process.env.TENANT_ID,
        clientSecret: '.4_N101myYbbsQh5R4a-8AraOo2XzM__Jx',
    }
};

/**
 * With client credentials flows permissions need to be granted in the portal by a tenant administrator.
 * The scope is always in the format '<resource>/.default'. For more, visit:
 * https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
 */
const tokenRequest = {
    scopes: [process.env.GRAPH_ENDPOINT + '.default'],
};

const apiConfig = {
    uri: process.env.GRAPH_ENDPOINT + 'v1.0/users/',
};

/**
 * Initialize a confidential client application. For more info, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md
 */
const cca = new msal.ConfidentialClientApplication(msalConfig);

/**
 * Acquires token with client credentials.
 * @param {object} tokenRequest
 */
async function getToken(tokenRequest) {
    return await cca.acquireTokenByClientCredential(tokenRequest);
}

/**
 * Calls the endpoint with authorization bearer token.
 * @param {string} endpoint
 * @param {string} accessToken
 */
 async function callApi(endpoint, accessToken) {

  const options = {
      headers: {
          Authorization: `Bearer ${accessToken}`
      }
  };

  console.log('request made to web API at: ' + new Date().toString());

  try {
      const response = await axios.default.get(endpoint, options);
      return response.data;
  } catch (error) {
      console.log(error)
      return error;
  }
};


app.get('/users', async (req, res) => {

  try {
    // here we get an access token
    const authResponse = await getToken(tokenRequest);

    // call the web API with the access token
    const users = await callApi(apiConfig.uri, authResponse.accessToken);

    // display result
    console.log(users);
    res.json(users);
} catch (error) {
    console.log(error);
}

});

app.get('/users/:id', async (req, res) => {

  try {
    // here we get an access token
    const authResponse = await getToken(tokenRequest);

    // call the web API with the access token
    const users = await callApi(`${apiConfig.uri}${req.params.id}`, authResponse.accessToken);

    // display result
    console.log(users);
    res.json(users);
} catch (error) {
    console.log(error);
}

});

//mailFolders/sentItems/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime
app.get('/users/:id/sent-mails', async (req, res) => {

  try {
    // here we get an access token
    const authResponse = await getToken(tokenRequest);

    // call the web API with the access token
    const sentMails = await callApi(`${apiConfig.uri}${req.params.id}/mailFolders/sentItems/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime`, authResponse.accessToken);

    // display result
    console.log(sentMails);
    res.json(sentMails);
} catch (error) {
    console.log(error);
}

});

app.get('/users/:id/draft-mails', async (req, res) => {

  try {
    // here we get an access token
    const authResponse = await getToken(tokenRequest);

    // call the web API with the access token
    // const sentMails = await callApi(`${apiConfig.uri}${req.params.id}/mailFolders/sentItems/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime`, authResponse.accessToken);

    // // display result
    // console.log(sentMails);
    

    const emails = await callApi(`${apiConfig.uri}${req.params.id}/mailFolders/drafts/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime`, authResponse.accessToken);
      const reesp = emails.value.map((email) => ({
          id: email.id,
          sentDateTime: email.sentDateTime,
          receivedDateTime: email.receivedDateTime,
          subject: email.subject,
          bodyPreview: email.bodyPreview,
          importance: email.importance,
          // createdDateTime: email.createdDateTime,
          // lastModifiedDateTime: email.lastModifiedDateTime,
          // categories: email.categories,
          isRead: email.isRead,
          // isDraft: email.isDraft,
          // body: email.body,
          sender: email.sender,
          // toRecipients: email.toRecipients,
          // ccRecipients:email.ccRecipients,
          // bccRecipients:email.bccRecipients,
          // replyTo: email.replyTo,
          // flag: email.flag,
        }));
        res.json(reesp);
} catch (error) {
    console.log(error);
}

});

