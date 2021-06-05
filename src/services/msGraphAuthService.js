import msal from '@azure/msal-node';
import graph from '@microsoft/microsoft-graph-client';
import config from '../config/index.js';
import 'isomorphic-fetch';

export default class MsGraphAuthService {

  constructor(options = config) {
    const { auth, system, scope, redirectURI } = options;
    this.msalConfig = {
      auth,
      system,
    };
    this.msalConfig.system.loggerOptions.logLevel = msal.LogLevel.Verbose;
    this.msalConfigClient = new msal.ConfidentialClientApplication(this.msalConfig);
    this.scope = scope;
    this.redirectURI = redirectURI;
  }

  getAuthURL () {
    const urlParams = {
      scopes: this.scope.split(','),
      redirectUri: this.redirectURI,
    };
    try {
      return this.msalConfigClient.getAuthCodeUrl(urlParams);
    } catch (error) {
      console.log(`getAuthURL error: ${error}`);
      throw new Error(error);
    }
  }

  async getAuthDetails (code) {
    const tokenRequest = {
      code: code,
      scopes: this.scope.split(','),
      redirectUri: this.redirectURI,
    };
    try {
      const response = await this.msalConfigClient.acquireTokenByCode(tokenRequest);
      const userId = response.account.homeAccountId;

      const client = await this.getAuthenticatedClient(response.accessToken);
      const userDetails = await client.api('/me').select('displayName,mail,mailboxSettings,userPrincipalName').get();
      // const userDetails = await graph.getUserDetails(response.accessToken);
  
        // Add the user details
      const user = {
        displayName: userDetails.displayName,
        email: userDetails.mail || userDetails.userPrincipalName,
        timeZone: userDetails.mailboxSettings.timeZone,
        accessToken: response.accessToken,
      };
      return user;
    } catch (error) {
      console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
      throw new Error(error);
    }
  }

  async getUserDetails (accessToken) {
    const client = this.getAuthenticatedClient(accessToken);
    const user = client.api('/me').select('displayName,mail,mailboxSettings,userPrincipalName').get();
    return user;
  }
  async removeToken(userId) {
    if (userId) {
      // Look up the user's account in the cache
      const accounts = await this.msalClient
        .getTokenCache()
        .getAllAccounts();

      const userAccount = accounts.find(a => a.homeAccountId === userId);

      // Remove the account
      if (userAccount) {
        this.msalClient
          .getTokenCache()
          .removeAccount(userAccount);
      }
    }
    return true;
  }

  getAuthenticatedClient (accessToken){
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
    return client;
  };

  async getAccessToken(userId) {
    // Look up the user's account in the cache
    try {
      const accounts = await this.msalClient
        .getTokenCache()
        .getAllAccounts();
  
      const userAccount = accounts.find(a => a.homeAccountId === userId);
  
      // Get the token silently
      const response = await this.msalClient.acquireTokenSilent({
        scopes: this.scope.split(','),
        redirectUri: this.redirectURI,
        account: userAccount
      });
  
      return response.accessToken;
    } catch (error) {
      console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
      throw new Error(error);
    }
  }
}
