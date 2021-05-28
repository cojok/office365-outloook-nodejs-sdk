import MsGraphAuthService from '../services/msGraphAuthService.js';

export default class EmailComponent {
  constructor(authService) {
    this.authService = authService;
  }

  async getEmailAddresses(accessToken) {
    if(accessToken && accessToken.length) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        const emails = await client.api('/me/profile/emails')
        .version('beta')
        .get();
        return emails;
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error)
      }
    } else {
      throw new Error('could not get access token');
    }
  }
}