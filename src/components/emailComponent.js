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
        return emails.value.map((email) => {
          return {
            address: email.address,
            displayName: email.displayName,
            type: email.type,
            id: email.id,
          }
        });
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async sendEmail (accessToken, data) {
    const {
      subject,
      content,
      contentType,
      toRecipients,
      ccRecipients,
      bcRecipients,
      options
    }  = data;

    const message = {
      subject,
      importance: options.importance || 'Low',
      body: {
        content,
        contentType: contentType || 'HTML',
      },
    }

    message.toRecipients = toRecipients.map((recipient) => {
      return { 
        emailAddress: { 
          name: recipient.name, 
          address: recipient.email 
        } 
      };
    });

    // TODO: enable later this options 
    // message.ccRecipients = ccRecipients.map((recipient) => {
    //   return { 
    //     emailAddress: { 
    //       name: recipient.name, 
    //       address: recipient.email 
    //     } 
    //   };
    // });

    // message.bcRecipients = bcRecipients.map((recipient) => {
    //   return { 
    //     emailAddress: { 
    //       name: recipient.name, 
    //       address: recipient.email 
    //     } 
    //   };
    // });

    const email = {
      message,
      saveToSentItems: options.saveToSentItems || 'true',
    }
    
    const client = this.authService.getAuthenticatedClient(accessToken);
    
    try {
      await client.api('/me/sendMail').post(email);
      return true;
    } catch (error) {
      console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
      throw new Error(error);
    }

  }
}