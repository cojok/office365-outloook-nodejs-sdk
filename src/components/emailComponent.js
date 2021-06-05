export default class EmailComponent {
  constructor(authService) {
    this.authService = authService;
  }

  async getEmailAddresses(accessToken) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        const emails = await client.api('/me/profile/emails')
          .version('beta')
          .get();
        return emails.value.map((email) => ({
          address: email.address,
          displayName: email.displayName,
          type: email.type,
          id: email.id,
        }));
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async getAllEmailsInbox(accessToken) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        // TODO try to see what makes sense here as standard params for the select query param
        // TODO maybe dynamic values with default config should be the way or default + extras
        const emails = await client.api('/me/mailFolders/inbox/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime').version('beta').get();
        return emails.value.map((email) => ({
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
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async getAllEmailsSentItems(accessToken) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        // TODO try to see what makes sense here as standard params for the select query param
        // TODO maybe dynamic values with default config should be the way or default + extras
        const emails = await client.api('/me/mailFolders/sentItems/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime').version('beta').get();
        return emails.value.map((email) => ({
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
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async getAllEmailsDrafts(accessToken) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        // TODO try to see what makes sense here as standard params for the select query param
        // TODO maybe dynamic values with default config should be the way or default + extras
        const emails = await client.api('/me/mailFolders/drafts/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime').version('beta').get();
        return emails.value.map((email) => ({
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
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async getAllEmailsDeleteItems(accessToken) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        // TODO: try to see what makes sense here as standard params for the select query param
        // TODO: maybe dynamic values with default config should be the way or default + extras
        const emails = await client.api('/me/mailFolders/deletedItems/messages?$select=bodyPreview,subject,sender,sentDateTime,importance,isRead,flag,receivedDateTime').version('beta').get();
        return emails.value.map((email) => ({
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
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async getEmailById(accessToken, id) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        // TODO: try to see what makes sense here as standard params for the select query param
        // TODO: maybe dynamic values with default config should be the way or default + extras
        const email = await client.api(`/me/messages/${id}?$select=body,categories,ccRecipients,bccRecipients,createdDateTime,flag,sender,sentDateTime,attachments,isDraft,isRead,receivedDateTime,replyTo`).version('beta').get();
        delete email['@odata.context'];
        delete email['@odata.etag'];
        return email;
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  async sendEmail(accessToken, data) {
    const {
      subject,
      content,
      contentType,
      toRecipients,
      ccRecipients,
      bcRecipients,
      options,
    } = data;

    const message = {
      subject,
      importance: options.importance || 'Low',
      body: {
        content,
        contentType: contentType || 'HTML',
      },
    };

    message.toRecipients = toRecipients.map((recipient) => ({
      emailAddress: {
        name: recipient.name,
        address: recipient.email,
      },
    }));

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
    };

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
