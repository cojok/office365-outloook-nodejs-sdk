import { config } from 'dotenv';

config();

const configObject = {
  auth: {
    clientId: process.env.OAUTH_APP_ID,
    authority: process.env.OAUTH_AUTHORITY,
    clientSecret: process.env.OAUTH_APP_SECRET
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        console.log(message);
      },
      piiLoggingEnabled: false,
    }
  },
  scope: process.env.OAUTH_SCOPES,
  redirectURI: process.env.OAUTH_REDIRECT_URI,
}

export default configObject;