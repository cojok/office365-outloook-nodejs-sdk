// import MsGraphAuthService from './services/msGraphAuthService.js';

// console.log(new MsGraphAuthService());

// import EmailComponent from './components/emailComponent.js';

// new EmailComponent();

// const mainApp = 'test';

// export default mainApp;


import express, { response } from 'express';
import path from 'path';

import MsGraphAuthService from '../src/services/msGraphAuthService.js';
import EmailComponent from '../src/components/emailComponent.js';

const app = express();
const port = process.env.PORT || 3000;

const authService = new MsGraphAuthService();
const emailComponent = new EmailComponent(authService);

app.use(express.static('./docs'));
app.get('/', (req, res) => {
  res.json('hey there');
  // res.sendFile('index.html', {
  //   root: path.join(__dirname, './docs'),
  // });
});

app.get('/auth/signin',
  async function (req, res) {
    try {
      const authUrl = await authService.getAuthURL();
      res.json({ authUrl });
      // res.redirect(authUrl);
    }
    catch (error) {
      console.log(`Error: ${error}`);
      res.json(error);
      // res.redirect('/');
    }
  }
);

app.get('/auth/callback', async (req, res) => {
  try {
    const response = await authService.getAuthDetails(req.query.code);
    // const userId = response.account.homeAccountId;
    // const client = await authService.getAuthenticatedClient(response.accessToken);
    // const user = await client.api('/me').select('displayName,mail,mailboxSettings,userPrincipalName').get();
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    res.json(error);
  }
  res.redirect('/');
});

app.get('/emails', async (req, res) => {
  try {
    const response = await emailComponent.getEmailAddresses(req.query.token);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});
app.listen(port, () => {
  console.log('DOCS running', port);
});

export default app;