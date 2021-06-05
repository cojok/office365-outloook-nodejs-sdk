import express from 'express';

import { MsGraphAuthService, EmailComponent } from '../src/app.js';

const app = express();
const port = process.env.PORT || 3000;

const authService = new MsGraphAuthService();
const emailComponent = new EmailComponent(authService);

app.use(express.json());

app.get('/', (req, res) => {
  res.json('hey there');
});

app.get('/auth/signin',
  async (req, res) => {
    try {
      const authUrl = await authService.getAuthURL();
      res.json({ authUrl });
    }
    catch (error) {
      console.log(`Error: ${error}`);
      res.json(error);
    }
  }
);

app.get('/auth/callback', async (req, res) => {
  try {
    const response = await authService.getAuthDetails(req.query.code);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    res.json(error);
  }
  res.redirect('/');
});

app.get('/emails/address', async (req, res) => {
  try {
    const response = await emailComponent.getEmailAddresses(req.query.token);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});

app.post('/emails/send', async (req, res) => {
  try {
    const accessToken = req.body.token;
    const { data } = req.body;
    await emailComponent.sendEmail(accessToken, data);
    return res.sendStatus(202);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});

app.get('/emails', async (req, res) => {
  try {
    const response = await emailComponent.getAllEmails(req.query.token);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});

app.get('/emails/:id', async (req, res) => {
  try {
    const response = await emailComponent.getEmailById(req.query.token, req.params.id);
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