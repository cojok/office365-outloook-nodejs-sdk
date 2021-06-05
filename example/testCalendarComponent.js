import express from 'express';

import { MsGraphAuthService, CalendarComponent } from '../src/app.js';

const app = express();
const port = process.env.PORT || 3000;

const authService = new MsGraphAuthService();
const calendarComponent = new CalendarComponent(authService);

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

app.get('/calendarView', async (req, res) => {
  try {
    const response = await calendarComponent.getAllEvents(req.query.token, req.query.timeZone);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});

app.get('/event-id', async (req, res) => {
  try {
    const response = await calendarComponent.getEventById(req.query.token, req.query.timeZone, req.query.id);
    return res.json(response);
  } catch (error) {
    console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
    return res.json(error);
  }
});

app.post('/new-event', async (req, res) => {
  try {
    const response = await calendarComponent.createNewEvent(req.query.token, req.query.timeZone, req.body.data);
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