import express from 'express';
import path from 'path';

import mainApp from './src/app.js';

const app = express();
const port = process.env.PORT || 3000;

app.use(express.static('./docs'));
app.use('/', (req, res) => {
  res.sendFile('index.html', {
    root: path.join(__dirname, './docs'),
  });
});

app.listen(port, () => {
  console.log('DOCS running', port);
});

export default app;

