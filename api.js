const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const port = process.env.PORT || 3000; // set port as environment variable or default to 3000
app.use(bodyParser.json()); // support application/json post requests
app.post('/me/sendMail', (req, res) => {
  const to = req.body.to; // extract 'to' from request body using the body-parser library
  const subject = req.body.subject; // extract 'subject' from request body using the body-parser library
  const message = req.body.message; // extract 'message' from request body using the body-parser library
  
  // send email using Microsoft Graph API
  const graphClient = new MicrosoftGraph.Client({
      authProvider: new MicrosoftGraph.AuthProvider('https://login.microsoftonline.com/yourdomain.onmicrosoft.com/yourtenantid/oauth2_client_id'),
      userAgent: 'YourAppName/1.0',
      requestTimeout: 60000
  });
  graphClient.api('/me/sendMail').method('POST')
      .header('Content-Type', 'application/json')
      .body(JSON.stringify({ to, subject, message }))
      .request()
      .then((response) => {
          res.status(200).send(response.body);
      })
      .catch((error) => {
          res.status(500).send(error);
      });
});
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
