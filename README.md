# mailByGraph

## 概述
本项目将介绍如何使用 Microsoft Graph API 实现网页邮件发送功能。


## 准备工作
注册 Microsoft Azure Active Directory(AAD)应用程序，获取客户端 ID 和客户端密钥。
了解 Microsoft Graph API 的基本概念和使用方法。
在本地安装 Node.js 和 npm。

## 步骤
1. 创建一个 HTML 文件，并引入 Microsoft Graph API 的 JavaScript SDK:

```html
<!DOCTYPE html>
<html>
<head>
 <meta charset="utf-8">
 <title>Mail Send Example</title>
 <script src="https://cdn.botframework.com/botframework-sdk/latest/botframework-sdk.js"></script>
 <script src="https://cdn.botframework.com/botframework-sdk/latest/microsoftgraph.js"></script>
</head>
<body>
 <h1>Mail Send Example</h1>
 <div id="status"></div>
 <form id="mailForm">
  <label for="to">To:</label>
  <input type="text" id="to" name="to">
  <br><br>
  <label for="subject">Subject:</label>
  <input type="text" id="subject" name="subject">
  <br><br>
  <label for="message">Message:</label>
  <textarea id="message" name="message"></textarea>
  <br><br>
  <button type="submit">Send Mail</button>
 </form>
 <script type="text/javascript">
        var status = document.getElementById('status');
        var mailForm = document.getElementById('mailForm');
        var to = document.getElementById('to');
        var subject = document.getElementById('subject');
        var message = document.getElementById('message');
        var formHandler = function (event) {
            event.preventDefault();
            fetch('/me/sendMail', { method: 'POST', body: JSON.stringify({ to: to.value, subject: subject.value, message: message.value }) }).then(function (response) {
                return response.json();
            }).then(function (data) {
                if (data.error) {
                    console.log(data);
                } else {
                    status.innerHTML = 'Mail sent successfully!';
                }
            }).catch(function (error) {
                console.error(error);
            });
        };
        mailForm.addEventListener('submit', formHandler);
    </script>
</body>
</html>
```
    
2.在服务器上部署网站，并将 JavaScript SDK 文件上传至网站根目录。在 HTML 文件中引用 JavaScript SDK。注意，需要使用 HTTPS 将请求发送到 Microsoft Graph API。如果没有 SSL,请考虑使用自签名证书或购买 SSL 证书。此外，还需要配置应用程序以允许公共访问。具体操作可参考 Microsoft Graph API 文档。

3.在服务器上创建一个 API,用于处理发送邮件的请求。可以使用任何 Web 框架，如 Express、Flask、Django 等。以下是一个使用 Express 实现的示例：
        
```javascript
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
```
    
4.将 API 部署到服务器上，并启动应用程序。在浏览器中访问 http://localhost:${port}/me/sendMail  ,即可发送邮件。注意，需要将 yourdomain.onmicrosoft.com、yourtenantid、YourAppName 替换为实际的值。如果没有安装 Microsoft Graph API SDK,请先从 Microsoft Graph API 文档中下载并安装。
