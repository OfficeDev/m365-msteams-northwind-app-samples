## Lab B06: Extend teams application with Messaging Extension

In this lab you will begin with the application in folder `B05-ConfigurableTab`, make changes as per the steps below to achieve what is in the folder `B06-MessagingExtension`.
See project structures comparison in Exercise 2.

* [Lab B01: Start an application with bespoke authentication](./Lab-B01.md)
* [Lab B02: Create a teams app](./Lab-B02.md)
* [Lab B03: Make existing teams app use Azure ADO SSO](./Lab-B03.md)
* [Lab B04: Teams styling and themes](./Lab-B04.md)
* [Lab B05: Add a Configurable Tab](./Lab-B05.md)
* [Lab B06: Add a Messaging Extension](./Lab-B06.md)**(You are here)**
* [Lab B07: Add a Task Module and Deep Link](./Lab-B07.md)
* [Lab B08: Add support for selling your app in the Microsoft Teams store](./Lab-B08.md)

In this exercise you will learn new concepts as below:

- [Messaging extensions](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=nodejs)
- [Bot Framework](https://github.com/microsoft/botframework-sdk)
- [Adaptive cards](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-cards#adaptive-cards)

### Features

- A search based messaging extension to search for products and share result in the form of a rich form card in a conversation
- In the card, provide a button for users to take action to update stock value in the same conversation

### Exercise 1: Bot registration

Messaging extensions allow users to bring the application into a conversation in Teams. You can search data in your application, perform actions on them and send back results of your interaction to your application as well as Teams to display all results in a rich card in the conversation.

Since it is a conversation between your application's web service and teams, you'll need a secure communication protocol to send and receive messages like the Bot Framework's messaging schema.

You'll need to register your web service as a bot in the Bot Framework and update the app manifest to define your web service so Teams knows about it.


#### Step 1: Register your web service as a bot in the Bot Framework in Azure portal

- Go to [https://portal.azure.com/](https://portal.azure.com/)
- In the right pane, select **Create a resource**.
- In the search box enter *bot*, then press Enter.
- Select the **Azure Bot** card.
- Select *Create*.
- Fill the form with all the required fields like *Bot handle*, *Subscription* etc.
- Choose **Multi Tenant** for the **Type of App** field.
- Leave everything else as is and select **Review + create**
- Once validation is passed, select **Create** to create the resources.
- Once deployment is complete, select **Go to resource**, this will take you to the bot resource.
- Once your are in the bot, on the left navigation , select **Configuration**.
- You will see the **Microsoft App ID**, copy the ID (we will need it later as BOT_REG_AAD_APP_ID in .env file)
- Select the link **Manage** next to the Microsoft App ID label. This will take us to Certificates & secrets page of the Azure AD app tied to the bot
- Create a new **Client secret** and copy the `Value` immediately (we will need this later as BOT_REG_AAD_APP_PASSWORD in .env file )
- Go to the registered bot, and on the left navigation select **Channels**
- In the given list of channels, select **Microsoft Teams**, agree to the terms if you wish too and select **Agree** to complete the configurations needed for the bot.

#### Step 2: Run ngrok 

Run below script and copy the tunneled url.

```nodejs
ngrok http 3978 
```

#### Step 3: Update the bot registration configuration

- Copy the url from the above step and go to the bot registered in the Azure portal in Step 1.
- Go to the **Configuration** page from the left navigation
- Immediately on the top of the page you will find a field **Messaging endpoint**
- Paste the url from Step 2 and append `/api/messages` to the url and select **Apply**

### Exercise 2: Code changes

The project structure when you start of this lab and end of this lab is as follows.
Use this depiction for comparison.

<table>
<tr>
<th>Project Structure Before </th>
<th>Project Structure After</th>
</tr>
<tr>
<td valign="top" >
<pre>
B05-ConfigurableTab
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
    │       └── userPanel.js
    ├── modules
    │   └── env.js
    │   └── northwindDataService.js
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── orderDetail.html
    │   └── orderDetail.js
    │   └── privacy.html
    │   └── productDetail.html
    │   └── productDetail.js
    │   └── tabConfig.html
    │   └── tabConfig.js
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── teamstyle.css
    ├── manifest
    │   └── 🔺makePackage.js
    │   └── 🔺manifest.template.jso
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── constants.js
    │   └── 🔺identityService.js
    │   └── 🔺northwindDataService.js</
    │   └── 🔺server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
<td>
<pre>
B06-MessagingExtension
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
    │       └── userPanel.js
    ├── 🆕images
    │   └── 🆕1.PNG
    │   └── 🆕2.PNG
    │   └── 🆕3.PNG
    │   └── 🆕4.PNG
    │   └── 🆕5.PNG
    │   └── 🆕6.PNG
    │   └── 🆕7.PNG
    │   └── 🆕8.PNG
    │   └── 🆕9.PNG
    ├── modules
    │   └── env.js
    │   └── northwindDataService.js
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── orderDetail.html
    │   └── orderDetail.js
    │   └── privacy.html
    │   └── productDetail.html
    │   └── productDetail.js
    │   └── tabConfig.html
    │   └── tabConfig.js
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── teamstyle.css
    ├── manifest
    │   └── 🔺makePackage.js
    │   └── 🔺manifest.template.json
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── 🆕cards
    │       └── 🆕errorCard.js
    │       └── 🆕productCard.js
    │       └── 🆕stockUpdateSuccess.js
    │   └── 🆕bot.js
    │   └── constants.js
    │   └── 🔺identityService.js
    │   └── 🔺northwindDataService.js
    │   └── 🔺server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
</tr>
</table>


#### Step 1: Add new files

In the project structure, on the right under `B06-MessagingExtension`, you will see emoji 🆕 near the files.
They are the new files and folders that you need to add into the project structure.
- `images` folder and it's contents of 9 image files are needed for the rich adaptive cards to display products.
- `cards` folder and the three files `errorCard.js`,`productCard.js` and `stockUpdateSuccess.js` are adaptive cards needed for the messaging extension to display in a conversation based on what state the cards are in.
For e.g. if it's a product card, the bot will use `productCard.js`, if the form is submitted by a user to update the stock value, the bot will use the `stockUpdateSuccess.js` card to let users know the action is completed and incase of any error `errorCard.js` will be displayed.

#### Step 2: Update existing files
In the project structure, on the right under `B06-MessagingExtension`, you will see emoji 🔺 near the files.
They are the files that were updated to add the new features.
Let's take files one by one to understand what changes you need to make for this exercise. 

**1. manifest\makePackage.js**
    When you run script `npm run package` find and replace the key  `BOT_REG_AAD_APP_ID` in the `manifest.template.json` with the value from the `.env` file while generating the new app package for this exercise.

<pre>
if (key.indexOf('TEAMS_APP_ID') === 0 ||
            key.indexOf('HOSTNAME') === 0 ||
            key.indexOf('CLIENT_ID') === 0) {
</pre>
becomes 
<pre>
 if (key.indexOf('TEAMS_APP_ID') === 0 ||
            key.indexOf('HOSTNAME') === 0 ||
            key.indexOf('CLIENT_ID') === 0||
           <b> key.indexOf('BOT_REG_AAD_APP_ID') === 0) {</b>
</pre>

**2.manifest\manifest.template.json**
Update the version number in the `manifest.template.json`.
<pre>
 "version": "1.5.0",
</pre>
becomes
<pre>
 "version": "1.<b>6</b>.0",
</pre>

Add the messaging extension command information
<pre>
"composeExtensions": [
    {
      "botId": "&lt;BOT_REG_AAD_APP_ID&gt;",
      "canUpdateConfiguration": true,
      "commands": [
        {
          "id": "productSearch",
          "type": "query",
          "title": "Find product",
          "description": "",
          "initialRun": false,
          "fetchTask": false,
          "context": [
            "commandBox",
            "compose"
          ],
          "parameters": [
            {
              "name": "productName",
              "title": "product name",
              "description": "Enter the product name",
              "inputType": "text"
            }
          ]
        }
      ]
    }
  ],
  "bots": [
    {
      "botId": "<BOT_REG_AAD_APP_ID>",
      "scopes": [ "personal", "team", "groupchat" ],
      "isNotificationOnly": false,
      "supportsFiles": false
    }
  ],
</pre>
**3.server\identityService.js**

Add an extra condition to let validation  be performed by Bot Framework Adapter if the bot is requested.
In the function `validateApiRequest()`, add another condition in the `if` statement and check if request is for `bot` then move to next step.

<pre>
 if (req.cookies.employeeId && parseInt(req.cookies.employeeId) > 0) {
</pre>
becomes
<pre>
if ((req.cookies.employeeId && parseInt(req.cookies.employeeId) > 0 )
              <b>|| req.path==="/messages")</b> {
</pre>
**4.server\northwindDataService.js**

Add two new functions as below
- <b>getProductByName()</b> - This will search products by name and bring the top 5 results back to the messaging extension's search results.
- <b>updateProductUnitStock()</b> - This will update the value of unit stock based on the input action of a user on the product result card.

**5.server\server.js**

Import the needed modules for bot related code.
Import required bot service from botbuilder package and the bot `StockManagerBot` from newly added file `bot.js`

```javascript
import {StockManagerBot} from './bot.js';
import { BotFrameworkAdapter } from 'botbuilder';
```

A bot adapter authenticates and connects a bot to a service endpoint to send and receive message.
So to authenticate, we'll need to pass the bot registration's AAD app id and app secret.
Add below code to initialize the bot adapter.
```javascript
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_REG_AAD_APP_ID,
  appPassword:process.env.BOT_REG_AAD_APP_PASSWORD
});
```
Create the bot that will handle incoming messages.
```javascript
const stockManagerBot = new StockManagerBot();
```
For the main dialog add error handling.
```javascript
// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${ error }`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
      'OnTurnError Trace',
      `${ error }`,
      'https://www.botframework.com/schemas/error',
      'TurnError'
  );

  // Send a message to the user
  await context.sendActivity('The bot encountered an error or bug.');
  await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};
// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;
```
Listen for incoming requests.
```javascript

const PORT = process.env.PORT || 3978;
app.listen(PORT, () => {
  console.log(`Server is Running on Port ${PORT}`);
});

```

**6.env_Sample**

The .env file you will be creating for your development will have two additional key-value pair for bot registration's Azure AD application credentials as below

```json
BOT_REG_AAD_APP_ID=00000000-0000-0000-0000-000000000000
BOT_REG_AAD_APP_PASSWORD=00000000-0000-0000-0000-000000000000
```

**package.json**

You'll need to install additional packages for adaptive cards and botbuilder.
Add below packages into the `package.json` file.

```json
  "adaptive-expressions": "^4.15.0",
    "adaptivecards": "^2.10.0",
    "adaptivecards-templating": "^2.2.0",   
    "botbuilder": "^4.15.0"
```
### Exercise 3: Test the changes

- Install new packages by running 

```nodejs
npm i
```
- Update .env file with `BOT_REG_AAD_APP_ID` and `BOT_REG_AAD_APP_PASSWORD` which were copied in Step 1.
- Create updated teams app package by running
```nodejs
npm run package
```
- Upload the zipped app package in `manifest` folder in team's app catalog.
- Start server by running
```nodejs
npm start
```
- Go to the app in the app catalog, and add it into a Microsoft Teams team or group chat.
- From the compose area, select the app icon to invoke the messaging extension to search a product.
- Once search results are displayed, select a product and hit send to send the rich card in the conversation.
- Notice the product information, with the form to update the unit stock of the product.
- Update the stock value and select **Update stock** button
- Notice the card being refreshed with new stock value.

### Known issues

The rich adaptive card does not preview in compose area in a Microsoft Teams team's context. This is a bug which is currently with the product team. Fixes will be applied in March '22

### References




