## Lab A06: Extend teams application with Messaging Extension

In this lab you will begin with the application in folder `A05-ConfigurableTab`, make changes as per the steps below to achieve what is in the folder `A06-MessagingExtension`.
See project structures comparison in Exercise 2.


You will learn new concepts:

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
- You will see the **Microsoft App ID**, copy the ID (we will need it later)
- Select the link **Manage** next to the Microsoft App ID label. This will take us to Certificates & secrets page of the Azure AD app tied to the bot
- Create a new **Client secret** and copy the `Value` immediately (we will need this later)
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
<th>A05-ConfigurableTab - Before </th>
<th>A06-MessagingExtension - After</th>
</tr>
<tr>
<td style=" vertical-align: top;">
<pre>
A05-ConfigurableTab
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
    │   └── <span style="color:red">makePackage.js</span>
    │   └── <span style="color:red">manifest.template.json</span>
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── constants.js
    │   └── <span style="color:red">identityService.js</span>
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── <span style="color:red">server.js</span>
    ├── <span style="color:red">.env_Sample</span>
    ├── .gitignore
    ├── <span style="color:red">package.json</span>
    ├── README.md
</pre>
</td>
<td>
<pre>
A06-MessagingExtension
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
    │       └── userPanel.js
    ├── <span style="color:purple">images</span>
    │   └── <span style="color:purple">1.PNG</span>
    │   └── <span style="color:purple">2.PNG</span>
    │   └── <span style="color:purple">3.PNG</span>
    │   └── <span style="color:purple">4.PNG</span>
    │   └── <span style="color:purple">5.PNG</span>
    │   └── <span style="color:purple">6.PNG</span>
    │   └── <span style="color:purple">7.PNG</span>
    │   └── <span style="color:purple">8.PNG</span>
    │   └── <span style="color:purple">9.PNG</span>
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
    │   └── <span style="color:red">makePackage.js</span>
    │   └── <span style="color:red">manifest.template.json</span>
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── <span style="color:purple">cards</span>
    │       └── <span style="color:purple">errorCard.js</span>
    │       └── <span style="color:purple">productCard.js</span>
    │       └── <span style="color:purple">stockUpdateSuccess.js</span>
    │   └── <span style="color:purple">bot.js</span>
    │   └── constants.js
    │   └── <span style="color:red">identityService.js</span>
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── <span style="color:red">server.js</span>
    ├── <span style="color:red">.env_Sample</span>
    ├── .gitignore
    ├── <span style="color:red">package.json</span>
    ├── README.md
</pre>
</td>
</tr>
</table>


#### Step 1: Add new files

In the project structure, on the right under `A06-MessagingExtension`, you will see <span style="color:purple"> purple </span> colored files.
They are the new files and folders that you need to add into the project structure.
- `images` folder and it's contents of 9 image files are needed for the rich adaptive cards to display products.
- `cards` folder and the three files `errorCard.js`,`productCard.js` and `stockUpdateSuccess.js` are adaptive cards needed for the messaging extension to display in a conversation based on what state the cards are in.
For e.g. if it's a product card, the bot will use `productCard.js`, if the form is submitted by a user to update the stock value, the bot will use the `stockUpdateSuccess.js` card to let users know the action is completed and incase of any error `errorCard.js` will be displayed.

#### Step 2: Update existing files
In the project structure, on the right under `A06-MessagingExtension`, you will see <span style="color:red">red</span> colored files.
They are the files that were updated to add the new features.
Let's take files one by one to understand what changes you need to make for this exercise.
##### 

### Exercise 3: Test the changes


### Known issues

### References



