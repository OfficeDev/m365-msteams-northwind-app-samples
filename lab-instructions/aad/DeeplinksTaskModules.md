![Teams App Camp](../../assets/code-lab-banner.png)

## Add a Task Module and Deep Link

This lab is an adventure should you choose to go on which begins with a Northwind Orders core application using the `aad` path.
> Complete labs A01-A03 to get to the Northwind Orders core application


Let's look at `Task modules` which are dialogues and `Deep links` which is a smart navigation mechanism within Teams.

**Task modules** are modal pop-up experiences in Teams application to run your app's own html or JavaScript code. 

Using **Deep links** your application can help users navigate easily and intelligently within your application.


In this exercise you will learn new concepts as below:

- [Task modules](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-task-modules?WT.mc_id=m365-58890-cxa)
- [Deep linking](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links?WT.mc_id=m365-58890-cxa)
- [Microsoft Graph to get a user's manager details](https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http&WT.mc_id=m365-58890-cxa)


### Features

- In the application's order details page, add a button to open a dialog with order details.
- In the dialog add a button to initiate a group chat with the order's sales representative and their manager using deep linking.

### Exercise 1: Code changes
---

#### Step 1: Add new files

There are new files and folders that you need to add into the project.

**1.\client\modules\orderChatCard.js**

Create a new file `orderChatCard.js` in the path `\client\modules`, which is the adaptive card template used by the task module (dialog). You can also use any HTML or JavsScript file for creating task modules. Here we will use adaptive cards.

> Adaptive cards are json files but in our project since we own these JSON files and do not use any modern bundlers, we have created JS files out of them for the ease of importing content.

Copy below content into the new file created.
```javascript
export default
{
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "refresh": {
        "userIds": [],
        "action": {
            "type": "Action.Execute",
            "verb": "refresh",
            "title": "Refresh",
            "data": {              
            }
        }
    },
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Chat with ${contact}: Order #${orderId}",
                            "horizontalAlignment": "left",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                              {
                                "title": "Sales rep:",
                                "value": "${salesRepName}"
                              },
                              {
                                "title": "Sales rep manager:",
                                "value": "${salesRepManagerName}"
                              } 
                                     ]
                            }
                        ]
                }
            ]
        }
      
    ],
    "actions": [
       
        {
            "type": "Action.OpenUrl",
            "title": "Chat with sales rep team",
            "id": "chatWithUser",
            "url": "https://teams.microsoft.com/l/chat/0/0?users=${salesRepEmail},${salesRepManagerEmail}&message=Questions%20on%20Order%20${orderId}%20&topicName=Enquire%20about%20Order%20${orderId}%20"
        }
    ]
}
```

#### Step 2: Update existing files

There are files that were updated to add the new features.
Let's take files one by one to understand what changes you need to make for this exercise. 

**1.\client\identity\identityClient.js**

Add two new functions 
-  **getAADUserFromEmployeeId()** - Get's the AAD user details mapped to the employeeId in the order. 
They are the sales representative for that particular order.
- **getUserDetailsFromAAD()** - Get the user data from Microsoft Graph using the aad userid.

Append below code in the `identityClient.js` file.
```javascript


export async function getAADUserFromEmployeeId(employeeId) {

    const response = await fetch (`/api/getAADUserFromEmployeeId?employeeId=${employeeId}`, {
        "method": "get",
        "headers": await getFetchHeadersAuth(),       
        "cache": "no-cache"
    });
    if (response.ok) {
        return response.text().then((data) => data ? JSON.parse(data).id : null);      
    } else {
        const error = "getAADUserFromEmployeeId failed"
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}
export async function getUserDetailsFromAAD(aadUserId) {

    const response = await fetch (`/api/getUserDetailsFromAAD?aadUserId=${aadUserId}`, {
        "method": "get",
        "headers": await getFetchHeadersAuth(),       
        "cache": "no-cache"
    });
    if (response.ok) {
        return response.text().then((data) => data ? JSON.parse(data) : null);          
    } else {
        const error = "getUserDetailsFromAAD failed"
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}
```
**2.\client\pages\orderDetail.html**

- Add scripts required for adaptive card templating, in the same order.
Copy below code and paste it above the `<top-nav-panel>` tag.
```html
<script src='https://unpkg.com/markdown-it/dist/markdown-it.min.js'></script>
<script src='https://unpkg.com/adaptive-expressions/lib/browser.js'></script>
<script src='https://unpkg.com/adaptivecards-templating/dist/adaptivecards-templating.min.js'></script>
<script src='https://unpkg.com/adaptivecards/dist/adaptivecards.min.js'></script>
 
```
- Add a new button to open the task module (dialog). Add below code above the `<table>` tag.
```html
 <button id="btnTaskModule">Chat here</button>
```
**3.\client\pages\orderDetail.js**

- Import the needed modules. Paste the below code before the function definition of **displayUI()**.
```javascript
import {
    getAADUserFromEmployeeId,getUserDetailsFromAAD
} from '../identity/identityClient.js'
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
import templatePayload from '../modules/orderChatCard.js';
```
- Define two variable above `displayUI()` function definition, which will be used later. Copy below code and paste below the imports.
```javascript
let orderId="0";
let orderDetails={};
```
- Update the **displayUI()** function by first creating a constant to get the button for the task module that we added earlier in **orderDetail.html**. Paste below code after constant definition for **detailsElement**.
```javascript
const btnTaskModuleElement = document.getElementById('btnTaskModule');
```
- Replace the **displayUI()** function definition with below code

<pre>
async function displayUI() {
    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    <b>const btnTaskModuleElement = document.getElementById('btnTaskModule');</b>
    try {
        <b>microsoftTeams.initialize(async () => {</b>
        const searchParams = new URLSearchParams(window.location.search);
        if (searchParams.has('orderId')) {
              orderId = searchParams.get('orderId');

            const order = await getOrder(orderId);    
            <b>//graph call to get AAD mapped employee details        
            let user=await getAADUserFromEmployeeId(order.employeeId);  
             if(!user)   
                user= await getAADUserFromEmployeeId("1");     //fall back to employee 1 
            const salesRepdetails=await getUserDetailsFromAAD(user);          

            orderDetails.orderId=orderId?orderId:"";
            orderDetails.contact=order.contactName&&order.contactTitle?`${order.contactName}(${order.contactTitle})`:"";
            orderDetails.salesRepEmail=salesRepdetails.mail?salesRepdetails.mail:"";
            orderDetails.salesRepManagerEmail=salesRepdetails.managerMail?salesRepdetails.managerMail:"";
            orderDetails.salesRepName=salesRepdetails.displayName?salesRepdetails.displayName:"";
            orderDetails.salesRepManagerName=salesRepdetails.managerDisplayName?salesRepdetails.managerDisplayName:"";</b>

            displayElement.innerHTML = `
                    &lt;h1&gt;Order ${order.orderId}&lt;/h1&gt;
                    &lt;p&gt;Customer: ${order.customerName} &lt;br /&gt;
                    Contact: ${order.contactName}, ${order.contactTitle}&lt;br /&gt;
                    Date: ${new Date(order.orderDate).toDateString()}&lt;br /&gt;
                    ${order.employeeTitle}: ${order.employeeName} (${order.employeeId})
                    &lt;/p&gt;
                `;

            order.details.forEach(item => {
                const orderRow = document.createElement('tr');
                orderRow.innerHTML = `&lt;tr&gt;
                        &lt;td&gt;${item.quantity}&lt;/td&gt;
                        &lt;td&gt;&lt;a href="/pages/productDetail.html?productId=${item.productId}">${item.productName}&lt;/a&gt;&lt;/td&gt;
                        &lt;td&gt;${item.unitPrice}&lt;/td&gt;
                        &lt;td&gt;${item.discount}&lt;/td&gt;
                    &lt;/tr&gt;`;
                detailsElement.append(orderRow);

            });

        }
    <b>btnTaskModuleElement.addEventListener('click',  ev => {
            let submitHandler = (err, result) => { console.log(result); };          
            var template = new ACData.Template(templatePayload); 
            // Expand the template with your `$root` data object.
            // This binds it to the data and produces the final Adaptive Card payload
            var cardPayload = template.expand({$root: orderDetails});             
             const taskInfo = {
                card:cardPayload,
                title:"chat",
                height:310,
                width:430,
                url: null,               
                fallbackUrl: null,
                completionBotId: null,
            }; 
            //invoke the task module (dialog)    
            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        });
        
    });</b>
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}

</pre>

**4.\server\identityService.js**
 Add below two functions:
- **getAADUserFromEmployeeId()** - To get the Azure AD user details based on employee id.
- **getUserDetailsFromAAD()** - To get user information including their manager details.
    
> We are using Microsoft Graph API to get user details from AAD.

Copy the below code and append to the end of the file.


```javascript
export async function getAADUserFromEmployeeId(employeeId) {
    let aadUserdata;

    try {
        const msalResponse = await msalClientApp.acquireTokenByClientCredential(msalRequest);
        const graphResponse = await fetch(
            `https://graph.microsoft.com/v1.0/users/?$filter=employeeId eq '${employeeId}'`,
            {
                "method": "GET",
                "headers": {
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${msalResponse.accessToken}`
                }
            });
        if (graphResponse.ok) {
            const userProfile = await graphResponse.json();
            aadUserdata = userProfile.value[0];

        } else {
            console.log(`Error ${graphResponse.status} calling Graph in getAADUserFromEmployeeId: ${graphResponse.statusText}`);
        }
    }
    catch (error) {
        console.log(`Error calling MSAL in getAADUserFromEmployeeId: ${error}`);
    }
    return aadUserdata;

}
const userCache = {}
export async function getUserDetailsFromAAD(aadUserId) {
    let graphResult = {};
    try {
        if (userCache[aadUserId]) return userCache[aadUserId];
        const msalResponse = await msalClientApp.acquireTokenByClientCredential(msalRequest);
        const graphUserUrl = `https://graph.microsoft.com/v1.0/users/${aadUserId}`
        const graphResponse = await fetch(graphUserUrl, {
            method: "GET",
            headers: {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": `Bearer ${msalResponse.accessToken}`
            }
        });
        if (graphResponse.ok) {
            const graphData = await graphResponse.json();
            graphResult.mail = graphData.mail;
            graphResult.displayName = graphData.displayName;
            const graphManagerUrl = `https://graph.microsoft.com/v1.0/users/${aadUserId}/manager`
            const graphResponse2 = await fetch(graphManagerUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${msalResponse.accessToken}`
                }
            });
            if (graphResponse2.ok) {
                const managerInfo = await graphResponse2.json();
                graphResult.managerMail = managerInfo.mail;
                graphResult.managerDisplayName = managerInfo.displayName;
            } else {
                console.log(`Error ${graphResponse2.status} calling Graph in getUserDetailsFromAAD: ${graphResponse2.statusText}`);
            }
            userCache[aadUserId] = graphResult;
        } else {
            console.log(`Error ${graphResponse.status} calling Graph in getUserDetailsFromAAD: ${graphResponse.statusText}`);
        }
    }
    catch (error) {
        console.log(`Error calling MSAL in getUserDetailsFromAAD: ${error}`);
    }
    return graphResult;
}
```
**5.\server\server.js**

- Import the two functions `getAADUserFromEmployeeId()` and `getUserDetailsFromAAD()` from the module `identityService.js`.
Replace below code block"
Import statement
```javascript
import {
  initializeIdentityService
} from './identityService.js';
```
With:
```javascript
import {
  initializeIdentityService,
  getAADUserFromEmployeeId,getUserDetailsFromAAD
} from './identityService.js';
```

- Append below code for routing the api calls for `getAADUserFromEmployeeId` and `getUserDetailsFromAAD`
```javascript
app.get('/api/getAADUserFromEmployeeId', async (req, res) => {

  try {
    const employeeId = req.query.employeeId;
    const employeeData = await getAADUserFromEmployeeId(employeeId);
    res.send(employeeData);
  }
  catch (error) {
      console.log(`Error in /api/getAADUserFromEmployeeId handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});
app.get('/api/getUserDetailsFromAAD', async (req, res) => {

  try {
    const aadUserId = req.query.aadUserId;
    const userData = await getUserDetailsFromAAD(aadUserId);
    res.send(userData);
  }
  catch (error) {
      console.log(`Error in /api/getUserDetailsFromAAD handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});
```
**6. manifest\manifest.template.json**

Update version number from `1.6.0` to `1.7.0`.
~~~json
"version": "1.7.0"
~~~
> NOTE: Make each new version greater than the last app so you can update the application in Teams!

### Exercise 3: Test the changes
---
Now that you have applied all code changes, let's test the features.


#### Step 1: Create new teams app package

Create updated teams app package by running below script:
```nodejs
npm run package
```

#### Step 2: Upload the app package
In the Teams web or desktop UI, click "Apps" in the sidebar 1️⃣, then "Manage your apps" 2️⃣. At this point you have three choices:

* Upload a custom app (upload the app for yourself or a specific team or group chat) - this only appears if you have enabled "Upload custom apps" in your setup policy; this was a step in the previous lab
* Upload an app to your org's app catalog (upload the app for use within your organization) - this only appears if you are a tenant administrator
* Submit an app to your org (initiate a workflow asking a tenant administrator to install your app) - this appears for everyone

In this case, choose the first option 3️⃣.

<img src="https://github.com/OfficeDev/m365-msteams-northwind-app-samples/assets/03-005-InstallApp-1.png?raw=true" alt="Upload the app"/>

Navigate to the Northwind.zip file in your manifest directory and upload it. 
The Teams client will display the application information, add the application to a team or a group chat.
<img src="https://github.com/OfficeDev/m365-msteams-northwind-app-samples/assets/07-001-addapp.png?raw=true" alt="Add the app"/>


#### Step 3: Start your local project

Now it's time to run your updated application and run it in Microsoft Teams. Start the application by running below command: 

```nodejs
npm start
```

#### Step 4 : Run the application in Teams client

Once you are in the application, go to `My orders` page and select any order as shown below:

<img src="https://github.com/OfficeDev/m365-msteams-northwind-app-samples/assets/07-004-selectorder.png?raw=true" alt="Order details page"/>
This will take you to the `Order details page`.

Notice the button `Chat here`1️⃣. This button opens a dialogue 2️⃣ to show the Sales representative and their team.
<img src="https://github.com/OfficeDev/m365-msteams-northwind-app-samples/assets/07-005-chat.png?raw=true" alt="chat"/>

Select `Chat with the sales rep team` button, this will initiate a group chat with the user, and the sales rep team.
    > Manually close the dialog once you are in the group chat.

1. You'll see that the chat is initiated to a group (Sales rep and the sales rep's manager)
    > Sales rep's manager information is taken from Microsoft365 data using Microsoft Graph's [Get manager](https://docs.microsoft.com/en-us/graph/api/orgcontact-get-manager?view=graph-rest-1.0&tabs=http) api.
1. The chat's topic has the order number from where the chat is initiated.
1. The chat's initial message is already typed and ready with the order number.
<img src="https://github.com/OfficeDev/m365-msteams-northwind-app-samples/assets/07-006-groupchat.png?raw=true" alt="Group chat"/>

### Known issues

The task module (dialog) has to be closed manually.
