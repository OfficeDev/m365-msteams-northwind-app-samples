## Lab A07: Add a Task Module and Deep Link

In this lab you will begin with the application in folder `B06-MessagingExtension`, make changes as per the steps below to achieve what is in the folder `B07-TaskModule`.
See project structures comparison in Exercise 2.

* [Lab B01: Start an application with bespoke authentication](./Lab-B01.md)
* [Lab B02: Create a teams app](./Lab-B02.md)
* [Lab B03: Make existing teams app use Azure ADO SSO](./Lab-B03.md)
* [Lab B04: Teams styling and themes](./Lab-B04.md)
* [Lab B05: Add a Configurable Tab](./Lab-B05.md)
* [Lab B06: Add a Messaging Extension](./Lab-B06.md)
* [Lab B07: Add a Task Module and Deep Link](./Lab-B07.md)**(You are here)**
* [Lab B08: Add support for selling your app in the Microsoft Teams store](./Lab-B08.md)

In this exercise you will learn new concepts as below:

- [Task modules](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-task-modules)
- [Deep linking](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links)
- [Microsoft Graph to get a user's manager details](https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http)


### Features

- In the order details page, add a button to open a dialog with order details.
- In the dialog add a button to initiate a group chat with it's sales representative and their manager using deep linking.

### Exercise 1: Code changes

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
B06-MessagingExtension
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── 🔺identityClient.js
    │       └── userPanel.js
    ├── images
    │   └── 1.PNG
    │   └── 2.PNG
    │   └── 3.PNG
    │   └── 4.PNG
    │   └── 5.PNG
    │   └── 6.PNG
    │   └── 7.PNG
    │   └── 8.PNG
    │   └── 9.PNG
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
    │   └── 🔺orderDetail.html
    │   └── 🔺orderDetail.js
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
    │   └── makePackage.js
    │   └── manifest.template.json
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── cards
    │       └── errorCard.js
    │       └── productCard.js
    │       └── stockUpdateSuccess.js
    │   └── bot.js
    │   └── constants.js
    │   └── 🔺identityService.js
    │   └── northwindDataService.js
    │   └── server.js
    ├── .env_Sample
    ├── .gitignore
    ├── package.json
    ├── README.md
</pre>
</td>
<td>
<pre>
B07-TaskModule
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── 🔺identityClient.js
    │       └── userPanel.js
    ├── images
    │   └── 1.PNG
    │   └── 2.PNG
    │   └── 3.PNG
    │   └── 4.PNG
    │   └── 5.PNG
    │   └── 6.PNG
    │   └── 7.PNG
    │   └── 8.PNG
    │   └── 9.PNG
    ├── modules
    │   └── env.js
    │   └── northwindDataService.js
    │   └── 🆕orderChatCard.js
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── 🔺orderDetail.html
    │   └── 🔺orderDetail.js
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
    │   └── makePackage.js
    │   └── manifest.template.json
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── cards
    │       └── errorCard.js
    │       └── productCard.js
    │       └── stockUpdateSuccess.js
    │   └── bot.js
    │   └── constants.js
    │   └── 🔺identityService.js
    │   └── northwindDataService.js
    │   └── server.js
    ├── .env_Sample
    ├── .gitignore
    ├── package.json
    ├── README.md
</pre>
</td>
</tr>
</table>


#### Step 1: Add new files

In the project structure, on the right under `A07-TaskModule`, you will see emoji 🆕 near the files.
They are the new files and folders that you need to add into the project structure.


**1.\client\modules\orderChatCard.js**

The adaptive card template used by the task module (dialog).


#### Step 2: Update existing files
In the project structure, on the right under `A07-TaskModule`, you will see emoji 🔺 near the files.
They are the files that were updated to add the new features.
Let's take files one by one to understand what changes you need to make for this exercise. 

**1.\client\identity\identityClient.js**

Add two new functions 
-  **getAADUserFromEmployeeId()** - Get's the AAD user details mapped to the employeeId in the order. 
They are the sales representative for that particular order.
- **getUserDetailsFromAAD()** - Get the user data from Microsoft Graph using the aad userid.
- 
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
import templatePayload from '../modules/orderChatCard.js'
```
- Define two variable, which will be used later. Copy below code and paste below the imports.
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
    const btnTaskModuleElement = document.getElementById('btnTaskModule');
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
            let taskInfo = {
                title: null,
                height: null,
                width: null,
                url: null,
                card: null,
                fallbackUrl: null,
                completionBotId: null,
            };         

            var template = new ACData.Template(templatePayload); 
            // Expand the template with your `$root` data object.
            // This binds it to the data and produces the final Adaptive Card payload
            var cardPayload = template.expand({$root: orderDetails});             
            taskInfo.card = cardPayload;
            taskInfo.title = "chat";
            taskInfo.height = 310;
            taskInfo.width = 430;     
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
        const graphAppUrl = `https://graph.microsoft.com/v1.0/users/${aadUserId}`
        const graphResponse = await fetch(graphAppUrl, {
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
            const graphAppUrl2 = `https://graph.microsoft.com/v1.0/users/${aadUserId}/manager`
            const graphResponse2 = await fetch(graphAppUrl2, {
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
Import statement
```javascript
import {
  initializeIdentityService
} from './identityService.js';
```
becomes
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
### Exercise 2: Test the changes

- Install new packages by running.

```nodejs
npm i
```

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
- Configure the teams tab, select a category for e.g. "Beverages".
- Select a product for e.g. "Chai"
- Go to one of it's orders, this will take you to the order details page.
- In the order details page, select the button `Chat here`. This will open up the task module (dialog)
- You can now view the name of the sales representative for this particular order and their manager's info along with a button `Chat with sales rep team`.
- Select the button `Chat with sales rep team` to start a group chat with this team about the order.


### Known issues

The task module (dialog) has to be closed manually.

### References




