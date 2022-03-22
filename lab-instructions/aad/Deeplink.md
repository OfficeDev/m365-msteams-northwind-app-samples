![Teams App Camp](../../assets/code-lab-banner.png)

## Add a Deep link to a personal Tab

This lab is part of extending app capabilities for your teams app which begins with a Northwind Orders core application using the `aad` path. The [core app](../../src/create-core-app/aad/A03-after-apply-styling/) is the boilerplate application with which you will do this lab.

> Complete labs [A01](A01-begin-app.md)-[A03](A03-after-apply-styling.md) for deeper understanding of how the core application works, to set up AAD application registration etc. to update the `.env` file as per the `.env_sample`. This configuration is required for the success of the lab.

Deep links help the user to directly navigate to the content.
In this lab we will create deep link to entities in Teams so the user can navigate to contents within the app's tab.

In this exercise you will learn new concepts as below:

- [Deep links](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links?WT.mc_id=m365-58890-cxa)


### How to build the deep link

Use below syntax, to create the deep link for this lab.

```
https://teams.microsoft.com/l/entity/<app-id>/<entitiyId>?context={"subEntityId": "<subEntityId>"}
```

- **app-id** - This teams app id from the manifest file
- **entityId** - This is defined in your manifest file in the `staticTabs` object for the particular entity (tab). In our case this is the entity id `Orders` of  `My Orders` tab.
- **subEntityId** - This is the ID for the item you are displaying information for. This is similar to query parameters. In our case in this lab, it will be the orderId.


### Features

- In the application's order details page, add a button to copy the order's tab link into clipboard, that helps users share the link via chat or outlook to colleague, for them to navigate easily to that specific order.


### Exercise 1: Code changes
---

#### Step 1: Update existing files


**1. client\page\orderDetail.html**

Let's add the copy to clipboard button and a div to display a message to show if the copy was successful.

Add below block of code and paste it above `orderDetails` div element.

```javascript
<div id="copySection" style="display: none;">
<div> <button id="btnCopyOrderUrl">Copy order url</button></div>
    <div style="flex-grow: 1;padding:5px;" id="copyMessage">Copy to clipboard</div>
</div> 
```

**2. client\page\orderDetail.js**
 
Import the teams SDK module as well as the environment variable we exported to use.

Paste below code above the displayUI() function definition.

```javascript
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
import { env } from '/modules/env.js';

```

Replace the displayUI() function with below definition:

```javascript
async function displayUI() {
    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    const copyUrlElement = document.getElementById('btnCopyOrderUrl');
    const copyMsgElement=document.getElementById('copyMessage');
    const copySectionElement=document.getElementById('copySection');
    const errorMsgElement=document.getElementById('message');
    try {

        const searchParams = new URLSearchParams(window.location.search);        
        microsoftTeams.initialize(async () => {
        microsoftTeams.getContext(async (context)=> {      
     
        if (searchParams.has('orderId')||context.subEntityId) {
            const orderId = searchParams.get('orderId')?searchParams.get('orderId'):context.subEntityId;
            const order = await getOrder(orderId);
            displayElement.innerHTML = `
                    <h1>Order ${order.orderId}</h1>
                    <p>Customer: ${order.customerName}<br />
                    Contact: ${order.contactName}, ${order.contactTitle}<br />
                    Date: ${new Date(order.orderDate).toDateString()}<br />
                    ${order.employeeTitle}: ${order.employeeName} (${order.employeeId})
                    </p>
                `;

            order.details.forEach(item => {
                const orderRow = document.createElement('tr');
                orderRow.innerHTML = `<tr>
                        <td>${item.quantity}</td>
                        <td><a href="/pages/productDetail.html?productId=${item.productId}">${item.productName}</a></td>
                        <td>${item.unitPrice}</td>
                        <td>${item.discount}</td>
                    </tr>`;
                detailsElement.append(orderRow);

            });
            
                copySectionElement.style.display = "flex";
                copyUrlElement.addEventListener('click', async ev => {
                    try { 
                        //temp textarea for copy to clipboard functionality
                        var textarea = document.createElement("textarea");
                        const encodedContext = encodeURI(`{"subEntityId": "${order.orderId}"}`);
                        //form the deeplink
                        const deeplink = `https://teams.microsoft.com/l/entity/${env.TEAMS_APP_ID}/Orders?&context=${encodedContext}`;
                        textarea.value = deeplink;
                        document.body.appendChild(textarea);
                        textarea.select();
                        document.execCommand("copy"); //deprecated but there is an issue with navigator.clipboard api
                        document.body.removeChild(textarea); 
                        copyMsgElement.innerHTML="Link copied!"
                    
                    } catch (err) {
                        console.error('Failed to copy: ', err);
                      }});
            
        }else{
            errorMsgElement.innerText = `No order to show`;
            displayElement.style.display="none";
            orderDetails.style.display="none";
        }
    });
});
       
    }
    catch (error) {            // If here, we had some other error
        errorMsgElement.innerText = `Error: ${JSON.stringify(error)}`;
    }
}
```
##### Explanation for above code changes

We will use Microsoft Teams SDK to get the current tems context through which we can get the entityId and userId of the current user in the personal tab they are in.
That is why we use below lines of code to initialize and get context from teams
<pre>
microsoftTeams.initialize(async () => {
microsoftTeams.getContext(async (context)=> {      
</pre>

If the tab is opened using a query parameter (as done in the core lab) **OR**
there is an **subEntityId** in the teams context (which is passed in the deep link) then get the order details and display to the user. 
This is why we added an extra condition to the if statement here:

<pre>
   if (searchParams.has('orderId')<b>||context.subEntityId</b>) {
            const orderId = searchParams.get('orderId')<b>?searchParams.get('orderId'):context.subEntityId;</b>
</pre>

In the `copyUrlElement.addEventListener()` what goes on is explained next.

The **entityId** is basically used by Teams apps to identify it's own tab. When creating the deep link the **entityId** we use in `Orders` as we will be using `My Orders` tab as the landing tab when the link is opened.
The teams app id is taken from `.env` file, which is the id in the manifest file.
`encodedContext` is a JSON constant that defines the parameter(subEntityId)  to be passed to the tab, to display order information.

**4. client\myOrders.js**

Add import statement for Microsoft Teams SDK. 

```javascript
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
```

Using the `My Orders` tab as the base, we will redirect the deeplink to `Order details` page to show the order only if the **subEntitiyId** is present in the teams context.
Add an *if* condition in the function displayUI() after getting the teams context.

```javascript
 microsoftTeams.initialize(async () => {
            microsoftTeams.getContext(async (context) => {                
                if (context.subEntityId) {
                    window.location.href = `/pages/orderDetail.html?orderId=${context.subEntityId}`;
                } else { //..rest of the code
              }});});

```

The updated function definition looks like below:

```javascript
async function displayUI() {

    const displayElement = document.getElementById('content');
    const ordersElement = document.getElementById('orders');
    const messageDiv = document.getElementById('message');

    try {
        microsoftTeams.initialize(async () => {
            microsoftTeams.getContext(async (context) => {                
                if (context.subEntityId) {
                    window.location.href = `/pages/orderDetail.html?orderId=${context.subEntityId}`;
                } else {
                    const employee = await getLoggedInEmployee();
                    if (employee) {
                        displayElement.innerHTML = `<h3>Orders for ${employee.displayName}<h3>`;
                        employee.orders.forEach(order => {
                        const orderRow = document.createElement('tr');
                        orderRow.innerHTML = `<tr>
                            <td><a href="/pages/orderDetail.html?orderId=${order.orderId}">${order.orderId}</a></td>
                            <td>${(new Date(order.orderDate)).toDateString()}</td>
                            <td>${order.shipName}</td>
                            <td>${order.shipAddress}, ${order.shipCity} ${order.shipRegion || ''} ${order.shipPostalCode || ''} ${order.shipCountry}</td>
                            </tr>`;
                            ordersElement.append(orderRow);
                        });
                    }
                }
            });
        });
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }
}
```

**5. manifest\manifest.template.json**

Update the version number so it's greater than it was; for example if your manifest was version 1.4, make it 1.4.1 or 1.5.0. This is required in order for you to update the app in Teams.

~~~json
"version": "1.5.0"
~~~

### Exercise 2: Test the changes
---
Now that you have applied all code changes, let's test the features.

#### Step 1 : Create new teams app package

Make sure the env file is configured as per the sample file .env_Sample.
Make sure all npm packages are installed, run below script in the command line tool:

```nodejs
npm i
```
Create updated teams app package by running below script:
```nodejs
npm run package
```

#### Step 2: Start your local project

Now it's time to run your updated application and run it in Microsoft Teams. Start the application by running below command: 

```nodejs
npm start
```

#### Step 3: Upload the app package to Teams

In the Teams web or desktop UI, click "Apps" in the sidebar 1️⃣, then "Manage your apps" 2️⃣. At this point you have three choices:

* Upload a custom app (upload the app for yourself or a specific team or group chat) - this only appears if you have enabled "Upload custom apps" in your setup policy; this was a step in the previous lab
* Upload an app to your org's app catalog (upload the app for use within your organization) - this only appears if you are a tenant administrator
* Submit an app to your org (initiate a workflow asking a tenant administrator to install your app) - this appears for everyone

In this case, choose the first option 3️⃣.

<img src="../../assets/03-005-InstallApp-1.png?raw=true" alt="Upload the app"/>

Once uploaded, add the personal tab.

[add app](../../assets/deeplink-add-app.png)

#### Step 4 : Run the application in Teams client

Once you are in the application, go to `My orders` page and select any order.
Select **Copy order url**.

On selection, the message next to button changes from *Copy to clipboard* to *Link copied!*

Login as another user who has Northwind Order app installed in their teams.
Open the link in the browser. It should open in the personal tab with the order information displayed.

![order](../../assets/deeplink-working.gif)

