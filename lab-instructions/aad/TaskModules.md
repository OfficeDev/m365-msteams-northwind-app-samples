![Teams App Camp](../../assets/code-lab-banner.png)

## Add a Task Module 

This lab is part of extending with capabilities for your teams app which begins with a Northwind Orders core application using the `aad` path.
> Complete labs [A01](A01-begin-app.md)-[A03](A03-after-apply-styling.md) to get to the Northwind Orders core application

**Task modules** are modal pop-up experiences in Teams application to run your app's own html or JavaScript code. 


In this exercise you will learn new concepts as below:

- [Task modules](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-task-modules?WT.mc_id=m365-58890-cxa)


### Features

- In the application's order details page, add a button to open a web based form to add notes into a order database. 


### Exercise 1: Code changes
---

#### Step 1: Add new files

There are new files and folders that you need to add into the project.

**1.\client\pages\orderNotesForm.html**

Create a new file `orderNotesForm.html` in the path `\client\pages`and copy below form HTML into it.

```HTML
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8" />
    <title>Task Module: Update order notes</title>
    <link rel="stylesheet" href="/northwind.css" />
    <link rel="icon" href="data:;base64,="> <!-- Suppress favicon error -->
    <script src="https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js"
        asp-append-version="true"></script>
</head>
    <body>
        <script>
            microsoftTeams.initialize();
            function validateForm() {
                var orderFormInfo = {
                    notes: document.forms["orderForm"]["notes"].value,
                }                        
                microsoftTeams.tasks.submitTask(orderFormInfo);
                return true;
            }
        </script>            
        <form id="orderForm" onSubmit="validateForm()">
            <div>                
               <textarea id="notes" name="notes" rows="4" cols="50"></textarea>               
            </div>
            <div>
                <button type="submit" tabindex="2">Save</button>
            </div>
        </form>
    </body>

</html>

```

#### Step 2: Update existing files


**1.\client\pages\orderDetail.html**

Add a content area to display the comments that will get added into each order.
Add a new button to open the web based form as a task module in the `orderDetail.html` page.

Copy below block of code and paste it above the `<table>` element.

```html
 <div id="orderContent">
    
</div>
    <br/>
 <br/>
    <button id="btnTaskModule">Add notes</button>
```
**2.\client\pages\orderDetail.js**

In the displayUI() function define two constants to get the above two HTML elements.

```javascript
 const btnTaskModuleElement = document.getElementById('btnTaskModule');
 const orderElement=document.getElementById('orderContent');
```
To open the task module (web form), add an event listener for the button we added earlier.
Paste below code in the dislayUI() function in the end before closing the `try`.

```javascript
 btnTaskModuleElement.addEventListener('click',  ev => {  
            let submitHandler = (err, result) => {                 
                const postDate = new Date().toLocaleString()
                const newComment = document.createElement('p');  
                newComment.innerHTML=`<div><b>Posted on:</b>${postDate}</div>
                <div><b>Notes:</b>${result.notes}</div><br/>
                -----------------------------` 
                orderElement.append(newComment);
            }                     
            let taskInfo = {
                title: null,
                height: null,
                width: null,
                url: null,
                card: null,
                fallbackUrl: null,
                completionBotId: null,
            };
            taskInfo.url = `https://${window.location.hostname}/pages/orderNotesForm.html`;
            taskInfo.title = "Task module order notes";
            taskInfo.height = 210;
            taskInfo.width = 400;
           
            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        });
```
This code `microsoftTeams.tasks.startTask(taskInfo, submitHandler);` will ensure that the web based form opens as a dialog and the call back function `submitHandler` will take care of the values passed from the form. You can perform any action on the results.
Here we are appending it to the content area to display. See line `orderElement.append(newComment);`.

The final look of displayUI() function is as below

```javascript
import {
    getOrder
} from '../modules/northwindDataService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    const btnTaskModuleElement = document.getElementById('btnTaskModule');
    const orderElement=document.getElementById('orderContent');
    try {

        const searchParams = new URLSearchParams(window.location.search);
        if (searchParams.has('orderId')) {
            const orderId = searchParams.get('orderId');

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

        }
        btnTaskModuleElement.addEventListener('click',  ev => {  
            let submitHandler = (err, result) => {                 
                const postDate = new Date().toLocaleString()
                const newComment = document.createElement('p');  
                newComment.innerHTML=`<div><b>Posted on:</b>${postDate}</div>
                <div><b>Notes:</b>${result.notes}</div><br/>
                -----------------------------` 
                orderElement.append(newComment);
            }                     
            let taskInfo = {
                title: null,
                height: null,
                width: null,
                url: null,
                card: null,
                fallbackUrl: null,
                completionBotId: null,
            };
            taskInfo.url = `https://${window.location.hostname}/pages/orderNotesForm.html`;
            taskInfo.title = "Task module order notes";
            taskInfo.height = 210;
            taskInfo.width = 400;
           
            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        });
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();
```


**6. manifest\manifest.template.json**

Update the version number so it's greater than it was; for example if your manifest was version 1.4, make it 1.4.1 or 1.5.0. This is required in order for you to update the app in Teams.

~~~json
"version": "1.5.0"
~~~

### Exercise 2: Test the changes
---
Now that you have applied all code changes, let's test the features.

#### Step 1 : Create new teams app package
Make sure the env file is configured as per the sample file .env_Sample.
Create updated teams app package by running below script:
```nodejs
npm run package
```

#### Step 2: Start your local project

Now it's time to run your updated application and run it in Microsoft Teams. Start the application by running below command: 

```nodejs
npm start
```

#### Step 3: Upload the app package
In the Teams web or desktop UI, click "Apps" in the sidebar 1️⃣, then "Manage your apps" 2️⃣. At this point you have three choices:

* Upload a custom app (upload the app for yourself or a specific team or group chat) - this only appears if you have enabled "Upload custom apps" in your setup policy; this was a step in the previous lab
* Upload an app to your org's app catalog (upload the app for use within your organization) - this only appears if you are a tenant administrator
* Submit an app to your org (initiate a workflow asking a tenant administrator to install your app) - this appears for everyone

In this case, choose the first option 3️⃣.

<img src="../../assets/03-005-InstallApp-1.png?raw=true" alt="Upload the app"/>

Navigate to the Northwind.zip file in your manifest directory and upload it. Add the personal tab.


#### Step 4 : Run the application in Teams client

Once you are in the application, go to `My orders` page and select any order as shown below:


### Known issues

The task module (dialog) has to be closed manually.
