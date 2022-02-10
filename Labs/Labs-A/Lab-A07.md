## Lab A07: Add a Task Module and Deep Link

In this lab you will begin with the application in folder `A06-MessagingExtension`, make changes as per the steps below to achieve what is in the folder `A07-TaskModule`.
See project structures comparison in Exercise 2.

* [Lab A01: Setting up the application with Azure AD](./Lab-A01.md)
* [Lab A02: Setting up your Microsoft 365 Tenant](./Lab-A02.md)
* [Lab A03: Creating a Teams app with Azure ADO SSO](./Lab-A03.md)
* [Lab A04: Teams styling and themes](./Lab-A04.md)
* [Lab A05: Add a Configurable Tab](./Lab-A05.md)
* [Lab A06: Add a Messaging Extension](./Lab-A06.md)
* [Lab A07: Add a Task Module and Deep Link](./Lab-A07.md)**(You are here)**
* [Lab A08: Add support for selling your app in the Microsoft Teams store](./Lab-A08.md)

In this exercise you will learn new concepts as below:

- Task modules
- Deep linking
- Microsoft Graph to pull Microsoft365 data


### Features


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
A06-MessagingExtension
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
    │   └── identityService.js
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
A07-TaskModule
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
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
    │   └── identityService.js
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

In the project structure, on the right under `A07-TaskModule`, you will see **bold** files.
They are the new files and folders that you need to add into the project structure.


#### Step 2: Update existing files
In the project structure, on the right under `A07-TaskModule`, you will see *italics* files.
They are the files that were updated to add the new features.
Let's take files one by one to understand what changes you need to make for this exercise. 

**1.**
 
### Exercise 2: Test the changes

- Install new packages by running 

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


### Known issues



### References




