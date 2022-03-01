![Teams App Camp](../Assets/code-lab-banner.png)

## Lab A03: Create a Teams app with Azure AD Single Sign-On

This lab is part of Path A, which begins with a Northwind Orders application that already uses Azure AD.

In this lab you will extend the Northwind Orders application as a personal tab in a Teams application. That means the application needs to run in an IFrame; if your app includes code that prevents this, you may need to modify it to look for the Teams referrer URL. 

The Northwind Orders app doesn't check for IFrames, but Azure AD does, so the existing site won't work without modifications. To accomodate this and to prevent extra user logins, you'll set this tab up to use Teams' Azure AD Single Sign-on. 

The completed solution can be found in the [A03-TeamsSSO](../../A03-TeamsSSO/) folder, but the instructions will guide you through modifying the app running in your working folder. 

Note that as you complete the labs, the original app should still work outside of Teams! This is often a requirement of ISV's who have an app in market and need to serve an existing customer base outside of Teams.

* [Lab A01: Setting up the application with Azure AD](./Lab-A01.md)
* Lab A02: (there is no lab A02; please skip to A03)
* [Lab A03: Creating a Teams app with Azure ADO SSO](./Lab-A03.md) (📍You are here)
* [Lab A04: Teams styling and themes](./Lab-A04.md)
* [Lab A05: Add a Configurable Tab](./Lab-A05.md)
* [Lab A06: Add a Messaging Extension](./Lab-A06.md)
* [Lab A07: Add a Task Module and Deep Link](./Lab-A07.md)
* [Lab A08: Add support for selling your app in the Microsoft Teams store](./Lab-A08.md)

In this lab you will learn to:

- Create a Microsoft Teams app manifest and package that can be installed into Teams
- Update your Azure AD app registration to allow Teams to issue tokens on behalf of your application
- Use the Microsoft Teams JavaScript SDK to request an Azure AD access token
- Install and test your application in Microsoft Teams

### Features

- Microsoft Teams personal tab application displays the Northwind Orders web application
- Users sign into the Teams application transparently using Azure AD SSO
- Application alters its appearance (hides the top navigation) when running in Teams, allowing Teams tab navigation instead

### Project structure

The project structure when you start of this lab and end of this lab is as follows.
Use this depiction for comparison.
On your left is the contents of folder  `A01-Start-AAD` and on your right is the contents of folder `A03-TeamsSSO`.

- 🆕 New files/folders

- 🔺Files changed


<table>
<tr>
<th >Project Structure Before </th>
<th>Project Structure After</th>
</tr>
<tr>
<td valign="top" >
<pre>
A01-Start-AAD
    ├── client
    │   ├── components
    │       ├── 🔺navigation.js
    │   └── identity
    │       ├── 🔺identityClient.js
    │       └── userPanel.js
    ├── modules
    │   └── env.js
    │   └── northwindDataService.js
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
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── server
    │   └── constants.js
    │   └── identityService.js
    │   └── northwindDataService.js
    │   └── server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
<td>
<pre>
A03-TeamsSSO
    ├── client
    │   ├── components
    │       ├── 🔺navigation.js
    │   └── identity
    │       ├── 🔺identityClient.js
    │       └── userPanel.js
    ├── modules
    │   └── env.js
    │   └── northwindDataService.js
    │   └── 🆕teamsHelpers.js
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
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── 🆕manifest
    │   └── 🆕makePackage.js
    │   └── 🆕manifest.template.json
    │   └── 🆕northwind32.png
    │   └── 🆕northwind192.png
    ├── server
    │   └── constants.js
    │   └── identityService.js
    │   └── northwindDataService.js
    │   └── server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
</tr>
</table>


### Exercise 1: Authorize Microsoft Teams to log users into your application

Microsoft Teams provides a Single Sign-On (SSO) capability so users are silently logged into your application using the same credentials they used to log into Microsoft Teams. This requires giving Microsoft Teams permission to issue Azure AD tokens on behalf of your application. In this exercise, you'll provide that permission.

#### Step 1: Return to your app registration

Return to the [Azure AD admin portal](https://aad.portal.azure.com/) and make sure you're logged in as the administrator of your development tenant. Click "Azure Active Directory" 1️⃣ and then "App Registrations" 2️⃣.

![Return to your app registration](../Assets/03-001-AppRegistrationUpdate-1.png)

Select the app you registered earlier to view the application overview.

#### Step 2: Add the Teams client applications

Click "Expose an API" 1️⃣ and then "+ Add a client application" 2️⃣.

![Open the Expose an API screen](../Assets/03-002-AppRegistrationUpdate-2.png)

Paste the ID for the Teams mobile or desktop app, `1fec8e78-bce4-4aaf-ab1b-5451cc387264` into the flyout 1️⃣ and check the scope you created earlier 2️⃣ to allow Teams to issue tokens for that scope. Then click "Add application" 3️⃣ to save your work.

Repeat the process for the Teams web application, `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`.

![Add a client application](../Assets/03-003-AppRegistrationUpdate-3.png)

### Exercise 2: Create the Teams application package

Microsoft Teams applications don't run "inside" of Microsoft Teams, they just appear in the Teams user interface. A tab in Teams is just a web page, which could be hosted anywhere as long as the Teams client can reach it. 

To create a Teams application, you need to create a file called *manifest.json* which contains the information Teams needs to display the app, such as the URL of the Northwind Orders application. This file is placed in a .zip file along with the application icons, and the resulting application package is uploaded into Teams or distributed through the Teams app store.

In this exercise you'll create a manifest.json file and application package for the Northwind Orders app and upload it into Microsoft Teams.

#### Step 1: Copy the *manifest* folder to your working directory

Many developers use the [Teams Developer Portal](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/teams-developer-portal?WT.mc_id=m365-58890-cxa) to create an app package; this is preferred by many enterprise developer and systems integrators. However ISV's may want to keep the app package settings in their source control system, and that's the approach used in the lab. It's just a zip file; you can create it any way you want!

Go to your local copy of the `A03-TeamsSSO` folder on your computer and copy the *manifest* folder into the working folder you used in the previous lab. This folder contains a template for building the manifest.json file.

#### Step 2: Examine the manifest template

In the manifest folder you just copied, open [manifest.template.json](../../A03-TeamsSSO/manifest/manifest.template.json) in your code editor. This is the JSON that Teams needs to display your application.

Notice that the template contains tokens such as`<HOSTNAME>` and `<CLIENT_ID>`. A small build script will take these values from your .env file and plug them into the manifest. However there's one token, `<TEAMS_APP_ID>` that's not yet in the .env file; we'll add that in the next step.

Examine the `staticTabs` property in the manifest. It defines two tabs, one for the "My Orders" page and one for the "Products" page. The `contentUrl` is used within the Teams application, and `websiteUrl` is used if Teams can't render the tab and needs to launch it in a regular web browser. The Northwind Orders app will use the same code URL's for both.

~~~json
"staticTabs": [
  {
    "entityId": "Orders",
    "name": "My Orders",
    "contentUrl": "https://<HOSTNAME>/pages/myOrders.html",
    "websiteUrl": "https://<HOSTNAME>/pages/myOrders.html",
    "scopes": [
      "personal"
    ]
  },
  {
    "entityId": "Products",
    "name": "Products",
    "contentUrl": "https://<HOSTNAME>/pages/categories.html",
    "websiteUrl": "https://<HOSTNAME>/pages/categories.html",
    "scopes": [
      "personal"
    ]

~~~

Now examine the `webApplicationInfo` property. It contains the information Teams needs to obtain an access token using Single Sign On.

~~~json
  "webApplicationInfo": {
      "id": "<CLIENT_ID>",
      "resource": "api://<HOSTNAME>/<CLIENT_ID>"
  }
~~~

#### Step 3: Add the Teams App ID to the .env file

Open the .env file in your working directory and add this line:

~~~text
TEAMS_APP_ID=1331dbd6-08eb-4123-9713-017d9e0fc04a
~~~

You should generate a different GUID for each application you register; this one is just here for your convenience. We could have hard-coded the app ID in the manifest.json template, but there are times when you need it in your code, so this will make that possible in the future.

#### Step 4: Update your package.json file

Open the **package.json** file in your working directory and add a script that will generate the app package. The [script code](../../A03-TeamsSSO/manifest/makePackage.js) is in the manifest folder you just copied, so we just need to declare it in package.json. This is what `scripts` property should look like when you're done.

~~~json
"scripts": {
  "start": "nodemon server/server.js",
  "debug": "nodemon --inspect server/server.js",
  "package": "node manifest/makePackage.js"
},
~~~

The script uses an npm package called "adm-zip" to create the .zip file, so you need to add that as a development dependency. Update the `devDependencies` property to include it like this:

~~~json
  "devDependencies": {
    "@types/express": "^4.17.2",
    "@types/request": "^2.48.3",
    "nodemon": "^2.0.13",
    "adm-zip": "^0.4.16"
  }
~~~

Then, from a command line in your working directory, install the package by typing

~~~shell
npm install
~~~

#### Step 5: Build the package

Now you can build a new package at any time with this command:

~~~shell
npm run package
~~~

Go ahead and run it, and two new files, manifest.json and northwind.zip (the app package) should appear in your manifest folder.

### Exercise 3: Modify the application source code

#### Step 1: Add a module with Teams helper functions

Create a file called teamsHelpers.js in the client/modules folder, and paste in this code:

~~~javascript
// async function returns true if we're running in Teams
export async function inTeams() {
    return (window.parent === window.self && window.nativeInterface) ||
        window.navigator.userAgent.includes("Teams/") ||
        window.name === "embedded-page-container" ||
        window.name === "extension-tab-frame";
}
~~~

This function will allow your code to determine if it's running in Microsoft Teams, and is used in this and future labs.

---
> NOTE: There is no official way to do this with the Teams JavaScript SDK v1. The official guidance is to pass some indication that the app is running in Teams via the app's URL path or query string. This is not a single-page app, however, so we're using a workaround to avoid updating every page to generate hyperlinks that support Teams-specific URLs. This workaround is used by the [yo teams generator](https://github.com/wictorwilen/msteams-react-base-component/blob/master/src/useTeams.ts#L10), so it's well tested and in wide use, though not officially supported.
---

#### Step 2: Update the login code for Teams SSO

Open the client/identity/identityClient.js file and add these import statements near the top.

~~~javascript
import { inTeams } from '/modules/teamsHelpers.js';
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
~~~

The first import, of course, is the inTeams() function we just added. It's a [JavaScript module](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Modules) so there is no bundling; the browser will resolve the import at runtime. 

The second import will load the Teams JavaScript SDK, which creates a global object `microsoftTeams` that we can use to access the SDK. You could also load it using a `<script>` tag or, if you bundle your client-side JavaScript, using the [@microsoft/teams-js](https://www.npmjs.com/package/@microsoft/teams-js) npm package.

Now modify the `getAccessToken2()` function to include this code at the top:

~~~javascript
    if (await inTeams()) {

        microsoftTeams.initialize();
        const accessToken = await new Promise((resolve, reject) => {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (result) => { resolve(result); },
                failureCallback: (error) => { reject(error); }
            });
        });
        return accessToken;

    } else {
      // existing code
    }
~~~

This codechecks to see if it's running in Teams and if so, uses the Teams JavaScript SDK's `getAuthToken()` function to get the access token needed to call the server.

The completed `getAccessToken2()` function should look like this:

~~~javascript
async function getAccessToken2() {

    if (await inTeams()) {

        microsoftTeams.initialize();
        const accessToken = await new Promise((resolve, reject) => {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (result) => { resolve(result); },
                failureCallback: (error) => { reject(error); }
            });
        });
        return accessToken;

    } else {

        // If we were waiting for a redirect with an auth code, handle it here
        await msalClient.handleRedirectPromise();

        try {
            await msalClient.ssoSilent(msalRequest);
        } catch (error) {
            await msalClient.loginRedirect(msalRequest);
        }

        const accounts = msalClient.getAllAccounts();
        if (accounts.length === 1) {
            msalRequest.account = accounts[0];
        } else {
            throw ("Error: Too many or no accounts logged in");
        }

        let accessToken;
        try {
            const tokenResponse = await msalClient.acquireTokenSilent(msalRequest);
            accessToken = tokenResponse.accessToken;
            return accessToken;
        } catch (error) {
            if (error instanceof msal.InteractionRequiredAuthError) {
                console.warn("Silent token acquisition failed; acquiring token using redirect");
                this.msalClient.acquireTokenRedirect(this.request);
            } else {
                throw (error);
            }
        }
    }
}
~~~

#### Step 3: Hide the navigation within Teams

Microsoft Teams already has multiple levels of navigation, including multiple tabs as configured in the previous exercise. So the applications' built-in navigation is redundant in Teams.

To hide the built-in navigation in Teams, open the client/components/navigation.js file and add this import statement at the top.

~~~javascript
import { inTeams } from '../modules/teamsHelpers.js';
~~~

Now modify the `connectedCallback()` function, which displays the navigation web component, to skip rendering if it's running in Teams. The resulting function should look like this:

~~~javascript
    async connectedCallback() {

        if (!(await inTeams())) {
            let listItemHtml = "";
            topNavLinks.forEach(link => {
                if (window.location.href.indexOf(link.url) < 0) {
                    listItemHtml += '<li><a href="' + link.url + '">' + link.text + '</a></li>';
                } else {
                    return listItemHtml += '<li><a href="' + link.url + '" class="selected">' + link.text + '</a></li>';
                }
            });
            this.innerHTML = `
                <ul class="topnav">${listItemHtml}</ul>
            `;
        }

    }
~~~

---
> NOTE: Web components are encapsulated custom HTML elements. They're not a Teams thing, nor do they use React or another UI library; they're built right into modern web browsers. You can learn more [in this article](https://developer.mozilla.org/en-US/docs/Web/Web_Components.)
---

### Exercise 4: Test your application in Microsoft Teams

#### Step 1: Start the application

Now it's time to run your updated application and run it in Microsoft Teams. Start the application with this command:

~~~shell
npm start
~~~

#### Step 2: Upload the app package

In the Teams web or desktop UI, click "Apps" in the sidebar 1️⃣, then "Manage your apps" 2️⃣. At this point you have three choices:

* Upload a custom app (upload the app for yourself or a specific team or group chat) - this only appears if you have enabled "Upload custom apps" in your setup policy; this was a step in the previous lab
* Upload an app to your org's app catalog (upload the app for use within your organization) - this only appears if you are a tenant administrator
* Submit an app to your org (initiate a workflow asking a tenant administrator to install your app) - this appears for everyone

In this case, choose the first option 3️⃣.

![Upload the app](../Assets/03-005-InstallApp-1.png)

Navigate to the Northwind.zip file in your manifest directory and upload it. Teams will display the application information; click the "Add" button to install it for your personal use.

![Upload the app](../Assets/03-006-InstallApp-2.png)

#### Step 3: Run the application

The application should appear without any login prompt. The app's navigation should not be displayed; instead users can navigate to "My Orders" or "Products" using the tabs in the Teams app.

![Run the app](../Assets/03-007-RunApp-1.png)

---
> CHALLENGE: Notice the logout button doesn't do anything in Teams! If you wish, hide the logout button just as you hid the navigation bar. The code is in client/identity/userPanel.js.
---

### Known issues

For the latest issues, or to file a bug report, see the [github issues list](https://github.com/OfficeDev/TeamsAppCamp1/issues) for this repository.

### References

[Single sign-on (SSO) support for Tabs](https://docs.microsoft.com/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso?WT.mc_id=m365-58890-cxa)


