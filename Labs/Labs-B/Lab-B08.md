## Lab B08: Set up and integrate with licensing sample and App Source simulator

This lab is part of Path B, which begins with a Northwind Orders application that uses an identity system other than Azure AD, and then adds Azure AD SSO.

* [Lab B01: Start an application with bespoke authentication](./Lab-B01.md)
* [Lab B02: Create a teams app](./Lab-B02.md)
* [Lab B03: Make existing teams app use Azure ADO SSO](./Lab-B03.md)
* [Lab B04: Teams styling and themes](./Lab-B04.md)
* [Lab B05: Add a Configurable Tab](./Lab-B05.md)
* [Lab B06: Add a Messaging Extension](./Lab-B06.md)
* [Lab B07: Add a Task Module and Deep Link](./Lab-B07.md)
* [Lab B08: Add support for selling your app in the Microsoft Teams store](./Lab-B08.md) (📍You are here)


In this lab you will learn to:

- Deploy the App Source simulator and sample SaaS fulfillment and licensing service in Microsoft Azure so you can test it
- Observe the interactions between App Source and a SaaS landing page in a simulated environment
- Connect the Northwind Orders application to the sample SaaS licensing service to enforce licenses for Microsoft Teams users

### Features

- App Source simulator where a customer can "purchase" a subscription to your application
- Sample web service that fulfills this purchase and manages licenses for Microsoft Teams users to use the Northwind Orders application
- Northwind Orders application checks to ensure Microsoft Teams users are licensed or displays an error page

### Exercise 1: Download and install the monetization sample

> Insert Rabia's exercises here
> Note the student should obtain and save these values (if it's hidden in scripts I'll write up how to find them in the AAD portal):
> 
> SAAS_API=https://mySaasApiProject.azurewebsites.net/api/Subscriptions/CheckOrActivateLicense
> SAAS_SCOPES=api://11111111-1111-1111-1111-111111111111/user_impersonation

### Exercise n: Grant the Northwind Orders app permission to call the licensing service in Azure

In this exercise and the next, you will connect the Northwind Orders application to the sample licensing service you just installed. This will allow you to simulate purchasing the Northwind Orders application in the App Source simulator and enforcing the licenses in Microsoft Teams.

The licensing web service is secured using Azure AD, so in order to call it the Northwind Orders app will acquire an access token to call the licenisng service on behalf of the logged in user.

#### Step 1: Return to the Northwind Orders app registration

Return to the [Azure AD admin portal](https://aad.portal.azure.com/) and make sure you're logged in as the administrator of your development tenant. Click "Azure Active Directory" 1️⃣ and then "App Registrations" 2️⃣.

![Return to your app registration](../Assets/03-001-AppRegistrationUpdate-1.png)

Select the app you registered earlier to view the application overview.

#### Step 2: Add permission to call the licensing application

In the left navigation, click "API permissions" 1️⃣ and then "+ Add a permission" 2️⃣.

![Add Permission](../Assets/08-100-Add-Permission-0.png)

In the flyout, select the "My APIs" tab 1️⃣ and then find the licensing service you installed earlier in this lab and click on it. By default, it will be called the "Contoso Monetization Code Sample Web API" 2️⃣.

![Add permission](../Assets/08-100-Add-Permission.png)

Now select "Delegated permissions" 1️⃣ and the one scope exposed by the licensing web API, "user_impersonation", will be displayed. Check this permission 2️⃣ and click "Add permissions" 3️⃣.

#### Step 3: Consent to the permission

You have added the permission but nobody has consented to it. Fortunately you're an administrator and can grant your consent from this same screen! Just click the "Grant admin consent for <tenant>" button 1️⃣ and then agree to grant the consent 2️⃣. When this is complete, the message "Granted for <tenant>" 3️⃣ should be displayed for each permission.

![Grant consent](../Assets//08-102-Grant-Consent.png)

### Exercise n+1: Update the Northwind Orders app to call the licensing service in Azure

#### Step 1: Add licensing service information to your .env file

In your working folder, open the .env file and add 

~~~text
SAAS_API=<URL of your SaaS service in Azure such as https://mySaasApiProject.azurewebsites.net/api/Subscriptions/CheckOrActivateLicense>
SAAS_SCOPES=<scope of your SaaS service such as api://11111111-1111-1111-1111-111111111111/user_impersonation>
OFFER_ID=contoso_o365_addin
~~~

#### Step 2: Add a server side function to validate the user has a license

In your working folder, create a new file /server/validateLicenseService.js and paste in this code (or copy the file from [here](../../B08-Monetization/server/northwindLicenseService.js)).
)
~~~javascript
import aad from 'azure-ad-jwt';
import fetch from 'node-fetch';

export async function validateLicense(thisAppAccessToken) {

    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    return new Promise((resolve, reject) => {

        aad.verify(thisAppAccessToken, { audience: audience }, async (err, result) => {
            if (result) {
                const licensingAppUrl = `${process.env.SAAS_API}/${process.env.OFFER_ID}`
                const licensingAppAccessToken = await getOboAccessToken(thisAppAccessToken);
                if (licensingAppAccessToken === "interaction_required") {
                    reject({ "status":401, "message": "Interaction required"});
                }
                
                const licensingResponse = await fetch(licensingAppUrl, {
                    method: "POST",
                    headers: {
                        "Accept": "application/json",
                        "Content-Type": "application/json",
                        "Authorization" :`Bearer ${licensingAppAccessToken}`
                    }
                });
                if (licensingResponse.ok) {
                    const licensingData = await licensingResponse.json();
                    console.log(licensingData.reason);
                    resolve(licensingData);
                } else {
                    reject({ "status": licensingResponse.status, "message": licensingResponse.statusText });
                }
            } else {
                reject({ "status": 401, "message": "Invalid client access token in northwindLicenseService.js"});
            }
        });
    });

}

// TODO: Securely the results of this function for the lifetime of the resulting token
async function getOboAccessToken(clientSideToken) {

    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const scopes = process.env.SAAS_SCOPES;

    // Use On Behalf Of flow to exchange the client-side token for an
    // access token with the needed permissions
    
    const INTERACTION_REQUIRED_STATUS_TEXT = "interaction_required";
    const url = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";
    const params = {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: clientSideToken,
        requested_token_use: "on_behalf_of",
        scope: scopes
    };

    const accessTokenQueryParams = new URLSearchParams(params).toString();
    try {
        const oboResponse = await fetch(url, {
            method: "POST",
            body: accessTokenQueryParams,
            headers: {
                Accept: "application/json",
                "Content-Type": "application/x-www-form-urlencoded"
            }
        });

        const oboData = await oboResponse.json();
        if (oboResponse.status !== 200) {
            // We got an error on the OBO request. Check if it is consent required.
            if (oboData.error.toLowerCase() === 'invalid_grant' ||
                oboData.error.toLowerCase() === 'interaction_required') {
                throw (INTERACTION_REQUIRED_STATUS_TEXT);
            } else {
                console.log(`Error returned in OBO: ${JSON.stringify(oboData)}`);
                throw (`Error in OBO exchange ${oboResponse.status}: ${oboResponse.statusText}`);
            }
        }
        return oboData.access_token;
    } catch (error) {
        return error;
    }

}
~~~

In Lab B03, you called the Microsoft Graph API using application permissions. This code calls the licensing service using delegated permissions, meaning that the application is acting on behalf of the user.

To do this, the code uses the [On Behalf Of flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the incoming access token (targeted for the Northwind Orders app) for a new access token that is targeted for the Licensing service application.

#### Step 3: Add a server side API to validate the user's license

Now that we have a function that checks the user's license on the server side, we need to add a POST request to our service that calls the function.

In your working folder, locate the file server/server.js and open it in your code editor.

Add these lines to the top of the file:

~~~javascript
import aad from 'azure-ad-jwt';
import { validateLicense } from './northwindLicenseService.js';
~~~

Now, immediately below the call to `await initializeIdentityService()`, add this code:

~~~javascript
// Web service validates a user's license
app.post('/api/validateLicense', async (req, res) => {

  try {
    const token = req.headers['authorization'].split(' ')[1];

    try {
      let hasLicense = await validateLicense(token);
      res.send(JSON.stringify({ "validLicense" : hasLicense }));
    }
    catch (error) {
      console.log (`Error ${error.status} in validateLicense(): ${error.message}`);
      res.status(error.status).send(error.message);
    }
  }
  catch (error) {
      console.log(`Error in /api/validateAadLogin handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});
~~~

#### Step 4: Add client side pages to display a license error

Add a new file, client/pages/needLicense.html and paste in this markup, or copy the file from [here](../../B08-Monetization/client/pages/needLicense.html).

~~~html
<!doctype html>
<html>

<head>
    <meta charset="UTF-8" />
    <title>Northwind Privacy</title>
    <link rel="icon" href="data:;base64,="> <!-- Suppress favicon error -->
    <link rel="stylesheet" href="/northwind.css" />
   
</head>

<body class="ms-Fabric" dir="ltr">   
    <h1>Sorry you need a valid license to use this application</h1>
    <p>Please purchase a license from the Microsoft Teams store.       
    </p>
    <div id="errorMsg"></div>
    <script type="module" src="needLicense.js"></script>
</body>

</html>
~~~

To provide the JavaScript for the new page, create a file /client/pages/needLicense.js and paste in this code, or copy the file from [here](../../B08-Monetization/client/pages/needLicense.js).

~~~javascript
const searchParams = new URLSearchParams(window.location.search);
if (searchParams.has('error')) {
    const error = searchParams.get('error');
    const displayElementError = document.getElementById('errorMsg');
    displayElementError.innerHTML = error;  
}
~~~










#### Step 5: Add client side function to check if the user has a license

Add a new file, client/modules/northwindLicensing.js and paste in the following code, or copy the file from [here](../../B08-Monetization/client/modules/northwindLicensing.js). This code calls the server-side API we just added using an Azure AD token obtained using Microsoft Teams SSO.

~~~javascript
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';

export async function hasValidLicense() {

    await new Promise((resolve, reject) => {
        microsoftTeams.initialize(() => { resolve(); });
    });

    const authToken = await new Promise((resolve, reject) => {
        microsoftTeams.authentication.getAuthToken({
            successCallback: (result) => { resolve(result); },
            failureCallback: (error) => { reject(error); }
        });
    });

    const response = await fetch(`/api/validateLicense`, {
        "method": "post",
        "headers": {
            "content-type": "application/json",
            "authorization": `Bearer ${authToken}`
        },
        "cache": "no-cache"
    });
   
    if (response.ok) {

        const data = await response.json();
        return data.validLicense;

    } else {

        const error = await response.json();
        console.log(`ERROR: ${error}`);

    }

}

~~~

#### Step 6: Add client side call to check the license on every request

Open the file client/identity/userPanel.js in your code editor. This is a web component that displays the user's picture and name on every page, so it's an easy place to check the license.

Add these lines at the top of the file:

~~~javascript
import { inTeams } from '../modules/teamsHelpers.js';
import { hasValidLicense } from '../modules/northwindLicensing.js';
~~~

Now add this code in the `else` clause within the `connectedCallback()` function:

~~~javascript
    if (await inTeams()) {
        const validLicense = await hasValidLicense();  
        if (validLicense.status && validLicense.status.toString().toLowerCase()==="failure") {
                window.location.href =`/pages/needLicense.html?error=${validLicense.reason}`;
        }    
    }
~~~

The completed userPanel.js should look like this:

~~~javascript
import {
    getLoggedInEmployee,
    logoff
} from './identityClient.js';
import { inTeams } from '../modules/teamsHelpers.js';
import { hasValidLicense } from '../modules/northwindLicensing.js';

class northwindUserPanel extends HTMLElement {

    async connectedCallback() {

        const employee = await getLoggedInEmployee();

        if (!employee) {

            logoff();

        } else {

            if (await inTeams()) {
                const validLicense = await hasValidLicense();  
                if (validLicense.status && validLicense.status.toString().toLowerCase()==="failure") {
                     window.location.href =`/pages/needLicense.html?error=${validLicense.reason}`;
                }    
            }

            this.innerHTML = `<div class="userPanel">
                <img src="data:image/bmp;base64,${employee.photo}"></img>
                <p>${employee.displayName}</p>
                <p>${employee.jobTitle}</p>
                <hr />
                <button id="logout">Log out</button>
            </div>
            `;

            const logoutButton = document.getElementById('logout');
            logoutButton.addEventListener('click', async ev => {
                logoff();
            });
        }
    }
}

// Define the web component and insert an instance at the top of the page
customElements.define('northwind-user-panel', northwindUserPanel);
const panel = document.createElement('northwind-user-panel');
document.body.insertBefore(panel, document.body.firstChild);
~~~

> NOTE: There are many ways to make the license check more robust, such as checking it on every web service call and caching this on the server side to avoid excessive calls to the licensing server, however this is just a lab so we wanted to keep it simple.

### Exercise n+2: Examine the Application Code

TO BE PROVIDED


### Known issues

For the latest issues, or to file a bug report, see the [github issues list](https://github.com/OfficeDev/TeamsAppCamp1/issues) for this repository.

### References

* [Create a configuration page](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/create-tab-pages/configuration-page)



