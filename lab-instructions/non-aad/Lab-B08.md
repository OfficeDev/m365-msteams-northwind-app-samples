![Teams App Camp](../Assets/code-lab-banner.png)

## Lab B08: Set up and integrate with licensing sample and App Source simulator

This lab is part of Path B, which begins with a Northwind Orders application that uses an identity system other than Azure AD, and then adds Azure AD SSO.

* [Lab B01: Start an application with bespoke authentication](./Lab-B01.md)
* [Lab B02: Create a teams app](./Lab-B02.md)
* [Lab B03: Make existing teams app use Azure AD SSO](./Lab-B03.md)
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

### Project structure
The project structure when you start of this lab and end of this lab is as follows.
Use this depiction for comparison.
- 🆕 New files/folders

- 🔺Files changed
<table>
<tr>
<th>Project Structure Before </th>
<th>Project Structure After</th>
</tr>
<tr>
<td valign="top" >
<pre>
B07-TaskModule
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── aadLogin.html
    │       └── aadLogin.js
    │       ├── identityClient.js
    │       └── login.html
    │       └── login.js
    │       └── teamsLoginLauncher.html
    │       └── teamsLoginLauncher.js
    │       └── 🔺userPanel.js
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
    │   └── orderChatCard.js
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
    │   └── 🔺manifest.template.json
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
    │   └── 🔺server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
<td>
<pre>
B08-Monetization
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── aadLogin.html
    │       └── aadLogin.js
    │       ├── identityClient.js
    │       └── login.html
    │       └── login.js
    │       └── teamsLoginLauncher.html
    │       └── teamsLoginLauncher.js
    │       └── 🔺userPanel.js
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
    │   └── 🆕northwindLicensing.js
    │   └── orderChatCard.js
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── 🆕needLicense.html
    │   └── 🆕needLicense.js
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
    │   └── 🔺manifest.template.json
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
    │   └── 🆕northwindLicenseService.js
    │   └── 🔺server.js
    ├── 🔺.env_Sample
    ├── .gitignore
    ├── 🔺package.json
    ├── README.md
</pre>
</td>
</tr>
</table>


### Exercise 1: Download and install the monetization sample

To complete this lab you'll need to set up a mock App source simulator, as we cannot test apps in Microsoft's real App source. You will also need a sample SaaS fulfillment and licensing service in Azure which can be later replaced by your company's services. 

To help you succeed at this, we have set up some scripts that you can run in PowerShell in order to deploy the needed resources in Azure as well as get your mock simulator and licensing services up and running in few minutes.

In this exercise you'll create three Azure Active Directory applications and their supporting infrastructure using automated deployment scripts called [ARM templates](https://docs.microsoft.com/en-us/azure/azure-resource-manager/templates/overview?WT.mc_id=m365-58890-cxa).

- Contoso Monetization Code Sample Web App
- Contoso Monetization Code Sample Web API
- Contoso Monetization Code Sample App source

 #### Step 1: Install the prerequisites

> You would not come this far without Microsoft365 developer tenant as Global Admin and an Azure subscription. Below are the rest of the prerequisites.
 
- Install [PowerShell 7](https://github.com/PowerShell/PowerShell/releases/tag/v7.1.4)

- Install the following PowerShell modules:
  - [Microsoft Graph PowerShell SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell#powershell-gallery)

      ``` command
      Install-Module Microsoft.Graph -AllowClobber -Force
      ```

  - [Azure Az PowerShell module](https://docs.microsoft.com/en-us/powershell/azure/install-az-ps?view=azps-6.4.0&WT.mc_id=m365-58890-cxa#installation)

      ``` command
      Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -AllowClobber -Force 
      ```
- Install [.NET Core 3.1 SDK](https://dotnet.microsoft.com/download/dotnet/3.1)

OPTIONAL: If you want to run these applications locally or modify them, you may find these tools helpful:

- [Visual Studio 2019](https://visualstudio.microsoft.com/vs/)
   >**Note:** use **Visual Studio Installer** to install the following development toolsets:
  - ASP.NET and web development
  - Azure development
  - Office/SharePoint development
  - .NET cross-platform development

- Install [.NET Framework 4.8 Developer Pack](https://dotnet.microsoft.com/download/dotnet-framework/thank-you/net48-developer-pack-offline-installer)

#### Step 2:  Download the source code needed to be deployed

Go to [https://github.com/OfficeDev/office-add-in-saas-monetization-sample](https://github.com/OfficeDev/office-add-in-saas-monetization-sample).
Clone or download the project into your local machine.

#### Step 3:  Get everything ready to run ARM template

- In the project you just downloaded in Step 2, go to folder `office-add-in-saas-monetization-sample/Deployment_SaaS_Resources/`.
- Open the `ARMParameters.json` file and update the following parameters with values you choose:
    - webAppSiteName
    - webApiSiteName
    - resourceMockWebSiteName
    - domainName
    - directoryId (Directory (tenant) ID)
    - sqlAdministratorLogin
    - sqlAdministratorLoginPassword
    - sqlMockDatabaseName
    - sqlSampleDatabaseName
    
> Leave the rest of the configuration in file `ARMParameters.json` as is, this will be automatically filled in after scripts deploy the resources.
You need to make sure enter a unique name for each web app and web site in the parameter list shown below because the script will create many Azure web apps and sites and each one must have a unique name.  All of the parameters that correspond to web apps and sites in the following list end in **SiteName**.
For **domainName** and **directoryId**, please refer to this [article](https://docs.microsoft.com/en-us/partner-center/find-ids-and-domain-names?WT.mc_id=m365-58890-cxa#find-the-microsoft-azure-ad-tenant-id-and-primary-domain-name) to find your Microsoft Azure AD tenant ID and primary domain name.

    
- In a Powershell 7 window, change to the **.\Deployment_SaaS_Resources** directory.

- In the same window run `Connect-Graph -Scopes "Application.ReadWrite.All, Directory.AccessAsUser.All DelegatedPermissionGrant.ReadWrite.All Directory.ReadWrite.All"`

- Click **Accept**.

 ![Graph consent](../Assets/08-001.png)

Once accepted, the browser will redirect and show below message. You can now close the browser and continue with the PowerShell command line.

 ![Graph consent redirect](../Assets/08-001-1.png)

> What this step does is add `Microsoft Graph PowerShell` in Azure Active Directory under [Enterprise Applications](https://docs.microsoft.com/en-us/azure/active-directory/manage-apps/add-application-portal?WT.mc_id=m365-58890-cxa) with the necessary permissions so we can create the needed applications for this particular exercise using its commands.

- In the same window run `.\InstallApps.ps1`

> You might get a warning as shown below. And it depends on the execution policy settings in the machine. 

 ![execution policy](../Assets/08-001-2.png)

Let's set it to be `bypass` for now. But please read more on Execution policies [here](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.2&WT.mc_id=m365-58890-cxa).

Run below script:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
Now re-run `.\InstallApps.ps1`

The script should now run to create all three applications in Azure AD. At the end of the script, your command line should display below information.:

> Based on the subscription you are using, you may change the location where your azure resources are deployed. To change this, find the `DeployTemlate.ps1` file and search for variable `$location`.
By default it is `centralus` but you can change it to `eastus` which works on both **Visual Studio Enterprise Subscription** and **Microsoft Azure Enterprise Subscription**.

 ![app id secret](../Assets/08-002.png)

- Copy the values from the output and later you will need  these values to update the code and .env file for deploying Add-ins. These values will also be pre-populated in `ARMParameters.json`. Do not change this file.
- Notice how the `ARMParameters.json` file is now updated with the values of applications deployed.

#### Step 4:  Deploy the ARM template with PowerShell

Open PowerShell 7 and run the Powershell command `Connect-AzAccount`.

This will redirect you to login page. Once you confirm with the Global admin credentials you have been using all along in this exercise, you will be redirected to a page displaying below:

![Azure CLI consent redirect](../Assets/08-001-1.png)

You can now close the browser and continue with the PowerShell command line. You will see similar output in your command line, if everything is okay:

![AZ CLI](../Assets/08-003.png)

Run the script `.\DeployTemplate.ps1`. When prompted, enter the name of the resource group to create. 

![AZ CLI](../Assets/08-004-1.png)

Your resourses will start to get deployed one after the other and you'll see the output as shown below if everything is okay:

![AZ CLI](../Assets/08-004.png)

You'll get a message on the command line, that the ARM Template deployment was successfully as shown below:

![AZ CLI](../Assets/08-005.png)

To confirm the creation of all three azure ad apps, go to the `App registrations` in Azure AD in Azure portal. Use this [link](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps) to navigate to it.

Under **All applications**,  filter with Display name `Contoso Monetization`.
You should see three apps as shown in the screen below:

![AZ AD Apps](../Assets/08-006.png)

#### Step 5: Deploy server side code

Now let's deploy the server side code for these three applications.

- In the command line, change to the **.\MonetizationCodeSample** directory.

- Run the script `.\PublishSaaSApps.ps1`.

- When prompted, enter the same resource group name.

You will see the source code in your local machine getting built and packaged.

  ![Build Apps](../Assets/08-007.png)



> **Note:** You may see some warnings about file expiration, please ignore.

The final messages may look like this:

 ![publish Apps](../Assets/08-008.png)


#### Step 6: Update .env file with deployed resources.

Add below entries into .env files in your working folder where you've done Labs A01-A07. Add below two keys, and replace the values (webApiSiteName) and (webApiClientId) with the values from your `ARMParameters.json` file:
```
 SAAS_API=https://(webApiSiteName).azurewebsites.net/api/Subscriptions/CheckOrActivateLicense
 SAAS_SCOPES=api://(webApiClientId)/user_impersonation
```

Where the values for `webApiSiteName` and `webApiClientId` are copied from the file `ARMParameters.json`.

Try visiting the App Source simulator, which is at https://(webAppSiteName).azurewebsites.net; you should be able to log in using your tenant administrator account. Don't purchase a subscription yet, however!

### Exercise 2: Grant the Northwind Orders app permission to call the licensing service in Azure

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

### Exercise 3: Update the Northwind Orders app to call the licensing service in Azure

#### Step 1: Add a server side function to validate the user has a license

In your working folder, create a new file /server/validateLicenseService.js and paste in this code (or copy the file from [here](../../A08-Monetization/server/northwindLicenseService.js)).

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

// TODO: Securely cache the results of this function for the lifetime of the resulting token
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

To do this, the code uses the [On Behalf Of flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow?WT.mc_id=m365-58890-cxa) to exchange the incoming access token (targeted for the Northwind Orders app) for a new access token that is targeted for the Licensing service application.

#### Step 2: Add a server side API to validate the user's license

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

#### Step 3: Add client side pages to display a license error

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



### Exercise 4: Run the application

#### Step 1: Run the app in Teams without a license

Return to your application in Microsoft Teams; refresh the tab or browser if necessary. The UI will begin to render, and then it will detect the license failure and display an error page.

![Run application](../Assets/08-201-RunApp-1.png)

--- 
>NOTE: The sample application checks the license in JavaScript, which is convenient for this lab but it would be easy for someone to bypass the license check. In a real application you'd probably check the license on all accesses to your application web site.
---

#### Step 2: "Purchase" a subscription and set licensing policy

Teams Store are also listed in the Microsoft App Source portal; users can purchase you app in either location. For this lab you will use an App Source simulator which you installed earlier in this lab, but users can purchase apps directly from the Teams user interface when they're listed in the Teams app store.

Browse to https://(resourceMockWebSiteName).azurewebsites.net where (resourceMockWebSiteName) is the name you chose in Exercise 1 Step 3. This should display the App Source simulator. 

---
> NOTE: The App Source simulator's background color is green to make it easy to see when you are redirected to your app's landing page, which has a blue background.
---

Click the "Purchase" button to purchase a subscription to the Northwind Orders Application.

![Run application](../Assets/08-202-RunApp-2.png)

---
> NOTE: The App Source simulator has a mock offer name, "Contoso Apps", rather than showing the "Northwind Orders" app. This is just a constant defined in the monetization project's SaasOfferMockData/Offers.cs file. The real App Source web page will show the application name and other information you configured in Partner Center.
---

Next, the App Source simulator displays the plans available for the offer; the simulator has two hard-coded plans, "SeatBasedPlan" (which uses a [per-user pricing model](https://docs.microsoft.com/en-us/azure/marketplace/create-new-saas-offer-plans?WT.mc_id=m365-58890-cxa#define-a-pricing-model)), and a "SiteBasedPlan" (which uses a [flat-rate pricing model](https://docs.microsoft.com/en-us/azure/marketplace/plan-saas-offer?WT.mc_id=m365-58890-cxa#saas-pricing-models)). The real App Source would show the plans you had defined in Partner Center.

Since Microsoft Teams only supports the per-user pricing model, choose the "SiteBasedPlan" and click the "Purchase" button. Because this is a simulator, your credit card will not be charged.

![Run application](../Assets/08-203-RunApp-3.png)

The simulated purchase is now complete, so you will be redirected to the app's landing page. You will need to supply a page like this as part of your application; it needs to interpret a token sent by App Source and log the user in with Azure Active Directory. This token can be sent to the Microsoft Commercial Marketplace API, which will respond with the details about what the customer has purchased. You can find [the code for this](https://github.com/OfficeDev/office-add-in-saas-monetization-sample/blob/7673db6c8e6c809ae7aa0ba894460183aed964fc/MonetizationCodeSample/SaaSSampleWebApp/Services/SubscriptionService.cs#L37) in the Monetization repo's SaaSSampleWebApp project under /Services/SubscriptionService.cs.

The landing page gives the app a chance to interact with the user and capture any configuration information it needs. Users who purchase the app in the Teams store would be brought to this same page. The sample app's landing page allows the user to select a region; the app stores this information in its own database based on the Microsoft 365 tenant ID.

![Run application](../Assets/08-204-RunApp-4.png)

Once the region has been selected, the sample app shows a welcome page with the user's name, which is obtained by [reading the user's profile with the Microsoft Graph API](https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&WT.mc_id=m365-58890-cxa). Click "License Settings" to view the license assignment screen.

![Run application](../Assets/08-205-RunApp-5.png)

On this screen you can add individual user licenses using the "Add User" button, or you can set a policy that allows users to claim licenses on a first come, first served basis. Turn on the "First come first served" switch to enable this option.

![Run application](../Assets/08-206-RunApp-6.png)

Note that everything on this screen is defined by this application. It's intended to be flexible since our partners have a wide range of licensing approaches. Apps can tell who's logging in via Azure AD and use the user ID and tenant ID to authorize access, provide personalization, etc.

#### Step 3: Run the application in Teams

Now that you've purchased a subscription, return to Microsoft Teams and refresh your application. The license will be checked and the user can interact with the application normally.

![Run application](../Assets/08-208-RunApp-8.png)

Now return to the licensing application. If you've closed the tab, you can find it at https://(webAppSiteName).azurewebsites.net where (webAppSiteName) is the name you chose in Exercise 1 Step 3. 

![Run application](../Assets/08-209-RunApp-9.png)

Notice that your username has been assigned a license. The sample app stored this in a SQL Server database. When the Teams application called the licensing service, the access token contained the tenant ID and user ID, enabling the licening service to determine that the user has a license.

## ** CONGRATULATIONS **

You have completed the Teams App Camp! Thanks very much; we hope this helps in the process of extending your application to Microsoft Teams!

### Known issues

For the latest issues, or to file a bug report, see the [github issues list](https://github.com/OfficeDev/TeamsAppCamp1/issues) for this repository.

### References



