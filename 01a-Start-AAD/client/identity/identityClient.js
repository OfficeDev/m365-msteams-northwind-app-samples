// User profiles are stored in the Northwind database
import { getEmployee } from '../modules/northwindDataService.js';
import 'https://alcdn.msauth.net/browser/2.21.0/js/msal-browser.min.js';
import { env } from '/modules/env.js';

// interface IIdentityClient {
//     async getLoggedinEmployeeId(): number;
//     async setLoggedinEmployeeId(employeeId: number): void;
//     async validateEmployeeLogin(surname: string, password: string): Number;
//     async getLoggedInEmployee(): IEmployee;
//     async Logoff(): void                    // Redirects and does not return
//     async getFetchHeadersAnon(): string;    // headers for calling Fetch anonymously
//     async getFetchHeadersAuth(): string;    // headers for calling Fetch with authentication
// }

const msalConfig = {
    auth: {
        clientId: env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${env.TENANT_ID}`,
        redirectUri: `https://${env.HOSTNAME}`
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false    // Set this to "true" if you are having issues on IE11 or Edge
    }
};

// MSAL request object to use over and over
const msalRequest = {
    scopes: [`api://${env.HOSTNAME}/${env.CLIENT_ID}/access_as_user`]
}

const msalClient = new msal.PublicClientApplication(msalConfig);

export async function getLoggedinEmployeeId() {

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
        console.log (accessToken);
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            console.warn("Silent token acquisition failed; acquiring token using redirect");
            this.msalClient.acquireTokenRedirect(this.request);
        } else {
            throw (error);
        }
    }

    const response = await fetch(`/api/validateAadLogin`, {
        "method": "post",
        "headers": {
            "content-type": "application/json",
            "authorization": `Bearer ${accessToken}`
        },
        "body": JSON.stringify({
            "employeeId": 0
        }),
        "cache": "no-cache"
    });
    if (response.ok) {
        const data = await response.json();
        if (data.employeeId) {
            console.log(`Employee ID ${data.employeeId}`);
            // If here, AAD user was mapped to a Northwind employee ID
            // setLoggedinEmployeeId(data.employeeId);
            // window.location.href = document.referrer;
        }
    }

    // const cookies = document.cookie.split(';');
    // for (const c of cookies) {
    //     const [name, value] = c.split('=');
    //     if (name.trim() === 'employeeId') {
    //         return Number(value);
    //     }
    // }
    // return null;
}

export async function setLoggedinEmployeeId(employeeId) {
    document.cookie = `employeeId=${employeeId};SameSite=None;Secure;path=/`;
}

export async function validateEmployeeLogin(surname, password) {

    const response = await fetch(`/api/validateEmployeeLogin`, {
        "method": "post",
        "headers": await getFetchHeadersAnon(),
        "body": JSON.stringify({
            "username": surname,
            "password": password
        }),
        "cache": "no-cache"
    });
    if (response.ok) {
        const data = await response.json();
        return data.employeeId;
    } else {
        const error = await response.json();
        console.log(`ERROR: ${error}`);
        throw (error);
    }
}

// Get the employee profile from our web service
export async function getLoggedInEmployee() {

    const employeeId = await getLoggedinEmployeeId();

    return await getEmployee(employeeId);

}

export async function logoff() {
    setLoggedinEmployeeId(0);
    // Redirect to the login page
    window.location.href = "/identity/login.html";
}

// Headers for use in Fetch (HTTP) requests when calling anonymous web services
// in the server side of this app.
export async function getFetchHeadersAnon() {
    return ({ "content-type": "application/json" });
}

// Headers for use in Fetch (HTTP) requests when calling authenticated web services
// in the server side of this app. Authentication is sent in a cookie, so no
// additional headers are required.
// Other implementations of this module may insert an Authorization header here
export async function getFetchHeadersAuth() {
    return ({ "content-type": "application/json" });
}
