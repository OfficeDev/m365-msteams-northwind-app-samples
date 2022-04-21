import dotenv from 'dotenv';
import fetch from 'node-fetch';
import aad from 'azure-ad-jwt';
import * as msal from '@azure/msal-node';

dotenv.config();

// Wire up middleware
export async function initializeIdentityService(app) {

    // Web service validates an Azure AD login
    app.post('/api/validateAadLogin', async (req, res) => {

        try {
            const employeeId = await validateAndMapAadLogin(req, res);
            if (employeeId) {
                res.send(JSON.stringify({ "employeeId": employeeId }));
            } else {
                res.status(401).send('Unknown authentication failure');
            }
        }
        catch (error) {
            console.log(`Error in /api/validateAadLogin handling: ${error.message}`);
            res.status(error.status).json({ status: error.status, statusText: error.statusMessage });
        }

    });

    // Middleware to validate all other API requests
    app.use('/api/', validateApiRequest);

}

async function validateApiRequest(req, res, next) {
    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    const token = req.headers['authorization'].split(' ')[1];
    if (req.path === "/messages") {
        console.log('Request for bot, validation will be performed by Bot Framework Adapter');
        next();
    } else {
        aad.verify(token, { audience: audience }, async (err, result) => {
            if (result) {
                console.log(`Validated authentication on /api${req.path}`);
                next();
            } else {
                console.error(`Invalid authentication on /api${req.path}: ${err.message}`);
                res.status(401).json({ status: 401, statusText: "Access denied" });
            }
        });
    }
}

// validateAndMapAadLogin() - Returns an employee ID of the logged in user based
// on an existing mapping OR the username/password passed from a client login.
// If there is no existing mapping and no username/password is specified, it will throw
// an exception.
async function validateAndMapAadLogin(req, res) {

    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    const token = req.headers['authorization'].split(' ')[1];

    const aadUserId = await new Promise((resolve, reject) => {
        aad.verify(token, { audience: audience }, async (err, result) => {
            if (result) {
                resolve(result.oid);
            } else {
                console.error(`Error validating access token: ${err.message}`);
                reject(err);
            }
        });
    });

    if (aadUserId) {
        // If here, user is logged into Azure AD
        let employeeId = await getEmployeeIdForUser(aadUserId);
        if (employeeId) {
            // We found the employee ID for the AAD user
            return employeeId;
        } else if (req.body.username) {
            // We did not find an employee ID for this user, try to 
            // get one using the legacy authentication
            const username = req.body.username;
            const password = req.body.password;
            const employeeId = await validateEmployeeLogin(username, password);
            if (employeeId) {
                // If here, user is logged into both Azure AD and the legacy
                // authentication. Save the employee ID in the user's AAD
                // profile for future use.
                await setEmployeeIdForUser(aadUserId, employeeId);
                return employeeId;
            } else {
                // If here, the employee login failed; throw an exception
                throw ({ status: 401, statusMessage: "Employee login failed" });
            }
        } else {
            // If here we don't have an employee ID and employee credentials were
            // not provided.
            throw ({ status: 404, statusMessage: "Employee ID not found for this user" });
        }
    } else {
        res.status(401).send('Invalid AAD token');
    }
}

const config = {
    auth: {
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET
    }
};
const msalClientApp = new msal.ConfidentialClientApplication(config);
const msalRequest = {
    scopes: ["https://graph.microsoft.com/.default"]
}

const idCache = {};     // The employee mapping shouldn't change over time, so cache it here
async function getEmployeeIdForUser(aadUserId) {

    let employeeId;
    if (idCache[aadUserId]) {
        employeeId = idCache[aadUserId];
    } else {
        try {
            const msalResponse =
                await msalClientApp.acquireTokenByClientCredential(msalRequest);

            const graphResponse = await fetch(
                `https://graph.microsoft.com/v1.0/users/${aadUserId}?$select=employeeId`,
                {
                    "method": "GET",
                    "headers": {
                        "Accept": "application/json",
                        "Content-Type": "application/json",
                        "Authorization": `Bearer ${msalResponse.accessToken}`
                    }
                });
            if (graphResponse.ok) {
                const employeeProfile = await graphResponse.json();
                employeeId = employeeProfile.employeeId;
                idCache[aadUserId] = employeeId;
            } else {
                console.log(`Error ${graphResponse.status} calling Graph in getEmployeeIdForUser: ${graphResponse.statusText}`);
            }
        }
        catch (error) {
            console.log(`Error calling MSAL in getEmployeeIdForUser: ${error}`);
        }
    }
    return employeeId;
}

async function setEmployeeIdForUser(aadUserId, employeeId) {
    try {
        const msalResponse =
            await msalClientApp.acquireTokenByClientCredential(msalRequest);

        const graphResponse = await fetch(
            `https://graph.microsoft.com/v1.0/users/${aadUserId}`,
            {
                "method": "PATCH",
                "headers": {
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${msalResponse.accessToken}`
                },
                "body": JSON.stringify({
                    "employeeId": employeeId.toString()
                })
            });
        if (graphResponse.ok) {
            const employeeProfile = await graphResponse.json();
            employeeId = employeeProfile.employeeId;
        } else {
            console.log(`Error ${graphResponse.status} calling Graph in setEmployeeIdForUser: ${graphResponse.statusText}`);
        }

    }
    catch (error) {
        console.log(`Error calling MSAL in getEmployeeIdForUser: ${error}`);
    }
    return employeeId;
}

// TODO: Securely cache the results of this function for the lifetime of the resulting token
export async function getOboAccessToken(clientSideToken, scopes) {

    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

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