import dotenv from 'dotenv';
import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE } from './constants.js';
import aad from 'azure-ad-jwt';
import * as msal from '@azure/msal-node';

dotenv.config();

// Mock identity service based on Northwind employees

// Wire up middleware
export async function initializeIdentityService(app) {

    // Web service validates a user login
    app.post('/api/validateEmployeeLogin', async (req, res) => {

        try {
            const username = req.body.username;
            const password = req.body.password;
            const employeeId = await validateEmployeeLogin(username, password);
            res.send(JSON.stringify({ "employeeId": employeeId }));
        }
        catch (error) {
            console.log(`Error in /api/validateEmployeeLogin handling: ${error}`);
            res.status(500).json({ status: 500, statusText: error });
        }

    });

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
            console.log(`Error in /api/validateAadLogin handling: ${error.statusMessage}`);
            res.status(error.status).json({ status: error.status, statusText: error.statusMessage });
        }

    });

    // Middleware to validate all other API requests
    app.use('/api/', validateApiRequest);

}

async function validateApiRequest(req, res, next) {
    try {
        if (req.cookies.employeeId && parseInt(req.cookies.employeeId) > 0) {
            console.log(`Validated authentication on /api${req.path}`);
            next();
        } else {
            console.log(`Invalid authentication on /api${req.path}`);
            res.status(401).json({ status: 401, statusText: "Access denied" });
        }
    }
    catch (error) {
        res.status(401).json({ status: 401, statusText: error });
    }
}

async function validateEmployeeLogin(username, password) {

    // For simplicity, the username is the employee's surname,
    // and the password is ignored
    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Employees?$filter=LastName eq '${username}'&$select=EmployeeID`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const employees = await response.json();
    if (employees?.value?.length === 1) {
        return employees.value[0].EmployeeID;
    } else {
        return null;
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