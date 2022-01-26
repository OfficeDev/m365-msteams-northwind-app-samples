import aad from 'azure-ad-jwt';
import * as msal from '@azure/msal-node';
import dotenv from 'dotenv';
import fetch from 'node-fetch';

import { validateEmployeeLogin } from './northwindIdentityService.js';
import e from 'express';

dotenv.config();

// validateAndMapAadLogin() - Returns an employee ID of the logged in user based
// on an existing mapping OR the username/password passed from a client login.
// If there is no existing mapping and no username/password is specified, it will throw
// an exception.
export async function validateAndMapAadLogin(req, res) {

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

const idMap = [];   // TODO: Make this go away
async function getEmployeeIdForUser(aadUserId, aadToken) {

    let employeeId;
    try {
        console.log(aadUserId);
        const msalResponse =
            await msalClientApp.acquireTokenByClientCredential(msalRequest);
        console.log(msalResponse.accessToken);

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
        } else {
            console.log(`Error ${graphResponse.status} calling Graph in getEmployeeIdForUser: ${graphResponse.statusText}`);
        }

    }
    catch (error) {
        console.log(`Error calling MSAL in getEmployeeIdForUser: ${error}`);
    }
    return employeeId; // idMap[aadUserId];
}

async function setEmployeeIdForUser(aadUserId, employeeId) {
    idMap[aadUserId] = employeeId;
}