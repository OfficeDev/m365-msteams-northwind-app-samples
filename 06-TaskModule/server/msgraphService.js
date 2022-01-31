import aad from 'azure-ad-jwt';
import fetch from 'node-fetch';

const userCache = {}
export async function getGraphUserDetails(teamsAccessToken,userId) {
    if (userCache[userId]) return userCache[userId];
    const graphResult = {};
    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    return new Promise((resolve, reject) => {

        aad.verify(teamsAccessToken, { audience: audience }, async (err, result) => {
            if (result) {
                const graphAppAccessToken = await getGraphOboAccessToken(teamsAccessToken);
               
                const graphAppUrl = `https://graph.microsoft.com/v1.0/users/?$filter=employeeId eq '${userId}'`

                const graphResponse = await fetch(graphAppUrl, {
                    method: "GET",
                    headers: {
                        "Accept": "application/json",
                        "Content-Type": "application/json",
                        "Authorization" :`Bearer ${graphAppAccessToken}`
                    }
                });
                if (graphResponse.ok) {
                    const graphData = await graphResponse.json();  
                    graphResult.mail=graphData.value[0].mail;
                    graphResult.displayName=graphData.value[0].displayName; 
                    const graphAppUrl2 = `https://graph.microsoft.com/v1.0/users/${graphData.value[0].mail}/manager`

                    const graphResponse2 = await fetch(graphAppUrl2, {
                        method: "GET",
                        headers: {
                            "Accept": "application/json",
                            "Content-Type": "application/json",
                            "Authorization" :`Bearer ${graphAppAccessToken}`
                        }
                    });
                    if (graphResponse.ok) {
                    const managerInfo = await graphResponse2.json(); 
                    graphResult.managerMail=managerInfo.mail;
                    graphResult.managerDisplayName=managerInfo.displayName;   
                    }else {
                        reject(`Error in graph manager of user obo`);
                    }
                    userCache[userId] = graphResult;              
                    resolve(graphResult);
                } else {
                    reject(`Error in graph obo`);
                }



            } else {
                reject("Error in getGraphUserDetails(): Invalid client access token");
            }
        });
    });

}


// TODO: Combine with Graph exercise which will need OBO flow
// TODO: Add caching?
async function getGraphOboAccessToken(clientSideToken) {

    //const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const graphScopes = process.env.GRAPH_SCOPE;

    // Use On Behalf Of flow to exchange the client-side token for an
    // access token with the needed permissions
    
    const INTERACTION_REQUIRED_STATUS_TEXT = "interaction_required";
    const url = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    const params = {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: clientSideToken,
        requested_token_use: "on_behalf_of",
        scope: graphScopes
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