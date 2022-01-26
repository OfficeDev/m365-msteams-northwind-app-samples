import aad from 'azure-ad-jwt';
import fetch from 'node-fetch';


export async function getGraphUserDetails(teamsAccessToken) {

    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    return new Promise((resolve, reject) => {

        aad.verify(teamsAccessToken, { audience: audience }, async (err, result) => {
            if (result) {
                const graphAppAccessToken = await getGraphOboAccessToken(teamsAccessToken);
               
                const graphAppUrl = `https://graph.microsoft.com/beta/users/?$filter=employeeId eq '1'`

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
                    resolve(graphData.value[0]);
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