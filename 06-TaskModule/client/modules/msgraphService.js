import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
export async function getGraphUserDetails() {

    await new Promise((resolve, reject) => {
        microsoftTeams.initialize(() => { resolve(); });
    });

    const authToken = await new Promise((resolve, reject) => {
        microsoftTeams.authentication.getAuthToken({
            successCallback: (result) => { resolve(result); },
            failureCallback: (error) => { reject(error); }
        });
    });

    const response = await fetch(`/api/getGraphUserDetails`, {
        "method": "get",
        "headers": {
            "content-type": "application/json",
            "authorization": `Bearer ${authToken}`
        },
        "cache": "no-cache"
    });
   
    if (response.ok) {

        const data = await response.json();
        return data;

    } else {

        const error = await response.json();
        console.log(`ERROR: ${error}`);

    }

}
