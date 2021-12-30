import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
import {
    setLoggedinEmployeeId
} from './identityService.js';

microsoftTeams.initialize(async () => {

    const authToken = await new Promise((resolve, reject) => {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (result) => { resolve (result); },
                failureCallback: (error) => { reject (error); }
            });
    });

    const response = await fetch (`/api/validateAadLogin`, {
        "method": "post",
        "headers": {
            "content-type": "application/json",
            "authorization": `Bearer ${authToken}`
        },
        "body": JSON.stringify({
            "authToken": authToken
        }),
        "cache": "no-cache"
    });
    if (response.ok) {
        const data = await response.json();
        setLoggedinEmployeeId(data.employeeId);
        window.location.href = document.referrer;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
    }
});



//    teamsLoginLauncherButton.addEventListener('click', async ev => {
//       microsoftTeams.authentication.authenticate({
//          url: `${window.location.origin}/northwindIdentity/login.html`,
//          width: 600,
//          height: 535,
//          successCallback: (response) => {
//             window.location.href = document.referrer;
//          },
//          failureCallback: (reason) => {
//             throw `Error in teams.authentication.authenticate: ${reason}`
//          }
//       });
//    });
   
// });
