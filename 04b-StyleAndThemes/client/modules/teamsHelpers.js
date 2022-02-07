// The Teams JavaScript SDK currently does not support a way to determine if code
// is running in or out of Teams. 

// *** OPTION 1 ***

// This method is not officially suppoprted but seems to work.
// We are using it in this lab to allow students to freely switch between the stand-alone
// app and the app running in Teams in a browser. For the supported approach, see Option 2
// below.

import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';

// Code to attempt to retrieve the Teams context, but unlike the Teams SDK
// it will return null if you're running outside of Teams. This makes it easier
// to share code between Teams and browser apps.
const cache = {};
export async function getTeamsContext() {
    if (cache?.teamsContext) return cache.teamsContext;
    const teamsContext = await Promise.race([
        // If we're in Teams, the context will be returned
        new Promise((resolve) => {
            microsoftTeams.initialize(() => {
                microsoftTeams.getContext((context) => {
                    resolve(context);
                });
            });    
        }),
        // If we're not in Teams, the Teams SDK calls never return.
        // This timeout will resolve the promise and return null.
        new Promise((resolve) => {
            setTimeout(() => resolve(null), 100);
        })
    ]); 
    cache.teamsContext = teamsContext;
    return teamsContext;
}
export async function inTeams() {
    const teamsContext = await getTeamsContext();
    return teamsContext !== null;
}

function setTheme (theme) {
    const el = document.documentElement;
    el.setAttribute('data-theme', theme); // switching CSS
};
  
microsoftTeams.initialize(() => {
    // Set the CSS to reflect the current Teams theme
    microsoftTeams.getContext((context) => {
        setTheme(context.theme);
    });
    // When the theme changes, update the CSS again
    microsoftTeams.registerOnThemeChangeHandler((theme) => {
        setTheme(theme);
    });
});
  
  

// *** OPTION 2 ***
// The recommended practice is to indicate if the app is running in Teams by passing
// this in the app URL - for example with a route (/pages/teamsTab.html) or a query string
// such as /home.html?teams=true.
//
// This app will use a query string parameter, teams=true.
// Since this is not a single-page app, it uses cookies to remember if it's running
// in or out of Teams. This works so long as the function is called on the first page,
// where it can check the query string parameter.

// const COOKIE_NAME = 'inTeams';
// export async function inTeams() {

//     const cookies = document.cookie.split(';');
//     for (const c of cookies) {
//         const [name, value] = c.split('=');
//         if (name === COOKIE_NAME) {
//             // We found the cookie, convert it to boolean and return
//             return (value.toLowerCase() === 'true');
//         }
//     }
//     // If we found no cookie, set one now based on the presence or absence
//     // of the query string parameter
//     const urlParams = new URLSearchParams(window.location.search);
//     const inTeams = urlParams.get('teams') === 'true' ? 'true' : 'false';
//     document.cookie = `${COOKIE_NAME}=${inTeams};SameSite=None;Secure;path=/`;
//     return inTeams === 'true';
    
// }
