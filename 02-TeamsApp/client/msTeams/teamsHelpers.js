import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';

// Code to attempt to retrieve the Teams context, but unlike the Teams SDK
// it will return null if you're running outside of Teams. This makes it easier
// to share code between Teams and browser apps.
const cache = {};
export async function getTeamsContext() {

    if (cache?.teamsContext) return cache.teamsContext;

    const teamsContext = await Promise.race([

        // If we're in Teams, the context will be returned
        new Promise((resolve, reject) => {
            microsoftTeams.initialize(() => {
                microsoftTeams.getContext((context) => {
                    resolve(context);
                });
            });    
        }),

        // If we're not in Teams, the context never returns.
        // This timeout will resolve the promise and return null.
        new Promise((resolve, reject) => {
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