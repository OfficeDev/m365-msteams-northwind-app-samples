// The Teams JavaScript SDK currently does not support a way to determine if code
// is running in or out of Teams. The recommended practice is to pass this context
// in the app URL - for example with a route (/pages/teamsTab.html) or a query string
// such as /home.html?teams=true.
//
// This app will use a query string parameter, teams=true.
// Since this is not a single-page app, it uses cookies to remember if it's running
// in or out of Teams. This works so long as the function is called on the first page,
// where it can check the query string parameter.

const COOKIE_NAME = 'inTeams';
export async function inTeams() {

    const cookies = document.cookie.split(';');
    for (const c of cookies) {
        const [name, value] = c.split('=');
        if (name === COOKIE_NAME) {
            // We found the cookie, convert it to boolean and return
            return (value.toLowerCase() === 'true');
        }
    }
    // If we found no cookie, set one now based on the presence or absence
    // of the query string parameter
    const urlParams = new URLSearchParams(window.location.search);
    const inTeams = urlParams.get('teams') === 'true' ? 'true' : 'false';
    document.cookie = `${COOKIE_NAME}=${inTeams};SameSite=None;Secure;path=/`;
    return inTeams === 'true';
    
}
