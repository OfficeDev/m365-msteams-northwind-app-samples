import 'https://res.cdn.office.net/teams-js/2.0.0/js/MicrosoftTeams.min.js';

// async function returns true if we're running in Teams
export async function inTeams() {
    try {
        await microsoftTeams.app.initialize();
        const context = await microsoftTeams.app.getContext();
        return (context.app.host.name === microsoftTeams.HostName.teams);
    }
    catch (e) {
        console.error(`ERROR in inTeams: ${e}`);
    }
}

// Set the CSS to reflect the current Teams theme
function setTheme(theme) {
    const el = document.documentElement;
    el.setAttribute('data-theme', theme); // switching CSS
};

// Inline code to set theme on any page using teamsHelpers
await microsoftTeams.app.initialize();
const context = await microsoftTeams.app.getContext();
setTheme(context.app.theme);

// When the theme changes, update the CSS again
microsoftTeams.registerOnThemeChangeHandler((theme) => {
    setTheme(theme);
});
