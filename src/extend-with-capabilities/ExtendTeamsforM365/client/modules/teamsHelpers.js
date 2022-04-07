import 'https://res.cdn.office.net/teams-js/2.0.0-beta.1/js/MicrosoftTeams.min.js';
// async function returns true if we're running in Teams
export async function inTeams() {
    // There is no supported way to do this in the Teams JavaScript SDK v1.
    // The official guidance is to indicate the app is running in Teams via
    // the URL. This workaround is used by the "yo teams" generator so is in
    // wide use. See the "yo teams" base class here:
    // https://github.com/wictorwilen/msteams-react-base-component/blob/master/src/useTeams.ts#L10
    return (window.parent === window.self && window.nativeInterface) ||
        window.navigator.userAgent.includes("Teams/") ||
        window.name === "embedded-page-container" ||
        window.name === "extension-tab-frame";
}

function setTheme (theme) {
    const el = document.documentElement;
    el.setAttribute('data-theme', theme); // switching CSS
};
if(microsoftTeams.app !== undefined) {
  microsoftTeams.app.initialize();
  microsoftTeams.app.getContext().then(context=> { 
    if(context) {
      setTheme(context.theme);     
      switch(context.app.host.name){
        case "Teams":{
          let element = document.createElement("link");
          element.setAttribute("rel", "stylesheet");
          element.setAttribute("type", "text/css");
          element.setAttribute("href", "/northwind.css");
          document.getElementsByTagName("head")[0].appendChild(element);
          // When the theme changes, update the CSS again: Only for teams
          microsoftTeams.app.registerOnThemeChangeHandler((theme) => {
          setTheme(theme);
        });
        };
        
        break;
        case "Outlook":{
          let element = document.createElement("link");
          element.setAttribute("rel", "stylesheet");
          element.setAttribute("type", "text/css");
          element.setAttribute("href", "/northwind-outlook.css");
          document.getElementsByTagName("head")[0].appendChild(element);
        };
          break;
        case "Office":{
          let element = document.createElement("link");
          element.setAttribute("rel", "stylesheet");
          element.setAttribute("type", "text/css");
          element.setAttribute("href", "/northwind-office.css");
          document.getElementsByTagName("head")[0].appendChild(element);
        }
        break;
        default:{
          let element = document.createElement("link");
          element.setAttribute("rel", "stylesheet");
          element.setAttribute("type", "text/css");
          element.setAttribute("href", "/northwind.css");
          document.getElementsByTagName("head")[0].appendChild(element);
        }
      }     
    }});
}
