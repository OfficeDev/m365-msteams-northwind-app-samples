import 'https://res.cdn.office.net/teams-js/2.0.0-beta.3/js/MicrosoftTeams.min.js';
// async function returns true if we're running in Teams, Outlook, Office

export async function inM365() { 
  const m365Apps=["Teams","Outlook","Office"];
  microsoftTeams.app.initialize();
  const context= await microsoftTeams.app.getContext();
  return m365Apps.includes(context.app.host.name);
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
          setHubTheme("/northwind.css");
          // When the theme changes, update the CSS again: Only for teams
          microsoftTeams.app.registerOnThemeChangeHandler((theme) => {
          setTheme(theme);
        });
        };        
        break;
        case "Outlook":{
          setHubTheme("/northwind-outlook.css");
        };
          break;
        case "Office":{
          setHubTheme("/northwind-office.css");
        }
        break;
        default:{
          setHubTheme("/northwind.css");
        }
      }     
    }

function setHubTheme(fileName) {
    let element = document.createElement("link");
    element.setAttribute("rel", "stylesheet");
    element.setAttribute("type", "text/css");
    element.setAttribute("href", fileName);
    document.getElementsByTagName("head")[0].appendChild(element);
  }
});
}
