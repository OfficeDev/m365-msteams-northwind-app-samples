import 'https://res-sdf.cdn.office.net/teams-js/2.0.0-beta.4/js/MicrosoftTeams.min.js';

// async function returns true if we're running in Teams, Outlook, Office
export async function inM365() {     
  await microsoftTeams.app.initialize();
  const context= await microsoftTeams.app.getContext(); 
  //check against the enum for hostnames
  return Object.values(microsoftTeams.HostName).includes(context.app.host.name);
}
 const setTheme = (theme)=>{
    const el = document.documentElement;
    el.setAttribute('data-theme', theme); // switching CSS
};
const displayTheme=async()=>{
  if(inM365()) {
      const context= await microsoftTeams.app.getContext();  
      if(context) {
        setTheme(context.theme);     
        switch(context.app.host.name){
          case microsoftTeams.HostName.teams:{
            setHubTheme("/northwind.css");
            // When the theme changes, update the CSS again: Only for teams
            microsoftTeams.app.registerOnThemeChangeHandler((theme) => {
            setTheme(theme);
          });
          };        
          break;
          case microsoftTeams.HostName.outlook:{
            setHubTheme("/northwind-outlook.css");
          };
            break;
          case microsoftTeams.HostName.office:{
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
  }
  

}
displayTheme();

