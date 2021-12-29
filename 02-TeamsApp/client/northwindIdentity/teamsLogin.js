import {
   validateEmployeeLogin,
   setLoggedinEmployeeId
} from './identityService.js';
import {
   getEmployees
} from '../northwindData/dataService.js';
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';

const loginPanel = document.getElementById('loginPanel');
const usernameInput = document.getElementById('username');
const passwordInput = document.getElementById('password');
const loginButton = document.getElementById('loginButton');
const messageDiv = document.getElementById('message');
const hintUL = document.getElementById('hintList');
const teamsLoginLauncher = document.getElementById('teamsLoginLauncher');
const teamsLoginLauncherButton = document.getElementById('teamsLoginLauncherButton');

microsoftTeams.initialize(() => {

   if (window.location !== window.parent.location) {

      // The page is in an Iframe - assume we're in a Teams IFrame

      teamsLoginLauncher.style.display = 'inline';
      teamsLoginLauncherButton.addEventListener('click', async ev => {
         microsoftTeams.authentication.authenticate({
            url: `${window.location.origin}/northwindIdentity/teamsLogin.html`,
            width: 600,
            height: 535,
            successCallback: (url) => {
               window.location.href = url;
            },
            failureCallback: (reason) => {
               throw `Error in teams.authentication.authenticate: ${reason}`
            }
         });
      });

   } else {

      // If here we're not in an IFrame - assume it's a Teams popup

      loginPanel.style.display = 'inline';
      loginButton.addEventListener('click', async ev => {

         messageDiv.innerText = "";
         const employeeId = await validateEmployeeLogin(
            usernameInput.value,
            passwordInput.value
         );
         if (employeeId) {
            setLoggedinEmployeeId(employeeId);
            microsoftTeams.authentication.notifySuccess("/");
         } else {
            microsoftTeams.authentication.notifyFailure("User not found");
         }
      });

      (async () => {
         const employees = await getEmployees();
         employees.forEach(employee => {
            const employeeListItem = document.createElement('li');
            employeeListItem.innerHTML = `<b>${employee.lastName.toLowerCase()}</b> (${employee.firstName} ${employee.lastName})`;
            hintUL.appendChild(employeeListItem);
         });
      })();

   }

});