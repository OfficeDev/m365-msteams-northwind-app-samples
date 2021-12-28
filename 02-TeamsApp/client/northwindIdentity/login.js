import {
   validateEmployeeLogin,
   setLoggedinEmployeeId
} from './identityService.js';
import {
   getEmployees
} from '../northwindData/dataService.js';

const loginPanel = document.getElementById('loginPanel');
const usernameInput = document.getElementById('username');
const passwordInput = document.getElementById('password');
const loginButton = document.getElementById('loginButton');
const messageDiv = document.getElementById('message');

if (window.location !== window.parent.location) {
   // The page is in an iframe - refuse service
   messageDiv.innerText = "ERROR: You cannot run this app in an IFrame";
} else {

   loginPanel.style.display = 'inline';
   loginButton.addEventListener('click', async ev => {

      messageDiv.innerText = "";
      const employeeId = await validateEmployeeLogin(
         usernameInput.value,
         passwordInput.value
      );
      if (employeeId) {
         setLoggedinEmployeeId(employeeId);
         window.location.href = "/";
      } else {
         messageDiv.innerText = "Error: user not found";
      }
   });

   const hintUL = document.getElementById('hintList');
   const employees = await getEmployees();

   employees.forEach(employee => {
      const employeeListItem = document.createElement('li');
      employeeListItem.innerHTML = `<b>${employee.lastName.toLowerCase()}</b> (${employee.firstName} ${employee.lastName})`;
      hintUL.appendChild(employeeListItem);
   });

}