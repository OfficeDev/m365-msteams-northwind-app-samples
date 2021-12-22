import {
   validateEmployeeLogin,
   setLoggedinEmployeeId
} from '../modules/northwindIdentityService.js';
import {
   getAllEmployees
} from '../modules/northwindDataService.js';

const usernameInput = document.getElementById('username');
const passwordInput = document.getElementById('password');
const loginButton = document.getElementById('loginButton');
const messageDiv = document.getElementById('message');

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
const employees = await getAllEmployees();

employees.forEach(employee => {
   const employeeListItem = document.createElement('li');
   employeeListItem.innerText = employee.lastName;
   hintUL.appendChild(employeeListItem);
});
