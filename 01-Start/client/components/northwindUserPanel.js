import {
    getLoggedinEmployeeId,
    getEmployeeProfile,
    logoff
} from '../modules/northwindIdentityService.js';


class northwindUserPanel extends HTMLElement {

    async connectedCallback() {

        const employeeId = await getLoggedinEmployeeId();

        if (!employeeId) {

            // If here, nobody is logged in - redirect to the login page
            window.location.href = "/pages/northwindLogin.html";

        } else {

            // If here, get the profile of the logged-in user and display
            // the user panel
            const employee = await getEmployeeProfile(employeeId);
            this.innerHTML = `<div class="userPanel">
                <img src="data:image/bmp;base64,${employee.photo}"></img>
                <p>${employee.displayName}</p>
                <p>${employee.jobTitle}</p>
                <hr />
                <button id="logout">Log out</button>
            </div>
            `;

        }

        const logoutButton = document.getElementById('logout');
        logoutButton.addEventListener('click', async ev => {
            logoff();
            window.location.href = "/pages/northwindLogin.html";
        });
        
    }
}

// Define the web component and insert an instance at the top of the page
customElements.define('northwind-user-panel', northwindUserPanel);
const panel = document.createElement('northwind-user-panel');
document.body.insertBefore(panel, document.body.firstChild);
