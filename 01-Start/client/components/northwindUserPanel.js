import {
    getLoggedinEmployeeId,
    getEmployeeProfile,
    logoff
} from '../modules/northwindIdentityService.js';


class northwindUserPanel extends HTMLElement {

    async connectedCallback() {

        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/pages/northwindLogin.html";
        }

        this.innerHTML = `<div class="userPanel">
            <p>Employee ${employeeId}</p>
            <button id="logout">Log out</button>
        </div>
        `;
        const logoutButton = document.getElementById('logout');
        logoutButton.addEventListener('click', async ev => {
            logoff();
            window.location.href = "/pages/northwindLogin.html";
        });
        
    }
}

customElements.define('northwind-user-panel', northwindUserPanel);
