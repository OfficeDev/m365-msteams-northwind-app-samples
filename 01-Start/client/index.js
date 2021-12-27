import {
    getLoggedinEmployeeId,
    getEmployeeProfile
} from './northwindIdentity/identityService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const messageDiv = document.getElementById('message');

    try {
        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/pages/northwindLogin.html";
        } else {
            const employeeProfile = await getEmployeeProfile(employeeId);
            displayElement.innerHTML = `
                <h3>Welcome ${employeeProfile.displayName}</h3>
                <ul>
                    <li><a href="/pages/myOrders.html">View my orders</a></li>
                    <li><a href="/pages/categories.html">Browse products</a></li>
                </ul>
            `;
        }
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();