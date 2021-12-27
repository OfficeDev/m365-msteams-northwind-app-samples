import {
    getLoggedInEmployee
} from './northwindIdentity/identityService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const messageDiv = document.getElementById('message');

    try {
        const employee = await getLoggedInEmployee();
        if (employee) {
            displayElement.innerHTML = `
                <h3>Welcome ${employee.displayName}</h3>
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