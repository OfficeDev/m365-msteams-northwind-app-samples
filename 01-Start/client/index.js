import {
    getLoggedinEmployeeId,
    getEmployeeProfile
} from './modules/northwindIdentityService.js';
import {
    getOrdersForEmployee
} from './modules/northwindDataService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const imageElement = document.getElementById('image');
    const ordersElement = document.getElementById('orders');

    try {
        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/northwindLogin.html";
        } else {
            const employeeProfile = await getEmployeeProfile(employeeId);
            const orders = await getOrdersForEmployee(employeeId);

            displayElement.innerHTML = `
            <h1>Hello ${employeeProfile.displayName}</h1>
            <h3>Profile Information</h3>
            <p>Mail: ${employeeProfile.mail}<br />
            Job Title: ${employeeProfile.jobTitle}<br />
        `;
            imageElement.src = `data:image/bmp;base64,${employeeProfile.photo}`;

            orders.forEach(order => {
                const orderRow = document.createElement('tr');
                orderRow.innerHTML = `<tr>
                <td>${order.OrderID}</td>
                <td>${(new Date(order.OrderDate)).toDateString()}</td>
                <td>${order.ShipName}</td>
                <td>${order.ShipAddress}, ${order.ShipCity} ${order.ShipRegion || ''} ${order.ShipPostalCode || ''} ${order.ShipCountry}</td>
            </tr>`;
                ordersElement.append(orderRow);

            });
        }
    }
    catch (error) {            // If here, we had some other error
        displayElement.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();