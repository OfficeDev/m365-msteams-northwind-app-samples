import {
    getLoggedinEmployeeId,
    getEmployeeProfile
} from '../northwindIdentity/identityService.js';
import {
    getOrdersForEmployee
} from '../modules/northwindDataService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const imageElement = document.getElementById('image');
    const ordersElement = document.getElementById('orders');
    const messageDiv = document.getElementById('message');

    try {
        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/pages/northwindLogin.html";
        } else {
            const employeeProfile = await getEmployeeProfile(employeeId);
            displayElement.innerHTML = `
                <h3>Orders for ${employeeProfile.displayName}<h3>
            `;

            const orders = await getOrdersForEmployee(employeeId);
            orders.forEach(order => {
                const orderRow = document.createElement('tr');
                orderRow.innerHTML = `<tr>
                <td><a href="/pages/orderDetail.html?orderId=${order.OrderID}">${order.OrderID}</a></td>
                <td>${(new Date(order.OrderDate)).toDateString()}</td>
                <td>${order.ShipName}</td>
                <td>${order.ShipAddress}, ${order.ShipCity} ${order.ShipRegion || ''} ${order.ShipPostalCode || ''} ${order.ShipCountry}</td>
            </tr>`;
                ordersElement.append(orderRow);

            });
        }
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();