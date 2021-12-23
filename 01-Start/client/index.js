import {
    getLoggedinEmployeeId,
    getEmployeeProfile,
    logoff
} from './modules/northwindIdentityService.js';
import {
    getOrdersForEmployee
} from './modules/northwindDataService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const imageElement = document.getElementById('image');
    const ordersElement = document.getElementById('orders');
    const logoutButton = document.getElementById('logout');
    const messageDiv = document.getElementById('message');

    try {
        logoutButton.addEventListener('click', async ev => {
            logoff();
            window.location.href = "/pages/northwindLogin.html";
        });
        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/pages/northwindLogin.html";
        } else {
            const employeeProfile = await getEmployeeProfile(employeeId);
            const orders = await getOrdersForEmployee(employeeId);

            displayElement.innerHTML = `
            <h1>Hello ${employeeProfile.displayName}</h1>
            <p>Mail: ${employeeProfile.mail}<br />
            Job Title: ${employeeProfile.jobTitle}<br />
        `;
            imageElement.src = `data:image/bmp;base64,${employeeProfile.photo}`;

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
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();