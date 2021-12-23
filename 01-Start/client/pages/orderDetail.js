import {
    getLoggedinEmployeeId,
    getEmployeeProfile,
    logoff
} from '../modules/northwindIdentityService.js';
import {
    getOrder
} from '../modules/northwindDataService.js';
// import { response } from 'express';

async function displayUI() {

    const displayElement = document.getElementById('content');
    // const imageElement = document.getElementById('image');
    const detailsElement = document.getElementById('orderDetails');
    const logoutButton = document.getElementById('logout');
    // const messageDiv = document.getElementById('message');

    try {
        logoutButton.addEventListener('click', async ev => {
            logoff();
            window.location.href = "/pages/northwindLogin.html";
        });
        const employeeId = await getLoggedinEmployeeId();
        if (!employeeId) {
            window.location.href = "/pages/northwindLogin.html";
        } else {
            // const employeeProfile = await getEmployeeProfile(employeeId);

            const searchParams = new URLSearchParams(window.location.search);
            if (searchParams.has('orderId')) {
                const orderId = searchParams.get('orderId');

                const order = await getOrder(orderId);

                displayElement.innerHTML = `
                    <h1>Order ${order.orderId}</h1>
                    <p>Customer: ${order.customerName}<br />
                    Contact: ${order.contactName}, (${order.contactTitle})<br />
                    Date: ${new Date(order.orderDate).toDateString()}</p>
                `;

                order.details.forEach(item => {
                    const orderRow = document.createElement('tr');
                    const pictureUrl = `data:image/bmp;base64,${item.categoryPicture}`
                    orderRow.innerHTML = `<tr>
                        <td>${item.quantity}</td>
                        <td><img src="${pictureUrl}" class="productImage"></image>
                            ${item.productName}</td>
                        <td>${item.unitPrice}</td>
                        <td>${item.discount}</td>
                    </tr>`;
                    detailsElement.append(orderRow);

                });

            }
        }
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();