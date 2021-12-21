import { getUserProfile } from './modules/userProfileService.js';
import { getOrdersForEmployee } from './modules/northwindService.js'; 

async function displayUI() {

    const displayElement = document.getElementById('content');
    const imageElement = document.getElementById('image');
    const ordersElement = document.getElementById('orders');

    try {
        const userProfile = await getUserProfile();
        const orders = await getOrdersForEmployee(2);

        displayElement.innerHTML = `
            <h1>Hello ${userProfile.displayName}</h1>
            <h3>Profile Information</h3>
            <p>Mail: ${userProfile.mail}<br />
            Job Title: ${userProfile.jobTitle}<br />
        `;
        imageElement.src = `data:image/bmp;base64,${userProfile.photo}`;

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
    catch (error) {            // If here, we had some other error
            displayElement.innerText = `Error: ${JSON.stringify(error)}`;
        }
    }

displayUI();