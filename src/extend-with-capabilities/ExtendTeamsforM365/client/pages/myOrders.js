import {
    getLoggedInEmployee
} from '../identity/identityClient.js';
async function displayUI() {   
    const messageDiv = document.getElementById('message');
    try {
        const employee = await getLoggedInEmployee();
        if (employee) {            
            displayAllMyOrders(employee);
        }
        if(microsoftTeams.app !== undefined) {
            microsoftTeams.app.initialize();
            microsoftTeams.app.getContext().then(context=> { 
              if(context &&context.app.host.name==="Office" ){
                displayMyRecentOrders(employee.orders);
              }else if(context &&context.app.host.name==="Teams"){
                const displayElement = document.getElementById('rOchart');
                displayElement.style.display="flex";
              }           
            });           
        }
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }   
}

displayUI();
function displayAllMyOrders(employee) {
    const ordersElement = document.getElementById('orders');
    const displayElement = document.getElementById('content');
    displayElement.innerHTML = `
            <h3>Orders for ${employee.displayName}<h3>
        `;

    employee.orders.forEach(order => {
        const orderRow = document.createElement('tr');
        orderRow.innerHTML = `<tr>
            <td><a href="/pages/orderDetail.html?orderId=${order.orderId}">${order.orderId}</a></td>
            <td>${(new Date(order.orderDate)).toDateString()}</td>
            <td>${order.shipName}</td>
            <td>${order.shipAddress}, ${order.shipCity} ${order.shipRegion || ''} ${order.shipPostalCode || ''} ${order.shipCountry}</td>
        </tr>`;
        ordersElement.append(orderRow);

    });
}
function displayMyRecentOrders(orders) {
    const displayElement = document.getElementById('rOContent');
    const rordersElement = document.getElementById('rOTable');
    const rOFilesElement = document.getElementById('rOFiles');
    const recentOrderArea = document.getElementById('rODiv');
    recentOrderArea.style.display = "block";
    displayElement.innerHTML = `
    <h3>My recent orders<h3>
`;
    orders.splice(1, 5).forEach(order => {
        const orderRow = document.createElement('tr');
        orderRow.innerHTML = `<tr>
    <td><a href="/pages/orderDetail.html?orderId=${order.orderId}">${order.orderId}</a></td>
    <td>${(new Date(order.orderDate)).toDateString()}</td>
    <td>${order.shipName}</td>
    <td>${order.shipAddress}, ${order.shipCity} ${order.shipRegion || ''} ${order.shipPostalCode || ''} ${order.shipCountry}</td>
</tr>`;
        rordersElement.append(orderRow);
    });
    recentFiles.forEach(file => {
        const orderRow = document.createElement('tr');
        orderRow.innerHTML = `<tr>
    <td><a href="/">${file.name}</a></td>             
    <td>${file.modified}</td>
          </tr>`;
        rOFilesElement.append(orderRow);
    });
}
const recentFiles=[{"name":"INV-order987","modified":"Today at 7.am"},
{"name":"Invitation to partners","modified":"Yesterday at 4.pm"},
{"name":"FAQ - Faulty deliveries","modified":"Yesterday at 2.pm"},
{"name":"Customer survey order 9787","modified":"Yesterday at 11.am"}];