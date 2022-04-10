import {
    getLoggedInEmployee
} from '../identity/identityClient.js';
async function displayUI() {   
const messageDiv = document.getElementById('message');
    try {
        const employee = await getLoggedInEmployee();
        //taos- Initialize and get context for host information to show/hide hub specific displays
        if(microsoftTeams.app !== undefined) {
            microsoftTeams.app.initialize();
            microsoftTeams.app.getContext().then(context=> {                
                if (employee && context.app.host.name!=="Office" ){   
                    document.getElementById("allOrders").style.display="block";     
                    displayAllMyOrders(employee);
                }
                if(context &&context.app.host.name==="Office" ){
                displayMyRecentOrders(employee.orders);
                }else if(context &&context.app.host.name==="Teams"){
                const displayElement = document.getElementById('rOchart');
                displayElement.style.display="flex";
                }           
            });          
            //taos- In a perfect world this button should only work and display in Outlook.
            // if(microsoftTeams.mail.isSupported()){
            //     const displayElementbtn = document.getElementById('btnMail');
            //     displayElementbtn.style.display="block";
            //     displayElementbtn.addEventListener('click', async ev => {
            //     //this is hardcoded for test, should be dynamic (Select and send followup mail)
            //     const input=[{type:"new",toRecipients:"adelev@m365404404.onmicrosoft.com",subject:"Order follow up"}]
            //     await microsoftTeams.mail.composeMail(input);
            //  });
            // } 
        }
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }   
}
//display the order detail tab
displayUI();
//fn definition for all of current user's orders from northwind database
const displayAllMyOrders=(employee)=> {
    const ordersElement = document.getElementById('orders');
    const displayElement = document.getElementById('content');
    displayElement.innerHTML = `<h3>Orders for ${employee.displayName}<h3>`;
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
//fn definition for recent order of current user 
const displayMyRecentOrders=(orders)=> {
    const displayElement = document.getElementById('rOContent');
    const rordersElement = document.getElementById('rOTable');
    const rOFilesElement = document.getElementById('rOFiles');
    const recentOrderArea = document.getElementById('rODiv');
    const recentOrderFilesArea = document.getElementById('rODiv2');
    recentOrderArea.style.display = "block";
    recentOrderFilesArea.style.display = "block";
    displayElement.innerHTML = `<h3>My recent orders<h3>`;
    //Show only top 5 as recent for demo purpose. 
    //Show show based on modified date sort for real world scenario
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
    <td>${file.modified}</td></tr>`;
    rOFilesElement.append(orderRow);
    });
}
//Northwind database does not have file so adding a json obj for demo.
const recentFiles=[{"name":"INV-order987","modified":"Today at 7.am"},
{"name":"Invitation to partners","modified":"Yesterday at 4.pm"},
{"name":"FAQ - Faulty deliveries 2022","modified":"Yesterday at 2.45 pm"},
{"name":"FAQ - Faulty deliveries","modified":"Yesterday at 2.pm"},
{"name":"Customer survey order 9787","modified":"Yesterday at 11.am"},
{"name":"Customer survey order 44587","modified":"Yesterday at 8.30.am"}];