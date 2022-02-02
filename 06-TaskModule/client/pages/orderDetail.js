import {
    getOrder
} from '../modules/northwindDataService.js';
import {
    getAADUserFromEmployeeId,getUserDetailsFromAAD
} from '../identity/identityClient.js'
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
import templatePayload from '../modules/orderChatCard.js'
let orderId="0";
let card="";
let customerContact="";
let salesRepdetails={};

async function displayUI() {
    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    const btnTaskModuleElement = document.getElementById('btnTaskModule');
    try {
        microsoftTeams.initialize(async () => {
        const searchParams = new URLSearchParams(window.location.search);
        if (searchParams.has('orderId')) {
              orderId = searchParams.get('orderId');

            const order = await getOrder(orderId);    
            //graph call to get AAD mapped employee details        
            const user=await getAADUserFromEmployeeId(order.employeeId);            
            salesRepdetails=await getUserDetailsFromAAD(user);
            customerContact=`${order.contactName}(${order.contactTitle})`;
            displayElement.innerHTML = `
                    <h1>Order ${order.orderId}</h1>
                    <p>Customer: ${order.customerName} <br />
                    Contact: ${order.contactName}, ${order.contactTitle}<br />
                    Date: ${new Date(order.orderDate).toDateString()}</p>
                `;

            order.details.forEach(item => {
                const orderRow = document.createElement('tr');
                orderRow.innerHTML = `<tr>
                        <td>${item.quantity}</td>
                        <td><a href="/pages/productDetail.html?productId=${item.productId}">${item.productName}</a></td>
                        <td>${item.unitPrice}</td>
                        <td>${item.discount}</td>
                    </tr>`;
                detailsElement.append(orderRow);

            });

        }
    btnTaskModuleElement.addEventListener('click',  ev => {
            let submitHandler = (err, result) => { console.log(result); };
            let taskInfo = {
                title: null,
                height: null,
                width: null,
                url: null,
                card: null,
                fallbackUrl: null,
                completionBotId: null,
            };
            console.log(salesRepdetails)
            //todo: not using AC templating due to an issue : Error Cannot read properties of undefined (reading 'AEL')
            card=JSON.stringify(templatePayload).replaceAll('${orderId}',orderId)
            .replaceAll("${contact}",customerContact).replaceAll("${salesRepEmail}",salesRepdetails.mail)
            .replaceAll("${salesRepManagerEmail}",salesRepdetails.managerMail)
            .replaceAll("${salesRepName}",salesRepdetails.displayName)   
            .replaceAll("${salesRepManagerName}",salesRepdetails.managerDisplayName);   
            taskInfo.card = JSON.parse(card);
            taskInfo.title = "chat";
            taskInfo.height = 310;
            taskInfo.width = 430;         
            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        });
    });
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();

