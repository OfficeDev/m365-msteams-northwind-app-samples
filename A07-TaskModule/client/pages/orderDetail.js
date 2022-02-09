import {
    getOrder
} from '../modules/northwindDataService.js';
import {
    getAADUserFromEmployeeId,getUserDetailsFromAAD
} from '../identity/identityClient.js'
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
import templatePayload from '../modules/orderChatCard.js'

let orderId="0";
let orderDetails={};
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
            let user=await getAADUserFromEmployeeId(order.employeeId);  
             if(!user)   
                user= await getAADUserFromEmployeeId("1");     //fall back to employee 1 
            const salesRepdetails=await getUserDetailsFromAAD(user);          

            orderDetails.orderId=orderId?orderId:"";
            orderDetails.contact=order.contactName&&order.contactTitle?`${order.contactName}(${order.contactTitle})`:"";
            orderDetails.salesRepEmail=salesRepdetails.mail?salesRepdetails.mail:"";
            orderDetails.salesRepManagerEmail=salesRepdetails.managerMail?salesRepdetails.managerMail:"";
            orderDetails.salesRepName=salesRepdetails.displayName?salesRepdetails.displayName:"";
            orderDetails.salesRepManagerName=salesRepdetails.managerDisplayName?salesRepdetails.managerDisplayName:"";

            displayElement.innerHTML = `
                    <h1>Order ${order.orderId}</h1>
                    <p>Customer: ${order.customerName} <br />
                    Contact: ${order.contactName}, ${order.contactTitle}<br />
                    Date: ${new Date(order.orderDate).toDateString()}<br />
                    ${order.employeeTitle}: ${order.employeeName} (${order.employeeId})
                    </p>
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

            var template = new ACData.Template(templatePayload); 
            // Expand the template with your `$root` data object.
            // This binds it to the data and produces the final Adaptive Card payload
            var cardPayload = template.expand({$root: orderDetails});             
            taskInfo.card = cardPayload;
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

