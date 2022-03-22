import {
    getOrder
} from '../modules/northwindDataService.js';
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';
async function displayUI() {

    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    const copyUrlElement = document.getElementById('btnCopyOrderUrl');
    const copyMsgElement=document.getElementById('copyMessage');
    try {

        const searchParams = new URLSearchParams(window.location.search);
        microsoftTeams.initialize(async () => {
        if (searchParams.has('orderId')) {
            const orderId = searchParams.get('orderId');

            const order = await getOrder(orderId);

            displayElement.innerHTML = `
                    <h1>Order ${order.orderId}</h1>
                    <p>Customer: ${order.customerName}<br />
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
            copyUrlElement.addEventListener('click',  ev => {
                try {
                    microsoftTeams.getContext(function (context) { 
                        console.log(context)
                        //entityId
                    var textarea = document.createElement("textarea");
                    var encodedWebUrl = encodeURI(`https://${window.location.hostname}/pages/orderDetail.html?orderId=${orderId}`);
                    const encodedContext = encodeURI(`{"subEntityId": "${orderId}"}`);
                    const deeplink = `https://teams.microsoft.com/l/entity/3d79f44f-c85e-4638-a49c-33b0033a2721/${context.entityId}?webUrl=${encodedWebUrl}&context=${encodedContext}`;
               
                    textarea.value = deeplink;
                    document.body.appendChild(textarea);
                    textarea.select();
                    document.execCommand("copy"); //deprecated but there is an issue with navigator.clipboard api
                    document.body.removeChild(textarea); 
                    copyMsgElement.innerHTML="Link copied!"
                    });
                } catch (err) {
                    console.error('Failed to copy: ', err);
                  }
            });

        }
    });
       
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();