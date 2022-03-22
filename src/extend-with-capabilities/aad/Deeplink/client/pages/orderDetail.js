import {
    getOrder
} from '../modules/northwindDataService.js';
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';


async function displayUI() {

    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    const copyUrlElement = document.getElementById('btnCopyOrderUrl');
    const copyMsgElement=document.getElementById('copyMessage');
    const copySectionElement=document.getElementById('copySection');
    const errorMsgElement=document.getElementById('message');
    try {

        const searchParams = new URLSearchParams(window.location.search);        
        microsoftTeams.initialize(async () => {
        microsoftTeams.getContext(async (context)=> {      
     
        if (searchParams.has('orderId')||context.subEntityId) {
            const orderId = searchParams.get('orderId')?searchParams.get('orderId'):context.subEntityId;
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
            if(searchParams.has('orderId')){
                copySectionElement.style.display = "flex";
                copyUrlElement.addEventListener('click', async ev => {
                    try { 
                        //temp textarea for copy to clipboard functionality
                        var textarea = document.createElement("textarea");
                        const encodedContext = encodeURI(`{"subEntityId": "${order.orderId}"}`);                            
                        //form the deeplink                       
                        const deeplink = `https://teams.microsoft.com/l/entity/c42d89e3-19b2-40a3-b20c-44cc05e6ee26/OrderDetails?&context=${encodedContext}`;
                        textarea.value = deeplink;
                        document.body.appendChild(textarea);
                        textarea.select();
                        document.execCommand("copy"); //deprecated but there is an issue with navigator.clipboard api
                        document.body.removeChild(textarea); 
                        copyMsgElement.innerHTML="Link copied!"
                    
                    } catch (err) {
                        console.error('Failed to copy: ', err);
                      }});
            }else{
                copySectionElement.style.display = "none";
               
            }
        }else{
            errorMsgElement.innerText = `No order to show`;
            displayElement.style.display="none";
            orderDetails.style.display="none";
        }
    });
});
       
    }
    catch (error) {            // If here, we had some other error
        errorMsgElement.innerText = `Error: ${JSON.stringify(error)}`;
    }
}


displayUI();