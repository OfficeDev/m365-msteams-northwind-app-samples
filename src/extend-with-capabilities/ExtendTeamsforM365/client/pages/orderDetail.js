import {
    getOrder
} from '../modules/northwindDataService.js';
import chatCard from '../cards/orderChatCard.js';
import orderTrackerCard from '../cards/orderTrackerCard.js';

let orderDetails={};
async function displayUI() {
    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    try {
        const searchParams = new URLSearchParams(window.location.search);
        if (searchParams.has('orderId')) {
            const orderId = searchParams.get('orderId');
            const order = await getOrder(orderId);           
            orderDetails.orderId=orderId?orderId:"";
            orderDetails.contact=order.contactName&&order.contactTitle?`${order.contactName}(${order.contactTitle})`:"";
            //get from graph, for demo hardcoded.
            orderDetails.salesRepEmail="adelev@m365404404.onmicrosoft.com,AlexW@m365404404.onmicrosoft.com";
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
           //taos- Task module button click based on dialog support in hubs
            if (microsoftTeams.dialog.isSupported()) {
                const btnTaskModuleElement = document.getElementById('btnTaskModule');
                btnTaskModuleElement.style.display = "block";
                btnTaskModuleElement.addEventListener('click', ev => {
                    let submitHandler = (err, result) => {
                        //empty for demo purpose. But this is where you get form results from the callback.
                     };
                    var template = new ACData.Template(orderTrackerCard);
                    var cardPayload = template.expand({ $root: orderDetails }); 
                    const taskInfo = {
                        card: cardPayload,
                        title: "chat",
                        height: 310,
                        width: 500,
                        url: null,
                        fallbackUrl: null,
                        completionBotId: null,
                    };
                    microsoftTeams.dialog.open(taskInfo, submitHandler);
                });
            }
            if(microsoftTeams.chat.isSupported()){              
                const chatArea=document.getElementById("chatBox");
                chatArea.style.display="block";
                var template = new ACData.Template(chatCard);
                var card = template.expand({$root: orderDetails});
                var adaptiveCard = new AdaptiveCards.AdaptiveCard();
                adaptiveCard.onExecuteAction = action=> {                   
                    const param=encodeURI(`users="${orderDetails.salesRepEmail}"&topicName=" Enquire about order  ${orderDetails.orderId}"&message=Regarding order #${orderDetails.orderId} `)
                    microsoftTeams.executeDeepLink(`https://teams.microsoft.com/l/chat/0/0?${param}`)
                }
                adaptiveCard.parse(card);                         
                chatArea.appendChild(adaptiveCard.render());
            }else { //add if condition for mail support.
                //for now use hostName, remove once mail support issue is fixed.
                if(microsoftTeams.app !== undefined) {
                    microsoftTeams.app.initialize();
                    microsoftTeams.app.getContext().then(context=> { 
                        if (context.app.host.name==="Outlook" ){ 
                            const mailArea=document.getElementById("mailBox");
                            mailArea.style.display="block";
                            const displayElementbtn = document.getElementById('btnMail');
                            displayElementbtn.addEventListener('click', async ev => {                                               
                            alert('send mail');
                            const input=[{type:"new",toRecipients:"adelev@m365404404.onmicrosoft.com",subject:"Order follow up"}]
                            await microsoftTeams.mail.composeMail(input);
                            });                           
                        }
                    });
                }              

            }
        }
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }}

//display the tab for order details
displayUI();

