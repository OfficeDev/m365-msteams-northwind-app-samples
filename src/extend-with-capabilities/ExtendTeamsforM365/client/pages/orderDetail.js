import {
    getOrder
} from '../modules/northwindDataService.js';
import chatCard from '../cards/orderChatCard.js';
import orderTrackerCard from '../cards/orderTrackerCard.js';
import mailCard from '../cards/orderMailCard.js';
let orderDetails = {};

async function displayUI() {
    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');
    try {
        const searchParams = new URLSearchParams(window.location.search);
        if (searchParams.has('orderId')) {
            const orderId = searchParams.get('orderId');
            const order = await getOrder(orderId);
            orderDetails.orderId = orderId ? orderId : "";
            orderDetails.contact = order.contactName && order.contactTitle ? `${order.contactName}(${order.contactTitle})` : "";
            //get from graph, for demo hardcoded.
            orderDetails.salesRepEmail = "adelev@m365404404.onmicrosoft.com,AlexW@m365404404.onmicrosoft.com";
            orderDetails.salesRepMailrecipients = "adelev@m365404404.onmicrosoft.com; AlexW@m365404404.onmicrosoft.com";
            displayElement.innerHTML = `
                    <h2>Order details for ${order.orderId}</h2>
                    <p><b>Customer:</b> ${order.customerName}<br />
                    <b>Contact:</b> ${order.contactName}, ${order.contactTitle}<br />
                    <b>Date:</b> ${new Date(order.orderDate).toDateString()}<br />
                    <b> ${order.employeeTitle}</b>: ${order.employeeName} (${order.employeeId})
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
            if(microsoftTeams.dialog.isSupported()) {
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
             //taos- chat support
            if(microsoftTeams.chat.isSupported()) {
                const chatArea = document.getElementById("chatBox");
                chatArea.style.display = "block";
                var template = new ACData.Template(chatCard);
                var card = template.expand({ $root: orderDetails });
                var adaptiveCard = new AdaptiveCards.AdaptiveCard();
                adaptiveCard.onExecuteAction = action => {
                    const param = encodeURI(`users="${orderDetails.salesRepEmail}"&topicName=" Enquire about order  ${orderDetails.orderId}"&message=Regarding order #${orderDetails.orderId} `)
                    microsoftTeams.executeDeepLink(`https://teams.microsoft.com/l/chat/0/0?${param}`)
                }
                adaptiveCard.parse(card);
                chatArea.appendChild(adaptiveCard.render());
            } else if(microsoftTeams.mail.isSupported()) {  //taos- mail support
                const mailArea = document.getElementById("mailBox");
                mailArea.style.display = "block";
                var template = new ACData.Template(mailCard);
                var card = template.expand({ $root: orderDetails });
                var adaptiveCard = new AdaptiveCards.AdaptiveCard();
                adaptiveCard.onExecuteAction = action => {
                    microsoftTeams.mail.composeMail({
                        type: microsoftTeams.mail.ComposeMailType.New,
                        subject: `Enquire about order ${orderDetails.orderId}`,
                        toRecipients: [orderDetails.salesRepMailrecipients],
                        message: "Hello",
                    });
                }
                adaptiveCard.parse(card);
                mailArea.appendChild(adaptiveCard.render());
            }
        }
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}

//display the tab for order details
displayUI();

