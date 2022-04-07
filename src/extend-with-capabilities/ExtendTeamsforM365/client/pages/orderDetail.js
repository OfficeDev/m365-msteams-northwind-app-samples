import {
    getOrder
} from '../modules/northwindDataService.js';

async function displayUI() {

    const displayElement = document.getElementById('content');
    const detailsElement = document.getElementById('orderDetails');

    try {

        const searchParams = new URLSearchParams(window.location.search);
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
            // microsoftTeams.app.initialize();
            if(microsoftTeams.dialog.isSupported()){               
                const btnTaskModuleElement = document.getElementById('btnTaskModule');
                btnTaskModuleElement.style.display="block";
                btnTaskModuleElement.addEventListener('click',  ev => {
                    let submitHandler = (err, result) => {  };             
                    var template = new ACData.Template(templatePayload); 
                    // Expand the template with your `$root` data object.
                    // This binds it to the data and produces the final Adaptive Card payload
                    var cardPayload = template.expand({$root:{}});    
                    const taskInfo = {
                        card:cardPayload,
                        title:"chat",
                        height:310,
                        width:430,
                        url: null,               
                        fallbackUrl: null,
                        completionBotId: null,
                    };  
                    microsoftTeams.dialog.open(taskInfo, submitHandler);
                });
            }

        }
    }
    catch (error) {            // If here, we had some other error
        message.innerText = `Error: ${JSON.stringify(error)}`;
    }
}
const templatePayload=`{
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "spacing": "None",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "2:00 PM",
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "text": "5 days ago",
                            "isSubtle": true,
                            "wrap": true
                        }
                    ],
                    "width": "110px"
                },
                {
                    "type": "Column",
                    "backgroundImage": {
                        "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                        "fillMode": "RepeatVertically",
                        "horizontalAlignment": "Center"
                    },
                    "items": [
                        {
                            "type": "Image",
                            "horizontalAlignment": "Center",
                            "url": "https://messagecardplayground.azurewebsites.net/assets/CircleGreen_coffee.png",
                            "altText": "Location A: Coffee"
                        }
                    ],
                    "width": "auto",
                    "spacing": "None"
                },
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "**Order out for delivery**",
                            "wrap": true
                        },
                        {
                            "type": "ColumnSet",
                            "spacing": "None",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://messagecardplayground.azurewebsites.net/assets/location_gray.png",
                                            "altText": "Location"
                                        }
                                    ],
                                    "width": "auto"
                                },
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Sydney",
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        },
                        {
                            "type": "ImageSet",
                            "spacing": "Small",
                            "imageSize": "Small",
                            "images": [
                                {
                                    "type": "Image",
                                    "url": "https://messagecardplayground.azurewebsites.net/assets/person_m1.png",
                                    "size": "Small",
                                    "altText": "Person with glasses and short hair"
                                }
                            ]
                        },
                        {
                            "type": "ColumnSet",
                            "spacing": "Small",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Delivery team",
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        }
                    ],
                    "width": 40
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "None",
            "columns": [
                {
                    "type": "Column",
                    "width": "110px"
                },
                {
                    "type": "Column",
                    "backgroundImage": {
                        "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                        "fillMode": "RepeatVertically",
                        "horizontalAlignment": "Center"
                    },
                    "items": [
                        {
                            "type": "Image",
                            "horizontalAlignment": "Center",
                            "url": "https://messagecardplayground.azurewebsites.net/assets/Gray_Dot.png"
                        }
                    ],
                    "width": "auto",
                    "spacing": "None"
                },
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://messagecardplayground.azurewebsites.net/assets/car.png",
                                            "altText": "Travel by car"
                                        }
                                    ],
                                    "width": "auto"
                                },
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "about 2 days ago",
                                            "isSubtle": true,
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        }
                    ],
                    "width": 40
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "None",
            "columns": [
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "text": "8:00 PM",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "text": "1 day ago",
                            "isSubtle": true,
                            "wrap": true
                        }
                    ],
                    "width": "110px"
                },
                {
                    "type": "Column",
                    "backgroundImage": {
                        "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                        "fillMode": "RepeatVertically",
                        "horizontalAlignment": "Center"
                    },
                    "items": [
                        {
                            "type": "Image",
                            "horizontalAlignment": "Center",
                            "url": "https://messagecardplayground.azurewebsites.net/assets/CircleBlue_flight.png",
                            "altText": "Location B: Flight"
                        }
                    ],
                    "width": "auto",
                    "spacing": "None"
                },
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "**Reached Brisbane warehouse**",
                            "wrap": true
                        },
                        {
                            "type": "ColumnSet",
                            "spacing": "None",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://messagecardplayground.azurewebsites.net/assets/location_gray.png",
                                            "altText": "Location"
                                        }
                                    ],
                                    "width": "auto"
                                },
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "400 George, Brisbane - 4000, QLD",
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        },
                        {
                            "type": "Image",
                            "url": "https://messagecardplayground.azurewebsites.net/assets/SeaTacMap.png",
                            "size": "Stretch",
                            "altText": "Map of the Seattle-Tacoma airport"
                        }
                    ],
                    "width": 40
                }
            ]
        }
    ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.5"
}`

displayUI();