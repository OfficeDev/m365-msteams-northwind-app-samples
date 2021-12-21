// Get the user profile from our web service
async function getUserProfile(clientSideToken) {

    const response = await fetch ("/userProfile?userId=2", {
        "method": "get",
        "headers": {
            "content-type": "application/json"
        },
        "cache": "default"
    });
    if (response.ok) {
        const userProfile = await response.json();
        return userProfile;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}

async function getOrdersForEmployee(employeeId)
{
    const response = await fetch ("/ordersForEmployee?employeeId=2", {
        "method": "get",
        "headers": {
            "content-type": "application/json"
        },
        "cache": "default"
    });
    if (response.ok) {
        const orders = await response.json();
        return orders.value;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }    
}

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
            const orderElement = document.createElement('p');
            orderElement.innerText = order.OrderID;
            ordersElement.appendChild(orderElement);
        });
    }
    catch (error) {            // If here, we had some other error
            displayElement.innerText = `Error: ${JSON.stringify(error)}`;
        }
    }

displayUI();