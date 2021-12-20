// Get the user profile from our web service
async function getUserProfile(clientSideToken) {

    const response = await fetch ("/userProfile", {
        "method": "post",
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

async function displayUI() {

    const displayElement = document.getElementById('content');
    const imageElement = document.getElementById('image');
    try {
        const userProfile = await getUserProfile();

        displayElement.innerHTML = `
            <h1>Hello ${userProfile.displayName}</h1>
            <h3>Profile Information</h3>
            <p>Mail: ${userProfile.mail}<br />
            <p>Company: ${userProfile.company}<br />
            Job Title: ${userProfile.jobTitle}<br />
        `;
        imageElement.src = `data:image/bmp;base64,${userProfile.photo}`;
    }
    catch (error) {            // If here, we had some other error
            displayElement.innerText = `Error: ${JSON.stringify(error)}`;
        }
    }

displayUI();