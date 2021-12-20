import { env } from './env.js';

// Get the user profile from our web service
async function getUserProfile(clientSideToken) {

    return {
        displayName: "Angel Brown",
        mail: "angel@northwind.com",
        company: env.COMPANY_NAME,
        jobTitle: "Team lead"
    }

}

async function displayUI() {

    const displayElement = document.getElementById('content');
    try {
        const userProfile = await getUserProfile();
        displayElement.innerHTML = `
            <h1>Hello ${userProfile.displayName}</h1>
            <h3>Profile Information</h3>
            <p>Mail: ${userProfile.mail}<br />
            <p>Company: ${userProfile.company}<br />
            Job Title: ${userProfile.jobTitle}<br />
        `;
    }
    catch (error) {            // If here, we had some other error
            displayElement.innerText = `Error: ${JSON.stringify(error)}`;
        }
    }

displayUI();