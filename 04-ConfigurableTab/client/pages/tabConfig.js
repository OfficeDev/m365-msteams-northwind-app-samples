// import {
//     getLoggedInEmployee
// } from '../northwindIdentity/identityService.js';

import {
    getCategories
} from '../modules/northwindDataService.js';

async function displayUI() {

    const categorySelect = document.getElementById('categorySelect');
    const messageDiv = document.getElementById('message');

    try {
        const employee = 1; // await getLoggedInEmployee();
        if (employee) {
            
            const categories = await getCategories();
            categories.forEach((category) => {
                const option = document.createElement('option');
                option.value = category.categoryId;
                option.innerText = category.displayName;
                categorySelect.appendChild(option);
            });

        }
    }
    catch (error) {            // If here, we had some other error
        messageDiv.innerText = `Error: ${JSON.stringify(error)}`;
    }
}

displayUI();