import { getEmployee } from '../modules/northwindDataService.js';

export async function getLoggedinEmployeeId() {
    const cookies = document.cookie.split(';');
    for (const c of cookies) {
        const [name, value] = c.split('=');
        if (name === 'employeeId') {
            return Number(value);
        }
    }
    return null;
}

export async function setLoggedinEmployeeId(employeeId) {
    document.cookie = `employeeId=${employeeId};SameSite=None;Secure;path=/`;
}

export async function validateEmployeeLogin(surname, password) {

    const response = await fetch (`/api/validateEmployeeLogin`, {
        "method": "post",
        "headers": {
            "content-type": "application/json"
        },
        "body": JSON.stringify({
            "username": surname,
            "password": password
        }),
        "cache": "no-cache"
    });
    if (response.ok) {
        const data = await response.json();
        return data.employeeId;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}

// Get the employee profile from our web service
export async function getLoggedInEmployee() {

    const employeeId = await getLoggedinEmployeeId();

    return await getEmployee(employeeId);

}

export async function logoff() {
    setLoggedinEmployeeId(0);
    window.location.href = "/northwindIdentity/login.html";
}
