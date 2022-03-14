import { getEmployee } from '../modules/northwindDataService.js';
import { inTeams } from '../modules/teamsHelpers.js';

export async function getLoggedinEmployeeId() {
    const cookies = document.cookie.split(';');
    for (const c of cookies) {
        const [name, value] = c.split('=');
        if (name.trim() === 'employeeId') {
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
        "headers": await getFetchHeadersAnon(),
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

    // Redirect to the login page
    if (!(await inTeams())) {
        window.location.href = "/identity/login.html";
    } else {
        window.location.href = "/identity/aadLogin.html";
    }
}

// Headers for use in Fetch (HTTP) requests when calling anonymous web services
// in the server side of this app.
export async function getFetchHeadersAnon() {
    return ({ "content-type": "application/json" });
}

// Headers for use in Fetch (HTTP) requests when calling authenticated web services
// in the server side of this app. Authentication is sent in a cookie, so no
// additional headers are required.
// Other implementations of this module may insert an Authorization header here
export async function getFetchHeadersAuth() {
    return ({ "content-type": "application/json" });
}

export async function getAADUserFromEmployeeId(employeeId) {

    const response = await fetch (`/api/getAADUserFromEmployeeId?employeeId=${employeeId}`, {
        "method": "get",
        "headers": await getFetchHeadersAuth(),       
        "cache": "no-cache"
    });
    if (response.ok) {
        return response.text().then((data) => data ? JSON.parse(data).id : null);      
    } else {
        const error = "getAADUserFromEmployeeId failed"
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}
export async function getUserDetailsFromAAD(aadUserId) {

    const response = await fetch (`/api/getUserDetailsFromAAD?aadUserId=${aadUserId}`, {
        "method": "get",
        "headers": await getFetchHeadersAuth(),       
        "cache": "no-cache"
    });
    if (response.ok) {
        return response.text().then((data) => data ? JSON.parse(data) : null);          
    } else {
        const error = "getUserDetailsFromAAD failed"
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}