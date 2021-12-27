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

let cached_employee = null;
export async function setLoggedinEmployeeId(employeeId) {
    cached_employee = null;
    document.cookie = `employeeId=${employeeId};path=/`;
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

    if (!employeeId) {
        return null;
    }

    if (cached_employee && cached_employee.id === employeeId) {
        return cached_employee;
    }

    const response = await fetch (`/api/employeeProfile?employeeId=${employeeId}`, {
        "method": "get",
        "headers": {
            "content-type": "application/json"
        },
        "cache": "default"
    });
    if (response.ok) {
        const employee = await response.json();
        employee.id = employeeId;
        return employee;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}

export async function logoff() {
    setLoggedinEmployeeId(0);
}
