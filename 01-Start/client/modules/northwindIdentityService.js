export async function getLoggedinEmployeeId() {
    const cookies = document.cookie.split(';');
    for (const c of cookies) {
        const [name, value] = c.split('=');
        if (name === 'employeeId') {
            return value;
        }
    }
    return null;
}

export async function setLoggedinEmployeeId(employeeId) {
    document.cookie = `employeeId=${employeeId};`
}

// Get the employee profile from our web service
export async function getEmployeeProfile(employeeId) {

    const response = await fetch (`/employeeProfile?employeeId=${employeeId}`, {
        "method": "get",
        "headers": {
            "content-type": "application/json"
        },
        "cache": "default"
    });
    if (response.ok) {
        const employeeProfile = await response.json();
        return employeeProfile;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}
