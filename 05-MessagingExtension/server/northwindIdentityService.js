import fetch from 'node-fetch';
import {
    NORTHWIND_ODATA_SERVICE
} from './constants.js';

// Mock identity service based on Northwind employees

export async function validateEmployeeLogin(username, password) {

    // For simplicity, the username is the employee's surname,
    // and the password is ignored
    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Employees?$filter=LastName eq '${username}'&$select=EmployeeID`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const employees = await response.json();
    if (employees?.value?.length === 1) {
        return employees.value[0].EmployeeID;
    } else {
        return null;
    }
}

