import fetch from 'node-fetch';
import {
    NORTHWIND_ODATA_SERVICE,
    EMAIL_DOMAIN
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

export async function getAllEmployees() {

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Employees/?$select=EmployeeID,FirstName,LastName`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });

    const employees = await response.json();
    return employees.value.map(employee => ({
        employeeId: employee.EmployeeID,
        firstName: employee.FirstName,
        lastName: employee.LastName
    }));
}

export async function getEmployeeProfile(employeeId) {

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Employees(${employeeId})`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const employeeProfile = await response.json();

    return {
        displayName: `${employeeProfile.FirstName} ${employeeProfile.LastName}`,
        mail: `${employeeProfile.FirstName}@${EMAIL_DOMAIN}`,
        photo: employeeProfile.Photo.substring(104), // Trim Northwind-specific junk
        jobTitle: employeeProfile.Title
    }

}

