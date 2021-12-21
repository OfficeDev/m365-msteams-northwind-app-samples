import fetch from 'node-fetch';
import {
    NORTHWIND_ODATA_SERVICE,
    EMAIL_DOMAIN
} from './constants.js';

export async function getEmployees() {

    const response = await fetch (
        `${NORTHWIND_ODATA_SERVICE}/Employees?$select=FirstName`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const data = await response.json();
    return data;

}

export async function getEmployeeProfile(employeeId) {

    const response = await fetch (
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

