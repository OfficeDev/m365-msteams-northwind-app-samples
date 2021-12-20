import fetch from 'node-fetch';
import { NORTHWIND_SERVICE } from './constants.js';

export async function getEmployees() {

    const response = await fetch (
        `${NORTHWIND_SERVICE}/Employees?$select=FirstName`,
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
        // `${NORTHWIND_SERVICE}/Employees?$filter=EmployeeID eq 1`,
        `${NORTHWIND_SERVICE}/Employees(${employeeId})`,
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
