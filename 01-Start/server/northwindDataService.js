import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE } from './constants.js';

export async function getOrdersForEmployee(employeeId) {

    const response = await fetch (
        `${NORTHWIND_ODATA_SERVICE}/Orders?$filter=EmployeeID eq ${employeeId}&$top=10`,
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