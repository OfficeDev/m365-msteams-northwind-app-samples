import fetch from 'node-fetch';
import {
    NORTHWIND_ODATA_SERVICE
} from './constants.js';

// Mock identity service based on Northwind employees

const anonPaths = ['/employees', '/validateEmployeeLogin']
export async function validateApiRequest(req, res, next) {
    try {
        if (anonPaths.includes(req.path)) {
            console.log(`Skipped authentication on /api${req.path}`);
            next();
        } else if (req.cookies.employeeId && parseInt(req.cookies.employeeId) > 0) {
            console.log(`Validated authentication on /api${req.path}`);
            next();
        } else {
            console.log(`Invalid authentication on /api${req.path}`);
            res.status(401).json({ status: 401, statusText: "Access denied" });
        }
    }
    catch (error) {
        res.status(401).json({ status: 401, statusText: error });
    }
}

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

