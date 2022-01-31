import dotenv from 'dotenv';
import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE } from './constants.js';

dotenv.config();

// Mock identity service based on Northwind employees

// Wire up middleware
export async function initializeIdentityService(app) {

    // Web service validates a user login
    app.post('/api/validateEmployeeLogin', async (req, res) => {

        try {
            const username = req.body.username;
            const password = req.body.password;
            const employeeId = await validateEmployeeLogin(username, password);
            res.send(JSON.stringify({ "employeeId": employeeId }));
        }
        catch (error) {
            console.log(`Error in /api/validateEmployeeLogin handling: ${error}`);
            res.status(500).json({ status: 500, statusText: error });
        }

    });

    // Middleware to validate all other API requests
    app.use('/api/', validateApiRequest);

}

async function validateApiRequest(req, res, next) {
    try {
        if (req.cookies.employeeId && parseInt(req.cookies.employeeId) > 0) {
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

async function validateEmployeeLogin(username, password) {

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

