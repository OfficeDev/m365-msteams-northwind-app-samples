import express from "express";
import dotenv from 'dotenv';

import {
  getEmployeeProfile,
  validateEmployeeLogin
} from './northwindIdentityService.js';
import {
  getAllEmployees,
  getOrdersForEmployee,
  getOrder
} from './northwindDataService.js';

dotenv.config();
const app = express();

// JSON middleware is needed if you want to parse request bodies
app.use(express.json());

// Web service validates a user login
app.post('/api/validateEmployeeLogin', async (req, res) => {

  try {
    const username = req.body.username;
    const password = req.body.password;
    const employeeId = await validateEmployeeLogin(username, password);
    res.send(JSON.stringify({ "employeeId" : employeeId }));
  }
  catch (error) {
      console.log(`Error in /validateEmployeeLogin handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns a list of employees
app.get('/api/employees', async (req, res) => {

  try {
    const employees = await getAllEmployees();
    res.send(employees);
  }
  catch (error) {
      console.log(`Error in /getEmployeeProfile handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns an employee's profile
app.get('/api/employeeProfile', async (req, res) => {

  try {
    const employeeId = req.query.employeeId;
    const employeeProfile = await getEmployeeProfile(employeeId);
    res.send(employeeProfile);
  }
  catch (error) {
      console.log(`Error in /getEmployeeProfile handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns an employee's orders
app.get('/api/ordersForEmployee', async (req, res) => {

  try {
    const employeeId = req.query.employeeId;
    const orders = await getOrdersForEmployee(employeeId);
    res.send(orders);
  }
  catch (error) {
      console.log(`Error in /userProfile handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns order details
app.get('/api/order', async (req, res) => {

  try {
    const orderId = req.query.orderId;
    const order = await getOrder(orderId);
    res.send(order);
  }
  catch (error) {
      console.log(`Error in /order handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Make environment values available on the client side
// NOTE: Do not pass any secret or sensitive values to the client!
app.get('/modules/env.js', (req, res) => {
  res.contentType("application/javascript");
  res.send(`
    export const env = {
    };
  `);
});

// Serve static pages from /client
app.use(express.static('client'));

//start listening to server side calls
const PORT = process.env.PORT || 3978;
app.listen(PORT, () => {
  console.log(`Server is Running on Port ${PORT}`);
});
