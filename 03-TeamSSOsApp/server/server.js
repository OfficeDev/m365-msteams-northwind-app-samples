import express from "express";
import dotenv from 'dotenv';
import cookieParser from 'cookie-parser';

import {
  validateApiRequest,
  validateEmployeeLogin
} from './northwindIdentityService.js';
import {
  getAllEmployees,
  getEmployee,
  getOrder,
  getCategories,
  getCategory,
  getProduct
} from './northwindDataService.js';
import {
  validateAndMapAadLogin
} from './aadIdentityService.js';

dotenv.config();
const app = express();

// JSON middleware is needed if you want to parse request bodies
app.use(express.json());
app.use(cookieParser());

app.use('/api/', validateApiRequest);

// Web service validates a user login
app.post('/api/validateEmployeeLogin', async (req, res) => {

  try {
    const username = req.body.username;
    const password = req.body.password;
    const employeeId = await validateEmployeeLogin(username, password);
    res.send(JSON.stringify({ "employeeId" : employeeId }));
  }
  catch (error) {
      console.log(`Error in /api/validateEmployeeLogin handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service validates an Azure AD login
app.post('/api/validateAadLogin', async (req, res) => {

  try {
    const employeeId = await validateAndMapAadLogin(req, res);
    if (employeeId) {
        res.send(JSON.stringify({ "employeeId" : employeeId }));
      } else {
        res.status(401).send('Unknown authentication failure');
      }
  }
  catch (error) {
      console.log(`Error in /api/validateAadLogin handling: ${error.statusMessage}`);
      res.status(error.status).json({ status: error.status, statusText: error.statusMessage });
  }

});


// Web service returns a list of employees
app.get('/api/employees', async (req, res) => {

  try {
    const employees = await getAllEmployees();
    res.send(employees);
  }
  catch (error) {
      console.log(`Error in /api/employees handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns an employee's profile
app.get('/api/employee', async (req, res) => {

  try {
    const employeeId = req.query.employeeId;
    const employeeProfile = await getEmployee(employeeId);
    res.send(employeeProfile);
  }
  catch (error) {
      console.log(`Error in /api/employee handling: ${error}`);
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
      console.log(`Error in /api/order handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

app.get('/api/categories', async (req, res) => {

  try {
    const categories = await getCategories();
    res.send(categories);
  }
  catch (error) {
      console.log(`Error in /api/categories handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

app.get('/api/category', async (req, res) => {

  try {
    const categoryId = req.query.categoryId;
    const categories = await getCategory(categoryId);
    res.send(categories);
  }
  catch (error) {
      console.log(`Error in /api/category handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

app.get('/api/product', async (req, res) => {

  try {
    const productId = req.query.productId;
    const product = await getProduct(productId);
    res.send(product);
  }
  catch (error) {
      console.log(`Error in /api/product handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});


// Make environment values available on the client side
// NOTE: Do not pass any secret or sensitive values to the client!
app.get('/modules/env.js', (req, res) => {
  res.contentType("application/javascript");
  res.send(`
    export const env = {
      CLIENT_ID: ${process.env.CLIENT_ID};
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
