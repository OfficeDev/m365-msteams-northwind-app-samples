import express from "express";
import dotenv from 'dotenv';

import { getEmployeeProfile } from './northwindIdentityService.js';
import { getOrdersForEmployee } from './northwindDataService.js';

dotenv.config();
const app = express();

// JSON middleware is needed if you want to parse request bodies
app.use(express.json());

// Web service returns the current user's profile
app.get('/employeeProfile', async (req, res) => {

  try {
    const employeeId = req.query.employeeId;
    const employeeProfile = await getEmployeeProfile(employeeId);
    res.send(employeeProfile);
  }
  catch (error) {
      console.log(`Error in /userProfile handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error });
  }

});

// Web service returns the current user's profile
app.get('/ordersForEmployee', async (req, res) => {

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
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server is Running on Port ${PORT}`);
});
