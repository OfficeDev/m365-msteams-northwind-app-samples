import express from "express";
import dotenv from 'dotenv';

import { getUserProfile } from './userProfileService.js';

dotenv.config();
const app = express();

// JSON middleware is needed if you want to parse request bodies
app.use(express.json());

// Web service returns the current user's profile
app.post('/userProfile', async (req, res) => {

  try {
    const userProfile = await getUserProfile();
    res.send(userProfile);
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
      COMPANY_NAME: "${process.env.COMPANY_NAME}"
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
