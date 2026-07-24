const express = require('express');
const path = require('path');
const apiHandler = require('./api/index');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files from the 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

// Middleware to parse JSON and URL-encoded data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Map Vercel-style Serverless function to Express route
app.all('/api', (req, res) => {
  apiHandler(req, res);
});

app.listen(PORT, () => {
  console.log(`Server is running locally at http://localhost:${PORT}`);
});
