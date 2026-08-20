const express = require('express');
const cors = require('cors');
require('dotenv').config();

let onedriveRoutes = require('./routes/API_OneDrive.js');

const app = express();
const PORT = process.env.PORT || 5500;

app.use(cors());
app.use(express.json());

app.use('/api',onedriveRoutes);

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});