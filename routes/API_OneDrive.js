const express = require('express');
const router = express.Router();
const onedriveController = require('../controllers/API_OneDrive');

router.post('/upload-excel', onedriveController.uploadExcelFiles);
router.post('/upload-image', onedriveController.uploadImageFiles);

module.exports = router;