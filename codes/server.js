const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const port = 3000;

// Set up multer for file upload
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Serve the HTML page
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

// Handle file upload and conversion
app.post('/upload', upload.single('fileUpload'), (req, res) => {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Extract original file name from uploaded file
    const originalFileName = req.file.originalname;
    const fileNameWithoutExtension = originalFileName.split('.').slice(0, -1).join('.');
    const outputFileName = `${fileNameWithoutExtension}.json`;

    // Assuming the first row contains column headers
    const headers = jsonData.shift(); // Remove headers from data

    // Filter out empty or "junk" entries
    const filteredData = jsonData.filter(row => {
        return Object.values(row).some(value => value !== '');
    });

    const formattedData = filteredData.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });

    const jsonString = JSON.stringify(formattedData, null, 2);

    // Send the JSON data as a downloadable file with the original file name
    res.set({
        'Content-Type': 'application/json',
        'Content-Disposition': `attachment; filename=${outputFileName}`
    });
    res.send(jsonString);
});

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
