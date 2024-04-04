const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const xlsx = require("xlsx");

const app = express();
const PORT = 3000;

// Parse URL-encoded bodies (as sent by HTML forms)
app.use(bodyParser.urlencoded({ extended: true }));

// Serve HTML file for the form
app.get("/", (req, res) => {
    res.sendFile(__dirname + "/index.html");
});

// Handle form submission
app.post("/submit", (req, res) => {
    const { name, email, phone, date, time, reason } = req.body;

    // Check if any field is empty
    if (!name || !email || !phone || !date || !time || !reason) {
        res.status(400).send("Please fill out all fields.");
        return;
    }

    // Prepare data for Excel
    const data = [name, email, phone, date, time, reason];

    // Define path to Excel file
    const fileName = "appointments.xlsx";

    // If file exists, load existing data
    let existingData = [];
    if (fs.existsSync(fileName)) {
        const workbook = xlsx.readFile(fileName);
        existingData = xlsx.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[0]], { header: 1 }
        );
    }

    // Append new data to existing data
    existingData.push(data);

    // Convert data to worksheet
    const ws = xlsx.utils.aoa_to_sheet(existingData);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Appointments");

    // Write workbook to a file
    xlsx.writeFile(wb, fileName);

    // Send response
    res.send("Form submitted successfully!");
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});