const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Endpoint to send email
app.post('/send-email', (req, res) => {
    const { to, cc, subject, text } = req.body;

    const transporter = nodemailer.createTransport({
        service: 'gmail', // Replace with your email provider
        auth: {
            user: 'your-email@gmail.com', // Replace with your email
            pass: 'your-password' // Replace with your email password
        }
    });

    const mailOptions = {
        from: 'your-email@gmail.com', // Replace with your email
        to,
        cc,
        subject,
        text
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('Error sending email:', error);
            return res.status(500).send('Failed to send email');
        }
        res.send('Email sent: ' + info.response);
    });
});

// Endpoint to save data to Excel
app.post('/save-data', (req, res) => {
    const { name, id, category, text, to } = req.body;

    // Load existing workbook or create a new one
    let wb;
    try {
        wb = XLSX.readFile('data.xlsx');
    } catch (error) {
        wb = XLSX.utils.book_new();
    }

    // Get the worksheet or create a new one
    let ws = wb.Sheets['Data'];
    if (!ws) {
        ws = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(wb, ws, 'Data');
    }

    // Append new data to the worksheet
    const newRow = { Name: name, ID: id, Category: category, Content: text, To: to };
    const existingData = XLSX.utils.sheet_to_json(ws);
    existingData.push(newRow);
    const newWs = XLSX.utils.json_to_sheet(existingData);
    wb.Sheets['Data'] = newWs;

    // Save the workbook to a file
    XLSX.writeFile(wb, 'data.xlsx');

    res.send({ status: 'success', message: 'Data saved to Excel sheet' });
});

app.listen(3001, () => {
    console.log('Server running on port 3001');
});
