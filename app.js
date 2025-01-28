const express = require("express");
const fs = require("fs");
const path = require("path");
const nodemailer = require("nodemailer");
const sql = require("mssql");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

// MSSQL configuration
const dbConfig = {
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    server: process.env.DB_HOST,
    database: process.env.DB_NAME,
    options: {
        encrypt: true, // Use this if you're connecting to Azure
        trustServerCertificate: true, // Use this if you're working on a local dev server
    },
};

const transporter = nodemailer.createTransport({
    host: "smtp.office365.com", // Outlook SMTP server
    port: 587,                 // SMTP port for secure communication
    secure: false,             // Use TLS (not SSL)
    auth: {
        user: process.env.EMAIL_USER, // Your Outlook email address
        pass: process.env.EMAIL_PASS, // Your Outlook email password or app password
    },
});


// Folder paths
const statementsDir = path.join(__dirname, "Statements");
const successDir = path.join(__dirname, "successful_sent_statements");

// Helper function to send emails
async function sendEmail(customerEmail, copyEmails, pdfPath) {
    try {
        await transporter.sendMail({
            from: process.env.EMAIL_USER,
            to: customerEmail,
            cc: copyEmails,
            subject: "Your Statement from Farmers Choice",
            text: "Please find attached your statement.",
            attachments: [
            {
                filename: path.basename(pdfPath),
                path: pdfPath,
            },
            ],
        });
        console.log(`Email sent to ${customerEmail}`);
        return true;
    } catch (error) {
        console.error(`Failed to send email to ${customerEmail}:`, error.message);
        return false;
    }
}

// Main function to process statements
async function processStatements() {
    try {
        const pool = await sql.connect(dbConfig);

        const files = fs.readdirSync(statementsDir);

        for (const file of files) {
            if (path.extname(file) === ".pdf") {
                const customerNumber = file.split("_")[1]?.replace(".pdf", "");
                const pdfPath = path.join(statementsDir, file);

                if (!customerNumber) {
                    console.error(`Invalid file name format: ${file}`);
                    continue;
                }

                // Fetch customer email from database
                const result = await pool
                    .request()
                    .input("customerNumber", sql.VarChar, customerNumber)
                    .query("SELECT TOP 1 [E-Mail] AS email FROM [FCL1$Customer$437dbf0e-84ff-417a-965d-ed2bb9650972] WHERE [No_] = @customerNumber");

                if (result.recordset.length === 0) {
                    console.error(`No email found for customer: ${customerNumber}`);
                    continue;
                }

                // Fetch default copy emails from database
                const copyResult = await pool
                    .request()
                    .query("SELECT [Email] AS email FROM [FCL1$Default Copy$23dc970e-11e8-4d9b-8613-b7582aec86ba]");

                const copyEmails = copyResult.recordset.map(record => record.email).join(", ");

                const customerEmail = result.recordset[0].email;

                // Send email and move file if successful
                const emailSent = await sendEmail(customerEmail, copyEmails, pdfPath);
                console.log(`customer email: ${customerEmail}`)

                if (emailSent) {
                    const successPath = path.join(successDir, file);
                    fs.renameSync(pdfPath, successPath);
                    console.log(`File moved to ${successPath}`);
                }
            }
        }

        await pool.close();
    } catch (error) {
        console.error("Error processing statements:", error.message);
    }
}

// Route to trigger processing
app.get("/send-statements", async (req, res) => {
    await processStatements();
    res.send("Statements processed.");
});

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
