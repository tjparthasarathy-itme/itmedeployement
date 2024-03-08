const express = require("express");
const exceljs = require("exceljs");
const path = require("path");
const bodyParser = require("body-parser");
const nodemailer = require("nodemailer");
require("dotenv").config();

const app = express();
const port = process.env.PORT || 3000;
const serverIP = "118.139.177.213";

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.use(express.static(__dirname));

function sendEmail(formData, formType) {
  let transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  let mailOptions = {
    from: process.env.EMAIL_USER,
    to: "raju@itmeindia.com", // Replace with the owner's email
    subject:
      formType === "booking"
        ? "New Booking Request"
        : "Contact Us Form Submission",
    html: generateEmailContent(formData, formType),
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log("Error sending email:", error);
    } else {
      console.log("Email sent:", info.response);
    }
  });
}

async function addToExcel(formData, fileName) {
  const excelPath = path.join(__dirname, fileName);
  const workbook = new exceljs.Workbook();

  try {
    // Load existing workbook if file exists
    await workbook.xlsx.readFile(excelPath);
  } catch (error) {
    // Create a new workbook if the file doesn't exist
    console.log("Creating a new Excel file");
  }

  const sheet =
    workbook.getWorksheet(1) || workbook.addWorksheet("Form Submissions");

  if (sheet.rowCount === 1) {
    // If the sheet is empty, add header row
    sheet.addRow(Object.keys(formData));
  }

  sheet.addRow(Object.values(formData));

  // Save the workbook
  await workbook.xlsx.writeFile(excelPath);
  console.log("Form data added to Excel file");
}
function generateEmailContent(formData, formType) {
  let content = ` 
        <p>New Booking Form Submission:</p> 
        <p>Name: ${formData["user-name"]}</p> 
        <p>Designation: ${formData["designation"]}</p> 
        <p>Company: ${formData["company"]}</p> 
        <p>Mobile Number: ${formData["mobile-number"]}</p> 
        <p>E-Mail: ${formData["email"]}</p> 
        <p>Address: ${formData["address"]}</p> 
        <p>City/State: ${formData["city-state"]}</p> 
        <p>Booking Type: ${formData["bookingType"]}</p> 
    `;

  // Only join cities if available and not a contactUs form
  if (
    formData["places"] &&
    Array.isArray(formData["places"]) &&
    formData["places"].length > 0 &&
    formType !== "contactUs"
  ) {
    const selectedCities = formData["places"].join(", ");
    content += `<p>Selected Cities: ${selectedCities}</p>`;
  }

  if (formType === "booking") {
    content = ` 
        <p>New Booking Form Submission:</p> 
        <p>Name: ${formData["user-name"]}</p> 
        <p>Designation: ${formData["designation"]}</p> 
        <p>Company: ${formData["company"]}</p> 
        <p>Mobile Number: ${formData["mobile-number"]}</p> 
        <p>E-Mail: ${formData["email"]}</p> 
        <p>Address: ${formData["address"]}</p> 
        <p>City/State: ${formData["city-state"]}</p> 
        <p>Booking Type: ${formData["bookingType"]}</p> 
    `;
    // Only join cities if available
    if (
      formData["places"] &&
      Array.isArray(formData["places"]) &&
      formData["places"].length > 0
    ) {
      const selectedCities = formData["places"].join(", ");
      content += `<p>Selected Cities: ${selectedCities}</p>`;
    }
  } else if (formType === "contactUs") {
    content = ` 
            <p>First Name: ${formData["firstname"]}</p> 
            <p>Last Name: ${formData["lastname"]}</p> 
            <p>Phone Number: ${formData["phone"]}</p> 
            <p>E-Mail: ${formData["email_id"]}</p> 
            <p>Message: ${formData["message"]}</p> 
        `;
  }

  return content;
}
app.get("/", (req, res) => {
  res.sendFile("index.html", { root: __dirname });
});

app.post("/contactUs", (req, res) => {
  const formData = req.body;
  sendEmail(formData, "contactUs");
  addToExcel(formData, "contactus.xlsx");
  res.status(200).json({ message: "Form submitted successfully" });
});

app.post("/submit-booking", (req, res) => {
  const formData = req.body;
  sendEmail(formData, "booking");
  addToExcel(formData, "form_submission.xlsx");
  res.status(200).json({ message: "Form submitted successfully" });
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

module.exports = { generateEmailContent, sendEmail };
