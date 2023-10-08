const nodemailer = require('nodemailer');
const XLSX = require('xlsx');

const path = require('path');
const fs = require('fs');

// Read the image file as a data URI
let imagePath = path.join(__dirname, 'image.png'); // Path to your image file
let imageContent = fs.readFileSync(imagePath, { encoding: 'base64' });
let dataUri = `data:image/jpeg;base64,${imageContent}`;
// console.log(dataUri);

let htmlContent = `
<html>

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tech Festival</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #282c35;
            color: #c8c8c8;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
        }

        .container {
            width: 80%;
            background-color: #3e3e3e;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            text-align: center;
        }

        h1 {
            color: #61dafb;
            margin-bottom: 20px;
        }

        .fest-content {
            margin-top: 20px;
        }

        .fest-content p {
            margin-bottom: 15px;
        }

        .fest-content a {
            color: #4caf50;
            text-decoration: underline;
        }

        .highlight {
            color: #f8c32b;
            font-weight: bold;
        }
    </style>
</head>

<body>

    <div class="container">
        <h1>Tech Festival 2023</h1>
        <div class="fest-content">
            <p>Welcome to Tech Festival 2023, where <span class="highlight">innovation</span> meets <span
                    class="highlight">excitement</span>! Join us for a thrilling journey into the world of technology.</p>
            <p>Event Highlights:</p>
            <ul>
                <li>Keynote speakers from tech giants</li>
                <li>Hands-on workshops and coding challenges</li>
                <li>Panel discussions on the future of technology</li>
                <li>Networking opportunities with industry experts</li>
            </ul>
            <p>Event Date: <span class="highlight">November 10-12, 2023</span></p>
            <p>Location: <span class="highlight">[Your Venue]</span></p>
            <p>For more information, visit our <a href="https://github.com/your-tech-festival">GitHub Repository</a>.
            </p>
        </div>
    </div>

</body>

</html>
  `;

// Rest of your email sending code...


let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'avinashshrivastavaofficial@gmail.com',
        pass: 'wcew zrjr ofqi yqwi' // Use the generated App Password here
    }
});
// C:\Users\hp\Desktop\mail\email-sender\sendmail.js
// Rest of your email sending code remains the same...
const workbook = XLSX.readFile('EmailList.xlsx'); // Replace 'emails.xlsx' with your Excel file name
const sheetName = workbook.SheetNames[0];
const emailData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
console.log(emailData)


// HTML content for the email
// let htmlContent = `
//     <h1>Hello!</h1>
//     <p>This is a test email sent from Node.js using Nodemailer with HTML content and an image.</p>
//     <img src="https://www.google.com/imgres?imgurl=https%3A%2F%2Fplay-lh.googleusercontent.com%2FKSuaRLiI_FlDP8cM4MzJ23ml3og5Hxb9AapaGTMZ2GgR103mvJ3AAnoOFz1yheeQBBI&tbnid=p0BwJ0uiMmavoM&vet=12ahUKEwid-7-9q-SBAxWikWMGHSrLDpMQMygAegQIARBt..i&imgrefurl=https%3A%2F%2Fplay.google.com%2Fstore%2Fapps%2Fdetails%3Fid%3Dcom.google.android.gm%26hl%3Den_US&docid=r4FMRStPB37O9M&w=512&h=512&q=gmail&ved=2ahUKEwid-7-9q-SBAxWikWMGHSrLDpMQMygAegQIARBt" alt="Example Image">
// `;

async function sendEmails() {
    for (let index = 0; index < emailData.length; index++) {
        const row = emailData[index];
        console.log("row", row.email);
        let mailOptions = {
            from: 'avinashshrivastavaofficial@gmail.com',
            to: row.email,
            subject: 'HTML Email Test', // Subject line
            html: htmlContent // HTML content as the email body
        };

        try {
            let info = await transporter.sendMail(mailOptions);
            console.log(`Email sent to ${row.email}: ${info.response}`);
            // Update status in the Excel file
            emailData[index].status = 'Sent';
        } catch (error) {
            console.log(`Error sending email to ${row.email}: ${error.message}`);
            // Update status in the Excel file
            emailData[index].status = 'Failed';
        }
    }

    // Update the Excel file with the status
    const updatedWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(updatedWorkbook, XLSX.utils.json_to_sheet(emailData), sheetName);
    XLSX.writeFile(updatedWorkbook, 'EmailList.xlsx'); // Save the updated data to the same Excel file
}

sendEmails();




