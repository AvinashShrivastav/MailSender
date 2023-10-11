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

<body>
  <table style="color: #500050; background-color: #f4f4f4; max-width: 660px;" role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
      <tr>
        <td style="background-color: #ffffff; background-position: 50% 50%; background-repeat: no-repeat; background-size: cover;" valign="top">
          <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
            <tbody>
              <tr>
                <td style="background-position: 50% 50%; background-repeat: no-repeat; background-size: cover;" valign="top">
                  <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tbody>
                      <tr>
                        <td style="padding-top: 0px; padding-bottom: 0px;" colspan="12" valign="top" width="100%">
                          <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0">
                            <tbody>
                              <tr>
                                <td style="background-color: #000000; padding: 12px 0px 0px;" align="full" valign="top">
                                  <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fdiscord.com%2Finvite%2FAfVYrSemzB&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=f897a5c6a6e24e3ebe76139576b0bb7aea4ab58c7cf4e5440bd5ca2708020788" target="_blank" style="display: block;" rel="noopener">
                                    <img class="gmail-CToWUd" style="border: 0px; width: 660px; height: auto; max-width: 100%; display: block;" role="presentation" src="https://ci6.googleusercontent.com/proxy/BuCjEWEnJEP5LSi_TZHPp7ZciNo_bUnFfyKt9AidvMU9fd6NhSsG_nA-RxE5njNtvNZFFw7wkL3ope8or-BAigxPLTfQzDMAE__d_-iKUexz9MEG_vWstXW9gsQLnIbXiVsvoCvbEW7OjWffNk66-rD1_zvgFg=s0-d-e1-ft#https://mcusercontent.com/c8c24fed9e80867c69eb3291b/images/56043619-7c7e-1f02-9aa0-03bb106aab61.png" alt="" width="660" height="auto">
                                  </a>
                                </td>
                              </tr>
                              <tr>
                                <td style="background-color: #000000; padding: 12px 24px;" align="center" valign="top">
                                  <table role="presentation" border="0" cellspacing="0" cellpadding="0" align="center">
                                    <tbody>
                                      <tr>
                                        <td style="background-color: #90c4f2; border-radius: 50px; text-align: center;" valign="top">
                                          <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fhackcbs.tech%2F&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=1af0cfa4e797275527a9832165894df4044cb24662d866bb2257ae933b9770fb" target="_blank" style="color: #ffffff; border-radius: 50px; border: 2px solid #ffffff; display: block; font-family: 'Helvetica Neue', Helvetica, Arial, Verdana, sans-serif; font-size: 16px; padding: 16px 28px; text-decoration-line: none; min-width: 30px; direction: ltr; letter-spacing: 0px;" rel="noopener">Register now!</a>
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                              <tr>
                                <td style="background-color: #000000; padding: 20px 24px;" valign="top">
                                  <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0">
                                    <tbody>
                                      <tr>
                                        <td style="min-width: 100%; border-top: 2px solid #ffffff;" valign="top">&nbsp;</td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                              <tr>
                                <td style="background-color: #000000; padding: 12px 0px;" valign="top">
                                  <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
                                    <tbody>
                                      <tr>
                                        <td style="background-position: 50% 50%; background-repeat: no-repeat; background-size: cover; padding-top: 0px; padding-bottom: 0px;" valign="top">
                                          <table role="presentation" border="0" width="100%" cellspacing="24" cellpadding="0">
                                            <tbody>
                                              <tr>
                                                <td style="margin-bottom: 24px;" colspan="12" valign="top" width="100%">
                                                  <table role="presentation" border="0" width="100%" cellspacing="0" cellpadding="0">
                                                    <tbody>
                                                      <tr>
                                                        <td align="center" valign="top">
                                                          <table role="presentation" border="0" width="" cellspacing="0" cellpadding="0">
                                                            <tbody>
                                                              <tr>
                                                                <td style="padding-left: 24px; padding-top: 0px; padding-right: 24px;" valign="top">
                                                                  <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fwww.facebook.com%2Fhackcbs%2F&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=576ee7982893a5cd6d464ea2e9f5f86c9103ddeb887d1030d1a9f4839775e5b6" target="_blank" style="display: block;" rel="noopener">
                                                                    <img class="gmail-CToWUd" style="border: 0px; width: 24px; height: auto; max-width: 100%; display: block;" src="https://ci3.googleusercontent.com/proxy/1LyF0Pnt-tl6Y94m1UPfpsfUBcjfThno-iIT9ybxGTgCuY0ZODJLuE0_WbvIFzLzaGGsdWjjdo57yQsaYumE2_8o-tiGzmHSS7gIBSVRUl3Cyc_sDeiAODqxIdXYIXU17CejuPp-mEhdeItgNIQ_cRNtPg=s0-d-e1-ft#https://cdn-images.mailchimp.com/icons/social-block-v3/block-icons-v3/facebook-icon-light-40.png" alt="Facebook icon" width="24" height="auto">
                                                                  </a>
                                                                </td>
                                                                <td style="padding-left: 24px; padding-top: 0px; padding-right: 24px;" valign="top">
                                                                  <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fwww.instagram.com%2Fhackcbs%2F&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=9dc3cc5ad13a89a2afa2b5c398303579fe0d6d6a9acba8ad870907edd4c1aec1" target="_blank" style="display: block;" rel="noopener">
                                                                    <img class="gmail-CToWUd" style="border: 0px; width: 24px; height: auto; max-width: 100%; display: block;" src="https://ci5.googleusercontent.com/proxy/gxE2oHBQLGP0r0fFyRdaC88UATtBqx_rOAMg2bdEI6hzbvVC9LC0hKkCZUTnOWEgvFu16x26G3WqUqQZEkRS50aq1T_3SJY8a8XA1hjO2Qlb28MNGirw3PNpBndDN_jWgLYunnuVykyVQ2eucdEyc5-Alzw=s0-d-e1-ft#https://cdn-images.mailchimp.com/icons/social-block-v3/block-icons-v3/instagram-icon-light-40.png" alt="Instagram icon" width="24" height="auto">
                                                                  </a>
                                                                </td>
                                                                <td style="padding-left: 24px; padding-top: 0px; padding-right: 24px;" valign="top">
                                                                  <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fwww.twitter.com%2Fhackcbs%2F&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=f1fd77ca3bcfb77728aeb952e55895203cdfd64780083f265440ec73e79bc3b2" target="_blank" style="display: block;" rel="noopener">
                                                                    <img class="gmail-CToWUd" style="border: 0px; width: 24px; height: auto; max-width: 100%; display: block;" src="https://ci6.googleusercontent.com/proxy/PolSFWGj8u-Sz39sTEeU7DGb-ccfDiT6hqMJURyijKCU2loUfUMr9W5vhBFfzh7_XtXy5wcER7BAUy_YX0A_PXNUrL1xNyKfuz4D-9EvdkE39yPb8R0jXMty6t5R6gV7HA0BlRCkKsq9gXPIRuwhrFII=s0-d-e1-ft#https://cdn-images.mailchimp.com/icons/social-block-v3/block-icons-v3/twitter-icon-light-40.png" alt="Twitter icon" width="24" height="auto">
                                                                  </a>
                                                                </td>
                                                                <td style="padding-left: 24px; padding-top: 0px; padding-right: 24px;" valign="top">
                                                                  <a href="https://us14.mailchimp.com/mctx/clicks?url=https%3A%2F%2Fwww.linkedin.com%2Fcompany%2Fhackcbs%2F&amp;xid=566c0b636e&amp;uid=198429754&amp;iid=10031656&amp;pool=template_test&amp;v=2&amp;c=1696959854&amp;h=58a0397cd74fccae99e66a8dd29a226d706a1f7e02a9dac37717ee05ce6c2f40" target="_blank" style="display: block;" rel="noopener">
                                                                    <img class="gmail-CToWUd" style="border: 0px; width: 24px; height: auto; max-width: 100%; display: block;" src="https://ci3.googleusercontent.com/proxy/OE-5hM_b9J7TiyoxNc5djsXL2dV6RGJrj9sfApLVV4vrWXn-dCp1ChEQdqxrmpBcgE5keTnFWqyiyfZxPNd4x_L-DvRxXLzVn0bJuixJgZkjJcJ1HIEmyEstOQXmYUanAbBXVpJdAlGGFvNRHQ0UohfFOA=s0-d-e1-ft#https://cdn-images.mailchimp.com/icons/social-block-v3/block-icons-v3/linkedin-icon-light-40.png" alt="LinkedIn icon" width="24" height="auto">
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </td>
                                                      </tr>
                                                    </tbody>
                                                  </table>
                                                </td>
                                              </tr>
                                            </tbody>
                                          </table>
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </td>
      </tr>
    </tbody>
  </table>
  
  
</body>

</html>
  `;

// Rest of your email sending code...


let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'hackcbs@gmail.com',
        pass: 'Your App Password Here' // Use the generated App Password here
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
            from: 'hackcbs@gmail.com',
            to: row.email,
            subject: 'Invitation to ', // Subject line
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




