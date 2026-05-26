const nodemailer = require('nodemailer');

// Environment variables
const gmailEmail = process.env.GMAIL_EMAIL;
const gmailAppPassword = process.env.GMAIL_APP_PASSWORD;
const recipientEmail = process.env.RECIPIENT_EMAIL;
const issueTitle = process.env.ISSUE_TITLE || 'No title';
const issueBody = process.env.ISSUE_BODY || 'No description';
const issueUrl = process.env.ISSUE_URL || 'No URL';
const issueCreator = process.env.ISSUE_CREATOR || 'Unknown';
const repositoryName = process.env.REPOSITORY_NAME || 'Unknown';
const creationDate = process.env.CREATION_DATE || 'Unknown';

// Validate required environment variables
if (!gmailEmail || !gmailAppPassword || !recipientEmail) {
  console.error('Error: Missing required environment variables');
  console.error('Required: GMAIL_EMAIL, GMAIL_APP_PASSWORD, RECIPIENT_EMAIL');
  process.exit(1);
}

// Create email transporter
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: gmailEmail,
    pass: gmailAppPassword,
  },
});

// Format issue body for HTML display
const formattedBody = issueBody
  .replace(/\n/g, '<br>')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;');

// Create email content
const mailOptions = {
  from: gmailEmail,
  to: recipientEmail,
  subject: `[GitHub Issue] ${issueTitle}`,
  html: `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
      <h2 style="color: #333; border-bottom: 2px solid #0366d6; padding-bottom: 10px;">
        New GitHub Issue Created
      </h2>
      
      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="background-color: #f6f8fa;">
          <td style="padding: 10px; font-weight: bold; width: 150px;">Issue Title:</td>
          <td style="padding: 10px;">${issueTitle}</td>
        </tr>
        <tr>
          <td style="padding: 10px; font-weight: bold; background-color: #f6f8fa;">Repository:</td>
          <td style="padding: 10px; background-color: #f6f8fa;">${repositoryName}</td>
        </tr>
        <tr>
          <td style="padding: 10px; font-weight: bold;">Created By:</td>
          <td style="padding: 10px;">${issueCreator}</td>
        </tr>
        <tr style="background-color: #f6f8fa;">
          <td style="padding: 10px; font-weight: bold;">Creation Date:</td>
          <td style="padding: 10px;">${creationDate}</td>
        </tr>
        <tr>
          <td style="padding: 10px; font-weight: bold; background-color: #f6f8fa;">Issue URL:</td>
          <td style="padding: 10px; background-color: #f6f8fa;">
            <a href="${issueUrl}" style="color: #0366d6; text-decoration: none;">${issueUrl}</a>
          </td>
        </tr>
      </table>
      
      <h3 style="color: #333; margin-top: 30px;">Description:</h3>
      <div style="background-color: #f6f8fa; padding: 15px; border-radius: 5px; margin: 10px 0;">
        ${formattedBody}
      </div>
      
      <div style="margin-top: 30px; text-align: center;">
        <a href="${issueUrl}" style="background-color: #0366d6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">
          View Issue on GitHub
        </a>
      </div>
      
      <p style="color: #666; font-size: 12px; margin-top: 30px; text-align: center;">
        This email was sent automatically by GitHub Actions workflow.
      </p>
    </div>
  `,
};

// Send email
transporter.sendMail(mailOptions, (error, info) => {
  if (error) {
    console.error('Error sending email:', error);
    process.exit(1);
  } else {
    console.log('Email sent successfully:', info.response);
    console.log('Message ID:', info.messageId);
    process.exit(0);
  }
});
