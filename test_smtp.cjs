const nodemailer = require('nodemailer');

async function test(email) {
  console.log(`Testing SMTP for: ${email}`);
  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false, // TLS
    auth: {
      user: email,
      pass: 'fmdrdczrxkpjrbsv'
    },
    tls: {
      ciphers: 'SSLv3',
      rejectUnauthorized: false
    }
  });

  try {
    await transporter.verify();
    console.log(`✅ Success verifying SMTP connection for: ${email}`);

    console.log('Sending test message...');
    const info = await transporter.sendMail({
      from: '"UtilitySense Verification" <verify.software2040@pgel.in>',
      to: 'verify.software2040@pgel.in',
      subject: 'Office 365 SMTP Test - UtilitySense',
      text: 'This is a test email verifying that Office 365 SMTP is working properly.'
    });
    console.log(`✅ Test email successfully sent! Message ID: ${info.messageId}`);
  } catch (err) {
    console.error(`❌ Failed for: ${email}. Error: ${err.message}`);
  }
}

async function run() {
  await test('verify.software2040@pgel.in');
}

run();
