const nodemailer = require('nodemailer');

async function test(email) {
  console.log(`Testing SMTP for: ${email}`);
  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false, // TLS
    auth: {
      user: email,
      pass: 'nsxfmjjkskdrbbtt'
    },
    tls: {
      ciphers: 'SSLv3',
      rejectUnauthorized: false
    }
  });

  try {
    await transporter.verify();
    console.log(`✅ Success for: ${email}`);
  } catch (err) {
    console.error(`❌ Failed for: ${email}. Error: ${err.message}`);
  }
}

async function run() {
  await test('verifysoftwre2040@pgel.in');
  await test('verifysoftware2040@pgel.in');
}

run();
