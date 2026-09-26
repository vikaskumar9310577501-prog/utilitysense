import nodemailer from 'nodemailer';

export default async function handler(req, res) {
  const logs = [];
  const capture = (msg) => {
    logs.push(typeof msg === 'object' ? JSON.stringify(msg) : String(msg));
  };

  const SMTP_USER = 'verify.software2040@pgel.in';
  const SMTP_PASS = 'fmdrdczrxkpjrbsv';

  // Test 1: Standard Office 365 config
  capture(`[DIAGNOSTIC] Node: ${process.version}, Arch: ${process.arch}, Platform: ${process.platform}`);
  
  const transporter1 = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false, // STARTTLS
    auth: {
      user: SMTP_USER,
      pass: SMTP_PASS
    },
    tls: {
      ciphers: 'SSLv3',
      rejectUnauthorized: false
    },
    logger: {
      debug: capture,
      info: capture,
      warn: capture,
      error: capture
    },
    debug: true
  });

  try {
    capture('[TEST 1] Verifying with SSLv3 ciphers...');
    await transporter1.verify();
    capture('[TEST 1] SUCCESS! Sending test email...');
    const info = await transporter1.sendMail({
      from: `"PGEL UtilitySense Verification" <${SMTP_USER}>`,
      to: 'software.2040@pgel.in',
      subject: 'Office 365 Live SMTP Test from Vercel',
      text: 'This is a live test from Vercel Serverless Function.'
    });
    capture(`[TEST 1] Mail Sent! MessageId: ${info.messageId}`);
    return res.status(200).json({ success: true, method: 'SSLv3', logs });
  } catch (err1) {
    capture(`[TEST 1 FAILED]: ${err1.message}`);

    // Test 2: Standard TLS without SSLv3
    const transporter2 = nodemailer.createTransport({
      host: 'smtp.office365.com',
      port: 587,
      secure: false,
      auth: {
        user: SMTP_USER,
        pass: SMTP_PASS
      },
      tls: {
        rejectUnauthorized: false
      },
      logger: {
        debug: capture,
        info: capture,
        warn: capture,
        error: capture
      },
      debug: true
    });

    try {
      capture('[TEST 2] Verifying without SSLv3...');
      await transporter2.verify();
      capture('[TEST 2] SUCCESS! Sending test email...');
      const info = await transporter2.sendMail({
        from: `"PGEL UtilitySense Verification" <${SMTP_USER}>`,
        to: 'software.2040@pgel.in',
        subject: 'Office 365 Live SMTP Test from Vercel (Standard TLS)',
        text: 'This is a live test from Vercel Serverless Function.'
      });
      capture(`[TEST 2] Mail Sent! MessageId: ${info.messageId}`);
      return res.status(200).json({ success: true, method: 'StandardTLS', logs });
    } catch (err2) {
      capture(`[TEST 2 FAILED]: ${err2.message}`);
      return res.status(500).json({ success: false, logs, error: err2.message });
    }
  }
}
