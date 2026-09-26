import nodemailer from 'nodemailer';

const SENDER_RELAY_URL = 'https://script.google.com/macros/s/AKfycbyO2guilzdohQC7V0IvAzsxUODjNKxQ3lpStHLZt5gr562Jqe27ZvXX4ybq6eXx49WP/exec';

export default async function handler(req, res) {
  // Set CORS headers
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { email, otp } = req.body;

    if (!email || !otp) {
      return res.status(400).json({ error: 'Email and OTP are required' });
    }

    const SMTP_USER = 'verify.software2040@pgel.in';
    const SMTP_PASS = 'fmdrdczrxkpjrbsv';

    const subject = `PGEL UtilitySense - Login Verification OTP: [${otp}]`;
    const plainText = `Hello,\n\nYour 6-digit secure portal verification code is: ${otp}\n\nThis verification code is valid for 5 minutes. If you did not request this code, please ignore this email or contact your IT Admin.\n\nPG Electroplast Ltd © 2026. All rights reserved.`;
    const htmlBody = `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #eee; border-radius: 10px; max-width: 500px; margin: 0 auto; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
          <div style="text-align: center; margin-bottom: 20px;">
            <h2 style="color: #0284c7; margin: 0; font-size: 24px; font-weight: 800; letter-spacing: -0.5px;">UTILITY SENSE</h2>
            <p style="color: #64748b; font-size: 12px; margin: 5px 0 0 0;">Utility Energy Consumption Management System</p>
          </div>
          <div style="border-top: 1px solid #f1f5f9; padding-top: 20px;">
            <p style="font-size: 14px; color: #334155; line-height: 1.5;">Hello,</p>
            <p style="font-size: 14px; color: #334155; line-height: 1.5;">Your 6-digit secure portal verification code is:</p>
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; text-align: center; margin: 20px 0;">
              <span style="font-size: 32px; font-weight: 800; color: #0f172a; letter-spacing: 4px; font-family: monospace;">${otp}</span>
            </div>
            <p style="font-size: 12px; color: #64748b; line-height: 1.5;">This verification code is valid for 5 minutes. If you did not request this code, please ignore this email or contact your IT Admin.</p>
          </div>
          <div style="border-top: 1px solid #f1f5f9; margin-top: 25px; padding-top: 15px; text-align: center;">
            <p style="font-size: 11px; color: #94a3b8; margin: 0;">PG Electroplast Ltd © 2026. All rights reserved.</p>
          </div>
        </div>
      `;

    const mailOptions = {
      from: `"PGEL UtilitySense Verification" <${SMTP_USER}>`,
      to: email,
      subject: subject,
      text: plainText,
      html: htmlBody
    };

    let emailSent = false;
    let messageId = null;
    let lastError = null;

    // Primary Attempt: Direct Office 365 SMTP
    try {
      const transporter = nodemailer.createTransport({
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
        }
      });
      const info = await transporter.sendMail(mailOptions);
      emailSent = true;
      messageId = info?.messageId;
    } catch (primaryErr) {
      console.warn('Primary Office 365 SMTP attempt failed:', primaryErr.message);
      lastError = primaryErr;

      // Secondary Attempt: Fallback TLS configuration
      try {
        const fallbackTransporter = nodemailer.createTransport({
          host: 'smtp.office365.com',
          port: 587,
          secure: false,
          auth: {
            user: SMTP_USER,
            pass: SMTP_PASS
          },
          tls: {
            rejectUnauthorized: false
          }
        });
        const info = await fallbackTransporter.sendMail(mailOptions);
        emailSent = true;
        messageId = info?.messageId;
      } catch (tlsErr) {
        console.warn('Fallback Office 365 SMTP failed:', tlsErr.message);
        lastError = tlsErr;
      }
    }

    // Tertiary Attempt / Cloud Serverless IP Bypass:
    // When Microsoft 365 blocks serverless/cloud hosting IPs (535 Authentication error),
    // automatically relay through the Corporate HTTP Webhook Relay
    if (!emailSent) {
      console.log('Office 365 direct SMTP blocked by cloud IP policy; engaging Corporate HTTP Relay for:', email);
      try {
        const relayRes = await fetch(SENDER_RELAY_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            to: email,
            email: email,
            subject: subject,
            body: plainText,
            html: htmlBody,
            otp: otp
          }),
          redirect: 'follow'
        });
        if (relayRes.status === 200 || relayRes.status === 302) {
          emailSent = true;
          messageId = `relay-${Date.now()}`;
          console.log('Corporate HTTP Relay delivered OTP successfully!');
        }
      } catch (relayErr) {
        console.error('Corporate HTTP Relay failed:', relayErr.message);
        lastError = relayErr;
      }
    }

    if (emailSent) {
      return res.status(200).json({ success: true, messageId: messageId });
    } else {
      return res.status(500).json({ error: lastError?.message || 'Failed to send verification email' });
    }
  } catch (error) {
    console.error('Error sending email:', error);
    return res.status(500).json({ error: error.message || 'Failed to send verification email' });
  }
}
