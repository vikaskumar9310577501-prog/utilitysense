import nodemailer from 'nodemailer';

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

    // Configure Microsoft 365 SMTP transport
    const transporter = nodemailer.createTransport({
      host: 'smtp.office365.com',
      port: 587,
      secure: false, // TLS
      auth: {
        user: 'verify.software2040@pgel.in',
        pass: 'nsxfmjjkskdrbbtt'
      },
      tls: {
        ciphers: 'SSLv3',
        rejectUnauthorized: false
      }
    });

    const mailOptions = {
      from: '"Utility Sense Support" <verify.software2040@pgel.in>',
      to: email,
      subject: 'Verify Account OTP - Utility Sense',
      html: `
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
      `
    };

    await transporter.sendMail(mailOptions);
    return res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error sending email:', error);
    return res.status(500).json({ error: error.message || 'Failed to send verification email' });
  }
}
