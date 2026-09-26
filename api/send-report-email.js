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
    const { to, cc, bcc, recipients, subject, html, attachments } = req.body;

    // Normalize recipients
    const toRecipients = to || recipients;
    if (!toRecipients || (Array.isArray(toRecipients) && toRecipients.length === 0)) {
      return res.status(400).json({ error: 'At least one valid "To" recipient is required' });
    }

    const toList = Array.isArray(toRecipients) ? toRecipients.join(', ') : toRecipients;
    const ccList = cc ? (Array.isArray(cc) ? cc.join(', ') : cc) : undefined;
    const bccList = bcc ? (Array.isArray(bcc) ? bcc.join(', ') : bcc) : undefined;

    const SMTP_USER = 'verify.software2040@pgel.in';
    const SMTP_PASS = 'fmdrdczrxkpjrbsv';

    const mailOptions = {
      from: `"UtilitySense Reports" <${SMTP_USER}>`,
      to: toList,
      subject: subject || 'UtilitySense Monthly Report',
      html: html || '<p>Please find attached the UtilitySense Report.</p>',
      attachments: attachments ? attachments.map(att => ({
        filename: att.filename || 'UtilitySense_Report.xlsx',
        content: att.content,
        encoding: att.encoding || 'base64',
        contentType: att.contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      })) : []
    };

    if (ccList && ccList.trim()) mailOptions.cc = ccList;
    if (bccList && bccList.trim()) mailOptions.bcc = bccList;

    let info = null;
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
      info = await transporter.sendMail(mailOptions);
    } catch (primaryErr) {
      console.warn('Primary SSLv3 SMTP attempt failed for reports, trying fallback TLS:', primaryErr.message);
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
      try {
        info = await fallbackTransporter.sendMail(mailOptions);
      } catch (tlsErr) {
        console.warn('Fallback Office 365 SMTP failed for reports:', tlsErr.message);
      }
    }

    // Tertiary Fallback: HTTP Webhook Relay
    if (!info) {
      try {
        const relayRes = await fetch('https://script.google.com/macros/s/AKfycbyO2guilzdohQC7V0IvAzsxUODjNKxQ3lpStHLZt5gr562Jqe27ZvXX4ybq6eXx49WP/exec', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            to: toList,
            email: toList,
            subject: subject || 'UtilitySense Monthly Report',
            body: html ? html.replace(/<[^>]+>/g, '') : 'Please find attached the UtilitySense Report.',
            html: html || '<p>Please find attached the UtilitySense Report.</p>'
          }),
          redirect: 'follow'
        });
        if (relayRes.status === 200 || relayRes.status === 302) {
          info = { messageId: `relay-report-${Date.now()}` };
        }
      } catch (relayErr) {
        console.error('Report HTTP relay failed:', relayErr.message);
      }
    }

    if (info) {
      return res.status(200).json({
        success: true,
        messageId: info.messageId,
        message: `Report email sent successfully to ${toList}`
      });
    } else {
      return res.status(500).json({ error: 'Failed to send report email' });
    }
  } catch (error) {
    console.error('Error sending report email:', error);
    return res.status(500).json({
      error: 'Failed to send report email',
      details: error.message
    });
  }
}
