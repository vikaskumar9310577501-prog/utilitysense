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
      from: '"UtilitySense Reports" <verify.software2040@pgel.in>',
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

    const info = await transporter.sendMail(mailOptions);
    return res.status(200).json({
      success: true,
      messageId: info.messageId,
      message: `Report email sent successfully to ${toList}`
    });
  } catch (error) {
    console.error('Error sending report email:', error);
    return res.status(500).json({
      error: 'Failed to send report email',
      details: error.message
    });
  }
}
