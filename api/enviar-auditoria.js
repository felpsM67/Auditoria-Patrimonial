import nodemailer from 'nodemailer';

export const config = {
  api: {
    bodyParser: false,
  },
};

async function readRequestBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, message: 'Método não permitido.' });
  }

  try {
    const buffer = await readRequestBody(req);
    if (!buffer || buffer.length === 0) {
      return res.status(400).json({ ok: false, message: 'Arquivo ZIP não recebido pela API.' });
    }

    const requiredEnvs = ['SMTP_HOST', 'SMTP_PORT', 'SMTP_USER', 'SMTP_PASS', 'MAIL_FROM', 'MAIL_TO'];
    const missing = requiredEnvs.filter((key) => !process.env[key]);
    if (missing.length) {
      return res.status(500).json({ ok: false, message: `Variáveis de ambiente ausentes: ${missing.join(', ')}` });
    }

    const fileName = req.headers['x-file-name'] || 'Auditoria.zip';
    const responsavel = req.headers['x-responsavel'] || 'Responsável não informado';

    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT || 587),
      secure: String(process.env.SMTP_SECURE) === 'true',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS,
      },
    });

    const info = await transporter.sendMail({
      from: process.env.MAIL_FROM,
      to: process.env.MAIL_TO,
      subject: `Auditoria Patrimonial - ${responsavel}`,
      text: `Segue em anexo o pacote ZIP da auditoria patrimonial realizada por ${responsavel}.`,
      attachments: [
        {
          filename: fileName,
          content: buffer,
          contentType: 'application/zip',
        },
      ],
    });

    return res.status(200).json({
      ok: true,
      message: 'E-mail aceito pelo servidor SMTP.',
      accepted: info.accepted,
      rejected: info.rejected,
      messageId: info.messageId,
    });
  } catch (error) {
    console.error('Erro ao enviar auditoria:', error);
    return res.status(500).json({ ok: false, message: 'Erro ao enviar e-mail.', error: error.message });
  }
}
