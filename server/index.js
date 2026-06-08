import express from 'express';
import cors from 'cors';
import multer from 'multer';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }
});

app.use(cors());
app.use(express.json());

app.get('/api/health', (req, res) => {
  res.json({ ok: true, service: 'auditoria-email-api' });
});

app.post('/api/enviar-auditoria', upload.single('arquivo'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'Arquivo Excel não recebido.' });
    }

    const requiredEnv = ['SMTP_HOST', 'SMTP_PORT', 'SMTP_USER', 'SMTP_PASS', 'MAIL_FROM', 'MAIL_TO'];
    const missing = requiredEnv.filter(key => !process.env[key]);

    if (missing.length) {
      return res.status(500).json({
        ok: false,
        message: `Configuração de e-mail incompleta. Faltando: ${missing.join(', ')}`
      });
    }

    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT),
      secure: String(process.env.SMTP_SECURE || 'false').toLowerCase() === 'true',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    });

    const responsavel = req.body.responsavel || 'Responsável não informado';
    const filial = req.body.filial || 'Filial não informada';
    const assunto = req.body.assunto || `Auditoria Patrimonial - ${responsavel}`;
    const mensagem = req.body.mensagem || `Segue em anexo a auditoria patrimonial realizada por ${responsavel}.`;

    await transporter.sendMail({
      from: process.env.MAIL_FROM,
      to: process.env.MAIL_TO,
      cc: process.env.MAIL_CC || undefined,
      subject: assunto,
      text: `${mensagem}\n\nResponsável: ${responsavel}\nFilial: ${filial}`,
      attachments: [
        {
          filename: req.file.originalname || 'auditoria.xlsx',
          content: req.file.buffer,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    });

    res.json({ ok: true, message: 'E-mail enviado com sucesso.' });
  } catch (error) {
    console.error('Erro ao enviar e-mail:', error);
    res.status(500).json({ ok: false, message: error.message });
  }
});

const PORT = Number(process.env.API_PORT || 3001);
app.listen(PORT, '0.0.0.0', () => {
  console.log(`API de e-mail rodando em http://localhost:${PORT}`);
});
