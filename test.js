const nodemailer = require("nodemailer");

async function sendEmail() {
  const transporter = nodemailer.createTransport({
    host: "qarirgenerator.com",
    port: 465,
    secure: true, // karena port 465 wajib SSL
    auth: {
      user: "no-reply@qarirgenerator.com",
      pass: "UQ.X-lg045Hr4LfE",
    },
  });

  const info = await transporter.sendMail({
    from: '"Qarir Generator" <no-reply@qarirgenerator.com>',
    to: "osama.work54@email.com",
    subject: "SMTP Nodemailer Test",
    html: "<b>Hai! Email ini dikirim menggunakan Nodemailer ✔</b>",
  });

  console.log("Message sent: ", info.messageId);
}

sendEmail().catch(console.error);
