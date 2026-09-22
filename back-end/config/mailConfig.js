const nodemailer = require("nodemailer");


const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT || 587),
  secure: String(process.env.SMTP_SECURE).toLowerCase() === "true",

  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },

  family: 4,

  connectionTimeout: 15000,
  greetingTimeout: 15000,
  socketTimeout: 30000,
});