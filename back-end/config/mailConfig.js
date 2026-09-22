const dns = require("dns");
const nodemailer = require("nodemailer");

/*
 * Railway/Gmail may resolve smtp.gmail.com to IPv6 first.
 * Prefer IPv4 to avoid ENETUNREACH in environments without IPv6 routing.
 */
dns.setDefaultResultOrder("ipv4first");

const smtpHost =
    String(process.env.SMTP_HOST || "smtp.gmail.com").trim();

const smtpPort =
    Number(process.env.SMTP_PORT || 587);

const smtpSecure =
    String(process.env.SMTP_SECURE || "false")
        .trim()
        .toLowerCase() === "true";

const smtpUser =
    String(process.env.SMTP_USER || "").trim();

const smtpPass =
    String(process.env.SMTP_PASS || "").replace(/\s+/g, "");

if (!smtpUser) {
    throw new Error("SMTP_USER is missing in environment variables.");
}

if (!smtpPass) {
    throw new Error("SMTP_PASS is missing in environment variables.");
}

const transporter = nodemailer.createTransport({
    host: smtpHost,
    port: smtpPort,
    secure: smtpSecure,

    auth: {
        user: smtpUser,
        pass: smtpPass,
    },

    /*
     * Port 587 uses STARTTLS.
     * secure=false is correct for port 587.
     */
    requireTLS: smtpPort === 587,

    connectionTimeout: 20000,
    greetingTimeout: 20000,
    socketTimeout: 30000,
});

/*
 * CRITICAL:
 * Export the Nodemailer transporter itself.
 * Without this line, require("../config/mailConfig")
 * returns an empty object and transporter.sendMail is unavailable.
 */
module.exports = transporter;
