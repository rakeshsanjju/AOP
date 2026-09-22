/*
 * back-end/config/mailConfig.js
 *
 * Railway-safe email transport using Resend HTTPS API.
 *
 * Why this version:
 * - Railway Free / Trial / Hobby plans block outbound SMTP ports.
 * - HTTPS (port 443) works, so this adapter sends mail through Resend's API.
 * - It exposes transporter.sendMail(...) so your existing
 *   reportMailController.js does NOT need to be rewritten.
 *
 * Required Railway variables:
 *
 *   RESEND_API_KEY=re_xxxxxxxxxxxxxxxxx
 *   MAIL_FROM_NAME=AVIS Reports
 *   MAIL_FROM_ADDRESS=reports@your-verified-domain.com
 *
 * IMPORTANT:
 * MAIL_FROM_ADDRESS must belong to a domain verified in Resend.
 */

const RESEND_API_URL = "https://api.resend.com/emails";

function getRequiredEnv(name) {
    const value = String(process.env[name] || "").trim();

    if (!value) {
        throw new Error(`${name} is missing from environment variables.`);
    }

    return value;
}

function normalizeAddress(value) {
    if (!value) {
        return "";
    }

    /*
     * Nodemailer-style:
     * {
     *   name: "AVIS Reports",
     *   address: "reports@example.com"
     * }
     */
    if (typeof value === "object") {
        const address = String(value.address || "").trim();
        const name = String(value.name || "").trim();

        if (!address) {
            return "";
        }

        return name
            ? `${name} <${address}>`
            : address;
    }

    return String(value).trim();
}

function normalizeRecipients(value) {
    if (!value) {
        return [];
    }

    if (Array.isArray(value)) {
        return value
            .flatMap(item => normalizeRecipients(item))
            .filter(Boolean);
    }

    return String(value)
        .split(",")
        .map(item => item.trim())
        .filter(Boolean);
}

function normalizeAttachments(attachments) {
    if (!Array.isArray(attachments)) {
        return [];
    }

    return attachments
        .map(attachment => {
            if (!attachment) {
                return null;
            }

            const filename =
                String(
                    attachment.filename ||
                    attachment.name ||
                    "attachment"
                ).trim();

            let content = attachment.content;

            if (Buffer.isBuffer(content)) {
                content = content.toString("base64");
            }
            else if (content instanceof Uint8Array) {
                content = Buffer
                    .from(content)
                    .toString("base64");
            }
            else if (typeof content === "string") {
                /*
                 * If your controller already supplied base64,
                 * keep it unchanged.
                 */
                content = content;
            }
            else {
                return null;
            }

            return {
                filename,
                content
            };
        })
        .filter(Boolean);
}

async function sendMail(options = {}) {
    const apiKey =
        getRequiredEnv("RESEND_API_KEY");

    const from =
        normalizeAddress(
            options.from ||
            {
                name:
                    process.env.MAIL_FROM_NAME ||
                    "AVIS Reports",

                address:
                    process.env.MAIL_FROM_ADDRESS
            }
        );

    if (!from) {
        throw new Error(
            "Email sender is missing. Configure MAIL_FROM_ADDRESS."
        );
    }

    const to =
        normalizeRecipients(options.to);

    if (!to.length) {
        throw new Error(
            "No recipient email address was supplied."
        );
    }

    const cc =
        normalizeRecipients(options.cc);

    const bcc =
        normalizeRecipients(options.bcc);

    const replyTo =
        normalizeRecipients(
            options.replyTo ||
            options.reply_to
        );

    const payload = {
        from,
        to,
        subject:
            String(
                options.subject ||
                "AVIS Report"
            ),

        html:
            options.html !== undefined
                ? String(options.html)
                : undefined,

        text:
            options.text !== undefined
                ? String(options.text)
                : undefined
    };

    if (cc.length) {
        payload.cc = cc;
    }

    if (bcc.length) {
        payload.bcc = bcc;
    }

    if (replyTo.length) {
        payload.reply_to = replyTo;
    }

    const attachments =
        normalizeAttachments(
            options.attachments
        );

    if (attachments.length) {
        payload.attachments =
            attachments;
    }

    /*
     * Remove undefined properties before sending.
     */
    Object.keys(payload).forEach(key => {
        if (
            payload[key] === undefined ||
            payload[key] === null
        ) {
            delete payload[key];
        }
    });

    let response;

    try {
        response =
            await fetch(
                RESEND_API_URL,
                {
                    method: "POST",

                    headers: {
                        Authorization:
                            `Bearer ${apiKey}`,

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify(
                            payload
                        )
                }
            );
    }
    catch (error) {
        throw new Error(
            `Email API connection failed: ${
                error.message ||
                "Unable to connect to Resend."
            }`
        );
    }

    let result = {};

    try {
        result =
            await response.json();
    }
    catch (error) {
        /*
         * Keep a useful fallback if API response is not JSON.
         */
        result = {};
    }

    if (!response.ok) {
        const message =
            result &&
            (
                result.message ||
                result.error
            );

        throw new Error(
            message
            ||
            `Resend email API returned HTTP ${response.status}.`
        );
    }

    return {
        accepted: to,
        rejected: [],
        messageId:
            result.id ||
            null,
        response:
            result,
        envelope: {
            from,
            to
        }
    };
}

/*
 * Compatibility helper.
 *
 * Existing code may call transporter.verify().
 * With an HTTPS API there is no SMTP socket to verify,
 * so we validate the required configuration instead.
 */
async function verify() {
    getRequiredEnv("RESEND_API_KEY");

    const from =
        String(
            process.env.MAIL_FROM_ADDRESS ||
            ""
        ).trim();

    if (!from) {
        throw new Error(
            "MAIL_FROM_ADDRESS is missing from environment variables."
        );
    }

    return true;
}

/*
 * CRITICAL:
 * Export an object with sendMail() so your existing controller can keep:
 *
 *   const transporter = require("../config/mailConfig");
 *   await transporter.sendMail({...});
 */
module.exports = {
    sendMail,
    verify
};
