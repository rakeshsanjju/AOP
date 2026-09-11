const { google } = require("googleapis");

function getGoogleSheetsClient() {
    const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const privateKey = process.env.GOOGLE_PRIVATE_KEY;

    if (!email) {
        throw new Error("GOOGLE_SERVICE_ACCOUNT_EMAIL is missing in .env");
    }

    if (!privateKey) {
        throw new Error("GOOGLE_PRIVATE_KEY is missing in .env");
    }

    const auth = new google.auth.JWT({
        email,
        key: privateKey.replace(/\\n/g, "\n"),
        scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    });

    return google.sheets({
        version: "v4",
        auth
    });
}

module.exports = {
    getGoogleSheetsClient
};
