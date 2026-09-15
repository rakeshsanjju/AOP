

const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const transporter = require("../config/mailConfig");

/*
 * ============================================================
 * GOOGLE APPS SCRIPT SHARED SECRET
 * ============================================================
 *
 * Never hard-code this secret in source code.
 *
 * Configure the SAME value in:
 *
 * back-end/.env
 *   DRIVE_UPLOAD_SECRET=...
 *
 * Apps Script
 *   Project Settings -> Script Properties
 *   DRIVE_UPLOAD_SECRET=...
 */
function getAppsScriptSharedSecret() {

    const secret =
        String(
            process.env.DRIVE_UPLOAD_SECRET
            ||
            ""
        ).trim();


    if (
        !secret
    ) {

        throw new Error(
            "DRIVE_UPLOAD_SECRET is missing in back-end/.env. "
            +
            "Add DRIVE_UPLOAD_SECRET to .env and use the exact same value "
            +
            "in Apps Script Project Settings -> Script Properties."
        );

    }


    return secret;

}





const REPORT_CONFIGS = {

    panIndia: {
        key: "PAN_INDIA",
        routeKey: "pan-india",
        mode: "PAN_INDIA",
        title: "Pan India Performance Report",
        description:
            "Please find the Pan India performance report below. The complete PDF and Excel reports are attached to this email.",
        sheetName: "PAN_INDIA",
        driveFolderEnv: "PAN_INDIA_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Pan India Reports",
        /*
         * Recommended:
         * PAN_INDIA_RECIPIENTS=email1@example.com,email2@example.com
         *
         * Backward-compatible:
         * PAN_INDIA_RECIPIENT=email1@example.com
         * PAN_INDIA_RECIPIENT_2=email2@example.com
         */
        recipientsEnv: "PAN_INDIA_RECIPIENTS",
        recipientEnv: "PAN_INDIA_RECIPIENT",
        recipientEnv2: "PAN_INDIA_RECIPIENT_2",
        recipientDefault: "",
        filePrefix: "Pan-India-Performance-Report",
        excelSheetName: "Pan India Report"
    },

    leads: {
        key: "LEADS_APTS",
        liveInputKey: "APTS",
        routeKey: "leads",
        mode: "LEADS",
        title: "APTS Leads Performance Report",
        description:
            "Please find the AP & TS Leads performance report below. The PDF and Excel reports contain IP Count Target, IP Count Achieved, IP Count Need to Achieve, IP Revenue Target, IP Revenue Generated, IP Revenue Need to Achieve, Leads Target, Leads Achieved, Leads Need to Achieve, OP Target, OP Achieved, Lead to OP Conversion, OP Revenue and OP Revenue Need to Achieve.",
        sheetName: "LEADS_APTS",
        driveFolderEnv: "LEADS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS Leads Reports",
        recipientEnv: "LEADS_RECIPIENT",
        recipientsEnv: "LEADS_RECIPIENTS",
        recipientEnv2: "LEADS_RECIPIENT_2",
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "APTS-Leads-Performance-Report",
        excelSheetName: "APTS Leads Report"
    },

    leadsKa: {
        key: "LEADS_KA",
        liveInputKey: "KA",
        routeKey: "leads-ka",
        mode: "LEADS",
        title: "Karnataka Leads Performance Report",
        description:
            "Please find the Karnataka Leads performance report below. The PDF and Excel reports contain IP Count Target, IP Count Achieved, IP Count Need to Achieve, IP Revenue Target, IP Revenue Generated, IP Revenue Need to Achieve, Leads Target, Leads Achieved, Leads Need to Achieve, OP Target, OP Achieved, Lead to OP Conversion, OP Revenue and OP Revenue Need to Achieve.",
        sheetName: "LEADS_KA",
        driveFolderEnv: "KA_LEADS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Karnataka Leads Reports",
        recipientEnv: "KA_LEADS_RECIPIENT",
        recipientsEnv: "KA_LEADS_RECIPIENTS",
        recipientEnv2: "KA_LEADS_RECIPIENT_2",
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Karnataka-Leads-Performance-Report",
        excelSheetName: "Karnataka Leads Report"
    },

    leadsTn: {
        key: "LEADS_TN",
        liveInputKey: "TN",
        routeKey: "leads-tn",
        mode: "LEADS",
        title: "Tamil Nadu Leads Performance Report",
        description:
            "Please find the Tamil Nadu Leads performance report below. The PDF and Excel reports contain IP Count Target, IP Count Achieved, IP Count Need to Achieve, IP Revenue Target, IP Revenue Generated, IP Revenue Need to Achieve, Leads Target, Leads Achieved, Leads Need to Achieve, OP Target, OP Achieved, Lead to OP Conversion, OP Revenue and OP Revenue Need to Achieve.",
        sheetName: "LEADS_TN",
        driveFolderEnv: "TN_LEADS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Tamil Nadu Leads Reports",
        recipientEnv: "TN_LEADS_RECIPIENT",
        recipientsEnv: "TN_LEADS_RECIPIENTS",
        recipientEnv2: "TN_LEADS_RECIPIENT_2",
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Tamil-Nadu-Leads-Performance-Report",
        excelSheetName: "Tamil Nadu Leads Report"
    },

    leadsWb: {
        key: "LEADS_WB",
        liveInputKey: "WB",
        routeKey: "leads-wb",
        mode: "LEADS",
        title: "West Bengal Leads Performance Report",
        description:
            "Please find the West Bengal Leads performance report below. The PDF and Excel reports contain IP Count Target, IP Count Achieved, IP Count Need to Achieve, IP Revenue Target, IP Revenue Generated, IP Revenue Need to Achieve, Leads Target, Leads Achieved, Leads Need to Achieve, OP Target, OP Achieved, Lead to OP Conversion, OP Revenue and OP Revenue Need to Achieve.",
        sheetName: "LEADS_WB",
        driveFolderEnv: "WB_LEADS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "West Bengal Leads Reports",
        recipientEnv: "WB_LEADS_RECIPIENT",
        recipientsEnv: "WB_LEADS_RECIPIENTS",
        recipientEnv2: "WB_LEADS_RECIPIENT_2",
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "West-Bengal-Leads-Performance-Report",
        excelSheetName: "West Bengal Leads Report"
    },

    leadsMh: {
        key: "LEADS_MH",
        liveInputKey: "MH",
        routeKey: "leads-mh",
        mode: "LEADS",
        title: "Maharashtra Leads Performance Report",
        description:
            "Please find the Maharashtra Leads performance report below. The PDF and Excel reports contain IP Count Target, IP Count Achieved, IP Count Need to Achieve, IP Revenue Target, IP Revenue Generated, IP Revenue Need to Achieve, Leads Target, Leads Achieved, Leads Need to Achieve, OP Target, OP Achieved, Lead to OP Conversion, OP Revenue and OP Revenue Need to Achieve.",
        sheetName: "LEADS_MH",
        driveFolderEnv: "MH_LEADS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Maharashtra Leads Reports",
        recipientEnv: "MH_LEADS_RECIPIENT",
        recipientsEnv: "MH_LEADS_RECIPIENTS",
        recipientEnv2: "MH_LEADS_RECIPIENT_2",
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Maharashtra-Leads-Performance-Report",
        excelSheetName: "Maharashtra Leads Report"
    },

    labApts: {
        key: "LAB_APTS",
        liveInputKey: "APTS",
        routeKey: "lab-apts",
        mode: "LAB",
        sourceLabel: "APTS",
        title: "AP & TS Lab Performance Report",
        description: "Please find the AP & TS Lab performance report below. The PDF and Excel reports contain Lab Count Target, Achieved and Unachieved; Lab Revenue Target, Achieved and Need to Achieve; Per Lab Revenue Generated; and Discount Given Per Lab Revenue.",
        sheetName: "LAB_APTS",
        driveFolderEnv: "LAB_APTS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS Lab Reports",
        recipientEnv: "LAB_APTS_RECIPIENT",
        recipientsEnv: "LAB_APTS_RECIPIENTS",
        recipientEnv2: "LAB_APTS_RECIPIENT_2",
        recipientEnvs: [
            "APTS_LAB_RECIPIENT",
            "APTS_LAB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "AP-TS-Lab-Performance-Report",
        excelSheetName: "APTS Lab Report"
    },

    labKa: {
        key: "LAB_KA",
        liveInputKey: "KA",
        routeKey: "lab-ka",
        mode: "LAB",
        sourceLabel: "Karnataka",
        title: "Karnataka Lab Performance Report",
        description: "Please find the Karnataka Lab performance report below. The PDF and Excel reports contain Lab Count Target, Achieved and Unachieved; Lab Revenue Target, Achieved and Need to Achieve; Per Lab Revenue Generated; and Discount Given Per Lab Revenue.",
        sheetName: "LAB_KA",
        driveFolderEnv: "LAB_KA_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Karnataka Lab Reports",
        recipientEnv: "LAB_KA_RECIPIENT",
        recipientsEnv: "LAB_KA_RECIPIENTS",
        recipientEnv2: "LAB_KA_RECIPIENT_2",
        recipientEnvs: [
            "KA_LAB_RECIPIENT",
            "KA_LAB_RECIPIENTS",
            "KARNATAKA_LAB_RECIPIENT",
            "KARNATAKA_LAB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Karnataka-Lab-Performance-Report",
        excelSheetName: "Karnataka Lab Report"
    },

    labTn: {
        key: "LAB_TN",
        liveInputKey: "TN",
        routeKey: "lab-tn",
        mode: "LAB",
        sourceLabel: "Tamil Nadu",
        title: "Tamil Nadu Lab Performance Report",
        description: "Please find the Tamil Nadu Lab performance report below. The PDF and Excel reports contain Lab Count Target, Achieved and Unachieved; Lab Revenue Target, Achieved and Need to Achieve; Per Lab Revenue Generated; and Discount Given Per Lab Revenue.",
        sheetName: "LAB_TN",
        driveFolderEnv: "LAB_TN_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Tamil Nadu Lab Reports",
        recipientEnv: "LAB_TN_RECIPIENT",
        recipientsEnv: "LAB_TN_RECIPIENTS",
        recipientEnv2: "LAB_TN_RECIPIENT_2",
        recipientEnvs: [
            "TN_LAB_RECIPIENT",
            "TN_LAB_RECIPIENTS",
            "TAMIL_NADU_LAB_RECIPIENT",
            "TAMIL_NADU_LAB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Tamil-Nadu-Lab-Performance-Report",
        excelSheetName: "Tamil Nadu Lab Report"
    },

    labWb: {
        key: "LAB_WB",
        liveInputKey: "WB",
        routeKey: "lab-wb",
        mode: "LAB",
        sourceLabel: "West Bengal",
        title: "West Bengal Lab Performance Report",
        description: "Please find the West Bengal Lab performance report below. The PDF and Excel reports contain Lab Count Target, Achieved and Unachieved; Lab Revenue Target, Achieved and Need to Achieve; Per Lab Revenue Generated; and Discount Given Per Lab Revenue.",
        sheetName: "LAB_WB",
        driveFolderEnv: "LAB_WB_DRIVE_FOLDER_NAME",
        driveFolderDefault: "West Bengal Lab Reports",
        recipientEnv: "LAB_WB_RECIPIENT",
        recipientsEnv: "LAB_WB_RECIPIENTS",
        recipientEnv2: "LAB_WB_RECIPIENT_2",
        recipientEnvs: [
            "WB_LAB_RECIPIENT",
            "WB_LAB_RECIPIENTS",
            "WEST_BENGAL_LAB_RECIPIENT",
            "WEST_BENGAL_LAB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "West-Bengal-Lab-Performance-Report",
        excelSheetName: "West Bengal Lab Report"
    },

    labMh: {
        key: "LAB_MH",
        liveInputKey: "MH",
        routeKey: "lab-mh",
        mode: "LAB",
        sourceLabel: "Maharashtra",
        title: "Maharashtra Lab Performance Report",
        description: "Please find the Maharashtra Lab performance report below. The PDF and Excel reports contain Lab Count Target, Achieved and Unachieved; Lab Revenue Target, Achieved and Need to Achieve; Per Lab Revenue Generated; and Discount Given Per Lab Revenue.",
        sheetName: "LAB_MH",
        driveFolderEnv: "LAB_MH_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Maharashtra Lab Reports",
        recipientEnv: "LAB_MH_RECIPIENT",
        recipientsEnv: "LAB_MH_RECIPIENTS",
        recipientEnv2: "LAB_MH_RECIPIENT_2",
        recipientEnvs: [
            "MH_LAB_RECIPIENT",
            "MH_LAB_RECIPIENTS",
            "MAHARASHTRA_LAB_RECIPIENT",
            "MAHARASHTRA_LAB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Maharashtra-Lab-Performance-Report",
        excelSheetName: "Maharashtra Lab Report"
    },

    pharmacyApts: {
        key: "PHARMACY_APTS",
        liveInputKey: "APTS",
        routeKey: "pharmacy-apts",
        mode: "PHARMACY",
        sourceLabel: "APTS",
        title: "AP & TS Pharmacy Performance Report",
        description:
            "Please find the AP & TS Pharmacy performance report below. The PDF and Excel reports contain Pharmacy Bill Count Target, Received and Unreceived; Pharmacy Revenue Target, Achieved and Need to Achieve; Pharmacy Per Bill Revenue Target, Generated; and Discount Given Per Bill Revenue.",
        sheetName: "PHARMACY_APTS",
        driveFolderEnv: "PHARMACY_APTS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS Pharmacy Reports",
        recipientEnv: "PHARMACY_APTS_RECIPIENT",
        recipientsEnv: "PHARMACY_APTS_RECIPIENTS",
        recipientEnv2: "PHARMACY_APTS_RECIPIENT_2",
        recipientEnvs: [
            "APTS_PHARMACY_RECIPIENT",
            "APTS_PHARMACY_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "AP-TS-Pharmacy-Performance-Report",
        excelSheetName: "APTS Pharmacy Report"
    },

    pharmacyKa: {
        key: "PHARMACY_KA",
        liveInputKey: "KA",
        routeKey: "pharmacy-ka",
        mode: "PHARMACY",
        sourceLabel: "Karnataka",
        title: "Karnataka Pharmacy Performance Report",
        description:
            "Please find the Karnataka Pharmacy performance report below. The PDF and Excel reports contain Pharmacy Bill Count Target, Received and Unreceived; Pharmacy Revenue Target, Achieved and Need to Achieve; Pharmacy Per Bill Revenue Target, Generated; and Discount Given Per Bill Revenue.",
        sheetName: "PHARMACY_KA",
        driveFolderEnv: "PHARMACY_KA_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Karnataka Pharmacy Reports",
        recipientEnv: "PHARMACY_KA_RECIPIENT",
        recipientsEnv: "PHARMACY_KA_RECIPIENTS",
        recipientEnv2: "PHARMACY_KA_RECIPIENT_2",
        recipientEnvs: [
            "KA_PHARMACY_RECIPIENT",
            "KA_PHARMACY_RECIPIENTS",
            "KARNATAKA_PHARMACY_RECIPIENT",
            "KARNATAKA_PHARMACY_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Karnataka-Pharmacy-Performance-Report",
        excelSheetName: "Karnataka Pharmacy Report"
    },


    pharmacyTn: {
        key: "PHARMACY_TN",
        liveInputKey: "TN",
        routeKey: "pharmacy-tn",
        mode: "PHARMACY",
        sourceLabel: "Tamil Nadu",
        title: "Tamil Nadu Pharmacy Performance Report",
        description:
            "Please find the Tamil Nadu Pharmacy performance report below. The PDF and Excel reports contain Pharmacy Bill Count Target, Received and Unreceived; Pharmacy Revenue Target, Achieved and Need to Achieve; Pharmacy Per Bill Revenue Target, Generated; and Discount Given Per Bill Revenue.",
        sheetName: "PHARMACY_TN",
        driveFolderEnv: "PHARMACY_TN_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Tamil Nadu Pharmacy Reports",
        recipientEnv: "PHARMACY_TN_RECIPIENT",
        recipientsEnv: "PHARMACY_TN_RECIPIENTS",
        recipientEnv2: "PHARMACY_TN_RECIPIENT_2",
        recipientEnvs: [
            "TN_PHARMACY_RECIPIENT",
            "TN_PHARMACY_RECIPIENTS",
            "TAMIL_NADU_PHARMACY_RECIPIENT",
            "TAMIL_NADU_PHARMACY_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Tamil-Nadu-Pharmacy-Performance-Report",
        excelSheetName: "Tamil Nadu Pharmacy Report"
    },


    pharmacyWb: {
        key: "PHARMACY_WB",
        liveInputKey: "WB",
        routeKey: "pharmacy-wb",
        mode: "PHARMACY",
        sourceLabel: "West Bengal",
        title: "West Bengal Pharmacy Performance Report",
        description:
            "Please find the West Bengal Pharmacy performance report below. The PDF and Excel reports contain Pharmacy Bill Count Target, Received and Unreceived; Pharmacy Revenue Target, Achieved and Need to Achieve; Pharmacy Per Bill Revenue Target, Generated; and Discount Given Per Bill Revenue.",
        sheetName: "PHARMACY_WB",
        driveFolderEnv: "PHARMACY_WB_DRIVE_FOLDER_NAME",
        driveFolderDefault: "West Bengal Pharmacy Reports",
        recipientEnv: "PHARMACY_WB_RECIPIENT",
        recipientsEnv: "PHARMACY_WB_RECIPIENTS",
        recipientEnv2: "PHARMACY_WB_RECIPIENT_2",
        recipientEnvs: [
            "WB_PHARMACY_RECIPIENT",
            "WB_PHARMACY_RECIPIENTS",
            "WEST_BENGAL_PHARMACY_RECIPIENT",
            "WEST_BENGAL_PHARMACY_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "West-Bengal-Pharmacy-Performance-Report",
        excelSheetName: "West Bengal Pharmacy Report"
    },


    pharmacyMh: {
        key: "PHARMACY_MH",
        liveInputKey: "MH",
        routeKey: "pharmacy-mh",
        mode: "PHARMACY",
        sourceLabel: "Maharashtra",
        title: "Maharashtra Pharmacy Performance Report",
        description:
            "Please find the Maharashtra Pharmacy performance report below. The PDF and Excel reports contain Pharmacy Bill Count Target, Received and Unreceived; Pharmacy Revenue Target, Achieved and Need to Achieve; Pharmacy Per Bill Revenue Target, Generated; and Discount Given Per Bill Revenue.",
        sheetName: "PHARMACY_MH",
        driveFolderEnv: "PHARMACY_MH_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Maharashtra Pharmacy Reports",
        recipientEnv: "PHARMACY_MH_RECIPIENT",
        recipientsEnv: "PHARMACY_MH_RECIPIENTS",
        recipientEnv2: "PHARMACY_MH_RECIPIENT_2",
        recipientEnvs: [
            "MH_PHARMACY_RECIPIENT",
            "MH_PHARMACY_RECIPIENTS",
            "MAHARASHTRA_PHARMACY_RECIPIENT",
            "MAHARASHTRA_PHARMACY_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Maharashtra-Pharmacy-Performance-Report",
        excelSheetName: "Maharashtra Pharmacy Report"
    },



    hrApts: {
        key: "HR_APTS",
        liveInputKey: "APTS",
        routeKey: "hr-apts",
        mode: "HR",
        sourceLabel: "APTS",
        title: "AP & TS HR Performance Report",
        description:
            "Please find the AP & TS HR performance report below. The PDF and Excel reports contain position-wise Target, Achieved, Have To Achieve and hiring status for all HR positions.",
        sheetName: "HR_APTS",
        driveFolderEnv: "HR_APTS_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS HR Reports",
        recipientEnv: "HR_APTS_RECIPIENT",
        recipientsEnv: "HR_APTS_RECIPIENTS",
        recipientEnv2: "HR_APTS_RECIPIENT_2",
        recipientEnvs: [
            "APTS_HR_RECIPIENT",
            "APTS_HR_RECIPIENTS",
            "HR_APTS_EMAIL_RECIPIENT",
            "HR_APTS_EMAIL_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "AP-TS-HR-Performance-Report",
        excelSheetName: "APTS HR Report"
    },

    hrKa: {
        key: "HR_KA",
        liveInputKey: "KA",
        routeKey: "hr-ka",
        mode: "HR",
        sourceLabel: "Karnataka",
        title: "Karnataka HR Performance Report",
        description:
            "Please find the Karnataka HR performance report below. The PDF and Excel reports contain position-wise Target, Achieved, Have To Achieve and hiring status for all HR positions.",
        sheetName: "HR_KA",
        driveFolderEnv: "HR_KA_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Karnataka HR Reports",
        recipientEnv: "HR_KA_RECIPIENT",
        recipientsEnv: "HR_KA_RECIPIENTS",
        recipientEnv2: "HR_KA_RECIPIENT_2",
        recipientEnvs: [
            "KA_HR_RECIPIENT",
            "KA_HR_RECIPIENTS",
            "KARNATAKA_HR_RECIPIENT",
            "KARNATAKA_HR_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Karnataka-HR-Performance-Report",
        excelSheetName: "Karnataka HR Report"
    },

    hrTn: {
        key: "HR_TN",
        liveInputKey: "TN",
        routeKey: "hr-tn",
        mode: "HR",
        sourceLabel: "Tamil Nadu",
        title: "Tamil Nadu HR Performance Report",
        description:
            "Please find the Tamil Nadu HR performance report below. The PDF and Excel reports contain position-wise Target, Achieved, Have To Achieve and hiring status for all HR positions.",
        sheetName: "HR_TN",
        driveFolderEnv: "HR_TN_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Tamil Nadu HR Reports",
        recipientEnv: "HR_TN_RECIPIENT",
        recipientsEnv: "HR_TN_RECIPIENTS",
        recipientEnv2: "HR_TN_RECIPIENT_2",
        recipientEnvs: [
            "TN_HR_RECIPIENT",
            "TN_HR_RECIPIENTS",
            "TAMIL_NADU_HR_RECIPIENT",
            "TAMIL_NADU_HR_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Tamil-Nadu-HR-Performance-Report",
        excelSheetName: "Tamil Nadu HR Report"
    },


    hrWb: {
        key: "HR_WB",
        liveInputKey: "WB",
        routeKey: "hr-wb",
        mode: "HR",
        sourceLabel: "West Bengal",
        title: "West Bengal HR Performance Report",
        description:
            "Please find the West Bengal HR performance report below. The PDF and Excel reports contain position-wise Target, Achieved, Have To Achieve and hiring status for all HR positions.",
        sheetName: "HR_WB",
        driveFolderEnv: "HR_WB_DRIVE_FOLDER_NAME",
        driveFolderDefault: "West Bengal HR Reports",
        recipientEnv: "HR_WB_RECIPIENT",
        recipientDefault: "",
        filePrefix: "West-Bengal-HR-Performance-Report",
        excelSheetName: "West Bengal HR Report"
    },


    hrMh: {
        key: "HR_MH",
        liveInputKey: "MH",
        routeKey: "hr-mh",
        mode: "HR",
        sourceLabel: "Maharashtra",
        title: "Maharashtra HR Performance Report",
        description:
            "Please find the Maharashtra HR performance report below. The PDF and Excel reports contain position-wise Target, Achieved, Have To Achieve and hiring status for all HR positions.",
        sheetName: "HR_MH",
        driveFolderEnv: "HR_MH_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Maharashtra HR Reports",
        recipientEnv: "HR_MH_RECIPIENT",
        recipientsEnv: "HR_MH_RECIPIENTS",
        recipientEnv2: "HR_MH_RECIPIENT_2",
        recipientEnvs: [
            "MH_HR_RECIPIENT",
            "MH_HR_RECIPIENTS",
            "MAHARASHTRA_HR_RECIPIENT",
            "MAHARASHTRA_HR_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "Maharashtra-HR-Performance-Report",
        excelSheetName: "Maharashtra HR Report"
    },




    apts: {
        key: "APTS",
        routeKey: "apts",
        title: "AP & TS IP Performance Report",
        description:
            "Please find the AP & TS IP performance report below. The detailed PDF and Excel reports are attached to this email.",
        sheetName: "APTS",
        driveFolderEnv: "GOOGLE_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS Reports",
        recipientEnv: "APTS_IP_RECIPIENT",
        recipientsEnv: "APTS_IP_RECIPIENTS",
        recipientEnv2: "APTS_IP_RECIPIENT_2",
        recipientEnvs: [
            "IP_APTS_RECIPIENT",
            "IP_APTS_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "AP-TS-IP-Performance-Report",
        excelSheetName: "APTS Report"
    },

    opApts: {
        key: "OP_APTS",
        liveInputKey: "APTS",
        routeKey: "op-apts",
        mode: "OP",
        title: "AP & TS OP Performance Report",
        description:
            "Please find the AP & TS OP performance report below. The OP PDF and Excel reports are attached to this email.",
        sheetName: "OP_APTS",
        driveFolderEnv: "APTS_OP_DRIVE_FOLDER_NAME",
        driveFolderDefault: "APTS OP Reports",
        recipientEnv: "APTS_OP_RECIPIENT",
        recipientsEnv: "APTS_OP_RECIPIENTS",
        recipientEnv2: "APTS_OP_RECIPIENT_2",
        recipientEnvs: [
            "OP_APTS_RECIPIENT",
            "OP_APTS_RECIPIENTS"
        ],
        minimumRecipients: 1,
        recipientDefault: "",
        filePrefix: "AP-TS-OP-Performance-Report",
        excelSheetName: "APTS OP Report"
    },

    opKa: {
        key: "OP_KA",
        liveInputKey: "KA",
        routeKey: "op-ka",
        mode: "OP",
        title: "Karnataka OP Performance Report",
        description:
            "Please find the Karnataka OP performance report below. The OP PDF and Excel reports are attached to this email.",
        sheetName: "OP_KA",
        driveFolderEnv: "KA_OP_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Karnataka OP Reports",
        recipientEnv: "KA_OP_RECIPIENT",
        recipientsEnv: "KA_OP_RECIPIENTS",
        recipientEnv2: "KA_OP_RECIPIENT_2",
        recipientEnvs: [
            "OP_KA_RECIPIENT",
            "OP_KA_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "KA-OP-Performance-Report",
        excelSheetName: "KA OP Report"
    },

    opTn: {
        key: "OP_TN",
        liveInputKey: "TN",
        routeKey: "op-tn",
        mode: "OP",
        title: "Tamil Nadu OP Performance Report",
        description:
            "Please find the Tamil Nadu OP performance report below. The OP PDF and Excel reports are attached to this email.",
        sheetName: "OP_TN",
        driveFolderEnv: "TN_OP_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Tamil Nadu OP Reports",
        recipientEnv: "TN_OP_RECIPIENT",
        recipientsEnv: "TN_OP_RECIPIENTS",
        recipientEnv2: "TN_OP_RECIPIENT_2",
        recipientEnvs: [
            "OP_TN_RECIPIENT",
            "OP_TN_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "TN-OP-Performance-Report",
        excelSheetName: "TN OP Report"
    },

    opWb: {
        key: "OP_WB",
        liveInputKey: "WB",
        routeKey: "op-wb",
        mode: "OP",
        title: "West Bengal OP Performance Report",
        description:
            "Please find the West Bengal OP performance report below. The OP PDF and Excel reports are attached to this email.",
        sheetName: "OP_WB",
        driveFolderEnv: "WB_OP_DRIVE_FOLDER_NAME",
        driveFolderDefault: "West Bengal OP Reports",
        recipientEnv: "WB_OP_RECIPIENT",
        recipientsEnv: "WB_OP_RECIPIENTS",
        recipientEnv2: "WB_OP_RECIPIENT_2",
        recipientEnvs: [
            "OP_WB_RECIPIENT",
            "OP_WB_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "WB-OP-Performance-Report",
        excelSheetName: "WB OP Report"
    },

    opMh: {
        key: "OP_MH",
        liveInputKey: "MH",
        routeKey: "op-mh",
        mode: "OP",
        title: "Maharashtra OP Performance Report",
        description:
            "Please find the Maharashtra OP performance report below. The OP PDF and Excel reports are attached to this email.",
        sheetName: "OP_MH",
        driveFolderEnv: "MH_OP_DRIVE_FOLDER_NAME",
        driveFolderDefault: "Maharashtra OP Reports",
        recipientEnv: "MH_OP_RECIPIENT",
        recipientsEnv: "MH_OP_RECIPIENTS",
        recipientEnv2: "MH_OP_RECIPIENT_2",
        recipientEnvs: [
            "OP_MH_RECIPIENT",
            "OP_MH_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "MH-OP-Performance-Report",
        excelSheetName: "MH OP Report"
    },

    ka: {
        key: "KA",
        routeKey: "ka",
        title: "Karnataka IP Performance Report",
        description:
            "Please find the Karnataka IP performance report below. The detailed PDF and Excel reports are attached to this email.",
        sheetName: "KA",
        driveFolderEnv: "KA_DRIVE_FOLDER_NAME",
        driveFolderDefault: "KA Reports",
        recipientEnv: "KA_IP_RECIPIENT",
        recipientDefault: "sanjurakeshdasari4@gmail.com",
        filePrefix: "KA-IP-Performance-Report",
        excelSheetName: "KA Report"
    },


    tn: {
        key: "TN",
        routeKey: "tn",
        title: "Tamil Nadu IP Performance Report",
        description:
            "Please find the Tamil Nadu IP performance report below. The detailed PDF and Excel reports are attached to this email.",
        sheetName: "TN",
        driveFolderEnv: "TN_DRIVE_FOLDER_NAME",
        driveFolderDefault: "TN Reports",
        recipientEnv: "TN_IP_RECIPIENT",
        recipientDefault: "",
        filePrefix: "TN-IP-Performance-Report",
        excelSheetName: "TN Report"
    },


    wb: {
        key: "WB",
        routeKey: "wb",
        title: "West Bengal IP Performance Report",
        description:
            "Please find the West Bengal IP performance report below. The detailed PDF and Excel reports are attached to this email.",
        sheetName: "WB",
        driveFolderEnv: "WB_DRIVE_FOLDER_NAME",
        driveFolderDefault: "WB Reports",
        recipientEnv: "WB_IP_RECIPIENT",
        recipientsEnv: "WB_IP_RECIPIENTS",
        recipientEnv2: "WB_IP_RECIPIENT_2",
        recipientEnvs: [
            "WEST_BENGAL_IP_RECIPIENT",
            "WEST_BENGAL_IP_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "WB-IP-Performance-Report",
        excelSheetName: "WB Report"
    },


    mh: {
        key: "MH",
        routeKey: "mh",
        title: "Maharashtra IP Performance Report",
        description:
            "Please find the Maharashtra IP performance report below. The detailed PDF and Excel reports are attached to this email.",
        sheetName: "MH",
        driveFolderEnv: "MH_DRIVE_FOLDER_NAME",
        driveFolderDefault: "MH Reports",
        recipientEnv: "MH_IP_RECIPIENT",
        recipientsEnv: "MH_IP_RECIPIENTS",
        recipientEnv2: "MH_IP_RECIPIENT_2",
        recipientEnvs: [
            "MAHARASHTRA_IP_RECIPIENT",
            "MAHARASHTRA_IP_RECIPIENTS"
        ],
        minimumRecipients: 2,
        recipientDefault: "",
        filePrefix: "MH-IP-Performance-Report",
        excelSheetName: "MH Report"
    }

};


const METRICS = [
    {
        key: "ipCount",
        label: "IP Count",
        percent: false
    },
    {
        key: "b2bCount",
        label: "B2B Count",
        percent: false
    },
    {
        key: "opCount",
        label: "OP Count",
        percent: false
    },
    {
        key: "leads",
        label: "Leads",
        percent: false
    },
    {
        key: "opToIp",
        label: "OP to IP Conversion",
        percent: true
    },
    {
        key: "leadToOp",
        label: "Lead to OP Conversion",
        percent: true
    }
];


/* ============================================================
   HELPERS
============================================================ */

function validNumber(value) {

    return (
        value !== undefined &&
        value !== null &&
        value !== "" &&
        Number.isFinite(
            Number(value)
        )
    );

}


function formatValue(
    value,
    percent = false
) {

    if (
        !validNumber(value)
    ) {
        return "N/A";
    }

    const number =
        Number(value);

    if (
        percent
    ) {
        return `${number.toFixed(2)}%`;
    }

    return new Intl.NumberFormat(
        "en-US",
        {
            maximumFractionDigits: 2
        }
    ).format(number);

}


function calculateStatus(
    target,
    achieved
) {

    if (
        !validNumber(target) ||
        !validNumber(achieved)
    ) {
        return "Data Missing";
    }

    return (
        Number(achieved) >=
        Number(target)
    )
        ?
        "Target Achieved"
        :
        "Target Not Achieved";

}


function getStatusStyle(
    status
) {

    if (
        status ===
        "Target Achieved"
    ) {

        return {
            emailBg: "#dcfce7",
            emailText: "#15803d",
            pdfBg: "#DCFCE7",
            pdfText: "#15803D",
            label: "Target Achieved"
        };

    }


    if (
        status ===
        "Target Not Achieved"
    ) {

        return {
            emailBg: "#fee2e2",
            emailText: "#b91c1c",
            pdfBg: "#FEE2E2",
            pdfText: "#B91C1C",
            label: "Target Not Achieved"
        };

    }


    if (
        status ===
        "Within Target Cost"
    ) {

        return {
            emailBg: "#dcfce7",
            emailText: "#15803d",
            pdfBg: "#DCFCE7",
            pdfText: "#15803D",
            label: "Within Target Cost"
        };

    }


    if (
        status ===
        "Above Target Cost"
    ) {

        return {
            emailBg: "#fee2e2",
            emailText: "#b91c1c",
            pdfBg: "#FEE2E2",
            pdfText: "#B91C1C",
            label: "Above Target Cost"
        };

    }


    if (
        status ===
        "Calculated"
    ) {

        return {
            emailBg: "#dbeafe",
            emailText: "#1d4ed8",
            pdfBg: "#DBEAFE",
            pdfText: "#1D4ED8",
            label: "Calculated"
        };

    }


    return {
        emailBg: "#fff7ed",
        emailText: "#c2410c",
        pdfBg: "#FFF7ED",
        pdfText: "#C2410C",
        label: "Data Missing"
    };

}


/* ============================================================
   CURRENCY / REVENUE HELPERS
============================================================ */

function formatCurrency(value) {
    if (!validNumber(value)) {
        return "N/A";
    }

    return (
        "Rs. "
        +
        new Intl.NumberFormat(
            "en-IN",
            {
                maximumFractionDigits: 2
            }
        ).format(Number(value))
    );
}


function calculateCostStatus(target, achieved) {
    if (!validNumber(target) || !validNumber(achieved)) {
        return "Data Missing";
    }

    return Number(achieved) <= Number(target)
        ? "Within Target Cost"
        : "Above Target Cost";
}


function calculatedStatus(value) {
    return validNumber(value)
        ? "Calculated"
        : "Data Missing";
}


function numberOrNull(
    value
) {

    if (
        value === undefined
        ||
        value === null
        ||
        value === ""
    ) {

        return null;

    }


    if (
        typeof value === "number"
    ) {

        return Number.isFinite(value)
            ?
            value
            :
            null;

    }


    const cleaned =
        String(value)
            .replace(/,/g, "")
            .replace(/%/g, "")
            .replace(/[^0-9.+-]/g, "");


    if (!cleaned) {
        return null;
    }


    const number =
        Number(cleaned);


    return Number.isFinite(number)
        ?
        number
        :
        null;

}


function firstNumber(
    ...values
) {

    for (
        const value
        of values
    ) {

        const number =
            numberOrNull(value);


        if (
            number !== null
        ) {

            return number;

        }

    }


    return null;

}


function normalizedMetric(
    metric,
    fallback = {}
) {

    const source =
        metric &&
        typeof metric === "object"
            ?
            metric
            :
            {};


    const backup =
        fallback &&
        typeof fallback === "object"
            ?
            fallback
            :
            {};


    return {

        target:
            firstNumber(
                source.target,
                backup.target
            ),

        achieved:
            firstNumber(
                source.achieved,
                source.reached,
                backup.achieved,
                backup.reached
            )

    };

}


function divideOrNull(
    numerator,
    denominator
) {

    const top =
        numberOrNull(numerator);


    const bottom =
        numberOrNull(denominator);


    if (
        top === null
        ||
        bottom === null
        ||
        bottom === 0
    ) {

        return null;

    }


    return top / bottom;

}


function normalizeCompleteReportData(
    inputData = {}
) {

    const data =
        inputData &&
        typeof inputData === "object"
            ?
            inputData
            :
            {};


    const parameters =
        data.parameters &&
        typeof data.parameters === "object"
            ?
            data.parameters
            :
            {};


    const revenueContainer =
        data.revenue &&
        typeof data.revenue === "object"
            ?
            data.revenue
            :
            {};


    const normalized = {
        ...data
    };


    /*
     * Core report metrics.
     */
    METRICS.forEach(
        metric => {

            normalized[metric.key] =
                normalizedMetric(
                    data[metric.key]
                );

        }
    );


    /*
     * IP Revenue.
     * "IP Revenue Generated" = achieved revenue.
     */
    normalized.ipRevenue =
        normalizedMetric(
            data.ipRevenue,
            revenueContainer.ipRevenue
        );


    normalized.ipRevenue.target =
        firstNumber(
            normalized.ipRevenue.target,
            data.ipRevenueTarget,
            data["IP Revenue Target"],
            parameters["IP Revenue Target"]
        );


    normalized.ipRevenue.achieved =
        firstNumber(
            normalized.ipRevenue.achieved,
            data.ipRevenueAchieved,
            data.ipRevenueGenerated,
            data["IP Revenue Generated"],
            data["IP Revenue Achieved"],
            parameters["IP Revenue Generated"],
            parameters["IP Revenue Achieved"]
        );


    /*
     * OP Revenue.
     */
    normalized.opRevenue =
        normalizedMetric(
            data.opRevenue,
            revenueContainer.opRevenue
        );


    normalized.opRevenue.target =
        firstNumber(
            normalized.opRevenue.target,
            data.opRevenueTarget,
            data["OP Revenue Target"],
            parameters["OP Revenue Target"]
        );


    normalized.opRevenue.achieved =
        firstNumber(
            normalized.opRevenue.achieved,
            data.opRevenueAchieved,
            data["OP Revenue Achieved"],
            parameters["OP Revenue Achieved"]
        );


    /*
     * IP cost/person.
     */
    normalized.ipCostPerPerson =
        normalizedMetric(
            data.ipCostPerPerson
        );


    normalized.ipCostPerPerson.target =
        firstNumber(
            normalized.ipCostPerPerson.target,

            divideOrNull(
                normalized.ipRevenue.target,
                normalized.ipCount.target
            )
        );


    normalized.ipCostPerPerson.achieved =
        firstNumber(
            normalized.ipCostPerPerson.achieved,

            divideOrNull(
                normalized.ipRevenue.achieved,
                normalized.ipCount.achieved
            )
        );


    /*
     * Discount/person.
     */
    const discountSource =
        data.discountPerPerson &&
        typeof data.discountPerPerson === "object"
            ?
            data.discountPerPerson
            :
            {};


    let discount =
        firstNumber(
            discountSource.achieved,
            discountSource.value,
            data.discountPerPersonValue,
            data["Discount Given Per Person"]
        );


    if (
        discount === null
        &&
        validNumber(
            normalized.ipCostPerPerson.target
        )
        &&
        validNumber(
            normalized.ipCostPerPerson.achieved
        )
    ) {

        discount =
            Math.max(
                0,

                Number(
                    normalized.ipCostPerPerson.target
                )
                -
                Number(
                    normalized.ipCostPerPerson.achieved
                )
            );

    }


    normalized.discountPerPerson = {

        target:
            null,

        achieved:
            discount,

        value:
            discount

    };


    return normalized;

}


function getAdditionalReportRows(
    inputData = {}
) {

    const data =
        normalizeCompleteReportData(
            inputData
        );


    const discountValue =
        data.discountPerPerson.achieved;


    return [

        {
            particular: "IP Revenue",

            target:
                formatCurrency(
                    data.ipRevenue.target
                ),

            achieved:
                formatCurrency(
                    data.ipRevenue.achieved
                ),

            result:
                calculateStatus(
                    data.ipRevenue.target,
                    data.ipRevenue.achieved
                )
        },


        {
            particular: "OP Revenue",

            target:
                formatCurrency(
                    data.opRevenue.target
                ),

            achieved:
                formatCurrency(
                    data.opRevenue.achieved
                ),

            result:
                calculateStatus(
                    data.opRevenue.target,
                    data.opRevenue.achieved
                )
        },


        {
            particular: "IP Cost Per Person",

            target:
                formatCurrency(
                    data.ipCostPerPerson.target
                ),

            achieved:
                formatCurrency(
                    data.ipCostPerPerson.achieved
                ),

            result:
                calculateCostStatus(
                    data.ipCostPerPerson.target,
                    data.ipCostPerPerson.achieved
                )
        },


        {
            particular:
                "Discount Given Per Person",

            target:
                "N/A",

            achieved:
                formatCurrency(
                    discountValue
                ),

            result:
                calculatedStatus(
                    discountValue
                )
        }

    ];

}


function getReportSummary(inputData = {}, config = null) {
    const data=normalizeCompleteReportData(inputData);
    if(config && config.mode === "OP") {
        const revenueNeed=reportNeedToAchieve(data.opRevenue.target,data.opRevenue.achieved);
        const leadsNeed=reportNeedToAchieve(data.leads.target,data.leads.achieved);
        return [
            "OP Count: "+formatValue(data.opCount.target,false)+" / "+formatValue(data.opCount.achieved,false),
            "OP Revenue: "+formatCurrency(data.opRevenue.target)+" / "+formatCurrency(data.opRevenue.achieved),
            "OP Revenue Need: "+formatCurrency(revenueNeed),
            "Leads: "+formatValue(data.leads.target,false)+" / "+formatValue(data.leads.achieved,false),
            "Leads Need: "+formatValue(leadsNeed,false),
            "Lead to OP: "+formatValue(data.leadToOp.target,true)+" / "+formatValue(data.leadToOp.achieved,true)
        ].join(" | ");
    }
    if(config && config.mode === "LEADS") {
        const ipNeed=reportNeedToAchieve(data.ipCount.target,data.ipCount.achieved);
        const ipRevenueNeed=reportNeedToAchieve(data.ipRevenue.target,data.ipRevenue.achieved);
        const leadsNeed=reportNeedToAchieve(data.leads.target,data.leads.achieved);
        const opRevenueNeed=reportNeedToAchieve(data.opRevenue.target,data.opRevenue.achieved);
        const leadsLabel=String(config.title || "Leads").replace(/ Performance Report$/i,"");

        return [
            "IP Count: "+formatValue(data.ipCount.target,false)+" / "+formatValue(data.ipCount.achieved,false),
            "IP Count Need: "+formatValue(ipNeed,false),
            "IP Revenue: "+formatCurrency(data.ipRevenue.target)+" / "+formatCurrency(data.ipRevenue.achieved),
            "IP Revenue Need: "+formatCurrency(ipRevenueNeed),
            leadsLabel+": "+formatValue(data.leads.target,false)+" / "+formatValue(data.leads.achieved,false),
            "Leads Need: "+formatValue(leadsNeed,false),
            "OP Count: "+formatValue(data.opCount.target,false)+" / "+formatValue(data.opCount.achieved,false),
            "Lead to OP: "+formatValue(data.leadToOp.target,true)+" / "+formatValue(data.leadToOp.achieved,true),
            "OP Revenue: "+formatCurrency(data.opRevenue.target)+" / "+formatCurrency(data.opRevenue.achieved),
            "OP Revenue Need: "+formatCurrency(opRevenueNeed)
        ].join(" | ");
    }
    if(config && config.mode === "LAB") {
        const labCount=data.labCount&&typeof data.labCount==="object"?data.labCount:{};
        const labRevenue=data.labRevenue&&typeof data.labRevenue==="object"?data.labRevenue:{};
        const perLabRevenue=data.perLabRevenue&&typeof data.perLabRevenue==="object"?data.perLabRevenue:{};
        return [
            "Lab Count: "+formatValue(labCount.target,false)+" / "+formatValue(labCount.achieved,false),
            "Lab Count Unachieved: "+formatValue(labCount.unachieved,false),
            "Lab Revenue: "+formatCurrency(labRevenue.target)+" / "+formatCurrency(labRevenue.achieved),
            "Lab Revenue Need to Achieve: "+formatCurrency(firstNumber(labRevenue.needToAchieve,labRevenue.unachieved)),
            "Per Lab Revenue Generated: "+formatCurrency(perLabRevenue.generated),
            "Discount Given Per Lab Revenue: "+formatCurrency(perLabRevenue.discount)
        ].join(" | ");
    }
    if(config && config.mode === "PHARMACY") {
        const pharmacyCount=data.pharmacyCount&&typeof data.pharmacyCount==="object"?data.pharmacyCount:{};
        const pharmacyRevenue=data.pharmacyRevenue&&typeof data.pharmacyRevenue==="object"?data.pharmacyRevenue:{};
        const perBill=data.pharmacyPerBillRevenue&&typeof data.pharmacyPerBillRevenue==="object"?data.pharmacyPerBillRevenue:{};

        return [
            "Pharmacy Bill Count: "+formatValue(pharmacyCount.target,false)+" / "+formatValue(pharmacyCount.achieved,false),
            "Pharmacy Bill Count Unreceived: "+formatValue(pharmacyCount.unreceived,false),
            "Pharmacy Revenue: "+formatCurrency(pharmacyRevenue.target)+" / "+formatCurrency(pharmacyRevenue.achieved),
            "Pharmacy Revenue Need to Achieve: "+formatCurrency(pharmacyRevenue.needToAchieve),
            "Pharmacy Per Bill Revenue Target: "+formatCurrency(perBill.target),
            "Pharmacy Per Bill Revenue Generated: "+formatCurrency(perBill.generated),
            "Discount Given Per Bill Revenue: "+formatCurrency(perBill.discount)
        ].join(" | ");
    }
    if(config && config.mode === "HR") {
        const positions=data.positions&&typeof data.positions==="object"?data.positions:{};
        const totals=data.totals&&typeof data.totals==="object"?data.totals:{};
        const items=Object.values(positions);
        const positionCount=firstNumber(totals.positionCount,items.length);
        const totalTarget=firstNumber(totals.target,items.reduce((sum,item)=>sum+(validNumber(item&&item.target)?Number(item.target):0),0));
        const totalAchieved=firstNumber(totals.achieved,items.reduce((sum,item)=>sum+(validNumber(item&&item.achieved)?Number(item.achieved):0),0));
        const totalRemaining=firstNumber(totals.remaining,items.reduce((sum,item)=>sum+(validNumber(item&&item.haveToAchieve)?Number(item.haveToAchieve):0),0));
        const positionsAchieved=firstNumber(totals.positionsAchieved,items.filter(item=>validNumber(item&&item.target)&&validNumber(item&&item.achieved)&&Number(item.achieved)>=Number(item.target)).length);
        const positionsPending=firstNumber(totals.positionsPending,items.filter(item=>validNumber(item&&item.target)&&validNumber(item&&item.achieved)&&Number(item.target)>0&&Number(item.achieved)<Number(item.target)).length);
        return [
            "Positions: "+formatValue(positionCount,false),
            "Total Target: "+formatValue(totalTarget,false),
            "Total Achieved: "+formatValue(totalAchieved,false),
            "Have To Achieve: "+formatValue(totalRemaining,false),
            "Positions Achieved: "+formatValue(positionsAchieved,false),
            "Hiring Pending: "+formatValue(positionsPending,false)
        ].join(" | ");
    }
    let achieved=0,failed=0,missing=0;
    METRICS.forEach(metric=>{const item=data[metric.key]; const status=calculateStatus(item.target,item.achieved); if(status==="Target Achieved") achieved++; else if(status==="Target Not Achieved") failed++; else missing++;});
    return [
        `${achieved} Achieved / ${failed} Not Achieved`+(missing>0?` / ${missing} Missing`:""),
        "IP Revenue: "+formatCurrency(data.ipRevenue.target)+" / "+formatCurrency(data.ipRevenue.achieved),
        "OP Revenue: "+formatCurrency(data.opRevenue.target)+" / "+formatCurrency(data.opRevenue.achieved),
        "IP Cost/Person: "+formatCurrency(data.ipCostPerPerson.target)+" / "+formatCurrency(data.ipCostPerPerson.achieved),
        "Discount/Person: "+formatCurrency(data.discountPerPerson.achieved)
    ].join(" | ");
}


function reportNeedToAchieve(target, achieved) {
    if (!validNumber(target) || !validNumber(achieved)) {
        return null;
    }

    /*
     * Never show a negative "Need to Achieve".
     * Once the target is achieved/exceeded, remaining = 0.
     */
    return Math.max(
        0,
        Number(target) - Number(achieved)
    );
}

function safeFileName(value) {

    return String(
        value ||
        "Monthly-Report"
    )
        .trim()
        .replace(
            /[<>:"/\\|?*]+/g,
            "-"
        )
        .replace(
            /\s+/g,
            "-"
        );

}


function escapeHtml(value) {

    return String(
        value ?? ""
    )
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

}


function buildReportRows(report, config = null) {
    const data=normalizeCompleteReportData(report.data || {});
    if(config && config.mode === "OP") {
        const opRevenueNeed=reportNeedToAchieve(data.opRevenue.target,data.opRevenue.achieved);
        const leadsNeed=reportNeedToAchieve(data.leads.target,data.leads.achieved);
        return [
            {particular:"OP Count",target:formatValue(data.opCount.target,false),achieved:formatValue(data.opCount.achieved,false),result:calculateStatus(data.opCount.target,data.opCount.achieved)},
            {particular:"OP Revenue",target:formatCurrency(data.opRevenue.target),achieved:formatCurrency(data.opRevenue.achieved),result:calculateStatus(data.opRevenue.target,data.opRevenue.achieved)},
            {particular:"OP Revenue Need to Achieve",target:"N/A",achieved:formatCurrency(opRevenueNeed),result:calculateStatus(data.opRevenue.target,data.opRevenue.achieved)},
            {particular:"Leads",target:formatValue(data.leads.target,false),achieved:formatValue(data.leads.achieved,false),result:calculateStatus(data.leads.target,data.leads.achieved)},
            {particular:"Leads Need to Achieve",target:"N/A",achieved:formatValue(leadsNeed,false),result:calculatedStatus(leadsNeed)},
            {particular:"Lead to OP Conversion",target:formatValue(data.leadToOp.target,true),achieved:formatValue(data.leadToOp.achieved,true),result:calculateStatus(data.leadToOp.target,data.leadToOp.achieved)}
        ];
    }
    if(config && config.mode === "LEADS") {
        const ipNeed=reportNeedToAchieve(data.ipCount.target,data.ipCount.achieved);
        const ipRevenueNeed=reportNeedToAchieve(data.ipRevenue.target,data.ipRevenue.achieved);
        const leadsNeed=reportNeedToAchieve(data.leads.target,data.leads.achieved);
        const opRevenueNeed=reportNeedToAchieve(data.opRevenue.target,data.opRevenue.achieved);

        return [
            {particular:"IP Count",target:formatValue(data.ipCount.target,false),achieved:formatValue(data.ipCount.achieved,false),result:calculateStatus(data.ipCount.target,data.ipCount.achieved)},
            {particular:"IP Count Need to Achieve",target:"N/A",achieved:formatValue(ipNeed,false),result:calculateStatus(data.ipCount.target,data.ipCount.achieved)},
            {particular:"IP Revenue",target:formatCurrency(data.ipRevenue.target),achieved:formatCurrency(data.ipRevenue.achieved),result:calculateStatus(data.ipRevenue.target,data.ipRevenue.achieved)},
            {particular:"IP Revenue Need to Achieve",target:"N/A",achieved:formatCurrency(ipRevenueNeed),result:calculateStatus(data.ipRevenue.target,data.ipRevenue.achieved)},
            {particular:"Leads",target:formatValue(data.leads.target,false),achieved:formatValue(data.leads.achieved,false),result:calculateStatus(data.leads.target,data.leads.achieved)},
            {particular:"Leads Need to Achieve",target:"N/A",achieved:formatValue(leadsNeed,false),result:calculateStatus(data.leads.target,data.leads.achieved)},
            {particular:"OP Count",target:formatValue(data.opCount.target,false),achieved:formatValue(data.opCount.achieved,false),result:calculateStatus(data.opCount.target,data.opCount.achieved)},
            {particular:"Lead to OP Conversion",target:formatValue(data.leadToOp.target,true),achieved:formatValue(data.leadToOp.achieved,true),result:calculateStatus(data.leadToOp.target,data.leadToOp.achieved)},
            {particular:"OP Revenue",target:formatCurrency(data.opRevenue.target),achieved:formatCurrency(data.opRevenue.achieved),result:calculateStatus(data.opRevenue.target,data.opRevenue.achieved)},
            {particular:"OP Revenue Need to Achieve",target:"N/A",achieved:formatCurrency(opRevenueNeed),result:calculateStatus(data.opRevenue.target,data.opRevenue.achieved)}
        ];
    }
    if(config && config.mode === "LAB") {
        const labCount=data.labCount&&typeof data.labCount==="object"?data.labCount:{};
        const labRevenue=data.labRevenue&&typeof data.labRevenue==="object"?data.labRevenue:{};
        const perLabRevenue=data.perLabRevenue&&typeof data.perLabRevenue==="object"?data.perLabRevenue:{};
        return [
            {particular:"Lab Count",target:formatValue(labCount.target,false),achieved:formatValue(labCount.achieved,false),result:calculateStatus(labCount.target,labCount.achieved)},
            {particular:"Lab Count Unachieved",target:"N/A",achieved:formatValue(labCount.unachieved,false),result:calculatedStatus(labCount.unachieved)},
            {particular:"Lab Revenue",target:formatCurrency(labRevenue.target),achieved:formatCurrency(labRevenue.achieved),result:calculateStatus(labRevenue.target,labRevenue.achieved)},
            {particular:"Lab Revenue Need to Achieve",target:"N/A",achieved:formatCurrency(firstNumber(labRevenue.needToAchieve,labRevenue.unachieved)),result:calculatedStatus(firstNumber(labRevenue.needToAchieve,labRevenue.unachieved))},
            {particular:"Per Lab Revenue Generated",target:"N/A",achieved:formatCurrency(perLabRevenue.generated),result:calculatedStatus(perLabRevenue.generated)},
            {particular:"Discount Given Per Lab Revenue",target:"N/A",achieved:formatCurrency(perLabRevenue.discount),result:calculatedStatus(perLabRevenue.discount)}
        ];
    }
    if(config && config.mode === "PHARMACY") {
        const pharmacyCount=data.pharmacyCount&&typeof data.pharmacyCount==="object"?data.pharmacyCount:{};
        const pharmacyRevenue=data.pharmacyRevenue&&typeof data.pharmacyRevenue==="object"?data.pharmacyRevenue:{};
        const perBill=data.pharmacyPerBillRevenue&&typeof data.pharmacyPerBillRevenue==="object"?data.pharmacyPerBillRevenue:{};

        return [
            {particular:"Pharmacy Bill Count",target:formatValue(pharmacyCount.target,false),achieved:formatValue(pharmacyCount.achieved,false),result:calculateStatus(pharmacyCount.target,pharmacyCount.achieved)},
            {particular:"Pharmacy Bill Count Unreceived",target:"N/A",achieved:formatValue(pharmacyCount.unreceived,false),result:calculatedStatus(pharmacyCount.unreceived)},
            {particular:"Pharmacy Revenue",target:formatCurrency(pharmacyRevenue.target),achieved:formatCurrency(pharmacyRevenue.achieved),result:calculateStatus(pharmacyRevenue.target,pharmacyRevenue.achieved)},
            {particular:"Pharmacy Revenue Need to Achieve",target:"N/A",achieved:formatCurrency(pharmacyRevenue.needToAchieve),result:calculatedStatus(pharmacyRevenue.needToAchieve)},
            {particular:"Pharmacy Per Bill Revenue Target",target:"N/A",achieved:formatCurrency(perBill.target),result:calculatedStatus(perBill.target)},
            {particular:"Pharmacy Per Bill Revenue Generated",target:"N/A",achieved:formatCurrency(perBill.generated),result:calculatedStatus(perBill.generated)},
            {particular:"Discount Given Per Bill Revenue",target:"N/A",achieved:formatCurrency(perBill.discount),result:calculatedStatus(perBill.discount)}
        ];
    }
    if(config && config.mode === "HR") {
        const positions=data.positions&&typeof data.positions==="object"?data.positions:{};
        const order=["telecallers","telesales","surgeons","counsellors","fieldRepresentative","frontOffice","pharmacist","mbbsDoctor","houseKeeping"];
        const rows=[];
        order.forEach(key=>{
            const item=positions[key];
            if(!item)return;
            rows.push({
                particular:item.label||key,
                target:formatValue(item.target,false),
                achieved:formatValue(item.achieved,false),
                result:calculateStatus(item.target,item.achieved)
            });
            rows.push({
                particular:(item.label||key)+" - Have To Achieve",
                target:"N/A",
                achieved:formatValue(item.haveToAchieve,false),
                result:calculatedStatus(item.haveToAchieve)
            });
        });
        return rows;
    }
    const coreRows=METRICS.map(metric=>{const item=data[metric.key]; return {particular:metric.label,target:formatValue(item.target,metric.percent),achieved:formatValue(item.achieved,metric.percent),result:calculateStatus(item.target,item.achieved)};});
    return coreRows.concat(getAdditionalReportRows(data));
}


function getDriveFolderName(
    config
) {

    return String(
        process.env[
            config.driveFolderEnv
        ]
        ||
        config.driveFolderDefault
    ).trim();

}


function getRecipient(
    config
) {

    return String(
        process.env[
            config.recipientEnv
        ]
        ||
        config.recipientDefault
        ||
        ""
    ).trim();

}


function getRecipients(
    config
) {

    const envNames =
        [];


    function addEnvName(
        name
    ) {

        const clean =
            String(
                name
                ||
                ""
            ).trim();


        if (
            clean
            &&
            !envNames.includes(
                clean
            )
        ) {

            envNames.push(
                clean
            );

        }

    }


    if (
        config
    ) {

        addEnvName(
            config.recipientsEnv
        );

        addEnvName(
            config.recipientEnv
        );

        addEnvName(
            config.recipientEnv2
        );


        if (
            Array.isArray(
                config.recipientEnvs
            )
        ) {

            config.recipientEnvs
                .forEach(
                    addEnvName
                );

        }

    }


    const recipientValues =
        envNames.map(
            envName =>
                process.env[
                    envName
                ]
                ||
                ""
        );


    if (
        config
        &&
        config.recipientDefault
    ) {

        recipientValues.push(
            config.recipientDefault
        );

    }


    const recipients =
        [];


    recipientValues.forEach(
        value => {

            String(
                value
                ||
                ""
            )
                .split(
                    /[;,]+/
                )
                .map(
                    item =>
                        String(
                            item
                            ||
                            ""
                        ).trim()
                )
                .filter(
                    Boolean
                )
                .forEach(
                    email => {

                        const normalizedEmail =
                            email.toLowerCase();


                        const exists =
                            recipients.some(
                                current =>
                                    String(
                                        current
                                    )
                                        .toLowerCase()
                                    ===
                                    normalizedEmail
                            );


                        if (
                            !exists
                        ) {

                            recipients.push(
                                email
                            );

                        }

                    }
                );

        }
    );


    return recipients;

}


function getRecipientConfigurationStatus(
    config
) {

    const envNames =
        [];


    function addEnvName(
        name
    ) {

        const clean =
            String(
                name
                ||
                ""
            ).trim();


        if (
            clean
            &&
            !envNames.includes(
                clean
            )
        ) {

            envNames.push(
                clean
            );

        }

    }


    if (
        config
    ) {

        addEnvName(
            config.recipientsEnv
        );

        addEnvName(
            config.recipientEnv
        );

        addEnvName(
            config.recipientEnv2
        );


        if (
            Array.isArray(
                config.recipientEnvs
            )
        ) {

            config.recipientEnvs
                .forEach(
                    addEnvName
                );

        }

    }


    const configured =
        {};


    envNames.forEach(
        envName => {

            configured[
                envName
            ] =
                Boolean(
                    String(
                        process.env[
                            envName
                        ]
                        ||
                        ""
                    ).trim()
                );

        }
    );


    return {

        recipients:
            getRecipients(
                config
            ),

        configured:
            configured

    };

}


function getPanIndiaEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.panIndia;


    const recipients =
        getRecipients(
            config
        );


    return res.json({

        success:
            recipients.length >= 2,

        recipientCount:
            recipients.length,

        configured: {

            PAN_INDIA_RECIPIENTS:
                Boolean(
                    String(
                        process.env
                            .PAN_INDIA_RECIPIENTS
                        ||
                        ""
                    ).trim()
                ),

            PAN_INDIA_RECIPIENT:
                Boolean(
                    String(
                        process.env
                            .PAN_INDIA_RECIPIENT
                        ||
                        ""
                    ).trim()
                ),

            PAN_INDIA_RECIPIENT_2:
                Boolean(
                    String(
                        process.env
                            .PAN_INDIA_RECIPIENT_2
                        ||
                        ""
                    ).trim()
                )

        },

        message:
            recipients.length >= 2
                ?
                "Pan India email recipients are configured correctly."
                :
                "Two unique Pan India recipients are required. "
                +
                "Recommended: set PAN_INDIA_RECIPIENTS=email1@example.com,email2@example.com in back-end/.env, then restart Node."

    });

}


/* ============================================================
   EMAIL HTML TABLE
============================================================ */

function createEmailHtml(
    report,
    config
) {

    const rows =
        buildReportRows(report, config);

    const reportMonth =
        report.reportMonthName ||
        report.reportMonth ||
        "N/A";

    const tableRows =
        rows.map(
            row => {

                const style =
                    getStatusStyle(
                        row.result
                    );

                return `
                    <tr>
                        <td style="padding:18px 14px;border:1px solid #dbe3ef;color:#111827;font-size:15px;">
                            ${escapeHtml(row.particular)}
                        </td>

                        <td style="padding:18px 14px;border:1px solid #dbe3ef;color:#111827;font-size:15px;">
                            ${escapeHtml(row.target)}
                        </td>

                        <td style="padding:18px 14px;border:1px solid #dbe3ef;color:#111827;font-size:15px;">
                            ${escapeHtml(row.achieved)}
                        </td>

                        <td style="padding:14px;border:1px solid #dbe3ef;">
                            <span style="display:inline-block;padding:8px 13px;border-radius:18px;background:${style.emailBg};color:${style.emailText};font-size:14px;font-weight:700;white-space:nowrap;">
                                ${escapeHtml(style.label)}
                            </span>
                        </td>
                    </tr>
                `;

            }
        )
        .join("");

    return `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>

<body style="margin:0;padding:0;background:#f5f7fb;font-family:Arial,Helvetica,sans-serif;color:#111827;">

<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f5f7fb;padding:24px 12px;">
<tr>
<td align="center">

<table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:920px;background:#ffffff;border-radius:10px;overflow:hidden;">

<tr>
<td style="padding:32px 36px 10px 36px;">
    <h1 style="margin:0;color:#111827;font-size:25px;line-height:1.3;">
        ${escapeHtml(config.title)}
    </h1>
</td>
</tr>

<tr>
<td style="padding:10px 36px 4px 36px;font-size:16px;color:#111827;">
    Report Month:
    <strong>${escapeHtml(reportMonth)}</strong>
</td>
</tr>

<tr>
<td style="padding:16px 36px 26px 36px;font-size:16px;line-height:1.6;color:#64748b;">
    ${escapeHtml(config.description)}
</td>
</tr>

<tr>
<td style="padding:0 36px 26px 36px;">

<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="border-collapse:collapse;border:1px solid #dbe3ef;">

<thead>
<tr style="background:#eef3f8;">
    <th align="left" style="padding:16px 14px;border:1px solid #dbe3ef;font-size:15px;color:#111827;">
        Particular
    </th>
    <th align="left" style="padding:16px 14px;border:1px solid #dbe3ef;font-size:15px;color:#111827;">
        Target
    </th>
    <th align="left" style="padding:16px 14px;border:1px solid #dbe3ef;font-size:15px;color:#111827;">
        Achieved
    </th>
    <th align="left" style="padding:16px 14px;border:1px solid #dbe3ef;font-size:15px;color:#111827;">
        Result
    </th>
</tr>
</thead>

<tbody>
${tableRows}
</tbody>

</table>

</td>
</tr>

<tr>
<td style="padding:0 36px 30px 36px;color:#64748b;font-size:14px;line-height:1.6;">
    Summary:
    <strong style="color:#111827;">
        ${escapeHtml(
            getReportSummary(
                report.data,
                config
            )
        )}
    </strong>
</td>
</tr>

</table>

</td>
</tr>
</table>

</body>
</html>
    `;

}


/* ============================================================
   PDF TABLE
============================================================ */

function drawPdfCell(
    doc,
    text,
    x,
    y,
    width,
    height,
    options = {}
) {

    const {
        background = "#FFFFFF",
        color = "#111827",
        bold = false,
        fontSize = 10,
        padding = 8,
        align = "left",
        borderColor = "#D7DEE8"
    } =
        options;

    doc
        .save()
        .rect(
            x,
            y,
            width,
            height
        )
        .fillAndStroke(
            background,
            borderColor
        )
        .restore();

    doc
        .fillColor(color)
        .font(
            bold
                ?
                "Helvetica-Bold"
                :
                "Helvetica"
        )
        .fontSize(fontSize)
        .text(
            String(
                text ?? ""
            ),
            x + padding,
            y + padding,
            {
                width:
                    width -
                    padding * 2,

                height:
                    height -
                    padding * 2,

                align,
                ellipsis: true
            }
        );

}


function createPdfAttachment(
    report,
    config
) {

    return new Promise(
        (
            resolve,
            reject
        ) => {

            try {

                const normalizedReport = {

                    ...report,

                    data:
                        normalizeCompleteReportData(
                            report.data
                            ||
                            {}
                        )

                };


                const doc =
                    new PDFDocument({

                        size:
                            "A4",

                        margin:
                            36,

                        bufferPages:
                            true

                    });


                const chunks = [];


                doc.on(
                    "data",
                    chunk =>
                        chunks.push(
                            chunk
                        )
                );


                doc.on(
                    "end",
                    () =>
                        resolve(
                            Buffer.concat(
                                chunks
                            )
                        )
                );


                doc.on(
                    "error",
                    reject
                );


                const rows =
                    buildReportRows(
                        normalizedReport,
                        config
                    );


                const reportMonth =
                    normalizedReport.reportMonthName
                    ||
                    normalizedReport.reportMonth
                    ||
                    "N/A";


                const margin =
                    36;


                const pageWidth =
                    doc.page.width;


                const tableWidth =
                    pageWidth
                    -
                    margin * 2;


                const colWidths = [

                    tableWidth * 0.34,

                    tableWidth * 0.18,

                    tableWidth * 0.18,

                    tableWidth * 0.30

                ];


                const headerLabels = [

                    "Particular",

                    "Target",

                    "Achieved",

                    "Result"

                ];


                doc
                    .fillColor(
                        "#111827"
                    )
                    .font(
                        "Helvetica-Bold"
                    )
                    .fontSize(
                        20
                    )
                    .text(
                        config.title,
                        margin,
                        42,
                        {
                            width:
                                tableWidth
                        }
                    );


                doc
                    .font(
                        "Helvetica"
                    )
                    .fontSize(
                        10.5
                    )
                    .fillColor(
                        "#111827"
                    )
                    .text(
                        `Report Month: ${reportMonth}`,
                        margin,
                        78
                    );


                let y =
                    108;


                const headerHeight =
                    34;


                const rowHeight =
                    42;


                function drawHeader() {

                    let x =
                        margin;


                    headerLabels.forEach(
                        (
                            label,
                            index
                        ) => {

                            drawPdfCell(
                                doc,
                                label,
                                x,
                                y,
                                colWidths[index],
                                headerHeight,
                                {
                                    background:
                                        "#EEF3F8",
                                    bold:
                                        true,
                                    fontSize:
                                        10
                                }
                            );


                            x +=
                                colWidths[index];

                        }
                    );


                    y +=
                        headerHeight;

                }


                function newPageIfNeeded(
                    requiredHeight
                ) {

                    if (
                        y
                        +
                        requiredHeight
                        >
                        doc.page.height
                        -
                        60
                    ) {

                        doc.addPage();

                        y =
                            42;

                        drawHeader();

                    }

                }


                drawHeader();


                rows.forEach(
                    row => {

                        newPageIfNeeded(
                            rowHeight
                        );


                        const style =
                            getStatusStyle(
                                row.result
                            );


                        let x =
                            margin;


                        drawPdfCell(
                            doc,
                            row.particular,
                            x,
                            y,
                            colWidths[0],
                            rowHeight,
                            {
                                fontSize:
                                    9.5
                            }
                        );


                        x +=
                            colWidths[0];


                        drawPdfCell(
                            doc,
                            row.target,
                            x,
                            y,
                            colWidths[1],
                            rowHeight,
                            {
                                fontSize:
                                    9.5
                            }
                        );


                        x +=
                            colWidths[1];


                        drawPdfCell(
                            doc,
                            row.achieved,
                            x,
                            y,
                            colWidths[2],
                            rowHeight,
                            {
                                fontSize:
                                    9.5
                            }
                        );


                        x +=
                            colWidths[2];


                        /*
                         * No arrows or icons.
                         */
                        drawPdfCell(
                            doc,
                            row.result,
                            x,
                            y,
                            colWidths[3],
                            rowHeight,
                            {
                                background:
                                    style.pdfBg,
                                color:
                                    style.pdfText,
                                bold:
                                    true,
                                fontSize:
                                    8.8,
                                align:
                                    "center"
                            }
                        );


                        y +=
                            rowHeight;

                    }
                );


                /*
                 * Summary is also inside bordered cells.
                 */
                const summary =
                    getReportSummary(
                        normalizedReport.data,
                        config
                    );


                const summaryLabelWidth =
                    76;


                const summaryTextWidth =
                    tableWidth
                    -
                    summaryLabelWidth;


                doc
                    .font(
                        "Helvetica"
                    )
                    .fontSize(
                        8.5
                    );


                const summaryHeight =
                    Math.max(

                        50,

                        doc.heightOfString(
                            summary,
                            {
                                width:
                                    summaryTextWidth
                                    -
                                    14
                            }
                        )
                        +
                        18

                    );


                newPageIfNeeded(
                    summaryHeight
                );


                drawPdfCell(
                    doc,
                    "Summary",
                    margin,
                    y,
                    summaryLabelWidth,
                    summaryHeight,
                    {
                        background:
                            "#F8FAFC",
                        bold:
                            true,
                        fontSize:
                            9
                    }
                );


                drawPdfCell(
                    doc,
                    summary,
                    margin
                    +
                    summaryLabelWidth,
                    y,
                    summaryTextWidth,
                    summaryHeight,
                    {
                        fontSize:
                            8.5
                    }
                );


                if (
                    config
                    &&
                    config.mode === "LEADS"
                    &&
                    Array.isArray(
                        normalizedReport.data.states
                    )
                    &&
                    normalizedReport.data.states.length
                ) {

                    doc.addPage({
                        size: "A4",
                        layout: "landscape",
                        margin: 30
                    });

                    const stateMargin = 30;
                    const stateTableWidth =
                        doc.page.width - stateMargin * 2;

                    doc
                        .fillColor("#111827")
                        .font("Helvetica-Bold")
                        .fontSize(15)
                        .text(
                            "State-wise / Source-wise Leads Breakdown",
                            stateMargin,
                            30,
                            { width: stateTableWidth }
                        );

                    let stateY = 60;

                    const stateHeaders = [
                        "State / Source",
                        "Leads Target",
                        "Leads Achieved",
                        "Leads Need",
                        "OP Target",
                        "OP Achieved",
                        "OP Need",
                        "Lead -> OP Target",
                        "Lead -> OP Achieved",
                        "Result"
                    ];

                    const stateWidths = [
                        120, 64, 72, 62, 56,
                        64, 54, 72, 80, 96
                    ];

                    const stateHeaderHeight = 36;
                    const stateRowHeight = 40;

                    function drawLeadsStateHeader() {
                        let stateX = stateMargin;

                        stateHeaders.forEach(
                            (label, index) => {
                                drawPdfCell(
                                    doc,
                                    label,
                                    stateX,
                                    stateY,
                                    stateWidths[index],
                                    stateHeaderHeight,
                                    {
                                        background: "#EEF3F8",
                                        bold: true,
                                        fontSize: 7.3,
                                        padding: 4
                                    }
                                );

                                stateX += stateWidths[index];
                            }
                        );

                        stateY += stateHeaderHeight;
                    }

                    drawLeadsStateHeader();

                    normalizedReport.data.states.forEach(
                        state => {

                            if (
                                stateY + stateRowHeight >
                                doc.page.height - 42
                            ) {
                                doc.addPage({
                                    size: "A4",
                                    layout: "landscape",
                                    margin: 30
                                });

                                stateY = 30;
                                drawLeadsStateHeader();
                            }

                            const stateLeads = state.leads || {};
                            const stateOp = state.opCount || {};
                            const stateLeadToOp = state.leadToOp || {};

                            const leadsNeed =
                                reportNeedToAchieve(
                                    stateLeads.target,
                                    stateLeads.achieved
                                );

                            const opNeed =
                                reportNeedToAchieve(
                                    stateOp.target,
                                    stateOp.achieved
                                );

                            const stateResult =
                                calculateStatus(
                                    stateLeadToOp.target,
                                    stateLeadToOp.achieved
                                );

                            const stateStyle =
                                getStatusStyle(stateResult);

                            const values = [
                                state.label ||
                                state.sourceColumn ||
                                state.key ||
                                "",

                                formatValue(stateLeads.target, false),
                                formatValue(stateLeads.achieved, false),
                                formatValue(leadsNeed, false),

                                formatValue(stateOp.target, false),
                                formatValue(stateOp.achieved, false),
                                formatValue(opNeed, false),

                                formatValue(stateLeadToOp.target, true),
                                formatValue(stateLeadToOp.achieved, true),
                                stateResult
                            ];

                            let stateX = stateMargin;

                            values.forEach(
                                (value, index) => {

                                    drawPdfCell(
                                        doc,
                                        value,
                                        stateX,
                                        stateY,
                                        stateWidths[index],
                                        stateRowHeight,
                                        {
                                            fontSize: 7.2,
                                            padding: 4,
                                            bold:
                                                index === 0 ||
                                                index === 9,
                                            background:
                                                index === 9
                                                    ? stateStyle.pdfBg
                                                    : "#FFFFFF",
                                            color:
                                                index === 9
                                                    ? stateStyle.pdfText
                                                    : "#111827",
                                            align:
                                                index === 9
                                                    ? "center"
                                                    : "left"
                                        }
                                    );

                                    stateX += stateWidths[index];
                                }
                            );

                            stateY += stateRowHeight;
                        }
                    );
                }


                doc
                    .font(
                        "Helvetica"
                    )
                    .fontSize(
                        8
                    )
                    .fillColor(
                        "#94A3B8"
                    )
                    .text(
                        `Generated: ${
                            normalizedReport.createdAt
                            ||
                            new Date().toISOString()
                        }`,
                        margin,
                        doc.page.height
                        -
                        42,
                        {
                            width:
                                tableWidth,
                            align:
                                "right"
                        }
                    );


                doc.end();

            }
            catch (
                error
            ) {

                reject(error);

            }

        }
    );

}


/* ============================================================
   EXCEL TABLE
============================================================ */

async function createExcelAttachment(
    report,
    config
) {

    const normalizedReport = {

        ...report,

        data:
            normalizeCompleteReportData(
                report.data
                ||
                {}
            )

    };


    const workbook =
        new ExcelJS.Workbook();


    const worksheet =
        workbook.addWorksheet(
            config.excelSheetName
        );


    workbook.creator =
        "AVIS Reports";


    workbook.created =
        new Date();


    worksheet.mergeCells(
        "A1:D1"
    );


    worksheet.getCell(
        "A1"
    ).value =
        config.title;


    worksheet.getCell(
        "A1"
    ).font = {

        bold:
            true,

        size:
            18,

        color: {
            argb:
                "FF111827"
        }

    };


    worksheet.getRow(
        1
    ).height =
        28;


    /*
     * Metadata section.
     */
    worksheet.getCell(
        "A3"
    ).value =
        "Report Month";


    worksheet.getCell(
        "B3"
    ).value =
        normalizedReport.reportMonthName
        ||
        normalizedReport.reportMonth
        ||
        "";


    worksheet.getCell(
        "A4"
    ).value =
        "Summary";


    worksheet.getCell(
        "B4"
    ).value =
        getReportSummary(
            normalizedReport.data,
            config
        );


    worksheet.getCell(
        "A3"
    ).font = {
        bold:
            true
    };


    worksheet.getCell(
        "A4"
    ).font = {
        bold:
            true
    };


    worksheet.getCell(
        "B4"
    ).alignment = {
        wrapText:
            true,
        vertical:
            "top"
    };


    const rows =
        buildReportRows(
            normalizedReport,
            config
        );


    /*
     * Real Excel table.
     */
    worksheet.addTable({

        name:
            "PerformanceSummaryTable",

        ref:
            "A6",

        headerRow:
            true,

        totalsRow:
            false,

        style: {
            theme:
                "TableStyleMedium2",
            showRowStripes:
                true
        },

        columns: [

            {
                name:
                    "Particular"
            },

            {
                name:
                    "Target"
            },

            {
                name:
                    "Achieved"
            },

            {
                name:
                    "Result"
            }

        ],

        rows:
            rows.map(
                row => [

                    row.particular,

                    row.target,

                    row.achieved,

                    /*
                     * Plain status only.
                     * No arrows/icons.
                     */
                    row.result

                ]
            )

    });


    const firstDataRow =
        7;


    rows.forEach(
        (
            row,
            index
        ) => {

            const excelRow =
                worksheet.getRow(
                    firstDataRow
                    +
                    index
                );


            excelRow.height =
                24;


            const resultCell =
                excelRow.getCell(
                    4
                );


            const statusStyle =
                getStatusStyle(
                    row.result
                );


            resultCell.font = {

                bold:
                    true,

                color: {
                    argb:
                        statusStyle
                            .pdfText
                            .replace(
                                "#",
                                "FF"
                            )
                }

            };


            resultCell.fill = {

                type:
                    "pattern",

                pattern:
                    "solid",

                fgColor: {
                    argb:
                        statusStyle
                            .pdfBg
                            .replace(
                                "#",
                                "FF"
                            )
                }

            };

        }
    );


    worksheet.columns = [

        {
            width:
                32
        },

        {
            width:
                22
        },

        {
            width:
                22
        },

        {
            width:
                27
        }

    ];


    worksheet.views = [

        {
            state:
                "frozen",

            ySplit:
                6
        }

    ];


    if (
        config
        &&
        config.mode === "LEADS"
        &&
        Array.isArray(
            normalizedReport.data.states
        )
        &&
        normalizedReport.data.states.length
    ) {

        const stateSheet =
            workbook.addWorksheet(
                "State Wise"
            );

        stateSheet.addRow([
            "State / Source",
            "Leads Target",
            "Leads Achieved",
            "Leads Need",
            "OP Target",
            "OP Achieved",
            "OP Need",
            "Lead -> OP Target",
            "Lead -> OP Achieved",
            "Result"
        ]);

        stateSheet.getRow(1).font = { bold: true };

        normalizedReport.data.states.forEach(
            state => {

                const stateLeads = state.leads || {};
                const stateOp = state.opCount || {};
                const stateLeadToOp = state.leadToOp || {};

                stateSheet.addRow([
                    state.label ||
                    state.sourceColumn ||
                    state.key ||
                    "",

                    stateLeads.target,
                    stateLeads.achieved,
                    reportNeedToAchieve(
                        stateLeads.target,
                        stateLeads.achieved
                    ),

                    stateOp.target,
                    stateOp.achieved,
                    reportNeedToAchieve(
                        stateOp.target,
                        stateOp.achieved
                    ),

                    stateLeadToOp.target,
                    stateLeadToOp.achieved,
                    calculateStatus(
                        stateLeadToOp.target,
                        stateLeadToOp.achieved
                    )
                ]);
            }
        );

        stateSheet.columns.forEach(
            column => {
                column.width = 20;
            }
        );

        stateSheet.getColumn(1).width = 30;

        stateSheet.views = [
            {
                state: "frozen",
                ySplit: 1,
                xSplit: 1
            }
        ];
    }


    const buffer =
        await workbook.xlsx
            .writeBuffer();


    return Buffer.from(
        buffer
    );

}


/* ============================================================
   APPS SCRIPT
============================================================ */

async function parseAppsScriptResponse(
    response
) {

    const text =
        await response.text();

    const contentType =
        String(
            response.headers.get(
                "content-type"
            ) ||
            ""
        )
            .toLowerCase();

    if (
        contentType.includes(
            "text/html"
        ) ||
        text.trim().startsWith(
            "<!DOCTYPE html"
        ) ||
        text.trim().startsWith(
            "<html"
        )
    ) {

        throw new Error(
            `Google Apps Script returned HTML instead of JSON (HTTP ${response.status}, ` +
            `host ${new URL(response.url).host}). Check the Web App deployment, access, ` +
            "authorization and Apps Script execution logs."
        );

    }

    let result;

    try {

        result =
            JSON.parse(text);

    }
    catch (
        error
    ) {

        throw new Error(
            "Google Apps Script returned invalid JSON: " +
            text.slice(
                0,
                300
            )
        );

    }

    if (
        !response.ok ||
        !result.success
    ) {

        if (
            result.code ===
            "MISSING_UPLOAD_SECRET"
        ) {

            throw new Error(
                "Apps Script DRIVE_UPLOAD_SECRET is missing. "
                +
                "Open Apps Script -> Project Settings -> Script Properties, "
                +
                "add DRIVE_UPLOAD_SECRET, save it, and use the exact same value "
                +
                "in back-end/.env."
            );

        }


        if (
            result.code ===
            "INVALID_UPLOAD_SECRET" ||
            String(
                result.message ||
                ""
            )
                .toLowerCase()
                .includes(
                    "invalid upload secret"
                )
        ) {

            throw new Error(
                "DRIVE_UPLOAD_SECRET mismatch. "
                +
                "The value in back-end/.env must exactly match "
                +
                "Apps Script Project Settings -> Script Properties -> DRIVE_UPLOAD_SECRET. "
                +
                "Also confirm APPS_SCRIPT_WEB_APP_URL points to the deployment "
                +
                "of the same Apps Script project you edited."
            );

        }


        if (
            result.code ===
            "SHEETS_PERMISSION_MISSING" ||
            result.code ===
            "AUTHORIZATION_REQUIRED" ||
            String(
                result.message ||
                ""
            ).includes(
                "SpreadsheetApp.openById"
            )
        ) {

            throw new Error(
                "Google Sheets permission is missing. Run setupAuthorization() in Apps Script, approve Sheets + Drive, then redeploy the Web App."
            );

        }


        throw new Error(
            result.message ||
            `Google Apps Script failed with HTTP ${response.status}`
        );

    }

    return result;

}


function getWebAppUrl() {

    const webAppUrl =
        String(
            process.env
                .APPS_SCRIPT_WEB_APP_URL
            ||
            ""
        )
            .trim();

    if (
        !webAppUrl
    ) {

        throw new Error(
            "APPS_SCRIPT_WEB_APP_URL is missing in .env"
        );

    }

    if (
        !/\/exec(?:\?|$)/
            .test(
                webAppUrl
            )
    ) {

        throw new Error(
            "APPS_SCRIPT_WEB_APP_URL must end in /exec."
        );

    }

    return webAppUrl;

}


async function callAppsScript(
    payload
) {

    const webAppUrl =
        getWebAppUrl();


    const securePayload = {

        ...(
            payload
            ||
            {}
        ),

        secret:
            getAppsScriptSharedSecret()

    };


    const form =
        new URLSearchParams();

    form.set(
        "payload",
        JSON.stringify(
            securePayload
        )
    );

    const response =
        await fetch(
            webAppUrl,
            {
                method: "POST",
                redirect: "follow",
                headers: {
                    "Content-Type":
                        "application/x-www-form-urlencoded;charset=UTF-8"
                },
                body:
                    form.toString()
            }
        );

    return parseAppsScriptResponse(
        response
    );

}


async function callAppsScriptGet(
    action,
    config
) {

    const url =
        new URL(
            getWebAppUrl()
        );

    url.searchParams.set(
        "action",
        action
    );

    url.searchParams.set(
        "secret",
        getAppsScriptSharedSecret()
    );

    const requestReportKey =
        (
            String(
                action ||
                ""
            )
                .trim()
                .toLowerCase() ===
            "liveinput"
        )
            ?
            (
                config.liveInputKey ||
                config.key
            )
            :
            config.key;


    url.searchParams.set(
        "reportKey",
        requestReportKey
    );

    const attempts = String(requestReportKey).toUpperCase() === "APTS" ? 2 : 1;

    for (let attempt = 1; attempt <= attempts; attempt++) {
        const response = await fetch(url.toString(), {
            method: "GET",
            redirect: "follow"
        });

        try {
            return await parseAppsScriptResponse(response);
        } catch (error) {
            if (
                attempt === attempts ||
                !error.message.includes("returned HTML instead of JSON")
            ) {
                throw error;
            }
            await new Promise(resolve => setTimeout(resolve, 750));
        }
    }

}


/* ============================================================
   PAN INDIA PDF / EXCEL / EMAIL
============================================================ */

function panIndiaFormatReportValue(
    value,
    type
) {

    if (
        type ===
        "currency"
    ) {

        return formatCurrency(
            value
        );

    }


    if (
        type ===
        "percent"
    ) {

        return formatValue(
            value,
            true
        );

    }


    return formatValue(
        value,
        false
    );

}


function buildPanIndiaGroupedReportSections(
    parameterCards
) {

    const cards =
        Array.isArray(
            parameterCards
        )
            ?
            parameterCards
            :
            [];


    const sectionOrder = [
        "Count Parameters",
        "Conversion Parameters",
        "Revenue Parameters",
        "Per-Unit Revenue & Discount Parameters"
    ];


    const grouped =
        {};


    cards.forEach(
        card => {

            const section =
                card.section
                ||
                "Other Parameters";


            if (
                !grouped[
                    section
                ]
            ) {

                grouped[
                    section
                ] =
                    [];

            }


            grouped[
                section
            ].push(
                card
            );

        }
    );


    const sectionConfig = {

        "Count Parameters": {
            title:
                "Count Parameters",

            headers: [
                "Particular",
                "Target",
                "Achieved / Received",
                "Not Achieved / Unreceived"
            ]
        },

        "Conversion Parameters": {
            title:
                "Conversion Parameters",

            headers: [
                "Particular",
                "Target",
                "Achieved",
                "Gap To Target"
            ]
        },

        "Revenue Parameters": {
            title:
                "Revenue Parameters",

            headers: [
                "Particular",
                "Target",
                "Achieved",
                "Yet To Achieve"
            ]
        },

        "Per-Unit Revenue & Discount Parameters": {
            title:
                "Per-Unit Revenue & Discount Parameters",

            headers: [
                "Particular",
                "Target",
                "Generated",
                "Discount"
            ]
        }

    };


    const orderedNames = [
        ...sectionOrder.filter(
            name =>
                grouped[
                    name
                ]
        ),

        ...Object.keys(
            grouped
        ).filter(
            name =>
                !sectionOrder.includes(
                    name
                )
        )
    ];


    return orderedNames.map(
        sectionName => {

            const config =
                sectionConfig[
                    sectionName
                ]
                ||
                {
                    title:
                        sectionName,

                    headers: [
                        "Particular",
                        "Value 1",
                        "Value 2",
                        "Value 3"
                    ]
                };


            const rows =
                grouped[
                    sectionName
                ].map(
                    card => {

                        const fields =
                            Array.isArray(
                                card.fields
                            )
                                ?
                                card.fields
                                :
                                [];


                        const first =
                            fields[0]
                            ||
                            {};


                        const second =
                            fields[1]
                            ||
                            {};


                        const third =
                            fields[2]
                            ||
                            {};


                        let thirdValue =
                            third.value;


                        let thirdType =
                            third.type
                            ||
                            first.type
                            ||
                            "number";


                        if (
                            sectionName ===
                            "Conversion Parameters"
                        ) {

                            thirdValue =
                                (
                                    validNumber(
                                        first.value
                                    )
                                    &&
                                    validNumber(
                                        second.value
                                    )
                                )
                                    ?
                                    Math.max(
                                        0,
                                        Number(
                                            first.value
                                        )
                                        -
                                        Number(
                                            second.value
                                        )
                                    )
                                    :
                                    null;


                            thirdType =
                                "percent";

                        }


                        return {

                            particular:
                                card.title
                                ||
                                "Parameter",

                            values: [
                                {
                                    value:
                                        first.value,

                                    type:
                                        first.type
                                        ||
                                        "number"
                                },

                                {
                                    value:
                                        second.value,

                                    type:
                                        second.type
                                        ||
                                        first.type
                                        ||
                                        "number"
                                },

                                {
                                    value:
                                        thirdValue,

                                    type:
                                        thirdType
                                }
                            ]

                        };

                    }
                );


            return {

                title:
                    config.title,

                headers:
                    config.headers,

                rows:
                    rows

            };

        }
    );

}


function normalizeStateWiseRevenueToken(
    value
) {

    return String(
        value
        ||
        ""
    )
        .trim()
        .toUpperCase()
        .replace(
            /[^A-Z0-9]/g,
            ""
        );

}


function findStateWiseRevenueState(
    states,
    aliases
) {

    const list =
        Array.isArray(
            states
        )
            ?
            states
            :
            [];


    const tokens =
        aliases.map(
            normalizeStateWiseRevenueToken
        );


    return list.find(
        state => {

            const values = [
                state && state.key,
                state && state.label,
                state && state.sourceColumn
            ]
                .map(
                    normalizeStateWiseRevenueToken
                );


            return tokens.some(
                token =>
                    values.includes(
                        token
                    )
            );

        }
    )
    ||
    null;

}


function getStateWiseRevenueNumericValue(
    state,
    metricKey,
    side
) {

    if (
        !state
        ||
        !state[
            metricKey
        ]
    ) {

        return 0;

    }


    const value =
        Number(
            state[
                metricKey
            ][
                side
            ]
        );


    return Number.isFinite(
        value
    )
        ?
        value
        :
        0;

}


function formatStateWiseRevenueNumber(
    value
) {

    const number =
        Number(
            value
        );


    if (
        !Number.isFinite(
            number
        )
    ) {

        return "0";

    }


    return Math.round(
        number
    )
        .toLocaleString(
            "en-IN"
        );

}


function buildStateWiseRevenueMatrix(
    states
) {

    const columns = [

        {
            label:
                "APTS",

            aliases: [
                "APTS",
                "AP & TS",
                "APTS Combined"
            ]
        },

        {
            label:
                "KA",

            aliases: [
                "KA",
                "Karnataka"
            ]
        },

        {
            label:
                "TN",

            aliases: [
                "TN",
                "Tamil Nadu",
                "TamilNadu"
            ]
        },

        {
            label:
                "WB",

            aliases: [
                "WB",
                "West Bengal",
                "WestBengal"
            ]
        },

        {
            label:
                "MH",

            aliases: [
                "MH",
                "Maharashtra"
            ]
        }

    ];


    const resolvedStates =
        columns.map(
            column =>
                findStateWiseRevenueState(
                    states,
                    column.aliases
                )
        );


    const rowDefinitions = [

        {
            label:
                "IP Revenue Target",

            metric:
                "ipRevenue",

            side:
                "target",

            style:
                "target"
        },

        {
            label:
                "IP Revenue Generated",

            metric:
                "ipRevenue",

            side:
                "achieved",

            style:
                "achieved-alt"
        },

        {
            label:
                "B2B Revenue Target",

            metric:
                "b2bRevenue",

            side:
                "target",

            style:
                "target"
        },

        {
            label:
                "B2B Revenue Achieved",

            metric:
                "b2bRevenue",

            side:
                "achieved",

            style:
                "achieved"
        },

        {
            label:
                "OP Revenue Target",

            metric:
                "opRevenue",

            side:
                "target",

            style:
                "target"
        },

        {
            label:
                "OP Revenue Achieved",

            metric:
                "opRevenue",

            side:
                "achieved",

            style:
                "achieved-alt"
        },

        {
            label:
                "Pharmacy Revenue Target",

            metric:
                "pharmacyRevenue",

            side:
                "target",

            style:
                "target"
        },

        {
            label:
                "Pharmacy Revenue Achieved",

            metric:
                "pharmacyRevenue",

            side:
                "achieved",

            style:
                "achieved"
        },

        {
            label:
                "Lab Revenue Target",

            metric:
                "labRevenue",

            side:
                "target",

            style:
                "target"
        },

        {
            label:
                "Lab Revenue Achieved",

            metric:
                "labRevenue",

            side:
                "achieved",

            style:
                "achieved-alt"
        },

        {
            label:
                "Total Target",

            metric:
                "totalRevenue",

            side:
                "target",

            style:
                "total-target"
        },

        {
            label:
                "Total Achieved",

            metric:
                "totalRevenue",

            side:
                "achieved",

            style:
                "total-achieved"
        }

    ];


    const rows =
        rowDefinitions.map(
            definition => {

                const values =
                    resolvedStates.map(
                        state =>
                            getStateWiseRevenueNumericValue(
                                state,
                                definition.metric,
                                definition.side
                            )
                    );


                return {

                    ...definition,

                    values:
                        values,

                    panIndiaTotal:
                        values.reduce(
                            (
                                sum,
                                value
                            ) =>
                                sum + value,
                            0
                        )

                };

            }
        );


    return {

        headers: [
            "Particulars",
            ...columns.map(
                column =>
                    column.label
            ),
            "Pan India Total"
        ],

        rows:
            rows

    };

}


function createPanIndiaPdfAttachment(
    report,
    config
) {

    return new Promise(
        (
            resolve,
            reject
        ) => {

            try {

                const data =
                    report.data
                    ||
                    {};


                const allParameters =
                    Array.isArray(
                        data.allParameters
                    )
                        ?
                        data.allParameters
                        :
                        [];


                const states =
                    Array.isArray(
                        data.states
                    )
                        ?
                        data.states
                        :
                        [];


                const doc =
                    new PDFDocument({

                        size:
                            "A4",

                        layout:
                            "landscape",

                        margin:
                            32,

                        bufferPages:
                            true

                    });


                const chunks =
                    [];


                doc.on(
                    "data",
                    chunk =>
                        chunks.push(
                            chunk
                        )
                );


                doc.on(
                    "end",
                    () =>
                        resolve(
                            Buffer.concat(
                                chunks
                            )
                        )
                );


                doc.on(
                    "error",
                    reject
                );


                /* =====================================================
                   PAGE 1+: GROUPED PAN INDIA PERFORMANCE SUMMARY
                ===================================================== */

                const parameterCards =
                    Array.isArray(
                        data.parameterCards
                    )
                        ?
                        data.parameterCards
                        :
                        [];


                const groupedSections =
                    buildPanIndiaGroupedReportSections(
                        parameterCards
                    );


                doc
                    .font(
                        "Helvetica-Bold"
                    )
                    .fontSize(
                        18
                    )
                    .text(
                        config.title
                    );


                doc
                    .moveDown(
                        0.4
                    )
                    .font(
                        "Helvetica"
                    )
                    .fontSize(
                        10
                    )
                    .text(
                        `Report Month: ${report.reportMonthName || report.reportMonth || "N/A"}`
                    );


                doc
                    .text(
                        `Generated: ${new Date(report.createdAt || Date.now()).toLocaleString("en-IN")}`
                    );


                doc
                    .moveDown(
                        0.8
                    );


                doc
                    .font(
                        "Helvetica-Bold"
                    )
                    .fontSize(
                        13
                    )
                    .text(
                        "Pan India Consolidated Performance Summary"
                    );


                doc
                    .moveDown(
                        0.45
                    );


                const left =
                    32;


                const groupedWidths = [
                    245,
                    170,
                    170,
                    185
                ];


                let y =
                    doc.y;


                function ensureGroupedRoom(
                    requiredHeight
                ) {

                    if (
                        y + requiredHeight >
                        doc.page.height - 42
                    ) {

                        doc.addPage();


                        doc
                            .font(
                                "Helvetica-Bold"
                            )
                            .fontSize(
                                12
                            )
                            .text(
                                "Pan India Consolidated Performance Summary - Continued"
                            );


                        doc.moveDown(
                            0.5
                        );


                        y =
                            doc.y;

                    }

                }


                function drawGroupedSection(
                    section
                ) {

                    const requiredHeight =
                        30
                        +
                        24
                        +
                        (
                            section.rows.length
                            *
                            24
                        )
                        +
                        14;


                    ensureGroupedRoom(
                        Math.min(
                            requiredHeight,
                            180
                        )
                    );


                    doc
                        .font(
                            "Helvetica-Bold"
                        )
                        .fontSize(
                            11
                        )
                        .text(
                            section.title,
                            left,
                            y
                        );


                    y +=
                        20;


                    let x =
                        left;


                    section.headers.forEach(
                        (
                            header,
                            index
                        ) => {

                            drawPdfCell(
                                doc,
                                header,
                                x,
                                y,
                                groupedWidths[index],
                                24,
                                {
                                    bold:
                                        true,

                                    background:
                                        "#e2e8f0",

                                    borderColor:
                                        "#cbd5e1",

                                    color:
                                        "#0f172a",

                                    fontSize:
                                        8,

                                    padding:
                                        5
                                }
                            );


                            x +=
                                groupedWidths[index];

                        }
                    );


                    y +=
                        24;


                    section.rows.forEach(
                        row => {

                            ensureGroupedRoom(
                                24
                            );


                            let rowX =
                                left;


                            const values = [
                                row.particular,
                                ...row.values.map(
                                    item =>
                                        panIndiaFormatReportValue(
                                            item.value,
                                            item.type
                                        )
                                )
                            ];


                            values.forEach(
                                (
                                    value,
                                    index
                                ) => {

                                    drawPdfCell(
                                        doc,
                                        value,
                                        rowX,
                                        y,
                                        groupedWidths[index],
                                        24,
                                        {
                                            bold:
                                                index === 0,

                                            background:
                                                "#ffffff",

                                            borderColor:
                                                "#cbd5e1",

                                            color:
                                                (
                                                    index === 3
                                                )
                                                    ?
                                                    "#c2410c"
                                                    :
                                                    "#0f172a",

                                            fontSize:
                                                8,

                                            padding:
                                                5
                                        }
                                    );


                                    rowX +=
                                        groupedWidths[index];

                                }
                            );


                            y +=
                                24;

                        }
                    );


                    y +=
                        14;

                }


                groupedSections.forEach(
                    drawGroupedSection
                );


                /* =====================================================
                   STATE WISE REVENUE REPORT - EXACT MATRIX FORMAT
                ===================================================== */

                const stateRevenueMatrix =
                    buildStateWiseRevenueMatrix(
                        states
                    );


                doc.addPage();


                doc
                    .font(
                        "Helvetica-Bold"
                    )
                    .fontSize(
                        16
                    )
                    .text(
                        "State Wise Revenue Report"
                    );


                doc.moveDown(
                    0.55
                );


                const revenueLeft =
                    32;


                const revenueColumnWidths = [
                    235,
                    90,
                    90,
                    90,
                    90,
                    90,
                    96
                ];


                const revenueRowHeight =
                    30;


                let revenueY =
                    doc.y;


                let revenueX =
                    revenueLeft;


                stateRevenueMatrix.headers.forEach(
                    (
                        header,
                        index
                    ) => {

                        drawPdfCell(
                            doc,
                            header,
                            revenueX,
                            revenueY,
                            revenueColumnWidths[
                                index
                            ],
                            revenueRowHeight,
                            {
                                bold:
                                    true,

                                background:
                                    "#ffffff",

                                borderColor:
                                    "#64748b",

                                color:
                                    "#111827",

                                fontSize:
                                    index === 0
                                        ?
                                        9
                                        :
                                        8,

                                padding:
                                    5
                            }
                        );


                        revenueX +=
                            revenueColumnWidths[
                                index
                            ];

                    }
                );


                revenueY +=
                    revenueRowHeight;


                const revenueBackgrounds = {

                    target:
                        "#9dc3e6",

                    achieved:
                        "#a9d18e",

                    "achieved-alt":
                        "#f4b183",

                    "total-target":
                        "#9dc3e6",

                    "total-achieved":
                        "#a9d18e"

                };


                stateRevenueMatrix.rows.forEach(
                    row => {

                        revenueX =
                            revenueLeft;


                        const values = [
                            row.label,

                            ...row.values.map(
                                formatStateWiseRevenueNumber
                            ),

                            formatStateWiseRevenueNumber(
                                row.panIndiaTotal
                            )
                        ];


                        values.forEach(
                            (
                                value,
                                index
                            ) => {

                                drawPdfCell(
                                    doc,
                                    value,
                                    revenueX,
                                    revenueY,
                                    revenueColumnWidths[
                                        index
                                    ],
                                    revenueRowHeight,
                                    {
                                        bold:
                                            index === 0
                                            ||
                                            row.style === "total-target"
                                            ||
                                            row.style === "total-achieved"
                                            ||
                                            index === values.length - 1,

                                        background:
                                            revenueBackgrounds[
                                                row.style
                                            ]
                                            ||
                                            "#ffffff",

                                        borderColor:
                                            "#64748b",

                                        color:
                                            "#111827",

                                        fontSize:
                                            7.2,

                                        padding:
                                            5
                                    }
                                );


                                revenueX +=
                                    revenueColumnWidths[
                                        index
                                    ];

                            }
                        );


                        revenueY +=
                            revenueRowHeight;

                    }
                );


                /* =====================================================
                   STATE-WISE / SOURCE-WISE BASE PARAMETER BREAKDOWN

                   The browser table contains 35 columns. Putting all 35
                   columns on one A4 page would make the PDF unreadable,
                   so the SAME complete breakdown is split into five
                   readable tables. Every table contains ALL states/sources.
                ===================================================== */

                function getStateMetricValue(
                    state,
                    key,
                    side,
                    type =
                        "number"
                ) {

                    const metric =
                        state
                        &&
                        state[key]
                            ?
                            state[key]
                            :
                            {};


                    return panIndiaFormatReportValue(
                        metric[side],
                        type
                    );

                }


                function getStateName(
                    state
                ) {

                    return String(
                        state.label
                        ||
                        state.sourceColumn
                        ||
                        state.key
                        ||
                        "N/A"
                    );

                }


                function drawStateBreakdownSection(
                    sectionTitle,
                    columns
                ) {

                    doc.addPage();


                    doc
                        .font(
                            "Helvetica-Bold"
                        )
                        .fontSize(
                            15
                        )
                        .text(
                            "State-wise / Source-wise Base Parameter Breakdown"
                        );


                    doc
                        .moveDown(
                            0.25
                        )
                        .font(
                            "Helvetica-Bold"
                        )
                        .fontSize(
                            10
                        )
                        .text(
                            sectionTitle
                        );


                    doc.moveDown(
                        0.45
                    );


                    let stateY =
                        doc.y;


                    const rowHeight =
                        31;


                    function drawStateHeader() {

                        let stateX =
                            left;


                        columns.forEach(
                            column => {

                                drawPdfCell(
                                    doc,
                                    column.header,
                                    stateX,
                                    stateY,
                                    column.width,
                                    rowHeight,
                                    {
                                        bold:
                                            true,

                                        background:
                                            "#e2e8f0",

                                        borderColor:
                                            "#cbd5e1",

                                        color:
                                            "#0f172a",

                                        fontSize:
                                            6.2,

                                        padding:
                                            4
                                    }
                                );


                                stateX +=
                                    column.width;

                            }
                        );


                        stateY +=
                            rowHeight;

                    }


                    function newContinuationPage() {

                        doc.addPage();


                        doc
                            .font(
                                "Helvetica-Bold"
                            )
                            .fontSize(
                                13
                            )
                            .text(
                                "State-wise / Source-wise Base Parameter Breakdown - Continued"
                            );


                        doc
                            .moveDown(
                                0.2
                            )
                            .font(
                                "Helvetica-Bold"
                            )
                            .fontSize(
                                9
                            )
                            .text(
                                sectionTitle
                            );


                        doc.moveDown(
                            0.35
                        );


                        stateY =
                            doc.y;


                        drawStateHeader();

                    }


                    drawStateHeader();


                    if (
                        states.length ===
                        0
                    ) {

                        const totalWidth =
                            columns.reduce(
                                (
                                    sum,
                                    column
                                ) =>
                                    sum + column.width,
                                0
                            );


                        drawPdfCell(
                            doc,
                            "No state/source data is available.",
                            left,
                            stateY,
                            totalWidth,
                            rowHeight,
                            {
                                background:
                                    "#ffffff",

                                borderColor:
                                    "#cbd5e1",

                                color:
                                    "#64748b",

                                fontSize:
                                    7,

                                padding:
                                    5
                            }
                        );


                        return;

                    }


                    states.forEach(
                        state => {

                            if (
                                stateY + rowHeight >
                                doc.page.height - 42
                            ) {

                                newContinuationPage();

                            }


                            let stateX =
                                left;


                            columns.forEach(
                                column => {

                                    let value;


                                    if (
                                        typeof column.value ===
                                        "function"
                                    ) {

                                        value =
                                            column.value(
                                                state
                                            );

                                    }
                                    else {

                                        value =
                                            getStateMetricValue(
                                                state,
                                                column.key,
                                                column.side,
                                                column.type
                                            );

                                    }


                                    drawPdfCell(
                                        doc,
                                        value,
                                        stateX,
                                        stateY,
                                        column.width,
                                        rowHeight,
                                        {
                                            bold:
                                                column.header ===
                                                "State / Source",

                                            background:
                                                "#ffffff",

                                            borderColor:
                                                "#cbd5e1",

                                            color:
                                                "#0f172a",

                                            fontSize:
                                                6.2,

                                            padding:
                                                4
                                        }
                                    );


                                    stateX +=
                                        column.width;

                                }
                            );


                            stateY +=
                                rowHeight;

                        }
                    );

                }


                /* -----------------------------------------------------
                   1) COUNTS
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "1. Counts",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                105,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "IP Target",
                            width:
                                55,
                            key:
                                "ipCount",
                            side:
                                "target"
                        },
                        {
                            header:
                                "IP Achieved",
                            width:
                                55,
                            key:
                                "ipCount",
                            side:
                                "achieved"
                        },
                        {
                            header:
                                "B2B Target",
                            width:
                                55,
                            key:
                                "b2bCount",
                            side:
                                "target"
                        },
                        {
                            header:
                                "B2B Achieved",
                            width:
                                55,
                            key:
                                "b2bCount",
                            side:
                                "achieved"
                        },
                        {
                            header:
                                "OP Target",
                            width:
                                55,
                            key:
                                "opCount",
                            side:
                                "target"
                        },
                        {
                            header:
                                "OP Achieved",
                            width:
                                55,
                            key:
                                "opCount",
                            side:
                                "achieved"
                        },
                        {
                            header:
                                "Pharmacy Bill Count",
                            width:
                                55,
                            key:
                                "pharmacy",
                            side:
                                "target"
                        },
                        {
                            header:
                                "Pharmacy Received",
                            width:
                                55,
                            key:
                                "pharmacy",
                            side:
                                "achieved"
                        },
                        {
                            header:
                                "Lab Count",
                            width:
                                55,
                            key:
                                "lab",
                            side:
                                "target"
                        },
                        {
                            header:
                                "Lab Achieved",
                            width:
                                55,
                            key:
                                "lab",
                            side:
                                "achieved"
                        },
                        {
                            header:
                                "Leads Target",
                            width:
                                55,
                            key:
                                "leads",
                            side:
                                "target"
                        },
                        {
                            header:
                                "Leads Achieved",
                            width:
                                55,
                            key:
                                "leads",
                            side:
                                "achieved"
                        }
                    ]
                );


                /* -----------------------------------------------------
                   2) CONVERSIONS
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "2. Conversion Parameters",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                125,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "OP -> IP Target",
                            width:
                                160,
                            key:
                                "opToIp",
                            side:
                                "target",
                            type:
                                "percent"
                        },
                        {
                            header:
                                "OP -> IP Achieved",
                            width:
                                160,
                            key:
                                "opToIp",
                            side:
                                "achieved",
                            type:
                                "percent"
                        },
                        {
                            header:
                                "Lead -> OP Target",
                            width:
                                160,
                            key:
                                "leadToOp",
                            side:
                                "target",
                            type:
                                "percent"
                        },
                        {
                            header:
                                "Lead -> OP Achieved",
                            width:
                                160,
                            key:
                                "leadToOp",
                            side:
                                "achieved",
                            type:
                                "percent"
                        }
                    ]
                );


                /* -----------------------------------------------------
                   3) IP / B2B / OP REVENUE
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "3. IP / B2B / OP Revenue",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                120,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "IP Revenue Target",
                            width:
                                105,
                            key:
                                "ipRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "IP Revenue Generated",
                            width:
                                105,
                            key:
                                "ipRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "B2B Revenue Target",
                            width:
                                105,
                            key:
                                "b2bRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "B2B Revenue Achieved",
                            width:
                                105,
                            key:
                                "b2bRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "OP Revenue Target",
                            width:
                                105,
                            key:
                                "opRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "OP Revenue Achieved",
                            width:
                                105,
                            key:
                                "opRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        }
                    ]
                );


                /* -----------------------------------------------------
                   4) PHARMACY / LAB REVENUE
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "4. Pharmacy / Lab Revenue",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                120,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "Pharmacy Revenue Target",
                            width:
                                105,
                            key:
                                "pharmacyRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Pharmacy Revenue Achieved",
                            width:
                                105,
                            key:
                                "pharmacyRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Pharmacy Revenue Yet To Be",
                            width:
                                105,
                            key:
                                "pharmacyRevenue",
                            side:
                                "remaining",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Lab Revenue Target",
                            width:
                                105,
                            key:
                                "labRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Lab Revenue Achieved",
                            width:
                                105,
                            key:
                                "labRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Lab Revenue Yet To Be",
                            width:
                                105,
                            key:
                                "labRevenue",
                            side:
                                "remaining",
                            type:
                                "currency"
                        }
                    ]
                );


                /* -----------------------------------------------------
                   5) TOTAL REVENUE
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "5. State-wise Total Revenue",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                150,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "Total Revenue Target",
                            width:
                                190,
                            key:
                                "totalRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Total Revenue Generated",
                            width:
                                190,
                            key:
                                "totalRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Total Revenue Yet To Achieve",
                            width:
                                210,
                            key:
                                "totalRevenue",
                            side:
                                "remaining",
                            type:
                                "currency"
                        }
                    ]
                );


                /* -----------------------------------------------------
                   6) PHARMACY / LAB PER-UNIT REVENUE & DISCOUNT
                ----------------------------------------------------- */

                drawStateBreakdownSection(
                    "5. Pharmacy / Lab Per-Unit Revenue & Discount",
                    [
                        {
                            header:
                                "State / Source",
                            width:
                                120,
                            value:
                                getStateName
                        },
                        {
                            header:
                                "Pharmacy Per Bill Revenue",
                            width:
                                105,
                            key:
                                "pharmacyPerBillRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Pharmacy Per Bill Revenue Generated",
                            width:
                                105,
                            key:
                                "pharmacyPerBillRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Discount Per Pharmacy Bill",
                            width:
                                105,
                            key:
                                "pharmacyPerBillRevenue",
                            side:
                                "discount",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Per Lab Revenue Target",
                            width:
                                105,
                            key:
                                "labPerBillRevenue",
                            side:
                                "target",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Per Lab Revenue Generated",
                            width:
                                105,
                            key:
                                "labPerBillRevenue",
                            side:
                                "achieved",
                            type:
                                "currency"
                        },
                        {
                            header:
                                "Discount Per Lab",
                            width:
                                105,
                            key:
                                "labPerBillRevenue",
                            side:
                                "discount",
                            type:
                                "currency"
                        }
                    ]
                );


                doc.end();

            }
            catch (
                error
            ) {

                reject(
                    error
                );

            }

        }
    );

}


async function createPanIndiaExcelAttachment(
    report,
    config
) {

    const workbook =
        new ExcelJS.Workbook();


    workbook.creator =
        "AVIS Reports";


    const summarySheet =
        workbook.addWorksheet(
            "Pan India Summary"
        );


    const parameterCards =
        report.data
        &&
        Array.isArray(
            report.data.parameterCards
        )
            ?
            report.data.parameterCards
            :
            [];


    const groupedSections =
        buildPanIndiaGroupedReportSections(
            parameterCards
        );


    summarySheet.columns = [
        {
            key:
                "particular",
            width:
                34
        },
        {
            key:
                "value1",
            width:
                24
        },
        {
            key:
                "value2",
            width:
                24
        },
        {
            key:
                "value3",
            width:
                28
        }
    ];


    summarySheet.addRow([
        config.title
    ]);


    summarySheet.mergeCells(
        "A1:D1"
    );


    summarySheet.getCell(
        "A1"
    ).font = {
        bold:
            true,
        size:
            16
    };


    summarySheet.addRow([
        `Report Month: ${report.reportMonthName || report.reportMonth || "N/A"}`
    ]);


    summarySheet.mergeCells(
        "A2:D2"
    );


    summarySheet.addRow([]);


    groupedSections.forEach(
        section => {

            const titleRow =
                summarySheet.addRow([
                    section.title
                ]);


            summarySheet.mergeCells(
                `A${titleRow.number}:D${titleRow.number}`
            );


            titleRow.font = {
                bold:
                    true,
                size:
                    12
            };


            const headerRow =
                summarySheet.addRow(
                    section.headers
                );


            headerRow.font = {
                bold:
                    true
            };


            section.rows.forEach(
                row => {

                    summarySheet.addRow([
                        row.particular,

                        ...row.values.map(
                            item =>
                                panIndiaFormatReportValue(
                                    item.value,
                                    item.type
                                )
                        )
                    ]);

                }
            );


            summarySheet.addRow([]);

        }
    );


    summarySheet.views = [
        {
            state:
                "frozen",
            ySplit:
                2
        }
    ];


    const revenueMatrixSheet =
        workbook.addWorksheet(
            "State Wise Revenue"
        );


    const stateRevenueMatrix =
        buildStateWiseRevenueMatrix(
            report.data
            &&
            Array.isArray(
                report.data.states
            )
                ?
                report.data.states
                :
                []
        );


    revenueMatrixSheet.mergeCells(
        "A1:G1"
    );


    revenueMatrixSheet.getCell(
        "A1"
    ).value =
        "state wise revenue report";


    revenueMatrixSheet.getCell(
        "A1"
    ).font = {
        bold:
            true,
        size:
            14
    };


    const revenueHeaderRow =
        revenueMatrixSheet.addRow(
            stateRevenueMatrix.headers
        );


    revenueHeaderRow.font = {
        bold:
            true,
        size:
            12
    };


    revenueMatrixSheet.columns = [
        {
            width:
                34
        },
        {
            width:
                16
        },
        {
            width:
                16
        },
        {
            width:
                16
        },
        {
            width:
                16
        },
        {
            width:
                16
        },
        {
            width:
                20
        }
    ];


    const revenueFillByStyle = {

        target:
            "FF9DC3E6",

        achieved:
            "FFA9D18E",

        "achieved-alt":
            "FFF4B183",

        "total-target":
            "FF9DC3E6",

        "total-achieved":
            "FFA9D18E"

    };


    stateRevenueMatrix.rows.forEach(
        row => {

            const excelRow =
                revenueMatrixSheet.addRow([
                    row.label,
                    ...row.values,
                    row.panIndiaTotal
                ]);


            excelRow.eachCell(
                (
                    cell,
                    columnNumber
                ) => {

                    cell.fill = {
                        type:
                            "pattern",

                        pattern:
                            "solid",

                        fgColor: {
                            argb:
                                revenueFillByStyle[
                                    row.style
                                ]
                                ||
                                "FFFFFFFF"
                        }
                    };


                    cell.border = {
                        top: {
                            style:
                                "thin",
                            color: {
                                argb:
                                    "FF64748B"
                            }
                        },

                        left: {
                            style:
                                "thin",
                            color: {
                                argb:
                                    "FF64748B"
                            }
                        },

                        bottom: {
                            style:
                                "thin",
                            color: {
                                argb:
                                    "FF64748B"
                            }
                        },

                        right: {
                            style:
                                "thin",
                            color: {
                                argb:
                                    "FF64748B"
                            }
                        }
                    };


                    cell.alignment = {
                        vertical:
                            "middle",

                        horizontal:
                            "right"
                    };


                    if (
                        columnNumber >
                        1
                    ) {

                        cell.numFmt =
                            "#,##,##0";

                    }


                    if (
                        columnNumber === 1
                        ||
                        columnNumber === 7
                        ||
                        row.style === "total-target"
                        ||
                        row.style === "total-achieved"
                    ) {

                        cell.font = {
                            bold:
                                true
                        };

                    }

                }
            );

        }
    );


    revenueMatrixSheet
        .getRow(
            2
        )
        .eachCell(
            cell => {

                cell.border = {
                    top: {
                        style:
                            "thin",
                        color: {
                            argb:
                                "FF64748B"
                        }
                    },

                    left: {
                        style:
                            "thin",
                        color: {
                            argb:
                                "FF64748B"
                        }
                    },

                    bottom: {
                        style:
                            "thin",
                        color: {
                            argb:
                                "FF64748B"
                        }
                    },

                    right: {
                        style:
                            "thin",
                        color: {
                            argb:
                                "FF64748B"
                        }
                    }
                };


                cell.alignment = {
                    vertical:
                        "middle",

                    horizontal:
                        "center"
                };

            }
        );


    revenueMatrixSheet.views = [
        {
            state:
                "frozen",
            ySplit:
                2,
            xSplit:
                1
        }
    ];


    const stateSheet =
        workbook.addWorksheet(
            "State Wise"
        );


    const stateHeaders = [
        "State / Source",
        "IP Target",
        "IP Achieved",
        "B2B Target",
        "B2B Achieved",
        "OP Target",
        "OP Achieved",
        "Pharmacy Bill Count",
        "Pharmacy Received",
        "Lab Count",
        "Lab Achieved",
        "Leads Target",
        "Leads Achieved",
        "OP -> IP Target",
        "OP -> IP Achieved",
        "Lead -> OP Target",
        "Lead -> OP Achieved",
        "IP Revenue Target",
        "IP Revenue Generated",
        "B2B Revenue Target",
        "B2B Revenue Achieved",
        "OP Revenue Target",
        "OP Revenue Achieved",
        "Pharmacy Revenue Target",
        "Pharmacy Revenue Achieved",
        "Pharmacy Revenue Yet To Be",
        "Lab Revenue Target",
        "Lab Revenue Achieved",
        "Lab Revenue Yet To Be",
        "Total Revenue Target",
        "Total Revenue Generated",
        "Total Revenue Yet To Achieve",
        "Pharmacy Per Bill Revenue",
        "Pharmacy Per Bill Revenue Generated",
        "Discount Per Pharmacy Bill",
        "Per Lab Revenue Target",
        "Per Lab Revenue Generated",
        "Discount Per Lab"
    ];


    stateSheet.addRow(
        stateHeaders
    );


    stateSheet
        .getRow(
            1
        )
        .font = {
            bold:
                true
        };


    const states =
        report.data
        &&
        Array.isArray(
            report.data.states
        )
            ?
            report.data.states
            :
            [];


    states.forEach(
        state => {

            const get =
                (
                    key,
                    side
                ) => {

                    return (
                        state[
                            key
                        ]
                        &&
                        state[
                            key
                        ][
                            side
                        ]
                        !==
                        undefined
                    )
                        ?
                        state[
                            key
                        ][
                            side
                        ]
                        :
                        null;

                };


            stateSheet.addRow([
                state.label
                ||
                state.sourceColumn
                ||
                state.key
                ||
                "",

                get("ipCount","target"),
                get("ipCount","achieved"),

                get("b2bCount","target"),
                get("b2bCount","achieved"),

                get("opCount","target"),
                get("opCount","achieved"),

                get("pharmacy","target"),
                get("pharmacy","achieved"),

                get("lab","target"),
                get("lab","achieved"),

                get("leads","target"),
                get("leads","achieved"),

                get("opToIp","target"),
                get("opToIp","achieved"),

                get("leadToOp","target"),
                get("leadToOp","achieved"),

                get("ipRevenue","target"),
                get("ipRevenue","achieved"),

                get("b2bRevenue","target"),
                get("b2bRevenue","achieved"),

                get("opRevenue","target"),
                get("opRevenue","achieved"),

                get("pharmacyRevenue","target"),
                get("pharmacyRevenue","achieved"),
                get("pharmacyRevenue","remaining"),

                get("labRevenue","target"),
                get("labRevenue","achieved"),
                get("labRevenue","remaining"),

                get("totalRevenue","target"),
                get("totalRevenue","achieved"),
                get("totalRevenue","remaining"),

                get("pharmacyPerBillRevenue","target"),
                get("pharmacyPerBillRevenue","achieved"),
                get("pharmacyPerBillRevenue","discount"),

                get("labPerBillRevenue","target"),
                get("labPerBillRevenue","achieved"),
                get("labPerBillRevenue","discount")
            ]);

        }
    );


    stateSheet.columns.forEach(
        column => {

            column.width =
                22;

        }
    );


    stateSheet
        .getColumn(
            1
        )
        .width =
            30;


    stateSheet.views = [
        {
            state:
                "frozen",
            ySplit:
                1,
            xSplit:
                1
        }
    ];


    return workbook.xlsx.writeBuffer();

}


function createPanIndiaEmailHtml(
    report,
    config
) {

    const allParameters =
        report.data
        &&
        Array.isArray(
            report.data.allParameters
        )
            ?
            report.data.allParameters
            :
            [];


    const rows =
        allParameters
            .map(
                parameter => `
                    <tr>
                        <td style="padding:8px;border:1px solid #e2e8f0;">
                            ${escapeHtml(parameter.section || "")}
                        </td>
                        <td style="padding:8px;border:1px solid #e2e8f0;">
                            ${escapeHtml(parameter.label || parameter.id || "")}
                        </td>
                        <td style="padding:8px;border:1px solid #e2e8f0;text-align:right;">
                            ${escapeHtml(
                                panIndiaFormatReportValue(
                                    parameter.value,
                                    parameter.type
                                )
                            )}
                        </td>
                    </tr>
                `
            )
            .join(
                ""
            );


    return `
        <div style="font-family:Arial,Helvetica,sans-serif;color:#0f172a;">
            <h2>${escapeHtml(config.title)}</h2>

            <p>
                Report Month:
                <strong>
                    ${escapeHtml(report.reportMonthName || report.reportMonth || "")}
                </strong>
            </p>

            <p>
                The complete Pan India PDF and Excel reports are attached.
            </p>

            <table
                style="border-collapse:collapse;width:100%;max-width:900px;"
            >
                <thead>
                    <tr>
                        <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:left;">
                            Category
                        </th>
                        <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:left;">
                            Parameter
                        </th>
                        <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:right;">
                            Pan India Value
                        </th>
                    </tr>
                </thead>

                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
    `;

}


async function savePanIndiaReportToDrive(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.panIndia;


    try {

        const {
            reportId,
            reportMonth,
            reportMonthName,
            createdAt,
            data
        } =
            req.body
            ||
            {};


        if (
            !data
            ||
            typeof data !==
            "object"
        ) {

            return res
                .status(
                    400
                )
                .json({

                    success:
                        false,

                    message:
                        "Pan India report data is required."

                });

        }


        const report = {

            reportId:
                reportId
                ||
                Date.now(),

            reportMonth:
                reportMonth
                ||
                "",

            reportMonthName:
                reportMonthName
                ||
                reportMonth
                ||
                "",

            createdAt:
                createdAt
                ||
                new Date()
                    .toISOString(),

            data:
                data

        };


        const [
            pdfBuffer,
            excelBuffer
        ] =
            await Promise.all([

                createPanIndiaPdfAttachment(
                    report,
                    config
                ),

                createPanIndiaExcelAttachment(
                    report,
                    config
                )

            ]);


        const monthName =
            safeFileName(
                report.reportMonthName
                ||
                report.reportMonth
            );


        const result =
            await callAppsScript({

                secret:
                    getAppsScriptSharedSecret(),

                reportKey:
                    config.key,

                folderName:
                    getDriveFolderName(
                        config
                    ),

                report:
                    report,

                files: [
                    {
                        name:
                            `${config.filePrefix}-${monthName}.pdf`,

                        mimeType:
                            "application/pdf",

                        base64:
                            pdfBuffer.toString(
                                "base64"
                            )
                    },
                    {
                        name:
                            `${config.filePrefix}-${monthName}.xlsx`,

                        mimeType:
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                        base64:
                            Buffer
                                .from(
                                    excelBuffer
                                )
                                .toString(
                                    "base64"
                                )
                    }
                ]

            });


        return res.json({

            success:
                true,

            message:
                "Pan India report saved to PAN_INDIA Google Sheet and Google Drive.",

            sheetSaved:
                result.sheetSaved ===
                true,

            sheetRow:
                result.sheetRow
                ||
                null,

            sheetUrl:
                result.sheetUrl
                ||
                "",

            folderUrl:
                result.folderUrl
                ||
                "",

            files:
                result.files
                ||
                []

        });

    }
    catch (
        error
    ) {

        console.error(
            "Pan India save error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    "Unable to save Pan India report."

            });

    }

}


async function sendPanIndiaReportEmail(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.panIndia;


    try {

        const {
            reportId,
            reportMonth,
            reportMonthName,
            createdAt,
            data
        } =
            req.body
            ||
            {};


        if (
            !data
            ||
            typeof data !==
            "object"
        ) {

            return res
                .status(
                    400
                )
                .json({

                    success:
                        false,

                    message:
                        "Pan India report data is required."

                });

        }


        const recipients =
            getRecipients(
                config
            );


        /*
         * Pan India is intentionally configured for TWO recipients.
         * If only one unique email is configured, do not silently
         * send to one person and report success.
         */
        if (
            recipients.length <
            2
        ) {

            throw new Error(
                `Only ${recipients.length} unique Pan India email recipient(s) were loaded. `
                +
                "Recommended fix: add PAN_INDIA_RECIPIENTS=email1@example.com,email2@example.com "
                +
                "to back-end/.env and restart Node. "
                +
                "The older PAN_INDIA_RECIPIENT + PAN_INDIA_RECIPIENT_2 format is also supported."
            );

        }


        const report = {

            reportId:
                reportId
                ||
                Date.now(),

            reportMonth:
                reportMonth
                ||
                "",

            reportMonthName:
                reportMonthName
                ||
                reportMonth
                ||
                "",

            createdAt:
                createdAt
                ||
                new Date()
                    .toISOString(),

            data:
                data

        };


        const [
            pdfBuffer,
            excelBuffer
        ] =
            await Promise.all([

                createPanIndiaPdfAttachment(
                    report,
                    config
                ),

                createPanIndiaExcelAttachment(
                    report,
                    config
                )

            ]);


        const monthName =
            safeFileName(
                report.reportMonthName
                ||
                report.reportMonth
            );


        const info =
            await transporter.sendMail({

                from: {
                    name:
                        process.env.MAIL_FROM_NAME
                        ||
                        "AVIS Reports",

                    address:
                        process.env.MAIL_FROM_ADDRESS
                        ||
                        process.env.SMTP_USER
                },

                to:
                    recipients,

                subject:
                    `${config.title} - ${report.reportMonthName || report.reportMonth}`,

                html:
                    createPanIndiaEmailHtml(
                        report,
                        config
                    ),

                attachments: [
                    {
                        filename:
                            `${config.filePrefix}-${monthName}.pdf`,

                        content:
                            pdfBuffer,

                        contentType:
                            "application/pdf"
                    },
                    {
                        filename:
                            `${config.filePrefix}-${monthName}.xlsx`,

                        content:
                            Buffer.from(
                                excelBuffer
                            ),

                        contentType:
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    }
                ]

            });


        return res.json({

            success:
                true,

            recipientCount:
                recipients.length,

            recipients:
                recipients,

            message:
                `Pan India report sent successfully to ${recipients.length} recipients: ${recipients.join(", ")}.`,

            messageId:
                info.messageId
                ||
                ""

        });

    }
    catch (
        error
    ) {

        console.error(
            "Pan India email error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    "Unable to send Pan India report email."

            });

    }

}


function getPanIndiaGeneratedReports(
    req,
    res
) {

    return getGeneratedReports(
        REPORT_CONFIGS.panIndia,
        req,
        res
    );

}


function clearPanIndiaGeneratedReports(
    req,
    res
) {

    return clearGeneratedReports(
        REPORT_CONFIGS.panIndia,
        req,
        res
    );

}



/* ============================================================
   GENERIC REPORT ACTIONS
============================================================ */

async function testReportConnection(
    config,
    req,
    res
) {

    try {

        const result =
            await callAppsScriptGet(
                "test",
                config
            );

        return res.json({
            success: true,
            message:
                `${config.title} Google Sheet + Drive connection is working.`,
            appsScript:
                result
        });

    }
    catch (
        error
    ) {

        return res
            .status(500)
            .json({
                success: false,
                message:
                    error.message
            });

    }

}


async function getGeneratedReports(
    config,
    req,
    res
) {

    try {

        const result =
            await callAppsScriptGet(
                "history",
                config
            );


        const reports =
            Array.isArray(
                result.reports
            )
                ?
                result.reports.map(
                    report => ({

                        ...report,

                        data:
                            normalizeCompleteReportData(
                                report.data
                                ||
                                {}
                            )

                    })
                )
                :
                [];


        return res.json({

            success:
                true,

            count:
                reports.length,

            reports:
                reports,

            sheetName:
                result.sheetName
                ||
                config.sheetName,

            sheetUrl:
                result.sheetUrl
                ||
                ""

        });

    }
    catch (
        error
    ) {

        console.error(
            `${config.key} report history error:`,
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    `Unable to load ${config.key} report history.`

            });

    }

}


async function clearGeneratedReports(
    config,
    req,
    res
) {

    try {

        const result =
            await callAppsScript({
                action:
                    "clearHistory",
                secret:
                    getAppsScriptSharedSecret(),
                reportKey:
                    config.key
            });

        return res.json({
            success: true,
            message:
                result.message ||
                `Generated report history cleared from ${config.sheetName}.`
        });

    }
    catch (
        error
    ) {

        console.error(
            `Clear ${config.key} history error:`,
            error
        );

        return res
            .status(500)
            .json({
                success: false,
                message:
                    error.message ||
                    `Unable to clear ${config.key} history.`
            });

    }

}


async function saveReport(
    config,
    req,
    res
) {

    try {

        const {
            reportId,
            reportMonth,
            reportMonthName,
            createdAt,
            data
        } =
            req.body
            ||
            {};


        if (
            !data
            ||
            typeof data !==
            "object"
        ) {

            return res
                .status(
                    400
                )
                .json({

                    success:
                        false,

                    message:
                        "Report data is required."

                });

        }


        const normalizedData =
            normalizeCompleteReportData(
                data
            );


        const report = {

            reportId:
                reportId
                ||
                Date.now(),

            reportMonth:
                reportMonth
                ||
                "",

            reportMonthName:
                reportMonthName
                ||
                reportMonth
                ||
                "",

            createdAt:
                createdAt
                ||
                new Date()
                    .toISOString(),

            data:
                normalizedData

        };


        const [
            pdfBuffer,
            excelBuffer
        ] =
            await Promise.all([

                createPdfAttachment(
                    report,
                    config
                ),

                createExcelAttachment(
                    report,
                    config
                )

            ]);


        const monthName =
            safeFileName(
                report.reportMonthName
                ||
                report.reportMonth
            );


        const result =
            await callAppsScript({

                secret:
                    getAppsScriptSharedSecret(),

                reportKey:
                    config.key,

                folderName:
                    getDriveFolderName(
                        config
                    ),

                report:
                    report,

                files: [

                    {
                        name:
                            `${config.filePrefix}-${monthName}.pdf`,

                        mimeType:
                            "application/pdf",

                        base64:
                            pdfBuffer
                                .toString(
                                    "base64"
                                )
                    },

                    {
                        name:
                            `${config.filePrefix}-${monthName}.xlsx`,

                        mimeType:
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                        base64:
                            excelBuffer
                                .toString(
                                    "base64"
                                )
                    }

                ]

            });


        return res.json({

            success:
                true,

            message:
                `Report saved to ${config.sheetName} Google Sheet and Google Drive.`,

            sheetSaved:
                result.sheetSaved ===
                true,

            sheetRow:
                result.sheetRow
                ||
                null,

            sheetUrl:
                result.sheetUrl
                ||
                "",

            folderUrl:
                result.folderUrl
                ||
                "",

            files:
                result.files
                ||
                [],

            data:
                normalizedData

        });

    }
    catch (
        error
    ) {

        console.error(
            `${config.key} Sheet/Drive save error:`,
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    "Unable to save report."

            });

    }

}


async function sendReportEmail(
    config,
    req,
    res
) {

    try {

        const {
            reportId,
            reportMonth,
            reportMonthName,
            createdAt,
            data
        } =
            req.body
            ||
            {};


        if (
            !data
            ||
            typeof data !==
            "object"
        ) {

            return res
                .status(
                    400
                )
                .json({

                    success:
                        false,

                    message:
                        "Report data is required."

                });

        }


        const normalizedData =
            normalizeCompleteReportData(
                data
            );


        const report = {

            reportId:
                reportId
                ||
                Date.now(),

            reportMonth:
                reportMonth
                ||
                "",

            reportMonthName:
                reportMonthName
                ||
                reportMonth
                ||
                "",

            createdAt:
                createdAt
                ||
                new Date()
                    .toISOString(),

            data:
                normalizedData

        };


        /*
         * Multi-recipient support.
         *
         * Recommended:
         * RECIPIENTS=email1@example.com,email2@example.com
         *
         * Backward-compatible:
         * RECIPIENT=email1@example.com
         * RECIPIENT_2=email2@example.com
         */
        const recipients =
            getRecipients(
                config
            );


        const minimumRecipients =
            Number(
                config.minimumRecipients
                ||
                1
            );


        if (
            recipients.length <
            minimumRecipients
        ) {

            const combinedEnvName =
                config.recipientsEnv
                ||
                "";


            const firstEnvName =
                config.recipientEnv
                ||
                "";


            const secondEnvName =
                config.recipientEnv2
                ||
                "";


            const status =
                getRecipientConfigurationStatus(
                    config
                );


            const checkedNames =
                Object.keys(
                    status.configured
                )
                    .join(
                        ", "
                    );


            throw new Error(
                `Only ${recipients.length} unique email recipient(s) were loaded for ${config.title}. `
                +
                `This report requires ${minimumRecipients}. `
                +
                `Checked .env variables: ${checkedNames || "none"}. `
                +
                (
                    combinedEnvName
                        ?
                        `Recommended: set ${combinedEnvName}=email1@example.com`
                        +
                        (
                            minimumRecipients > 1
                                ?
                                ",email2@example.com"
                                :
                                ""
                        )
                        +
                        ". "
                        :
                        ""
                )
                +
                (
                    firstEnvName
                        ?
                        `You can also set ${firstEnvName}=email1@example.com. `
                        :
                        ""
                )
                +
                "Save .env and restart Node with Ctrl+C then npm run dev."
            );

        }


        const [
            pdfBuffer,
            excelBuffer
        ] =
            await Promise.all([

                createPdfAttachment(
                    report,
                    config
                ),

                createExcelAttachment(
                    report,
                    config
                )

            ]);


        const monthName =
            safeFileName(
                report.reportMonthName
                ||
                report.reportMonth
            );


        const reportMonthLabel =
            report.reportMonthName
            ||
            report.reportMonth
            ||
            "";


        const info =
            await transporter.sendMail({

                from:
                    `"${process.env.MAIL_FROM_NAME || "AVIS Reports"}" <${process.env.MAIL_FROM_ADDRESS || process.env.SMTP_USER}>`,

                /*
                 * Nodemailer accepts an array.
                 * One click sends the same email + attachments
                 * to all configured recipients.
                 */
                to:
                    recipients,

                subject:
                    `${config.title} - ${reportMonthLabel}`,

                html:
                    createEmailHtml(
                        report,
                        config
                    ),

                attachments: [

                    {
                        filename:
                            `${config.filePrefix}-${monthName}.pdf`,

                        content:
                            pdfBuffer,

                        contentType:
                            "application/pdf"
                    },

                    {
                        filename:
                            `${config.filePrefix}-${monthName}.xlsx`,

                        content:
                            excelBuffer,

                        contentType:
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    }

                ]

            });


        return res.json({

            success:
                true,

            /*
             * Keep recipient string for compatibility with old frontend code.
             */
            recipient:
                recipients.join(
                    ", "
                ),

            recipients:
                recipients,

            recipientCount:
                recipients.length,

            messageId:
                info.messageId
                ||
                "",

            message:
                `${config.title} sent successfully to ${recipients.length} recipients: ${recipients.join(", ")}.`,

            data:
                normalizedData

        });

    }
    catch (
        error
    ) {

        console.error(
            `${config.key} send email error:`,
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    "Unable to send email."

            });

    }

}


/* ============================================================
   LIVE INPUT SHEET

   Reads current values from the separate Input spreadsheet
   through Google Apps Script.
============================================================ */

async function getLiveInput(
    config,
    req,
    res
) {

    try {

        const result =
            await callAppsScriptGet(
                "liveInput",
                config
            );

        return res.json({
            success: true,
            reportKey:
                config.key,
            sourceColumn:
                result.sourceColumn ||
                config.liveInputKey ||
                config.key,
            report:
                result.report ||
                {},
            parameters:
                result.parameters ||
                {},
            revenue:
                result.revenue ||
                {},
            detectedHeaderRow:
                result.detectedHeaderRow ||
                null,
            detectedParticularsColumn:
                result.detectedParticularsColumn ||
                "",
            detectedSourceColumn:
                result.detectedSourceColumn ||
                "",
            inputSpreadsheetName:
                result.inputSpreadsheetName ||
                "Input",
            inputSheetName:
                result.inputSheetName ||
                "Sheet1",
            inputSheetUrl:
                result.inputSheetUrl ||
                "",
            refreshedAt:
                result.refreshedAt ||
                new Date().toISOString()
        });

    }
    catch (
        error
    ) {

        console.error(
            `${config.key} live input error:`,
            error
        );

        return res
            .status(500)
            .json({
                success: false,
                message:
                    error.message ||
                    `Unable to load ${config.key} live input data.`
            });

    }

}




/* ============================================================
   PAN INDIA LIVE DASHBOARD
============================================================ */

const PAN_INDIA_LIVE_SOURCES = [

    {
        key: "APTS",
        label: "Andhra Pradesh & Telangana"
    },

    {
        key: "KA",
        label: "Karnataka"
    },

    {
        key: "TN",
        label: "Tamil Nadu"
    },

    {
        key: "WB",
        label: "West Bengal"
    },

    {
        key: "MH",
        label: "Maharashtra"
    },

    {
        key: "DELHI",
        label: "Delhi"
    }

];

async function mapWithConcurrency(items, limit, worker) {
    const results = new Array(items.length);
    let nextIndex = 0;
    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
        while (nextIndex < items.length) {
            const index = nextIndex++;
            results[index] = await worker(items[index]);
        }
    }));
    return results;
}


function normalizePanIndiaParameterLabel(
    value
) {

    return String(
        value ||
        ""
    )
        .trim()
        .toLowerCase()
        .replace(
            /[_–—-]+/g,
            " "
        )
        .replace(
            /[^a-z0-9 ]+/g,
            " "
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim();

}


function panIndiaNumber(
    value
) {

    if (
        value === null
        ||
        value === undefined
        ||
        value === ""
    ) {

        return null;

    }


    if (
        typeof value ===
        "number"
    ) {

        return Number.isFinite(
            value
        )
            ?
            value
            :
            null;

    }


    const cleaned =
        String(
            value
        )
            .trim()
            .replace(
                /,/g,
                ""
            )
            .replace(
                /%/g,
                ""
            )
            .replace(
                /[^0-9.+-]/g,
                ""
            );


    if (
        !cleaned
    ) {

        return null;

    }


    const number =
        Number(
            cleaned
        );


    return Number.isFinite(
        number
    )
        ?
        number
        :
        null;

}


function getPanIndiaParameter(
    parameters,
    aliases
) {

    const source =
        parameters &&
        typeof parameters ===
        "object"
            ?
            parameters
            :
            {};


    const normalized =
        {};


    Object.keys(
        source
    )
        .forEach(
            key => {

                normalized[
                    normalizePanIndiaParameterLabel(
                        key
                    )
                ] =
                    source[
                        key
                    ];

            }
        );


    for (
        const alias
        of
        aliases
    ) {

        const key =
            normalizePanIndiaParameterLabel(
                alias
            );


        if (
            Object.prototype
                .hasOwnProperty
                .call(
                    normalized,
                    key
                )
        ) {

            return panIndiaNumber(
                normalized[
                    key
                ]
            );

        }

    }


    return null;

}


function normalizePanIndiaMetric(
    metric
) {

    const source =
        metric &&
        typeof metric ===
        "object"
            ?
            metric
            :
            {};


    return {

        target:
            panIndiaNumber(
                source.target
            ),

        achieved:
            panIndiaNumber(
                source.achieved
            )

    };

}


function panIndiaPercentFromCounts(
    numerator,
    denominator
) {

    const top =
        panIndiaNumber(
            numerator
        );


    const bottom =
        panIndiaNumber(
            denominator
        );


    if (
        top ===
        null
        ||
        bottom ===
        null
        ||
        Number(
            bottom
        ) ===
        0
    ) {

        return null;

    }


    return (
        Number(
            top
        )
        /
        Number(
            bottom
        )
    )
    *
    100;

}


function normalizePanIndiaConversion(
    metric,
    numeratorMetric,
    denominatorMetric
) {

    const normalized =
        normalizePanIndiaMetric(
            metric
        );


    const numerator =
        normalizePanIndiaMetric(
            numeratorMetric
        );


    const denominator =
        normalizePanIndiaMetric(
            denominatorMetric
        );


    return {

        target:
            normalized.target !==
            null
                ?
                normalized.target
                :
                panIndiaPercentFromCounts(
                    numerator.target,
                    denominator.target
                ),

        achieved:
            normalized.achieved !==
            null
                ?
                normalized.achieved
                :
                panIndiaPercentFromCounts(
                    numerator.achieved,
                    denominator.achieved
                )

    };

}


function getPanIndiaLabMetric(
    parameters
) {

    return {

        target:
            getPanIndiaParameter(
                parameters,
                [
                    "Lab Count Target",
                    "Lab Count"
                ]
            ),

        achieved:
            getPanIndiaParameter(
                parameters,
                [
                    "Lab Count Achieved",
                    "Lab Count Reached",
                    "Lab Count Received"
                ]
            )

    };

}


function getPanIndiaPharmacyMetric(
    parameters
) {

    return {

        target:
            getPanIndiaParameter(
                parameters,
                [
                    "Pharmacy Bill Count Target",
                    "Pharmacy Bill Count",
                    "Pharmacy Count Target",
                    "Pharmacy Count"
                ]
            ),

        achieved:
            getPanIndiaParameter(
                parameters,
                [
                    "Pharmacy Bill Count Received",
                    "Pharmacy Bill Count Achieved",
                    "Pharmacy Bill Count Reached",
                    "Pharmacy Count Received",
                    "Pharmacy Count Achieved",
                    "Pharmacy Count Reached"
                ]
            )

    };

}


function getPanIndiaRevenueMetric(
    parameters,
    targetAliases,
    achievedAliases
) {

    return {

        target:
            getPanIndiaParameter(
                parameters,
                targetAliases
            ),

        achieved:
            getPanIndiaParameter(
                parameters,
                achievedAliases
            )

    };

}


function resolvePanIndiaRevenueMetric(
    rawMetric,
    parameters,
    targetAliases,
    achievedAliases,
    remainingAliases
) {

    const raw =
        normalizePanIndiaMetric(
            rawMetric
        );


    const parameterMetric =
        getPanIndiaRevenueMetric(
            parameters,
            targetAliases,
            achievedAliases
        );


    const achieved =
        firstNumber(
            raw.achieved,
            parameterMetric.achieved
        );


    const rawRemaining =
        rawMetric
        &&
        typeof rawMetric ===
        "object"
            ?
            panIndiaNumber(
                rawMetric.remaining
            )
            :
            null;


    const parameterRemaining =
        getPanIndiaParameter(
            parameters,
            remainingAliases
        );


    const remaining =
        firstNumber(
            rawRemaining,
            parameterRemaining
        );


    /*
     * Critical fallback:
     *
     * Target = Achieved + Yet To Be
     *
     * "Yet To Be" is signed. A negative value means
     * achievement is above target, so it MUST NOT be clamped.
     */
    const calculatedTarget =
        (
            validNumber(
                achieved
            )
            &&
            validNumber(
                remaining
            )
        )
            ?
            (
                Number(
                    achieved
                )
                +
                Number(
                    remaining
                )
            )
            :
            null;


    const target =
        firstNumber(
            raw.target,
            parameterMetric.target,
            calculatedTarget
        );


    const resolvedRemaining =
        firstNumber(
            remaining,
            panIndiaSignedDifference(
                target,
                achieved
            )
        );


    return {

        target:
            target,

        achieved:
            achieved,

        remaining:
            panIndiaNeedToAchieve(
                target,
                achieved
            )

    };

}


function panIndiaSignedDifference(
    target,
    achieved
) {

    const targetNumber =
        panIndiaNumber(
            target
        );


    const achievedNumber =
        panIndiaNumber(
            achieved
        );


    if (
        targetNumber ===
        null
        ||
        achievedNumber ===
        null
    ) {

        return null;

    }


    return (
        Number(
            targetNumber
        )
        -
        Number(
            achievedNumber
        )
    );

}


function panIndiaNeedToAchieve(
    target,
    achieved
) {

    const difference =
        panIndiaSignedDifference(
            target,
            achieved
        );


    if (
        difference ===
        null
    ) {

        return null;

    }


    /*
     * Business rule:
     * once target is achieved or exceeded,
     * Yet To Achieve / Unachieved must display 0.
     */
    return Math.max(
        0,
        Number(
            difference
        )
    );

}


function panIndiaSafeDivide(
    numerator,
    denominator,
    zeroWhenBothZero =
        false
) {

    const top =
        panIndiaNumber(
            numerator
        );


    const bottom =
        panIndiaNumber(
            denominator
        );


    if (
        top ===
        null
        ||
        bottom ===
        null
    ) {

        return null;

    }


    if (
        Number(
            bottom
        ) ===
        0
    ) {

        if (
            zeroWhenBothZero
            &&
            Number(
                top
            ) ===
            0
        ) {

            return 0;

        }


        return null;

    }


    return (
        Number(
            top
        )
        /
        Number(
            bottom
        )
    );

}


function panIndiaParameterField(
    id,
    label,
    value,
    type
) {

    return {

        id:
            id,

        label:
            label,

        value:
            value,

        type:
            type
            ||
            "number"

    };

}


function buildPanIndiaParameterCards(
    totals
) {

    const ipCount =
        totals.ipCount
        ||
        {};


    const b2bCount =
        totals.b2bCount
        ||
        {};


    const opCount =
        totals.opCount
        ||
        {};


    const pharmacy =
        totals.pharmacy
        ||
        {};


    const lab =
        totals.lab
        ||
        {};


    const leads =
        totals.leads
        ||
        {};


    const opToIp =
        totals.opToIp
        ||
        {};


    const leadToOp =
        totals.leadToOp
        ||
        {};


    const ipRevenue =
        totals.ipRevenue
        ||
        {};


    const b2bRevenue =
        totals.b2bRevenue
        ||
        {};


    const opRevenue =
        totals.opRevenue
        ||
        {};


    const pharmacyRevenue =
        totals.pharmacyRevenue
        ||
        {};


    const labRevenue =
        totals.labRevenue
        ||
        {};


    const totalRevenue =
        totals.totalRevenue
        ||
        {};


    /*
     * Signed spreadsheet-style remaining values.
     * Negative values are intentionally preserved because the live
     * Input sheet also shows negative Unreceived / Yet To Be values
     * when achievement is above target.
     */
    const ipCountUnachieved =
        panIndiaNeedToAchieve(
            ipCount.target,
            ipCount.achieved
        );


    const b2bCountUnachieved =
        panIndiaNeedToAchieve(
            b2bCount.target,
            b2bCount.achieved
        );


    const opCountUnachieved =
        panIndiaNeedToAchieve(
            opCount.target,
            opCount.achieved
        );


    const pharmacyUnreceived =
        panIndiaNeedToAchieve(
            pharmacy.target,
            pharmacy.achieved
        );


    const labUnachieved =
        panIndiaNeedToAchieve(
            lab.target,
            lab.achieved
        );


    const leadsUnachieved =
        panIndiaNeedToAchieve(
            leads.target,
            leads.achieved
        );


    const ipRevenueYet =
        panIndiaNeedToAchieve(
            ipRevenue.target,
            ipRevenue.achieved
        );


    const b2bRevenueYet =
        panIndiaNeedToAchieve(
            b2bRevenue.target,
            b2bRevenue.achieved
        );


    const opRevenueYet =
        panIndiaNeedToAchieve(
            opRevenue.target,
            opRevenue.achieved
        );


    const pharmacyRevenueYet =
        panIndiaNeedToAchieve(
            pharmacyRevenue.target,
            pharmacyRevenue.achieved
        );


    const labRevenueYet =
        panIndiaNeedToAchieve(
            labRevenue.target,
            labRevenue.achieved
        );


    const totalRevenueYet =
        panIndiaNeedToAchieve(
            totalRevenue.target,
            totalRevenue.achieved
        );


    /*
     * Per-unit revenue / discount parameters.
     */
    const perIpRevenueTarget =
        panIndiaSafeDivide(
            ipRevenue.target,
            ipCount.target
        );


    const perIpRevenueGenerated =
        panIndiaSafeDivide(
            ipRevenue.achieved,
            ipCount.achieved
        );


    const discountGivenPerIp =
        panIndiaSignedDifference(
            perIpRevenueTarget,
            perIpRevenueGenerated
        );


    const perB2bRevenueTarget =
        panIndiaSafeDivide(
            b2bRevenue.target,
            b2bCount.target
        );


    const perB2bRevenueGenerated =
        panIndiaSafeDivide(
            b2bRevenue.achieved,
            b2bCount.achieved,
            true
        );


    const discountGivenPerB2b =
        panIndiaSignedDifference(
            perB2bRevenueTarget,
            perB2bRevenueGenerated
        );


    const perOpRevenueTarget =
        panIndiaSafeDivide(
            opRevenue.target,
            opCount.target
        );


    const perOpRevenueGenerated =
        panIndiaSafeDivide(
            opRevenue.achieved,
            opCount.achieved
        );


    const discountGivenPerOp =
        panIndiaSignedDifference(
            perOpRevenueTarget,
            perOpRevenueGenerated
        );


    const pharmacyPerBillRevenue =
        panIndiaSafeDivide(
            pharmacyRevenue.target,
            pharmacy.target
        );


    const pharmacyPerBillRevenueGenerated =
        panIndiaSafeDivide(
            pharmacyRevenue.achieved,
            pharmacy.achieved
        );


    const discountGivenPerBill =
        panIndiaSignedDifference(
            pharmacyPerBillRevenue,
            pharmacyPerBillRevenueGenerated
        );


    const perLabRevenueTarget =
        panIndiaSafeDivide(
            labRevenue.target,
            lab.target
        );


    const perLabRevenueGenerated =
        panIndiaSafeDivide(
            labRevenue.achieved,
            lab.achieved
        );


    const discountGivenPerLab =
        panIndiaSignedDifference(
            perLabRevenueTarget,
            perLabRevenueGenerated
        );


    return [

        /*
         * COUNTS
         */
        {
            section:
                "Count Parameters",

            title:
                "IP Count",

            showStatus:
                true,

            targetValue:
                ipCount.target,

            achievedValue:
                ipCount.achieved,

            fields: [
                panIndiaParameterField(
                    "ipCountTarget",
                    "IP Count Target",
                    ipCount.target,
                    "number"
                ),

                panIndiaParameterField(
                    "ipCountAchieved",
                    "IP Count Achieved",
                    ipCount.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "ipCountUnachieved",
                    "IP Count Unachieved",
                    ipCountUnachieved,
                    "number"
                )
            ]
        },


        {
            section:
                "Count Parameters",

            title:
                "B2B Count",

            showStatus:
                true,

            targetValue:
                b2bCount.target,

            achievedValue:
                b2bCount.achieved,

            fields: [
                panIndiaParameterField(
                    "b2bCountTarget",
                    "B2B Count Target",
                    b2bCount.target,
                    "number"
                ),

                panIndiaParameterField(
                    "b2bCountAchieved",
                    "B2B Count Achieved",
                    b2bCount.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "b2bCountUnachieved",
                    "B2B Count Unachieved",
                    b2bCountUnachieved,
                    "number"
                )
            ]
        },


        {
            section:
                "Count Parameters",

            title:
                "OP Count",

            showStatus:
                true,

            targetValue:
                opCount.target,

            achievedValue:
                opCount.achieved,

            fields: [
                panIndiaParameterField(
                    "opCountTarget",
                    "OP Count Target",
                    opCount.target,
                    "number"
                ),

                panIndiaParameterField(
                    "opCountAchieved",
                    "OP Count Achieved",
                    opCount.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "opCountUnachieved",
                    "OP Count Unachieved",
                    opCountUnachieved,
                    "number"
                )
            ]
        },


        {
            section:
                "Count Parameters",

            title:
                "Pharmacy Bill Count",

            showStatus:
                true,

            targetValue:
                pharmacy.target,

            achievedValue:
                pharmacy.achieved,

            fields: [
                panIndiaParameterField(
                    "pharmacyBillCount",
                    "Pharmacy Bill Count",
                    pharmacy.target,
                    "number"
                ),

                panIndiaParameterField(
                    "pharmacyBillCountReceived",
                    "Pharmacy Bill Count Received",
                    pharmacy.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "pharmacyBillCountUnreceived",
                    "Pharmacy Bill Count Unreceived",
                    pharmacyUnreceived,
                    "number"
                )
            ]
        },


        {
            section:
                "Count Parameters",

            title:
                "Lab Count",

            showStatus:
                true,

            targetValue:
                lab.target,

            achievedValue:
                lab.achieved,

            fields: [
                panIndiaParameterField(
                    "labCount",
                    "Lab Count",
                    lab.target,
                    "number"
                ),

                panIndiaParameterField(
                    "labCountAchieved",
                    "Lab Count Achieved",
                    lab.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "labCountUnachieved",
                    "Lab Count Unachieved",
                    labUnachieved,
                    "number"
                )
            ]
        },


        {
            section:
                "Count Parameters",

            title:
                "Leads",

            showStatus:
                true,

            targetValue:
                leads.target,

            achievedValue:
                leads.achieved,

            fields: [
                panIndiaParameterField(
                    "leadsTarget",
                    "Leads Target",
                    leads.target,
                    "number"
                ),

                panIndiaParameterField(
                    "leadsAchieved",
                    "Leads Achieved",
                    leads.achieved,
                    "number"
                ),

                panIndiaParameterField(
                    "leadsUnachieved",
                    "Leads Unachieved",
                    leadsUnachieved,
                    "number"
                )
            ]
        },


        /*
         * CONVERSIONS
         */
        {
            section:
                "Conversion Parameters",

            title:
                "OP → IP Conversion",

            showStatus:
                true,

            targetValue:
                opToIp.target,

            achievedValue:
                opToIp.achieved,

            fields: [
                panIndiaParameterField(
                    "opToIpConversionTarget",
                    "OP to IP Conversion Target",
                    opToIp.target,
                    "percent"
                ),

                panIndiaParameterField(
                    "opToIpConversionAchieved",
                    "OP to IP Conversion Achieved",
                    opToIp.achieved,
                    "percent"
                )
            ]
        },


        {
            section:
                "Conversion Parameters",

            title:
                "Lead → OP Conversion",

            showStatus:
                true,

            targetValue:
                leadToOp.target,

            achievedValue:
                leadToOp.achieved,

            fields: [
                panIndiaParameterField(
                    "leadToOpConversionTarget",
                    "Lead to OP Conversion Target",
                    leadToOp.target,
                    "percent"
                ),

                panIndiaParameterField(
                    "leadToOpConversionAchieved",
                    "Lead to OP Conversion Achieved",
                    leadToOp.achieved,
                    "percent"
                )
            ]
        },


        /*
         * REVENUE
         */
        {
            section:
                "Revenue Parameters",

            title:
                "IP Revenue",

            showStatus:
                true,

            targetValue:
                ipRevenue.target,

            achievedValue:
                ipRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "ipRevenueTarget",
                    "IP Revenue Target",
                    ipRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "ipRevenueGenerated",
                    "IP Revenue Generated",
                    ipRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "ipRevenueYetToBe",
                    "IP Revenue Yet To Be",
                    ipRevenueYet,
                    "currency"
                )
            ]
        },


        {
            section:
                "Revenue Parameters",

            title:
                "B2B Revenue",

            showStatus:
                true,

            targetValue:
                b2bRevenue.target,

            achievedValue:
                b2bRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "b2bRevenueTarget",
                    "B2B Revenue Target",
                    b2bRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "b2bRevenueAchieved",
                    "B2B Revenue Achieved",
                    b2bRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "b2bRevenueYetToBe",
                    "B2B Revenue Yet To Be",
                    b2bRevenueYet,
                    "currency"
                )
            ]
        },


        {
            section:
                "Revenue Parameters",

            title:
                "OP Revenue",

            showStatus:
                true,

            targetValue:
                opRevenue.target,

            achievedValue:
                opRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "opRevenueTarget",
                    "OP Revenue Target",
                    opRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "opRevenueAchieved",
                    "OP Revenue Achieved",
                    opRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "opRevenueYetToBe",
                    "OP Revenue Yet To Be",
                    opRevenueYet,
                    "currency"
                )
            ]
        },


        {
            section:
                "Revenue Parameters",

            title:
                "Pharmacy Revenue",

            showStatus:
                true,

            targetValue:
                pharmacyRevenue.target,

            achievedValue:
                pharmacyRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "pharmacyRevenueTarget",
                    "Pharmacy Revenue Target",
                    pharmacyRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "pharmacyRevenueAchieved",
                    "Pharmacy Revenue Achieved",
                    pharmacyRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "pharmacyRevenueYetToBe",
                    "Pharmacy Revenue Yet To Be",
                    pharmacyRevenueYet,
                    "currency"
                )
            ]
        },


        {
            section:
                "Revenue Parameters",

            title:
                "Lab Revenue",

            showStatus:
                true,

            targetValue:
                labRevenue.target,

            achievedValue:
                labRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "labRevenueTarget",
                    "Lab Revenue Target",
                    labRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "labRevenueAchieved",
                    "Lab Revenue Achieved",
                    labRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "labRevenueYetToBe",
                    "Lab Revenue Yet To Be",
                    labRevenueYet,
                    "currency"
                )
            ]
        },


        {
            section:
                "Revenue Parameters",

            title:
                "Total Revenue",

            showStatus:
                true,

            targetValue:
                totalRevenue.target,

            achievedValue:
                totalRevenue.achieved,

            fields: [
                panIndiaParameterField(
                    "totalRevenueTarget",
                    "Total Revenue Target",
                    totalRevenue.target,
                    "currency"
                ),

                panIndiaParameterField(
                    "totalRevenueGenerated",
                    "Total Revenue Generated",
                    totalRevenue.achieved,
                    "currency"
                ),

                panIndiaParameterField(
                    "totalRevenueYetToAchieve",
                    "Total Revenue Yet To Achieve",
                    totalRevenueYet,
                    "currency"
                )
            ]
        },


        /*
         * PER-UNIT / DISCOUNT
         */
        {
            section:
                "Per-Unit Revenue & Discount Parameters",

            title:
                "Per IP Revenue",

            showStatus:
                false,

            fields: [
                panIndiaParameterField(
                    "perIpRevenueTarget",
                    "Per IP Revenue Target",
                    perIpRevenueTarget,
                    "currency"
                ),

                panIndiaParameterField(
                    "perIpRevenueGenerated",
                    "Per IP Revenue Generated",
                    perIpRevenueGenerated,
                    "currency"
                ),

                panIndiaParameterField(
                    "discountGivenPerIp",
                    "Discount Given Per IP",
                    discountGivenPerIp,
                    "currency"
                )
            ]
        },


        {
            section:
                "Per-Unit Revenue & Discount Parameters",

            title:
                "B2B Per Revenue",

            showStatus:
                false,

            fields: [
                panIndiaParameterField(
                    "b2bPerRevenue",
                    "B2B Revenue",
                    perB2bRevenueTarget,
                    "currency"
                ),

                panIndiaParameterField(
                    "b2bRevenueGeneratedPerUnit",
                    "B2B Revenue Generated",
                    perB2bRevenueGenerated,
                    "currency"
                ),

                panIndiaParameterField(
                    "discountGivenPerB2b",
                    "Discount Given Per IP-(B2B)",
                    discountGivenPerB2b,
                    "currency"
                )
            ]
        },


        {
            section:
                "Per-Unit Revenue & Discount Parameters",

            title:
                "OP Per Revenue",

            showStatus:
                false,

            fields: [
                panIndiaParameterField(
                    "opPerRevenueTarget",
                    "OP Revenue (Target Per OP)",
                    perOpRevenueTarget,
                    "currency"
                ),

                panIndiaParameterField(
                    "opPerRevenueGenerated",
                    "OP Revenue (Generated Per OP)",
                    perOpRevenueGenerated,
                    "currency"
                ),

                panIndiaParameterField(
                    "discountGivenPerOpRevenue",
                    "Discount Given Per OP Revenue",
                    discountGivenPerOp,
                    "currency"
                )
            ]
        },


        {
            section:
                "Per-Unit Revenue & Discount Parameters",

            title:
                "Pharmacy Per Bill Revenue",

            showStatus:
                false,

            fields: [
                panIndiaParameterField(
                    "pharmacyPerBillRevenue",
                    "Pharmacy Per Bill Revenue",
                    pharmacyPerBillRevenue,
                    "currency"
                ),

                panIndiaParameterField(
                    "pharmacyPerBillRevenueGenerated",
                    "Pharmacy Per Bill Revenue Generated",
                    pharmacyPerBillRevenueGenerated,
                    "currency"
                ),

                panIndiaParameterField(
                    "discountGivenPerBillRevenue",
                    "Discount Given Per Bill Revenue",
                    discountGivenPerBill,
                    "currency"
                )
            ]
        },


        {
            section:
                "Per-Unit Revenue & Discount Parameters",

            title:
                "Per Lab Revenue",

            showStatus:
                false,

            fields: [
                panIndiaParameterField(
                    "perLabRevenueTarget",
                    "Per Lab Revenue Target",
                    perLabRevenueTarget,
                    "currency"
                ),

                panIndiaParameterField(
                    "perLabRevenueGenerated",
                    "Per Lab Revenue Generated",
                    perLabRevenueGenerated,
                    "currency"
                ),

                panIndiaParameterField(
                    "discountGivenPerLabRevenue",
                    "Discount Given Per Lab Revenue",
                    discountGivenPerLab,
                    "currency"
                )
            ]
        }

    ];

}


function flattenPanIndiaParameterCards(
    cards
) {

    const output =
        [];


    (
        Array.isArray(
            cards
        )
            ?
            cards
            :
            []
    )
        .forEach(
            card => {

                (
                    Array.isArray(
                        card.fields
                    )
                        ?
                        card.fields
                        :
                        []
                )
                    .forEach(
                        field => {

                            output.push({

                                section:
                                    card.section
                                    ||
                                    "",

                                cardTitle:
                                    card.title
                                    ||
                                    "",

                                id:
                                    field.id,

                                label:
                                    field.label,

                                value:
                                    field.value,

                                type:
                                    field.type

                            });

                        }
                    );

            }
        );


    return output;

}


function sumPanIndiaValues(
    values
) {

    let total =
        0;


    let found =
        false;


    values.forEach(
        value => {

            const number =
                panIndiaNumber(
                    value
                );


            if (
                number !==
                null
            ) {

                total +=
                    number;


                found =
                    true;

            }

        }
    );


    return found
        ?
        total
        :
        null;

}


function sumPanIndiaMetric(
    states,
    key
) {

    return {

        target:
            sumPanIndiaValues(
                states.map(
                    state =>
                        state[key]
                            ?
                            state[key].target
                            :
                            null
                )
            ),

        achieved:
            sumPanIndiaValues(
                states.map(
                    state =>
                        state[key]
                            ?
                            state[key].achieved
                            :
                            null
                )
            )

    };

}


function buildPanIndiaTotalRevenueMetric(
    source
) {

    const state =
        source
        &&
        typeof source ===
        "object"
            ?
            source
            :
            {};


    const revenueKeys = [
        "ipRevenue",
        "b2bRevenue",
        "opRevenue",
        "pharmacyRevenue",
        "labRevenue"
    ];


    const target =
        sumPanIndiaValues(
            revenueKeys.map(
                key =>
                    state[key]
                        ?
                        state[key].target
                        :
                        null
            )
        );


    const achieved =
        sumPanIndiaValues(
            revenueKeys.map(
                key =>
                    state[key]
                        ?
                        state[key].achieved
                        :
                        null
            )
        );


    return {

        target:
            target,

        achieved:
            achieved,

        remaining:
            panIndiaNeedToAchieve(
                target,
                achieved
            )

    };

}


async function getPanIndiaLiveInput(
    req,
    res
) {

    try {

        const results =
            await mapWithConcurrency(
                PAN_INDIA_LIVE_SOURCES,
                2,
                    async source => {

                        try {

                            const result =
                                await callAppsScriptGet(
                                    "liveInput",
                                    {
                                        key:
                                            source.key,

                                        liveInputKey:
                                            source.key
                                    }
                                );


                            const report =
                                result.report &&
                                typeof result.report ===
                                "object"
                                    ?
                                    result.report
                                    :
                                    {};


                            const revenue =
                                result.revenue &&
                                typeof result.revenue ===
                                "object"
                                    ?
                                    result.revenue
                                    :
                                    {};


                            return {

                                success:
                                    true,

                                key:
                                    source.key,

                                label:
                                    source.label,

                                sourceColumn:
                                    result.sourceColumn
                                    ||
                                    source.key,

                                detectedSourceColumn:
                                    result.detectedSourceColumn
                                    ||
                                    "",

                                ipCount:
                                    normalizePanIndiaMetric(
                                        report.ipCount
                                    ),

                                b2bCount:
                                    normalizePanIndiaMetric(
                                        report.b2bCount
                                    ),

                                opCount:
                                    normalizePanIndiaMetric(
                                        report.opCount
                                    ),

                                leads:
                                    normalizePanIndiaMetric(
                                        report.leads
                                    ),

                                leadToOp:
                                    normalizePanIndiaConversion(
                                        report.leadToOp,
                                        report.opCount,
                                        report.leads
                                    ),

                                opToIp:
                                    normalizePanIndiaConversion(
                                        report.opToIp,
                                        report.ipCount,
                                        report.opCount
                                    ),

                                lab:
                                    normalizePanIndiaMetric(
                                        report.lab
                                    ).target !== null
                                    ||
                                    normalizePanIndiaMetric(
                                        report.lab
                                    ).achieved !== null
                                        ?
                                        normalizePanIndiaMetric(
                                            report.lab
                                        )
                                        :
                                        getPanIndiaLabMetric(
                                            result.parameters
                                        ),

                                pharmacy:
                                    normalizePanIndiaMetric(
                                        report.pharmacy
                                    ).target !== null
                                    ||
                                    normalizePanIndiaMetric(
                                        report.pharmacy
                                    ).achieved !== null
                                        ?
                                        normalizePanIndiaMetric(
                                            report.pharmacy
                                        )
                                        :
                                        getPanIndiaPharmacyMetric(
                                            result.parameters
                                        ),

                                ipRevenue:
                                    resolvePanIndiaRevenueMetric(
                                        revenue.ipRevenue,
                                        result.parameters,
                                        [
                                            "IP Revenue Target"
                                        ],
                                        [
                                            "IP Revenue Generated",
                                            "IP Revenue Achieved"
                                        ],
                                        [
                                            "IP Revenue Yet To Be",
                                            "IP Revenue Yet to Be",
                                            "IP Revenue Need to Achieve"
                                        ]
                                    ),

                                b2bRevenue:
                                    resolvePanIndiaRevenueMetric(
                                        revenue.b2bRevenue,
                                        result.parameters,
                                        [
                                            "B2B Revenue Target"
                                        ],
                                        [
                                            "B2B Revenue Achieved",
                                            "B2B Revenue Generated"
                                        ],
                                        [
                                            "B2B Revenue Yet To Be",
                                            "B2B Revenue Yet to Be",
                                            "B2B Revenue Need to Achieve"
                                        ]
                                    ),

                                opRevenue:
                                    resolvePanIndiaRevenueMetric(
                                        revenue.opRevenue,
                                        result.parameters,
                                        [
                                            "OP Revenue Target"
                                        ],
                                        [
                                            "OP Revenue Achieved",
                                            "OP Revenue Generated"
                                        ],
                                        [
                                            "OP Revenue Yet To Be",
                                            "OP Revenue Yet to Be",
                                            "OP Revenue Need to Achieve"
                                        ]
                                    ),

                                pharmacyRevenue:
                                    resolvePanIndiaRevenueMetric(
                                        revenue.pharmacyRevenue,
                                        result.parameters,
                                        [
                                            "Pharmacy Revenue Target"
                                        ],
                                        [
                                            "Pharmacy Revenue Achieved"
                                        ],
                                        [
                                            "Pharmacy Revenue Yet To Be",
                                            "Pharmacy Revenue Yet to Be",
                                            "Pharmacy Revenue Need to Achieve",
                                            "Pharmacy Revenue Unachieved",
                                            "Pharmacy Revenue Remaining"
                                        ]
                                    ),

                                labRevenue:
                                    resolvePanIndiaRevenueMetric(
                                        revenue.labRevenue,
                                        result.parameters,
                                        [
                                            "Lab Revenue Target"
                                        ],
                                        [
                                            "Lab Revenue Achieved"
                                        ],
                                        [
                                            "Lab Revenue Yet To Be",
                                            "Lab Revenue Yet to Be",
                                            "Lab Revenue Need to Achieve",
                                            "Lab Revenue Unachieved",
                                            "Lab Revenue Remaining"
                                        ]
                                    ),

                                refreshedAt:
                                    result.refreshedAt
                                    ||
                                    null

                            };

                        }
                        catch (
                            error
                        ) {

                            return {

                                success:
                                    false,

                                key:
                                    source.key,

                                label:
                                    source.label,

                                message:
                                    error.message
                                    ||
                                    "Unable to load this source."

                            };

                        }

                    }
            );


        const states =
            results.filter(
                item =>
                    item.success
            );


        states.forEach(
            state => {

                const ipRevenue =
                    state.ipRevenue
                    ||
                    {};


                const b2bRevenue =
                    state.b2bRevenue
                    ||
                    {};


                const opRevenue =
                    state.opRevenue
                    ||
                    {};


                const pharmacyRevenue =
                    state.pharmacyRevenue
                    ||
                    {};


                const labRevenue =
                    state.labRevenue
                    ||
                    {};


                const pharmacy =
                    state.pharmacy
                    ||
                    {};


                const lab =
                    state.lab
                    ||
                    {};


                state.ipRevenue = {

                    ...ipRevenue,

                    remaining:
                        panIndiaNeedToAchieve(
                            ipRevenue.target,
                            ipRevenue.achieved
                        )

                };


                state.b2bRevenue = {

                    ...b2bRevenue,

                    remaining:
                        panIndiaNeedToAchieve(
                            b2bRevenue.target,
                            b2bRevenue.achieved
                        )

                };


                state.opRevenue = {

                    ...opRevenue,

                    remaining:
                        panIndiaNeedToAchieve(
                            opRevenue.target,
                            opRevenue.achieved
                        )

                };


                state.pharmacyRevenue = {

                    ...pharmacyRevenue,

                    remaining:
                        panIndiaNeedToAchieve(
                            pharmacyRevenue.target,
                            pharmacyRevenue.achieved
                        )

                };


                state.labRevenue = {

                    ...labRevenue,

                    remaining:
                        panIndiaNeedToAchieve(
                            labRevenue.target,
                            labRevenue.achieved
                        )

                };



                state.totalRevenue =
                    buildPanIndiaTotalRevenueMetric(
                        state
                    );


                const pharmacyTargetPerBill =
                    panIndiaSafeDivide(
                        pharmacyRevenue.target,
                        pharmacy.target
                    );


                const pharmacyAchievedPerBill =
                    panIndiaSafeDivide(
                        pharmacyRevenue.achieved,
                        pharmacy.achieved
                    );


                state.pharmacyPerBillRevenue = {

                    target:
                        pharmacyTargetPerBill,

                    achieved:
                        pharmacyAchievedPerBill,

                    discount:
                        panIndiaSignedDifference(
                            pharmacyTargetPerBill,
                            pharmacyAchievedPerBill
                        )

                };


                const labTargetPerBill =
                    panIndiaSafeDivide(
                        labRevenue.target,
                        lab.target
                    );


                const labAchievedPerBill =
                    panIndiaSafeDivide(
                        labRevenue.achieved,
                        lab.achieved
                    );


                state.labPerBillRevenue = {

                    target:
                        labTargetPerBill,

                    achieved:
                        labAchievedPerBill,

                    discount:
                        panIndiaSignedDifference(
                            labTargetPerBill,
                            labAchievedPerBill
                        )

                };

            }
        );


        const failures =
            results
                .filter(
                    item =>
                        !item.success
                )
                .map(
                    item => ({

                        key:
                            item.key,

                        label:
                            item.label,

                        message:
                            item.message

                    })
                );


        if (
            states.length ===
            0
        ) {

            return res
                .status(
                    500
                )
                .json({

                    success:
                        false,

                    message:
                        "Unable to load Pan India data from any live source.",

                    failures:
                        failures

                });

        }


        const totals = {

            ipCount:
                sumPanIndiaMetric(
                    states,
                    "ipCount"
                ),

            b2bCount:
                sumPanIndiaMetric(
                    states,
                    "b2bCount"
                ),

            opCount:
                sumPanIndiaMetric(
                    states,
                    "opCount"
                ),

            pharmacy:
                sumPanIndiaMetric(
                    states,
                    "pharmacy"
                ),

            lab:
                sumPanIndiaMetric(
                    states,
                    "lab"
                ),

            leads:
                sumPanIndiaMetric(
                    states,
                    "leads"
                ),

            ipRevenue:
                sumPanIndiaMetric(
                    states,
                    "ipRevenue"
                ),

            b2bRevenue:
                sumPanIndiaMetric(
                    states,
                    "b2bRevenue"
                ),

            opRevenue:
                sumPanIndiaMetric(
                    states,
                    "opRevenue"
                ),

            pharmacyRevenue:
                sumPanIndiaMetric(
                    states,
                    "pharmacyRevenue"
                ),

            labRevenue:
                sumPanIndiaMetric(
                    states,
                    "labRevenue"
                ),

            totalRevenue:
                sumPanIndiaMetric(
                    states,
                    "totalRevenue"
                )

        };


        totals.totalRevenue = {

            ...totals.totalRevenue,

            remaining:
                panIndiaNeedToAchieve(
                    totals.totalRevenue.target,
                    totals.totalRevenue.achieved
                )

        };


        /*
         * Do NOT add state conversion percentages together.
         * Pan India conversion is calculated from consolidated counts.
         *
         * Lead -> OP:
         *   OP / Leads * 100
         *
         * OP -> IP:
         *   IP / OP * 100
         */
        totals.leadToOp = {

            target:
                panIndiaPercentFromCounts(
                    totals.opCount.target,
                    totals.leads.target
                ),

            achieved:
                panIndiaPercentFromCounts(
                    totals.opCount.achieved,
                    totals.leads.achieved
                )

        };


        totals.opToIp = {

            target:
                panIndiaPercentFromCounts(
                    totals.ipCount.target,
                    totals.opCount.target
                ),

            achieved:
                panIndiaPercentFromCounts(
                    totals.ipCount.achieved,
                    totals.opCount.achieved
                )

        };


        const parameterCards =
            buildPanIndiaParameterCards(
                totals
            );


        const allParameters =
            flattenPanIndiaParameterCards(
                parameterCards
            );


        return res.json({

            success:
                true,

            reportKey:
                "PAN_INDIA",

            expectedSources:
                PAN_INDIA_LIVE_SOURCES.length,

            loadedSources:
                states.length,

            partial:
                failures.length > 0,

            totals:
                totals,

            parameterCards:
                parameterCards,

            allParameters:
                allParameters,

            states:
                states,

            failures:
                failures,

            refreshedAt:
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "Pan India live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                message:
                    error.message
                    ||
                    "Unable to load Pan India live input."

            });

    }

}


/* ============================================================
   LEADS LIVE DASHBOARD
============================================================ */

function resolveLeadsIpCountMetric(
    report,
    parameters
) {

    const source =
        report
        &&
        report.ipCount
        &&
        typeof report.ipCount ===
        "object"
            ?
            report.ipCount
            :
            {};


    const raw =
        normalizePanIndiaMetric(
            source
        );


    const parameterMetric =
        getPanIndiaRevenueMetric(
            parameters,
            [
                "IP Count Target"
            ],
            [
                "IP Count Achieved"
            ]
        );


    const achieved =
        firstNumber(
            raw.achieved,
            parameterMetric.achieved
        );


    const remaining =
        firstNumber(
            panIndiaNumber(
                source.remaining
            ),
            panIndiaNumber(
                source.unachieved
            ),
            panIndiaNumber(
                source.needToAchieve
            ),
            getPanIndiaParameter(
                parameters,
                [
                    "IP Count Unachieved",
                    "IP Count Need to Achieve",
                    "IP Count Have To Achieve"
                ]
            )
        );


    const calculatedTarget =
        (
            validNumber(
                achieved
            )
            &&
            validNumber(
                remaining
            )
        )
            ?
            Number(
                achieved
            )
            +
            Number(
                remaining
            )
            :
            null;


    const target =
        firstNumber(
            raw.target,
            parameterMetric.target,
            calculatedTarget
        );


    return {

        target:
            target,

        achieved:
            achieved,

        remaining:
            firstNumber(
                remaining,
                (
                    validNumber(
                        target
                    )
                    &&
                    validNumber(
                        achieved
                    )
                )
                    ?
                    Number(
                        target
                    )
                    -
                    Number(
                        achieved
                    )
                    :
                    null
            )

    };

}


function resolveLeadsIpRevenueMetric(
    report,
    revenue,
    parameters
) {

    const revenueSource =
        revenue
        &&
        revenue.ipRevenue
        &&
        typeof revenue.ipRevenue ===
        "object"
            ?
            revenue.ipRevenue
            :
            {};


    const reportSource =
        report
        &&
        report.ipRevenue
        &&
        typeof report.ipRevenue ===
        "object"
            ?
            report.ipRevenue
            :
            {};


    const revenueMetric =
        normalizePanIndiaMetric(
            revenueSource
        );


    const reportMetric =
        normalizePanIndiaMetric(
            reportSource
        );


    const parameterMetric =
        getPanIndiaRevenueMetric(
            parameters,
            [
                "IP Revenue Target"
            ],
            [
                "IP Revenue Generated",
                "IP Revenue Achieved"
            ]
        );


    const achieved =
        firstNumber(
            revenueMetric.achieved,
            reportMetric.achieved,
            parameterMetric.achieved
        );


    const remaining =
        firstNumber(
            panIndiaNumber(
                revenueSource.remaining
            ),
            panIndiaNumber(
                revenueSource.yetToBe
            ),
            panIndiaNumber(
                revenueSource.needToAchieve
            ),
            panIndiaNumber(
                reportSource.remaining
            ),
            panIndiaNumber(
                reportSource.yetToBe
            ),
            panIndiaNumber(
                reportSource.needToAchieve
            ),
            getPanIndiaParameter(
                parameters,
                [
                    "IP Revenue Yet To Be",
                    "IP Revenue Need to Achieve"
                ]
            )
        );


    /*
     * If the target cell is unavailable:
     *
     * Target = Generated + Yet To Be
     *
     * This mirrors the live Input sheet.
     */
    const calculatedTarget =
        (
            validNumber(
                achieved
            )
            &&
            validNumber(
                remaining
            )
        )
            ?
            Number(
                achieved
            )
            +
            Number(
                remaining
            )
            :
            null;


    const target =
        firstNumber(
            revenueMetric.target,
            reportMetric.target,
            parameterMetric.target,
            calculatedTarget
        );


    return {

        target:
            target,

        achieved:
            achieved,

        remaining:
            firstNumber(
                remaining,
                (
                    validNumber(
                        target
                    )
                    &&
                    validNumber(
                        achieved
                    )
                )
                    ?
                    Number(
                        target
                    )
                    -
                    Number(
                        achieved
                    )
                    :
                    null
            )

    };

}


async function getLeadsLiveInput(
    req,
    res
) {

    try {

        /*
         * IMPORTANT:
         * This Leads dashboard is APTS ONLY.
         *
         * We intentionally call the APTS live source directly.
         * No Pan India source loop is used here.
         */
        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        "APTS",

                    liveInputKey:
                        "APTS"
                }
            );


        const report =
            result.report
            &&
            typeof result.report ===
            "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue ===
            "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters ===
            "object"
                ?
                result.parameters
                :
                {};



        /*
         * FIX:
         * report, revenue and parameters must exist before these helpers run.
         * This prevents:
         * ReferenceError: Cannot access 'report' before initialization
         */
        const ipCount =
            resolveLeadsIpCountMetric(
                report,
                parameters
            );


        const ipRevenue =
            resolveLeadsIpRevenueMetric(
                report,
                revenue,
                parameters
            );


        /*
         * APTS LEADS
         */
        const leads =
            normalizePanIndiaMetric(
                report.leads
            );


        const opCount =
            normalizePanIndiaMetric(
                report.opCount
            );


        /*
         * APTS LEAD -> OP CONVERSION
         *
         * Prefer the actual sheet conversion values.
         * If blank, normalizePanIndiaConversion() derives them
         * from the APTS OP and Leads counts.
         */
        const leadToOp =
            normalizePanIndiaConversion(
                report.leadToOp,
                report.opCount,
                report.leads
            );


        /*
         * APTS OP REVENUE
         *
         * Read it robustly from all supported live-response locations:
         * 1) result.revenue.opRevenue      (current Apps Script)
         * 2) result.report.opRevenue       (older/alternate response)
         * 3) exact parameter map           (safe fallback)
         */
        const revenueMetric =
            normalizePanIndiaMetric(
                revenue.opRevenue
            );


        const reportRevenueMetric =
            normalizePanIndiaMetric(
                report.opRevenue
            );


        const parameterRevenueMetric =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "OP Revenue Target"
                ],
                [
                    "OP Revenue Achieved"
                ]
            );


        const opRevenue = {

            target:
                firstNumber(
                    revenueMetric.target,
                    reportRevenueMetric.target,
                    parameterRevenueMetric.target
                ),

            achieved:
                firstNumber(
                    revenueMetric.achieved,
                    reportRevenueMetric.achieved,
                    parameterRevenueMetric.achieved
                )

        };


        return res.json({

            success:
                true,

            reportKey:
                "LEADS_APTS",

            sourceKey:
                "APTS",

            sourceColumn:
                result.sourceColumn
                ||
                "APTS",

            totals: {

                ipCount:
                    ipCount,

                ipRevenue:
                    ipRevenue,

                leads:
                    leads,

                opCount:
                    opCount,

                leadToOp:
                    leadToOp,

                opRevenue:
                    opRevenue

            },

            failures:
                [],

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "APTS Leads live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    "LEADS_APTS",

                sourceKey:
                    "APTS",

                message:
                    error.message
                    ||
                    "Unable to load APTS Leads live input."

            });

    }

}


/* ============================================================
   LEADS REPORT WRAPPERS
============================================================ */

function getLeadsGeneratedReports(req, res) {
    return getGeneratedReports(
        REPORT_CONFIGS.leads,
        req,
        res
    );
}

function clearLeadsGeneratedReports(req, res) {
    return clearGeneratedReports(
        REPORT_CONFIGS.leads,
        req,
        res
    );
}

function saveLeadsReportToDrive(req, res) {
    return saveReport(
        REPORT_CONFIGS.leads,
        req,
        res
    );
}

function sendLeadsReportEmail(req, res) {
    return sendReportEmail(
        REPORT_CONFIGS.leads,
        req,
        res
    );
}


/* ============================================================
   KARNATAKA LEADS LIVE DASHBOARD
============================================================ */

async function getKarnatakaLeadsLiveInput(
    req,
    res
) {

    try {

        /*
         * KARNATAKA ONLY.
         * No Pan India aggregation is used here.
         */
        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        "KA",

                    liveInputKey:
                        "KA"
                }
            );


        const report =
            result.report
            &&
            typeof result.report ===
            "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue ===
            "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters ===
            "object"
                ?
                result.parameters
                :
                {};

        const ipCount =
            resolveLeadsIpCountMetric(
                report,
                parameters
            );


        const ipRevenue =
            resolveLeadsIpRevenueMetric(
                report,
                revenue,
                parameters
            );





        const leads =
            normalizePanIndiaMetric(
                report.leads
            );


        const opCount =
            normalizePanIndiaMetric(
                report.opCount
            );


        const leadToOp =
            normalizePanIndiaConversion(
                report.leadToOp,
                report.opCount,
                report.leads
            );


        /*
         * OP Revenue fallback order:
         * 1) revenue.opRevenue
         * 2) report.opRevenue
         * 3) exact parameter map
         */
        const revenueMetric =
            normalizePanIndiaMetric(
                revenue.opRevenue
            );


        const reportRevenueMetric =
            normalizePanIndiaMetric(
                report.opRevenue
            );


        const parameterRevenueMetric =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "OP Revenue Target"
                ],
                [
                    "OP Revenue Achieved"
                ]
            );


        const opRevenue = {

            target:
                firstNumber(
                    revenueMetric.target,
                    reportRevenueMetric.target,
                    parameterRevenueMetric.target
                ),

            achieved:
                firstNumber(
                    revenueMetric.achieved,
                    reportRevenueMetric.achieved,
                    parameterRevenueMetric.achieved
                )

        };


        return res.json({

            success:
                true,

            reportKey:
                "LEADS_KA",

            sourceKey:
                "KA",

            sourceColumn:
                result.sourceColumn
                ||
                "Karnataka",

            totals: {

                ipCount:
                    ipCount,

                ipRevenue:
                    ipRevenue,

                leads:
                    leads,

                opCount:
                    opCount,

                leadToOp:
                    leadToOp,

                opRevenue:
                    opRevenue

            },

            failures:
                [],

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "Karnataka Leads live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    "LEADS_KA",

                sourceKey:
                    "KA",

                message:
                    error.message
                    ||
                    "Unable to load Karnataka Leads live input."

            });

    }

}


/* ============================================================
   KARNATAKA LEADS REPORT WRAPPERS
============================================================ */

function getKarnatakaLeadsGeneratedReports(
    req,
    res
) {

    return getGeneratedReports(
        REPORT_CONFIGS.leadsKa,
        req,
        res
    );

}


function clearKarnatakaLeadsGeneratedReports(
    req,
    res
) {

    return clearGeneratedReports(
        REPORT_CONFIGS.leadsKa,
        req,
        res
    );

}


function saveKarnatakaLeadsReportToDrive(
    req,
    res
) {

    return saveReport(
        REPORT_CONFIGS.leadsKa,
        req,
        res
    );

}


function sendKarnatakaLeadsReportEmail(
    req,
    res
) {

    return sendReportEmail(
        REPORT_CONFIGS.leadsKa,
        req,
        res
    );

}


async function getTamilNaduLeadsLiveInput(
    req,
    res
) {

    try {

        /*
         * TAMIL NADU ONLY.
         * No Pan India aggregation is used here.
         */
        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        "TN",

                    liveInputKey:
                        "TN"
                }
            );


        const report =
            result.report
            &&
            typeof result.report ===
            "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue ===
            "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters ===
            "object"
                ?
                result.parameters
                :
                {};

        const ipCount =
            resolveLeadsIpCountMetric(
                report,
                parameters
            );


        const ipRevenue =
            resolveLeadsIpRevenueMetric(
                report,
                revenue,
                parameters
            );





        const leads =
            normalizePanIndiaMetric(
                report.leads
            );


        const opCount =
            normalizePanIndiaMetric(
                report.opCount
            );


        const leadToOp =
            normalizePanIndiaConversion(
                report.leadToOp,
                report.opCount,
                report.leads
            );


        /*
         * OP Revenue fallback order:
         * 1) revenue.opRevenue
         * 2) report.opRevenue
         * 3) exact parameter map
         */
        const revenueMetric =
            normalizePanIndiaMetric(
                revenue.opRevenue
            );


        const reportRevenueMetric =
            normalizePanIndiaMetric(
                report.opRevenue
            );


        const parameterRevenueMetric =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "OP Revenue Target"
                ],
                [
                    "OP Revenue Achieved"
                ]
            );


        const opRevenue = {

            target:
                firstNumber(
                    revenueMetric.target,
                    reportRevenueMetric.target,
                    parameterRevenueMetric.target
                ),

            achieved:
                firstNumber(
                    revenueMetric.achieved,
                    reportRevenueMetric.achieved,
                    parameterRevenueMetric.achieved
                )

        };


        return res.json({

            success:
                true,

            reportKey:
                "LEADS_TN",

            sourceKey:
                "TN",

            sourceColumn:
                result.sourceColumn
                ||
                "Tamil Nadu",

            totals: {

                ipCount:
                    ipCount,

                ipRevenue:
                    ipRevenue,

                leads:
                    leads,

                opCount:
                    opCount,

                leadToOp:
                    leadToOp,

                opRevenue:
                    opRevenue

            },

            failures:
                [],

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "Tamil Nadu Leads live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    "LEADS_TN",

                sourceKey:
                    "TN",

                message:
                    error.message
                    ||
                    "Unable to load Tamil Nadu Leads live input."

            });

    }

}


/* ============================================================
   TAMIL NADU LEADS REPORT WRAPPERS
============================================================ */

function getTamilNaduLeadsGeneratedReports(
    req,
    res
) {

    return getGeneratedReports(
        REPORT_CONFIGS.leadsTn,
        req,
        res
    );

}


function clearTamilNaduLeadsGeneratedReports(
    req,
    res
) {

    return clearGeneratedReports(
        REPORT_CONFIGS.leadsTn,
        req,
        res
    );

}


function saveTamilNaduLeadsReportToDrive(
    req,
    res
) {

    return saveReport(
        REPORT_CONFIGS.leadsTn,
        req,
        res
    );

}


function sendTamilNaduLeadsReportEmail(
    req,
    res
) {

    return sendReportEmail(
        REPORT_CONFIGS.leadsTn,
        req,
        res
    );

}




async function getWestBengalLeadsLiveInput(
    req,
    res
) {

    try {

        /*
         * WEST BENGAL ONLY.
         * No Pan India aggregation is used here.
         */
        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        "WB",

                    liveInputKey:
                        "WB"
                }
            );


        const report =
            result.report
            &&
            typeof result.report ===
            "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue ===
            "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters ===
            "object"
                ?
                result.parameters
                :
                {};

        const ipCount =
            resolveLeadsIpCountMetric(
                report,
                parameters
            );


        const ipRevenue =
            resolveLeadsIpRevenueMetric(
                report,
                revenue,
                parameters
            );





        const leads =
            normalizePanIndiaMetric(
                report.leads
            );


        const opCount =
            normalizePanIndiaMetric(
                report.opCount
            );


        const leadToOp =
            normalizePanIndiaConversion(
                report.leadToOp,
                report.opCount,
                report.leads
            );


        /*
         * OP Revenue fallback order:
         * 1) revenue.opRevenue
         * 2) report.opRevenue
         * 3) exact parameter map
         */
        const revenueMetric =
            normalizePanIndiaMetric(
                revenue.opRevenue
            );


        const reportRevenueMetric =
            normalizePanIndiaMetric(
                report.opRevenue
            );


        const parameterRevenueMetric =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "OP Revenue Target"
                ],
                [
                    "OP Revenue Achieved"
                ]
            );


        const opRevenue = {

            target:
                firstNumber(
                    revenueMetric.target,
                    reportRevenueMetric.target,
                    parameterRevenueMetric.target
                ),

            achieved:
                firstNumber(
                    revenueMetric.achieved,
                    reportRevenueMetric.achieved,
                    parameterRevenueMetric.achieved
                )

        };


        return res.json({

            success:
                true,

            reportKey:
                "LEADS_WB",

            sourceKey:
                "WB",

            sourceColumn:
                result.sourceColumn
                ||
                "West Bengal",

            totals: {

                ipCount:
                    ipCount,

                ipRevenue:
                    ipRevenue,

                leads:
                    leads,

                opCount:
                    opCount,

                leadToOp:
                    leadToOp,

                opRevenue:
                    opRevenue

            },

            failures:
                [],

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "West Bengal Leads live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    "LEADS_WB",

                sourceKey:
                    "WB",

                message:
                    error.message
                    ||
                    "Unable to load West Bengal Leads live input."

            });

    }

}


/* ============================================================
   WEST BENGAL LEADS REPORT WRAPPERS
============================================================ */

function getWestBengalLeadsGeneratedReports(
    req,
    res
) {

    return getGeneratedReports(
        REPORT_CONFIGS.leadsWb,
        req,
        res
    );

}


function clearWestBengalLeadsGeneratedReports(
    req,
    res
) {

    return clearGeneratedReports(
        REPORT_CONFIGS.leadsWb,
        req,
        res
    );

}


function saveWestBengalLeadsReportToDrive(
    req,
    res
) {

    return saveReport(
        REPORT_CONFIGS.leadsWb,
        req,
        res
    );

}


function sendWestBengalLeadsReportEmail(
    req,
    res
) {

    return sendReportEmail(
        REPORT_CONFIGS.leadsWb,
        req,
        res
    );

}




async function getMaharashtraLeadsLiveInput(
    req,
    res
) {

    try {

        /*
         * MAHARASHTRA ONLY.
         * No Pan India aggregation is used here.
         */
        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        "MH",

                    liveInputKey:
                        "MH"
                }
            );


        const report =
            result.report
            &&
            typeof result.report ===
            "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue ===
            "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters ===
            "object"
                ?
                result.parameters
                :
                {};

        const ipCount =
            resolveLeadsIpCountMetric(
                report,
                parameters
            );


        const ipRevenue =
            resolveLeadsIpRevenueMetric(
                report,
                revenue,
                parameters
            );





        const leads =
            normalizePanIndiaMetric(
                report.leads
            );


        const opCount =
            normalizePanIndiaMetric(
                report.opCount
            );


        const leadToOp =
            normalizePanIndiaConversion(
                report.leadToOp,
                report.opCount,
                report.leads
            );


        /*
         * OP Revenue fallback order:
         * 1) revenue.opRevenue
         * 2) report.opRevenue
         * 3) exact parameter map
         */
        const revenueMetric =
            normalizePanIndiaMetric(
                revenue.opRevenue
            );


        const reportRevenueMetric =
            normalizePanIndiaMetric(
                report.opRevenue
            );


        const parameterRevenueMetric =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "OP Revenue Target"
                ],
                [
                    "OP Revenue Achieved"
                ]
            );


        const opRevenue = {

            target:
                firstNumber(
                    revenueMetric.target,
                    reportRevenueMetric.target,
                    parameterRevenueMetric.target
                ),

            achieved:
                firstNumber(
                    revenueMetric.achieved,
                    reportRevenueMetric.achieved,
                    parameterRevenueMetric.achieved
                )

        };


        return res.json({

            success:
                true,

            reportKey:
                "LEADS_MH",

            sourceKey:
                "MH",

            sourceColumn:
                result.sourceColumn
                ||
                "Maharashtra",

            totals: {

                ipCount:
                    ipCount,

                ipRevenue:
                    ipRevenue,

                leads:
                    leads,

                opCount:
                    opCount,

                leadToOp:
                    leadToOp,

                opRevenue:
                    opRevenue

            },

            failures:
                [],

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            "Maharashtra Leads live input error:",
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    "LEADS_MH",

                sourceKey:
                    "MH",

                message:
                    error.message
                    ||
                    "Unable to load Maharashtra Leads live input."

            });

    }

}


/* ============================================================
   MAHARASHTRA LEADS REPORT WRAPPERS
============================================================ */

function getMaharashtraLeadsGeneratedReports(
    req,
    res
) {

    return getGeneratedReports(
        REPORT_CONFIGS.leadsMh,
        req,
        res
    );

}


function clearMaharashtraLeadsGeneratedReports(
    req,
    res
) {

    return clearGeneratedReports(
        REPORT_CONFIGS.leadsMh,
        req,
        res
    );

}


function saveMaharashtraLeadsReportToDrive(
    req,
    res
) {

    return saveReport(
        REPORT_CONFIGS.leadsMh,
        req,
        res
    );

}


function sendMaharashtraLeadsReportEmail(
    req,
    res
) {

    return sendReportEmail(
        REPORT_CONFIGS.leadsMh,
        req,
        res
    );

}




/* ============================================================
   AP & TS WRAPPERS
============================================================ */

function testAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function getAptsLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function getAptsGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function clearAptsGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function saveAptsIpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function sendAptsIpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.apts,
        req,
        res
    );
}


function getAptsIpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.apts;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `AP & TS IP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `AP & TS IP requires ${minimumRecipients} unique recipients. Configure APTS_IP_RECIPIENTS or APTS_IP_RECIPIENT + APTS_IP_RECIPIENT_2 in back-end/.env, then restart Node.`

    });

}


/* ============================================================
   KARNATAKA WRAPPERS
============================================================ */

function testKaAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


function getKaLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


function getKaGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


function clearKaGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


function saveKaIpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


function sendKaIpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.ka,
        req,
        res
    );
}


/* ============================================================
   TAMIL NADU / TN WRAPPERS
============================================================ */

function testTnAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


function getTnLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


function getTnGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


function clearTnGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


function saveTnIpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


function sendTnIpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.tn,
        req,
        res
    );
}


/* ============================================================
   WEST BENGAL / WB WRAPPERS
============================================================ */

function testWbAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function getWbLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function getWbGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function clearWbGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function saveWbIpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function sendWbIpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.wb,
        req,
        res
    );
}


function getWbIpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.wb;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `West Bengal IP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `West Bengal IP requires ${minimumRecipients} unique recipients. Configure WB_IP_RECIPIENTS or WB_IP_RECIPIENT + WB_IP_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}


/* ============================================================
   MAHARASHTRA / MH WRAPPERS
============================================================ */

function testMhAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function getMhLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function getMhGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function clearMhGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function saveMhIpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function sendMhIpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.mh,
        req,
        res
    );
}


function getMhIpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.mh;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Maharashtra IP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Maharashtra IP requires ${minimumRecipients} unique recipients. Configure MH_IP_RECIPIENTS or MH_IP_RECIPIENT + MH_IP_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}


/* ============================================================
   AP & TS OP REPORT WRAPPERS
============================================================ */
function testAptsOpAppsScriptWebApp(req,res){return testReportConnection(REPORT_CONFIGS.opApts,req,res);}
function getAptsOpLiveInput(req,res){return getLiveInput(REPORT_CONFIGS.opApts,req,res);}
function getAptsOpGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.opApts,req,res);}
function clearAptsOpGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.opApts,req,res);}
function saveAptsOpReportToDrive(req,res){return saveReport(REPORT_CONFIGS.opApts,req,res);}
function sendAptsOpReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.opApts,req,res);}


function getAptsOpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.opApts;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            1
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                "AP & TS OP email recipient configuration is valid."
                :
                "No AP & TS OP email recipient was loaded. Set APTS_OP_RECIPIENT or APTS_OP_RECIPIENTS in back-end/.env, save it, then restart Node."

    });

}

/* ============================================================
   KARNATAKA OP REPORT WRAPPERS
============================================================ */

function testKaOpAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function getKaOpLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function getKaOpGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function clearKaOpGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function saveKaOpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function sendKaOpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.opKa,
        req,
        res
    );
}


function getKaOpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.opKa;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        recipients:
            status.recipients,

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Karnataka OP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Karnataka OP requires ${minimumRecipients} unique recipients. Configure KA_OP_RECIPIENTS or KA_OP_RECIPIENT + KA_OP_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}


/* ============================================================
   TAMIL NADU OP REPORT WRAPPERS
============================================================ */

function testTnOpAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function getTnOpLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function getTnOpGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function clearTnOpGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function saveTnOpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function sendTnOpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.opTn,
        req,
        res
    );
}


function getTnOpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.opTn;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        recipients:
            status.recipients,

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Tamil Nadu OP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Tamil Nadu OP requires ${minimumRecipients} unique recipients. Configure TN_OP_RECIPIENTS or TN_OP_RECIPIENT + TN_OP_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}


/* ============================================================
   WEST BENGAL OP REPORT WRAPPERS
============================================================ */

function testWbOpAppsScriptWebApp(
    req,
    res
) {
    return testReportConnection(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function getWbOpLiveInput(
    req,
    res
) {
    return getLiveInput(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function getWbOpGeneratedReports(
    req,
    res
) {
    return getGeneratedReports(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function clearWbOpGeneratedReports(
    req,
    res
) {
    return clearGeneratedReports(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function saveWbOpReportToDrive(
    req,
    res
) {
    return saveReport(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function sendWbOpReportEmail(
    req,
    res
) {
    return sendReportEmail(
        REPORT_CONFIGS.opWb,
        req,
        res
    );
}


function getWbOpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.opWb;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `West Bengal OP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `West Bengal OP requires ${minimumRecipients} unique recipients. Configure WB_OP_RECIPIENTS or WB_OP_RECIPIENT + WB_OP_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


/* ============================================================
   MAHARASHTRA OP REPORT WRAPPERS
============================================================ */

function testMhOpAppsScriptWebApp(req, res) {
    return testReportConnection(REPORT_CONFIGS.opMh, req, res);
}

function getMhOpLiveInput(req, res) {
    return getLiveInput(REPORT_CONFIGS.opMh, req, res);
}

function getMhOpGeneratedReports(req, res) {
    return getGeneratedReports(REPORT_CONFIGS.opMh, req, res);
}

function clearMhOpGeneratedReports(req, res) {
    return clearGeneratedReports(REPORT_CONFIGS.opMh, req, res);
}

function saveMhOpReportToDrive(req, res) {
    return saveReport(REPORT_CONFIGS.opMh, req, res);
}

function sendMhOpReportEmail(req, res) {
    return sendReportEmail(REPORT_CONFIGS.opMh, req, res);
}


function getMhOpEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.opMh;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Maharashtra OP email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Maharashtra OP requires ${minimumRecipients} unique recipients. Configure MH_OP_RECIPIENTS or MH_OP_RECIPIENT + MH_OP_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}



/* ============================================================
   LAB LIVE INPUT + WRAPPERS
============================================================ */
function labDifference(explicitValue,target,achieved){
    const exact=panIndiaNumber(explicitValue);
    if(exact!==null)return exact;
    const t=panIndiaNumber(target),a=panIndiaNumber(achieved);
    return (t===null||a===null)?null:(t-a);
}
async function getLabLiveInput(config, req, res) {

    try {

        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key:
                        config.liveInputKey,

                    liveInputKey:
                        config.liveInputKey
                }
            );


        const report =
            result.report
            &&
            typeof result.report === "object"
                ?
                result.report
                :
                {};


        const revenue =
            result.revenue
            &&
            typeof result.revenue === "object"
                ?
                result.revenue
                :
                {};


        const parameters =
            result.parameters
            &&
            typeof result.parameters === "object"
                ?
                result.parameters
                :
                {};


        /* =====================================================
           LAB COUNT
        ===================================================== */

        const labCount =
            normalizePanIndiaMetric(
                report.lab
            );


        const labCountUnachieved =
            labDifference(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Lab Count Unachieved",
                        "Lab Count Unreceived"
                    ]
                ),
                labCount.target,
                labCount.achieved
            );


        /* =====================================================
           LAB REVENUE - ROBUST RESOLUTION

           Supports sheet layouts where there is NO explicit
           "Lab Revenue Target" row and only these rows exist:

           Lab Revenue Achieved
           Lab Revenue Yet To Be

           Karnataka current values:
           Achieved = 12,860
           Yet To Be = 9,640

           Therefore:
           Target = 12,860 + 9,640 = 22,500
        ===================================================== */

        const rawLabRevenue =
            normalizePanIndiaMetric(
                revenue.labRevenue
            );


        const labRevenueAchieved =
            firstNumber(
                rawLabRevenue.achieved,

                getPanIndiaParameter(
                    parameters,
                    [
                        "Lab Revenue Achieved"
                    ]
                )
            );


        const explicitLabRevenueNeed =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Lab Revenue Need to Achieve",
                        "Lab Revenue Yet To Be",
                        "Lab Revenue Yet to Be",
                        "Lab Revenue Unachieved",
                        "Lab Revenue Remaining"
                    ]
                )
            );


        /*
         * Per Lab Revenue Target can also reconstruct
         * Lab Revenue Target:
         *
         * Per Lab Target × Lab Count Target
         *
         * Karnataka:
         * 2,250 × 10 = 22,500
         */
        const explicitPerLabRevenueTarget =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Per Lab Revenue Target"
                    ]
                )
            );


        const labRevenueTarget =
            firstNumber(
                rawLabRevenue.target,

                getPanIndiaParameter(
                    parameters,
                    [
                        "Lab Revenue Target"
                    ]
                ),

                (
                    validNumber(
                        labRevenueAchieved
                    )
                    &&
                    validNumber(
                        explicitLabRevenueNeed
                    )
                )
                    ?
                    (
                        Number(
                            labRevenueAchieved
                        )
                        +
                        Number(
                            explicitLabRevenueNeed
                        )
                    )
                    :
                    null,

                (
                    validNumber(
                        explicitPerLabRevenueTarget
                    )
                    &&
                    validNumber(
                        labCount.target
                    )
                )
                    ?
                    (
                        Number(
                            explicitPerLabRevenueTarget
                        )
                        *
                        Number(
                            labCount.target
                        )
                    )
                    :
                    null
            );


        const labRevenueNeedToAchieve =
            firstNumber(
                explicitLabRevenueNeed,

                (
                    validNumber(
                        labRevenueTarget
                    )
                    &&
                    validNumber(
                        labRevenueAchieved
                    )
                )
                    ?
                    (
                        Number(
                            labRevenueTarget
                        )
                        -
                        Number(
                            labRevenueAchieved
                        )
                    )
                    :
                    null
            );


        /* =====================================================
           PER LAB REVENUE
        ===================================================== */

        const perLabRevenueTarget =
            firstNumber(
                explicitPerLabRevenueTarget,

                divideOrNull(
                    labRevenueTarget,
                    labCount.target
                )
            );


        const perLabRevenueGenerated =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Per Lab Revenue Generated",
                        "Per Lab Revenue Achieved"
                    ]
                ),

                divideOrNull(
                    labRevenueAchieved,
                    labCount.achieved
                )
            );


        const discountPerLabRevenue =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Discount Given Per Lab Revenue",
                        "Discount Per Lab"
                    ]
                ),

                (
                    validNumber(
                        perLabRevenueTarget
                    )
                    &&
                    validNumber(
                        perLabRevenueGenerated
                    )
                )
                    ?
                    (
                        Number(
                            perLabRevenueTarget
                        )
                        -
                        Number(
                            perLabRevenueGenerated
                        )
                    )
                    :
                    null
            );


        return res.json({

            success:
                true,

            reportKey:
                config.key,

            sourceKey:
                config.liveInputKey,

            sourceColumn:
                result.sourceColumn
                ||
                config.sourceLabel
                ||
                config.liveInputKey,

            totals: {

                labCount: {

                    target:
                        labCount.target,

                    achieved:
                        labCount.achieved,

                    unachieved:
                        labCountUnachieved

                },


                labRevenue: {

                    target:
                        labRevenueTarget,

                    achieved:
                        labRevenueAchieved,

                    needToAchieve:
                        labRevenueNeedToAchieve,

                    /*
                     * Old frontend compatibility.
                     */
                    unachieved:
                        labRevenueNeedToAchieve

                },


                perLabRevenue: {

                    target:
                        perLabRevenueTarget,

                    generated:
                        perLabRevenueGenerated,

                    discount:
                        discountPerLabRevenue

                }

            },

            /*
             * Useful while verifying state-specific live data.
             */
            labRevenueResolution: {

                rawTarget:
                    rawLabRevenue.target,

                parameterTarget:
                    getPanIndiaParameter(
                        parameters,
                        [
                            "Lab Revenue Target"
                        ]
                    ),

                achieved:
                    labRevenueAchieved,

                sheetNeedToAchieve:
                    explicitLabRevenueNeed,

                perLabRevenueTarget:
                    explicitPerLabRevenueTarget,

                calculatedTarget:
                    labRevenueTarget,

                calculatedNeedToAchieve:
                    labRevenueNeedToAchieve

            },

            expectedSources:
                1,

            loadedSources:
                1,

            partial:
                false,

            inputSpreadsheetName:
                result.inputSpreadsheetName
                ||
                "Input",

            inputSheetName:
                result.inputSheetName
                ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt
                ||
                new Date()
                    .toISOString()

        });

    }
    catch (
        error
    ) {

        console.error(
            `${config.key} Lab live input error:`,
            error
        );


        return res
            .status(
                500
            )
            .json({

                success:
                    false,

                reportKey:
                    config.key,

                sourceKey:
                    config.liveInputKey,

                message:
                    error.message
                    ||
                    `Unable to load ${config.title} live input.`

            });

    }

}

function getAptsLabLiveInput(req,res){return getLabLiveInput(REPORT_CONFIGS.labApts,req,res);}
function getAptsLabGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.labApts,req,res);}
function clearAptsLabGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.labApts,req,res);}
function saveAptsLabReportToDrive(req,res){return saveReport(REPORT_CONFIGS.labApts,req,res);}
function sendAptsLabReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.labApts,req,res);}


function getAptsLabEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.labApts;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `AP & TS Lab email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `AP & TS Lab requires ${minimumRecipients} unique recipients. Configure LAB_APTS_RECIPIENTS or LAB_APTS_RECIPIENT + LAB_APTS_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}

function getKaLabLiveInput(req,res){return getLabLiveInput(REPORT_CONFIGS.labKa,req,res);}
function getKaLabGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.labKa,req,res);}
function clearKaLabGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.labKa,req,res);}
function saveKaLabReportToDrive(req,res){return saveReport(REPORT_CONFIGS.labKa,req,res);}
function sendKaLabReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.labKa,req,res);}


function getKaLabEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.labKa;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Karnataka Lab email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Karnataka Lab requires ${minimumRecipients} unique recipients. Configure LAB_KA_RECIPIENTS or LAB_KA_RECIPIENT + LAB_KA_RECIPIENT_2 in back-end/.env, save it, and restart Node.`

    });

}

function getTnLabLiveInput(req,res){return getLabLiveInput(REPORT_CONFIGS.labTn,req,res);}
function getTnLabGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.labTn,req,res);}
function clearTnLabGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.labTn,req,res);}
function saveTnLabReportToDrive(req,res){return saveReport(REPORT_CONFIGS.labTn,req,res);}
function sendTnLabReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.labTn,req,res);}


function getTnLabEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.labTn;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Tamil Nadu Lab email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Tamil Nadu Lab requires ${minimumRecipients} unique recipients. Configure LAB_TN_RECIPIENTS or LAB_TN_RECIPIENT + LAB_TN_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}

function getWbLabLiveInput(req,res){return getLabLiveInput(REPORT_CONFIGS.labWb,req,res);}
function getWbLabGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.labWb,req,res);}
function clearWbLabGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.labWb,req,res);}
function saveWbLabReportToDrive(req,res){return saveReport(REPORT_CONFIGS.labWb,req,res);}
function sendWbLabReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.labWb,req,res);}


function getWbLabEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.labWb;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `West Bengal Lab email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `West Bengal Lab requires ${minimumRecipients} unique recipients. Configure LAB_WB_RECIPIENTS or LAB_WB_RECIPIENT + LAB_WB_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}

function getMhLabLiveInput(req,res){return getLabLiveInput(REPORT_CONFIGS.labMh,req,res);}
function getMhLabGeneratedReports(req,res){return getGeneratedReports(REPORT_CONFIGS.labMh,req,res);}
function clearMhLabGeneratedReports(req,res){return clearGeneratedReports(REPORT_CONFIGS.labMh,req,res);}
function saveMhLabReportToDrive(req,res){return saveReport(REPORT_CONFIGS.labMh,req,res);}
function sendMhLabReportEmail(req,res){return sendReportEmail(REPORT_CONFIGS.labMh,req,res);}


function getMhLabEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.labMh;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Maharashtra Lab email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Maharashtra Lab requires ${minimumRecipients} unique recipients. Configure LAB_MH_RECIPIENTS or LAB_MH_RECIPIENT + LAB_MH_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


/* ============================================================
   PHARMACY LIVE INPUT - ALL STATES
============================================================ */

async function getPharmacyLiveInput(config, req, res) {

    try {

        const result =
            await callAppsScriptGet(
                "liveInput",
                {
                    key: config.liveInputKey,
                    liveInputKey: config.liveInputKey
                }
            );


        const report =
            result.report &&
            typeof result.report === "object"
                ? result.report
                : {};


        const revenue =
            result.revenue &&
            typeof result.revenue === "object"
                ? result.revenue
                : {};


        const parameters =
            result.parameters &&
            typeof result.parameters === "object"
                ? result.parameters
                : {};


        /* =====================================================
           PHARMACY BILL COUNT
        ===================================================== */

        const reportPharmacy =
            normalizePanIndiaMetric(
                report.pharmacy
            );


        const parameterPharmacy =
            getPanIndiaPharmacyMetric(
                parameters
            );


        const pharmacyCountTarget =
            firstNumber(
                reportPharmacy.target,
                parameterPharmacy.target
            );


        const pharmacyCountAchieved =
            firstNumber(
                reportPharmacy.achieved,
                parameterPharmacy.achieved
            );


        const pharmacyCountUnreceived =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Pharmacy Bill Count Unreceived",
                        "Pharmacy Count Unreceived"
                    ]
                ),

                (
                    validNumber(pharmacyCountTarget) &&
                    validNumber(pharmacyCountAchieved)
                )
                    ? Number(pharmacyCountTarget) - Number(pharmacyCountAchieved)
                    : null
            );


        /* =====================================================
           PHARMACY REVENUE
        ===================================================== */

        const reportRevenue =
            normalizePanIndiaMetric(
                revenue.pharmacyRevenue
            );


        const parameterRevenue =
            getPanIndiaRevenueMetric(
                parameters,
                [
                    "Pharmacy Revenue Target"
                ],
                [
                    "Pharmacy Revenue Achieved"
                ]
            );


        const pharmacyRevenueAchieved =
            firstNumber(
                reportRevenue.achieved,
                parameterRevenue.achieved
            );


        const explicitRevenueNeed =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Pharmacy Revenue Need to Achieve",
                        "Pharmacy Revenue Yet To Be",
                        "Pharmacy Revenue Yet to Be",
                        "Pharmacy Revenue Unachieved",
                        "Pharmacy Revenue Remaining"
                    ]
                )
            );


        const pharmacyRevenueTarget =
            firstNumber(
                reportRevenue.target,
                parameterRevenue.target,

                (
                    validNumber(pharmacyRevenueAchieved) &&
                    validNumber(explicitRevenueNeed)
                )
                    ? Number(pharmacyRevenueAchieved) + Number(explicitRevenueNeed)
                    : null
            );


        const pharmacyRevenueNeedToAchieve =
            firstNumber(
                explicitRevenueNeed,

                (
                    validNumber(pharmacyRevenueTarget) &&
                    validNumber(pharmacyRevenueAchieved)
                )
                    ? Number(pharmacyRevenueTarget) - Number(pharmacyRevenueAchieved)
                    : null
            );


        /* =====================================================
           PER PHARMACY BILL REVENUE
        ===================================================== */

        const perBillTarget =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Pharmacy Per Bill Revenue",
                        "Pharmacy Per Bill Revenue Target"
                    ]
                ),

                divideOrNull(
                    pharmacyRevenueTarget,
                    pharmacyCountTarget
                )
            );


        const perBillGenerated =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Pharmacy Per Bill Revenue Generated",
                        "Pharmacy Per Bill Revenue Achieved"
                    ]
                ),

                divideOrNull(
                    pharmacyRevenueAchieved,
                    pharmacyCountAchieved
                )
            );


        const discountPerBill =
            firstNumber(
                getPanIndiaParameter(
                    parameters,
                    [
                        "Discount Given Per Bill Revenue",
                        "Discount Per Pharmacy Bill"
                    ]
                ),

                (
                    validNumber(perBillTarget) &&
                    validNumber(perBillGenerated)
                )
                    ? Number(perBillTarget) - Number(perBillGenerated)
                    : null
            );


        return res.json({

            success: true,

            reportKey:
                config.key,

            sourceKey:
                config.liveInputKey,

            sourceColumn:
                result.sourceColumn ||
                config.sourceLabel ||
                config.liveInputKey,

            totals: {

                pharmacyCount: {
                    target: pharmacyCountTarget,
                    achieved: pharmacyCountAchieved,
                    unreceived: pharmacyCountUnreceived
                },

                pharmacyRevenue: {
                    target: pharmacyRevenueTarget,
                    achieved: pharmacyRevenueAchieved,
                    needToAchieve: pharmacyRevenueNeedToAchieve
                },

                pharmacyPerBillRevenue: {
                    target: perBillTarget,
                    generated: perBillGenerated,
                    discount: discountPerBill
                }

            },

            expectedSources: 1,
            loadedSources: 1,
            partial: false,

            inputSpreadsheetName:
                result.inputSpreadsheetName ||
                "Input",

            inputSheetName:
                result.inputSheetName ||
                "Sheet1",

            refreshedAt:
                result.refreshedAt ||
                new Date().toISOString()

        });

    }
    catch (error) {

        console.error(
            `${config.key} Pharmacy live input error:`,
            error
        );

        return res.status(500).json({

            success: false,

            reportKey:
                config.key,

            sourceKey:
                config.liveInputKey,

            message:
                error.message ||
                `Unable to load ${config.title} live input.`

        });

    }

}


/* ============================================================
   PHARMACY WRAPPERS
============================================================ */

function getAptsPharmacyLiveInput(req,res){
    return getPharmacyLiveInput(REPORT_CONFIGS.pharmacyApts,req,res);
}
function getAptsPharmacyGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.pharmacyApts,req,res);
}
function clearAptsPharmacyGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.pharmacyApts,req,res);
}
function saveAptsPharmacyReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.pharmacyApts,req,res);
}
function sendAptsPharmacyReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.pharmacyApts,req,res);
}


function getAptsPharmacyEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.pharmacyApts;

    const status =
        getRecipientConfigurationStatus(
            config
        );

    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );

    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `AP & TS Pharmacy email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `AP & TS Pharmacy requires ${minimumRecipients} unique recipients. Configure PHARMACY_APTS_RECIPIENTS or PHARMACY_APTS_RECIPIENT + PHARMACY_APTS_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getKaPharmacyLiveInput(req,res){
    return getPharmacyLiveInput(REPORT_CONFIGS.pharmacyKa,req,res);
}
function getKaPharmacyGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.pharmacyKa,req,res);
}
function clearKaPharmacyGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.pharmacyKa,req,res);
}
function saveKaPharmacyReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.pharmacyKa,req,res);
}
function sendKaPharmacyReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.pharmacyKa,req,res);
}


function getKaPharmacyEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.pharmacyKa;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Karnataka Pharmacy email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Karnataka Pharmacy requires ${minimumRecipients} unique recipients. Configure PHARMACY_KA_RECIPIENTS or PHARMACY_KA_RECIPIENT + PHARMACY_KA_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getTnPharmacyLiveInput(req,res){
    return getPharmacyLiveInput(REPORT_CONFIGS.pharmacyTn,req,res);
}
function getTnPharmacyGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.pharmacyTn,req,res);
}
function clearTnPharmacyGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.pharmacyTn,req,res);
}
function saveTnPharmacyReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.pharmacyTn,req,res);
}
function sendTnPharmacyReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.pharmacyTn,req,res);
}


function getTnPharmacyEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.pharmacyTn;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Tamil Nadu Pharmacy email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Tamil Nadu Pharmacy requires ${minimumRecipients} unique recipients. Configure PHARMACY_TN_RECIPIENTS or PHARMACY_TN_RECIPIENT + PHARMACY_TN_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getWbPharmacyLiveInput(req,res){
    return getPharmacyLiveInput(REPORT_CONFIGS.pharmacyWb,req,res);
}
function getWbPharmacyGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.pharmacyWb,req,res);
}
function clearWbPharmacyGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.pharmacyWb,req,res);
}
function saveWbPharmacyReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.pharmacyWb,req,res);
}
function sendWbPharmacyReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.pharmacyWb,req,res);
}


function getWbPharmacyEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.pharmacyWb;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `West Bengal Pharmacy email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `West Bengal Pharmacy requires ${minimumRecipients} unique recipients. Configure PHARMACY_WB_RECIPIENTS or PHARMACY_WB_RECIPIENT + PHARMACY_WB_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getMhPharmacyLiveInput(req,res){
    return getPharmacyLiveInput(REPORT_CONFIGS.pharmacyMh,req,res);
}
function getMhPharmacyGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.pharmacyMh,req,res);
}
function clearMhPharmacyGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.pharmacyMh,req,res);
}
function saveMhPharmacyReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.pharmacyMh,req,res);
}
function sendMhPharmacyReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.pharmacyMh,req,res);
}


function getMhPharmacyEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.pharmacyMh;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Maharashtra Pharmacy email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Maharashtra Pharmacy requires ${minimumRecipients} unique recipients. Configure PHARMACY_MH_RECIPIENTS or PHARMACY_MH_RECIPIENT + PHARMACY_MH_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}



/* ============================================================
   HR PORTAL - POSITION CARDS
============================================================ */

const HR_SOURCE_CONFIG = {
    APTS:{key:"APTS",label:"AP & TS"},
    KA:{key:"KA",label:"Karnataka"},
    TN:{key:"TN",label:"Tamil Nadu"},
    WB:{key:"WB",label:"West Bengal"},
    MH:{key:"MH",label:"Maharashtra"},
    DELHI:{key:"DELHI",label:"Delhi"}
};

const HR_POSITION_CONFIG = [
    {key:"telecallers",label:"Telecallers",target:["Telecallers Target"],achieved:["Telecallers Achieved"],remaining:["Telecallers Have To Achieve","Telecallers Have to Achieve"]},
    {key:"telesales",label:"Telesales",target:["Telesales Target"],achieved:["Telesales Achieved"],remaining:["Telesales Have To Achieve","Telesales Have to Achieve"]},
    {key:"surgeons",label:"Surgeons",target:["Surgeons Target"],achieved:["Surgeons Achieved"],remaining:["Surgeons Have To Achieve","Surgeons Have to Achieve"]},
    {key:"counsellors",label:"Counsellors",target:["Counsellors Target","Counselors Target"],achieved:["Counsellors Achieved","Counselors Achieved"],remaining:["Counsellors Have To Achieve","Counsellors Have to Achieve","Counselors Have To Achieve","Counselors Have to Achieve"]},
    {key:"fieldRepresentative",label:"Field Representative",target:["Field Representative Target"],achieved:["Field Representative Achieved"],remaining:["Field Representative Have To Achieve","Field Representative Have to Achieve"]},
    {key:"frontOffice",label:"Front Office",target:["Front Office Target"],achieved:["Front Office Achieved"],remaining:["Front Office Have To Achieve","Front Office Have to Achieve","Front OfficeHave To Achieve","Front OfficeHave to Achieve"]},
    {key:"pharmacist",label:"Pharmacist",target:["Pharmacist Target"],achieved:["Pharmacist Achieved"],remaining:["Pharmacist Have To Achieve","Pharmacist Have to Achieve"]},
    {key:"mbbsDoctor",label:"MBBS Doctor",target:["MBBS Doctor Target","Mbbs Doctor Target","Mbbbs Doctor Target"],achieved:["MBBS Doctor Achieved","Mbbs Doctor Achieved","Mbbbs Doctor Achieved"],remaining:["MBBS Doctor Have To Achieve","MBBS Doctor Have to Achieve","Mbbs Doctor Have To Achieve","Mbbs Doctor Have to Achieve","Mbbbs Doctor Have To Achieve","Mbbbs Doctor Have to Achieve"]},
    {key:"houseKeeping",label:"House Keeping",target:["House Keeping Target","Housekeeping Target"],achieved:["House Keeping Achieved","Housekeeping Achieved"],remaining:["House Keeping Have To Achieve","House Keeping Have to Achieve","Housekeeping Have To Achieve","Housekeeping Have to Achieve"]}
];

function normalizeHrParameterKey(value){
    return String(value||"").trim().toLowerCase().replace(/[_–—-]+/g," ").replace(/\s+/g," ");
}

function hrNumber(value){
    if(value===null||value===undefined)return null;
    const text=String(value).trim();
    if(!text||text==="-"||/^n\/?a$/i.test(text))return null;
    const parsed=Number(text.replace(/,/g,"").replace(/[^0-9.+-]/g,""));
    return Number.isFinite(parsed)?parsed:null;
}

function createHrParameterMap(parameters){
    const map={};
    Object.entries(parameters&&typeof parameters==="object"?parameters:{}).forEach(([label,value])=>{
        map[normalizeHrParameterKey(label)]=value;
    });
    return map;
}

function getHrParameterValue(parameterMap,aliases){
    for(const alias of aliases){
        const key=normalizeHrParameterKey(alias);
        if(Object.prototype.hasOwnProperty.call(parameterMap,key))return parameterMap[key];
    }
    return null;
}

function buildHrPositionMetrics(parameters){
    const parameterMap=createHrParameterMap(parameters);
    const positions={};

    HR_POSITION_CONFIG.forEach(position=>{
        const target=hrNumber(getHrParameterValue(parameterMap,position.target));
        const achieved=hrNumber(getHrParameterValue(parameterMap,position.achieved));
        const explicitRemaining=hrNumber(getHrParameterValue(parameterMap,position.remaining));
        let haveToAchieve=explicitRemaining;

        if(haveToAchieve===null&&target!==null&&achieved!==null){
            haveToAchieve=Math.max(0,Number(target)-Number(achieved));
        }

        positions[position.key]={
            label:position.label,
            target,
            achieved,
            haveToAchieve
        };
    });

    return positions;
}

async function getHrLiveInput(req,res){
    const requestedSource=String(req.hrSourceOverride||req.query.source||"APTS").trim().toUpperCase();
    const sourceConfig=HR_SOURCE_CONFIG[requestedSource];

    if(!sourceConfig){
        return res.status(400).json({
            success:false,
            message:"Unsupported HR source. Use APTS, KA, TN, WB, MH or DELHI."
        });
    }

    try{
        const result=await callAppsScriptGet(
            "liveInput",
            {key:sourceConfig.key,liveInputKey:sourceConfig.key}
        );

        const positions=buildHrPositionMetrics(result.parameters||{});

        return res.json({
            success:true,
            sourceKey:sourceConfig.key,
            sourceLabel:sourceConfig.label,
            sourceColumn:result.sourceColumn||sourceConfig.label,
            positions,
            positionCount:Object.keys(positions).length,
            inputSpreadsheetName:result.inputSpreadsheetName||"Input",
            inputSheetName:result.inputSheetName||"Sheet1",
            refreshedAt:result.refreshedAt||new Date().toISOString()
        });
    }catch(error){
        console.error(`${sourceConfig.key} HR live input error:`,error);
        return res.status(500).json({
            success:false,
            sourceKey:sourceConfig.key,
            message:error.message||`Unable to load ${sourceConfig.label} HR live input.`
        });
    }
}


function getAptsHrGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.hrApts,req,res);
}

function clearAptsHrGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.hrApts,req,res);
}

function saveAptsHrReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.hrApts,req,res);
}

function sendAptsHrReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.hrApts,req,res);
}


function getAptsHrEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.hrApts;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `AP & TS HR email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `AP & TS HR requires ${minimumRecipients} unique recipients. Configure HR_APTS_RECIPIENTS or HR_APTS_RECIPIENT + HR_APTS_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getKaHrGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.hrKa,req,res);
}

function clearKaHrGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.hrKa,req,res);
}

function saveKaHrReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.hrKa,req,res);
}

function sendKaHrReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.hrKa,req,res);
}


function getKaHrEmailConfiguration(
    req,
    res
) {
    const config = REPORT_CONFIGS.hrKa;
    const status = getRecipientConfigurationStatus(config);
    const minimumRecipients = Number(config.minimumRecipients || 2);

    return res.json({
        success: status.recipients.length >= minimumRecipients,
        reportKey: config.key,
        reportTitle: config.title,
        recipientCount: status.recipients.length,
        requiredRecipientCount: minimumRecipients,
        configured: status.configured,
        checkedEnvVariables: Object.keys(status.configured),
        message: status.recipients.length >= minimumRecipients
            ? `Karnataka HR email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
            : `Karnataka HR requires ${minimumRecipients} unique recipients. Configure HR_KA_RECIPIENTS or HR_KA_RECIPIENT + HR_KA_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`
    });
}



function getTnHrGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.hrTn,req,res);
}

function clearTnHrGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.hrTn,req,res);
}

function saveTnHrReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.hrTn,req,res);
}

function sendTnHrReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.hrTn,req,res);
}


function getTnHrEmailConfiguration(
    req,
    res
) {

    const config =
        REPORT_CONFIGS.hrTn;


    const status =
        getRecipientConfigurationStatus(
            config
        );


    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );


    return res.json({

        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >=
            minimumRecipients
                ?
                `Tamil Nadu HR email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                :
                `Tamil Nadu HR requires ${minimumRecipients} unique recipients. Configure HR_TN_RECIPIENTS or HR_TN_RECIPIENT + HR_TN_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`

    });

}


function getWbHrGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.hrWb,req,res);
}

function clearWbHrGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.hrWb,req,res);
}

function saveWbHrReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.hrWb,req,res);
}

function sendWbHrReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.hrWb,req,res);
}


function getMhHrGeneratedReports(req,res){
    return getGeneratedReports(REPORT_CONFIGS.hrMh,req,res);
}

function clearMhHrGeneratedReports(req,res){
    return clearGeneratedReports(REPORT_CONFIGS.hrMh,req,res);
}

function saveMhHrReportToDrive(req,res){
    return saveReport(REPORT_CONFIGS.hrMh,req,res);
}

function sendMhHrReportEmail(req,res){
    return sendReportEmail(REPORT_CONFIGS.hrMh,req,res);
}


function getMhHrEmailConfiguration(
    req,
    res
) {

    const config = REPORT_CONFIGS.hrMh;

    const status =
        getRecipientConfigurationStatus(
            config
        );

    const minimumRecipients =
        Number(
            config.minimumRecipients
            ||
            2
        );

    return res.json({
        success:
            status.recipients.length >=
            minimumRecipients,

        reportKey:
            config.key,

        reportTitle:
            config.title,

        recipientCount:
            status.recipients.length,

        requiredRecipientCount:
            minimumRecipients,

        configured:
            status.configured,

        checkedEnvVariables:
            Object.keys(
                status.configured
            ),

        message:
            status.recipients.length >= minimumRecipients
                ? `Maharashtra HR email configuration is valid. ${status.recipients.length} unique recipients are loaded.`
                : `Maharashtra HR requires ${minimumRecipients} unique recipients. Configure HR_MH_RECIPIENTS or HR_MH_RECIPIENT + HR_MH_RECIPIENT_2 in back-end/.env, save the file, and restart Node.`
    });

}

/* ============================================================
   EXPORT
============================================================ */


function getAptsHrLiveInput(req,res){
    req.hrSourceOverride="APTS";
    return getHrLiveInput(req,res);
}


function getKaHrLiveInput(req,res){
    req.hrSourceOverride="KA";
    return getHrLiveInput(req,res);
}


function getTnHrLiveInput(req,res){
    req.hrSourceOverride="TN";
    return getHrLiveInput(req,res);
}


function getWbHrLiveInput(req,res){
    req.hrSourceOverride="WB";
    return getHrLiveInput(req,res);
}


function getMhHrLiveInput(req,res){
    req.hrSourceOverride="MH";
    return getHrLiveInput(req,res);
}

module.exports = {
    getPanIndiaLiveInput,
    getPanIndiaGeneratedReports,
    clearPanIndiaGeneratedReports,
    savePanIndiaReportToDrive,
    sendPanIndiaReportEmail,
    getPanIndiaEmailConfiguration,

    getLeadsLiveInput,
    getLeadsGeneratedReports,
    clearLeadsGeneratedReports,
    saveLeadsReportToDrive,
    sendLeadsReportEmail,

    getKarnatakaLeadsLiveInput,
    getKarnatakaLeadsGeneratedReports,
    clearKarnatakaLeadsGeneratedReports,
    saveKarnatakaLeadsReportToDrive,
    sendKarnatakaLeadsReportEmail,

    getTamilNaduLeadsLiveInput,
    getTamilNaduLeadsGeneratedReports,
    clearTamilNaduLeadsGeneratedReports,
    saveTamilNaduLeadsReportToDrive,
    sendTamilNaduLeadsReportEmail,

    getWestBengalLeadsLiveInput,
    getWestBengalLeadsGeneratedReports,
    clearWestBengalLeadsGeneratedReports,
    saveWestBengalLeadsReportToDrive,
    sendWestBengalLeadsReportEmail,

    getMaharashtraLeadsLiveInput,
    getMaharashtraLeadsGeneratedReports,
    clearMaharashtraLeadsGeneratedReports,
    saveMaharashtraLeadsReportToDrive,
    sendMaharashtraLeadsReportEmail,

    testAppsScriptWebApp,
    getAptsLiveInput,
    getAptsGeneratedReports,
    clearAptsGeneratedReports,
    saveAptsIpReportToDrive,
    sendAptsIpReportEmail,
    getAptsIpEmailConfiguration,

    testAptsOpAppsScriptWebApp,
    getAptsOpLiveInput,
    getAptsOpGeneratedReports,
    clearAptsOpGeneratedReports,
    saveAptsOpReportToDrive,
    sendAptsOpReportEmail,
    getAptsOpEmailConfiguration,

    testKaOpAppsScriptWebApp,
    getKaOpLiveInput,
    getKaOpGeneratedReports,
    clearKaOpGeneratedReports,
    saveKaOpReportToDrive,
    sendKaOpReportEmail,
    getKaOpEmailConfiguration,

    testTnOpAppsScriptWebApp,
    getTnOpLiveInput,
    getTnOpGeneratedReports,
    clearTnOpGeneratedReports,
    saveTnOpReportToDrive,
    sendTnOpReportEmail,
    getTnOpEmailConfiguration,

    testWbOpAppsScriptWebApp,
    getWbOpLiveInput,
    getWbOpGeneratedReports,
    clearWbOpGeneratedReports,
    saveWbOpReportToDrive,
    sendWbOpReportEmail,
    getWbOpEmailConfiguration,

    testMhOpAppsScriptWebApp,
    getMhOpLiveInput,
    getMhOpGeneratedReports,
    clearMhOpGeneratedReports,
    saveMhOpReportToDrive,
    sendMhOpReportEmail,
    getMhOpEmailConfiguration,

    testKaAppsScriptWebApp,
    getKaLiveInput,
    getKaGeneratedReports,
    clearKaGeneratedReports,
    saveKaIpReportToDrive,
    sendKaIpReportEmail,

    testTnAppsScriptWebApp,
    getTnLiveInput,
    getTnGeneratedReports,
    clearTnGeneratedReports,
    saveTnIpReportToDrive,
    sendTnIpReportEmail,

    testWbAppsScriptWebApp,
    getWbLiveInput,
    getWbGeneratedReports,
    clearWbGeneratedReports,
    saveWbIpReportToDrive,
    sendWbIpReportEmail,
    getWbIpEmailConfiguration,

    testMhAppsScriptWebApp,
    getMhLiveInput,
    getMhGeneratedReports,
    clearMhGeneratedReports,
    saveMhIpReportToDrive,
    sendMhIpReportEmail,
    getMhIpEmailConfiguration,

    getAptsLabLiveInput,
    getAptsLabGeneratedReports,
    clearAptsLabGeneratedReports,
    saveAptsLabReportToDrive,
    sendAptsLabReportEmail,
    getAptsLabEmailConfiguration,

    getKaLabLiveInput,
    getKaLabGeneratedReports,
    clearKaLabGeneratedReports,
    saveKaLabReportToDrive,
    sendKaLabReportEmail,
    getKaLabEmailConfiguration,

    getTnLabLiveInput,
    getTnLabGeneratedReports,
    clearTnLabGeneratedReports,
    saveTnLabReportToDrive,
    sendTnLabReportEmail,
    getTnLabEmailConfiguration,

    getWbLabLiveInput,
    getWbLabGeneratedReports,
    clearWbLabGeneratedReports,
    saveWbLabReportToDrive,
    sendWbLabReportEmail,
    getWbLabEmailConfiguration,

    getMhLabLiveInput,
    getMhLabGeneratedReports,
    clearMhLabGeneratedReports,
    saveMhLabReportToDrive,
    sendMhLabReportEmail,
    getMhLabEmailConfiguration,

    getAptsPharmacyLiveInput,
    getAptsPharmacyGeneratedReports,
    clearAptsPharmacyGeneratedReports,
    saveAptsPharmacyReportToDrive,
    sendAptsPharmacyReportEmail,
    getAptsPharmacyEmailConfiguration,

    getKaPharmacyLiveInput,
    getKaPharmacyGeneratedReports,
    clearKaPharmacyGeneratedReports,
    saveKaPharmacyReportToDrive,
    sendKaPharmacyReportEmail,
    getKaPharmacyEmailConfiguration,

    getTnPharmacyLiveInput,
    getTnPharmacyGeneratedReports,
    clearTnPharmacyGeneratedReports,
    saveTnPharmacyReportToDrive,
    sendTnPharmacyReportEmail,
    getTnPharmacyEmailConfiguration,

    getWbPharmacyLiveInput,
    getWbPharmacyGeneratedReports,
    clearWbPharmacyGeneratedReports,
    saveWbPharmacyReportToDrive,
    sendWbPharmacyReportEmail,
    getWbPharmacyEmailConfiguration,

    getMhPharmacyLiveInput,
    getMhPharmacyGeneratedReports,
    clearMhPharmacyGeneratedReports,
    saveMhPharmacyReportToDrive,
    sendMhPharmacyReportEmail,
    getMhPharmacyEmailConfiguration,

    getHrLiveInput,
    getAptsHrLiveInput,
    getAptsHrGeneratedReports,
    clearAptsHrGeneratedReports,
    saveAptsHrReportToDrive,
    sendAptsHrReportEmail,
    getAptsHrEmailConfiguration,

    getKaHrLiveInput,
    getKaHrGeneratedReports,
    clearKaHrGeneratedReports,
    saveKaHrReportToDrive,
    sendKaHrReportEmail,
    getKaHrEmailConfiguration,

    getTnHrLiveInput,
    getTnHrGeneratedReports,
    clearTnHrGeneratedReports,
    saveTnHrReportToDrive,
    sendTnHrReportEmail,
    getTnHrEmailConfiguration,

    getWbHrLiveInput,
    getWbHrGeneratedReports,
    clearWbHrGeneratedReports,
    saveWbHrReportToDrive,
    sendWbHrReportEmail,

    getMhHrLiveInput,
    getMhHrGeneratedReports,
    clearMhHrGeneratedReports,
    saveMhHrReportToDrive,
    sendMhHrReportEmail,
    getMhHrEmailConfiguration
};


