const express = require("express");
const router = express.Router();

const authMiddleware =
    require("../middleware/authMiddleware");

const {
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
} =
    require("../controllers/reportMailController");


/* ============================================================
   PAN INDIA LIVE DASHBOARD
============================================================ */

router.get(
    "/pan-india/live-input",
    authMiddleware,
    getPanIndiaLiveInput
);

router.get(
    "/pan-india/generated-reports",
    authMiddleware,
    getPanIndiaGeneratedReports
);

router.delete(
    "/pan-india/generated-reports",
    authMiddleware,
    clearPanIndiaGeneratedReports
);

router.post(
    "/pan-india/save-drive",
    authMiddleware,
    savePanIndiaReportToDrive
);

router.post(
    "/pan-india/send-email",
    authMiddleware,
    sendPanIndiaReportEmail
);

router.get(
    "/pan-india/email-config",
    authMiddleware,
    getPanIndiaEmailConfiguration
);


/* ============================================================
   LEADS DASHBOARD
============================================================ */

/*
 * APTS-only Leads live source.
 * The dedicated path prevents accidental Pan India data reuse.
 */
router.get(
    "/leads/apts/live-input",
    authMiddleware,
    getLeadsLiveInput
);


/*
 * Backward-compatible alias.
 * This also returns APTS-only data now.
 */
router.get(
    "/leads/live-input",
    authMiddleware,
    getLeadsLiveInput
);

router.get(
    "/leads/generated-reports",
    authMiddleware,
    getLeadsGeneratedReports
);

router.delete(
    "/leads/generated-reports",
    authMiddleware,
    clearLeadsGeneratedReports
);

router.post(
    "/leads/save-drive",
    authMiddleware,
    saveLeadsReportToDrive
);

router.post(
    "/leads/send-email",
    authMiddleware,
    sendLeadsReportEmail
);


/* ============================================================
   KARNATAKA GOOGLE & FB / LEADS DASHBOARD
============================================================ */

router.get(
    "/leads/ka/live-input",
    authMiddleware,
    getKarnatakaLeadsLiveInput
);

router.get(
    "/leads/ka/generated-reports",
    authMiddleware,
    getKarnatakaLeadsGeneratedReports
);

router.delete(
    "/leads/ka/generated-reports",
    authMiddleware,
    clearKarnatakaLeadsGeneratedReports
);

router.post(
    "/leads/ka/save-drive",
    authMiddleware,
    saveKarnatakaLeadsReportToDrive
);

router.post(
    "/leads/ka/send-email",
    authMiddleware,
    sendKarnatakaLeadsReportEmail
);


/* ============================================================
   TAMIL NADU GOOGLE & FB / LEADS DASHBOARD
============================================================ */

router.get(
    "/leads/tn/live-input",
    authMiddleware,
    getTamilNaduLeadsLiveInput
);

router.get(
    "/leads/tn/generated-reports",
    authMiddleware,
    getTamilNaduLeadsGeneratedReports
);

router.delete(
    "/leads/tn/generated-reports",
    authMiddleware,
    clearTamilNaduLeadsGeneratedReports
);

router.post(
    "/leads/tn/save-drive",
    authMiddleware,
    saveTamilNaduLeadsReportToDrive
);

router.post(
    "/leads/tn/send-email",
    authMiddleware,
    sendTamilNaduLeadsReportEmail
);


/* ============================================================
   WEST BENGAL GOOGLE & FB / LEADS DASHBOARD
============================================================ */

router.get(
    "/leads/wb/live-input",
    authMiddleware,
    getWestBengalLeadsLiveInput
);

router.get(
    "/leads/wb/generated-reports",
    authMiddleware,
    getWestBengalLeadsGeneratedReports
);

router.delete(
    "/leads/wb/generated-reports",
    authMiddleware,
    clearWestBengalLeadsGeneratedReports
);

router.post(
    "/leads/wb/save-drive",
    authMiddleware,
    saveWestBengalLeadsReportToDrive
);

router.post(
    "/leads/wb/send-email",
    authMiddleware,
    sendWestBengalLeadsReportEmail
);


/* ============================================================
   MAHARASHTRA GOOGLE & FB / LEADS DASHBOARD
============================================================ */

router.get(
    "/leads/mh/live-input",
    authMiddleware,
    getMaharashtraLeadsLiveInput
);

router.get(
    "/leads/mh/generated-reports",
    authMiddleware,
    getMaharashtraLeadsGeneratedReports
);

router.delete(
    "/leads/mh/generated-reports",
    authMiddleware,
    clearMaharashtraLeadsGeneratedReports
);

router.post(
    "/leads/mh/save-drive",
    authMiddleware,
    saveMaharashtraLeadsReportToDrive
);

router.post(
    "/leads/mh/send-email",
    authMiddleware,
    sendMaharashtraLeadsReportEmail
);



/* ============================================================
   AP & TS
============================================================ */

router.get(
    "/apts/drive-test",
    authMiddleware,
    testAppsScriptWebApp
);

router.get(
    "/apts/live-input",
    authMiddleware,
    getAptsLiveInput
);


router.get(
    "/apts/generated-reports",
    authMiddleware,
    getAptsGeneratedReports
);

router.delete(
    "/apts/generated-reports",
    authMiddleware,
    clearAptsGeneratedReports
);

router.post(
    "/apts/save-drive",
    authMiddleware,
    saveAptsIpReportToDrive
);

router.post(
    "/apts/send-email",
    authMiddleware,
    sendAptsIpReportEmail
);
router.get(
    "/apts/email-config",
    authMiddleware,
    getAptsIpEmailConfiguration
);


/* ============================================================
   AP & TS OP REPORT
============================================================ */
router.get("/op-apts/drive-test",authMiddleware,testAptsOpAppsScriptWebApp);
router.get("/op-apts/live-input",authMiddleware,getAptsOpLiveInput);
router.get("/op-apts/generated-reports",authMiddleware,getAptsOpGeneratedReports);
router.delete("/op-apts/generated-reports",authMiddleware,clearAptsOpGeneratedReports);
router.post("/op-apts/save-drive",authMiddleware,saveAptsOpReportToDrive);
router.post("/op-apts/send-email",authMiddleware,sendAptsOpReportEmail);
router.get("/op-apts/email-config",authMiddleware,getAptsOpEmailConfiguration);

/* ============================================================
   KARNATAKA OP REPORT
============================================================ */

router.get(
    "/op-ka/drive-test",
    authMiddleware,
    testKaOpAppsScriptWebApp
);

router.get(
    "/op-ka/live-input",
    authMiddleware,
    getKaOpLiveInput
);

router.get(
    "/op-ka/generated-reports",
    authMiddleware,
    getKaOpGeneratedReports
);

router.delete(
    "/op-ka/generated-reports",
    authMiddleware,
    clearKaOpGeneratedReports
);

router.post(
    "/op-ka/save-drive",
    authMiddleware,
    saveKaOpReportToDrive
);

router.post(
    "/op-ka/send-email",
    authMiddleware,
    sendKaOpReportEmail
);

router.get(
    "/op-ka/email-config",
    authMiddleware,
    getKaOpEmailConfiguration
);


/* ============================================================
   KARNATAKA / KA
============================================================ */

router.get(
    "/ka/drive-test",
    authMiddleware,
    testKaAppsScriptWebApp
);

router.get(
    "/ka/live-input",
    authMiddleware,
    getKaLiveInput
);


router.get(
    "/ka/generated-reports",
    authMiddleware,
    getKaGeneratedReports
);

router.delete(
    "/ka/generated-reports",
    authMiddleware,
    clearKaGeneratedReports
);

router.post(
    "/ka/save-drive",
    authMiddleware,
    saveKaIpReportToDrive
);

router.post(
    "/ka/send-email",
    authMiddleware,
    sendKaIpReportEmail
);


/* ============================================================
   TAMIL NADU OP REPORT
============================================================ */

router.get(
    "/op-tn/drive-test",
    authMiddleware,
    testTnOpAppsScriptWebApp
);

router.get(
    "/op-tn/live-input",
    authMiddleware,
    getTnOpLiveInput
);

router.get(
    "/op-tn/generated-reports",
    authMiddleware,
    getTnOpGeneratedReports
);

router.delete(
    "/op-tn/generated-reports",
    authMiddleware,
    clearTnOpGeneratedReports
);

router.post(
    "/op-tn/save-drive",
    authMiddleware,
    saveTnOpReportToDrive
);

router.post(
    "/op-tn/send-email",
    authMiddleware,
    sendTnOpReportEmail
);

router.get(
    "/op-tn/email-config",
    authMiddleware,
    getTnOpEmailConfiguration
);


/* ============================================================
   TAMIL NADU / TN
============================================================ */

router.get(
    "/tn/drive-test",
    authMiddleware,
    testTnAppsScriptWebApp
);

router.get(
    "/tn/live-input",
    authMiddleware,
    getTnLiveInput
);


router.get(
    "/tn/generated-reports",
    authMiddleware,
    getTnGeneratedReports
);

router.delete(
    "/tn/generated-reports",
    authMiddleware,
    clearTnGeneratedReports
);

router.post(
    "/tn/save-drive",
    authMiddleware,
    saveTnIpReportToDrive
);

router.post(
    "/tn/send-email",
    authMiddleware,
    sendTnIpReportEmail
);


/* ============================================================
   WEST BENGAL OP REPORT
============================================================ */

router.get(
    "/op-wb/drive-test",
    authMiddleware,
    testWbOpAppsScriptWebApp
);

router.get(
    "/op-wb/live-input",
    authMiddleware,
    getWbOpLiveInput
);

router.get(
    "/op-wb/generated-reports",
    authMiddleware,
    getWbOpGeneratedReports
);

router.delete(
    "/op-wb/generated-reports",
    authMiddleware,
    clearWbOpGeneratedReports
);

router.post(
    "/op-wb/save-drive",
    authMiddleware,
    saveWbOpReportToDrive
);

router.post(
    "/op-wb/send-email",
    authMiddleware,
    sendWbOpReportEmail
);

router.get(
    "/op-wb/email-config",
    authMiddleware,
    getWbOpEmailConfiguration
);


/* ============================================================
   WEST BENGAL / WB
============================================================ */

router.get(
    "/wb/drive-test",
    authMiddleware,
    testWbAppsScriptWebApp
);

router.get(
    "/wb/live-input",
    authMiddleware,
    getWbLiveInput
);


router.get(
    "/wb/generated-reports",
    authMiddleware,
    getWbGeneratedReports
);

router.delete(
    "/wb/generated-reports",
    authMiddleware,
    clearWbGeneratedReports
);

router.post(
    "/wb/save-drive",
    authMiddleware,
    saveWbIpReportToDrive
);

router.post(
    "/wb/send-email",
    authMiddleware,
    sendWbIpReportEmail
);

router.get(
    "/wb/email-config",
    authMiddleware,
    getWbIpEmailConfiguration
);


/* ============================================================
   MAHARASHTRA OP REPORT
============================================================ */

router.get("/op-mh/drive-test", authMiddleware, testMhOpAppsScriptWebApp);
router.get("/op-mh/live-input", authMiddleware, getMhOpLiveInput);
router.get("/op-mh/generated-reports", authMiddleware, getMhOpGeneratedReports);
router.delete("/op-mh/generated-reports", authMiddleware, clearMhOpGeneratedReports);
router.post("/op-mh/save-drive", authMiddleware, saveMhOpReportToDrive);
router.post("/op-mh/send-email", authMiddleware, sendMhOpReportEmail);

router.get(
    "/op-mh/email-config",
    authMiddleware,
    getMhOpEmailConfiguration
);


/* ============================================================
   MAHARASHTRA / MH
============================================================ */

router.get(
    "/mh/drive-test",
    authMiddleware,
    testMhAppsScriptWebApp
);

router.get(
    "/mh/live-input",
    authMiddleware,
    getMhLiveInput
);


router.get(
    "/mh/generated-reports",
    authMiddleware,
    getMhGeneratedReports
);

router.delete(
    "/mh/generated-reports",
    authMiddleware,
    clearMhGeneratedReports
);

router.post(
    "/mh/save-drive",
    authMiddleware,
    saveMhIpReportToDrive
);

router.post(
    "/mh/send-email",
    authMiddleware,
    sendMhIpReportEmail
);

router.get(
    "/mh/email-config",
    authMiddleware,
    getMhIpEmailConfiguration
);


/* ============================================================
   LAB REPORTS - 5 STATES
============================================================ */

router.get("/lab/apts/live-input",authMiddleware,getAptsLabLiveInput);
router.get("/lab/apts/generated-reports",authMiddleware,getAptsLabGeneratedReports);
router.delete("/lab/apts/generated-reports",authMiddleware,clearAptsLabGeneratedReports);
router.post("/lab/apts/save-drive",authMiddleware,saveAptsLabReportToDrive);
router.post("/lab/apts/send-email",authMiddleware,sendAptsLabReportEmail);
router.get("/lab/apts/email-config",authMiddleware,getAptsLabEmailConfiguration);

router.get("/lab/ka/live-input",authMiddleware,getKaLabLiveInput);
router.get("/lab/ka/generated-reports",authMiddleware,getKaLabGeneratedReports);
router.delete("/lab/ka/generated-reports",authMiddleware,clearKaLabGeneratedReports);
router.post("/lab/ka/save-drive",authMiddleware,saveKaLabReportToDrive);
router.post("/lab/ka/send-email",authMiddleware,sendKaLabReportEmail);
router.get("/lab/ka/email-config",authMiddleware,getKaLabEmailConfiguration);

router.get("/lab/tn/live-input",authMiddleware,getTnLabLiveInput);
router.get("/lab/tn/generated-reports",authMiddleware,getTnLabGeneratedReports);
router.delete("/lab/tn/generated-reports",authMiddleware,clearTnLabGeneratedReports);
router.post("/lab/tn/save-drive",authMiddleware,saveTnLabReportToDrive);
router.post("/lab/tn/send-email",authMiddleware,sendTnLabReportEmail);
router.get("/lab/tn/email-config",authMiddleware,getTnLabEmailConfiguration);

router.get("/lab/wb/live-input",authMiddleware,getWbLabLiveInput);
router.get("/lab/wb/generated-reports",authMiddleware,getWbLabGeneratedReports);
router.delete("/lab/wb/generated-reports",authMiddleware,clearWbLabGeneratedReports);
router.post("/lab/wb/save-drive",authMiddleware,saveWbLabReportToDrive);
router.post("/lab/wb/send-email",authMiddleware,sendWbLabReportEmail);
router.get("/lab/wb/email-config",authMiddleware,getWbLabEmailConfiguration);

router.get("/lab/mh/live-input",authMiddleware,getMhLabLiveInput);
router.get("/lab/mh/generated-reports",authMiddleware,getMhLabGeneratedReports);
router.delete("/lab/mh/generated-reports",authMiddleware,clearMhLabGeneratedReports);
router.post("/lab/mh/save-drive",authMiddleware,saveMhLabReportToDrive);
router.post("/lab/mh/send-email",authMiddleware,sendMhLabReportEmail);
router.get("/lab/mh/email-config",authMiddleware,getMhLabEmailConfiguration);


/* ============================================================
   AP & TS PHARMACY REPORT
============================================================ */

router.get(
    "/pharmacy/apts/live-input",
    authMiddleware,
    getAptsPharmacyLiveInput
);

router.get(
    "/pharmacy/apts/generated-reports",
    authMiddleware,
    getAptsPharmacyGeneratedReports
);

router.delete(
    "/pharmacy/apts/generated-reports",
    authMiddleware,
    clearAptsPharmacyGeneratedReports
);

router.post(
    "/pharmacy/apts/save-drive",
    authMiddleware,
    saveAptsPharmacyReportToDrive
);

router.post(
    "/pharmacy/apts/send-email",
    authMiddleware,
    sendAptsPharmacyReportEmail
);

router.get(
    "/pharmacy/apts/email-config",
    authMiddleware,
    getAptsPharmacyEmailConfiguration
);


/* ============================================================
   KARNATAKA PHARMACY REPORT
============================================================ */

router.get("/pharmacy/ka/live-input",authMiddleware,getKaPharmacyLiveInput);
router.get("/pharmacy/ka/generated-reports",authMiddleware,getKaPharmacyGeneratedReports);
router.delete("/pharmacy/ka/generated-reports",authMiddleware,clearKaPharmacyGeneratedReports);
router.post("/pharmacy/ka/save-drive",authMiddleware,saveKaPharmacyReportToDrive);
router.post("/pharmacy/ka/send-email",authMiddleware,sendKaPharmacyReportEmail);

router.get(
    "/pharmacy/ka/email-config",
    authMiddleware,
    getKaPharmacyEmailConfiguration
);


/* ============================================================
   TAMIL NADU PHARMACY REPORT
============================================================ */

router.get("/pharmacy/tn/live-input",authMiddleware,getTnPharmacyLiveInput);
router.get("/pharmacy/tn/generated-reports",authMiddleware,getTnPharmacyGeneratedReports);
router.delete("/pharmacy/tn/generated-reports",authMiddleware,clearTnPharmacyGeneratedReports);
router.post("/pharmacy/tn/save-drive",authMiddleware,saveTnPharmacyReportToDrive);
router.post("/pharmacy/tn/send-email",authMiddleware,sendTnPharmacyReportEmail);

router.get(
    "/pharmacy/tn/email-config",
    authMiddleware,
    getTnPharmacyEmailConfiguration
);


/* ============================================================
   WEST BENGAL PHARMACY REPORT
============================================================ */

router.get("/pharmacy/wb/live-input",authMiddleware,getWbPharmacyLiveInput);
router.get("/pharmacy/wb/generated-reports",authMiddleware,getWbPharmacyGeneratedReports);
router.delete("/pharmacy/wb/generated-reports",authMiddleware,clearWbPharmacyGeneratedReports);
router.post("/pharmacy/wb/save-drive",authMiddleware,saveWbPharmacyReportToDrive);
router.post("/pharmacy/wb/send-email",authMiddleware,sendWbPharmacyReportEmail);

router.get(
    "/pharmacy/wb/email-config",
    authMiddleware,
    getWbPharmacyEmailConfiguration
);


/* ============================================================
   MAHARASHTRA PHARMACY REPORT
============================================================ */

router.get("/pharmacy/mh/live-input",authMiddleware,getMhPharmacyLiveInput);
router.get("/pharmacy/mh/generated-reports",authMiddleware,getMhPharmacyGeneratedReports);
router.delete("/pharmacy/mh/generated-reports",authMiddleware,clearMhPharmacyGeneratedReports);
router.post("/pharmacy/mh/save-drive",authMiddleware,saveMhPharmacyReportToDrive);
router.post("/pharmacy/mh/send-email",authMiddleware,sendMhPharmacyReportEmail);

router.get(
    "/pharmacy/mh/email-config",
    authMiddleware,
    getMhPharmacyEmailConfiguration
);


/* ============================================================
   HR PORTAL - AP & TS ONLY
============================================================ */

/*
 * AP & TS only.
 *
 * getHrLiveInput already defaults to APTS when no source query
 * parameter is supplied, so no extra wrapper handler is needed.
 */
router.get(
    "/hr/apts/live-input",
    authMiddleware,
    getHrLiveInput
);

router.get(
    "/hr/apts/generated-reports",
    authMiddleware,
    getAptsHrGeneratedReports
);

router.delete(
    "/hr/apts/generated-reports",
    authMiddleware,
    clearAptsHrGeneratedReports
);

router.post(
    "/hr/apts/save-drive",
    authMiddleware,
    saveAptsHrReportToDrive
);

router.post(
    "/hr/apts/send-email",
    authMiddleware,
    sendAptsHrReportEmail
);

router.get(
    "/hr/apts/email-config",
    authMiddleware,
    getAptsHrEmailConfiguration
);


/* ============================================================
   HR PORTAL - KARNATAKA
============================================================ */

router.get(
    "/hr/ka/live-input",
    authMiddleware,
    getKaHrLiveInput
);

router.get(
    "/hr/ka/generated-reports",
    authMiddleware,
    getKaHrGeneratedReports
);

router.delete(
    "/hr/ka/generated-reports",
    authMiddleware,
    clearKaHrGeneratedReports
);

router.post(
    "/hr/ka/save-drive",
    authMiddleware,
    saveKaHrReportToDrive
);

router.post(
    "/hr/ka/send-email",
    authMiddleware,
    sendKaHrReportEmail
);

router.get(
    "/hr/ka/email-config",
    authMiddleware,
    getKaHrEmailConfiguration
);


/* ============================================================
   HR PORTAL - TAMIL NADU
============================================================ */

router.get(
    "/hr/tn/live-input",
    authMiddleware,
    getTnHrLiveInput
);

router.get(
    "/hr/tn/generated-reports",
    authMiddleware,
    getTnHrGeneratedReports
);

router.delete(
    "/hr/tn/generated-reports",
    authMiddleware,
    clearTnHrGeneratedReports
);

router.post(
    "/hr/tn/save-drive",
    authMiddleware,
    saveTnHrReportToDrive
);

router.post(
    "/hr/tn/send-email",
    authMiddleware,
    sendTnHrReportEmail
);

router.get(
    "/hr/tn/email-config",
    authMiddleware,
    getTnHrEmailConfiguration
);


/* ============================================================
   HR PORTAL - WEST BENGAL
============================================================ */

router.get(
    "/hr/wb/live-input",
    authMiddleware,
    getWbHrLiveInput
);

router.get(
    "/hr/wb/generated-reports",
    authMiddleware,
    getWbHrGeneratedReports
);

router.delete(
    "/hr/wb/generated-reports",
    authMiddleware,
    clearWbHrGeneratedReports
);

router.post(
    "/hr/wb/save-drive",
    authMiddleware,
    saveWbHrReportToDrive
);

router.post(
    "/hr/wb/send-email",
    authMiddleware,
    sendWbHrReportEmail
);


/* ============================================================
   HR PORTAL - MAHARASHTRA
============================================================ */

router.get(
    "/hr/mh/live-input",
    authMiddleware,
    getMhHrLiveInput
);

router.get(
    "/hr/mh/generated-reports",
    authMiddleware,
    getMhHrGeneratedReports
);

router.delete(
    "/hr/mh/generated-reports",
    authMiddleware,
    clearMhHrGeneratedReports
);

router.post(
    "/hr/mh/save-drive",
    authMiddleware,
    saveMhHrReportToDrive
);

router.post(
    "/hr/mh/send-email",
    authMiddleware,
    sendMhHrReportEmail
);

router.get(
    "/hr/mh/email-config",
    authMiddleware,
    getMhHrEmailConfiguration
);


/* ============================================================
   ROUTE TEST
============================================================ */

router.get(
    "/test",
    (
        req,
        res
    ) => {

        res.json({
            success: true,
            message:
                "APTS IP, AP & TS OP, Karnataka, Tamil Nadu, West Bengal and Maharashtra report routes are working."
        });

    }
);


module.exports = router;
