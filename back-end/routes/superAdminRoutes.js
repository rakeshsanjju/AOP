const express =
    require("express");


const router =
    express.Router();


const {

    requireSuperAdmin

} = require(
    "../middleware/authMiddleware"
);


const {

    getDashboardData

} = require(
    "../controllers/superAdminController"
);


// =====================================================
// DASHBOARD DATA
// =====================================================

router.get(

    "/dashboard",

    requireSuperAdmin,

    getDashboardData

);


module.exports =
    router;