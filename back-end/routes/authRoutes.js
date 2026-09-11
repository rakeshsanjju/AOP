const express =
    require("express");


const router =
    express.Router();


const {

    loginSuperAdmin,

    logoutSuperAdmin,

    checkLoginStatus

} = require(
    "../controllers/authController"
);


// =====================================================
// LOGIN
// =====================================================

router.post(
    "/login",
    loginSuperAdmin
);


// =====================================================
// LOGIN STATUS
// =====================================================

router.get(
    "/auth/status",
    checkLoginStatus
);


// =====================================================
// LOGOUT
// =====================================================

router.post(
    "/logout",
    logoutSuperAdmin
);



module.exports =
    router;