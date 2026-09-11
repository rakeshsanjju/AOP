const SUPER_ADMIN =
    require("../config/adminConfig");


// =====================================================
// SUPER ADMIN LOGIN
// =====================================================

exports.loginSuperAdmin = (req, res) => {

    try {

        console.log(
            "Login request received:",
            req.body
        );


        const {
            email,
            password
        } = req.body;


        // =============================================
        // VALIDATE INPUT
        // =============================================

        if (!email || !password) {

            return res.status(400).json({

                success: false,

                message:
                    "Email and password are required"

            });

        }


        // =============================================
        // NORMALIZE EMAIL
        // =============================================

        const enteredEmail =
            email
                .trim()
                .toLowerCase();


        const superAdminEmail =
            SUPER_ADMIN.email
                .trim()
                .toLowerCase();


        // =============================================
        // CHECK STATIC LOGIN
        // =============================================

        if (
            enteredEmail === superAdminEmail &&
            password === SUPER_ADMIN.password
        ) {

            // =========================================
            // CREATE SESSION
            // =========================================

            req.session.isSuperAdmin =
                true;


            req.session.admin = {

                name:
                    SUPER_ADMIN.name,

                email:
                    SUPER_ADMIN.email,

                role:
                    SUPER_ADMIN.role

            };


            // Explicitly save session

            req.session.save(
                (error) => {

                    if (error) {

                        console.error(
                            "Session save error:",
                            error
                        );


                        return res
                            .status(500)
                            .json({

                                success:
                                    false,

                                message:
                                    "Unable to create login session"

                            });

                    }


                    console.log(
                        "Super Admin login successful"
                    );


                    return res
                        .status(200)
                        .json({

                            success:
                                true,

                            message:
                                "Login successful",

                            redirect:
                                "super-Admin.html"

                        });

                }
            );


            return;

        }


        // =============================================
        // WRONG CREDENTIALS
        // =============================================

        console.log(
            "Invalid login attempt"
        );


        return res
            .status(401)
            .json({

                success: false,

                message:
                    "Invalid email or password"

            });


    } catch (error) {

        console.error(
            "Login controller error:",
            error
        );


        return res
            .status(500)
            .json({

                success: false,

                message:
                    "Internal server error"

            });

    }

};



// =====================================================
// CHECK LOGIN STATUS
// =====================================================

exports.checkLoginStatus = (
    req,
    res
) => {

    try {

        if (
            req.session &&
            req.session.isSuperAdmin &&
            req.session.admin
        ) {

            return res
                .status(200)
                .json({

                    success: true,

                    loggedIn: true,

                    admin:
                        req.session.admin

                });

        }


        return res
            .status(401)
            .json({

                success: false,

                loggedIn: false,

                message:
                    "Not authenticated"

            });


    } catch (error) {

        console.error(
            "Check login status error:",
            error
        );


        return res
            .status(500)
            .json({

                success: false,

                message:
                    "Internal server error"

            });

    }

};



// =====================================================
// LOGOUT
// =====================================================

exports.logoutSuperAdmin = (
    req,
    res
) => {

    try {

        req.session.destroy(
            (error) => {

                if (error) {

                    console.error(
                        "Logout error:",
                        error
                    );


                    return res
                        .status(500)
                        .json({

                            success:
                                false,

                            message:
                                "Unable to logout"

                        });

                }


                res.clearCookie(
                    "superAdminSession"
                );


                return res
                    .status(200)
                    .json({

                        success: true,

                        message:
                            "Logout successful"

                    });

            }
        );


    } catch (error) {

        console.error(
            "Logout error:",
            error
        );


        return res
            .status(500)
            .json({

                success: false,

                message:
                    "Internal server error"

            });

    }

};


