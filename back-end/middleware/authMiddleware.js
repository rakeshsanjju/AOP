/* =========================================
   SUPER ADMIN AUTH MIDDLEWARE
========================================= */

const authMiddleware = (
    req,
    res,
    next
) => {

    try {

        console.log(
            "Session:",
            req.session
        );


        /*
         * User is logged in if admin
         * information exists in session.
         */

        if (
            req.session &&
            req.session.admin
        ) {

            return next();

        }


        return res
            .status(401)
            .json({

                success:
                    false,

                message:
                    "Unauthorized. Please login again."

            });

    }
    catch (error) {

        console.error(
            "Auth middleware error:",
            error
        );


        return res
            .status(500)
            .json({

                success:
                    false,

                message:
                    "Authentication error."

            });

    }

};


/*
 * IMPORTANT:
 *
 * Export the function directly.
 *
 * Then this works:
 *
 * const authMiddleware =
 * require("../middleware/authMiddleware");
 */

module.exports =
    authMiddleware;


/*
 * Also expose named versions
 * so existing code will not break
 * if another route uses destructuring.
 */

module.exports.authMiddleware =
    authMiddleware;

module.exports.requireSuperAdmin =
    authMiddleware;