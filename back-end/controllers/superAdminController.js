exports.getDashboardData = (
    req,
    res
) => {

    try {

        return res.status(200).json({

            success: true,

            message:
                "Super Admin Dashboard",

            admin:
                req.session.admin,

            dashboard: {

                totalUsers: 120,

                activeUsers: 98,

                enquiries: 250,

                appointments: 84

            }

        });


    } catch (error) {

        console.error(
            "Dashboard controller error:",
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