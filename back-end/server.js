require("dotenv").config();

const express = require("express");
const cors = require("cors");
const session = require("express-session");

const authRoutes = require("./routes/authRoutes");
const superAdminRoutes = require(
    "./routes/superAdminRoutes"
);
const reportMailRoutes = require(
    "./routes/reportMailRoutes"
);

const app = express();

const allowedOrigins = [
    "http://127.0.0.1:5500",
    "http://localhost:5500"
];

app.use(
    cors({
        origin(origin, callback) {
            if (
                !origin ||
                allowedOrigins.includes(origin)
            ) {
                return callback(null, true);
            }

            return callback(
                new Error(
                    `CORS origin is not allowed: ${origin}`
                )
            );
        },
        credentials: true
    })
);

app.use(
    express.json({
        limit: "20mb"
    })
);

app.use(
    express.urlencoded({
        extended: true,
        limit: "20mb"
    })
);

app.use(
    session({
        name: "superAdmin.sid",
        secret:
            process.env.SESSION_SECRET ||
            "super-admin-secret",
        resave: false,
        saveUninitialized: false,
        cookie: {
            httpOnly: true,
            secure: false,
            sameSite: "lax",
            maxAge:
                24 *
                60 *
                60 *
                1000
        }
    })
);

app.use(
    "/api",
    authRoutes
);

app.use(
    "/api/super-admin",
    superAdminRoutes
);

app.use(
    "/api/reports",
    reportMailRoutes
);

app.get(
    "/",
    (req, res) => {
        res.json({
            success: true,
            message:
                "AVIS backend is running."
        });
    }
);

app.get(
    "/api/test",
    (req, res) => {
        res.json({
            success: true,
            message:
                "API working successfully."
        });
    }
);

app.use(
    (req, res) => {
        res.status(404).json({
            success: false,
            message:
                `Route not found: ${req.method} ${req.originalUrl}`
        });
    }
);

app.use(
    (err, req, res, next) => {
        console.error(
            "Global server error:",
            err
        );

        res
            .status(err.status || 500)
            .json({
                success: false,
                message:
                    err.message ||
                    "Internal server error"
            });
    }
);

const PORT =
    Number(process.env.PORT) ||
    3000;

app.listen(
    PORT,
    "127.0.0.1",
    () => {
        console.log(
            `Server running: http://127.0.0.1:${PORT}`
        );

        console.log(
            `Drive save API: http://127.0.0.1:${PORT}/api/reports/apts/save-drive`
        );

        console.log(
            `Email API: http://127.0.0.1:${PORT}/api/reports/apts/send-email`
        );
    }
);