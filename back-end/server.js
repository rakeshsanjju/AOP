require("dotenv").config();

const express = require("express");
const cors = require("cors");
const session = require("express-session");

const authRoutes = require("./routes/authRoutes");
const superAdminRoutes = require("./routes/superAdminRoutes");
const reportMailRoutes = require("./routes/reportMailRoutes");

const app = express();
const isProduction = process.env.NODE_ENV === "production";

if (isProduction && !process.env.SESSION_SECRET) {
    throw new Error("SESSION_SECRET must be set in production");
}

const allowedOrigins = new Set([
    "https://lively-solace-production-10fd.up.railway.app",
    "http://localhost:5500",
    "http://127.0.0.1:5500",
    "http://localhost:3000",
    "http://127.0.0.1:3000"
]);

const corsOptions = {
    origin(origin, callback) {
        if (!origin || allowedOrigins.has(origin)) {
            return callback(null, true);
        }
        return callback(new Error(`CORS origin is not allowed: ${origin}`));
    },
    credentials: true,
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"]
};

app.use(cors(corsOptions));

// Confirm whether report requests reach this Express process.
app.use((req, res, next) => {
    if (req.path.startsWith("/api/reports/")) {
        const startedAt = Date.now();
        res.on("finish", () => {
            console.log(
                `${req.method} ${req.originalUrl} ${res.statusCode} ` +
                `${Date.now() - startedAt}ms origin=${req.headers.origin || "none"}`
            );
        });
    }
    next();
});
app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended: true, limit: "20mb" }));

// Railway terminates HTTPS at its proxy. Trust it so secure session cookies work.
app.set("trust proxy", 1);

app.use(session({
    name: "superAdmin.sid",
    secret: process.env.SESSION_SECRET || "local-development-only-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
        httpOnly: true,
        secure: isProduction,
        sameSite: isProduction ? "none" : "lax",
        maxAge: 24 * 60 * 60 * 1000
    }
}));

app.use("/api", authRoutes);
app.use("/api/super-admin", superAdminRoutes);
app.use("/api/reports", reportMailRoutes);

app.get("/", (req, res) => {
    res.json({ success: true, message: "AVIS backend is running." });
});

app.get("/api/test", (req, res) => {
    res.json({ success: true, message: "API working successfully." });
});

app.get("/api/cors-check", (req, res) => {
    res.json({
        success: true,
        origin: req.headers.origin || null,
        message: "This response came from the Express backend."
    });
});

app.use((req, res) => {
    res.status(404).json({
        success: false,
        message: `Route not found: ${req.method} ${req.originalUrl}`
    });
});

app.use((err, req, res, next) => {
    console.error("Global server error:", err);
    res.status(err.status || 500).json({
        success: false,
        message: err.message || "Internal server error"
    });
});

const PORT = Number(process.env.PORT) || 3000;
app.listen(PORT, "0.0.0.0", () => {
    console.log(`Backend listening on port ${PORT}`);
});
