const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
const port = process.env.PORT || 3000;

app.use(express.static('.'));
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    next();
});

const certPath = path.join(os.homedir(), '.office-addin-dev-certs');

let sslOptions;
try {
    sslOptions = {
        key: fs.readFileSync(path.join(certPath, 'localhost.key')),
        cert: fs.readFileSync(path.join(certPath, 'localhost.crt'))
    };
} catch (error) {
    console.error('❌ SSL certificates not found. Run: npx office-addin-dev-certs install --machine');
    process.exit(1);
}

https.createServer(sslOptions, app).listen(port, () => {
    console.log(`📧 Email Summarizer Dev Server`);
    console.log(`📋 Manifest: https://localhost:${port}/manifest.xml`);
    console.log(`🔧 Backend: Ensure Python server is running on port 5000`);
    console.log(`🌐 Interface: https://localhost:${port}/taskpane.html`);
});