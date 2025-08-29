const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Serve static files
app.use(express.static('.'));

// CORS headers for Office Add-ins
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    next();
});

// Try to load SSL certificates
let sslOptions;
try {
    sslOptions = {
        key: fs.readFileSync(path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.key')),
        cert: fs.readFileSync(path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.crt'))
    };
} catch (error) {
    console.error('SSL certificates not found. Run: npx office-addin-dev-certs install --machine');
    process.exit(1);
}

const server = https.createServer(sslOptions, app);

server.listen(port, () => {
    console.log('📋 Manifest URL: https://localhost:3000/manifest.xml');
    console.log('🔧 Make sure Python backend is running on http://localhost:5000');
});