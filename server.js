const express = require('express');
const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
const HTTPS_PORT = 3000;
const HTTP_PORT = 3001;

console.log('🚀 Starting DUAL SERVER (HTTP + HTTPS) for maximum compatibility');

// 共通ミドルウェア設定
function setupCommonMiddleware(app) {
    // 非常に寛容なCORS設定
    app.use((req, res, next) => {
        res.header('Access-Control-Allow-Origin', '*');
        res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS, HEAD');
        res.header('Access-Control-Allow-Headers', '*');
        res.header('Access-Control-Allow-Credentials', 'true');
        res.header('X-Frame-Options', 'ALLOWALL');
        res.header('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.header('Pragma', 'no-cache');
        res.header('Expires', '0');

        if (req.method === 'OPTIONS') {
            res.sendStatus(200);
            return;
        }
        next();
    });

    // ログ出力
    app.use((req, res, next) => {
        console.log(`[${new Date().toISOString()}] ${req.protocol.toUpperCase()} ${req.method} ${req.url}`);
        next();
    });

    app.use(express.json({ limit: '50mb' }));
    app.use(express.urlencoded({ extended: true, limit: '50mb' }));

    // 静的ファイル配信
    app.use('/src', express.static(path.join(__dirname, 'src')));
    app.use('/assets', express.static(path.join(__dirname, 'assets')));
    app.use('/manifest', express.static(path.join(__dirname, 'manifest')));

    // ルート定義
    app.get('/', (req, res) => {
        res.sendFile(path.join(__dirname, 'src', 'taskpane.html'));
    });

    app.get('/src/taskpane.html', (req, res) => {
        console.log('=== Serving taskpane.html directly ===');
        res.sendFile(path.join(__dirname, 'src', 'taskpane.html'));
    });

    app.get('/callback', (req, res) => {
        res.sendFile(path.join(__dirname, 'src', 'cllback.html'));
    });

    app.get('/manifest.xml', (req, res) => {
        res.type('application/xml');
        res.sendFile(path.join(__dirname, 'manifest', 'manifest.xml'));
    });

    // API エンドポイント
    app.get('/api/config', (req, res) => {
        res.json({
            client_id: process.env.SALESFORCE_CLIENT_ID,
            redirect_uri: process.env.SALESFORCE_REDIRECT_URI || `${req.protocol}://localhost:${req.get('host').split(':')[1] || (req.secure ? HTTPS_PORT : HTTP_PORT)}/callback`
        });
    });

    app.post('/api/oauth/callback', async (req, res) => {
        try {
            const { code } = req.body;
            if (!code) {
                return res.status(400).json({ error: 'Authorization code is required' });
            }

            const params = new URLSearchParams();
            params.append('grant_type', 'authorization_code');
            params.append('client_id', process.env.SALESFORCE_CLIENT_ID);
            params.append('client_secret', process.env.SALESFORCE_CLIENT_SECRET);
            params.append('redirect_uri', process.env.SALESFORCE_REDIRECT_URI);
            params.append('code', code);

            const response = await fetch('https://login.salesforce.com/services/oauth2/token', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: params
            });

            if (response.ok) {
                const tokenData = await response.json();
                res.json({
                    success: true,
                    access_token: tokenData.access_token,
                    instance_url: tokenData.instance_url
                });
            } else {
                const error = await response.json();
                throw new Error(error.error_description || error.error);
            }
        } catch (error) {
            console.error('OAuth error:', error);
            res.status(500).json({ error: error.message });
        }
    });

    // ヘルス チェック
    app.get('/health', (req, res) => {
        res.json({
            status: 'OK',
            protocol: req.protocol,
            port: req.get('host'),
            timestamp: new Date().toISOString(),
            server_type: 'DUAL (HTTP + HTTPS)'
        });
    });

    // 404ハンドラー
    app.use((req, res) => {
        console.log(`404 - ${req.protocol.toUpperCase()}: ${req.url}`);
        res.status(404).json({
            error: 'Not Found',
            url: req.url,
            protocol: req.protocol,
            available_urls: [
                '/',
                '/src/taskpane.html',
                '/callback',
                '/api/config',
                '/health'
            ]
        });
    });
}

// 共通ミドルウェアを設定
setupCommonMiddleware(app);

// HTTP サーバー起動
const httpServer = http.createServer(app);
httpServer.listen(HTTP_PORT, () => {
    console.log(`📡 HTTP Server running at http://localhost:${HTTP_PORT}`);
});

// HTTPS サーバー起動（証明書がある場合のみ）
try {
    const certPath = path.join(require('os').homedir(), '.office-addin-dev-certs');
    const keyPath = path.join(certPath, 'localhost.key');
    const certFilePath = path.join(certPath, 'localhost.crt');

    if (fs.existsSync(keyPath) && fs.existsSync(certFilePath)) {
        const sslOptions = {
            key: fs.readFileSync(keyPath),
            cert: fs.readFileSync(certFilePath)
        };

        const httpsServer = https.createServer(sslOptions, app);

        httpsServer.on('clientError', (err, socket) => {
            console.log('HTTPS Client error (ignored):', err.message);
            if (!socket.destroyed) {
                socket.end('HTTP/1.1 400 Bad Request\r\n\r\n');
            }
        });

        httpsServer.listen(HTTPS_PORT, () => {
            console.log(`🔒 HTTPS Server running at https://localhost:${HTTPS_PORT}`);
            console.log('\n=== 🎯 READY FOR OUTLOOK ADD-IN ===');
            console.log('✅ Both HTTP and HTTPS servers are running');
            console.log('✅ Use HTTPS for production: https://localhost:3000');
            console.log('✅ Use HTTP for fallback: http://localhost:3001');
            console.log('=====================================\n');
        });
    } else {
        console.log('⚠️  HTTPS certificates not found - running HTTP only');
        console.log('📡 HTTP Server ready at http://localhost:3001');
    }
} catch (error) {
    console.error('HTTPS setup failed:', error.message);
    console.log('📡 Falling back to HTTP only at http://localhost:3001');
}