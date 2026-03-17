const express = require('express');
const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
const HTTPS_PORT = 3000;
const HTTP_PORT = 3001;

console.log('🚀 Starting HYBRID SERVER (HTTP + HTTPS) for Office Add-in compatibility');

// 共通ミドルウェア設定
function setupCommonMiddleware(app) {
    // CORS設定
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
        const protocol = req.secure ? 'https' : 'http';
        const port = req.secure ? HTTPS_PORT : HTTP_PORT;
        res.json({
            client_id: process.env.SALESFORCE_CLIENT_ID,
            redirect_uri: process.env.SALESFORCE_REDIRECT_URI || `${protocol}://localhost:${port}/callback`
        });
    });

    app.post('/api/oauth/callback', async (req, res) => {
        try {
            console.log('[OAuth] コールバック開始');
            const { code } = req.body;
            console.log('[OAuth] 認証コード取得:', code ? 'あり' : 'なし');

            if (!code) {
                console.log('[OAuth] エラー: 認証コードなし');
                return res.status(400).json({ error: 'Authorization code is required' });
            }

            const params = new URLSearchParams();
            params.append('grant_type', 'authorization_code');
            params.append('client_id', process.env.SALESFORCE_CLIENT_ID);
            params.append('client_secret', process.env.SALESFORCE_CLIENT_SECRET);
            params.append('redirect_uri', process.env.SALESFORCE_REDIRECT_URI || 'http://localhost:3001/callback');
            params.append('code', code);

            console.log('[OAuth] Salesforceトークンリクエスト開始');
            console.log('[OAuth] Client ID:', process.env.SALESFORCE_CLIENT_ID);
            console.log('[OAuth] Redirect URI:', process.env.SALESFORCE_REDIRECT_URI || 'http://localhost:3001/callback');

            const response = await fetch('https://login.salesforce.com/services/oauth2/token', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: params
            });

            console.log('[OAuth] Salesforceレスポンス:', response.status, response.statusText);

            if (response.ok) {
                const tokenData = await response.json();
                console.log('[OAuth] トークン取得成功');
                console.log('[OAuth] Instance URL:', tokenData.instance_url);
                res.json({
                    success: true,
                    access_token: tokenData.access_token,
                    instance_url: tokenData.instance_url
                });
            } else {
                const error = await response.json();
                console.error('[OAuth] Salesforceエラー:', error);
                throw new Error(error.error_description || error.error);
            }
        } catch (error) {
            console.error('[OAuth] 全体エラー:', error.message);
            res.status(500).json({
                error: `認証に失敗しました: ${error.message}`,
                details: error.toString()
            });
        }
    });

    // Salesforce APIプロキシエンドポイント
    app.post('/api/salesforce/:objectType', async (req, res) => {
        try {
            console.log('[Salesforce API] リクエスト開始');
            console.log('[Salesforce API] オブジェクトタイプ:', req.params.objectType);

            const { recordData, accessToken, instanceUrl } = req.body;

            if (!recordData || !accessToken || !instanceUrl) {
                console.log('[Salesforce API] エラー: 必要なパラメータが不足');
                return res.status(400).json({
                    error: 'recordData, accessToken, instanceUrl が必要です'
                });
            }

            console.log('[Salesforce API] Salesforceリクエスト送信中...');
            console.log('[Salesforce API] URL:', `${instanceUrl}/services/data/v57.0/sobjects/${req.params.objectType}/`);

            const salesforceResponse = await fetch(`${instanceUrl}/services/data/v57.0/sobjects/${req.params.objectType}/`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(recordData)
            });

            console.log('[Salesforce API] Salesforceレスポンス:', salesforceResponse.status, salesforceResponse.statusText);

            if (salesforceResponse.ok) {
                const result = await salesforceResponse.json();
                console.log('[Salesforce API] レコード作成成功:', result.id);
                res.json({
                    success: true,
                    id: result.id,
                    message: 'レコードが正常に作成されました'
                });
            } else {
                const errorData = await salesforceResponse.json();
                console.error('[Salesforce API] Salesforceエラー:', errorData);
                res.status(salesforceResponse.status).json({
                    error: 'Salesforceでエラーが発生しました',
                    details: errorData
                });
            }
        } catch (error) {
            console.error('[Salesforce API] 全体エラー:', error.message);
            res.status(500).json({
                error: `Salesforce API呼び出しでエラーが発生しました: ${error.message}`,
                details: error.toString()
            });
        }
    });

    // ヘルス チェック
    app.get('/health', (req, res) => {
        res.json({
            status: 'OK',
            protocol: req.protocol,
            port: req.get('host'),
            timestamp: new Date().toISOString(),
            server_type: 'HYBRID (HTTP + HTTPS)'
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

// HTTP サーバー起動（開発・デバッグ用）
const httpServer = http.createServer(app);
httpServer.listen(HTTP_PORT, () => {
    console.log(`📡 HTTP Server running at http://localhost:${HTTP_PORT} (development/debugging)`);
});

// HTTPS サーバー起動（Office Add-in用）
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
                socket.end('HTTP/1.1 400 Bad Request\\r\\n\\r\\n');
            }
        });

        httpsServer.on('error', (err) => {
            console.log('HTTPS Server error:', err.message);
        });

        httpsServer.listen(HTTPS_PORT, () => {
            console.log(`🔒 HTTPS Server running at https://localhost:${HTTPS_PORT} (Office Add-in)`);
            console.log('\\n=== 🎯 READY FOR OUTLOOK ADD-IN ===');
            console.log('✅ Both HTTP and HTTPS servers are running');
            console.log('✅ Use HTTPS for Office Add-in: https://localhost:3000');
            console.log('✅ Use HTTP for development: http://localhost:3001');
            console.log('⚠️  Accept SSL certificate warning in browser for HTTPS to work');
            console.log('=====================================\\n');
        });
    } else {
        console.log('⚠️  HTTPS certificates not found - running HTTP only');
        console.log('💡 Run: npx office-addin-dev-certs install');
        console.log('📡 HTTP Server ready at http://localhost:3001');
    }
} catch (error) {
    console.error('HTTPS setup failed:', error.message);
    console.log('📡 Falling back to HTTP only at http://localhost:3001');
    console.log('💡 Run: npx office-addin-dev-certs install');
}