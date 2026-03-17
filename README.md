# Outlook Salesforce 連携アドイン

Web版Outlook.comで動作するアドインです。メール内容をワンクリックでSalesforceの各種オブジェクト（リード、取引先責任者、ケースなど）に反映することができます。

## 機能

- 📧 **メール情報の自動取得**: 件名、送信者、受信日時、本文を自動で取得
- 🔐 **Salesforce OAuth認証**: 安全なOAuth 2.0フローでSalesforceに接続
- 📋 **複数オブジェクト対応**: リード、取引先責任者、取引先、商談、ケース、活動に対応
- 🎨 **直感的なUI**: Bootstrap+Officeデザインシステムを使用したモダンなインターフェース
- 🔄 **リアルタイム更新**: 入力内容に応じて送信ボタンが自動で有効/無効化
- 🚀 **CORS回避**: サーバープロキシ経由でSalesforce APIに安全にアクセス
- 🔧 **詳細ログ**: 日本語での詳細なデバッグ情報を出力

## 前提条件

- Node.js 16.0.0以上
- Salesforce開発者アカウント
- Microsoft 365 アカウント（Outlook Web Access権限）

## セットアップ

### 1. プロジェクトのクローンとインストール

```bash
git clone <repository-url>
cd outlook-salesforce-addon
npm install
```

### 2. Salesforce Connected App の設定

1. Salesforce Setup → App Manager → New Connected App
2. 以下の設定を行う:
   - **Connected App Name**: `Outlook Salesforce Addon`
   - **API Name**: `Outlook_Salesforce_Addon`
   - **Contact Email**: あなたのメールアドレス
   - **Enable OAuth Settings**: チェック
   - **Callback URL**: `http://localhost:3001/callback` ⚠️ **HTTP**を使用
   - **Selected OAuth Scopes**:
     - `Full access (full)`
     - `Perform requests at any time (refresh_token, offline_access)`
     - `Access the identity URL service (id)`

3. **重要**: ポリシー設定で以下を設定:
   - **IP Restrictions**: Relax IP restrictions
   - **Permitted Users**: All users may self-authorize
   - **Require Secret for Web Server Flow**: チェックを外す
   - **Require Proof Key for Code Exchange (PKCE)**: チェックを外す

4. Consumer Key と Consumer Secret を控える

📋 **詳細な設定手順**: `CREATE_CONNECTED_APP.md` ファイルも参照してください。

### 3. 環境変数の設定

```bash
cp .env.example .env
```

`.env`ファイルを編集:

```env
SALESFORCE_CLIENT_ID=your_connected_app_consumer_key_here
SALESFORCE_CLIENT_SECRET=your_connected_app_consumer_secret_here
SALESFORCE_REDIRECT_URI=http://localhost:3001/callback
PORT=3000
NODE_ENV=development
```

### 4. SSL証明書の生成（開発用）

Office Add-inはHTTPS必須のため、開発用SSL証明書を生成:

```bash
npx office-addin-dev-certs install
```

### 5. 開発サーバーの起動

```bash
npm start
```

デュアルサーバーが起動します：
- HTTPS: `https://localhost:3000` (Office Add-in用)
- HTTP: `http://localhost:3001` (OAuth認証用)

**初回起動時**: ブラウザで `https://localhost:3000` にアクセスし、SSL証明書の警告を受け入れてください。

⚠️ **重要**: OAuth認証は`http://localhost:3001`で動作し、Office Add-inは`https://localhost:3000`で動作します。

### 6. Outlook にアドインを追加

1. Web版Outlook (https://outlook.office365.com) にアクセス
2. 設定 (⚙️) → アドインの管理 → カスタムアドインを追加
3. ファイルから追加を選択
4. `manifest/manifest.xml` ファイルをアップロード
5. アドインが追加されたら、メールを開いてリボンの「Salesforceに送信」ボタンを確認

## 使用方法

1. **メールを選択**: Outlookでメールを開く
2. **アドインを起動**: リボンの「Salesforceに送信」ボタンをクリック
3. **Salesforce認証**: 初回使用時はSalesforceにログイン
4. **オブジェクト選択**: 作成したいSalesforceオブジェクトを選択
5. **内容確認**: 自動取得されたメール情報を確認・編集
6. **送信**: 「Salesforceに送信」ボタンをクリック

## サポートするSalesforceオブジェクト

| オブジェクト | 用途 |
|------------|------|
| **Lead (リード)** | 見込み客の管理 |
| **Contact (取引先責任者)** | 既存顧客の連絡先 |
| **Account (取引先)** | 会社・組織の管理 |
| **Opportunity (商談)** | セールス機会の追跡 |
| **Case (ケース)** | サポートリクエスト |
| **Task (活動)** | フォローアップ活動 |

## 開発

### ディレクトリ構造

```
outlook-salesforce-addon/
├── manifest/
│   └── manifest.xml          # Office Add-in マニフェスト
├── src/
│   ├── taskpane.html         # メインUI
│   ├── taskpane.js           # メイン処理（Office.js統合）
│   ├── taskpane.css          # スタイル
│   ├── commands.html         # コマンド用HTML
│   ├── commands.js           # コマンド処理
│   └── cllback.html          # OAuth認証コールバック
├── assets/                   # 画像・アイコン（PNG形式）
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   └── logo-filled.png
├── server.js                 # デュアル（HTTP+HTTPS）サーバー + Salesforce API プロキシ
├── package.json
├── .env                      # 環境変数設定
└── README.md
```

### スクリプト

```bash
npm start            # デュアルサーバー起動（HTTP:3001 + HTTPS:3000）
npm run dev          # 開発サーバー起動（nodemon使用）
npm run lint         # ESLintでコード解析
npm run lint:fix     # ESLintで自動修正
npm run validate     # マニフェストファイル検証
```

### サーバー仕様

このプロジェクトは最大互換性のためデュアルサーバーを実装：

- **HTTPS Server**: `https://localhost:3000` (Office Add-in用)
- **HTTP Server**: `http://localhost:3001` (OAuth認証用)
- **自動SSL証明書**: Office Add-in開発用証明書を自動検出
- **CORS完全対応**: すべてのオリジンを許可（開発環境）
- **Salesforce APIプロキシ**: `/api/salesforce/:objectType` でCORSエラーを回避

### デバッグ

このアドインは詳細なデバッグ機能を実装しています：

1. **ブラウザ開発者ツール**: F12キーで開発者ツールを起動
2. **コンソールログ**: 段階的デバッグ情報をコンソールに出力
   - Office.js初期化プロセス
   - メール情報取得の各段階
   - Salesforce API通信
   - 認証フロー
3. **ネットワークタブ**: API通信を監視
4. **エラー表示**: UI上でエラー原因を表示

**重要なデバッグポイント**:
- メール情報が取得できない場合は、コンソールで`=== Office.onReady called ===`などのログを確認
- 認証エラーの場合は、Salesforce Connected Appの設定を確認

## トラブルシューティング

### よくある問題

**Q: アドインが表示されない**
- マニフェストファイルが正しくアップロードされているか確認
- HTTPSでサーバーが起動しているか確認（`https://localhost:3000`）
- ブラウザのキャッシュをクリア
- SSL証明書の警告を受け入れる

**Q: Salesforce認証でエラーが出る**
- Connected Appの設定を確認
- Callback URLが正確に設定されているか確認（`http://localhost:3001/callback` ⚠️**HTTP**）
- Client IDとClient Secretが正しいか確認
- .envファイルの環境変数を確認
- ブラウザのポップアップブロックを無効化
- **「Cross-org OAuth flows not supported」**エラー: 自分のSalesforce組織でConnected Appを作成

**Q: メール情報が取得できない（"(件名API利用不可)"等が表示）**
- ブラウザのコンソールでOffice.jsの初期化ログを確認
- Outlookでメールを選択してからアドインを起動
- マニフェストファイルの権限設定を確認（`ReadWriteMailbox`）
- Office.jsのバージョン互換性を確認

**Q: レコード作成でエラーが出る / 「Failed to fetch」エラー**
- ⚠️ **修正済み**: サーバープロキシ経由でCORSエラーを回避
- サーバーログで`[Salesforce API]`メッセージを確認
- Salesforceのユーザー権限を確認
- 必須フィールドが適切に設定されているか確認
- APIバージョンの互換性を確認（v57.0使用）

### ログ確認

**サーバーログ**:
```bash
npm start  # コンソールでサーバーログを確認
# 以下のような出力が表示されます：
# 🚀 Starting HYBRID SERVER (HTTP + HTTPS) for Office Add-in compatibility
# 📡 HTTP Server running at http://localhost:3001 (development/debugging)
# 🔒 HTTPS Server running at https://localhost:3000 (Office Add-in)
# [OAuth] コールバック開始
# [OAuth] 認証コード取得: あり
# [Salesforce API] リクエスト開始
```

**ブラウザログ**:
1. F12 → Console タブ
2. 以下のようなデバッグ情報を確認：
   - `=== Office.onReady called ===`
   - `=== loadEmailInfo() called ===`
   - `=== Attempting to get subject ===`
3. ネットワークエラーの場合は Network タブも確認

**主要なログパターン**:
- ✅ 成功: `✅ Subject successfully retrieved:`、`[OAuth] トークン取得成功`、`[Salesforce API] レコード作成成功:`
- ❌ エラー: `❌ item.subject.getAsync is not available`、`[OAuth] エラー: 認証コードなし`
- ⚠️ 警告: `❌ No subject property available`、`[Salesforce API] エラー: 必要なパラメータが不足`

## セキュリティ

- OAuth 2.0による安全な認証
- HTTPS必須（本番環境）
- CORS設定によるオリジン制限
- CSP（Content Security Policy）による XSS 防御
- 機密情報の適切な管理
- サーバープロキシによるAPIセキュリティ強化
- トークンの適切なライフサイクル管理

## ライセンス

MIT License

## サポート

問題が発生した場合は、GitHubのIssueで報告してください。

## 更新履歴

### v1.1.0 (2026-03-17)
- ✅ **CORSエラー修正**: サーバープロキシ経由でSalesforce API呼び出し
- ✅ **詳細ログ実装**: 日本語での詳細なデバッグ情報
- ✅ **OAuth認証改善**: HTTP認証フローに変更
- ✅ **エラーハンドリング強化**: より詳細なエラー情報を提供

### v1.0.0 (2026-03-16)
- 🎉 **初期リリース**: 基本的なOutlook-Salesforce連携機能
- 📧 メール情報の自動取得
- 🔐 OAuth 2.0認証
- 📋 複数Salesforceオブジェクト対応

## 貢献

プルリクエストを歓迎します。大きな変更の場合は、まずIssueで議論してください。

## 技術スタック

- **フロントエンド**: HTML5, CSS3, JavaScript ES6+, Bootstrap 5
- **バックエンド**: Node.js, Express.js
- **Office統合**: Office.js API
- **認証**: Salesforce OAuth 2.0
- **API**: Salesforce REST API v57.0
- **開発ツール**: SSL証明書、デュアルサーバー構成
