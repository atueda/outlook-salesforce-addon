# Salesforce Connected App 作成手順

## 手順

### 1. Salesforceにログイン
- Setup → Apps → App Manager → New Connected App

### 2. 基本情報
- **Connected App Name**: `Outlook Salesforce Addon`
- **API Name**: `Outlook_Salesforce_Addon`
- **Contact Email**: あなたのメールアドレス

### 3. OAuth設定 ✅
- **Enable OAuth Settings**: チェック
- **Callback URL**:
```
http://localhost:3001/callback
```

- **Selected OAuth Scopes**:
  - Full access (full)
  - Perform requests at any time (refresh_token, offline_access)
  - Access the identity URL service (id)

### 4. 追加設定
- **Require Secret for Web Server Flow**: チェックを外す（開発環境用）
- **Require Proof Key for Code Exchange (PKCE)**: チェックを外す

### 5. 保存後
1. Consumer Keyをコピー
2. Consumer Secretをコピー
3. .envファイルを更新

### 6. ポリシー設定（重要！）
保存後、「Manage」をクリック:
- **IP Restrictions**: Relax IP restrictions
- **Permitted Users**: All users may self-authorize

## 作成後に必要な操作

1. Consumer Key と Consumer Secret を下記の.envに更新:

```env
SALESFORCE_CLIENT_ID=【新しいConsumer Key】
SALESFORCE_CLIENT_SECRET=【新しいConsumer Secret】
SALESFORCE_REDIRECT_URI=http://localhost:3001/callback
```

2. サーバーを再起動