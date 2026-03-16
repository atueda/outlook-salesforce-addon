/**
 * Salesforce REST API クライアント
 * Salesforceとの通信を管理するクラス
 */
class SalesforceClient {
    constructor() {
        this.accessToken = null;
        this.instanceUrl = null;
        this.apiVersion = 'v57.0';
    }

    /**
     * 認証状態を初期化
     */
    initialize() {
        this.accessToken = localStorage.getItem('salesforce_access_token');
        this.instanceUrl = localStorage.getItem('salesforce_instance_url');
        return this.isAuthenticated();
    }

    /**
     * 認証状態を確認
     */
    isAuthenticated() {
        return !!(this.accessToken && this.instanceUrl);
    }

    /**
     * 認証情報をクリア
     */
    clearAuthentication() {
        this.accessToken = null;
        this.instanceUrl = null;
        localStorage.removeItem('salesforce_access_token');
        localStorage.removeItem('salesforce_instance_url');
    }

    /**
     * APIリクエストのベースメソッド
     */
    async makeApiRequest(endpoint, options = {}) {
        if (!this.isAuthenticated()) {
            throw new Error('Salesforceに認証されていません');
        }

        const url = `${this.instanceUrl}/services/data/${this.apiVersion}/${endpoint}`;
        const headers = {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
            ...options.headers
        };

        try {
            const response = await fetch(url, {
                ...options,
                headers
            });

            // 認証エラーの場合
            if (response.status === 401) {
                this.clearAuthentication();
                throw new Error('認証が無効です。再度ログインしてください。');
            }

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.message || `API Error: ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.error('Salesforce API Error:', error);
            throw error;
        }
    }

    /**
     * レコードを作成
     */
    async createRecord(sobjectType, recordData) {
        return await this.makeApiRequest(`sobjects/${sobjectType}/`, {
            method: 'POST',
            body: JSON.stringify(recordData)
        });
    }

    /**
     * レコードを取得
     */
    async getRecord(sobjectType, recordId, fields = null) {
        let endpoint = `sobjects/${sobjectType}/${recordId}`;
        if (fields) {
            endpoint += `?fields=${fields.join(',')}`;
        }
        return await this.makeApiRequest(endpoint);
    }

    /**
     * レコードを更新
     */
    async updateRecord(sobjectType, recordId, recordData) {
        return await this.makeApiRequest(`sobjects/${sobjectType}/${recordId}`, {
            method: 'PATCH',
            body: JSON.stringify(recordData)
        });
    }

    /**
     * レコードを削除
     */
    async deleteRecord(sobjectType, recordId) {
        return await this.makeApiRequest(`sobjects/${sobjectType}/${recordId}`, {
            method: 'DELETE'
        });
    }

    /**
     * SOQLクエリを実行
     */
    async query(soql) {
        const encodedQuery = encodeURIComponent(soql);
        return await this.makeApiRequest(`query/?q=${encodedQuery}`);
    }

    /**
     * オブジェクト情報を取得
     */
    async describeSObject(sobjectType) {
        return await this.makeApiRequest(`sobjects/${sobjectType}/describe/`);
    }

    /**
     * 利用可能なオブジェクト一覧を取得
     */
    async getSObjects() {
        return await this.makeApiRequest('sobjects/');
    }

    /**
     * ユーザー情報を取得
     */
    async getUserInfo() {
        const identity = await this.makeApiRequest('../oauth2/userinfo', {
            headers: {
                'Authorization': `Bearer ${this.accessToken}`
            }
        });
        return identity;
    }

    /**
     * 組織情報を取得
     */
    async getOrganizationInfo() {
        const result = await this.query("SELECT Id, Name, Country FROM Organization LIMIT 1");
        return result.records.length > 0 ? result.records[0] : null;
    }

    /**
     * メールからリードを検索
     */
    async findLeadByEmail(email) {
        const soql = `SELECT Id, Name, Email, Company FROM Lead WHERE Email = '${email}' LIMIT 1`;
        const result = await this.query(soql);
        return result.records.length > 0 ? result.records[0] : null;
    }

    /**
     * メールから取引先責任者を検索
     */
    async findContactByEmail(email) {
        const soql = `SELECT Id, Name, Email, AccountId, Account.Name FROM Contact WHERE Email = '${email}' LIMIT 1`;
        const result = await this.query(soql);
        return result.records.length > 0 ? result.records[0] : null;
    }

    /**
     * カスタムオブジェクトにレコードを作成（フレキシブル）
     */
    async createCustomRecord(objectApiName, recordData) {
        return await this.createRecord(objectApiName, recordData);
    }

    /**
     * ファイルをSalesforceにアップロード
     */
    async uploadFile(recordId, fileName, fileContent, contentType = 'application/octet-stream') {
        // ContentVersionを作成してファイルをアップロード
        const contentVersion = {
            Title: fileName,
            PathOnClient: fileName,
            VersionData: btoa(fileContent), // Base64エンコード
            IsMajorVersion: true
        };

        const versionResult = await this.createRecord('ContentVersion', contentVersion);

        // ContentDocumentLinkを作成してレコードにリンク
        if (recordId && versionResult.Id) {
            const documentId = await this.query(`SELECT ContentDocumentId FROM ContentVersion WHERE Id = '${versionResult.Id}'`);
            if (documentId.records.length > 0) {
                const link = {
                    ContentDocumentId: documentId.records[0].ContentDocumentId,
                    LinkedEntityId: recordId,
                    ShareType: 'V'
                };
                await this.createRecord('ContentDocumentLink', link);
            }
        }

        return versionResult;
    }
}

// グローバルインスタンスを作成
const salesforceClient = new SalesforceClient();