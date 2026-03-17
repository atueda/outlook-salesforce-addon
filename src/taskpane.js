/* global Office */

// デバッグ情報を更新する関数
function updateDebugInfo(status, host, url) {
    if (document.getElementById('office-status')) {
        document.getElementById('office-status').textContent = status || 'Office.js 読み込み済み';
    }
    if (document.getElementById('host-info')) {
        document.getElementById('host-info').textContent = host || 'ホスト情報不明';
    }
    if (document.getElementById('url-display')) {
        document.getElementById('url-display').textContent = url || window.location.href;
    }
}

// ページ読み込み時の初期化
window.addEventListener('load', () => {
    updateDebugInfo('ページ読み込み完了', 'ホスト情報を待機中', window.location.href);
});

// Office.js の初期化
Office.onReady((info) => {
    console.log('=== Office.onReady called ===');
    console.log('Host info:', info);
    console.log('Host type:', info.host);
    console.log('Platform:', info.platform);

    updateDebugInfo(
        'Office.js 初期化完了',
        `ホスト: ${info.host}, プラットフォーム: ${info.platform}`,
        window.location.href
    );

    if (info.host === Office.HostType.Outlook) {
        console.log('✅ Outlook host confirmed');
        document.getElementById('sideload-msg').style.display = 'none';
        document.getElementById('app-body').style.display = 'block';

        // イベントリスナーを設定
        setupEventListeners();

        // メール情報取得前にOfficeコンテキストをさらに詳しく確認
        console.log('=== Checking Office context before loading email ===');
        console.log('Office.context:', Office.context);
        console.log('Office.context.mailbox:', Office.context.mailbox);
        console.log('Office.context.mailbox.item:', Office.context.mailbox.item);

        // 少し遅延してメール情報を取得（コンテキストが完全に初期化されるまで待つ）
        setTimeout(() => {
            console.log('=== Attempting to load email info after delay ===');
            loadEmailInfo();
        }, 1000);

        // Salesforce認証状態をチェック
        checkAuthenticationStatus();
    } else {
        console.error('❌ Not running in Outlook host, current host:', info.host);
        updateDebugInfo(
            'エラー: Outlook以外のホストで実行中',
            `現在のホスト: ${info.host}`,
            window.location.href
        );
    }
}).catch(error => {
    console.error('❌ Office.onReady failed:', error);
    updateDebugInfo(
        `Office.js エラー: ${error.message}`,
        'エラーが発生しました',
        window.location.href
    );
});

// イベントリスナーの設定
function setupEventListeners() {
    document.getElementById('login-btn').addEventListener('click', authenticateWithSalesforce);
    document.getElementById('logout-btn').addEventListener('click', logout);
    document.getElementById('object-type').addEventListener('change', onObjectTypeChange);
    document.getElementById('send-to-salesforce').addEventListener('click', sendToSalesforce);
}

// メール情報の取得
function loadEmailInfo() {
    console.log('=== loadEmailInfo() called ===');
    console.log('Office object available:', typeof Office !== 'undefined');
    console.log('Office.context available:', !!Office?.context);
    console.log('Office.context.mailbox available:', !!Office?.context?.mailbox);
    console.log('Office.context.mailbox.item available:', !!Office?.context?.mailbox?.item);

    // Office.js の基本チェック
    if (typeof Office === 'undefined') {
        console.error('Office.js is not loaded');
        setDefaultEmailInfo();
        return;
    }

    if (!Office.context || !Office.context.mailbox) {
        console.error('Office.context.mailbox is not available');
        setDefaultEmailInfo();
        return;
    }

    if (!Office.context.mailbox.item) {
        console.error('Office.context.mailbox.item is not available');
        setDefaultEmailInfo();
        return;
    }

    const item = Office.context.mailbox.item;
    console.log('Mail item available:', !!item);
    console.log('Mail item type:', item?.itemType);
    console.log('Mail item subject property:', typeof item?.subject);
    console.log('Mail item from property:', typeof item?.from);
    console.log('Mail item dateTimeCreated property:', typeof item?.dateTimeCreated);
    console.log('Mail item body property:', typeof item?.body);

    // 直接プロパティアクセスを試す（非同期APIが利用できない場合の代替手段）
    console.log('=== Checking direct property access ===');
    try {
        if (item.subject && typeof item.subject === 'string') {
            console.log('Direct subject access available:', item.subject);
        } else if (item.subject && item.subject.value) {
            console.log('Subject value property available:', item.subject.value);
        }

        if (item.from && typeof item.from === 'object') {
            console.log('Direct from access available:', item.from);
        }

        if (item.dateTimeCreated) {
            console.log('Direct dateTimeCreated access:', typeof item.dateTimeCreated, item.dateTimeCreated);
        }
    } catch (error) {
        console.log('Direct property access failed:', error);
    }

    // 件名を取得（Office.js 1.7以降の新しいAPI方式を試す）
    console.log('=== Attempting to get subject ===');
    try {
        // 新しいAPI方式 (Office.js 1.7+) を試す
        if (item.subject) {
            console.log('Subject property exists, checking access methods...');

            // 直接値アクセスを試す（読み取り専用モードでは値が直接取得可能）
            if (typeof item.subject === 'string') {
                console.log('✅ Direct subject access available:', item.subject);
                document.getElementById('email-subject').textContent = item.subject;
                const recordNameField = document.getElementById('record-name');
                if (recordNameField) {
                    recordNameField.value = item.subject;
                }
            }
            // getAsync方式を試す
            else if (typeof item.subject.getAsync === 'function') {
                console.log('Subject getAsync method available, calling...');

                item.subject.getAsync((result) => {
                    console.log('Subject API response received:', result);
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const subject = result.value || '(件名なし)';
                        console.log('✅ Subject successfully retrieved:', subject);
                        document.getElementById('email-subject').textContent = subject;
                        const recordNameField = document.getElementById('record-name');
                        if (recordNameField) {
                            recordNameField.value = subject;
                        }
                    } else {
                        console.error('❌ Failed to get subject:', result.error);
                        document.getElementById('email-subject').textContent = '(件名取得エラー)';
                    }
                });
            }
            // 現在の読み取り値を試す
            else if (item.subject.value !== undefined) {
                console.log('✅ Subject value property available:', item.subject.value);
                document.getElementById('email-subject').textContent = item.subject.value || '(件名なし)';
                const recordNameField = document.getElementById('record-name');
                if (recordNameField) {
                    recordNameField.value = item.subject.value || '(件名なし)';
                }
            } else {
                console.warn('❌ Unknown subject property structure');
                console.log('Subject object:', item.subject);
                document.getElementById('email-subject').textContent = '(件名形式不明)';
            }
        } else {
            console.warn('❌ No subject property available');
            document.getElementById('email-subject').textContent = '(件名プロパティなし)';
        }
    } catch (error) {
        console.error('❌ Exception getting subject:', error);
        document.getElementById('email-subject').textContent = '(件名例外エラー)';
    }

    // 送信者を取得
    console.log('=== Attempting to get sender ===');
    try {
        if (item.from) {
            console.log('From property exists, checking access methods...');
            console.log('From property type:', typeof item.from);
            console.log('From property structure:', item.from);

            // 直接オブジェクトアクセスを試す
            if (item.from.emailAddress || item.from.displayName) {
                console.log('✅ Direct from access available');
                const senderText = `${item.from.displayName || item.from.emailAddress || 'Unknown'} (${item.from.emailAddress || 'no-email'})`;
                document.getElementById('email-sender').textContent = senderText;
            }
            // getAsync方式を試す
            else if (typeof item.from.getAsync === 'function') {
                console.log('From getAsync method available, calling...');
                item.from.getAsync((result) => {
                    console.log('From API response received:', result);
                    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                        const sender = result.value;
                        const senderText = `${sender.displayName || sender.emailAddress} (${sender.emailAddress})`;
                        console.log('✅ Sender successfully retrieved:', senderText);
                        document.getElementById('email-sender').textContent = senderText;
                    } else {
                        console.error('❌ Failed to get sender:', result.error);
                        document.getElementById('email-sender').textContent = '(送信者取得エラー)';
                    }
                });
            }
            // valueプロパティを試す
            else if (item.from.value) {
                console.log('✅ From value property available');
                const sender = item.from.value;
                const senderText = `${sender.displayName || sender.emailAddress} (${sender.emailAddress})`;
                document.getElementById('email-sender').textContent = senderText;
            } else {
                console.warn('❌ Unknown from property structure');
                document.getElementById('email-sender').textContent = '(送信者形式不明)';
            }
        } else {
            console.warn('❌ No from property available');
            document.getElementById('email-sender').textContent = '(送信者プロパティなし)';
        }
    } catch (error) {
        console.error('❌ Exception getting sender:', error);
        document.getElementById('email-sender').textContent = '(送信者例外エラー)';
    }

    // 日時を取得
    console.log('=== Attempting to get date ===');
    try {
        if (item.dateTimeCreated) {
            console.log('DateTimeCreated property exists, checking access methods...');
            console.log('DateTimeCreated property type:', typeof item.dateTimeCreated);
            console.log('DateTimeCreated property value:', item.dateTimeCreated);

            // 直接Dateオブジェクトアクセスを試す
            if (item.dateTimeCreated instanceof Date) {
                console.log('✅ Direct Date object access available');
                document.getElementById('email-date').textContent = item.dateTimeCreated.toLocaleString('ja-JP');
            }
            // 直接文字列アクセスを試す
            else if (typeof item.dateTimeCreated === 'string') {
                console.log('✅ Direct date string access available');
                const date = new Date(item.dateTimeCreated);
                document.getElementById('email-date').textContent = date.toLocaleString('ja-JP');
            }
            // getAsync方式を試す
            else if (typeof item.dateTimeCreated.getAsync === 'function') {
                console.log('DateTimeCreated getAsync method available, calling...');
                item.dateTimeCreated.getAsync((result) => {
                    console.log('Date API response received:', result);
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const date = new Date(result.value);
                        console.log('✅ Date successfully retrieved:', date);
                        document.getElementById('email-date').textContent = date.toLocaleString('ja-JP');
                    } else {
                        console.error('❌ Failed to get date:', result.error);
                        document.getElementById('email-date').textContent = '(日時取得エラー)';
                    }
                });
            }
            // valueプロパティを試す
            else if (item.dateTimeCreated.value) {
                console.log('✅ DateTimeCreated value property available');
                const date = new Date(item.dateTimeCreated.value);
                document.getElementById('email-date').textContent = date.toLocaleString('ja-JP');
            } else {
                console.warn('❌ Unknown dateTimeCreated property structure');
                document.getElementById('email-date').textContent = '(日時形式不明)';
            }
        } else {
            console.warn('❌ No dateTimeCreated property available');
            document.getElementById('email-date').textContent = '(日時プロパティなし)';
        }
    } catch (error) {
        console.error('❌ Exception getting date:', error);
        document.getElementById('email-date').textContent = '(日時例外エラー)';
    }

    // 本文を取得
    try {
        if (item.body && typeof item.body.getAsync === 'function') {
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                console.log('Body result:', result);
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const descriptionField = document.getElementById('description');
                    if (descriptionField) {
                        const subject = document.getElementById('email-subject').textContent;
                        const sender = document.getElementById('email-sender').textContent;
                        const date = document.getElementById('email-date').textContent;

                        const emailInfo = `件名: ${subject}
送信者: ${sender}
受信日時: ${date}

${result.value || 'メール本文がありません'}`;

                        descriptionField.value = emailInfo;
                    }
                } else {
                    console.error('Failed to get body:', result.error);
                    const descriptionField = document.getElementById('description');
                    if (descriptionField) {
                        const subject = document.getElementById('email-subject').textContent;
                        const sender = document.getElementById('email-sender').textContent;
                        const date = document.getElementById('email-date').textContent;

                        descriptionField.value = `件名: ${subject}
送信者: ${sender}
受信日時: ${date}

(メール本文の取得でエラーが発生しました)`;
                    }
                }
            });
        } else {
            console.warn('item.body.getAsync is not available');
            const descriptionField = document.getElementById('description');
            if (descriptionField) {
                const subject = document.getElementById('email-subject').textContent;
                const sender = document.getElementById('email-sender').textContent;
                const date = document.getElementById('email-date').textContent;

                descriptionField.value = `件名: ${subject}
送信者: ${sender}
受信日時: ${date}

(メール本文を取得できません)`;
            }
        }
    } catch (error) {
        console.error('Error getting body:', error);
        const descriptionField = document.getElementById('description');
        if (descriptionField) {
            descriptionField.value = '(メール本文取得エラー)';
        }
    }
}

// デフォルトのメール情報を設定
function setDefaultEmailInfo() {
    console.log('Setting default email info - no email context available');
    document.getElementById('email-subject').textContent = '(メールコンテキストなし)';
    document.getElementById('email-sender').textContent = '(メールコンテキストなし)';
    document.getElementById('email-date').textContent = '(メールコンテキストなし)';

    const descriptionField = document.getElementById('description');
    if (descriptionField) {
        descriptionField.value = '(メールコンテキストが利用できません。Outlookでメールを選択してから実行してください。)';
    }
}

// Salesforce認証状態のチェック
function checkAuthenticationStatus() {
    console.log('checkAuthenticationStatus() called');

    const accessToken = localStorage.getItem('salesforce_access_token');
    const instanceUrl = localStorage.getItem('salesforce_instance_url');

    console.log('Auth check:', {
        hasAccessToken: !!accessToken,
        hasInstanceUrl: !!instanceUrl
    });

    if (accessToken && instanceUrl) {
        console.log('Tokens found, validating...');
        // トークンの有効性を確認
        validateToken(accessToken, instanceUrl).then(isValid => {
            console.log('Token validation result:', isValid);
            if (isValid) {
                showAuthenticatedState();
            } else {
                showUnauthenticatedState();
            }
        }).catch(error => {
            console.error('Token validation failed:', error);
            showUnauthenticatedState();
        });
    } else {
        console.log('No tokens found, showing unauthenticated state');
        showUnauthenticatedState();
    }
}

// トークンの有効性確認
async function validateToken(accessToken, instanceUrl) {
    try {
        const response = await fetch(`${instanceUrl}/services/data/v57.0/sobjects/`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        return response.ok;
    } catch (error) {
        console.error('Token validation failed:', error);
        return false;
    }
}

// 認証済み状態の表示
function showAuthenticatedState() {
    console.log('showAuthenticatedState() called');

    document.getElementById('not-authenticated').style.display = 'none';
    document.getElementById('authenticated').style.display = 'block';
    document.getElementById('salesforce-section').style.display = 'block';

    // メール情報を確実に取得するため、少し遅延
    setTimeout(() => {
        console.log('Loading email info after authentication...');
        loadEmailInfo();

        // 送信ボタンの状態を更新
        setTimeout(() => {
            updateSendButtonState();
        }, 100);
    }, 200);

    // ステータスメッセージを表示
    showStatusMessage('Salesforceに正常に接続されました', 'success');
}

// 未認証状態の表示
function showUnauthenticatedState() {
    document.getElementById('not-authenticated').style.display = 'block';
    document.getElementById('authenticated').style.display = 'none';
    document.getElementById('salesforce-section').style.display = 'none';
    document.getElementById('send-to-salesforce').disabled = true;

    // ローカルストレージをクリア
    localStorage.removeItem('salesforce_access_token');
    localStorage.removeItem('salesforce_instance_url');
}

// Salesforce認証
function authenticateWithSalesforce() {
    showStatusMessage('Salesforce認証を開始しています...', 'info');

    // OAuth 2.0 Web Server Flow を使用
    // 環境変数からクライアントIDを取得（サーバー経由）
    fetch('/api/config')
        .then(response => response.json())
        .then(config => {
            const clientId = config.client_id;
            const redirectUri = encodeURIComponent('http://localhost:3001/callback');
            const authUrl = `https://login.salesforce.com/services/oauth2/authorize?response_type=code&client_id=${clientId}&redirect_uri=${redirectUri}&scope=api`;

            console.log('Opening auth window:', authUrl);

            // 新しいウィンドウで認証ページを開く（より大きなサイズで開く）
            const authWindow = window.open(authUrl, 'salesforce-auth', 'width=800,height=700,scrollbars=yes,resizable=yes,status=yes,toolbar=no,menubar=no');

            if (!authWindow) {
                showStatusMessage('ポップアップがブロックされました。ブラウザの設定を確認してください。', 'error');
                return;
            }

            // デバッグ用：生成されたURLをコンソールに表示
            console.log('Generated OAuth URL:', authUrl);

            // 認証完了を待機
            let checkCount = 0;
            const maxChecks = 600; // 5分間（500ms × 600回）

            const checkClosed = setInterval(() => {
                checkCount++;

                try {
                    // ウィンドウが閉じられたかチェック
                    if (authWindow.closed) {
                        console.log(`Auth window closed after ${checkCount * 0.5} seconds, checking for tokens...`);
                        clearInterval(checkClosed);

                        // 複数回認証状態をチェック（トークンが保存されるまで時間がかかる場合がある）
                        let authCheckCount = 0;
                        const authCheckInterval = setInterval(() => {
                            authCheckCount++;
                            console.log(`Authentication check attempt ${authCheckCount}...`);

                            const accessToken = localStorage.getItem('salesforce_access_token');
                            const instanceUrl = localStorage.getItem('salesforce_instance_url');

                            if (accessToken && instanceUrl) {
                                console.log('✅ Tokens found on attempt', authCheckCount);
                                clearInterval(authCheckInterval);
                                showAuthenticatedState();
                            } else if (authCheckCount >= 10) {
                                console.log('❌ No tokens found after 10 attempts');
                                clearInterval(authCheckInterval);
                                showStatusMessage('認証が完了していません。再度お試しください。', 'warning');
                            }
                        }, 1000);
                        return;
                    }

                    // URLが変更されたかチェック（認証完了の検出）
                    try {
                        const currentUrl = authWindow.location.href;
                        if (currentUrl && currentUrl.includes('localhost:3001/callback')) {
                            console.log('✅ Callback URL detected, authentication may be complete');
                        }
                    } catch (error) {
                        // Cross-origin エラーは無視（正常な動作）
                    }

                } catch (error) {
                    console.error('Error checking auth window:', error);
                    clearInterval(checkClosed);
                    showStatusMessage('認証ウィンドウの監視でエラーが発生しました', 'error');
                }

                // タイムアウトチェック
                if (checkCount >= maxChecks) {
                    console.log('Auth window timeout, closing...');
                    clearInterval(checkClosed);
                    if (!authWindow.closed) {
                        authWindow.close();
                    }
                    showStatusMessage('認証がタイムアウトしました。再試行してください。', 'warning');
                }
            }, 500);

            // ウィンドウフォーカス維持の試行
            try {
                authWindow.focus();
            } catch (error) {
                console.log('Could not focus auth window:', error.message);
            }
        })
        .catch(error => {
            console.error('設定の取得に失敗:', error);
            showStatusMessage('設定の取得に失敗しました。サーバーが起動しているか確認してください。', 'error');
        });
}

// ログアウト
function logout() {
    localStorage.removeItem('salesforce_access_token');
    localStorage.removeItem('salesforce_instance_url');
    showUnauthenticatedState();
    showStatusMessage('ログアウトしました', 'success');
}

// オブジェクトタイプ変更時の処理
function onObjectTypeChange() {
    const objectType = document.getElementById('object-type').value;
    const recordNameInput = document.getElementById('record-name');

    if (objectType) {
        // オブジェクトタイプに応じてデフォルト名を設定
        const subject = document.getElementById('email-subject').textContent;
        let defaultName = subject !== '-' ? subject : 'メールからの問い合わせ';

        switch (objectType) {
            case 'Lead':
                defaultName = `リード: ${defaultName}`;
                break;
            case 'Case':
                defaultName = `ケース: ${defaultName}`;
                break;
            case 'Task':
                defaultName = `活動: ${defaultName}`;
                break;
            case 'Opportunity':
                defaultName = `商談: ${defaultName}`;
                break;
        }

        recordNameInput.value = defaultName;
    }

    updateSendButtonState();
}

// 送信ボタン状態の更新
function updateSendButtonState() {
    console.log('updateSendButtonState() called');

    const objectTypeElement = document.getElementById('object-type');
    const recordNameElement = document.getElementById('record-name');
    const sendButton = document.getElementById('send-to-salesforce');

    if (!objectTypeElement || !recordNameElement || !sendButton) {
        console.error('Required elements not found');
        return;
    }

    const objectType = objectTypeElement.value;
    const recordName = recordNameElement.value;
    const isAuthenticated = localStorage.getItem('salesforce_access_token');

    console.log('Button state check:', {
        objectType: objectType,
        recordName: recordName,
        isAuthenticated: !!isAuthenticated
    });

    const canSend = isAuthenticated && objectType && recordName && recordName.trim().length > 0;
    sendButton.disabled = !canSend;

    console.log('Send button enabled:', !sendButton.disabled);

    // ボタンのスタイルも更新
    if (canSend) {
        sendButton.classList.remove('btn-secondary');
        sendButton.classList.add('btn-success');
    } else {
        sendButton.classList.remove('btn-success');
        sendButton.classList.add('btn-secondary');
    }
}

// Salesforceに送信
async function sendToSalesforce() {
    const objectType = document.getElementById('object-type').value;
    const recordName = document.getElementById('record-name').value;
    const description = document.getElementById('description').value;

    const accessToken = localStorage.getItem('salesforce_access_token');
    const instanceUrl = localStorage.getItem('salesforce_instance_url');

    if (!accessToken || !instanceUrl) {
        showStatusMessage('Salesforce認証が必要です', 'error');
        return;
    }

    showStatusMessage('Salesforceに送信中...', 'info');

    try {
        // Salesforce オブジェクトのデータを準備
        const recordData = buildRecordData(objectType, recordName, description);

        console.log('[送信] サーバー経由でSalesforceにレコードを作成中...');
        console.log('[送信] オブジェクトタイプ:', objectType);
        console.log('[送信] レコードデータ:', recordData);

        // サーバー経由でSalesforceにリクエストを送信（CORSエラー回避）
        const response = await fetch(`/api/salesforce/${objectType}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                recordData: recordData,
                accessToken: accessToken,
                instanceUrl: instanceUrl
            })
        });

        console.log('[送信] サーバーレスポンス:', response.status, response.statusText);

        if (response.ok) {
            const result = await response.json();
            console.log('[送信] レコード作成成功:', result);
            showStatusMessage(`正常に作成されました！ ID: ${result.id}`, 'success');

            // 成功後にフィールドをリセット（オプション）
            // resetForm();
        } else {
            const errorData = await response.json();
            console.error('[送信] サーバーエラー:', errorData);
            throw new Error(errorData.error || 'Salesforceへの送信に失敗しました');
        }
    } catch (error) {
        console.error('[送信] 全体エラー:', error);
        showStatusMessage(`エラー: ${error.message}`, 'error');
    }
}

// レコードデータの構築
function buildRecordData(objectType, recordName, description) {
    const baseData = {
        Description: description
    };

    // メール情報を追加
    const subject = document.getElementById('email-subject').textContent;
    const sender = document.getElementById('email-sender').textContent;
    const date = document.getElementById('email-date').textContent;

    const emailInfo = `件名: ${subject}\n送信者: ${sender}\n受信日時: ${date}\n\n${description}`;

    switch (objectType) {
        case 'Lead':
            return {
                ...baseData,
                LastName: recordName,
                Company: '不明',
                Description: emailInfo,
                LeadSource: 'Email'
            };
        case 'Contact':
            return {
                ...baseData,
                LastName: recordName,
                Description: emailInfo
            };
        case 'Account':
            return {
                ...baseData,
                Name: recordName,
                Description: emailInfo
            };
        case 'Opportunity':
            return {
                ...baseData,
                Name: recordName,
                StageName: 'Prospecting',
                CloseDate: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // 30日後
                Description: emailInfo
            };
        case 'Case':
            return {
                ...baseData,
                Subject: recordName,
                Description: emailInfo,
                Origin: 'Email',
                Status: 'New'
            };
        case 'Task':
            return {
                ...baseData,
                Subject: recordName,
                Description: emailInfo,
                Status: 'Not Started',
                Priority: 'Normal',
                ActivityDate: new Date().toISOString().split('T')[0]
            };
        default:
            return baseData;
    }
}

// ステータスメッセージの表示
function showStatusMessage(message, type) {
    const statusDiv = document.getElementById('status-message');
    let alertClass = 'alert-info';

    switch (type) {
        case 'success':
            alertClass = 'alert-success';
            break;
        case 'error':
            alertClass = 'alert-danger';
            break;
        case 'warning':
            alertClass = 'alert-warning';
            break;
    }

    statusDiv.innerHTML = `<div class="alert ${alertClass} alert-dismissible fade show" role="alert">
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    </div>`;

    // 5秒後に自動で非表示
    setTimeout(() => {
        statusDiv.innerHTML = '';
    }, 5000);
}

// フォームのリセット
function resetForm() {
    document.getElementById('object-type').value = '';
    document.getElementById('record-name').value = '';
    document.getElementById('description').value = '';
    updateSendButtonState();
}

// リアルタイム入力チェック
document.addEventListener('DOMContentLoaded', () => {
    if (document.getElementById('record-name')) {
        document.getElementById('record-name').addEventListener('input', updateSendButtonState);
    }
});

// クロスウィンドウメッセージリスナー（認証完了通知を受信）
window.addEventListener('message', (event) => {
    console.log('Received message:', event.data);

    if (event.data && event.data.type === 'SALESFORCE_AUTH_SUCCESS') {
        console.log('Authentication success message received!');

        // トークンをローカルストレージに保存
        localStorage.setItem('salesforce_access_token', event.data.access_token);
        localStorage.setItem('salesforce_instance_url', event.data.instance_url);

        console.log('Tokens saved from message:', {
            access_token: event.data.access_token ? 'Present' : 'Missing',
            instance_url: event.data.instance_url
        });

        // 認証状態を即座に更新
        showAuthenticatedState();
    }
}, false);