// Configuration
const msalConfig = {
    auth: {
        // ▼▼▼ ここにAzure Portalで取得した「アプリケーション(クライアント)ID」を入力してください ▼▼▼
        clientId: "644497d5-b09a-4eb7-91c6-1c8c95d1d0b3",
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        // 組織アカウント(Business Standard)の場合は "https://login.microsoftonline.com/organizations/"
        // 個人アカウントも含めるなら "https://login.microsoftonline.com/common/"
        authority: "https://login.microsoftonline.com/organizations/",

        // ローカルサーバーのURLに合わせて変更してください。
        // Pythonのhttp.serverデフォルトポート(8000)に合わせています。
        // デプロイ後も動くように、現在のURLを自動設定します（クエリパラメータ除く）
        redirectUri: window.location.origin + window.location.pathname.replace(/\/$/, "").replace("/index.html", ""),
    },
    cache: {
        cacheLocation: "localStorage", // sessionStorageから変更（ブラウザを閉じても維持）
        storeAuthStateInCookie: true, // Safari/iOSでの対策としてCookieを使用
    }
};

const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite.All"]
};

// Excel Config
// パス: OneDrive root -> study-log.xlsx -> Table1
const EXCEL_FILE_PATH = "/study-log.xlsx";
const TABLE_NAME = "Table1";

// UI Elements
const loginSection = document.getElementById('login-section');
const actionSection = document.getElementById('action-section');
const statusSection = document.getElementById('status-section');
const signInButton = document.getElementById('signIn');
const logNowButton = document.getElementById('log-now');
const welcomeMsg = document.getElementById('welcome-msg');
const statusText = document.getElementById('status-text');
const spinner = document.getElementById('spinner');
const successIcon = document.getElementById('success-icon');
const consoleLog = document.getElementById('console-log');
const errorMsg = document.getElementById('error-msg');

let myMSALObj;
let username = "";

// Initialize MSAL
function initializeMsal() {
    try {
        myMSALObj = new msal.PublicClientApplication(msalConfig);

        // Handle redirect promise
        myMSALObj.handleRedirectPromise()
            .then(handleResponse)
            .catch(err => {
                console.error(err);
                showError("初期化エラー: " + err);
                alert("認証処理でエラー: " + err); // スマホ用に追加
            });
    } catch (e) {
        showError("MSAL初期化失敗: " + e);
        alert("初期化失敗: " + e); // スマホ用に追加
    }
}

function handleResponse(response) {
    if (response !== null) {
        username = response.account.username;
        showWelcomeMessage(username);
        checkAutoLog(); // 自動記録チェック
    } else {
        // Try to verify if we are already logged in
        const currentAccounts = myMSALObj.getAllAccounts();
        if (currentAccounts.length === 0) {
            showLogin();
        } else if (currentAccounts.length === 1) {
            username = currentAccounts[0].username;
            showWelcomeMessage(username);
            checkAutoLog(); // 自動記録チェック
        } else {
            // Multiple accounts - pick the first one
            username = currentAccounts[0].username;
            showWelcomeMessage(username);
            checkAutoLog(); // 自動記録チェック
        }
    }
}

// URLパラメータに ?auto=true があれば自動で記録ボタンを押す
function checkAutoLog() {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('auto') === 'true') {
        console.log("Auto log mode detected.");
        // 少し待ってから実行（画面描画待ち）
        setTimeout(() => {
            addRowToExcel();
        }, 500);
    }
}

function signIn() {
    myMSALObj.loginRedirect(loginRequest);
}

function getTokenRedirect(request) {
    request.account = myMSALObj.getAccountByUsername(username);
    return myMSALObj.acquireTokenSilent(request)
        .catch(error => {
            console.warn("silent token acquisition fails. acquiring token using redirect");
            if (error instanceof msal.InteractionRequiredAuthError) {
                // fallback to interaction when silent call fails
                return myMSALObj.acquireTokenRedirect(request);
            } else {
                console.error(error);
                showError("トークン取得エラー: " + error);
            }
        });
}

// Main Logic: Add Row to Excel
async function addRowToExcel() {
    showStatus("記録中...", true);

    try {
        const tokenResponse = await getTokenRedirect(loginRequest);
        if (!tokenResponse) return;

        const accessToken = tokenResponse.accessToken;

        // Prepare Data (Date, Time)
        const now = new Date();
        const dateStr = now.toLocaleDateString('ja-JP'); // YYYY/MM/DD
        // 時間を HH:mm 形式で取得 (秒は省略)
        const timeStr = now.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });

        const rowData = {
            values: [
                [dateStr, timeStr] // A列: 日付, B列: 時間
            ]
        };

        // Graph API Call
        const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/tables/${TABLE_NAME}/rows/add`;

        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(rowData)
        });

        if (response.ok) {
            const data = await response.json();
            log("Excelに行を追加しました: " + JSON.stringify(data));
            showSuccess("記録しました！");

            // 3秒後に元の画面に戻る
            setTimeout(() => {
                showAction();
            }, 3000);
        } else {
            const errorData = await response.json();
            console.error("Graph API Error:", errorData);
            if (errorData.error && errorData.error.code === "ItemNotFound") {
                showError("エラー: Excelファイルまたはテーブルが見つかりません。OneDrive直下に 'study-log.xlsx' があり、'Table1' が作成されているか確認してください。");
            } else {
                showError("書き込みエラー: " + (errorData.error ? errorData.error.message : response.statusText));
            }
            // エラー時は戻らない
        }

    } catch (error) {
        console.error(error);
        showError("通信エラー: " + error);
    }
}


// UI Transitions
function showLogin() {
    loginSection.classList.remove('hidden');
    actionSection.classList.add('hidden');
    statusSection.classList.add('hidden');
}

function showAction() {
    loginSection.classList.add('hidden');
    actionSection.classList.remove('hidden');
    statusSection.classList.add('hidden');
}

function showStatus(msg, isLoading) {
    loginSection.classList.add('hidden');
    actionSection.classList.add('hidden');
    statusSection.classList.remove('hidden');

    statusText.innerText = msg;
    if (isLoading) {
        spinner.classList.remove('hidden');
        successIcon.classList.add('hidden');
    } else {
        spinner.classList.add('hidden');
        successIcon.classList.add('hidden');
    }
}

function showSuccess(msg) {
    statusText.innerText = msg;
    spinner.classList.add('hidden');
    successIcon.classList.remove('hidden');
}

function showWelcomeMessage(name) {
    // メールアドレスを表示してアカウント確認を促す
    welcomeMsg.innerHTML = `ようこそ<br><span style="font-size: 0.8em; color: #cbd5e1;">${name}</span> さん`;
    showAction();
}

function showError(msg) {
    errorMsg.innerText = msg;
    errorMsg.classList.remove('hidden');
    showStatus("エラー発生", false);
}

function log(msg) {
    // consoleLog.classList.remove('hidden');
    const p = document.createElement('div');
    p.innerText = "[" + new Date().toLocaleTimeString() + "] " + msg;
    consoleLog.prepend(p);
}


// Event Listeners
signInButton.addEventListener('click', signIn);
logNowButton.addEventListener('click', addRowToExcel);
document.getElementById('debug-btn').addEventListener('click', debugConnection);
document.getElementById('logout-btn').addEventListener('click', signOut);

function signOut() {
    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(username),
        postLogoutRedirectUri: msalConfig.auth.redirectUri,
    };
    myMSALObj.logoutRedirect(logoutRequest);
}

// Debug Function
async function debugConnection() {
    showStatus("デバッグ中...", true);
    try {
        const tokenResponse = await getTokenRedirect(loginRequest);
        if (!tokenResponse) return;
        const accessToken = tokenResponse.accessToken;

        let msg = "【デバッグ結果: ルートフォルダ一覧】\n";

        // ルートフォルダの子供を全部取得してみる
        const rootUrl = "https://graph.microsoft.com/v1.0/me/drive/root/children";
        const rootRes = await fetch(rootUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });

        if (!rootRes.ok) {
            const err = await rootRes.json();
            showError("フォルダ取得エラー: " + err.error.message);
            return;
        }

        const rootData = await rootRes.json();
        const files = rootData.value;

        msg += `📁 ファイル数: ${files.length} 個\n\n`;

        if (files.length === 0) {
            msg += "⚠️ フォルダは空っぽです。（認証したアカウントのOneDriveが正しいか確認してください）\n";
        } else {
            const targetFile = files.find(f => f.name === "study-log.xlsx");

            if (targetFile) {
                msg += "✅ 'study-log.xlsx' を発見しました！\n";
                msg += `ID: ${targetFile.id}\n\n`;

                // テーブル確認
                msg += "--- テーブル確認 ---\n";
                const tablesUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${targetFile.id}/workbook/tables`;
                try {
                    const tablesRes = await fetch(tablesUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
                    if (tablesRes.ok) {
                        const tablesData = await tablesRes.json();
                        const tableNames = tablesData.value.map(t => t.name);
                        msg += `📊 テーブル: ${tableNames.join(", ") || "(なし)"}\n`;
                    } else {
                        const tErr = await tablesRes.json();
                        msg += `❌ テーブル取得エラー: ${tErr.error.code}\n`;
                    }
                } catch (e) {
                    msg += "❌ テーブル確認失敗\n";
                }

            } else {
                msg += "❌ 'study-log.xlsx' が見つかりません。\n\n";
                msg += "↓ 見えているファイル一覧 ↓\n";
                files.forEach(f => {
                    msg += `・${f.name} (${f.folder ? 'フォルダ' : 'ファイル'})\n`;
                });
            }
        }

        alert(msg);
        showAction();

    } catch (e) {
        showError("デバッグエラー: " + e);
        console.error(e);
    }
}

// Start
initializeMsal();
