export const msalConfig = {
    auth: {
        // ▼▼▼ ここにAzure Portalで取得した「アプリケーション(クライアント)ID」を入力してください ▼▼▼
        clientId: "YOUR_CLIENT_ID_HERE", 
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
        
        // 組織アカウント(Business Standard)の場合は "https://login.microsoftonline.com/organizations/"
        // もし個人アカウントも含めるなら "https://login.microsoftonline.com/common/"
        authority: "https://login.microsoftonline.com/organizations/",
        
        // Viteのデフォルトポート。Azure Portalの「認証」>「リダイレクトURI」にもこれを追加する必要があります。
        redirectUri: "http://localhost:5173",
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    }
};

// Add scopes here for ID token to be used at Microsoft identity platform endpoints.
export const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite.All"]
};

// Add the endpoints here for Microsoft Graph API services you'd like to use.
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    // Excelへのパス: 自分のOneDriveのルートにある "study-log.xlsx" 内のテーブル "Table1" を想定
    excelAppendEndpoint: "https://graph.microsoft.com/v1.0/me/drive/root:/study-log.xlsx:/workbook/tables/Table1/rows/add"
};
