// Set clientId to an Entra app registration with the SPA redirect URI brk-multihub://localhost:3000 (see README.md).
export const config = {
    clientId: "3f693d4a-7477-412f-937a-8f4ad48d9c72",
    authority: "https://login.microsoftonline.com/organizations",
    scopes: ["User.Read"],

    // The startup calls. Vite reloads the task pane when you save this file.
    settings: {
        useNaa: false,
        callGetAuthContext: true
    }
};
