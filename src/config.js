// Set clientId to an Entra app registration with the SPA redirect URI brk-multihub://localhost:3000 (see README.md).
export const config = {
    clientId: "00000000-0000-0000-0000-000000000000",
    authority: "https://login.microsoftonline.com/organizations",
    scopes: ["User.Read"],

    // The startup calls. Vite reloads the task pane when you save this file.
    settings: {
        useNaa: false,
        callGetAuthContext: true
    }
};
