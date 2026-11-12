/* global Office */
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { config } from "./config.js";

const timeoutMs = 10000;
const startTime = performance.now();

const logElement = document.getElementById("log");
let msalInstance = null;
const lines = [];

function log(message, detail) {
    const elapsed = `${Math.round(performance.now() - startTime)}ms`.padStart(8);
    const suffix = detail === undefined ? "" : ` ${formatDetail(detail)}`;
    const line = `[${elapsed}] ${message}${suffix}`;

    lines.push(line);
    console.log(`[PptWebAddinTest] ${line}`);
    logElement.textContent += `${line}\n`;
}

function dumpLog() {
    console.log(`[PptWebAddinTest] Full log\n${lines.join("\n")}`);
}

function formatDetail(detail) {
    if (detail instanceof Error) {
        return `${detail.name}: ${detail.errorCode ?? ""} ${detail.message}`;
    }

    return JSON.stringify(detail);
}

function withTimeout(promise, name) {
    const timeout = new Promise((_, reject) => {
        setTimeout(() => reject(new Error(`${name} did not settle in ${timeoutMs}ms`)), timeoutMs);
    });

    return Promise.race([promise, timeout]);
}

async function startNaaAsync() {
    if (config.clientId.startsWith("00000000")) {
        log("NAA skipped: set clientId in src/config.js");
        return;
    }

    log("createNestablePublicClientApplication: start");
    msalInstance = await createNestablePublicClientApplication({
        auth: { clientId: config.clientId, authority: config.authority }
    });
    log("createNestablePublicClientApplication: done");

    try {
        log("ssoSilent: start");
        const result = await withTimeout(msalInstance.ssoSilent({ scopes: config.scopes }), "ssoSilent");
        log("ssoSilent: success", result.account?.username);
    } catch (err) {
        log("ssoSilent: failed", err);
    }
}

async function callGetAuthContextAsync() {
    try {
        log("getAuthContext: start");
        const context = await withTimeout(Office.auth.getAuthContext(), "getAuthContext");
        log("getAuthContext: success", context);
    } catch (err) {
        log("getAuthContext: failed", err);
    }
}

async function signInAsync() {
    if (!msalInstance) {
        log("Sign in: turn on NAA and set clientId first");
        return;
    }

    try {
        log("acquireTokenPopup: start");
        const result = await msalInstance.acquireTokenPopup({ scopes: config.scopes });
        log("acquireTokenPopup: success", result.account?.username);
    } catch (err) {
        log("acquireTokenPopup: failed", err);
    }
}

async function startAsync() {
    const settings = config.settings;
    document.getElementById("signIn").addEventListener("click", () => void signInAsync());
    document.getElementById("clearLog").addEventListener("click", () => { logElement.textContent = ""; lines.length = 0; });

    log("Settings", settings);

    const info = await Office.onReady();
    log("Office.onReady", {
        host: info.host,
        platform: info.platform,
        nestedAppAuth11: Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")
    });

    if (settings.useNaa) {
        await startNaaAsync();
    }

    if (settings.callGetAuthContext) {
        await callGetAuthContextAsync();
    }

    log("Startup complete");
    dumpLog();
}

startAsync().catch((err) => {
    log("Startup failed", err);
    dumpLog();
});
