# PowerPoint Web Add-in Test

A minimal PowerPoint task pane add-in that reproduces this dialog in PowerPoint on the web:

> **Microsoft PowerPoint** - Sorry, we ran into a problem. Please try again later. [Refresh]

The task pane keeps loading behind the dialog. Word and Excel on the web do not show the dialog with the same add-in code.

`src/config.js` has 2 settings, so you can run the add-in with and without auth calls at startup:

| Setting | What it does at startup |
|---|---|
| `useNaa` | `createNestablePublicClientApplication`, then `ssoSilent` (`@azure/msal-browser` 4.30.0). Needs a real `clientId`. |
| `callGetAuthContext` | `Office.auth.getAuthContext()` |

With both settings `false`, the add-in only calls `Office.onReady`. Vite reloads the task pane when you save `src/config.js`.

The add-in loads Office.js from `https://officeapis.public.onecdn.static.microsoft/1/office.js`, the current CDN address in the Microsoft documentation.

The task pane log shows each step with the time since the page loaded. The browser console shows the same lines with the prefix `[PptWebAddinTest]`.

## Prerequisites

- Node.js 22.2 or later.
- A Microsoft 365 work account with PowerPoint on the web.
- For NAA only: an Entra app registration (see below).

## 1. Make the app registration (NAA only)

1. In the Entra admin center, make an app registration for "Accounts in any organizational directory".
2. Under **Authentication**, add the platform **Single-page application**.
3. Add the redirect URI `brk-multihub://localhost:3000`.
4. Under **API permissions**, make sure that Microsoft Graph `User.Read` (delegated) is there.
5. Copy the application (client) ID into `clientId` in `src/config.js`.

If you leave `clientId` as the empty GUID, the add-in skips the NAA step and writes this in the log.

## 2. Sideload the add-in

The add-in uses the unified manifest for Microsoft 365 (`appPackage/manifest.json`). The Microsoft 365 Agents Toolkit CLI installs it to your account.

```sh
npm install
npm run signin
npm run sideload
```

- `signin`: sign in with the work account that you use for PowerPoint on the web.
- `sideload`: makes `appPackage.zip` and runs `atk install --file-path appPackage.zip`.

`office-addin-debugging start ... web` does not work with a unified manifest: its web sideload needs the `SourceLocation` from an XML manifest.

You can also upload `appPackage.zip` (from `npm run package`) in Teams: **Apps** > **Manage your apps** > **Upload an app** > **Upload a custom app**.

## 3. Start the add-in

```sh
npm start
```

The dev server runs on `https://localhost:3000`. On the first start, `office-addin-dev-certs` asks you to trust a development certificate.

Open a presentation in PowerPoint on the web, in a new browser tab. The **PPT Web Test** button is on the **Home** tab. If it is not there, select **Home** > **Add-ins** and find **PPT Web Test**. The add-in can take some minutes to show after the install.

## 4. Reproduce

1. Select **PPT Web Test** on the **Home** tab.
2. Wait about 1 minute.
3. Write down if the dialog shows, and the time on the task pane log when it showed.
4. Change the settings in `src/config.js` and do the test again. Use a new browser tab for each test.

Do these tests:

| Test | `useNaa` | `callGetAuthContext` |
|---|---|---|
| A, no auth calls | off | off |
| B | on | off |
| C | off | on |
| D | on | on |

## Issue text

```markdown
### Summary
PowerPoint on the web shows "Sorry, we ran into a problem. Please try again later." with a Refresh
button after a task pane add-in opens. The task pane continues to load behind the dialog.
Word and Excel on the web do not show the dialog with the same add-in code.

### Repro
Minimal add-in: <link to repo or zip>. Steps are in README.md, sections 1 to 4.

### Results
| Test | useNaa | callGetAuthContext | Dialog shows | Time |
|---|---|---|---|---|
| A | off | off | | |
| B | on | off | | |
| C | off | on | | |
| D | on | on | | |

### Expected
No error dialog.

### Actual
The PowerPoint error dialog shows. The console has this error from a frame that is not the add-in:
`taskpane.html?preload=1: Uncaught (in promise) TimeoutError: Timed out after 30000ms at taskpane.js`

### Environment
- Platform: PowerPoint on the web
- Browser and version:
- Manifest type: unified manifest for Microsoft 365 (JSON)
- Office.onReady result (from the task pane log):
- Account type: Microsoft 365 work account
```
