# Prototype 2 — Mock Demo (sideload into Outlook)

This runs the SharePoint-filing add-in **inside Outlook** with sample data:
a realistic folder tree and fake AI recommendations. **No sign-in, no backend,
no Azure** — nothing you're waiting on (client secret / admin consent) is needed.
When those clear later, you fill in the API + AUTH values in `taskpane.html`
and it goes live with zero other changes.

Everything is already wired to `https://localhost:3000`.

---

## What you need (already installed)
- Node.js  ✓ (you have v24)
- That's it for the demo. (Python / func / azurite are for the *live* backend, not this.)

## Step 1 — Open PowerShell in this folder
Unzip this folder somewhere (e.g. `C:\demo_p2`), then:
```powershell
cd C:\demo_p2
```

## Step 2 — Create a trusted local HTTPS certificate
Outlook refuses plain http, even locally. This tool makes a trusted localhost cert:
```powershell
npx office-addin-dev-certs install
```
Say yes to any trust prompt. It writes a cert + key into your user profile and
prints their paths — copy those two paths, you need them in Step 3.
(Typical location: `C:\Users\<you>\.office-addin-dev-certs\`)

## Step 3 — Serve this folder over HTTPS on port 3000
Using the cert paths from Step 2:
```powershell
npx http-server . -S -C "C:\Users\<you>\.office-addin-dev-certs\localhost.crt" -K "C:\Users\<you>\.office-addin-dev-certs\localhost.key" -p 3000 --cors
```
Leave this window running. Test it: open a browser to
**https://localhost:3000/taskpane.html** — you should see the dark
"SharePoint · File this attachment" pane with three ranked recommendations
(92% / 74% / 41%) and a sample folder tree. If that renders, hosting works.

## Step 4 — Sideload the manifest into Outlook
**New Outlook / Outlook on the web:**
1. Open Outlook, open any email (with or without an attachment).
2. Top-right of the reading pane: **... (More actions) → Get Add-ins**
   (or Settings gear → **Manage add-ins / integrations**).
3. Choose **My add-ins → Add a custom add-in → Add from file**.
4. Pick `manifest.xml` from this folder → **Install**.

**Classic Outlook (desktop):** Home ribbon → **Get Add-ins** → **My add-ins**
→ **Custom add-ins → Add a custom add-in → Add from file** → `manifest.xml`.

## Step 5 — Open it
On an open email, find the **"Link to SharePoint"** button
(ribbon, or the ... More actions menu) → the pane opens on the right.

You'll see the mock recommendations and can browse the sample tree, pick a
folder, and "file" — it shows the success state. All simulated.

---

## Troubleshooting
- **Pane is blank / won't load:** Step 3 server isn't running, or the cert
  wasn't trusted. Re-open https://localhost:3000/taskpane.html directly — if the
  browser warns about the cert, the trust step didn't take; re-run Step 2.
- **"Add from file" missing:** your tenant may restrict sideloading. As a global-
  admin alternative, deploy via **M365 admin center → Integrated apps**, or use
  the Teams/Outlook **Developer Portal** to sideload.
- **Button doesn't appear:** give Outlook a minute after install, then reopen the
  email. New Outlook sometimes needs a refresh.
- **Nothing about sign-in appears:** correct — mock mode skips auth entirely.

## Going live later (for reference)
When the client secret + admin consent are ready, and the backend is running:
open `taskpane.html`, fill in the `API` URLs (your Function endpoints) and the
`AUTH` block (`clientId`, `tenantId`, `backendScope`). No other changes.
