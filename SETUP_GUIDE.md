# UniTrack — Setup Guide

## Overview

UniTrack stores data locally in **IndexedDB** (in-browser) and can optionally sync to
**Google Sheets** for cloud backup and cross-device access.  
The app is deployed as a static site on **GitHub Pages** — no server or monthly fees needed.

---

## Part 1 — Google Sheets Cloud Sync

### Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **new blank spreadsheet**.
2. Name it `UniTrack Data` (or anything you like).
3. Copy the **Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/  ← SHEET_ID_IS_HERE  /edit
   ```

### Step 2 — Create the Apps Script Web App

1. Inside the spreadsheet, click **Extensions → Apps Script**.
2. Delete the default `Code.gs` content.
3. Open the file `Code.gs` from this project and **paste the entire contents** into the editor.
4. Replace the two constants at the top:
   ```js
   const SHEET_ID = "paste-your-sheet-id-here";
   const API_KEY = "choose-any-secret-string"; // e.g. "MyFactory2024!"
   ```
5. Click **Save** (floppy disk icon).
6. Click **Deploy → New Deployment**.
7. Set:
   - **Type**: Web App
   - **Execute as**: Me (your Google account)
   - **Who has access**: Anyone
8. Click **Deploy** and authorize when prompted.
9. Copy the **Web App URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfyc.../exec
   ```

> **Every time you edit Code.gs**, you must create a NEW deployment to apply changes.

### Step 3 — Connect UniTrack to Google Sheets

1. Open UniTrack in your browser.
2. Go to **Settings → Cloud Sync — Google Sheets**.
3. Paste the **Web App URL** and your **API Key** (the same one from `Code.gs`).
4. Click **Save Config**, then **Test Connection**.
5. A green dot confirms the connection is working.

### How sync works

| Button                | What it does                                               |
| --------------------- | ---------------------------------------------------------- |
| **Test Connection**   | Verifies the URL and key are correct                       |
| **Push All to Cloud** | Overwrites Google Sheets with everything in your local app |
| **Pull from Cloud**   | Overwrites your local app with what's in Google Sheets     |

New scans and employee saves are **automatically pushed** to Google Sheets in the background.  
Use **Pull from Cloud** on a second device to load the latest data.

---

## Part 2 — GitHub Pages Deployment

### Prerequisites

- A free [GitHub account](https://github.com)
- Git installed on your machine (`git --version` to check)

### Step 1 — Initialize the local git repo

```bash
cd /path/to/uniform-tracker
git init
git add index.html app.js styles.css
git commit -m "Initial commit: UniTrack uniform management app"
```

> Do **not** commit `Code.gs` to a public repo — it contains your API key.  
> The `.gitignore` file in this project already excludes it.

### Step 2 — Create the GitHub repository

1. Go to [github.com/new](https://github.com/new).
2. Name the repo `uniform-tracker` (or anything you like).
3. Set it to **Public** (required for free GitHub Pages).
4. Do **not** initialize with a README or .gitignore (you already have one locally).
5. Click **Create repository**.

### Step 3 — Push to GitHub

Copy the commands GitHub shows you under "push an existing repository":

```bash
git remote add origin https://github.com/YOUR-USERNAME/uniform-tracker.git
git branch -M main
git push -u origin main
```

### Step 4 — Enable GitHub Pages

1. In your GitHub repo, click **Settings → Pages**.
2. Under **Source**, choose **Deploy from a branch**.
3. Set branch to `main`, folder to `/ (root)`.
4. Click **Save**.
5. After 1–2 minutes, your app is live at:
   ```
   https://YOUR-USERNAME.github.io/uniform-tracker/
   ```

### Step 5 — Every future update

```bash
git add .
git commit -m "Describe your change"
git push
```

GitHub Pages auto-deploys within ~1 minute.

---

## Security Notes

- Your API Key is stored in the browser's `localStorage` on each device — it is **not** in the source code that goes to GitHub.
- The Google Apps Script web app URL is publicly accessible, but requests without the correct API key are rejected.
- Do not commit `Code.gs` (with your secrets) to a public GitHub repository.
- For higher security, consider restricting the Apps Script to specific Google accounts.

---

## Troubleshooting

| Problem                              | Solution                                                                |
| ------------------------------------ | ----------------------------------------------------------------------- |
| "Unauthorized" error                 | API Key in the app doesn't match `Code.gs`                              |
| CORS error                           | Make sure you deployed as **Anyone** not "Anyone with a Google Account" |
| Changes to Code.gs not working       | You must create a **New Deployment** — editing and saving is not enough |
| App data gone after clearing browser | Use **Pull from Cloud** to restore from Google Sheets                   |
