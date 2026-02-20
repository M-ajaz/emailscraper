# Connecting Your Outlook / Microsoft 365 Account

## Option A: OAuth2 Login (Recommended — No Admin Required)

This method uses Microsoft's official login flow. No app passwords or admin approval needed.

### Step 1: Register a Free Azure App (One-time setup, 5 minutes)

1. Go to https://portal.azure.com
2. Sign in with your Microsoft/work account
3. Search for "App registrations" in the top search bar
4. Click "New registration"
5. Fill in:
   - Name: `Local Mail Scraper`
   - Supported account types: **"Accounts in any organizational directory and personal Microsoft accounts"**
   - Redirect URI: Select **"Web"** and enter: `http://localhost:8000/auth/microsoft/callback`
6. Click **Register**
7. On the app overview page, copy the **"Application (client) ID"**

### Step 2: Configure the App

1. In the left menu click **"Authentication"**
2. Under "Advanced settings" set **"Allow public client flows"** to **YES**
3. Click **Save**

### Step 3: Add Client ID to the Tool

1. Open `backend/.env` in a text editor
2. Add this line:
   ```
   MICROSOFT_CLIENT_ID=paste-your-client-id-here
   ```
3. Save the file
4. Restart the backend

### Step 4: Sign In

1. Open the tool at http://localhost:5173
2. Click **"Sign in with Microsoft"**
3. Complete the Microsoft login in the browser
4. You will be redirected back to the tool automatically

---

## Option B: IMAP with App Password (Requires Admin Approval)

Use this if your IT admin enables IMAP access for your account.

1. Ask your IT admin to enable IMAP for your mailbox and provide an app password
2. IMAP Server: `imap.office365.com`, Port: `993`
3. Enter your email and app password in the IMAP login form

---

## Troubleshooting

- **"MICROSOFT_CLIENT_ID not configured"**: Add the client ID to `backend/.env` and restart the backend
- **"redirect_uri mismatch"**: Make sure the redirect URI in Azure exactly matches `http://localhost:8000/auth/microsoft/callback`
- **"Need admin approval"**: In Azure app registration, under Authentication, ensure "Allow public client flows" is set to Yes
