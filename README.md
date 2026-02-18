# Outlook Mail Scraper

A full-stack web dashboard to browse, search, scrape, and export emails from Outlook accounts via IMAP.

## Features

- **IMAP Authentication** — Connect with email + app password (no OAuth required)
- **Email Browsing** — Browse all folders with real-time search and pagination
- **Full Email View** — Subject, body (HTML/text), headers, metadata, recipients
- **Attachment Download** — View and download all attachments
- **Bulk Scrape** — Filter by date range, sender, subject, folder
- **Statistics** — Mailbox stats, folder breakdown, top senders
- **Export** — Download scraped data as JSON or CSV
- **Dark Theme** — Professional developer-focused UI

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌───────────────────┐
│  React Frontend │────▶│  FastAPI Backend  │────▶│  IMAP Server      │
│  (Vite + React) │     │  (Python 3.11+)  │     │  (IMAP4 over SSL) │
└─────────────────┘     └──────────────────┘     └───────────────────┘
```

## Prerequisites

- Python 3.11+
- Node.js 18+ (for frontend)
- An Outlook email account with an app password

### Supported Email Providers

- Outlook.com
- Hotmail
- Microsoft 365 (work/school accounts)

---

## Step 1: Generate an App Password

App passwords are required because IMAP doesn't support modern OAuth interactive login. You need to generate one from your Microsoft account.

1. Go to [Microsoft Account Security](https://account.microsoft.com/security)
2. Sign in with your Outlook/Microsoft account
3. Navigate to **Security** → **Advanced security options**
4. Under **App passwords**, click **Create a new app password**
5. Copy the generated password (format: `xxxx-xxxx-xxxx-xxxx`)

> **Note**: You may need to enable two-factor authentication (2FA) first before the app password option appears.

---

## Step 2: Backend Setup

```bash
cd backend

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Create .env file
cp .env.example .env
```

Edit `.env` with your credentials:
```env
OUTLOOK_EMAIL=you@outlook.com
OUTLOOK_APP_PASSWORD=xxxx-xxxx-xxxx-xxxx
IMAP_HOST=outlook.office365.com
IMAP_PORT=993
FRONTEND_URL=http://localhost:5173
```

Start the backend:
```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

The API docs are available at `http://localhost:8000/docs` (Swagger UI).

---

## Step 3: Frontend Setup

```bash
cd frontend

# Install dependencies
npm install

# Start dev server
npm run dev
```

Open `http://localhost:5173` in your browser.

---

## Step 4: First Login

1. Open the dashboard at `http://localhost:5173`
2. Enter your Outlook email address
3. Enter the app password you generated in Step 1
4. Click **Sign In**
5. The backend will test the IMAP connection and load your mailbox

---

## Docker

You can also run the full stack with Docker:

```bash
docker compose up
```

This starts both the backend (port 8000) and frontend (port 5173).

---

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/auth/login` | Authenticate with email + app password |
| `GET` | `/auth/status` | Check authentication status |
| `POST` | `/auth/logout` | Clear authentication |
| `GET` | `/api/folders` | List all mail folders |
| `GET` | `/api/emails` | List emails with filters & pagination |
| `GET` | `/api/emails/{id}` | Get full email detail with headers |
| `GET` | `/api/emails/{id}/attachments/{idx}` | Download attachment by index |
| `POST` | `/api/scrape` | Bulk scrape with filters |
| `POST` | `/api/export/json` | Export as JSON file |
| `POST` | `/api/export/csv` | Export as CSV file |
| `GET` | `/api/stats` | Mailbox statistics |
| `GET` | `/health` | Health check |

### Query Parameters for `/api/emails`

| Param | Type | Description |
|-------|------|-------------|
| `folder_id` | string | Filter by folder |
| `search` | string | Full-text search |
| `from_date` | string | YYYY-MM-DD start date |
| `to_date` | string | YYYY-MM-DD end date |
| `sender` | string | Filter by sender email |
| `importance` | string | low/normal/high |
| `has_attachments` | bool | Filter emails with attachments |
| `is_read` | bool | Filter read/unread |
| `skip` | int | Pagination offset |
| `top` | int | Page size (max 50) |

---

## Production Deployment Notes

1. **HTTPS**: Use nginx/Caddy with SSL certificates
2. **CORS**: Restrict `allow_origins` to your domain
3. **Credentials**: Stored in memory only — lost on server restart (by design)
4. **Secrets**: Use environment variables or a vault, never commit `.env`

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "IMAP authentication failed" | Verify your app password is correct and IMAP is enabled |
| "App password" option not visible | Enable two-factor authentication on your Microsoft account first |
| CORS errors | Check `FRONTEND_URL` in `.env` matches your dev server URL |
| Connection timeout | Check `IMAP_HOST` and `IMAP_PORT` — default is `outlook.office365.com:993` |
| "Not authenticated" errors | Session expired — sign in again (credentials are in-memory only) |

---

## Tech Stack

- **Backend**: Python, FastAPI, imaplib (IMAP4 over SSL)
- **Frontend**: React, Vite
- **Protocol**: IMAP4rev1 with STARTTLS/SSL
- **Auth**: Email + App Password (IMAP LOGIN)
