# Outlook Mail Scraper - Update Log

Reference document for all feature updates made to the project.

---

## Update 1: Independent & Combinable Filters
**Commit:** `30f45eb` | **Date:** Feb 2026

### Problem
- Scrape endpoint lacked keyword search
- Frontend didn't pass advanced filters (date range, sender) to the email listing API, even though the backend already supported them
- All Mail view had no filter UI beyond a search bar

### Backend Changes (`main.py`)
- Added `search: Optional[str] = None` field to `ScrapeRequest` model
- Threaded `search` parameter through `_scrape_impl` function to `_build_search_criteria`
- Passed `search=request.search` in the `/api/scrape` endpoint handler

### Frontend Changes (`App.jsx`)
- Added `inboxFilters` state (`from_date`, `to_date`, `sender`) and `showInboxFilters` toggle
- Updated `loadEmails` to accept a `filters` parameter and send `from_date`, `to_date`, `sender` as query params
- Updated all 7 `loadEmails` call sites to pass `inboxFilters` (debounced search, explicit search, folder select, pagination prev/next, refresh, All Mail nav)
- Added collapsible **Advanced Filters** panel in inbox view with From Date, To Date, Sender inputs + Apply/Clear buttons
- Added `search: ""` to `scrapeConfig` initial state
- Added "All filters are optional and work independently or combined" label in Scrape view
- Added full-width **Keyword Search** input to scrape filter grid
- Added **Clear Filters** button next to Export buttons

### Design Decisions
- No `has_attachments`/`importance` in inbox filters (IMAP can't search these natively; post-fetch filtering breaks pagination)
- Explicit "Apply" button for inbox filters (date pickers are awkward to debounce)
- Search bar keeps 300ms debounce and carries along current filters
- `subject_filter` stays scrape-only (inbox search bar via IMAP TEXT already covers subject searching)

---

## Update 2: Attachment Storage & Preview System
**Commit:** `328917e` | **Date:** Feb 2026

### Problem
- Scraped attachments were saved to disk but had no browsing/preview UI
- No way to view attachments outside of the email detail view
- No metadata tracking for stored files

### Backend Changes (`main.py`)
- Added `import mimetypes`
- Added constants: `_METADATA_FILE`, `IMAGE_TYPES`, `PDF_TYPES`, `DOC_TYPES`
- Added metadata helpers:
  - `_read_attachment_metadata()` / `_write_attachment_metadata()` — JSON sidecar pattern
  - `_save_attachment_with_metadata()` — saves file + records metadata (original name, content type, size, email context, timestamp)
  - `_classify_file_type()` — categorizes into image/pdf/document/other
- Updated `ScrapedAttachmentInfo` model with 3 new fields: `filename`, `download_url`, `preview_url`
- Added `StoredAttachmentInfo` model (filename, original_name, content_type, file_type, size, email context, URLs)
- Updated `_scrape_impl` to use `_save_attachment_with_metadata` and populate URL fields
- Added `_validate_attachment_filename()` — path traversal prevention
- Added 3 new endpoints:
  - `GET /api/attachments` — list all stored attachments with optional `file_type` filter
  - `GET /api/attachments/{filename}` — download a stored attachment
  - `GET /api/attachments/{filename}/preview` — inline preview for images/PDFs, metadata JSON for others

### Frontend Changes (`App.jsx`)
- Added state: `storedAttachments`, `attachmentFilter`, `previewModal`
- Added `loadAttachments()` function (fetches with optional type filter)
- Added `fileTypeIcon()` and `isPreviewable()` helper functions
- Added **Attachments** nav item in sidebar with count badge
- Added **Attachments View**:
  - Type filter tabs (All, Images, PDFs, Documents, Other)
  - Grid layout with cards showing thumbnails (images), icons (PDFs/docs), metadata
  - Download and Preview buttons per card
- Added **Preview Modal** (renders at top level):
  - Images: `<img>` tag with full-size preview
  - PDFs: `<iframe>` inline viewer
  - Others: file info card with download button
- Updated **Email Detail** attachments section:
  - Inline image thumbnails for image attachments
  - PDF indicator with "View PDF" label
  - Click to open preview modal for images/PDFs

### File Naming Pattern
- Stored as: `{uid}_{index}_{safe_filename}` in `./attachments/` directory
- Metadata stored in `./attachments/_metadata.json`

---

## Update 3: Improved Scrape Results Display
**Commit:** `e19802d` | **Date:** Feb 2026

### Problem
- Scrape results showed a basic table (subject, from, date, attachment count)
- No way to preview email content or attachment thumbnails from scrape results
- No summary statistics after scraping
- No bulk download for all scraped attachments
- CSV/JSON exports didn't include local attachment filenames or file paths

### Backend Changes (`main.py`)
- Added `import zipfile`
- Added `ZipDownloadRequest` model (accepts optional list of filenames)
- Added `POST /api/attachments/download-zip` endpoint:
  - Accepts specific filenames or downloads all stored attachments
  - Creates ZIP with original filenames (handles duplicates by prefixing UID)
  - Returns ZIP file with timestamped filename
- Updated `POST /api/export/csv`:
  - Added "Attachment Filenames" column (server-side filenames)
  - Added "Attachment Paths" column (local file paths)

### Frontend Changes (`App.jsx`)
- Added `expandedScrapeCards` state (Set tracking expanded indices, resets on new scrape)
- Added `downloadAllAttachments()` function (calls ZIP endpoint)
- Added `stripHtml()` helper (extracts plain text from HTML for body preview)
- Replaced scrape results table with:
  1. **Scrape Summary** — stat cards showing: total emails, total attachments, images count, PDFs count, documents count, other count
  2. **Action Row** — completion badge + timestamp + "Download All Attachments" ZIP button with file count
  3. **Expandable Email Cards**:
     - Card header: subject, sender, date, attachment count badge
     - Click to expand/collapse
     - Expanded body: first 500 characters of email body (HTML stripped to text)
     - Expanded attachments: clickable thumbnails with inline image previews, file type icons, download links
     - Clicking image/PDF thumbnails opens the preview modal

---

## API Endpoints Summary

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/auth/login` | IMAP login with email + app password |
| GET | `/auth/status` | Check authentication status |
| POST | `/auth/logout` | Logout and close IMAP connection |
| GET | `/api/folders` | List all IMAP folders |
| GET | `/api/emails` | List emails with filters and pagination |
| GET | `/api/emails/{id}` | Get full email detail |
| GET | `/api/emails/{id}/attachments/{idx}` | Download email attachment by index |
| GET | `/api/attachments` | List all stored attachments |
| GET | `/api/attachments/{filename}` | Download a stored attachment |
| GET | `/api/attachments/{filename}/preview` | Preview a stored attachment |
| POST | `/api/attachments/download-zip` | Download attachments as ZIP |
| POST | `/api/scrape` | Scrape emails with filters |
| POST | `/api/export/json` | Export scraped emails as JSON |
| POST | `/api/export/csv` | Export scraped emails as CSV |
| GET | `/api/stats` | Mailbox statistics |
| GET | `/health` | Health check |

## Tech Stack
- **Backend:** Python 3 / FastAPI / IMAP (imaplib)
- **Frontend:** React 18 / Vite 7
- **Auth:** IMAP with email + app password (no OAuth)
- **Repo:** https://github.com/M-ajaz/emailscraper.git
