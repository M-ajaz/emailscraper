import React, { useState, useEffect, useCallback, useRef, memo } from "react";

const API = (typeof import.meta !== "undefined" && import.meta.env?.VITE_API_URL) || "http://localhost:8000";

// ─── Theme & Design Tokens ──────────────────────────────────────────────────
const theme = {
  bg: "#0a0e17",
  surface: "#111827",
  surfaceHover: "#1a2236",
  border: "#1e293b",
  borderActive: "#3b82f6",
  text: "#e2e8f0",
  textMuted: "#64748b",
  textDim: "#475569",
  accent: "#3b82f6",
  accentGlow: "rgba(59,130,246,0.15)",
  success: "#10b981",
  warning: "#f59e0b",
  danger: "#ef4444",
  purple: "#8b5cf6",
};

// ─── Debounce Hook ──────────────────────────────────────────────────────────
function useDebounce(value, delay = 300) {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    const timer = setTimeout(() => setDebounced(value), delay);
    return () => clearTimeout(timer);
  }, [value, delay]);
  return debounced;
}

// ─── Utility Components ─────────────────────────────────────────────────────
const Spinner = () => (
  <div style={{ display: "flex", justifyContent: "center", padding: 40 }}>
    <div
      style={{
        width: 32, height: 32, border: `3px solid ${theme.border}`,
        borderTopColor: theme.accent, borderRadius: "50%",
        animation: "spin 0.8s linear infinite",
      }}
    />
  </div>
);

const Badge = ({ children, color = theme.accent, bg }) => (
  <span
    style={{
      display: "inline-flex", alignItems: "center", padding: "2px 8px",
      borderRadius: 4, fontSize: 11, fontWeight: 600, letterSpacing: 0.5,
      color: color, backgroundColor: bg || `${color}20`,
      textTransform: "uppercase",
    }}
  >
    {children}
  </span>
);

const IconBtn = ({ children, onClick, active, title, style: s }) => (
  <button
    onClick={onClick}
    title={title}
    style={{
      background: active ? theme.accentGlow : "transparent",
      border: `1px solid ${active ? theme.accent : theme.border}`,
      color: active ? theme.accent : theme.textMuted,
      borderRadius: 6, padding: "6px 10px", cursor: "pointer",
      transition: "all 0.15s", fontSize: 13, ...s,
    }}
    onMouseEnter={(e) => {
      if (!active) e.target.style.borderColor = theme.textMuted;
    }}
    onMouseLeave={(e) => {
      if (!active) e.target.style.borderColor = theme.border;
    }}
  >
    {children}
  </button>
);

// ─── Safe HTML Email Renderer (sandboxed iframe) ────────────────────────────
const SafeEmailBody = ({ html, text }) => {
  const iframeRef = useRef(null);
  const [iframeHeight, setIframeHeight] = useState(400);

  useEffect(() => {
    if (!iframeRef.current) return;
    const checkHeight = () => {
      try {
        const doc = iframeRef.current?.contentDocument;
        if (doc?.body) {
          const h = doc.body.scrollHeight;
          if (h > 0) setIframeHeight(Math.min(h + 32, 2000));
        }
      } catch (e) {
        // sandbox may block access
      }
    };
    const timer = setInterval(checkHeight, 500);
    setTimeout(() => clearInterval(timer), 5000);
    return () => clearInterval(timer);
  }, [html, text]);

  if (html) {
    const wrappedHtml = `
      <!DOCTYPE html>
      <html><head>
        <meta charset="utf-8">
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                 font-size: 14px; line-height: 1.6; color: #1a1a1a; padding: 16px; margin: 0; }
          img { max-width: 100%; height: auto; }
          a { color: #3b82f6; }
          table { max-width: 100%; }
        </style>
      </head><body>${html}</body></html>`;

    return (
      <iframe
        ref={iframeRef}
        sandbox="allow-same-origin"
        srcDoc={wrappedHtml}
        style={{
          width: "100%", height: iframeHeight, border: "none",
          borderRadius: 8, background: "#fff",
        }}
        title="Email content"
      />
    );
  }

  return (
    <pre style={{
      whiteSpace: "pre-wrap", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      background: "#fff", borderRadius: 8, padding: 20, color: "#1a1a1a",
      border: `1px solid ${theme.border}`, fontSize: 14, lineHeight: 1.6,
    }}>
      {text}
    </pre>
  );
};

// ─── Memoized Email List Item ───────────────────────────────────────────────
const EmailListItem = memo(({ email, isSelected, onClick, formatDate }) => (
  <div
    onClick={() => onClick(email.id)}
    style={{
      padding: "12px 20px", borderBottom: `1px solid ${theme.border}`,
      cursor: "pointer",
      background: isSelected ? theme.accentGlow
        : !email.is_read ? `${theme.accent}08` : "transparent",
      transition: "background 0.15s",
    }}
    onMouseEnter={(e) => {
      if (!isSelected) e.currentTarget.style.background = theme.surfaceHover;
    }}
    onMouseLeave={(e) => {
      if (!isSelected) {
        e.currentTarget.style.background = !email.is_read ? `${theme.accent}08` : "transparent";
      }
    }}
  >
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
      {!email.is_read && (
        <div style={{
          width: 6, height: 6, borderRadius: "50%",
          background: theme.accent, flexShrink: 0,
        }} />
      )}
      <span style={{
        fontWeight: email.is_read ? 400 : 600, fontSize: 12,
        flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {email.sender}
      </span>
      <span style={{ fontSize: 10, color: theme.textDim, flexShrink: 0 }}>
        {formatDate(email.received)}
      </span>
    </div>
    <div style={{
      fontSize: 12, fontWeight: email.is_read ? 400 : 500,
      color: email.is_read ? theme.textMuted : theme.text,
      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      marginBottom: 4,
    }}>
      {email.subject || "(No Subject)"}
    </div>
    <div style={{
      fontSize: 11, color: theme.textDim, overflow: "hidden",
      textOverflow: "ellipsis", whiteSpace: "nowrap",
    }}>
      {email.preview}
    </div>
    <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
      {email.has_attachments && <Badge color={theme.purple}>📎 Attachment</Badge>}
      {email.importance === "high" && <Badge color={theme.danger}>❗ High</Badge>}
      {email.categories?.map((c, i) => <Badge key={i} color={theme.success}>{c}</Badge>)}
    </div>
  </div>
));

// ─── Main App ───────────────────────────────────────────────────────────────
export default function OutlookScraper() {
  const [auth, setAuth] = useState({ loading: true, authenticated: false, user: null });
  const [loginForm, setLoginForm] = useState({ email: "", password: "" });
  const [loginLoading, setLoginLoading] = useState(false);
  const [authMethod, setAuthMethod] = useState("imap"); // "imap" or "oauth2"
  const [msLoginLoading, setMsLoginLoading] = useState(false);
  const [view, setView] = useState("inbox"); // inbox | email | scrape | stats | attachments | candidates | candidateProfile | jobs | recruit-export | settings
  const [folders, setFolders] = useState([]);
  const [emails, setEmails] = useState([]);
  const [selectedEmail, setSelectedEmail] = useState(null);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [activeFolder, setActiveFolder] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const debouncedSearch = useDebounce(searchQuery, 300);
  const [pagination, setPagination] = useState({ skip: 0, top: 25, total: 0 });
  const [inboxFilters, setInboxFilters] = useState({ from_date: "", to_date: "", sender: "" });
  const [showInboxFilters, setShowInboxFilters] = useState(false);
  const [scrapeConfig, setScrapeConfig] = useState({
    folder_id: "", from_date: "", to_date: "",
    sender_filter: "", subject_filter: "", search: "",
    max_results: 50, include_attachments: true,
  });
  const [scrapeResult, setScrapeResult] = useState(null);
  const [storedAttachments, setStoredAttachments] = useState([]);
  const [attachmentFilter, setAttachmentFilter] = useState("all");
  const [previewModal, setPreviewModal] = useState(null);
  const [expandedScrapeCards, setExpandedScrapeCards] = useState(new Set());
  const [notification, setNotification] = useState(null);

  // ─── Recruitment State ──────────────────────────────────────────────────
  const [candidates, setCandidates] = useState([]);
  const [candidateSearch, setCandidateSearch] = useState("");
  const debouncedCandSearch = useDebounce(candidateSearch, 300);
  const [showDuplicatesOnly, setShowDuplicatesOnly] = useState(false);
  const [jobs, setJobs] = useState([]);
  const [selectedJobId, setSelectedJobId] = useState(null);
  const [matchResults, setMatchResults] = useState([]);
  const [matchLoading, setMatchLoading] = useState(false);
  const [jobForm, setJobForm] = useState({
    title: "", jd_raw: "", required_skills: "",
    min_exp: 0, location: "", remote_ok: false,
  });
  const [exportJobId, setExportJobId] = useState("");
  const [jdUploading, setJdUploading] = useState(false);
  const [tagFilter, setTagFilter] = useState("");
  const [selectedCandidate, setSelectedCandidate] = useState(null);
  const [selectedCandidateId, setSelectedCandidateId] = useState(null);
  const [candidateProfile, setCandidateProfile] = useState(null);
  const [profileActiveTab, setProfileActiveTab] = useState("resume");
  const [profileLoading, setProfileLoading] = useState(false);
  const [profileNotesSaved, setProfileNotesSaved] = useState(false);
  const [expandedCandidateId, setExpandedCandidateId] = useState(null);
  const [expandedMatchIdx, setExpandedMatchIdx] = useState(null);
  const [candidateNotesDraft, setCandidateNotesDraft] = useState({});
  const [candidateTagInput, setCandidateTagInput] = useState({});
  const SUGGESTED_TAGS = ["Shortlisted", "Rejected", "On Hold", "Interview Scheduled", "Offer Sent"];
  const [dbStats, setDbStats] = useState({ candidates: 0, emails: 0 });
  const [emailsFromCache, setEmailsFromCache] = useState(false);
  const [storageHealth, setStorageHealth] = useState(null);
  const [storageLastUpdated, setStorageLastUpdated] = useState(null);
  const [clearConfirmText, setClearConfirmText] = useState("");
  const [showClearModal, setShowClearModal] = useState(false);
  const [schedulerConfig, setSchedulerConfig] = useState(null);
  const [schedulerLoading, setSchedulerLoading] = useState(false);
  const [schedulerForm, setSchedulerForm] = useState({ interval_minutes: 30, folder: "INBOX", subject_filter: "" });
  const [schedulerRunNow, setSchedulerRunNow] = useState(false);
  const [schedulerCountdown, setSchedulerCountdown] = useState(null);

  // ─── Notification System State ──────────────────────────────────────
  const [notifications, setNotifications] = useState([]);
  const [unreadCount, setUnreadCount] = useState(0);
  const [notifDropdownOpen, setNotifDropdownOpen] = useState(false);
  const [toasts, setToasts] = useState([]);
  const notifBellRef = useRef(null);
  const notifDropdownRef = useRef(null);
  const prevUnreadCountRef = useRef(0);

  const notify = (msg, type = "info") => {
    setNotification({ msg, type });
    setTimeout(() => setNotification(null), 4000);
  };

  // ─── Notification System Functions ──────────────────────────────────
  const loadNotificationCount = async () => {
    try {
      const res = await fetch(`${API}/api/notifications/count`);
      if (!res.ok) return;
      const data = await res.json();
      const newCount = data.unread_count ?? data.count ?? 0;
      setUnreadCount((prev) => {
        prevUnreadCountRef.current = prev;
        return newCount;
      });
      return newCount;
    } catch (_) { /* silent */ }
  };

  const loadNotifications = async () => {
    try {
      const res = await fetch(`${API}/api/notifications?limit=20`);
      if (!res.ok) return;
      const data = await res.json();
      setNotifications(Array.isArray(data) ? data : data.notifications || []);
    } catch (_) { /* silent */ }
  };

  const showToast = (message, type = "info") => {
    const id = Date.now() + Math.random();
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 5000);
  };

  // ─── Poll notification count every 60s ──────────────────────────────
  useEffect(() => {
    loadNotificationCount();
    const interval = setInterval(async () => {
      const newCount = await loadNotificationCount();
      if (newCount != null && newCount > prevUnreadCountRef.current) {
        // New notifications arrived — fetch and toast high_fit ones
        try {
          const res = await fetch(`${API}/api/notifications?limit=10`);
          if (res.ok) {
            const data = await res.json();
            const items = Array.isArray(data) ? data : data.notifications || [];
            items
              .filter((n) => !n.is_read && n.type === "new_high_fit")
              .slice(0, 3)
              .forEach((n) => showToast(n.message || n.title, "new_high_fit"));
          }
        } catch (_) { /* silent */ }
      }
    }, 60000);
    return () => clearInterval(interval);
  }, []);

  // ─── Load notifications when dropdown opens ─────────────────────────
  useEffect(() => {
    if (notifDropdownOpen) loadNotifications();
  }, [notifDropdownOpen]);

  // ─── Close dropdown on outside click ────────────────────────────────
  useEffect(() => {
    if (!notifDropdownOpen) return;
    const handleClick = (e) => {
      if (
        notifDropdownRef.current && !notifDropdownRef.current.contains(e.target) &&
        notifBellRef.current && !notifBellRef.current.contains(e.target)
      ) {
        setNotifDropdownOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, [notifDropdownOpen]);

  const formatRelativeTime = (iso) => {
    if (!iso) return "";
    const diff = Date.now() - new Date(iso).getTime();
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return "Just now";
    if (mins < 60) return `${mins}m ago`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return `${hrs}h ago`;
    const days = Math.floor(hrs / 24);
    if (days < 7) return `${days}d ago`;
    return `${Math.floor(days / 7)}w ago`;
  };

  // ─── Pre-load cached data from SQLite on mount ─────────────────────────
  useEffect(() => {
    (async () => {
      try {
        const [candRes, emailRes] = await Promise.all([
          fetch(`${API}/api/candidates`).catch(() => null),
          fetch(`${API}/api/emails?source=cache`).catch(() => null),
        ]);
        if (candRes && candRes.ok) {
          const data = await candRes.json();
          setCandidates(data);
          setDbStats((prev) => ({ ...prev, candidates: data.length }));
        }
        if (emailRes && emailRes.ok) {
          const data = await emailRes.json();
          if (data.emails?.length) {
            setEmails(data.emails);
            setPagination({ skip: 0, top: 25, total: data.total || 0 });
            setEmailsFromCache(true);
            setDbStats((prev) => ({ ...prev, emails: data.total || 0 }));
          }
        }
      } catch (_) { /* silent pre-load */ }
    })();
  }, []);

  // ─── Auth ───────────────────────────────────────────────────────────────
  useEffect(() => {
    // Check for OAuth2 redirect callback
    const params = new URLSearchParams(window.location.search);
    if (params.get("auth") === "success") {
      window.history.replaceState({}, "", window.location.pathname);
      (async () => {
        try {
          const msRes = await fetch(`${API}/auth/microsoft/status`);
          if (msRes.ok) {
            const msData = await msRes.json();
            if (msData.connected) {
              setAuthMethod("oauth2");
              const email = msData.email || params.get("email") || "";
              setAuth({
                loading: false, authenticated: true,
                user: { name: email.split("@")[0], email },
              });
              loadFolders();
              loadEmails();
              return;
            }
          }
        } catch (_) {}
        checkAuth();
      })();
    } else {
      checkAuth();
    }
  }, []);

  // ─── Auto-refresh storage health every 30s when on settings view ────────
  useEffect(() => {
    if (view !== "settings") return;
    loadStorageHealth();
    loadSchedulerConfig();
    const interval = setInterval(() => { loadStorageHealth(); loadSchedulerConfig(); }, 30000);
    return () => clearInterval(interval);
  }, [view]);

  // ─── Scheduler countdown timer (ticks every second) ───────────────────
  useEffect(() => {
    if (view !== "settings" || schedulerCountdown == null) return;
    const tick = setInterval(() => {
      setSchedulerCountdown((prev) => (prev != null && prev > 0) ? prev - 1 : 0);
    }, 1000);
    return () => clearInterval(tick);
  }, [view, schedulerCountdown != null]);

  // ─── Auto-search on debounced input ────────────────────────────────────
  useEffect(() => {
    if (auth.authenticated && debouncedSearch !== undefined) {
      loadEmails(activeFolder, 0, debouncedSearch, inboxFilters);
    }
  }, [debouncedSearch]);

  // ─── Auto-search candidates on debounced input ───────────────────────
  useEffect(() => {
    if (auth.authenticated && view === "candidates") {
      loadCandidates(debouncedCandSearch);
    }
  }, [debouncedCandSearch]);

  const checkAuth = async () => {
    try {
      const res = await fetch(`${API}/auth/status`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (data.authenticated) {
        setAuth({ loading: false, ...data });
        setAuthMethod("imap");
        loadFolders();
        loadEmails();
        return;
      }
      // Check Microsoft OAuth2 session
      const msRes = await fetch(`${API}/auth/microsoft/status`).catch(() => null);
      if (msRes && msRes.ok) {
        const msData = await msRes.json();
        if (msData.connected) {
          setAuthMethod("oauth2");
          setAuth({
            loading: false, authenticated: true,
            user: { name: msData.email.split("@")[0], email: msData.email },
          });
          loadFolders();
          loadEmails();
          return;
        }
      }
      setAuth({ loading: false, authenticated: false, user: null });
    } catch (e) {
      setAuth({ loading: false, authenticated: false, user: null });
      setError("Cannot connect to backend server");
    }
  };

  const login = async (e) => {
    e.preventDefault();
    setLoginLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API}/auth/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: loginForm.email, password: loginForm.password }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.detail || `HTTP ${res.status}`);
      }
      setAuthMethod("imap");
      await checkAuth();
    } catch (err) {
      setError(err.message || "Login failed. Check your credentials.");
      notify(err.message || "Login failed", "error");
    }
    setLoginLoading(false);
  };

  const handleMicrosoftLogin = async () => {
    setMsLoginLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API}/auth/microsoft/url`);
      const data = await res.json();
      if (data.error) {
        setError(data.error);
        setMsLoginLoading(false);
        return;
      }
      if (data.url) {
        window.location.href = data.url;
      }
    } catch (err) {
      setError("Failed to start Microsoft login");
      setMsLoginLoading(false);
    }
  };

  const logout = async () => {
    if (authMethod === "oauth2") {
      await fetch(`${API}/auth/microsoft/logout`, { method: "POST" });
    } else {
      await fetch(`${API}/auth/logout`, { method: "POST" });
    }
    setAuth({ loading: false, authenticated: false, user: null });
    setAuthMethod("imap");
    setEmails([]);
    setFolders([]);
    setSelectedEmail(null);
    setStats(null);
  };

  // ─── Data Loading ──────────────────────────────────────────────────────
  const loadFolders = async () => {
    try {
      const res = await fetch(`${API}/api/folders`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setFolders(await res.json());
    } catch (e) {
      console.error("Failed to load folders:", e);
    }
  };

  const loadEmails = async (folderId = null, skip = 0, search = "", filters = {}) => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams({ skip: String(skip), top: "25" });
      if (folderId) params.set("folder_id", folderId);
      if (search) params.set("search", search);
      if (filters.from_date) params.set("from_date", filters.from_date);
      if (filters.to_date) params.set("to_date", filters.to_date);
      if (filters.sender) params.set("sender", filters.sender);
      const res = await fetch(`${API}/api/emails?${params}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setEmails(data.emails || []);
      setPagination({ skip, top: 25, total: data.total || 0 });
      setEmailsFromCache(false);
    } catch (e) {
      // IMAP failed — try cache fallback if we don't already have cached emails
      if (!emailsFromCache) {
        try {
          const cacheParams = new URLSearchParams({ skip: String(skip), top: "25", source: "cache" });
          if (folderId) cacheParams.set("folder_id", folderId);
          if (search) cacheParams.set("search", search);
          if (filters.sender) cacheParams.set("sender", filters.sender);
          const cRes = await fetch(`${API}/api/emails?${cacheParams}`);
          if (cRes.ok) {
            const cData = await cRes.json();
            if (cData.emails?.length) {
              setEmails(cData.emails);
              setPagination({ skip, top: 25, total: cData.total || 0 });
              setEmailsFromCache(true);
            }
          }
        } catch (_) { /* cache fallback also failed */ }
      }
      notify("IMAP unavailable — showing cached emails", "error");
    }
    setLoading(false);
  };

  const loadEmail = async (id) => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API}/api/emails/${id}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setSelectedEmail(await res.json());
      setView("email");
    } catch (e) {
      notify("Failed to load email", "error");
    }
    setLoading(false);
  };

  const loadStats = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API}/api/stats`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setStats(await res.json());
    } catch (e) {
      notify("Failed to load stats", "error");
    }
    setLoading(false);
  };

  const loadAttachments = async (typeFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (typeFilter && typeFilter !== "all") params.set("file_type", typeFilter);
      const res = await fetch(`${API}/api/attachments?${params}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setStoredAttachments(await res.json());
    } catch (e) {
      notify("Failed to load attachments", "error");
    }
    setLoading(false);
  };

  const runScrape = async () => {
    setLoading(true);
    setScrapeResult(null);
    setExpandedScrapeCards(new Set());
    setError(null);
    try {
      const body = { ...scrapeConfig };
      Object.keys(body).forEach((k) => {
        if (body[k] === "" || body[k] === null) delete body[k];
      });
      const res = await fetch(`${API}/api/scrape`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setScrapeResult(data);
      notify(`Scraped ${data.total_scraped} emails`, "success");
    } catch (e) {
      notify("Scrape failed", "error");
    }
    setLoading(false);
  };

  const exportData = async (format) => {
    try {
      const body = { ...scrapeConfig };
      Object.keys(body).forEach((k) => {
        if (body[k] === "" || body[k] === null) delete body[k];
      });
      const res = await fetch(`${API}/api/export/${format}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `outlook_export.${format}`;
      a.click();
      URL.revokeObjectURL(url);
      notify(`Exported as ${format.toUpperCase()}`, "success");
    } catch (e) {
      notify("Export failed", "error");
    }
  };

  const downloadAllAttachments = async (filenames = null) => {
    try {
      const res = await fetch(`${API}/api/attachments/download-zip`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filenames }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "attachments.zip";
      a.click();
      URL.revokeObjectURL(url);
      notify("Downloading attachments ZIP", "success");
    } catch (e) {
      notify("Failed to download attachments", "error");
    }
  };

  const stripHtml = (html) => {
    const tmp = document.createElement("div");
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || "";
  };

  // ─── Folder Click ──────────────────────────────────────────────────────
  const selectFolder = (f) => {
    setActiveFolder(f.id);
    setView("inbox");
    setSelectedEmail(null);
    loadEmails(f.id, 0, searchQuery, inboxFilters);
  };

  // ─── Search (form submit for explicit search) ─────────────────────────
  const handleSearch = (e) => {
    e.preventDefault();
    loadEmails(activeFolder, 0, searchQuery, inboxFilters);
  };

  // ─── Recruitment Data Loading ─────────────────────────────────────────
  const loadCandidates = async (search = "", tagF = tagFilter) => {
    try {
      const params = new URLSearchParams();
      if (search) params.set("name", search);
      if (tagF) params.set("tag", tagF);
      const res = await fetch(`${API}/api/candidates?${params}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setCandidates(data);
      if (!search && !tagF) setDbStats((prev) => ({ ...prev, candidates: data.length }));
    } catch (e) {
      notify("Failed to load candidates", "error");
    }
  };

  const deleteCandidate = async (id) => {
    try {
      const res = await fetch(`${API}/api/candidates/${id}`, { method: "DELETE" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      notify("Candidate removed", "success");
      loadCandidates(candidateSearch);
    } catch (e) {
      notify("Failed to delete candidate", "error");
    }
  };

  const saveCandidateNotes = async (id, notes) => {
    try {
      const res = await fetch(`${API}/api/candidates/${id}/notes`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ notes }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setCandidates((prev) => prev.map((c) => c.id === id ? { ...c, notes } : c));
    } catch (e) {
      notify("Failed to save notes", "error");
    }
  };

  const saveCandidateTags = async (id, tags) => {
    try {
      const res = await fetch(`${API}/api/candidates/${id}/tags`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tags }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setCandidates((prev) => prev.map((c) => c.id === id ? { ...c, tags } : c));
    } catch (e) {
      notify("Failed to save tags", "error");
    }
  };

  const addTagToCandidate = (id, tag) => {
    const c = candidates.find((x) => x.id === id);
    if (!c) return;
    const current = c.tags || [];
    if (current.some((t) => t.toLowerCase() === tag.toLowerCase())) return;
    const updated = [...current, tag];
    saveCandidateTags(id, updated);
  };

  const removeTagFromCandidate = (id, tag) => {
    const c = candidates.find((x) => x.id === id);
    if (!c) return;
    const updated = (c.tags || []).filter((t) => t !== tag);
    saveCandidateTags(id, updated);
  };

  const loadStorageHealth = async () => {
    try {
      const res = await fetch(`${API}/api/storage/health`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setStorageHealth(await res.json());
      setStorageLastUpdated(new Date());
    } catch (e) {
      notify("Failed to load storage health", "error");
    }
  };

  const loadSchedulerConfig = async () => {
    try {
      const res = await fetch(`${API}/api/scheduler/status`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setSchedulerConfig(data);
      setSchedulerForm({
        interval_minutes: data.interval_minutes || 30,
        folder: data.folder || "INBOX",
        subject_filter: data.subject_filter || "",
      });
      if (data.time_until_next_run_seconds != null) {
        setSchedulerCountdown(data.time_until_next_run_seconds);
      } else {
        setSchedulerCountdown(null);
      }
    } catch (_) { /* silent */ }
  };

  const openCandidateProfile = async (id) => {
    setSelectedCandidateId(id);
    setCandidateProfile(null);
    setProfileActiveTab("resume");
    setExpandedMatchIdx(null);
    setProfileLoading(true);
    setView("candidateProfile");
    try {
      const res = await fetch(`${API}/api/candidates/${id}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setCandidateProfile(data);
      setSelectedCandidate(data);
    } catch (e) {
      notify(`Failed to load candidate profile: ${e.message}`, "error");
    } finally {
      setProfileLoading(false);
    }
  };

  const loadJobs = async () => {
    try {
      const res = await fetch(`${API}/api/jobs`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setJobs(await res.json());
    } catch (e) {
      notify("Failed to load jobs", "error");
    }
  };

  const createJob = async () => {
    try {
      const skills = jobForm.required_skills
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
      const res = await fetch(`${API}/api/jobs`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...jobForm,
          required_skills: skills,
          min_exp: parseFloat(jobForm.min_exp) || 0,
        }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      notify("Job created", "success");
      setJobForm({ title: "", jd_raw: "", required_skills: "", min_exp: 0, location: "", remote_ok: false });
      loadJobs();
    } catch (e) {
      notify("Failed to create job", "error");
    }
  };

  const uploadJD = async (file) => {
    setJdUploading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API}/api/jobs/upload-jd`, { method: "POST", body: fd });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || `HTTP ${res.status}`);
      }
      const data = await res.json();
      setJobForm((prev) => ({
        ...prev,
        jd_raw: data.raw_text || prev.jd_raw,
        title: data.title || prev.title,
        required_skills: (data.required_skills || []).join(", ") || prev.required_skills,
        min_exp: data.min_exp ?? prev.min_exp,
        location: data.location || prev.location,
        remote_ok: data.remote_ok ?? prev.remote_ok,
      }));
      notify("JD parsed — review and save", "success");
    } catch (e) {
      notify(`JD upload failed: ${e.message}`, "error");
    }
    setJdUploading(false);
  };

  const runJobMatch = async (jobId) => {
    setMatchLoading(true);
    try {
      const res = await fetch(`${API}/api/jobs/${jobId}/match`, { method: "POST" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setMatchResults(data.results || []);
      notify(`Matched ${data.total_candidates} candidates`, "success");
    } catch (e) {
      notify("Matching failed", "error");
    }
    setMatchLoading(false);
  };

  const loadMatchResults = async (jobId) => {
    try {
      const res = await fetch(`${API}/api/jobs/${jobId}/results`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setMatchResults(data.results || []);
    } catch (e) {
      setMatchResults([]);
    }
  };

  const exportRecruitCSV = async (jobId) => {
    try {
      const res = await fetch(`${API}/api/export/candidates-csv?job_id=${jobId}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `match_results_job_${jobId}.csv`;
      a.click();
      URL.revokeObjectURL(url);
      notify("CSV downloaded", "success");
    } catch (e) {
      notify("CSV export failed", "error");
    }
  };

  // ─── Helpers ──────────────────────────────────────────────────────────
  const folderIcon = (name) => {
    const n = name.toLowerCase();
    if (n.includes("inbox")) return "📥";
    if (n.includes("sent")) return "📤";
    if (n.includes("draft")) return "📝";
    if (n.includes("delete") || n.includes("trash")) return "🗑️";
    if (n.includes("junk") || n.includes("spam")) return "⚠️";
    if (n.includes("archive")) return "📦";
    return "📁";
  };

  const formatDate = useCallback((d) => {
    if (!d) return "";
    const dt = new Date(d);
    const now = new Date();
    const diff = now - dt;
    if (diff < 86400000) return dt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    if (diff < 604800000) return dt.toLocaleDateString([], { weekday: "short" });
    return dt.toLocaleDateString([], { month: "short", day: "numeric" });
  }, []);

  const formatSize = (b) => {
    if (b < 1024) return `${b} B`;
    if (b < 1048576) return `${(b / 1024).toFixed(1)} KB`;
    return `${(b / 1048576).toFixed(1)} MB`;
  };

  const fileTypeIcon = (ct) => {
    if (!ct) return "📄";
    if (ct.startsWith("image/")) return "🖼️";
    if (ct === "application/pdf") return "📕";
    if (ct.includes("word") || ct.includes("document")) return "📘";
    if (ct.includes("sheet") || ct.includes("excel")) return "📗";
    if (ct.includes("presentation") || ct.includes("powerpoint")) return "📙";
    if (ct.startsWith("text/")) return "📝";
    if (ct.startsWith("video/")) return "🎬";
    if (ct.startsWith("audio/")) return "🎵";
    if (ct.includes("zip") || ct.includes("rar") || ct.includes("tar")) return "📦";
    return "📄";
  };

  const isPreviewable = (ct) => {
    if (!ct) return false;
    return ct.startsWith("image/") || ct === "application/pdf";
  };

  // ─── Login Screen ─────────────────────────────────────────────────────
  if (auth.loading) return (
    <div style={{ background: theme.bg, color: theme.text, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <Spinner />
    </div>
  );

  if (!auth.authenticated) return (
    <div style={{
      background: theme.bg, color: theme.text, minHeight: "100vh",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'JetBrains Mono', 'SF Mono', monospace",
    }}>
      <div style={{
        textAlign: "center", padding: 48,
        background: theme.surface, borderRadius: 16,
        border: `1px solid ${theme.border}`, maxWidth: 440,
        boxShadow: "0 25px 60px rgba(0,0,0,0.5)",
      }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>📧</div>
        <h1 style={{ fontSize: 22, fontWeight: 700, marginBottom: 8, letterSpacing: -0.5 }}>
          Outlook Mail Scraper
        </h1>
        <p style={{ color: theme.textMuted, fontSize: 13, marginBottom: 32, lineHeight: 1.6 }}>
          Connect your Outlook account to browse, search, and export your emails with full metadata and attachments.
        </p>
        {error && (
          <div style={{
            background: `${theme.danger}15`, border: `1px solid ${theme.danger}40`,
            borderRadius: 8, padding: "10px 16px", marginBottom: 20,
            fontSize: 12, color: theme.danger,
          }}>
            {error}
          </div>
        )}
        <form onSubmit={login} style={{ textAlign: "left" }}>
          <label style={{ display: "block", fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>
            Email
          </label>
          <input
            type="email"
            required
            placeholder="you@outlook.com"
            value={loginForm.email}
            onChange={(e) => setLoginForm({ ...loginForm, email: e.target.value })}
            style={{
              width: "100%", padding: "10px 14px", borderRadius: 8, marginBottom: 16,
              background: theme.bg, border: `1px solid ${theme.border}`,
              color: theme.text, fontSize: 13, outline: "none", boxSizing: "border-box",
            }}
            onFocus={(e) => e.target.style.borderColor = theme.accent}
            onBlur={(e) => e.target.style.borderColor = theme.border}
          />
          <label style={{ display: "block", fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>
            App Password
          </label>
          <input
            type="password"
            required
            placeholder="xxxx-xxxx-xxxx-xxxx"
            value={loginForm.password}
            onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })}
            style={{
              width: "100%", padding: "10px 14px", borderRadius: 8, marginBottom: 24,
              background: theme.bg, border: `1px solid ${theme.border}`,
              color: theme.text, fontSize: 13, outline: "none", boxSizing: "border-box",
            }}
            onFocus={(e) => e.target.style.borderColor = theme.accent}
            onBlur={(e) => e.target.style.borderColor = theme.border}
          />
          <button
            type="submit"
            disabled={loginLoading}
            style={{
              width: "100%", background: theme.accent, color: "#fff", border: "none",
              padding: "12px 32px", borderRadius: 8, fontSize: 14,
              fontWeight: 600, cursor: loginLoading ? "not-allowed" : "pointer",
              letterSpacing: 0.3, transition: "all 0.2s",
              boxShadow: `0 4px 20px ${theme.accentGlow}`,
              opacity: loginLoading ? 0.6 : 1,
            }}
          >
            {loginLoading ? "Connecting..." : "Sign In"}
          </button>
        </form>

        {/* OR divider */}
        <div style={{
          display: "flex", alignItems: "center", gap: 12,
          margin: "24px 0",
        }}>
          <div style={{ flex: 1, height: 1, background: theme.border }} />
          <span style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, letterSpacing: 1 }}>OR</span>
          <div style={{ flex: 1, height: 1, background: theme.border }} />
        </div>

        {/* Microsoft OAuth2 Login */}
        <button
          onClick={handleMicrosoftLogin}
          disabled={msLoginLoading}
          style={{
            width: "100%", display: "flex", alignItems: "center", justifyContent: "center",
            gap: 10, padding: "11px 20px", borderRadius: 8, fontSize: 13, fontWeight: 600,
            background: "#f0f0f0", color: "#1a1a1a", border: "1px solid #ccc",
            cursor: msLoginLoading ? "not-allowed" : "pointer",
            opacity: msLoginLoading ? 0.7 : 1,
            transition: "all 0.2s",
          }}
          onMouseEnter={(e) => { if (!msLoginLoading) e.currentTarget.style.background = "#e0e0e0"; }}
          onMouseLeave={(e) => { e.currentTarget.style.background = "#f0f0f0"; }}
        >
          {msLoginLoading ? (
            <div style={{
              width: 16, height: 16, border: "2px solid #ccc",
              borderTopColor: "#333", borderRadius: "50%",
              animation: "spin 0.8s linear infinite",
            }} />
          ) : (
            <svg width="18" height="18" viewBox="0 0 21 21">
              <rect x="1" y="1" width="9" height="9" fill="#f25022" />
              <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
              <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
              <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
            </svg>
          )}
          {msLoginLoading ? "Redirecting..." : "Sign in with Microsoft"}
        </button>

        <div style={{ marginTop: 20, fontSize: 11, color: theme.textDim, lineHeight: 1.6 }}>
          Works with Outlook.com, Hotmail, and Microsoft 365 accounts
          <br />
          <a href="https://account.microsoft.com/security" target="_blank" rel="noopener noreferrer"
            style={{ color: theme.accent, textDecoration: "none" }}>
            Generate an app password
          </a>{" "}
          · Data stays on your server
        </div>
      </div>
    </div>
  );

  // ─── Main Dashboard ───────────────────────────────────────────────────
  return (
    <div style={{
      display: "flex", height: "100vh", overflow: "hidden",
      background: theme.bg, color: theme.text,
      fontFamily: "'JetBrains Mono', 'SF Mono', 'Fira Code', monospace",
      fontSize: 13,
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600;700&display=swap');
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes slideIn { from { opacity:0; transform:translateY(-8px); } to { opacity:1; transform:translateY(0); } }
        @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
        @keyframes toastIn { from { opacity:0; transform:translateX(40px); } to { opacity:1; transform:translateX(0); } }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${theme.border}; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: ${theme.textDim}; }
        input, select { font-family: inherit; }
      `}</style>

      {/* ─── Notification ─────────────────────────────────────────────── */}
      {notification && (
        <div style={{
          position: "fixed", top: 16, right: 16, zIndex: 1000,
          padding: "10px 20px", borderRadius: 8, fontSize: 12, fontWeight: 500,
          animation: "slideIn 0.3s ease",
          background: notification.type === "error" ? theme.danger
            : notification.type === "success" ? theme.success : theme.accent,
          color: "#fff", boxShadow: "0 8px 30px rgba(0,0,0,0.4)",
        }}>
          {notification.msg}
        </div>
      )}

      {/* ─── Notification Dropdown ────────────────────────────────────── */}
      {notifDropdownOpen && (
        <div
          ref={notifDropdownRef}
          style={{
            position: "fixed", top: 52, left: sidebarOpen ? 120 : 10,
            width: 360, maxHeight: 480, overflowY: "auto",
            background: theme.surface, border: `1px solid ${theme.border}`,
            borderRadius: 10, boxShadow: "0 12px 40px rgba(0,0,0,0.5)",
            zIndex: 2000, animation: "slideIn 0.2s ease",
            display: "flex", flexDirection: "column",
          }}
        >
          {/* Dropdown Header */}
          <div style={{
            padding: "12px 16px", borderBottom: `1px solid ${theme.border}`,
            display: "flex", justifyContent: "space-between", alignItems: "center",
          }}>
            <span style={{ fontWeight: 700, fontSize: 13, color: theme.text }}>Notifications</span>
            <button
              onClick={async () => {
                try {
                  await fetch(`${API}/api/notifications/read-all`, { method: "PATCH" });
                  setUnreadCount(0);
                  loadNotifications();
                } catch (_) {}
              }}
              style={{
                background: "none", border: "none", color: theme.accent,
                cursor: "pointer", fontSize: 11, fontWeight: 500,
              }}
            >
              Mark all read
            </button>
          </div>

          {/* Notification Items */}
          {notifications.length === 0 ? (
            <div style={{
              padding: "40px 16px", textAlign: "center",
              color: theme.textDim, fontSize: 12,
            }}>
              No notifications yet
            </div>
          ) : (
            notifications.map((n) => {
              const typeIcon = n.type === "new_high_fit" ? "🎯"
                : n.type === "new_candidate" ? "👤"
                : n.type === "scrape_complete" ? "📧" : "🔔";
              return (
                <div
                  key={n.id}
                  onClick={async () => {
                    if (!n.is_read) {
                      try {
                        await fetch(`${API}/api/notifications/${n.id}/read`, { method: "PATCH" });
                      } catch (_) {}
                      setNotifications((prev) =>
                        prev.map((x) => x.id === n.id ? { ...x, is_read: true } : x)
                      );
                      setUnreadCount((c) => Math.max(0, c - 1));
                    }
                    setNotifDropdownOpen(false);
                    if (n.type === "new_high_fit") setView("jobs");
                    else if (n.type === "new_candidate") setView("candidates");
                  }}
                  style={{
                    padding: "10px 16px", cursor: "pointer",
                    background: n.is_read ? "transparent" : `${theme.accent}08`,
                    borderBottom: `1px solid ${theme.border}`,
                    display: "flex", gap: 10, alignItems: "flex-start",
                    transition: "background 0.15s", position: "relative",
                  }}
                  onMouseEnter={(e) => { e.currentTarget.style.background = theme.surfaceHover; }}
                  onMouseLeave={(e) => { e.currentTarget.style.background = n.is_read ? "transparent" : `${theme.accent}08`; }}
                >
                  {!n.is_read && (
                    <div style={{
                      position: "absolute", left: 6, top: "50%", transform: "translateY(-50%)",
                      width: 6, height: 6, borderRadius: "50%", background: theme.accent,
                    }} />
                  )}
                  <span style={{ fontSize: 16, flexShrink: 0, marginTop: 2 }}>{typeIcon}</span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{
                      fontSize: 12, fontWeight: 600, color: theme.text,
                      whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                    }}>
                      {n.title}
                    </div>
                    <div style={{
                      fontSize: 11, color: theme.textMuted, marginTop: 2,
                      display: "-webkit-box", WebkitLineClamp: 2,
                      WebkitBoxOrient: "vertical", overflow: "hidden",
                      lineHeight: 1.4,
                    }}>
                      {n.message}
                    </div>
                    <div style={{
                      fontSize: 10, color: theme.textDim, marginTop: 4, textAlign: "right",
                    }}>
                      {formatRelativeTime(n.created_at)}
                    </div>
                  </div>
                </div>
              );
            })
          )}

          {/* Dropdown Footer */}
          <div style={{
            padding: "10px 16px", borderTop: `1px solid ${theme.border}`,
            textAlign: "center",
          }}>
            <button
              onClick={async () => {
                try {
                  await fetch(`${API}/api/notifications/clear?mode=read`, { method: "DELETE" });
                  loadNotifications();
                } catch (_) {}
              }}
              style={{
                background: "none", border: "none", color: theme.textDim,
                cursor: "pointer", fontSize: 11,
              }}
            >
              Clear read notifications
            </button>
          </div>
        </div>
      )}

      {/* ─── Sidebar ──────────────────────────────────────────────────── */}
      <aside style={{
        width: sidebarOpen ? 240 : 56, flexShrink: 0,
        background: theme.surface, borderRight: `1px solid ${theme.border}`,
        display: "flex", flexDirection: "column", transition: "width 0.2s",
        overflow: "hidden",
      }}>
        {/* Header */}
        <div style={{
          padding: sidebarOpen ? "16px 16px 12px" : "16px 8px 12px",
          borderBottom: `1px solid ${theme.border}`,
          display: "flex", alignItems: "center", gap: 10,
        }}>
          <button
            onClick={() => setSidebarOpen(!sidebarOpen)}
            style={{
              background: "none", border: "none", color: theme.textMuted,
              cursor: "pointer", fontSize: 18, padding: 4, flexShrink: 0,
            }}
          >
            {sidebarOpen ? "◀" : "▶"}
          </button>
          {sidebarOpen && (
            <>
              <span style={{ fontWeight: 700, fontSize: 14, letterSpacing: -0.5, whiteSpace: "nowrap", flex: 1 }}>
                📧 Mail Scraper
              </span>
              <button
                ref={notifBellRef}
                onClick={() => setNotifDropdownOpen((v) => !v)}
                style={{
                  background: "none", border: "none", cursor: "pointer",
                  position: "relative", fontSize: 16, padding: 4, flexShrink: 0,
                  color: theme.textMuted, lineHeight: 1,
                }}
              >
                🔔
                {unreadCount > 0 && (
                  <span style={{
                    position: "absolute", top: -2, right: -4,
                    background: theme.danger, color: "#fff",
                    fontSize: 9, fontWeight: 700, minWidth: 16, height: 16,
                    borderRadius: 8, display: "flex", alignItems: "center",
                    justifyContent: "center", padding: "0 3px",
                    lineHeight: 1, border: `2px solid ${theme.surface}`,
                  }}>
                    {unreadCount > 9 ? "9+" : unreadCount}
                  </span>
                )}
              </button>
            </>
          )}
          {!sidebarOpen && (
            <button
              ref={notifBellRef}
              onClick={() => setNotifDropdownOpen((v) => !v)}
              style={{
                background: "none", border: "none", cursor: "pointer",
                position: "relative", fontSize: 14, padding: 4, flexShrink: 0,
                color: theme.textMuted, lineHeight: 1,
              }}
            >
              🔔
              {unreadCount > 0 && (
                <span style={{
                  position: "absolute", top: -2, right: -4,
                  background: theme.danger, color: "#fff",
                  fontSize: 8, fontWeight: 700, minWidth: 14, height: 14,
                  borderRadius: 7, display: "flex", alignItems: "center",
                  justifyContent: "center", padding: "0 2px",
                  lineHeight: 1, border: `2px solid ${theme.surface}`,
                }}>
                  {unreadCount > 9 ? "9+" : unreadCount}
                </span>
              )}
            </button>
          )}
        </div>

        {/* User Info */}
        {sidebarOpen && auth.user && (
          <div style={{ padding: "12px 16px", borderBottom: `1px solid ${theme.border}` }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: theme.text, marginBottom: 2 }}>
              {auth.user.name}
            </div>
            <div style={{ fontSize: 10, color: theme.textDim, wordBreak: "break-all", marginBottom: 6 }}>
              {auth.user.email}
            </div>
            <span style={{
              display: "inline-flex", alignItems: "center", gap: 4,
              padding: "2px 7px", borderRadius: 4, fontSize: 9, fontWeight: 600,
              letterSpacing: 0.3, textTransform: "uppercase",
              background: authMethod === "oauth2" ? `${theme.accent}20` : `${theme.textDim}20`,
              color: authMethod === "oauth2" ? theme.accent : theme.textDim,
            }}>
              {authMethod === "oauth2" && (
                <svg width="10" height="10" viewBox="0 0 21 21" style={{ flexShrink: 0 }}>
                  <rect x="1" y="1" width="9" height="9" fill="#f25022" />
                  <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
                  <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
                  <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
                </svg>
              )}
              {authMethod === "oauth2" ? "Microsoft" : "IMAP"}
            </span>
          </div>
        )}

        {/* Nav */}
        <nav style={{ padding: "8px", flex: 1, overflowY: "auto" }}>
          {[
            { id: "inbox", icon: "📥", label: "All Mail", action: () => { setView("inbox"); setActiveFolder(null); setSelectedEmail(null); loadEmails(null, 0, searchQuery, inboxFilters); } },
            { id: "scrape", icon: "🔍", label: "Scrape & Export", action: () => { setView("scrape"); setSelectedEmail(null); } },
            { id: "attachments", icon: "📎", label: "Attachments", badge: storedAttachments.length || null, action: () => { setView("attachments"); setSelectedEmail(null); loadAttachments(attachmentFilter); } },
            { id: "stats", icon: "📊", label: "Statistics", action: () => { setView("stats"); setSelectedEmail(null); loadStats(); } },
            { id: "candidates", icon: "👤", label: "Candidates", action: () => { setView("candidates"); setSelectedEmail(null); setSelectedCandidateId(null); setCandidateProfile(null); loadCandidates(candidateSearch); } },
            { id: "jobs", icon: "💼", label: "Jobs & Match", action: () => { setView("jobs"); setSelectedEmail(null); loadJobs(); } },
            { id: "recruit-export", icon: "📤", label: "Recruit Export", action: () => { setView("recruit-export"); setSelectedEmail(null); loadJobs(); } },
            { id: "settings", icon: "⚙️", label: "Settings", action: () => { setView("settings"); setSelectedEmail(null); } },
          ].map((item) => (
            <button
              key={item.id}
              onClick={item.action}
              style={{
                display: "flex", alignItems: "center", gap: 10, width: "100%",
                padding: sidebarOpen ? "8px 12px" : "8px", borderRadius: 6,
                background: (view === item.id || (item.id === "candidates" && view === "candidateProfile")) && !activeFolder ? theme.accentGlow : "transparent",
                border: "none", color: (view === item.id || (item.id === "candidates" && view === "candidateProfile")) && !activeFolder ? theme.accent : theme.textMuted,
                cursor: "pointer", fontSize: 12, fontWeight: 500, textAlign: "left",
                transition: "all 0.15s", justifyContent: sidebarOpen ? "flex-start" : "center",
              }}
            >
              <span style={{ fontSize: 15, flexShrink: 0 }}>{item.icon}</span>
              {sidebarOpen && item.label}
              {sidebarOpen && item.badge > 0 && (
                <span style={{
                  background: theme.purple, color: "#fff", fontSize: 9,
                  padding: "1px 5px", borderRadius: 8, fontWeight: 600, marginLeft: "auto",
                }}>
                  {item.badge}
                </span>
              )}
            </button>
          ))}

          {sidebarOpen && (
            <>
              <div style={{
                fontSize: 10, fontWeight: 600, color: theme.textDim, padding: "16px 12px 6px",
                letterSpacing: 1, textTransform: "uppercase",
              }}>
                Folders
              </div>
              {folders.map((f) => (
                <button
                  key={f.id}
                  onClick={() => selectFolder(f)}
                  style={{
                    display: "flex", alignItems: "center", gap: 8, width: "100%",
                    padding: "6px 12px", borderRadius: 6,
                    background: activeFolder === f.id ? theme.accentGlow : "transparent",
                    border: "none", color: activeFolder === f.id ? theme.accent : theme.textMuted,
                    cursor: "pointer", fontSize: 12, textAlign: "left",
                    transition: "all 0.15s",
                  }}
                >
                  <span style={{ fontSize: 13 }}>{folderIcon(f.name)}</span>
                  <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {f.name}
                  </span>
                  {f.unread_count > 0 && (
                    <span style={{
                      background: theme.accent, color: "#fff", fontSize: 10,
                      padding: "1px 6px", borderRadius: 10, fontWeight: 600,
                    }}>
                      {f.unread_count}
                    </span>
                  )}
                </button>
              ))}
            </>
          )}
        </nav>

        {/* DB Stats */}
        {sidebarOpen && (dbStats.candidates > 0 || dbStats.emails > 0) && (
          <div style={{
            padding: "8px 16px", borderTop: `1px solid ${theme.border}`,
            fontSize: 10, color: theme.textDim, lineHeight: 1.5,
            display: "flex", alignItems: "center", gap: 6,
          }}>
            <span style={{ opacity: 0.6 }}>DB</span>
            <span>
              {dbStats.candidates > 0 && `${dbStats.candidates} candidate${dbStats.candidates !== 1 ? "s" : ""}`}
              {dbStats.candidates > 0 && dbStats.emails > 0 && " · "}
              {dbStats.emails > 0 && `${dbStats.emails} email${dbStats.emails !== 1 ? "s" : ""} stored`}
            </span>
          </div>
        )}

        {/* Logout */}
        <div style={{ padding: 8, borderTop: `1px solid ${theme.border}` }}>
          <button
            onClick={logout}
            style={{
              width: "100%", padding: "8px", borderRadius: 6,
              background: "transparent", border: `1px solid ${theme.border}`,
              color: theme.textMuted, cursor: "pointer", fontSize: 11,
            }}
          >
            {sidebarOpen ? "Sign Out" : "⏻"}
          </button>
        </div>
      </aside>

      {/* ─── Main Content ─────────────────────────────────────────────── */}
      <main style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

        {/* ─── Error Banner ───────────────────────────────────────────── */}
        {error && (
          <div style={{
            padding: "10px 20px", background: `${theme.danger}15`,
            borderBottom: `1px solid ${theme.danger}40`,
            fontSize: 12, color: theme.danger, display: "flex",
            justifyContent: "space-between", alignItems: "center",
          }}>
            <span>{error}</span>
            <button
              onClick={() => setError(null)}
              style={{ background: "none", border: "none", color: theme.danger, cursor: "pointer", fontSize: 14 }}
            >
              ✕
            </button>
          </div>
        )}

        {/* ─── Inbox View ─────────────────────────────────────────────── */}
        {(view === "inbox" || view === "email") && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
            {/* Search Bar */}
            <div style={{
              padding: "12px 20px", borderBottom: `1px solid ${theme.border}`,
              display: "flex", alignItems: "center", gap: 12,
            }}>
              <form onSubmit={handleSearch} style={{ flex: 1, display: "flex", gap: 8 }}>
                <input
                  type="text"
                  placeholder="Search emails... (auto-searches as you type)"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{
                    flex: 1, padding: "8px 14px", borderRadius: 6,
                    background: theme.bg, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12, outline: "none",
                  }}
                  onFocus={(e) => e.target.style.borderColor = theme.accent}
                  onBlur={(e) => e.target.style.borderColor = theme.border}
                />
                <IconBtn type="submit" onClick={handleSearch}>Search</IconBtn>
              </form>
              <IconBtn onClick={() => loadEmails(activeFolder, pagination.skip, searchQuery, inboxFilters)}>↻ Refresh</IconBtn>
            </div>

            {/* Advanced Filters */}
            <div style={{ borderBottom: `1px solid ${theme.border}` }}>
              <button
                onClick={() => setShowInboxFilters(!showInboxFilters)}
                style={{
                  display: "flex", alignItems: "center", gap: 8, width: "100%",
                  padding: "8px 20px", background: "transparent", border: "none",
                  color: theme.textMuted, cursor: "pointer", fontSize: 11, fontWeight: 500,
                }}
              >
                <span>{showInboxFilters ? "▾" : "▸"}</span>
                <span>Advanced Filters</span>
                {(inboxFilters.from_date || inboxFilters.to_date || inboxFilters.sender) && (
                  <Badge color={theme.accent}>Active</Badge>
                )}
              </button>
              {showInboxFilters && (
                <div style={{ padding: "0 20px 12px", display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      From Date
                    </label>
                    <input
                      type="date"
                      value={inboxFilters.from_date}
                      onChange={(e) => setInboxFilters({ ...inboxFilters, from_date: e.target.value })}
                      style={{
                        padding: "6px 10px", borderRadius: 6,
                        background: theme.bg, border: `1px solid ${theme.border}`,
                        color: theme.text, fontSize: 11,
                      }}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      To Date
                    </label>
                    <input
                      type="date"
                      value={inboxFilters.to_date}
                      onChange={(e) => setInboxFilters({ ...inboxFilters, to_date: e.target.value })}
                      style={{
                        padding: "6px 10px", borderRadius: 6,
                        background: theme.bg, border: `1px solid ${theme.border}`,
                        color: theme.text, fontSize: 11,
                      }}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      Sender
                    </label>
                    <input
                      type="text"
                      placeholder="sender@example.com"
                      value={inboxFilters.sender}
                      onChange={(e) => setInboxFilters({ ...inboxFilters, sender: e.target.value })}
                      style={{
                        padding: "6px 10px", borderRadius: 6, minWidth: 180,
                        background: theme.bg, border: `1px solid ${theme.border}`,
                        color: theme.text, fontSize: 11,
                      }}
                    />
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button
                      onClick={() => loadEmails(activeFolder, 0, searchQuery, inboxFilters)}
                      style={{
                        padding: "6px 14px", borderRadius: 6, border: "none",
                        background: theme.accent, color: "#fff", fontSize: 11,
                        fontWeight: 600, cursor: "pointer",
                      }}
                    >
                      Apply
                    </button>
                    <button
                      onClick={() => {
                        const cleared = { from_date: "", to_date: "", sender: "" };
                        setInboxFilters(cleared);
                        loadEmails(activeFolder, 0, searchQuery, cleared);
                      }}
                      style={{
                        padding: "6px 14px", borderRadius: 6,
                        background: "transparent", border: `1px solid ${theme.border}`,
                        color: theme.textMuted, fontSize: 11, cursor: "pointer",
                      }}
                    >
                      Clear
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Cache banner */}
            {emailsFromCache && emails.length > 0 && (
              <div style={{
                padding: "8px 20px", background: "#f59e0b18",
                borderBottom: `1px solid #f59e0b40`,
                display: "flex", alignItems: "center", gap: 8,
                fontSize: 12, color: "#f59e0b",
              }}>
                <span style={{ fontSize: 14 }}>&#9888;</span>
                Showing cached emails — IMAP reconnecting...
              </div>
            )}

            <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>
              {/* Email List */}
              <div style={{
                width: view === "email" ? 360 : "100%",
                borderRight: view === "email" ? `1px solid ${theme.border}` : "none",
                overflowY: "auto", transition: "width 0.2s",
              }}>
                {loading && !emails.length ? <Spinner /> : emails.length === 0 ? (
                  <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                    {searchQuery ? "No emails match your search." : "No emails found."}
                  </div>
                ) : emails.map((email) => (
                  <EmailListItem
                    key={email.id}
                    email={email}
                    isSelected={selectedEmail?.id === email.id}
                    onClick={loadEmail}
                    formatDate={formatDate}
                  />
                ))}

                {/* Pagination */}
                {emails.length > 0 && (
                  <div style={{
                    padding: "12px 20px", display: "flex", alignItems: "center",
                    justifyContent: "space-between", borderTop: `1px solid ${theme.border}`,
                  }}>
                    <span style={{ fontSize: 11, color: theme.textDim }}>
                      {pagination.skip + 1}–{Math.min(pagination.skip + pagination.top, pagination.total)} of {pagination.total}
                    </span>
                    <div style={{ display: "flex", gap: 6 }}>
                      <IconBtn
                        onClick={() => {
                          if (pagination.skip > 0) loadEmails(activeFolder, Math.max(0, pagination.skip - 25), searchQuery, inboxFilters);
                        }}
                        style={{ opacity: pagination.skip === 0 ? 0.3 : 1, pointerEvents: pagination.skip === 0 ? "none" : "auto" }}
                      >
                        ← Prev
                      </IconBtn>
                      <IconBtn
                        onClick={() => {
                          if (pagination.skip + 25 < pagination.total) loadEmails(activeFolder, pagination.skip + 25, searchQuery, inboxFilters);
                        }}
                        style={{ opacity: pagination.skip + 25 >= pagination.total ? 0.3 : 1, pointerEvents: pagination.skip + 25 >= pagination.total ? "none" : "auto" }}
                      >
                        Next →
                      </IconBtn>
                    </div>
                  </div>
                )}
              </div>

              {/* Email Detail Panel */}
              {view === "email" && selectedEmail && (
                <div style={{ flex: 1, overflowY: "auto", animation: "fadeIn 0.2s ease" }}>
                  <div style={{ padding: 24 }}>
                    {/* Close button */}
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 16 }}>
                      <IconBtn onClick={() => { setView("inbox"); setSelectedEmail(null); }}>← Back to list</IconBtn>
                    </div>

                    {/* Subject */}
                    <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 16, letterSpacing: -0.5, lineHeight: 1.3 }}>
                      {selectedEmail.subject}
                    </h2>

                    {/* Metadata */}
                    <div style={{
                      background: theme.surface, borderRadius: 8, padding: 16,
                      border: `1px solid ${theme.border}`, marginBottom: 20,
                    }}>
                      <div style={{ display: "grid", gridTemplateColumns: "80px 1fr", gap: "8px 12px", fontSize: 12 }}>
                        <span style={{ color: theme.textDim, fontWeight: 500 }}>From</span>
                        <span>{selectedEmail.sender.name} &lt;{selectedEmail.sender.email}&gt;</span>

                        <span style={{ color: theme.textDim, fontWeight: 500 }}>To</span>
                        <span>{selectedEmail.to_recipients.map((r) => `${r.name} <${r.email}>`).join(", ")}</span>

                        {selectedEmail.cc_recipients.length > 0 && (
                          <>
                            <span style={{ color: theme.textDim, fontWeight: 500 }}>CC</span>
                            <span>{selectedEmail.cc_recipients.map((r) => `${r.name} <${r.email}>`).join(", ")}</span>
                          </>
                        )}

                        <span style={{ color: theme.textDim, fontWeight: 500 }}>Date</span>
                        <span>{new Date(selectedEmail.received).toLocaleString()}</span>

                        <span style={{ color: theme.textDim, fontWeight: 500 }}>Msg ID</span>
                        <span style={{ wordBreak: "break-all", color: theme.textDim, fontSize: 10 }}>
                          {selectedEmail.internet_message_id}
                        </span>
                      </div>
                    </div>

                    {/* Attachments */}
                    {selectedEmail.attachments.length > 0 && (
                      <div style={{
                        background: theme.surface, borderRadius: 8, padding: 16,
                        border: `1px solid ${theme.border}`, marginBottom: 20,
                      }}>
                        <div style={{ fontSize: 11, fontWeight: 600, color: theme.textDim, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
                          📎 Attachments ({selectedEmail.attachments.length})
                        </div>
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                          {selectedEmail.attachments.map((att) => {
                            const ct = att.content_type || "";
                            const isImage = ct.startsWith("image/");
                            const isPdf = ct === "application/pdf";
                            const downloadHref = `${API}/api/emails/${selectedEmail.id}/attachments/${att.id}`;
                            return (
                              <div
                                key={att.id}
                                style={{
                                  borderRadius: 6, background: theme.bg,
                                  border: `1px solid ${theme.border}`,
                                  overflow: "hidden", cursor: isImage || isPdf ? "pointer" : "default",
                                  transition: "border-color 0.15s", width: isImage ? 160 : "auto",
                                }}
                                onClick={() => {
                                  if (isImage || isPdf) {
                                    setPreviewModal({
                                      filename: att.name,
                                      original_name: att.name,
                                      content_type: ct,
                                      size: att.size,
                                      email_subject: selectedEmail.subject,
                                      preview_url: `/api/emails/${selectedEmail.id}/attachments/${att.id}`,
                                      download_url: `/api/emails/${selectedEmail.id}/attachments/${att.id}`,
                                    });
                                  }
                                }}
                                onMouseEnter={(e) => e.currentTarget.style.borderColor = theme.accent}
                                onMouseLeave={(e) => e.currentTarget.style.borderColor = theme.border}
                              >
                                {/* Inline image preview */}
                                {isImage && (
                                  <div style={{
                                    width: "100%", height: 100, overflow: "hidden",
                                    display: "flex", alignItems: "center", justifyContent: "center",
                                    background: "#fff",
                                  }}>
                                    <img
                                      src={downloadHref}
                                      alt={att.name}
                                      style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }}
                                      loading="lazy"
                                    />
                                  </div>
                                )}
                                {/* PDF mini indicator */}
                                {isPdf && (
                                  <div style={{
                                    padding: "8px 12px", display: "flex", alignItems: "center", gap: 6,
                                    background: `${theme.danger}10`,
                                  }}>
                                    <span style={{ fontSize: 16 }}>📕</span>
                                    <span style={{ fontSize: 10, color: theme.danger, fontWeight: 600 }}>View PDF</span>
                                  </div>
                                )}
                                <div style={{ padding: "8px 12px", display: "flex", alignItems: "center", gap: 8 }}>
                                  {!isImage && !isPdf && <span>{fileTypeIcon(ct)}</span>}
                                  <div style={{ flex: 1, overflow: "hidden" }}>
                                    <div style={{ fontWeight: 500, fontSize: 11, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                      {att.name}
                                    </div>
                                    <div style={{ fontSize: 10, color: theme.textDim }}>{formatSize(att.size)}</div>
                                  </div>
                                  <a
                                    href={downloadHref}
                                    onClick={(e) => e.stopPropagation()}
                                    style={{ fontSize: 14, textDecoration: "none", color: theme.textMuted }}
                                    title="Download"
                                  >
                                    ⬇
                                  </a>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* Headers (collapsible) */}
                    {selectedEmail.headers.length > 0 && (
                      <details style={{ marginBottom: 20 }}>
                        <summary style={{
                          fontSize: 11, fontWeight: 600, color: theme.textDim,
                          cursor: "pointer", textTransform: "uppercase", letterSpacing: 1,
                          marginBottom: 8,
                        }}>
                          View Email Headers ({selectedEmail.headers.length})
                        </summary>
                        <div style={{
                          background: theme.surface, borderRadius: 8, padding: 12,
                          border: `1px solid ${theme.border}`, maxHeight: 300,
                          overflowY: "auto", fontSize: 10, lineHeight: 1.6,
                        }}>
                          {selectedEmail.headers.map((h, i) => (
                            <div key={i} style={{ marginBottom: 4 }}>
                              <span style={{ color: theme.accent, fontWeight: 600 }}>{h.name}:</span>{" "}
                              <span style={{ color: theme.textMuted, wordBreak: "break-all" }}>{h.value}</span>
                            </div>
                          ))}
                        </div>
                      </details>
                    )}

                    {/* Email Body — rendered safely in sandboxed iframe */}
                    <SafeEmailBody html={selectedEmail.body_html} text={selectedEmail.body_text} />
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ─── Scrape & Export View ───────────────────────────────────── */}
        {view === "scrape" && (
          <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 4, letterSpacing: -0.5 }}>
              🔍 Scrape & Export
            </h2>
            <p style={{ color: theme.textMuted, fontSize: 12, marginBottom: 8 }}>
              Configure filters to bulk-scrape emails with all metadata and attachments.
            </p>
            <p style={{ color: theme.textDim, fontSize: 11, marginBottom: 24, fontStyle: "italic" }}>
              All filters are optional and work independently or combined.
            </p>

            <div style={{
              display: "grid", gridTemplateColumns: "1fr 1fr",
              gap: 16, maxWidth: 700, marginBottom: 24,
            }}>
              {/* Folder */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Folder
                </label>
                <select
                  value={scrapeConfig.folder_id}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, folder_id: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                >
                  <option value="">All Folders</option>
                  {folders.map((f) => (
                    <option key={f.id} value={f.id}>{f.name} ({f.total_count})</option>
                  ))}
                </select>
              </div>

              {/* Max Results */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Max Results
                </label>
                <input
                  type="number"
                  min="1"
                  max="500"
                  value={scrapeConfig.max_results}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, max_results: Math.min(500, Math.max(1, parseInt(e.target.value) || 50)) })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>

              {/* From Date */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  From Date
                </label>
                <input
                  type="date"
                  value={scrapeConfig.from_date}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, from_date: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>

              {/* To Date */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  To Date
                </label>
                <input
                  type="date"
                  value={scrapeConfig.to_date}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, to_date: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>

              {/* Sender Filter */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Sender Email
                </label>
                <input
                  type="email"
                  placeholder="sender@example.com"
                  value={scrapeConfig.sender_filter}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, sender_filter: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>

              {/* Subject Filter */}
              <div>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Subject Contains
                </label>
                <input
                  type="text"
                  placeholder="keyword..."
                  value={scrapeConfig.subject_filter}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, subject_filter: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>

              {/* Keyword Search (full width) */}
              <div style={{ gridColumn: "1 / -1" }}>
                <label style={{ fontSize: 11, color: theme.textDim, fontWeight: 600, marginBottom: 6, display: "block", textTransform: "uppercase", letterSpacing: 0.5 }}>
                  Keyword Search
                </label>
                <input
                  type="text"
                  placeholder="Search in email body and headers..."
                  value={scrapeConfig.search}
                  onChange={(e) => setScrapeConfig({ ...scrapeConfig, search: e.target.value })}
                  style={{
                    width: "100%", padding: "8px 12px", borderRadius: 6,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                    color: theme.text, fontSize: 12,
                  }}
                />
              </div>
            </div>

            {/* Include Attachments Toggle */}
            <label style={{
              display: "flex", alignItems: "center", gap: 10,
              fontSize: 12, color: theme.textMuted, marginBottom: 24, cursor: "pointer",
            }}>
              <input
                type="checkbox"
                checked={scrapeConfig.include_attachments}
                onChange={(e) => setScrapeConfig({ ...scrapeConfig, include_attachments: e.target.checked })}
                style={{ accentColor: theme.accent }}
              />
              Download and save all attachments
            </label>

            {/* Action Buttons */}
            <div style={{ display: "flex", gap: 10, marginBottom: 32 }}>
              <button
                onClick={runScrape}
                disabled={loading}
                style={{
                  padding: "10px 24px", borderRadius: 8, border: "none",
                  background: theme.accent, color: "#fff", fontSize: 13,
                  fontWeight: 600, cursor: loading ? "not-allowed" : "pointer",
                  boxShadow: `0 4px 16px ${theme.accentGlow}`,
                  opacity: loading ? 0.6 : 1,
                }}
              >
                {loading ? "Scraping..." : "🚀 Start Scrape"}
              </button>
              <IconBtn onClick={() => exportData("json")}>Export JSON</IconBtn>
              <IconBtn onClick={() => exportData("csv")}>Export CSV</IconBtn>
              <IconBtn onClick={() => setScrapeConfig({
                folder_id: "", from_date: "", to_date: "",
                sender_filter: "", subject_filter: "", search: "",
                max_results: 50, include_attachments: true,
              })}>Clear Filters</IconBtn>
            </div>

            {/* Scrape Results */}
            {loading && <Spinner />}
            {scrapeResult && (() => {
              const allAtts = scrapeResult.emails.flatMap((e) => e.attachments || []);
              const totalAtts = allAtts.length;
              const imgCount = allAtts.filter((a) => a.content_type?.startsWith("image/")).length;
              const pdfCount = allAtts.filter((a) => a.content_type === "application/pdf").length;
              const docCount = allAtts.filter((a) =>
                a.content_type?.includes("word") || a.content_type?.includes("sheet") ||
                a.content_type?.includes("excel") || a.content_type?.includes("presentation") ||
                a.content_type?.includes("powerpoint") || a.content_type?.startsWith("text/")
              ).length;
              const otherCount = totalAtts - imgCount - pdfCount - docCount;
              const savedFilenames = allAtts.filter((a) => a.filename).map((a) => a.filename);

              return (
                <div>
                  {/* Scrape Summary */}
                  <div style={{
                    display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
                    gap: 12, marginBottom: 20,
                  }}>
                    {[
                      { label: "Total Emails", value: scrapeResult.total_scraped, color: theme.accent },
                      { label: "Total Attachments", value: totalAtts, color: theme.purple },
                      { label: "Images", value: imgCount, color: theme.success },
                      { label: "PDFs", value: pdfCount, color: theme.danger },
                      { label: "Documents", value: docCount, color: theme.warning },
                      ...(otherCount > 0 ? [{ label: "Other", value: otherCount, color: theme.textMuted }] : []),
                    ].map((s) => (
                      <div key={s.label} style={{
                        background: theme.surface, borderRadius: 8, padding: "14px 16px",
                        border: `1px solid ${theme.border}`,
                      }}>
                        <div style={{ fontSize: 9, color: theme.textDim, textTransform: "uppercase", letterSpacing: 1, fontWeight: 600, marginBottom: 6 }}>
                          {s.label}
                        </div>
                        <div style={{ fontSize: 22, fontWeight: 700, color: s.color, letterSpacing: -0.5 }}>
                          {s.value}
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Action row */}
                  <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
                    <Badge color={theme.success}>✓ Complete</Badge>
                    <span style={{ fontSize: 12, color: theme.textMuted }}>
                      Scraped {scrapeResult.total_scraped} emails at {new Date(scrapeResult.exported_at).toLocaleString()}
                    </span>
                    {savedFilenames.length > 0 && (
                      <button
                        onClick={() => downloadAllAttachments(savedFilenames)}
                        style={{
                          marginLeft: "auto", padding: "6px 14px", borderRadius: 6, border: "none",
                          background: theme.purple, color: "#fff", fontSize: 11,
                          fontWeight: 600, cursor: "pointer",
                        }}
                      >
                        📦 Download All Attachments ({savedFilenames.length})
                      </button>
                    )}
                  </div>

                  {/* Expandable Email Cards */}
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {scrapeResult.emails.map((em, i) => {
                      const isExpanded = expandedScrapeCards.has(i);
                      const bodyText = em.body_type === "html" ? stripHtml(em.body) : em.body;
                      const preview = bodyText ? bodyText.substring(0, 500) : "";
                      return (
                        <div key={i} style={{
                          background: theme.surface, borderRadius: 8,
                          border: `1px solid ${theme.border}`, overflow: "hidden",
                        }}>
                          {/* Card Header */}
                          <div
                            onClick={() => {
                              setExpandedScrapeCards((prev) => {
                                const next = new Set(prev);
                                if (next.has(i)) next.delete(i); else next.add(i);
                                return next;
                              });
                            }}
                            style={{
                              padding: "12px 16px", cursor: "pointer", display: "flex",
                              alignItems: "center", gap: 12, transition: "background 0.15s",
                            }}
                            onMouseEnter={(e) => e.currentTarget.style.background = theme.surfaceHover}
                            onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}
                          >
                            <span style={{ fontSize: 11, color: theme.textDim, flexShrink: 0 }}>
                              {isExpanded ? "▾" : "▸"}
                            </span>
                            <div style={{ flex: 1, overflow: "hidden" }}>
                              <div style={{
                                fontSize: 12, fontWeight: 600, overflow: "hidden",
                                textOverflow: "ellipsis", whiteSpace: "nowrap",
                              }}>
                                {em.subject || "(No Subject)"}
                              </div>
                              <div style={{ fontSize: 10, color: theme.textMuted, marginTop: 2 }}>
                                {em.sender_name} &lt;{em.sender_email}&gt;
                              </div>
                            </div>
                            <span style={{ fontSize: 10, color: theme.textDim, flexShrink: 0 }}>
                              {formatDate(em.received)}
                            </span>
                            {em.attachments?.length > 0 && (
                              <Badge color={theme.purple}>{em.attachments.length} 📎</Badge>
                            )}
                          </div>

                          {/* Expanded Content */}
                          {isExpanded && (
                            <div style={{
                              padding: "0 16px 16px", borderTop: `1px solid ${theme.border}`,
                              animation: "fadeIn 0.2s ease",
                            }}>
                              {/* Body Preview */}
                              {em.body && (
                                <div style={{ margin: "12px 0", borderRadius: 6, overflow: "hidden" }}>
                                  {em.body_type === "html" ? (
                                    <SafeEmailBody html={em.body} />
                                  ) : (
                                    <div style={{
                                      padding: "12px 14px", borderRadius: 6,
                                      background: theme.bg, fontSize: 11, color: theme.textMuted,
                                      lineHeight: 1.6, whiteSpace: "pre-wrap", wordBreak: "break-word",
                                    }}>
                                      {preview}{bodyText.length > 500 ? "..." : ""}
                                    </div>
                                  )}
                                </div>
                              )}

                              {/* Attachment Thumbnails */}
                              {em.attachments?.length > 0 && (
                                <div>
                                  <div style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>
                                    Attachments ({em.attachments.length})
                                  </div>
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                                    {em.attachments.map((att, ai) => {
                                      const ct = att.content_type || "";
                                      const isImage = ct.startsWith("image/");
                                      const canPreview = isPreviewable(ct);
                                      return (
                                        <div
                                          key={ai}
                                          style={{
                                            borderRadius: 6, background: theme.bg,
                                            border: `1px solid ${theme.border}`,
                                            overflow: "hidden", cursor: canPreview ? "pointer" : "default",
                                            transition: "border-color 0.15s", width: isImage ? 120 : "auto",
                                          }}
                                          onClick={() => {
                                            if (canPreview && att.filename) {
                                              setPreviewModal({
                                                filename: att.filename,
                                                original_name: att.name,
                                                content_type: ct,
                                                size: att.size,
                                                email_subject: em.subject,
                                                preview_url: att.preview_url,
                                                download_url: att.download_url,
                                              });
                                            }
                                          }}
                                          onMouseEnter={(e) => e.currentTarget.style.borderColor = theme.accent}
                                          onMouseLeave={(e) => e.currentTarget.style.borderColor = theme.border}
                                        >
                                          {isImage && att.preview_url && (
                                            <div style={{
                                              width: "100%", height: 80, overflow: "hidden",
                                              display: "flex", alignItems: "center", justifyContent: "center",
                                              background: "#fff",
                                            }}>
                                              <img
                                                src={`${API}${att.preview_url}`}
                                                alt={att.name}
                                                style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }}
                                                loading="lazy"
                                              />
                                            </div>
                                          )}
                                          <div style={{ padding: "6px 10px", display: "flex", alignItems: "center", gap: 6 }}>
                                            {!isImage && <span style={{ fontSize: 14 }}>{fileTypeIcon(ct)}</span>}
                                            <div style={{ flex: 1, overflow: "hidden" }}>
                                              <div style={{ fontSize: 10, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                                {att.name}
                                              </div>
                                              <div style={{ fontSize: 9, color: theme.textDim }}>{formatSize(att.size)}</div>
                                            </div>
                                            {att.download_url && (
                                              <a
                                                href={`${API}${att.download_url}`}
                                                onClick={(e) => e.stopPropagation()}
                                                style={{ fontSize: 12, textDecoration: "none", color: theme.textMuted }}
                                                title="Download"
                                              >
                                                ⬇
                                              </a>
                                            )}
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* ─── Stats View ─────────────────────────────────────────────── */}
        {view === "stats" && (
          <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 24, letterSpacing: -0.5 }}>
              📊 Mailbox Statistics
            </h2>

            {loading ? <Spinner /> : stats ? (
              <>
                {/* Stat Cards */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 24 }}>
                  {[
                    { label: "Total Emails", value: stats.total_emails?.toLocaleString(), color: theme.accent },
                    { label: "Unread", value: stats.total_unread?.toLocaleString(), color: theme.warning },
                    { label: "Last 7 Days", value: stats.emails_last_7_days?.toLocaleString(), color: theme.success },
                  ].map((s) => (
                    <div key={s.label} style={{
                      background: theme.surface, borderRadius: 10, padding: 20,
                      border: `1px solid ${theme.border}`,
                    }}>
                      <div style={{ fontSize: 10, color: theme.textDim, textTransform: "uppercase", letterSpacing: 1, fontWeight: 600, marginBottom: 8 }}>
                        {s.label}
                      </div>
                      <div style={{ fontSize: 28, fontWeight: 700, color: s.color, letterSpacing: -1 }}>
                        {s.value}
                      </div>
                    </div>
                  ))}
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                  {/* Folder Breakdown */}
                  <div style={{
                    background: theme.surface, borderRadius: 10, padding: 20,
                    border: `1px solid ${theme.border}`,
                  }}>
                    <div style={{ fontSize: 11, fontWeight: 600, color: theme.textDim, marginBottom: 16, textTransform: "uppercase", letterSpacing: 1 }}>
                      Folder Breakdown
                    </div>
                    {stats.folder_stats?.map((f) => (
                      <div key={f.name} style={{
                        display: "flex", justifyContent: "space-between", alignItems: "center",
                        padding: "8px 0", borderBottom: `1px solid ${theme.border}`,
                      }}>
                        <span style={{ fontSize: 12 }}>
                          {folderIcon(f.name)} {f.name}
                        </span>
                        <div style={{ display: "flex", gap: 12, fontSize: 11 }}>
                          <span style={{ color: theme.textMuted }}>{f.total.toLocaleString()}</span>
                          {f.unread > 0 && <Badge color={theme.warning}>{f.unread} unread</Badge>}
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Top Senders */}
                  <div style={{
                    background: theme.surface, borderRadius: 10, padding: 20,
                    border: `1px solid ${theme.border}`,
                  }}>
                    <div style={{ fontSize: 11, fontWeight: 600, color: theme.textDim, marginBottom: 16, textTransform: "uppercase", letterSpacing: 1 }}>
                      Top Senders (Recent)
                    </div>
                    {stats.top_senders?.map((s, i) => {
                      const maxCount = stats.top_senders[0]?.count || 1;
                      return (
                        <div key={s.email} style={{ marginBottom: 12 }}>
                          <div style={{
                            display: "flex", justifyContent: "space-between",
                            fontSize: 12, marginBottom: 4,
                          }}>
                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
                              {s.name}
                            </span>
                            <span style={{ color: theme.accent, fontWeight: 600, flexShrink: 0, marginLeft: 8 }}>
                              {s.count}
                            </span>
                          </div>
                          <div style={{
                            height: 4, borderRadius: 2, background: theme.bg,
                          }}>
                            <div style={{
                              height: "100%", borderRadius: 2,
                              width: `${(s.count / maxCount) * 100}%`,
                              background: `linear-gradient(90deg, ${theme.accent}, ${theme.purple})`,
                              transition: "width 0.5s ease",
                            }} />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </>
            ) : (
              <div style={{ padding: 40, textAlign: "center", color: theme.textDim }}>
                No statistics available. Try refreshing.
              </div>
            )}
          </div>
        )}
        {/* ─── Attachments View ─────────────────────────────────────── */}
        {view === "attachments" && (
          <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 4, letterSpacing: -0.5 }}>
              📎 Attachments
            </h2>
            <p style={{ color: theme.textMuted, fontSize: 12, marginBottom: 20 }}>
              Browse and preview all downloaded attachments from scraped emails.
            </p>

            {/* Type Filter Tabs */}
            <div style={{ display: "flex", gap: 6, marginBottom: 20, flexWrap: "wrap" }}>
              {[
                { id: "all", label: "All" },
                { id: "image", label: "Images" },
                { id: "pdf", label: "PDFs" },
                { id: "document", label: "Documents" },
                { id: "other", label: "Other" },
              ].map((f) => (
                <button
                  key={f.id}
                  onClick={() => { setAttachmentFilter(f.id); loadAttachments(f.id); }}
                  style={{
                    padding: "6px 14px", borderRadius: 6, fontSize: 11, fontWeight: 500,
                    cursor: "pointer", transition: "all 0.15s",
                    background: attachmentFilter === f.id ? theme.accent : "transparent",
                    color: attachmentFilter === f.id ? "#fff" : theme.textMuted,
                    border: `1px solid ${attachmentFilter === f.id ? theme.accent : theme.border}`,
                  }}
                >
                  {f.label}
                </button>
              ))}
              <span style={{ fontSize: 11, color: theme.textDim, alignSelf: "center", marginLeft: 8 }}>
                {storedAttachments.length} file{storedAttachments.length !== 1 ? "s" : ""}
              </span>
            </div>

            {loading ? <Spinner /> : storedAttachments.length === 0 ? (
              <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                No attachments found. Scrape emails with "Download and save all attachments" enabled.
              </div>
            ) : (
              <div style={{
                display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260, 1fr))",
                gap: 12,
              }}>
                {storedAttachments.map((att) => (
                  <div
                    key={att.filename}
                    style={{
                      background: theme.surface, borderRadius: 8, padding: 16,
                      border: `1px solid ${theme.border}`, cursor: "pointer",
                      transition: "border-color 0.15s",
                    }}
                    onClick={() => setPreviewModal(att)}
                    onMouseEnter={(e) => e.currentTarget.style.borderColor = theme.accent}
                    onMouseLeave={(e) => e.currentTarget.style.borderColor = theme.border}
                  >
                    {/* Image thumbnail */}
                    {att.file_type === "image" && (
                      <div style={{
                        width: "100%", height: 120, borderRadius: 6, marginBottom: 10,
                        background: theme.bg, overflow: "hidden",
                        display: "flex", alignItems: "center", justifyContent: "center",
                      }}>
                        <img
                          src={`${API}${att.preview_url}`}
                          alt={att.original_name}
                          style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }}
                          loading="lazy"
                        />
                      </div>
                    )}
                    {/* PDF indicator */}
                    {att.file_type === "pdf" && (
                      <div style={{
                        width: "100%", height: 60, borderRadius: 6, marginBottom: 10,
                        background: `${theme.danger}15`, display: "flex",
                        alignItems: "center", justifyContent: "center", gap: 8,
                      }}>
                        <span style={{ fontSize: 24 }}>📕</span>
                        <span style={{ fontSize: 11, color: theme.danger, fontWeight: 600 }}>PDF</span>
                      </div>
                    )}
                    {/* Non-image/pdf icon */}
                    {att.file_type !== "image" && att.file_type !== "pdf" && (
                      <div style={{
                        width: "100%", height: 60, borderRadius: 6, marginBottom: 10,
                        background: theme.bg, display: "flex",
                        alignItems: "center", justifyContent: "center",
                      }}>
                        <span style={{ fontSize: 28 }}>{fileTypeIcon(att.content_type)}</span>
                      </div>
                    )}
                    <div style={{
                      fontSize: 12, fontWeight: 500, overflow: "hidden",
                      textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 4,
                    }}>
                      {att.original_name}
                    </div>
                    <div style={{ fontSize: 10, color: theme.textDim, marginBottom: 6 }}>
                      {formatSize(att.size)} · {att.content_type.split("/").pop()}
                    </div>
                    {att.email_subject && (
                      <div style={{
                        fontSize: 10, color: theme.textMuted, overflow: "hidden",
                        textOverflow: "ellipsis", whiteSpace: "nowrap",
                      }}>
                        From: {att.email_subject}
                      </div>
                    )}
                    <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                      <a
                        href={`${API}${att.download_url}`}
                        onClick={(e) => e.stopPropagation()}
                        style={{
                          padding: "4px 10px", borderRadius: 4, fontSize: 10,
                          background: theme.bg, border: `1px solid ${theme.border}`,
                          color: theme.textMuted, textDecoration: "none",
                        }}
                      >
                        Download
                      </a>
                      {isPreviewable(att.content_type) && (
                        <button
                          onClick={(e) => { e.stopPropagation(); setPreviewModal(att); }}
                          style={{
                            padding: "4px 10px", borderRadius: 4, fontSize: 10,
                            background: `${theme.accent}20`, border: `1px solid ${theme.accent}40`,
                            color: theme.accent, cursor: "pointer",
                          }}
                        >
                          Preview
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {/* ─── Candidates View ─────────────────────────────────────── */}
        {view === "candidates" && (() => {
          const dupCount = candidates.filter((c) => c.is_duplicate || c.duplicate_group_id).length;
          const displayed = showDuplicatesOnly
            ? candidates.filter((c) => c.duplicate_group_id != null)
            : candidates;
          const allTags = [...new Set(candidates.flatMap((c) => c.tags || []))];
          const TAG_COLORS = {
            "shortlisted": "#22c55e", "rejected": "#ef4444", "on hold": "#f59e0b",
            "interview scheduled": "#3b82f6", "offer sent": "#8b5cf6",
          };
          const getTagColor = (tag) => TAG_COLORS[tag.toLowerCase()] || theme.accent;
          return (
          <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 4, letterSpacing: -0.5 }}>
              Candidates
            </h2>
            <p style={{ color: theme.textMuted, fontSize: 12, marginBottom: 20 }}>
              Browse all extracted candidates from scraped email attachments.
            </p>

            {/* Search + Filters */}
            <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 20, flexWrap: "wrap" }}>
              <input
                type="text"
                placeholder="Search by name or skill..."
                value={candidateSearch}
                onChange={(e) => setCandidateSearch(e.target.value)}
                style={{
                  flex: 1, minWidth: 200, maxWidth: 400, padding: "8px 14px", borderRadius: 6,
                  background: theme.surface, border: `1px solid ${theme.border}`,
                  color: theme.text, fontSize: 12, outline: "none",
                }}
                onFocus={(e) => e.target.style.borderColor = theme.accent}
                onBlur={(e) => e.target.style.borderColor = theme.border}
              />
              <select
                value={tagFilter}
                onChange={(e) => { setTagFilter(e.target.value); loadCandidates(candidateSearch, e.target.value); }}
                style={{
                  padding: "8px 12px", borderRadius: 6, fontSize: 11,
                  background: theme.surface, border: `1px solid ${tagFilter ? theme.accent : theme.border}`,
                  color: tagFilter ? theme.accent : theme.textMuted, outline: "none", cursor: "pointer",
                }}
              >
                <option value="">All Tags</option>
                {[...new Set([...SUGGESTED_TAGS, ...allTags])].map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
              <button
                onClick={() => setShowDuplicatesOnly(!showDuplicatesOnly)}
                style={{
                  padding: "8px 14px", borderRadius: 6, fontSize: 11, fontWeight: 500,
                  cursor: "pointer", transition: "all 0.15s",
                  background: showDuplicatesOnly ? `${theme.warning}20` : "transparent",
                  color: showDuplicatesOnly ? theme.warning : theme.textMuted,
                  border: `1px solid ${showDuplicatesOnly ? theme.warning : theme.border}`,
                }}
              >
                Show Duplicates Only{dupCount > 0 ? ` (${dupCount})` : ""}
              </button>
            </div>

            <span style={{ fontSize: 11, color: theme.textDim, marginBottom: 12, display: "block" }}>
              {displayed.length} candidate{displayed.length !== 1 ? "s" : ""}
              {showDuplicatesOnly ? " (filtered)" : ""}{tagFilter ? ` — tag: "${tagFilter}"` : ""}
            </span>

            {/* Table */}
            {displayed.length === 0 ? (
              <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                {showDuplicatesOnly
                  ? "No duplicate candidates found."
                  : tagFilter
                    ? `No candidates with tag "${tagFilter}".`
                    : (
                      <div>
                        <div style={{ marginBottom: 12 }}>No candidates yet — go to Scrape & Export to scan your inbox</div>
                        <button
                          onClick={() => { setView("scrape"); setSelectedEmail(null); }}
                          style={{
                            padding: "8px 20px", borderRadius: 8, fontSize: 12, fontWeight: 600,
                            background: theme.accent, color: "#fff", border: "none", cursor: "pointer",
                          }}
                        >
                          Go to Scrape & Export
                        </button>
                      </div>
                    )}
              </div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ borderBottom: `2px solid ${theme.border}` }}>
                      {["Name", "Email", "Phone", "Skills", "Tags", "Exp", "Location", ""].map((h) => (
                        <th key={h || "actions"} style={{
                          padding: "8px 12px", textAlign: "left", fontSize: 10,
                          color: theme.textDim, fontWeight: 600, textTransform: "uppercase",
                          letterSpacing: 0.5, whiteSpace: "nowrap",
                        }}>
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {displayed.map((c) => (
                      <React.Fragment key={c.id}>
                      <tr
                        onClick={() => setExpandedCandidateId(expandedCandidateId === c.id ? null : c.id)}
                        style={{
                          borderBottom: expandedCandidateId === c.id ? "none" : `1px solid ${theme.border}`,
                          borderLeft: c.duplicate_group_id != null
                            ? `3px solid ${c.is_duplicate ? theme.warning : theme.success}`
                            : "3px solid transparent",
                          cursor: "pointer", transition: "background 0.1s",
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = `${theme.accent}08`}
                        onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}
                      >
                        <td style={{ padding: "10px 12px", fontWeight: 500 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ fontSize: 10, color: theme.textDim, marginRight: 2 }}>
                              {expandedCandidateId === c.id ? "\u25BC" : "\u25B6"}
                            </span>
                            {c.name}
                            {c.is_duplicate && <Badge color={theme.warning}>DUPLICATE</Badge>}
                            {c.duplicate_group_id != null && !c.is_duplicate && (
                              <Badge color={theme.success}>ORIGINAL</Badge>
                            )}
                          </div>
                        </td>
                        <td style={{ padding: "10px 12px", color: theme.accent }}>{c.email || "\u2014"}</td>
                        <td style={{ padding: "10px 12px", color: theme.textMuted }}>{c.phone || "\u2014"}</td>
                        <td style={{ padding: "10px 12px", maxWidth: 200 }}>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                            {(c.skills || []).slice(0, 5).map((s, i) => (
                              <Badge key={i} color={theme.accent}>{s}</Badge>
                            ))}
                            {(c.skills || []).length > 5 && (
                              <Badge color={theme.textDim}>+{c.skills.length - 5}</Badge>
                            )}
                          </div>
                        </td>
                        <td style={{ padding: "10px 12px", maxWidth: 180 }}>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                            {(c.tags || []).slice(0, 3).map((t, i) => (
                              <span key={i} style={{
                                padding: "2px 8px", borderRadius: 12, fontSize: 10, fontWeight: 600,
                                background: `${getTagColor(t)}20`, color: getTagColor(t),
                                border: `1px solid ${getTagColor(t)}40`,
                              }}>{t}</span>
                            ))}
                            {(c.tags || []).length > 3 && (
                              <span style={{ fontSize: 10, color: theme.textDim }}>+{c.tags.length - 3}</span>
                            )}
                          </div>
                        </td>
                        <td style={{ padding: "10px 12px", color: theme.textMuted }}>
                          {c.years_exp != null ? `${c.years_exp}y` : "\u2014"}
                        </td>
                        <td style={{ padding: "10px 12px", color: theme.textMuted }}>{c.location || "\u2014"}</td>
                        <td style={{ padding: "10px 12px" }}>
                          {c.is_duplicate && (
                            <button
                              onClick={(e) => { e.stopPropagation(); deleteCandidate(c.id); }}
                              style={{
                                padding: "4px 10px", borderRadius: 4, fontSize: 10,
                                background: `${theme.danger}15`, border: `1px solid ${theme.danger}40`,
                                color: theme.danger, cursor: "pointer", fontWeight: 600,
                                whiteSpace: "nowrap",
                              }}
                            >
                              Merge (Remove)
                            </button>
                          )}
                        </td>
                      </tr>
                      {/* Expanded detail panel */}
                      {expandedCandidateId === c.id && (
                        <tr style={{
                          borderBottom: `1px solid ${theme.border}`,
                          borderLeft: c.duplicate_group_id != null
                            ? `3px solid ${c.is_duplicate ? theme.warning : theme.success}`
                            : "3px solid transparent",
                        }}>
                          <td colSpan={8} style={{ padding: "0 12px 16px 32px" }}>
                            <div style={{ display: "flex", gap: 24, flexWrap: "wrap" }}>
                              {/* Tags Section */}
                              <div style={{ flex: "1 1 320px", minWidth: 280 }}>
                                <label style={{
                                  fontSize: 10, color: theme.textDim, fontWeight: 600,
                                  textTransform: "uppercase", letterSpacing: 0.5,
                                  display: "block", marginBottom: 8,
                                }}>Tags</label>
                                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                                  {(c.tags || []).map((t, i) => (
                                    <span key={i} style={{
                                      display: "inline-flex", alignItems: "center", gap: 4,
                                      padding: "3px 10px", borderRadius: 12, fontSize: 11, fontWeight: 600,
                                      background: `${getTagColor(t)}20`, color: getTagColor(t),
                                      border: `1px solid ${getTagColor(t)}40`,
                                    }}>
                                      {t}
                                      <span
                                        onClick={(e) => { e.stopPropagation(); removeTagFromCandidate(c.id, t); }}
                                        style={{ cursor: "pointer", fontSize: 13, lineHeight: 1, marginLeft: 2, opacity: 0.7 }}
                                        title="Remove tag"
                                      >&times;</span>
                                    </span>
                                  ))}
                                </div>
                                {/* Tag input */}
                                <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                                  <input
                                    type="text"
                                    placeholder="Add tag..."
                                    value={candidateTagInput[c.id] || ""}
                                    onClick={(e) => e.stopPropagation()}
                                    onChange={(e) => setCandidateTagInput((p) => ({ ...p, [c.id]: e.target.value }))}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter" && (candidateTagInput[c.id] || "").trim()) {
                                        addTagToCandidate(c.id, candidateTagInput[c.id].trim());
                                        setCandidateTagInput((p) => ({ ...p, [c.id]: "" }));
                                      }
                                    }}
                                    style={{
                                      padding: "5px 10px", borderRadius: 6, fontSize: 11, width: 140,
                                      background: theme.bg, border: `1px solid ${theme.border}`,
                                      color: theme.text, outline: "none",
                                    }}
                                  />
                                  {/* Suggested tags */}
                                  {SUGGESTED_TAGS.filter((st) => !(c.tags || []).some((t) => t.toLowerCase() === st.toLowerCase())).map((st) => (
                                    <button
                                      key={st}
                                      onClick={(e) => { e.stopPropagation(); addTagToCandidate(c.id, st); }}
                                      style={{
                                        padding: "3px 8px", borderRadius: 10, fontSize: 10,
                                        background: "transparent", border: `1px dashed ${theme.border}`,
                                        color: theme.textDim, cursor: "pointer",
                                      }}
                                      title={`Add "${st}"`}
                                    >+ {st}</button>
                                  ))}
                                </div>
                              </div>
                              {/* Notes Section */}
                              <div style={{ flex: "1 1 320px", minWidth: 280 }}>
                                <label style={{
                                  fontSize: 10, color: theme.textDim, fontWeight: 600,
                                  textTransform: "uppercase", letterSpacing: 0.5,
                                  display: "block", marginBottom: 8,
                                }}>Notes</label>
                                <textarea
                                  placeholder="Add recruiter notes..."
                                  value={candidateNotesDraft[c.id] !== undefined ? candidateNotesDraft[c.id] : (c.notes || "")}
                                  onClick={(e) => e.stopPropagation()}
                                  onChange={(e) => setCandidateNotesDraft((p) => ({ ...p, [c.id]: e.target.value }))}
                                  onBlur={() => {
                                    const draft = candidateNotesDraft[c.id];
                                    if (draft !== undefined && draft !== (c.notes || "")) {
                                      saveCandidateNotes(c.id, draft);
                                    }
                                    setCandidateNotesDraft((p) => { const n = { ...p }; delete n[c.id]; return n; });
                                  }}
                                  rows={3}
                                  style={{
                                    width: "100%", padding: "8px 12px", borderRadius: 6, fontSize: 12,
                                    background: theme.bg, border: `1px solid ${theme.border}`,
                                    color: theme.text, outline: "none", resize: "vertical",
                                    fontFamily: "inherit",
                                  }}
                                />
                                <div style={{ fontSize: 10, color: theme.textDim, marginTop: 4 }}>
                                  Auto-saves on blur
                                </div>
                              </div>
                            </div>
                            {/* Extra info + actions row */}
                            <div style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                              <div style={{ fontSize: 11, color: theme.textDim }}>
                                Source UID: {c.source_email_uid || "\u2014"} &nbsp;|&nbsp;
                                Titles: {(c.titles || []).join(", ") || "\u2014"} &nbsp;|&nbsp;
                                Added: {c.created_at ? new Date(c.created_at).toLocaleDateString() : "\u2014"}
                              </div>
                              <button
                                onClick={(e) => { e.stopPropagation(); openCandidateProfile(c.id); }}
                                style={{
                                  padding: "6px 16px", borderRadius: 6, fontSize: 11, fontWeight: 600,
                                  background: theme.accent, color: "#fff", border: "none",
                                  cursor: "pointer",
                                }}
                              >
                                View Full Profile
                              </button>
                            </div>
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
          );
        })()}

        {/* ─── Candidate Profile View ─────────────────────────────── */}
        {view === "candidateProfile" && selectedCandidateId && (() => {
          const p = candidateProfile;
          // Use list data as fallback while full profile loads
          const preview = p || candidates.find((x) => x.id === selectedCandidateId) || {};
          const c = p || preview;
          const TAG_COLORS = {
            "shortlisted": "#22c55e", "rejected": "#ef4444", "on hold": "#f59e0b",
            "interview scheduled": "#3b82f6", "offer sent": "#8b5cf6",
          };
          const getTagColor = (tag) => TAG_COLORS[tag.toLowerCase()] || theme.accent;
          const bestMatch = (p?.match_history || []).length > 0 ? p.match_history[0] : null;
          const fitColors = { high: "#22c55e", medium: "#f59e0b", low: "#ef4444" };
          const copyToClipboard = (text) => { navigator.clipboard.writeText(text); notify("Copied to clipboard", "success"); };

          return (
          <div style={{ flex: 1, overflowY: "auto", padding: "24px 32px" }}>
            {/* Top Header Bar */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24 }}>
              <button
                onClick={() => { setView("candidates"); setSelectedCandidateId(null); setCandidateProfile(null); setSelectedCandidate(null); }}
                style={{
                  padding: "6px 14px", borderRadius: 6, fontSize: 12, fontWeight: 500,
                  background: "transparent", border: `1px solid ${theme.border}`,
                  color: theme.textMuted, cursor: "pointer",
                  display: "flex", alignItems: "center", gap: 6,
                }}
              >
                ← Back
              </button>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flex: 1, justifyContent: "center" }}>
                <h2 style={{ fontSize: 20, fontWeight: 700, letterSpacing: -0.5, margin: 0 }}>{c.name || "Loading..."}</h2>
                {bestMatch && (
                  <span style={{
                    padding: "3px 10px", borderRadius: 12, fontSize: 10, fontWeight: 700,
                    background: `${fitColors[bestMatch.fit_level] || theme.textDim}20`,
                    color: fitColors[bestMatch.fit_level] || theme.textDim,
                    border: `1px solid ${fitColors[bestMatch.fit_level] || theme.textDim}40`,
                    textTransform: "uppercase",
                  }}>
                    {bestMatch.fit_level} fit ({bestMatch.score}%)
                  </span>
                )}
              </div>
              {c.email && (
                <a
                  href={`mailto:${c.email}`}
                  style={{
                    padding: "7px 16px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                    background: theme.accent, color: "#fff", border: "none",
                    cursor: "pointer", textDecoration: "none", whiteSpace: "nowrap",
                  }}
                >
                  ✉ Send Email
                </a>
              )}
            </div>

            {/* Two Column Layout */}
            <div style={{ display: "flex", gap: 24 }}>
              {/* LEFT PANEL (35%) */}
              <div style={{ width: "35%", flexShrink: 0, position: "sticky", top: 24, alignSelf: "flex-start" }}>
                {/* Contact Card */}
                <div style={{ padding: 16, borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, marginBottom: 16 }}>
                  <div style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                    Contact
                  </div>
                  {c.email && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                      <span style={{ fontSize: 12, color: theme.accent, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{c.email}</span>
                      <span onClick={() => copyToClipboard(c.email)} style={{ cursor: "pointer", fontSize: 13, opacity: 0.5, marginLeft: 8, flexShrink: 0 }} title="Copy email">📋</span>
                    </div>
                  )}
                  {c.phone && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                      <span style={{ fontSize: 12, color: theme.text }}>{c.phone}</span>
                      <span onClick={() => copyToClipboard(c.phone)} style={{ cursor: "pointer", fontSize: 13, opacity: 0.5, marginLeft: 8, flexShrink: 0 }} title="Copy phone">📋</span>
                    </div>
                  )}
                  {c.location && (
                    <div style={{ fontSize: 12, color: theme.textMuted }}>{c.location}</div>
                  )}
                  {!c.email && !c.phone && !c.location && (
                    <div style={{ fontSize: 12, color: theme.textDim }}>No contact info</div>
                  )}
                </div>

                {/* Skills */}
                <div style={{ padding: 16, borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, marginBottom: 16 }}>
                  <div style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                    Skills ({(c.skills || []).length})
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {(c.skills || []).length > 0 ? (c.skills || []).map((s, i) => (
                      <span key={i} style={{
                        padding: "2px 8px", borderRadius: 4, fontSize: 10, fontWeight: 600,
                        background: "rgba(20,184,166,0.15)", color: "#14b8a6",
                        textTransform: "uppercase", letterSpacing: 0.3,
                      }}>{s}</span>
                    )) : <span style={{ fontSize: 11, color: theme.textDim }}>No skills extracted</span>}
                  </div>
                </div>

                {/* Experience & Titles */}
                <div style={{ padding: 16, borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, marginBottom: 16 }}>
                  <div style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                    Experience
                  </div>
                  <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>
                    {c.years_exp != null ? `${c.years_exp} years` : "Not specified"}
                  </div>
                  {(c.titles || []).length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                      {(c.titles || []).map((t, i) => (
                        <div key={i} style={{ fontSize: 11, color: theme.textMuted, paddingLeft: 8, borderLeft: `2px solid ${theme.border}` }}>{t}</div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Tags */}
                <div style={{ padding: 16, borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, marginBottom: 16 }}>
                  <div style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                    Tags
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 8 }}>
                    {(c.tags || []).map((t, i) => (
                      <span key={i} style={{
                        display: "inline-flex", alignItems: "center", gap: 3,
                        padding: "3px 10px", borderRadius: 12, fontSize: 11, fontWeight: 600,
                        background: `${getTagColor(t)}20`, color: getTagColor(t),
                        border: `1px solid ${getTagColor(t)}40`,
                      }}>
                        {t}
                        <span
                          onClick={() => {
                            removeTagFromCandidate(c.id, t);
                            setCandidateProfile((prev) => prev ? { ...prev, tags: (prev.tags || []).filter((x) => x !== t) } : prev);
                            setSelectedCandidate((prev) => prev ? { ...prev, tags: (prev.tags || []).filter((x) => x !== t) } : prev);
                          }}
                          style={{ cursor: "pointer", fontSize: 13, lineHeight: 1, opacity: 0.7 }}
                          title="Remove tag"
                        >&times;</span>
                      </span>
                    ))}
                  </div>
                  {/* Suggested tags */}
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 8 }}>
                    {SUGGESTED_TAGS.filter((st) => !(c.tags || []).some((t) => t.toLowerCase() === st.toLowerCase())).map((st) => (
                      <button
                        key={st}
                        onClick={() => {
                          addTagToCandidate(c.id, st);
                          setCandidateProfile((prev) => prev ? { ...prev, tags: [...(prev.tags || []), st] } : prev);
                          setSelectedCandidate((prev) => prev ? { ...prev, tags: [...(prev.tags || []), st] } : prev);
                        }}
                        style={{
                          padding: "3px 8px", borderRadius: 10, fontSize: 10,
                          background: "transparent", border: `1px dashed ${theme.border}`,
                          color: theme.textDim, cursor: "pointer",
                        }}
                      >+ {st}</button>
                    ))}
                  </div>
                  {/* Custom tag input */}
                  <input
                    type="text"
                    placeholder="Custom tag + Enter"
                    value={candidateTagInput[c.id] || ""}
                    onChange={(e) => setCandidateTagInput((pr) => ({ ...pr, [c.id]: e.target.value }))}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && (candidateTagInput[c.id] || "").trim()) {
                        const newTag = candidateTagInput[c.id].trim();
                        addTagToCandidate(c.id, newTag);
                        setCandidateProfile((prev) => {
                          if (!prev) return prev;
                          const cur = prev.tags || [];
                          if (cur.some((t) => t.toLowerCase() === newTag.toLowerCase())) return prev;
                          return { ...prev, tags: [...cur, newTag] };
                        });
                        setSelectedCandidate((prev) => {
                          if (!prev) return prev;
                          const cur = prev.tags || [];
                          if (cur.some((t) => t.toLowerCase() === newTag.toLowerCase())) return prev;
                          return { ...prev, tags: [...cur, newTag] };
                        });
                        setCandidateTagInput((pr) => ({ ...pr, [c.id]: "" }));
                      }
                    }}
                    style={{
                      padding: "5px 10px", borderRadius: 6, fontSize: 11, width: "100%",
                      background: theme.bg, border: `1px solid ${theme.border}`,
                      color: theme.text, outline: "none", boxSizing: "border-box",
                    }}
                  />
                </div>

                {/* Notes */}
                <div style={{ padding: 16, borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                    <span style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      Notes
                    </span>
                    {profileNotesSaved && <span style={{ fontSize: 10, color: theme.success, fontWeight: 600 }}>✓ Saved</span>}
                  </div>
                  <textarea
                    placeholder="Add recruiter notes..."
                    value={candidateNotesDraft[c.id] !== undefined ? candidateNotesDraft[c.id] : (c.notes || "")}
                    onChange={(e) => setCandidateNotesDraft((pr) => ({ ...pr, [c.id]: e.target.value }))}
                    onBlur={() => {
                      const draft = candidateNotesDraft[c.id];
                      if (draft !== undefined && draft !== (c.notes || "")) {
                        saveCandidateNotes(c.id, draft);
                        setCandidateProfile((prev) => prev ? { ...prev, notes: draft } : prev);
                        setSelectedCandidate((prev) => prev ? { ...prev, notes: draft } : prev);
                        setProfileNotesSaved(true);
                        setTimeout(() => setProfileNotesSaved(false), 2000);
                      }
                      setCandidateNotesDraft((pr) => { const n = { ...pr }; delete n[c.id]; return n; });
                    }}
                    rows={4}
                    style={{
                      width: "100%", padding: "8px 12px", borderRadius: 8, fontSize: 12,
                      background: theme.bg, border: `1px solid ${theme.border}`,
                      color: theme.text, outline: "none", resize: "vertical",
                      fontFamily: "inherit", lineHeight: 1.5, boxSizing: "border-box",
                    }}
                  />
                  <div style={{ fontSize: 10, color: theme.textDim, marginTop: 4 }}>Auto-saves on blur</div>
                </div>
              </div>

              {/* RIGHT PANEL (65%) */}
              <div style={{ flex: 1, minWidth: 0 }}>
                {/* Tab buttons */}
                <div style={{ display: "flex", gap: 4, marginBottom: 16, borderBottom: `1px solid ${theme.border}`, paddingBottom: 0 }}>
                  {[
                    { id: "resume", label: "Resume" },
                    { id: "matches", label: "Match History" },
                    { id: "email", label: "Source Email" },
                  ].map((tab) => (
                    <button
                      key={tab.id}
                      onClick={() => setProfileActiveTab(tab.id)}
                      style={{
                        padding: "8px 16px", borderRadius: "6px 6px 0 0", fontSize: 12, fontWeight: 600,
                        background: profileActiveTab === tab.id ? theme.surface : "transparent",
                        border: profileActiveTab === tab.id ? `1px solid ${theme.border}` : "1px solid transparent",
                        borderBottom: profileActiveTab === tab.id ? `1px solid ${theme.surface}` : "1px solid transparent",
                        color: profileActiveTab === tab.id ? theme.text : theme.textDim,
                        cursor: "pointer", marginBottom: -1,
                      }}
                    >{tab.label}</button>
                  ))}
                </div>

                {/* Loading state */}
                {profileLoading && (
                  <div style={{
                    padding: 40, textAlign: "center", borderRadius: 10,
                    background: theme.surface, border: `1px solid ${theme.border}`,
                  }}>
                    <div style={{ fontSize: 13, color: theme.textDim, animation: "pulse 1.5s infinite" }}>Loading profile data...</div>
                  </div>
                )}

                {/* Resume tab */}
                {!profileLoading && profileActiveTab === "resume" && (
                  <div style={{ borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, overflow: "hidden" }}>
                    {p?.resume_text ? (
                      <pre style={{
                        padding: 20, margin: 0, maxHeight: 600, overflowY: "auto",
                        fontSize: 12, lineHeight: 1.6, color: theme.text,
                        fontFamily: "'Consolas', 'Monaco', 'Courier New', monospace",
                        whiteSpace: "pre-wrap", wordBreak: "break-word",
                        background: theme.surface,
                      }}>
                        {p.resume_text}
                      </pre>
                    ) : (
                      <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                        {p ? "No resume text available" : "Loading..."}
                      </div>
                    )}
                  </div>
                )}

                {/* Match History tab */}
                {!profileLoading && profileActiveTab === "matches" && (
                  <div style={{ borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, overflow: "hidden" }}>
                    {(p?.match_history || []).length > 0 ? (
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                        <thead>
                          <tr style={{ borderBottom: `1px solid ${theme.border}` }}>
                            {["#", "Job Title", "Score", "Fit Level", "Date"].map((h) => (
                              <th key={h} style={{
                                padding: "10px 14px", textAlign: "left", fontSize: 10,
                                color: theme.textDim, fontWeight: 600, textTransform: "uppercase",
                                letterSpacing: 0.5,
                              }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {p.match_history.map((m, i) => (
                            <React.Fragment key={i}>
                              <tr
                                onClick={() => setExpandedMatchIdx(expandedMatchIdx === i ? null : i)}
                                style={{
                                  borderBottom: `1px solid ${theme.border}`, cursor: "pointer",
                                  background: expandedMatchIdx === i ? theme.bg : "transparent",
                                }}
                              >
                                <td style={{ padding: "10px 14px", color: theme.textDim }}>{i + 1}</td>
                                <td style={{ padding: "10px 14px", fontWeight: 600 }}>{m.job_title}</td>
                                <td style={{ padding: "10px 14px" }}>
                                  <span style={{
                                    padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700,
                                    background: `${fitColors[m.fit_level] || theme.textDim}20`,
                                    color: fitColors[m.fit_level] || theme.textDim,
                                  }}>{m.score}%</span>
                                </td>
                                <td style={{ padding: "10px 14px", color: fitColors[m.fit_level] || theme.textDim, fontWeight: 600, textTransform: "capitalize" }}>
                                  {m.fit_level}
                                </td>
                                <td style={{ padding: "10px 14px", color: theme.textDim }}>
                                  {m.matched_at ? new Date(m.matched_at).toLocaleDateString() : "\u2014"}
                                </td>
                              </tr>
                              {expandedMatchIdx === i && (
                                <tr>
                                  <td colSpan={5} style={{ padding: "12px 14px 16px 42px", background: theme.bg }}>
                                    <div style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>
                                      Match Reasons
                                    </div>
                                    <ul style={{ margin: 0, paddingLeft: 16, listStyle: "disc" }}>
                                      {(m.match_reasons || []).map((r, ri) => (
                                        <li key={ri} style={{ fontSize: 11, color: theme.textMuted, lineHeight: 1.8 }}>{r}</li>
                                      ))}
                                    </ul>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                        {p ? "No match runs yet \u2014 go to Jobs & Match to run matching" : "Loading..."}
                      </div>
                    )}
                  </div>
                )}

                {/* Source Email tab */}
                {!profileLoading && profileActiveTab === "email" && (
                  <div style={{ borderRadius: 10, background: theme.surface, border: `1px solid ${theme.border}`, overflow: "hidden" }}>
                    {p?.source_email ? (
                      <div style={{ padding: 20 }}>
                        <div style={{ display: "grid", gridTemplateColumns: "80px 1fr", gap: "8px 12px", marginBottom: 16 }}>
                          <span style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", paddingTop: 2 }}>Subject</span>
                          <span style={{ fontSize: 13, fontWeight: 600 }}>{p.source_email.subject || "(No Subject)"}</span>
                          <span style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", paddingTop: 2 }}>From</span>
                          <span style={{ fontSize: 12, color: theme.textMuted }}>
                            {p.source_email.sender || "Unknown"} {p.source_email.sender_email && <span style={{ color: theme.accent }}>&lt;{p.source_email.sender_email}&gt;</span>}
                          </span>
                          <span style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", paddingTop: 2 }}>Date</span>
                          <span style={{ fontSize: 12, color: theme.textMuted }}>
                            {p.source_email.date ? new Date(p.source_email.date).toLocaleString() : "\u2014"}
                          </span>
                        </div>
                        <div style={{
                          maxHeight: 400, overflowY: "auto", padding: 16, borderRadius: 8,
                          background: theme.bg, border: `1px solid ${theme.border}`,
                          fontSize: 12, lineHeight: 1.6, color: theme.textMuted,
                          whiteSpace: "pre-wrap", wordBreak: "break-word",
                        }}>
                          {p.source_email.body_text || "(No body text)"}
                        </div>
                      </div>
                    ) : (
                      <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                        {p ? "Source email not available" : "Loading..."}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
          );
        })()}

        {/* ─── Jobs & Match View ───────────────────────────────────── */}
        {view === "jobs" && (
          <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>
            {/* Left Panel: Create Job + Job List */}
            <div style={{
              width: 380, flexShrink: 0, borderRight: `1px solid ${theme.border}`,
              overflowY: "auto", padding: 20,
            }}>
              <h3 style={{ fontSize: 14, fontWeight: 700, marginBottom: 16, letterSpacing: -0.3 }}>
                Create Job Requisition
              </h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 24 }}>
                <div>
                  <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                    Job Title
                  </label>
                  <input
                    type="text"
                    placeholder="e.g. Senior Software Engineer"
                    value={jobForm.title}
                    onChange={(e) => setJobForm({ ...jobForm, title: e.target.value })}
                    style={{
                      width: "100%", padding: "8px 12px", borderRadius: 6,
                      background: theme.bg, border: `1px solid ${theme.border}`,
                      color: theme.text, fontSize: 12, outline: "none",
                    }}
                  />
                </div>
                <div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      Job Description
                    </label>
                    <label style={{
                      fontSize: 10, fontWeight: 600, color: theme.accent,
                      cursor: jdUploading ? "not-allowed" : "pointer",
                      opacity: jdUploading ? 0.5 : 1,
                      display: "flex", alignItems: "center", gap: 4,
                    }}>
                      {jdUploading ? (
                        <><span style={{
                          display: "inline-block", width: 10, height: 10,
                          border: `2px solid ${theme.border}`, borderTopColor: theme.accent,
                          borderRadius: "50%", animation: "spin 0.8s linear infinite",
                        }} /> Parsing...</>
                      ) : (
                        <>&#8593; Upload PDF/DOCX</>
                      )}
                      <input
                        type="file"
                        accept=".pdf,.docx,.doc"
                        disabled={jdUploading}
                        style={{ display: "none" }}
                        onChange={(e) => {
                          if (e.target.files[0]) uploadJD(e.target.files[0]);
                          e.target.value = "";
                        }}
                      />
                    </label>
                  </div>
                  <textarea
                    placeholder="Paste the full JD text here or upload a file..."
                    value={jobForm.jd_raw}
                    onChange={(e) => setJobForm({ ...jobForm, jd_raw: e.target.value })}
                    rows={4}
                    style={{
                      width: "100%", padding: "8px 12px", borderRadius: 6,
                      background: theme.bg, border: `1px solid ${theme.border}`,
                      color: theme.text, fontSize: 12, outline: "none", resize: "vertical",
                      fontFamily: "inherit",
                    }}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                    Required Skills (comma-separated)
                  </label>
                  <input
                    type="text"
                    placeholder="python, react, sql, docker"
                    value={jobForm.required_skills}
                    onChange={(e) => setJobForm({ ...jobForm, required_skills: e.target.value })}
                    style={{
                      width: "100%", padding: "8px 12px", borderRadius: 6,
                      background: theme.bg, border: `1px solid ${theme.border}`,
                      color: theme.text, fontSize: 12, outline: "none",
                    }}
                  />
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <div>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      Min Experience (yrs)
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={jobForm.min_exp}
                      onChange={(e) => setJobForm({ ...jobForm, min_exp: e.target.value })}
                      style={{
                        width: "100%", padding: "8px 12px", borderRadius: 6,
                        background: theme.bg, border: `1px solid ${theme.border}`,
                        color: theme.text, fontSize: 12, outline: "none",
                      }}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      Location
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. New York, NY"
                      value={jobForm.location}
                      onChange={(e) => setJobForm({ ...jobForm, location: e.target.value })}
                      style={{
                        width: "100%", padding: "8px 12px", borderRadius: 6,
                        background: theme.bg, border: `1px solid ${theme.border}`,
                        color: theme.text, fontSize: 12, outline: "none",
                      }}
                    />
                  </div>
                </div>
                <label style={{
                  display: "flex", alignItems: "center", gap: 8,
                  fontSize: 12, color: theme.textMuted, cursor: "pointer",
                }}>
                  <input
                    type="checkbox"
                    checked={jobForm.remote_ok}
                    onChange={(e) => setJobForm({ ...jobForm, remote_ok: e.target.checked })}
                    style={{ accentColor: theme.accent }}
                  />
                  Remote OK
                </label>
                <button
                  onClick={createJob}
                  disabled={!jobForm.title.trim()}
                  style={{
                    padding: "10px 20px", borderRadius: 8, border: "none",
                    background: theme.accent, color: "#fff", fontSize: 12,
                    fontWeight: 600, cursor: !jobForm.title.trim() ? "not-allowed" : "pointer",
                    opacity: !jobForm.title.trim() ? 0.5 : 1,
                  }}
                >
                  Save Job
                </button>
              </div>

              {/* Job List */}
              <div style={{
                fontSize: 10, fontWeight: 600, color: theme.textDim, marginBottom: 10,
                textTransform: "uppercase", letterSpacing: 1,
              }}>
                Saved Jobs ({jobs.length})
              </div>
              {jobs.length === 0 ? (
                <div style={{ padding: 20, textAlign: "center", color: theme.textDim, fontSize: 11 }}>
                  No jobs yet. Create one above.
                </div>
              ) : jobs.map((j) => (
                <div
                  key={j.id}
                  onClick={() => { setSelectedJobId(j.id); loadMatchResults(j.id); }}
                  style={{
                    padding: "10px 12px", borderRadius: 6, marginBottom: 6,
                    background: selectedJobId === j.id ? theme.accentGlow : theme.bg,
                    border: `1px solid ${selectedJobId === j.id ? theme.accent : theme.border}`,
                    cursor: "pointer", transition: "all 0.15s",
                  }}
                >
                  <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>{j.title}</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {(j.required_skills || []).slice(0, 4).map((s, i) => (
                      <Badge key={i} color={theme.purple}>{s}</Badge>
                    ))}
                    {(j.required_skills || []).length > 4 && (
                      <Badge color={theme.textDim}>+{j.required_skills.length - 4}</Badge>
                    )}
                  </div>
                  <div style={{ fontSize: 10, color: theme.textDim, marginTop: 6 }}>
                    {j.min_exp ? `${j.min_exp}+ yrs` : "Any exp"} · {j.location || "Any location"}
                    {j.remote_ok && " · Remote"}
                  </div>
                </div>
              ))}
            </div>

            {/* Right Panel: Match Results */}
            <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
              {selectedJobId ? (
                <>
                  <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
                    <h3 style={{ fontSize: 14, fontWeight: 700, flex: 1, letterSpacing: -0.3 }}>
                      Match Results
                    </h3>
                    <button
                      onClick={() => runJobMatch(selectedJobId)}
                      disabled={matchLoading}
                      style={{
                        padding: "8px 20px", borderRadius: 6, border: "none",
                        background: theme.success, color: "#fff", fontSize: 12,
                        fontWeight: 600, cursor: matchLoading ? "not-allowed" : "pointer",
                        opacity: matchLoading ? 0.6 : 1,
                      }}
                    >
                      {matchLoading ? "Matching..." : "Run Match"}
                    </button>
                  </div>

                  {matchLoading ? <Spinner /> : matchResults.length === 0 ? (
                    <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                      No results yet. Click "Run Match" to score all candidates against this job.
                    </div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {matchResults.map((r, i) => {
                        const fitColor = r.fit_level === "high" ? theme.success
                          : r.fit_level === "medium" ? theme.warning : theme.danger;
                        const c = r.candidate;
                        const matchedSkills = new Set();
                        (r.match_reasons || []).forEach((reason) => {
                          const m = reason.match(/Skills matched.*?:\s*(.+)/i);
                          if (m) m[1].split(",").forEach((s) => matchedSkills.add(s.trim().toLowerCase()));
                        });
                        return (
                          <div key={i} style={{
                            background: theme.surface, borderRadius: 8, padding: 16,
                            border: `1px solid ${theme.border}`,
                          }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                              {/* Score Badge */}
                              <div style={{
                                width: 48, height: 48, borderRadius: "50%",
                                background: `${fitColor}20`, border: `2px solid ${fitColor}`,
                                display: "flex", alignItems: "center", justifyContent: "center",
                                flexShrink: 0,
                              }}>
                                <span style={{ fontSize: 14, fontWeight: 700, color: fitColor }}>
                                  {r.score}
                                </span>
                              </div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 2 }}>
                                  {c.name}
                                </div>
                                <div style={{ fontSize: 11, color: theme.textMuted }}>
                                  {c.email || ""}{c.phone ? ` · ${c.phone}` : ""}
                                  {c.location ? ` · ${c.location}` : ""}
                                </div>
                              </div>
                              <Badge color={fitColor}>{r.fit_level}</Badge>
                            </div>

                            {/* Skills with match highlighting */}
                            {c.skills && c.skills.length > 0 && (
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 10 }}>
                                {c.skills.map((s, si) => (
                                  <Badge
                                    key={si}
                                    color={matchedSkills.has(s.toLowerCase()) ? theme.success : theme.textDim}
                                  >
                                    {s}
                                  </Badge>
                                ))}
                              </div>
                            )}

                            {/* Match Reasons */}
                            <div style={{ fontSize: 11, color: theme.textMuted, lineHeight: 1.6 }}>
                              {(r.match_reasons || []).map((reason, ri) => (
                                <div key={ri} style={{ paddingLeft: 8 }}>· {reason}</div>
                              ))}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </>
              ) : (
                <div style={{ padding: 40, textAlign: "center", color: theme.textDim, fontSize: 13 }}>
                  Select a job from the left panel to view or run matches.
                </div>
              )}
            </div>
          </div>
        )}

        {/* ─── Recruit Export View ─────────────────────────────────── */}
        {view === "recruit-export" && (
          <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 4, letterSpacing: -0.5 }}>
              Recruit Export
            </h2>
            <p style={{ color: theme.textMuted, fontSize: 12, marginBottom: 24 }}>
              Download match results as CSV for a specific job requisition.
            </p>

            <div style={{ maxWidth: 400 }}>
              <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>
                Select Job
              </label>
              <select
                value={exportJobId}
                onChange={(e) => setExportJobId(e.target.value)}
                style={{
                  width: "100%", padding: "8px 12px", borderRadius: 6, marginBottom: 16,
                  background: theme.surface, border: `1px solid ${theme.border}`,
                  color: theme.text, fontSize: 12,
                }}
              >
                <option value="">— Select a job —</option>
                {jobs.map((j) => (
                  <option key={j.id} value={j.id}>{j.title}</option>
                ))}
              </select>
              <button
                onClick={() => exportRecruitCSV(exportJobId)}
                disabled={!exportJobId}
                style={{
                  padding: "10px 24px", borderRadius: 8, border: "none",
                  background: theme.accent, color: "#fff", fontSize: 13,
                  fontWeight: 600, cursor: !exportJobId ? "not-allowed" : "pointer",
                  opacity: !exportJobId ? 0.5 : 1,
                  boxShadow: `0 4px 16px ${theme.accentGlow}`,
                }}
              >
                Download CSV
              </button>
            </div>
          </div>
        )}

        {/* ─── Settings View ──────────────────────────────────────── */}
        {view === "settings" && (() => {
          const h = storageHealth;
          const timeAgo = (iso) => {
            if (!iso) return "Never";
            const diff = Date.now() - new Date(iso).getTime();
            const mins = Math.floor(diff / 60000);
            if (mins < 1) return "Just now";
            if (mins < 60) return `${mins}m ago`;
            const hrs = Math.floor(mins / 60);
            if (hrs < 24) return `${hrs}h ago`;
            return `${Math.floor(hrs / 24)}d ago`;
          };
          const statusColor = h?.health_status === "healthy" ? "#22c55e"
            : h?.health_status === "warning" ? "#f59e0b" : "#ef4444";
          const statusLabel = h?.health_status === "healthy" ? "Healthy"
            : h?.health_status === "warning" ? "Warning" : "Empty";
          const typeColors = { pdf: "#ef4444", docx: "#3b82f6", doc: "#60a5fa", image: "#22c55e", other: "#6b7280" };

          return (
          <div style={{ flex: 1, overflowY: "auto", padding: 24, maxWidth: 900, margin: "0 auto" }}>
            <h2 style={{ fontSize: 18, fontWeight: 700, marginBottom: 4, letterSpacing: -0.5 }}>Settings</h2>
            <p style={{ color: theme.textMuted, fontSize: 12, marginBottom: 24 }}>Storage health, sync status, and data management.</p>

            {/* ─── Auto-Scrape Scheduler Section ──────────────────── */}
            {(() => {
              const sc = schedulerConfig;
              const isEnabled = sc?.enabled || false;
              const isRunning = sc?.is_running || false;
              const formatCountdown = (secs) => {
                if (secs == null || secs <= 0) return "0s";
                const m = Math.floor(secs / 60);
                const s = secs % 60;
                return m > 0 ? `${m}m ${s}s` : `${s}s`;
              };
              const INTERVAL_OPTIONS = [
                { value: 15, label: "Every 15 mins" },
                { value: 30, label: "Every 30 mins" },
                { value: 60, label: "Every hour" },
                { value: 120, label: "Every 2 hours" },
                { value: 360, label: "Every 6 hours" },
              ];
              return (
              <div style={{
                background: theme.surface,
                border: `1px solid ${theme.border}`,
                borderLeft: `3px solid ${isEnabled ? "#22c55e" : theme.textDim}`,
                borderRadius: 12, overflow: "hidden", marginBottom: 20,
              }}>
                {/* Top row: title + toggle */}
                <div style={{
                  padding: "16px 20px", borderBottom: `1px solid ${theme.border}`,
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{
                      width: 8, height: 8, borderRadius: "50%",
                      background: isEnabled && isRunning ? "#22c55e" : theme.textDim,
                      boxShadow: isEnabled && isRunning ? "0 0 6px #22c55e" : "none",
                      display: "inline-block",
                    }} />
                    <span style={{ fontSize: 15, fontWeight: 700 }}>Auto-Scrape Scheduler</span>
                  </div>
                  {/* Toggle switch */}
                  <div
                    onClick={async () => {
                      if (schedulerLoading) return;
                      setSchedulerLoading(true);
                      try {
                        const res = await fetch(`${API}/api/scheduler/config`, {
                          method: "POST",
                          headers: { "Content-Type": "application/json" },
                          body: JSON.stringify({
                            enabled: !isEnabled,
                            interval_minutes: schedulerForm.interval_minutes,
                            folder: schedulerForm.folder || "INBOX",
                            subject_filter: schedulerForm.subject_filter || null,
                          }),
                        });
                        if (!res.ok) throw new Error(`HTTP ${res.status}`);
                        notify(isEnabled ? "Scheduler disabled" : "Scheduler enabled", "success");
                        await loadSchedulerConfig();
                      } catch (e) {
                        notify("Failed to toggle scheduler", "error");
                      } finally {
                        setSchedulerLoading(false);
                      }
                    }}
                    style={{
                      width: 42, height: 22, borderRadius: 11, cursor: "pointer",
                      background: isEnabled ? "#22c55e" : theme.textDim,
                      position: "relative", transition: "background 0.2s",
                    }}
                  >
                    <div style={{
                      width: 16, height: 16, borderRadius: "50%", background: "#fff",
                      position: "absolute", top: 3,
                      left: isEnabled ? 23 : 3,
                      transition: "left 0.2s",
                    }} />
                  </div>
                </div>

                {/* Settings form — visible when enabled */}
                {isEnabled && (
                  <div style={{ padding: "16px 20px", borderBottom: `1px solid ${theme.border}` }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
                      <div>
                        <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                          Interval
                        </label>
                        <select
                          value={schedulerForm.interval_minutes}
                          onChange={(e) => setSchedulerForm((p) => ({ ...p, interval_minutes: parseInt(e.target.value) }))}
                          style={{
                            width: "100%", padding: "7px 10px", borderRadius: 6, fontSize: 12,
                            background: theme.bg, border: `1px solid ${theme.border}`,
                            color: theme.text, outline: "none",
                          }}
                        >
                          {INTERVAL_OPTIONS.map((o) => (
                            <option key={o.value} value={o.value}>{o.label}</option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                          Folder
                        </label>
                        <input
                          type="text"
                          value={schedulerForm.folder}
                          onChange={(e) => setSchedulerForm((p) => ({ ...p, folder: e.target.value }))}
                          placeholder="INBOX"
                          style={{
                            width: "100%", padding: "7px 10px", borderRadius: 6, fontSize: 12,
                            background: theme.bg, border: `1px solid ${theme.border}`,
                            color: theme.text, outline: "none", boxSizing: "border-box",
                          }}
                        />
                      </div>
                    </div>
                    <div style={{ marginBottom: 14 }}>
                      <label style={{ fontSize: 10, color: theme.textDim, fontWeight: 600, display: "block", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>
                        Subject Filter
                      </label>
                      <input
                        type="text"
                        value={schedulerForm.subject_filter}
                        onChange={(e) => setSchedulerForm((p) => ({ ...p, subject_filter: e.target.value }))}
                        placeholder="Optional keyword e.g. Candidate"
                        style={{
                          width: "100%", padding: "7px 10px", borderRadius: 6, fontSize: 12,
                          background: theme.bg, border: `1px solid ${theme.border}`,
                          color: theme.text, outline: "none", boxSizing: "border-box",
                        }}
                      />
                    </div>
                    <div style={{ display: "flex", gap: 10 }}>
                      <button
                        onClick={async () => {
                          setSchedulerLoading(true);
                          try {
                            const res = await fetch(`${API}/api/scheduler/config`, {
                              method: "POST",
                              headers: { "Content-Type": "application/json" },
                              body: JSON.stringify({
                                enabled: true,
                                interval_minutes: schedulerForm.interval_minutes,
                                folder: schedulerForm.folder || "INBOX",
                                subject_filter: schedulerForm.subject_filter || null,
                              }),
                            });
                            if (!res.ok) throw new Error(`HTTP ${res.status}`);
                            notify("Scheduler settings saved", "success");
                            await loadSchedulerConfig();
                          } catch (e) {
                            notify("Failed to save settings", "error");
                          } finally {
                            setSchedulerLoading(false);
                          }
                        }}
                        disabled={schedulerLoading}
                        style={{
                          padding: "7px 16px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                          background: theme.accent, color: "#fff", border: "none",
                          cursor: schedulerLoading ? "not-allowed" : "pointer",
                          opacity: schedulerLoading ? 0.6 : 1,
                        }}
                      >Save Settings</button>
                      <button
                        onClick={async () => {
                          setSchedulerRunNow(true);
                          try {
                            const res = await fetch(`${API}/api/scheduler/run-now`, { method: "POST" });
                            if (!res.ok) throw new Error(`HTTP ${res.status}`);
                            notify("Scrape started", "success");
                          } catch (e) {
                            notify("Failed to start scrape", "error");
                          }
                          setTimeout(async () => {
                            setSchedulerRunNow(false);
                            await loadSchedulerConfig();
                          }, 3000);
                        }}
                        disabled={schedulerRunNow}
                        style={{
                          padding: "7px 16px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                          background: "transparent", border: `1px solid ${theme.border}`,
                          color: schedulerRunNow ? theme.success : theme.textMuted,
                          cursor: schedulerRunNow ? "not-allowed" : "pointer",
                        }}
                      >{schedulerRunNow ? "Running..." : "▶ Run Now"}</button>
                    </div>
                  </div>
                )}

                {/* Status row */}
                <div style={{ padding: "12px 20px", fontSize: 12, color: theme.textMuted, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
                  {sc?.last_run_at ? (
                    <span>
                      Last run: {(() => {
                        const diff = Date.now() - new Date(sc.last_run_at).getTime();
                        const mins = Math.floor(diff / 60000);
                        if (mins < 1) return "Just now";
                        if (mins < 60) return `${mins}m ago`;
                        const hrs = Math.floor(mins / 60);
                        if (hrs < 24) return `${hrs}h ago`;
                        return `${Math.floor(hrs / 24)}d ago`;
                      })()} · Found {sc.emails_found_last_run} emails · {sc.candidates_added_last_run} new candidates
                    </span>
                  ) : (
                    <span style={{ color: theme.textDim }}>Not run yet</span>
                  )}
                  {isEnabled && schedulerCountdown != null && schedulerCountdown > 0 && (
                    <span style={{ color: theme.accent, fontWeight: 600 }}>
                      Next run in: {formatCountdown(schedulerCountdown)}
                    </span>
                  )}
                </div>
              </div>
              );
            })()}

            {!h ? (
              <Spinner />
            ) : (
              <div style={{
                background: theme.surface, border: `1px solid ${theme.border}`,
                borderRadius: 12, overflow: "hidden",
              }}>
                {/* Header */}
                <div style={{
                  padding: "16px 20px", borderBottom: `1px solid ${theme.border}`,
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 18 }}>💾</span>
                    <span style={{ fontSize: 15, fontWeight: 700 }}>Storage Health</span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{
                      padding: "3px 10px", borderRadius: 12, fontSize: 11, fontWeight: 700,
                      background: `${statusColor}20`, color: statusColor,
                      border: `1px solid ${statusColor}40`,
                    }}>{statusLabel}</span>
                    <button
                      onClick={loadStorageHealth}
                      title="Refresh"
                      style={{
                        width: 30, height: 30, borderRadius: "50%", border: `1px solid ${theme.border}`,
                        background: "transparent", color: theme.textMuted, cursor: "pointer",
                        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14,
                      }}
                    >↻</button>
                  </div>
                </div>

                {/* Database Section */}
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${theme.border}` }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", letterSpacing: 1, marginBottom: 12 }}>
                    Database
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 12 }}>
                    {[
                      { icon: "📁", label: "Database Size", value: h.database.db_size_human },
                      { icon: "👤", label: "Candidates", value: h.record_counts.candidates },
                      { icon: "📧", label: "Emails Scraped", value: h.record_counts.scraped_emails },
                      { icon: "💼", label: "Jobs Created", value: h.record_counts.jobs },
                      { icon: "🎯", label: "Match Results", value: h.record_counts.match_results },
                      { icon: "📝", label: "With Notes", value: h.record_counts.notes_count },
                      { icon: "🏷️", label: "Tagged", value: h.record_counts.tagged_count },
                    ].map((s) => (
                      <div key={s.label} style={{
                        padding: "10px 12px", borderRadius: 8, background: theme.bg,
                        border: `1px solid ${theme.border}`,
                      }}>
                        <div style={{ fontSize: 11, color: theme.textDim, marginBottom: 4 }}>
                          {s.icon} {s.label}
                        </div>
                        <div style={{ fontSize: 18, fontWeight: 700 }}>{s.value}</div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Attachments Section */}
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${theme.border}` }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", letterSpacing: 1, marginBottom: 12 }}>
                    Attachments
                  </div>
                  <div style={{ display: "flex", gap: 20, marginBottom: 14 }}>
                    <div style={{ padding: "10px 12px", borderRadius: 8, background: theme.bg, border: `1px solid ${theme.border}`, flex: 1 }}>
                      <div style={{ fontSize: 11, color: theme.textDim, marginBottom: 4 }}>📎 Total Files</div>
                      <div style={{ fontSize: 18, fontWeight: 700 }}>{h.attachments.total_files}</div>
                    </div>
                    <div style={{ padding: "10px 12px", borderRadius: 8, background: theme.bg, border: `1px solid ${theme.border}`, flex: 1 }}>
                      <div style={{ fontSize: 11, color: theme.textDim, marginBottom: 4 }}>💿 Storage Used</div>
                      <div style={{ fontSize: 18, fontWeight: 700 }}>{h.attachments.total_size_human}</div>
                    </div>
                  </div>
                  {/* Type breakdown bar */}
                  {h.attachments.total_files > 0 && (() => {
                    const types = h.attachments.by_type;
                    const segments = Object.entries(types).sort((a, b) => b[1] - a[1]);
                    return (
                      <div>
                        <div style={{ display: "flex", width: "100%", height: 10, borderRadius: 6, overflow: "hidden", marginBottom: 8 }}>
                          {segments.map(([type, count]) => (
                            <div key={type} style={{
                              flex: count,
                              background: typeColors[type] || typeColors.other,
                            }} />
                          ))}
                        </div>
                        <div style={{ fontSize: 11, color: theme.textDim }}>
                          {segments.filter(([, c]) => c > 0).map(([type, count]) => `${count} ${type.toUpperCase()}`).join(" · ")}
                        </div>
                      </div>
                    );
                  })()}
                </div>

                {/* Sync Status Section */}
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${theme.border}` }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: theme.textDim, textTransform: "uppercase", letterSpacing: 1, marginBottom: 12 }}>
                    Sync Status
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    {[
                      { icon: "🕐", label: "Last Email Scraped", value: timeAgo(h.sync.last_scraped_at) },
                      { icon: "👤", label: "Last Candidate Added", value: timeAgo(h.sync.last_candidate_added) },
                      { icon: "🎯", label: "Last Match Run", value: timeAgo(h.sync.last_match_run) },
                    ].map((s) => (
                      <div key={s.label} style={{
                        padding: "10px 12px", borderRadius: 8, background: theme.bg,
                        border: `1px solid ${theme.border}`,
                      }}>
                        <div style={{ fontSize: 11, color: theme.textDim, marginBottom: 4 }}>{s.icon} {s.label}</div>
                        <div style={{ fontSize: 13, fontWeight: 600 }}>{s.value}</div>
                      </div>
                    ))}
                    <div style={{
                      padding: "10px 12px", borderRadius: 8, background: theme.bg,
                      border: `1px solid ${theme.border}`, display: "flex", flexDirection: "column", gap: 6,
                    }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
                        <span style={{ fontSize: 14 }}>🔐</span>
                        <span style={{ color: theme.textDim }}>Session Saved</span>
                        <span style={{ marginLeft: "auto", fontSize: 14 }}>
                          {h.sync.session_file_exists
                            ? <span style={{ color: "#22c55e" }}>✓</span>
                            : <span style={{ color: "#ef4444" }}>✗</span>
                          }
                        </span>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
                        <span style={{ fontSize: 14 }}>📡</span>
                        <span style={{ color: theme.textDim }}>IMAP Connected</span>
                        <span style={{
                          marginLeft: "auto", width: 8, height: 8, borderRadius: "50%",
                          background: h.sync.imap_connected ? "#22c55e" : "#6b7280",
                          display: "inline-block",
                        }} />
                      </div>
                    </div>
                  </div>
                </div>

                {/* Action Buttons */}
                <div style={{
                  padding: "16px 20px",
                  display: "flex", gap: 12, alignItems: "center", justifyContent: "space-between",
                }}>
                  <div style={{ display: "flex", gap: 12 }}>
                    <button
                      onClick={() => { setShowClearModal(true); setClearConfirmText(""); }}
                      style={{
                        padding: "8px 18px", borderRadius: 8, fontSize: 12, fontWeight: 600,
                        background: "transparent", border: `1px solid ${theme.danger}`,
                        color: theme.danger, cursor: "pointer",
                      }}
                    >
                      🔄 Clear All Data
                    </button>
                    <button
                      onClick={async () => {
                        try {
                          const res = await fetch(`${API}/api/storage/backup`);
                          if (!res.ok) throw new Error(`HTTP ${res.status}`);
                          const blob = await res.blob();
                          const url = URL.createObjectURL(blob);
                          const a = document.createElement("a");
                          a.href = url;
                          a.download = `mailscraper_backup_${new Date().toISOString().slice(0, 10)}.db`;
                          a.click();
                          URL.revokeObjectURL(url);
                          notify("Backup downloaded", "success");
                        } catch (e) {
                          notify("Failed to download backup", "error");
                        }
                      }}
                      style={{
                        padding: "8px 18px", borderRadius: 8, fontSize: 12, fontWeight: 600,
                        background: "transparent", border: `1px solid ${theme.border}`,
                        color: theme.textMuted, cursor: "pointer",
                      }}
                    >
                      📦 Export Backup
                    </button>
                  </div>
                  {storageLastUpdated && (
                    <div style={{ fontSize: 10, color: theme.textDim }}>
                      Updated {timeAgo(storageLastUpdated.toISOString())}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
          );
        })()}

        {/* ─── Clear Data Confirmation Modal ─────────────────────────── */}
        {showClearModal && (
          <div
            style={{
              position: "fixed", inset: 0, zIndex: 3000,
              background: "rgba(0,0,0,0.7)", display: "flex",
              alignItems: "center", justifyContent: "center",
            }}
            onClick={() => setShowClearModal(false)}
          >
            <div
              style={{
                background: theme.surface, borderRadius: 12, padding: 24,
                border: `1px solid ${theme.border}`, maxWidth: 440, width: "90%",
                boxShadow: "0 20px 50px rgba(0,0,0,0.5)",
              }}
              onClick={(e) => e.stopPropagation()}
            >
              <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 8, color: theme.danger }}>
                🔄 Clear All Data
              </h3>
              <p style={{ fontSize: 12, color: theme.textMuted, lineHeight: 1.6, marginBottom: 16 }}>
                This will delete all candidates, emails, jobs, and match results from the local database.
                Attachment files will <strong>NOT</strong> be deleted.
              </p>
              <p style={{ fontSize: 12, color: theme.textMuted, marginBottom: 12 }}>
                Type <strong style={{ color: theme.danger }}>CONFIRM</strong> to proceed:
              </p>
              <input
                type="text"
                value={clearConfirmText}
                onChange={(e) => setClearConfirmText(e.target.value)}
                placeholder="Type CONFIRM"
                style={{
                  width: "100%", padding: "8px 12px", borderRadius: 6, fontSize: 13,
                  background: theme.bg, border: `1px solid ${theme.border}`,
                  color: theme.text, outline: "none", marginBottom: 16,
                }}
                autoFocus
              />
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button
                  onClick={() => setShowClearModal(false)}
                  style={{
                    padding: "8px 18px", borderRadius: 8, fontSize: 12,
                    background: "transparent", border: `1px solid ${theme.border}`,
                    color: theme.textMuted, cursor: "pointer",
                  }}
                >Cancel</button>
                <button
                  disabled={clearConfirmText !== "CONFIRM"}
                  onClick={async () => {
                    try {
                      const res = await fetch(`${API}/api/storage/clear-data`, {
                        method: "DELETE",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ confirm: "CONFIRM" }),
                      });
                      if (!res.ok) throw new Error(`HTTP ${res.status}`);
                      notify("All data cleared", "success");
                      setShowClearModal(false);
                      setCandidates([]);
                      setJobs([]);
                      setMatchResults([]);
                      loadStorageHealth();
                    } catch (e) {
                      notify("Failed to clear data", "error");
                    }
                  }}
                  style={{
                    padding: "8px 18px", borderRadius: 8, fontSize: 12, fontWeight: 600,
                    background: clearConfirmText === "CONFIRM" ? theme.danger : `${theme.danger}30`,
                    border: "none", color: "#fff", cursor: clearConfirmText === "CONFIRM" ? "pointer" : "not-allowed",
                    opacity: clearConfirmText === "CONFIRM" ? 1 : 0.5,
                  }}
                >Delete Everything</button>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* ─── Preview Modal ──────────────────────────────────────────── */}
      {previewModal && (
        <div
          style={{
            position: "fixed", inset: 0, zIndex: 2000,
            background: "rgba(0,0,0,0.8)", display: "flex",
            alignItems: "center", justifyContent: "center",
            animation: "fadeIn 0.2s ease",
          }}
          onClick={() => setPreviewModal(null)}
        >
          <div
            style={{
              background: theme.surface, borderRadius: 12, maxWidth: "90vw",
              maxHeight: "90vh", overflow: "hidden", display: "flex",
              flexDirection: "column", border: `1px solid ${theme.border}`,
              boxShadow: "0 25px 60px rgba(0,0,0,0.5)", minWidth: 400,
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div style={{
              padding: "14px 20px", borderBottom: `1px solid ${theme.border}`,
              display: "flex", alignItems: "center", gap: 12,
            }}>
              <span style={{ fontSize: 18 }}>{fileTypeIcon(previewModal.content_type)}</span>
              <div style={{ flex: 1, overflow: "hidden" }}>
                <div style={{ fontSize: 13, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {previewModal.original_name || previewModal.name || previewModal.filename}
                </div>
                <div style={{ fontSize: 10, color: theme.textDim }}>
                  {formatSize(previewModal.size)} · {previewModal.content_type}
                  {previewModal.email_subject && ` · ${previewModal.email_subject}`}
                </div>
              </div>
              <a
                href={`${API}${previewModal.download_url || `/api/attachments/${previewModal.filename}`}`}
                style={{
                  padding: "6px 12px", borderRadius: 6, fontSize: 11,
                  background: theme.accent, color: "#fff", textDecoration: "none",
                  fontWeight: 500,
                }}
              >
                Download
              </a>
              <button
                onClick={() => setPreviewModal(null)}
                style={{
                  background: "none", border: "none", color: theme.textMuted,
                  cursor: "pointer", fontSize: 18, padding: 4,
                }}
              >
                ✕
              </button>
            </div>

            {/* Modal Body */}
            <div style={{ flex: 1, overflow: "auto", padding: 20, display: "flex", alignItems: "center", justifyContent: "center", minHeight: 300 }}>
              {previewModal.content_type?.startsWith("image/") ? (
                <img
                  src={`${API}${previewModal.preview_url || `/api/attachments/${previewModal.filename}/preview`}`}
                  alt={previewModal.original_name || previewModal.name}
                  style={{ maxWidth: "100%", maxHeight: "70vh", objectFit: "contain", borderRadius: 4 }}
                />
              ) : previewModal.content_type === "application/pdf" ? (
                <iframe
                  src={`${API}${previewModal.preview_url || `/api/attachments/${previewModal.filename}/preview`}`}
                  style={{ width: "100%", height: "70vh", border: "none", borderRadius: 4, background: "#fff" }}
                  title="PDF Preview"
                />
              ) : (
                <div style={{ textAlign: "center", padding: 40 }}>
                  <div style={{ fontSize: 48, marginBottom: 16 }}>{fileTypeIcon(previewModal.content_type)}</div>
                  <div style={{ fontSize: 14, fontWeight: 500, marginBottom: 8 }}>
                    {previewModal.original_name || previewModal.name || previewModal.filename}
                  </div>
                  <div style={{ fontSize: 12, color: theme.textMuted, marginBottom: 20 }}>
                    {formatSize(previewModal.size)} · {previewModal.content_type}
                  </div>
                  <div style={{ fontSize: 11, color: theme.textDim, marginBottom: 20 }}>
                    Preview is not available for this file type.
                  </div>
                  <a
                    href={`${API}${previewModal.download_url || `/api/attachments/${previewModal.filename}`}`}
                    style={{
                      padding: "10px 24px", borderRadius: 8, fontSize: 13,
                      background: theme.accent, color: "#fff", textDecoration: "none",
                      fontWeight: 600,
                    }}
                  >
                    Download File
                  </a>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ─── Toast Notifications ─────────────────────────────────────── */}
      {toasts.length > 0 && (
        <div style={{
          position: "fixed", bottom: 20, right: 20, zIndex: 3000,
          display: "flex", flexDirection: "column-reverse", gap: 8,
          pointerEvents: "none",
        }}>
          {toasts.map((t) => {
            const borderColor = t.type === "new_high_fit" ? theme.success
              : t.type === "new_candidate" ? theme.accent : theme.textDim;
            const icon = t.type === "new_high_fit" ? "🎯"
              : t.type === "new_candidate" ? "👤" : "📧";
            return (
              <div
                key={t.id}
                style={{
                  background: theme.surface, border: `1px solid ${theme.border}`,
                  borderLeft: `4px solid ${borderColor}`,
                  borderRadius: 8, padding: "10px 14px",
                  boxShadow: "0 8px 30px rgba(0,0,0,0.4)",
                  display: "flex", alignItems: "center", gap: 10,
                  animation: "toastIn 0.3s ease",
                  minWidth: 260, maxWidth: 360, pointerEvents: "auto",
                }}
              >
                <span style={{ fontSize: 16, flexShrink: 0 }}>{icon}</span>
                <span style={{ fontSize: 12, color: theme.text, flex: 1 }}>{t.message}</span>
                <button
                  onClick={() => setToasts((prev) => prev.filter((x) => x.id !== t.id))}
                  style={{
                    background: "none", border: "none", color: theme.textDim,
                    cursor: "pointer", fontSize: 14, flexShrink: 0, padding: 2,
                  }}
                >
                  ✕
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
