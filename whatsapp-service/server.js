const express = require("express");
const cors = require("cors");
const qrcode = require("qrcode");
const pino = require("pino");
const path = require("path");
const fs = require("fs");

const {
  default: makeWASocket,
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
  makeCacheableSignalKeyStore,
} = require("@whiskeysockets/baileys");

const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

const PORT = process.env.PORT || 3001;
const API_KEY = process.env.WA_API_KEY || "jpsi-wa-service";
const SESSIONS_DIR = path.join(__dirname, "wa-session");

// Silent logger
const logger = pino({ level: "silent" });

// ── Sessions map: sessionId -> { sock, isReady, qrCodeData, chatCache } ───────
const sessions = new Map();

// ── Auth middleware ───────────────────────────────────────────────────────────
app.use((req, res, next) => {
  if (req.path === "/health") return next();
  if (req.headers["x-api-key"] !== API_KEY)
    return res.status(401).json({ error: "Unauthorized" });
  next();
});

// ── Get session ID from request ───────────────────────────────────────────────
function sessionId(req) {
  return (req.headers["x-session-id"] || req.query.session || "default")
    .replace(/[^a-zA-Z0-9_\- ]/g, "_"); // sanitize for filesystem
}

// ── Connect a session ─────────────────────────────────────────────────────────
async function connectSession(sid) {
  // If already connecting/connected, skip
  if (sessions.has(sid) && sessions.get(sid)._connecting) return;

  const SESSION_PATH = path.join(SESSIONS_DIR, sid);
  if (!fs.existsSync(SESSION_PATH)) fs.mkdirSync(SESSION_PATH, { recursive: true });

  // Init or reuse session state
  if (!sessions.has(sid)) {
    sessions.set(sid, { sock: null, isReady: false, qrCodeData: null, chatCache: [], _connecting: true });
  } else {
    sessions.get(sid)._connecting = true;
  }
  const session = sessions.get(sid);

  const { state, saveCreds } = await useMultiFileAuthState(SESSION_PATH);
  const { version } = await fetchLatestBaileysVersion();

  const sock = makeWASocket({
    version,
    auth: {
      creds: state.creds,
      keys: makeCacheableSignalKeyStore(state.keys, logger),
    },
    logger,
    printQRInTerminal: false,
    browser: ["Teams Broadcast", "Chrome", "1.0"],
    generateHighQualityLinkPreview: false,
  });

  session.sock = sock;
  session._connecting = false;

  sock.ev.on("creds.update", saveCreds);

  sock.ev.on("connection.update", async (update) => {
    const { connection, lastDisconnect, qr } = update;

    if (qr) {
      console.log(`[${sid}] QR code ready.`);
      session.qrCodeData = await qrcode.toDataURL(qr);
      session.isReady = false;
    }

    if (connection === "open") {
      console.log(`[${sid}] WhatsApp connected.`);
      session.isReady = true;
      session.qrCodeData = null;
      try {
        const groups = await sock.groupFetchAllParticipating();
        session.chatCache = Object.values(groups).map((g) => ({
          id: g.id,
          name: g.subject,
          isGroup: true,
        }));
      } catch (_) {}
    }

    if (connection === "close") {
      session.isReady = false;
      const code = lastDisconnect?.error?.output?.statusCode;
      const shouldReconnect = code !== DisconnectReason.loggedOut;
      console.log(`[${sid}] Connection closed. Reconnect:`, shouldReconnect, "Code:", code);
      if (shouldReconnect) {
        setTimeout(() => connectSession(sid), 5000);
      } else {
        // Logged out — clear session files
        fs.rmSync(SESSION_PATH, { recursive: true, force: true });
        session.qrCodeData = null;
        session.isReady = false;
        session.chatCache = [];
        setTimeout(() => connectSession(sid), 2000);
      }
    }
  });

  // Cache DM chats
  sock.ev.on("chats.set", ({ chats }) => {
    const newChats = chats
      .filter((c) => c.name)
      .map((c) => ({
        id: c.id,
        name: c.name,
        isGroup: c.id.endsWith("@g.us"),
      }));
    const ids = new Set(session.chatCache.map((c) => c.id));
    newChats.forEach((c) => { if (!ids.has(c.id)) session.chatCache.push(c); });
  });
}

// ── Auto-restore saved sessions on startup ────────────────────────────────────
if (fs.existsSync(SESSIONS_DIR)) {
  fs.readdirSync(SESSIONS_DIR).forEach((dir) => {
    const full = path.join(SESSIONS_DIR, dir);
    if (fs.statSync(full).isDirectory()) {
      console.log(`Restoring session: ${dir}`);
      connectSession(dir).catch(console.error);
    }
  });
}

// ── Routes ────────────────────────────────────────────────────────────────────

app.get("/health", (req, res) => res.json({ status: "ok" }));

// Ensure a session exists (creates it if not), returns its state
app.get("/status", async (req, res) => {
  const sid = sessionId(req);
  if (!sessions.has(sid)) {
    connectSession(sid).catch(console.error);
    return res.json({ ready: false, has_qr: false, initializing: true, error: null });
  }
  const s = sessions.get(sid);
  res.json({ ready: s.isReady, has_qr: !!s.qrCodeData, error: null });
});

app.get("/qr", (req, res) => {
  const sid = sessionId(req);
  if (!sessions.has(sid)) {
    connectSession(sid).catch(console.error);
    return res.json({ ready: false, qr: null, message: "Initializing, please wait..." });
  }
  const s = sessions.get(sid);
  if (s.isReady) return res.json({ ready: true, qr: null });
  if (!s.qrCodeData) return res.json({ ready: false, qr: null, message: "Generating QR code, please wait..." });
  res.json({ ready: false, qr: s.qrCodeData });
});

app.get("/chats", async (req, res) => {
  const sid = sessionId(req);
  const s = sessions.get(sid);
  if (!s || !s.isReady) return res.status(503).json({ error: "WhatsApp not ready" });
  try {
    const groups = await s.sock.groupFetchAllParticipating();
    const groupList = Object.values(groups).map((g) => ({
      id: g.id,
      name: g.subject,
      isGroup: true,
    }));
    const groupIds = new Set(groupList.map((g) => g.id));
    const dms = s.chatCache.filter((c) => !groupIds.has(c.id) && !c.isGroup);
    const all = [...groupList, ...dms].sort((a, b) => a.name.localeCompare(b.name));
    res.json(all);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/send", async (req, res) => {
  const sid = sessionId(req);
  const s = sessions.get(sid);
  if (!s || !s.isReady) return res.status(503).json({ error: "WhatsApp not ready" });

  const { chat_ids, message, image_base64, image_mime, images } = req.body;
  if (!chat_ids?.length) return res.status(400).json({ error: "chat_ids required" });

  const results = [];
  for (const chatId of chat_ids) {
    try {
      const imgList = images && images.length > 0
        ? images
        : image_base64
          ? [{ base64: image_base64, mime: image_mime || "image/png" }]
          : [];

      if (imgList.length > 0) {
        for (let i = 0; i < imgList.length; i++) {
          const buf = Buffer.from(imgList[i].base64, "base64");
          const mime = imgList[i].mime || "image/png";
          await s.sock.sendMessage(chatId, {
            image: buf,
            mimetype: mime,
            caption: i === 0 ? (message || undefined) : undefined,
          });
        }
      } else if (message) {
        await s.sock.sendMessage(chatId, { text: message });
      }
      results.push({ chatId, ok: true });
    } catch (err) {
      results.push({ chatId, ok: false, error: err.message });
    }
  }
  res.json({ results });
});

app.post("/logout", async (req, res) => {
  const sid = sessionId(req);
  const s = sessions.get(sid);
  if (!s) return res.json({ ok: true });
  try {
    await s.sock.logout();
  } catch (_) {}
  const SESSION_PATH = path.join(SESSIONS_DIR, sid);
  fs.rmSync(SESSION_PATH, { recursive: true, force: true });
  sessions.delete(sid);
  res.json({ ok: true });
  // Recreate fresh session (will generate new QR)
  setTimeout(() => connectSession(sid), 2000);
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => console.log(`WhatsApp service running on port ${PORT}`));
