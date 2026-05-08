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
const SESSION_DIR = path.join(__dirname, "wa-session");

// Silent logger
const logger = pino({ level: "silent" });

// ── State ─────────────────────────────────────────────────────────────────────
let sock = null;
let qrCodeData = null;
let isReady = false;
let chatCache = [];

// ── Auth middleware ───────────────────────────────────────────────────────────
app.use((req, res, next) => {
  if (req.path === "/health") return next();
  if (req.headers["x-api-key"] !== API_KEY)
    return res.status(401).json({ error: "Unauthorized" });
  next();
});

// ── WhatsApp connection ───────────────────────────────────────────────────────
async function connectWA() {
  if (!fs.existsSync(SESSION_DIR)) fs.mkdirSync(SESSION_DIR, { recursive: true });

  const { state, saveCreds } = await useMultiFileAuthState(SESSION_DIR);
  const { version } = await fetchLatestBaileysVersion();

  sock = makeWASocket({
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

  sock.ev.on("creds.update", saveCreds);

  sock.ev.on("connection.update", async (update) => {
    const { connection, lastDisconnect, qr } = update;

    if (qr) {
      console.log("QR code ready — scan in the app.");
      qrCodeData = await qrcode.toDataURL(qr);
      isReady = false;
    }

    if (connection === "open") {
      console.log("WhatsApp connected.");
      isReady = true;
      qrCodeData = null;
      // Pre-load chats
      try {
        const chats = await sock.groupFetchAllParticipating();
        chatCache = Object.values(chats).map((g) => ({
          id: g.id,
          name: g.subject,
          isGroup: true,
        }));
      } catch (_) {}
    }

    if (connection === "close") {
      isReady = false;
      const code = lastDisconnect?.error?.output?.statusCode;
      const shouldReconnect = code !== DisconnectReason.loggedOut;
      console.log("Connection closed. Reconnect:", shouldReconnect, "Code:", code);
      if (shouldReconnect) {
        setTimeout(connectWA, 5000);
      } else {
        // Logged out — clear session
        fs.rmSync(SESSION_DIR, { recursive: true, force: true });
        qrCodeData = null;
        isReady = false;
        setTimeout(connectWA, 2000);
      }
    }
  });

  // Cache DM chats when they come in
  sock.ev.on("chats.set", ({ chats }) => {
    const newChats = chats
      .filter((c) => c.name)
      .map((c) => ({
        id: c.id,
        name: c.name,
        isGroup: c.id.endsWith("@g.us"),
      }));
    // Merge with existing
    const ids = new Set(chatCache.map((c) => c.id));
    newChats.forEach((c) => { if (!ids.has(c.id)) chatCache.push(c); });
  });
}

connectWA().catch(console.error);

// ── Routes ────────────────────────────────────────────────────────────────────

app.get("/health", (req, res) => res.json({ status: "ok" }));

app.get("/status", (req, res) => {
  res.json({ ready: isReady, has_qr: !!qrCodeData, error: null });
});

app.get("/qr", (req, res) => {
  if (isReady) return res.json({ ready: true, qr: null });
  if (!qrCodeData) return res.json({ ready: false, qr: null, message: "Generating QR code, please wait..." });
  res.json({ ready: false, qr: qrCodeData });
});

app.get("/chats", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp not ready" });
  try {
    // Refresh group list
    const groups = await sock.groupFetchAllParticipating();
    const groupList = Object.values(groups).map((g) => ({
      id: g.id,
      name: g.subject,
      isGroup: true,
    }));
    // Merge with DM cache
    const groupIds = new Set(groupList.map((g) => g.id));
    const dms = chatCache.filter((c) => !groupIds.has(c.id) && !c.isGroup);
    const all = [...groupList, ...dms].sort((a, b) => a.name.localeCompare(b.name));
    res.json(all);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/send", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp not ready" });

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
        // Send first image with caption, subsequent images bare
        for (let i = 0; i < imgList.length; i++) {
          const buf = Buffer.from(imgList[i].base64, "base64");
          const mime = imgList[i].mime || "image/png";
          await sock.sendMessage(chatId, {
            image: buf,
            mimetype: mime,
            caption: i === 0 ? (message || undefined) : undefined,
          });
        }
      } else if (message) {
        await sock.sendMessage(chatId, { text: message });
      }
      results.push({ chatId, ok: true });
    } catch (err) {
      results.push({ chatId, ok: false, error: err.message });
    }
  }
  res.json({ results });
});

app.post("/logout", async (req, res) => {
  try {
    await sock.logout();
    fs.rmSync(SESSION_DIR, { recursive: true, force: true });
    isReady = false;
    qrCodeData = null;
    res.json({ ok: true });
    setTimeout(connectWA, 2000);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => console.log(`WhatsApp service running on port ${PORT}`));
