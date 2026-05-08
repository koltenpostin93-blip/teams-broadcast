const express = require("express");
const cors = require("cors");
const { Client, LocalAuth, MessageMedia } = require("whatsapp-web.js");
const qrcode = require("qrcode");

const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

const PORT = process.env.PORT || 3001;
const API_KEY = process.env.WA_API_KEY || "jpsi-wa-service";

// ── Auth middleware ────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  if (req.path === "/health") return next();
  const key = req.headers["x-api-key"];
  if (key !== API_KEY) return res.status(401).json({ error: "Unauthorized" });
  next();
});

// ── WhatsApp client ───────────────────────────────────────────────────────────
let qrCodeData = null;
let isReady = false;
let clientError = null;

const client = new Client({
  authStrategy: new LocalAuth({ dataPath: "./wa-session" }),
  puppeteer: {
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"],
  },
});

client.on("qr", async (qr) => {
  isReady = false;
  clientError = null;
  try {
    qrCodeData = await qrcode.toDataURL(qr);
    console.log("QR code ready — scan in the app.");
  } catch (err) {
    console.error("QR error:", err);
  }
});

client.on("ready", () => {
  isReady = true;
  qrCodeData = null;
  console.log("WhatsApp client ready.");
});

client.on("disconnected", (reason) => {
  isReady = false;
  clientError = reason;
  console.log("WhatsApp disconnected:", reason);
});

client.on("auth_failure", (msg) => {
  isReady = false;
  clientError = msg;
  console.error("Auth failure:", msg);
});

client.initialize();

// ── Routes ────────────────────────────────────────────────────────────────────

// Health check (no auth needed)
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

// WhatsApp connection status
app.get("/status", (req, res) => {
  res.json({
    ready: isReady,
    has_qr: !!qrCodeData,
    error: clientError,
  });
});

// Get QR code as base64 data URL
app.get("/qr", (req, res) => {
  if (isReady) return res.json({ ready: true, qr: null });
  if (!qrCodeData) return res.json({ ready: false, qr: null, message: "QR not generated yet, please wait..." });
  res.json({ ready: false, qr: qrCodeData });
});

// Get all chats
app.get("/chats", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp not ready" });
  try {
    const chats = await client.getChats();
    const result = chats
      .filter((c) => c.name)
      .map((c) => ({
        id: c.id._serialized,
        name: c.name,
        isGroup: c.isGroup,
        unreadCount: c.unreadCount,
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    res.json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Send message to one or more chats
app.post("/send", async (req, res) => {
  if (!isReady) return res.status(503).json({ error: "WhatsApp not ready" });

  const { chat_ids, message, image_base64, image_mime } = req.body;
  if (!chat_ids || !chat_ids.length) {
    return res.status(400).json({ error: "chat_ids required" });
  }

  const results = [];

  for (const chatId of chat_ids) {
    try {
      if (image_base64) {
        const media = new MessageMedia(
          image_mime || "image/png",
          image_base64,
          "image"
        );
        await client.sendMessage(chatId, media, {
          caption: message || undefined,
        });
      } else if (message) {
        await client.sendMessage(chatId, message);
      }
      results.push({ chatId, ok: true });
    } catch (err) {
      results.push({ chatId, ok: false, error: err.message });
    }
  }

  res.json({ results });
});

// Logout / reset session
app.post("/logout", async (req, res) => {
  try {
    await client.logout();
    isReady = false;
    qrCodeData = null;
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Start server ──────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`WhatsApp service running on port ${PORT}`);
});
