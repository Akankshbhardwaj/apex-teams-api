import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";
import pkg from "pg";
const { Pool } = pkg;

dotenv.config();
const app = express();
app.use(express.json());

// 🗄️ PostgreSQL connection
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// ⚙️ Microsoft Identity Config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = "https://apex-teams-api.onrender.com/redirect";
const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/OnlineMeetings.ReadWrite",
  "offline_access"
];

const pca = new msal.ConfidentialClientApplication(msalConfig);

// 🌐 Step 1️⃣ Login (per user)
app.get("/login/:apex_user", async (req, res) => {
  const { apex_user } = req.params;
  const authCodeUrlParameters = {
    scopes: SCOPES,
    redirectUri: `${REDIRECT_URI}?apex_user=${encodeURIComponent(apex_user)}`
  };

  try {
    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("❌ Error generating auth URL:", err);
    res.status(500).send("Error generating auth URL");
  }
});

// 🌐 Step 2️⃣ Redirect — store tokens
app.get("/redirect", async (req, res) => {
  const { apex_user, code } = req.query;
  if (!code || !apex_user)
    return res.status(400).send("Missing authorization code or user");

  try {
    const response = await pca.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: `${REDIRECT_URI}?apex_user=${encodeURIComponent(apex_user)}`
    });

    const accessToken = response.accessToken;
    const refreshToken = response.refreshToken;
    const expiresAt = new Date(Date.now() + response.expiresIn * 1000);

    await pool.query(
      `INSERT INTO ms_graph_tokens (apex_user, access_token, refresh_token, expires_at)
       VALUES ($1,$2,$3,$4)
       ON CONFLICT (apex_user)
       DO UPDATE SET access_token=$2, refresh_token=$3, expires_at=$4`,
      [apex_user, accessToken, refreshToken, expiresAt]
    );

    console.log(`✅ Tokens saved for ${apex_user}`);
    res.send(`✅ ${apex_user} authenticated successfully!`);
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

// 🔄 Ensure valid token
async function ensureAccessToken(apex_user) {
  const result = await pool.query(
    "SELECT access_token, refresh_token, expires_at FROM ms_graph_tokens WHERE apex_user=$1",
    [apex_user]
  );
  if (result.rows.length === 0) throw new Error("User not authenticated");

  let { access_token, refresh_token, expires_at } = result.rows[0];

  if (new Date(expires_at) < new Date()) {
    console.log(`🔄 Refreshing token for ${apex_user}...`);
    const tokenResponse = await pca.acquireTokenByRefreshToken({
      refreshToken: refresh_token,
      scopes: SCOPES
    });
    access_token = tokenResponse.accessToken;
    refresh_token = tokenResponse.refreshToken;
    expires_at = new Date(Date.now() + tokenResponse.expiresIn * 1000);

    await pool.query(
      `UPDATE ms_graph_tokens SET access_token=$1, refresh_token=$2, expires_at=$3 WHERE apex_user=$4`,
      [access_token, refresh_token, expires_at, apex_user]
    );
    console.log(`✅ Token refreshed for ${apex_user}`);
  }

  return access_token;
}

// 📧 Send Mail (per APEX user)
app.post("/send-mail/:apex_user", async (req, res) => {
  const { apex_user } = req.params;
  try {
    const token = await ensureAccessToken(apex_user);

    const mail = {
      message: {
        subject: req.body.subject || "Mail from Oracle APEX",
        body: { contentType: "HTML", content: req.body.body || "APEX Mail" },
        toRecipients: (req.body.toEmails || []).map(email => ({
          emailAddress: { address: email }
        }))
      },
      saveToSentItems: "true"
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(mail)
    });

    const result = await response.text();
    if (!response.ok) return res.status(400).json({ error: "Mail failed", details: result });

    res.json({ success: true, message: "📧 Mail sent successfully!" });
  } catch (err) {
    res.status(401).json({ error: err.message });
  }
});

// 📅 Create Meeting (per APEX user)
app.post("/create-meeting/:apex_user", async (req, res) => {
  const { apex_user } = req.params;
  try {
    const token = await ensureAccessToken(apex_user);

    const attendees = (req.body.attendees || []).map(email => ({
      emailAddress: { address: email, name: email },
      type: "required"
    }));

    const event = {
      subject: req.body.subject,
      body: { contentType: "HTML", content: req.body.description },
      start: { dateTime: req.body.start, timeZone: "India Standard Time" },
      end: { dateTime: req.body.end, timeZone: "India Standard Time" },
      location: { displayName: req.body.location },
      attendees,
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(event)
    });

    const result = await response.json();
    if (!response.ok)
      return res.status(400).json({ error: "Failed to create event", details: result });

    res.json({
      success: true,
      eventId: result.id,
      joinUrl: result.onlineMeeting?.joinUrl || "No Teams link"
    });
  } catch (err) {
    res.status(401).json({ error: err.message });
  }
});

// 🚀 Start
app.listen(10000, () => console.log("🚀 Multi-user Teams API running on port 10000"));
