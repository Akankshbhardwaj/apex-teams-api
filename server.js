import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();
const app = express();
app.use(express.json());

// 🔹 Microsoft Identity Config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = process.env.REDIRECT_URI || "https://apex-teams-api.onrender.com/redirect";
const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/OnlineMeetings.ReadWrite",
  "offline_access"
];

const pca = new msal.ConfidentialClientApplication(msalConfig);

// 🧠 Store tokens per user
let users = {}; // { username: { accessToken, refreshToken, expiresOn } }

// 🏠 Root route
app.get("/", (req, res) => {
  res.send("✅ APEX Teams API is running. Visit /login/{username} to authenticate.");
});

// 1️⃣ LOGIN
app.get("/login/:username", async (req, res) => {
  const username = req.params.username;
  const authCodeUrlParameters = {
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    state: username
  };

  try {
    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("❌ Error generating auth URL:", err);
    res.status(500).send("Error generating auth URL");
  }
});

// 2️⃣ REDIRECT (Token exchange)
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  const username = req.query.state || "default";

  if (!code) return res.status(400).send("Missing authorization code");

  const tokenRequest = {
    code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);

    users[username] = {
      accessToken: response.accessToken,
      refreshToken: response.refreshToken,
      expiresOn: response.expiresOn
    };

    console.log(`✅ User ${username} authenticated successfully!`);
    res.send(`✅ Authentication successful for ${username}! You can now create meetings.`);
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

// 🧰 Helper: get valid token (auto-refresh)
async function getValidToken(username) {
  const user = users[username];
  if (!user) return null;

  const now = new Date();
  if (user.expiresOn && user.expiresOn > now) {
    return user.accessToken;
  }

  if (!user.refreshToken) return null;

  try {
    const tokenResponse = await pca.acquireTokenByRefreshToken({
      refreshToken: user.refreshToken,
      scopes: SCOPES
    });

    users[username] = {
      accessToken: tokenResponse.accessToken,
      refreshToken: tokenResponse.refreshToken || user.refreshToken,
      expiresOn: tokenResponse.expiresOn
    };

    console.log(`🔁 Token refreshed for ${username}`);
    return tokenResponse.accessToken;
  } catch (err) {
    console.error(`❌ Token refresh failed for ${username}:`, err);
    return null;
  }
}

// 3️⃣ Create Meeting (per-user)
app.post("/create-meeting/:username", async (req, res) => {
  const username = req.params.username;
  const accessToken = await getValidToken(username);

  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login/" + username });

  const attendees = (req.body.attendees || []).map(email => ({
    emailAddress: { address: email, name: email },
    type: "required"
  }));

  const event = {
    subject: req.body.subject || "Meeting from Oracle APEX",
    body: { contentType: "HTML", content: req.body.description || "Meeting scheduled via APEX" },
    start: { dateTime: req.body.start, timeZone: "India Standard Time" },
    end: { dateTime: req.body.end, timeZone: "India Standard Time" },
    location: { displayName: req.body.location || "Online" },
    attendees: attendees,
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness"
  };

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(event)
    });

    const result = await response.json();
    if (!response.ok) return res.status(400).json({ error: "Failed to create event", details: result });

    res.json({
      success: true,
      eventId: result.id,
      joinUrl: result.onlineMeeting?.joinUrl || "No Teams join link available"
    });
  } catch (err) {
    console.error("Error creating event:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// 4️⃣ Send Mail (per-user)
app.post("/send-mail/:username", async (req, res) => {
  const username = req.params.username;
  const accessToken = await getValidToken(username);

  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login/" + username });

  const mail = {
    message: {
      subject: req.body.subject || "Hello from Oracle APEX",
      body: { contentType: "HTML", content: req.body.body || "<p>This email was sent via Graph API!</p>" },
      toRecipients: (req.body.toEmails || []).map(email => ({
        emailAddress: { address: email }
      }))
    },
    saveToSentItems: "true"
  };

  try {
    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(mail)
    });

    const result = await graphResponse.text();
    if (!graphResponse.ok)
      return res.status(400).json({ error: "Mail send failed", details: result });

    res.json({ success: true, message: "📧 Mail sent successfully!" });
  } catch (err) {
    console.error("Error sending mail:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// 🚀 Start server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
