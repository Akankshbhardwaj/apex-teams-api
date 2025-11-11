import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";
import oracledb from "oracledb";

dotenv.config();
const app = express();
app.use(express.json());

// 🗄️ Oracle DB connection
const dbConfig = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  connectString: process.env.DB_CONNECT
};

// ⚙️ Microsoft Config
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

async function getConnection() {
  return await oracledb.getConnection(dbConfig);
}

//
// 🔹 STEP 1: Microsoft Login (per APEX user)
//
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

//
// 🔹 STEP 2: Redirect — Save Tokens in Oracle
//
app.get("/redirect", async (req, res) => {
  const { apex_user, code } = req.query;
  if (!code || !apex_user) {
    return res.status(400).send("Missing authorization code or user");
  }

  try {
    const response = await pca.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: `${REDIRECT_URI}?apex_user=${encodeURIComponent(apex_user)}`
    });

    const conn = await getConnection();
    const expiresAt = new Date(Date.now() + response.expiresIn * 1000);

    await conn.execute(
      `
      MERGE INTO MS_GRAPH_TOKENS t
      USING (SELECT :user AS apex_user FROM dual) src
      ON (t.apex_user = src.apex_user)
      WHEN MATCHED THEN UPDATE SET
          access_token = :access_token,
          refresh_token = :refresh_token,
          expires_at = :expires_at
      WHEN NOT MATCHED THEN INSERT
          (apex_user, access_token, refresh_token, expires_at)
          VALUES (:user, :access_token, :refresh_token, :expires_at)
      `,
      {
        user: apex_user,
        access_token: response.accessToken,
        refresh_token: response.refreshToken,
        expires_at: expiresAt
      },
      { autoCommit: true }
    );

    await conn.close();

    console.log(`✅ Tokens saved for ${apex_user}`);
    res.send(`<h3>✅ ${apex_user} connected successfully to Microsoft!</h3>`);
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

//
// 🔹 Utility — Ensure Valid Token
//
async function ensureAccessToken(apex_user) {
  const conn = await getConnection();
  const result = await conn.execute(
    `SELECT access_token, refresh_token, expires_at FROM MS_GRAPH_TOKENS WHERE apex_user = :user`,
    [apex_user]
  );
  await conn.close();

  if (result.rows.length === 0) throw new Error("User not authenticated");
  let [access_token, refresh_token, expires_at] = result.rows[0];

  if (new Date(expires_at) < new Date()) {
    console.log(`🔄 Refreshing token for ${apex_user}...`);
    const tokenResponse = await pca.acquireTokenByRefreshToken({
      refreshToken: refresh_token,
      scopes: SCOPES
    });

    const conn2 = await getConnection();
    const newExpires = new Date(Date.now() + tokenResponse.expiresIn * 1000);

    await conn2.execute(
      `
      UPDATE MS_GRAPH_TOKENS
      SET access_token = :a, refresh_token = :r, expires_at = :e
      WHERE apex_user = :u
      `,
      {
        a: tokenResponse.accessToken,
        r: tokenResponse.refreshToken,
        e: newExpires,
        u: apex_user
      },
      { autoCommit: true }
    );
    await conn2.close();
    console.log(`✅ Token refreshed for ${apex_user}`);
    return tokenResponse.accessToken;
  }

  return access_token;
}

//
// 🔹 Send Email per user
//
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

//
// 🔹 Create Teams Meeting per user
//
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

//
// 🚀 Start server
//
app.listen(10000, () => console.log("🚀 Oracle APEX Teams API running on port 10000"));
