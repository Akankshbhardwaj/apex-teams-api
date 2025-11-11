import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();
const app = express();
app.use(express.json());

// ⚙️ Microsoft Identity Config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const REDIRECT_URI = process.env.REDIRECT_URI || "https://apex-teams-api.onrender.com/redirect";

const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/OnlineMeetings.ReadWrite",
  "offline_access",
];

const pca = new msal.ConfidentialClientApplication(msalConfig);

// 🧠 Store tokens for each logged-in user (in-memory, can later move to DB)
let userTokens = {};

// 🌐 Root route
app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph API is running. Visit /login?userEmail=your@email.com to authenticate.");
});


// Step 1️⃣: Login — Generate Microsoft OAuth URL (no query in redirect)
app.get("/login", async (req, res) => {
  const userEmail = req.query.userEmail;
  if (!userEmail) return res.status(400).send("❌ Missing userEmail in query string.");

  try {
    // ✅ use `state` instead of query param to safely carry the user email
    const authCodeUrlParameters = {
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
      state: userEmail
    };

    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("❌ Error generating auth URL:", err);
    res.status(500).send("Error generating auth URL");
  }
});


// Step 2️⃣: Redirect — Exchange code for tokens (email comes from `state`)
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  const userEmail = req.query.state; // ✅ email passed via state

  if (!code || !userEmail) {
    return res.status(400).send("❌ Missing authorization code or user context (state).");
  }

  const tokenRequest = {
    code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);

    userTokens[userEmail] = {
      accessToken: response.accessToken,
      refreshToken: response.refreshToken,
      expiresOn: response.expiresOn,
      account: response.account
    };

    console.log(`✅ Token acquired for ${userEmail}`);
    res.send(`✅ ${userEmail} authenticated successfully! You can now create meetings and send emails.`);
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});


// ♻️ Utility — Refresh or retrieve existing access token
async function getAccessToken(userEmail) {
  const tokenData = userTokens[userEmail];
  if (!tokenData) return null;

  try {
    const refreshed = await pca.acquireTokenSilent({
      scopes: SCOPES,
      account: tokenData.account,
    });
    userTokens[userEmail].accessToken = refreshed.accessToken;
    return refreshed.accessToken;
  } catch (err) {
    console.error(`🔁 Token refresh failed for ${userEmail}:`, err.message);
    return null;
  }
}


// Step 3️⃣: Send Email
app.post("/send-mail", async (req, res) => {
  const senderEmail = req.body.sender_email;
  if (!senderEmail)
    return res.status(400).json({ error: "Missing sender_email in body." });

  let accessToken = await getAccessToken(senderEmail);
  if (!accessToken)
    return res.status(401).json({
      error: `User ${senderEmail} not authenticated. Visit /login?userEmail=${senderEmail}`,
    });

  const mail = {
    message: {
      subject: req.body.subject || "Hello from Oracle APEX + Microsoft Graph",
      body: {
        contentType: "HTML",
        content: req.body.body || "<p>This email was sent via Microsoft Graph API!</p>",
      },
      toRecipients: (req.body.toEmails || []).map((email) => ({
        emailAddress: { address: email },
      })),
    },
    saveToSentItems: "true",
  };

  try {
    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(mail),
    });

    const text = await graphResponse.text();
    if (!graphResponse.ok) {
      return res.status(400).json({ error: "Mail send failed", details: text });
    }

    res.json({ success: true, message: "📧 Mail sent successfully!" });
  } catch (err) {
    console.error("Error sending mail:", err);
    res.status(500).json({ error: "Internal server error", details: err.message });
  }
});


// Step 4️⃣: Create Meeting
app.post("/create-meeting", async (req, res) => {
  const senderEmail = req.body.sender_email;
  if (!senderEmail)
    return res.status(400).json({ error: "Missing sender_email in body." });

  let accessToken = await getAccessToken(senderEmail);
  if (!accessToken)
    return res.status(401).json({
      error: `User ${senderEmail} not authenticated. Visit /login?userEmail=${senderEmail}`,
    });

  const attendees = (req.body.attendees || []).map((email) => ({
    emailAddress: { address: email, name: email },
    type: "required",
  }));

  const event = {
    subject: req.body.subject || "Meeting from Oracle APEX",
    body: {
      contentType: "HTML",
      content: req.body.description || "Meeting scheduled via Oracle APEX",
    },
    start: { dateTime: req.body.start, timeZone: "India Standard Time" },
    end: { dateTime: req.body.end, timeZone: "India Standard Time" },
    location: { displayName: req.body.location || "Online" },
    attendees,
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(event),
    });

    const result = await response.json();
    if (!response.ok) {
      return res.status(400).json({ error: "Failed to create meeting", details: result });
    }

    res.json({
      success: true,
      message: "📅 Meeting created successfully!",
      eventId: result.id,
      joinUrl: result.onlineMeeting?.joinUrl || "No Teams join link available",
    });
  } catch (err) {
    console.error("Error creating event:", err);
    res.status(500).json({ error: "Internal server error", details: err.message });
  }
});


// 🚀 Start Server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
