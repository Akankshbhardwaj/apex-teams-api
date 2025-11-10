// server.js
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
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = "https://apex-teams-api.onrender.com/redirect"; // Must match Azure redirect URL exactly

// ✅ Scopes needed for Mail, Calendar & Teams meetings
const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/OnlineMeetings.ReadWrite",
  "offline_access"
];

const pca = new msal.ConfidentialClientApplication(msalConfig);
let accessToken = null;

// 🌐 Root route
app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph API server is running. Visit /login to authenticate.");
});

// Step 1️⃣: Login - Generate Microsoft OAuth URL
app.get("/login", async (req, res) => {
  const authCodeUrlParameters = {
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("❌ Error generating auth URL:", err);
    res.status(500).send("Error generating auth URL");
  }
});

// Step 2️⃣: Redirect from Microsoft - Exchange code for access token
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  if (!code) {
    console.error("❌ Missing authorization code in redirect");
    return res.status(400).send("Error: Missing authorization code. Retry /login.");
  }

  const tokenRequest = {
    code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);
    accessToken = response.accessToken;
    console.log("✅ Access token acquired successfully!");
    res.send("✅ Authentication successful! You can now send emails and create meetings!");
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

// Step 3️⃣: Send Mail
app.post("/send-mail", async (req, res) => {
  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  const mail = {
    message: {
      subject: req.body.subject || "Hello from Oracle APEX + Microsoft Graph",
      body: {
        contentType: "HTML",
        content: req.body.body || "<p>This email was sent via Microsoft Graph API!</p>"
      },
      toRecipients: (req.body.toEmails || []).map(email => ({
        emailAddress: { address: email }
      }))
    },
    saveToSentItems: "true"
  };

  try {
    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(mail)
    });

    if (!graphResponse.ok) {
      const errText = await graphResponse.text();
      return res.status(400).json({ error: "Mail send failed", details: errText });
    }

    res.json({ success: true, message: "📧 Mail sent successfully!" });
  } catch (err) {
    console.error("Error sending mail:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// Step 4️⃣: Create Meeting (Multiple attendees + Teams join link)
app.post("/create-meeting", async (req, res) => {
  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  // Extract attendees array from request
  const attendees = (req.body.attendees || []).map(email => ({
    emailAddress: { address: email, name: email },
    type: "required"
  }));

  const event = {
    subject: req.body.subject || "Meeting from Oracle APEX",
    body: {
      contentType: "HTML",
      content: req.body.description || "Meeting scheduled via Oracle APEX"
    },
    start: {
      dateTime: req.body.start,
      timeZone: "India Standard Time"
    },
    end: {
      dateTime: req.body.end,
      timeZone: "India Standard Time"
    },
    location: {
      displayName: req.body.location || "Online"
    },
    attendees: attendees,
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness"
  };

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(event)
    });

    const result = await response.json();

    if (!response.ok) {
      return res.status(400).json({ error: "Failed to create event", details: result });
    }

    res.json({
      success: true,
      message: "📅 Meeting created successfully!",
      eventId: result.id,
      joinUrl: result.onlineMeeting?.joinUrl || "No Teams join link available"
    });
  } catch (err) {
    console.error("Error creating event:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// 🚀 Start server
app.listen(10000, () => console.log("🚀 Server running on port 10000"));
