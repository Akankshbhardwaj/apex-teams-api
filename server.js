// server.js
import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();
const app = express();
app.use(express.json());

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = "https://apex-teams-api.onrender.com/redirect"; // Must match exactly in Azure
const SCOPES = ["https://graph.microsoft.com/.default"];

const pca = new msal.ConfidentialClientApplication(msalConfig);
let accessToken = null;

app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph Mail API is running. Visit /login to authenticate.");
});

// Step 1: Login
app.get("/login", async (req, res) => {
  const authCodeUrlParameters = {
    scopes: ["https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read"],
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

// Step 2: Redirect from Microsoft
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  if (!code) {
    console.error("❌ Missing authorization code in redirect");
    return res.status(400).send("Error: Missing authorization code in redirect. Please retry /login.");
  }

  const tokenRequest = {
    code,
    scopes: ["https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read"],
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);
    accessToken = response.accessToken;
    console.log("✅ Access token acquired successfully!");
    res.send("✅ Authentication successful! You can now send emails via POST /send-mail");
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

// Step 3: Send Mail
app.post("/send-mail", async (req, res) => {
  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  const mail = {
    message: {
      subject: "Hello from Render + Microsoft Graph",
      body: { contentType: "Text", content: "This email was sent via Microsoft Graph API!" },
      toRecipients: [
        {
          emailAddress: {
            address: req.body.to || "your-email@faramond.in"
          }
        }
      ]
    }
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

    res.json({ success: true, message: "Mail sent successfully!" });
  } catch (err) {
    console.error("Error sending mail:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// ✅ Step 4: Create Meeting (ADD THIS HERE)
app.post("/create-meeting", async (req, res) => {
  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  const { subject, start, end, location, description, attendee } = req.body;

  const event = {
    subject,
    body: {
      contentType: "HTML",
      content: description || "Meeting scheduled via Oracle APEX"
    },
    start: {
      dateTime: start,
      timeZone: "India Standard Time"
    },
    end: {
      dateTime: end,
      timeZone: "India Standard Time"
    },
    location: {
      displayName: location || "Online"
    },
    attendees: [
      {
        emailAddress: {
          address: attendee,
          name: attendee
        },
        type: "required"
      }
    ]
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

    res.json({ success: true, message: "Meeting created successfully!", eventId: result.id });
  } catch (err) {
    console.error("Error creating event:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// ✅ Start server (keep this at the very end)
app.listen(10000, () => console.log("🚀 Server running on port 10000"));
