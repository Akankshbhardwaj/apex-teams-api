import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config(); // Load environment variables

const app = express();
app.use(express.json());

// 🟩 MSAL configuration
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

// 🟩 Redirect URI — must exactly match Azure’s “Redirect URIs” entry
const REDIRECT_URI = "https://apex-teams-api.onrender.com/redirect";

// 🟩 Scopes (use /.default for app permissions)
const SCOPES = ["https://graph.microsoft.com/.default"];

const pca = new msal.ConfidentialClientApplication(msalConfig);

let accessToken = null;

// 🟢 Step 1: Login route
app.get("/login", async (req, res) => {
  const authCodeUrlParameters = {
    scopes: ["User.Read", "Mail.Send"],
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

// 🟢 Step 2: Redirect handler
app.get("/redirect", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: ["User.Read", "Mail.Send"],
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);
    accessToken = response.accessToken;
    console.log("✅ Access token acquired successfully!");
    res.send("✅ Authentication successful! You can now send emails via /send-mail");
  } catch (err) {
    console.error("❌ Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err.message);
  }
});

// 🟢 Step 3: Send mail
app.post("/send-mail", async (req, res) => {
  if (!accessToken)
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  const mail = {
    message: {
      subject: req.body.subject || "Hello from Render + APEX 🚀",
      body: {
        contentType: "Text",
        content: req.body.body || "This email was sent using Microsoft Graph API!"
      },
      toRecipients: [
        { emailAddress: { address: req.body.to || "akanksh@faramond.in" } }
      ]
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
      console.error("❌ Graph API error:", errText);
      return res.status(400).json({ error: "Mail send failed", details: errText });
    }

    console.log("✅ Mail sent successfully!");
    res.json({ success: true, message: "Mail sent successfully!" });
  } catch (err) {
    console.error("❌ Error sending mail:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

// 🟢 Root route
app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph API - APEX Bridge is running 🚀");
});

// 🟢 Start server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on port ${PORT} 🚀`));
