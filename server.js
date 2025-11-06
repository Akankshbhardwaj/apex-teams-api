import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = "https://apex-teams-api.onrender.com/redirect"; // must match Azure exactly
const SCOPES = ["https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read"];

const pca = new msal.ConfidentialClientApplication(msalConfig);

let accessToken = null;

app.get("/login", async (req, res) => {
  const authCodeUrlParameters = {
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };
  try {
    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("Error generating auth URL:", err);
    res.status(500).send("Error generating auth URL");
  }
});

app.get("/redirect", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);
    accessToken = response.accessToken;
    console.log("✅ Access token acquired successfully!");
    res.send("✅ Authentication successful! You can now send emails via /send-mail");
  } catch (err) {
    console.error("Error acquiring token:", err);
    res.status(500).send("Error acquiring token: " + err);
  }
});

app.post("/send-mail", async (req, res) => {
  if (!accessToken) return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });

  const mail = {
    message: {
      subject: "Hello from Render + APEX",
      body: { contentType: "Text", content: "This email was sent using Microsoft Graph API!" },
      toRecipients: [{ emailAddress: { address: req.body.to || "your-email@faramond.in" } }]
    }
  };

  try {
    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
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

app.listen(10000, () => console.log("Server running on port 10000 🚀"));
