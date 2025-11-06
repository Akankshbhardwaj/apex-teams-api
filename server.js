import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
import { ConfidentialClientApplication } from "@azure/msal-node";

dotenv.config();
const app = express();
app.use(express.json());

// ===================
//  Microsoft Config
// ===================
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const REDIRECT_URI = "https://apex-teams-api.onrender.com/auth/callback";
const SCOPES = ["https://graph.microsoft.com/Mail.Send", "offline_access"];
const cca = new ConfidentialClientApplication(msalConfig);

let tokenResponse = null;

// ===================
//  1️⃣  Login endpoint
// ===================
app.get("/login", (req, res) => {
  const authCodeUrlParameters = {
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  cca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  });
});

// ===================
//  2️⃣  Callback endpoint
// ===================
app.get("/auth/callback", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    tokenResponse = await cca.acquireTokenByCode(tokenRequest);
    res.send("✅ Authentication successful! You can now send emails.");
  } catch (error) {
    console.error("Error acquiring token:", error);
    res.status(500).send("Error during authentication.");
  }
});

// ===================
//  3️⃣  Send Mail endpoint
// ===================
app.post("/send-mail", async (req, res) => {
  if (!tokenResponse) {
    return res.status(401).json({ error: "User not authenticated yet. Visit /login first." });
  }

  const { to, subject, body } = req.body;

  const emailData = {
    message: {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body
      },
      toRecipients: [
        {
          emailAddress: { address: to }
        }
      ]
    },
    saveToSentItems: "true"
  };

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${tokenResponse.accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(emailData)
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ error: "Mail send failed", details: err });
    }

    res.json({ success: true, message: "Email sent successfully 🚀" });
  } catch (err) {
    console.error("Error sending mail:", err);
    res.status(500).json({ error: "Internal server error" });
  }
});

app.listen(3000, () => console.log("✅ APEX → Microsoft Graph (Delegated) running on port 3000"));
