import express from "express";
import fetch from "node-fetch";
import bodyParser from "body-parser";

const app = express();
app.use(bodyParser.json());

const {
  CLIENT_ID,
  CLIENT_SECRET,
  TENANT_ID,
  GRAPH_SCOPE,
  EMAIL_USER
} = process.env;

// Optional: for debugging — only in dev
console.log("Loaded ENV:", { CLIENT_ID, TENANT_ID, EMAIL_USER });

let cachedToken = null;
let tokenExpiry = null;

// 🔐 Get access token dynamically from Azure
async function getAccessToken() {
  if (cachedToken && new Date() < tokenExpiry) {
    return cachedToken;
  }

  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: GRAPH_SCOPE,
      grant_type: "client_credentials"
    }),
  });

  const data = await res.json();

  if (!res.ok) {
    console.error("❌ Failed to get token:", data);
    throw new Error(data.error_description || "Unable to fetch access token");
  }

  cachedToken = data.access_token;
  tokenExpiry = new Date(Date.now() + data.expires_in * 1000);
  return cachedToken;
}

// 📧 API endpoint to send email
app.post("/send-mail", async (req, res) => {
  try {
    const { to, subject, body } = req.body;
    const token = await getAccessToken();

    const graphRes = await fetch(`https://graph.microsoft.com/v1.0/users/${EMAIL_USER}/sendMail`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: "HTML", content: body },
          toRecipients: [{ emailAddress: { address: to } }]
        },
        saveToSentItems: true
      })
    });

    if (!graphRes.ok) {
      const error = await graphRes.text();
      console.error("Graph error:", error);
      return res.status(500).send({ error });
    }

    res.send({ status: "✅ Mail sent successfully via Graph API" });
  } catch (e) {
    console.error("Exception:", e);
    res.status(500).send({ error: e.message });
  }
});

app.listen(3000, () => console.log("🚀 Server running on port 3000"));
