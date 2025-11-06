import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.json());

// Root endpoint
app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph API - APEX Bridge is running 🚀");
});

// Send Mail route
app.post("/send-mail", async (req, res) => {
  const { to, subject, body } = req.body;

  try {
    // Step 1: Get access token from Microsoft
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.CLIENT_ID,
          scope: process.env.GRAPH_SCOPE,
          client_secret: process.env.CLIENT_SECRET,
          grant_type: "client_credentials",
        }),
      }
    );

    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;

    if (!accessToken) {
      console.error("Failed to get access token:", tokenData);
      return res.status(500).json({ error: "Failed to obtain access token" });
    }

    // Step 2: Send email via Microsoft Graph
    const mailPayload = {
      message: {
        subject: subject || "No subject",
        body: {
          contentType: "HTML",
          content: body || "No content",
        },
        toRecipients: [
          {
            emailAddress: {
              address: to || process.env.EMAIL_USER,
            },
          },
        ],
      },
      saveToSentItems: true,
    };

    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/users/${process.env.EMAIL_USER}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(mailPayload),
      }
    );

    if (!mailResponse.ok) {
      const error = await mailResponse.text();
      console.error("Mail send failed:", error);
      return res.status(500).json({ error: "Mail send failed", details: error });
    }

    res.json({ message: "✅ Email sent successfully!" });
  } catch (err) {
    console.error("Error sending email:", err);
    res.status(500).json({ error: "Internal Server Error", details: err.message });
  }
});

// Start server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
