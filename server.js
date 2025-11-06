// server.js
import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();
const app = express();
app.use(express.json());

app.get("/", (req, res) => {
  res.send("✅ Microsoft Graph API - APEX Bridge is running 🚀");
});

app.post("/createMeeting", async (req, res) => {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const graphScope = process.env.GRAPH_SCOPE || "https://graph.microsoft.com/.default";

    // Step 1: Get Access Token
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: graphScope,
          grant_type: "client_credentials",
        }),
      }
    );

    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) {
      console.error("Error fetching token:", tokenData);
      return res.status(500).json({ error: "Failed to get access token", details: tokenData });
    }

    const accessToken = tokenData.access_token;

    // Step 2: Create Meeting
    const meetingResponse = await fetch("https://graph.microsoft.com/v1.0/me/onlineMeetings", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        subject: "APEX Auto Meeting via Render",
        startDateTime: "2025-11-06T13:30:00Z",
        endDateTime: "2025-11-06T14:00:00Z",
      }),
    });

    const meetingData = await meetingResponse.json();
    if (!meetingResponse.ok) {
      console.error("Error creating meeting:", meetingData);
      return res.status(500).json({ error: meetingData });
    }

    res.json({ meeting: meetingData });
  } catch (err) {
    console.error("Unexpected error:", err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
