import express from "express";
import axios from "axios";
import bodyParser from "body-parser";
import dotenv from "dotenv";

dotenv.config(); // Load env vars from Render or .env

const app = express();
app.use(bodyParser.json());

// Read from environment variables
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;
const GRAPH_SCOPE = process.env.GRAPH_SCOPE || "https://graph.microsoft.com/.default";
const EMAIL_USER = process.env.EMAIL_USER;

// Health check route
app.get("/", (req, res) => {
  res.send("🚀 Microsoft Graph API - APEX Bridge is running fine!");
});

// Create Microsoft Teams meeting
app.post("/createMeeting", async (req, res) => {
  try {
    // 1️⃣ Get Access Token from Azure AD
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: GRAPH_SCOPE,
        grant_type: "client_credentials",
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const accessToken = tokenResponse.data.access_token;

    // 2️⃣ Create Online Meeting
    const meetingResponse = await axios.post(
      `https://graph.microsoft.com/v1.0/users/${EMAIL_USER}/onlineMeetings`,
      {
        subject: "APEX Auto Meeting",
        startDateTime: "2025-11-06T13:30:00Z",
        endDateTime: "2025-11-06T14:00:00Z",
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    res.json({ meeting: meetingResponse.data });
  } catch (error) {
    console.error("❌ Error creating meeting:", error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

// Start server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
