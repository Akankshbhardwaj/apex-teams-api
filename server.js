import express from "express";
import axios from "axios";

const app = express();
app.use(express.json());

// Root route
app.get("/", (req, res) => {
  res.send("APEX → Render API → Microsoft Graph working 🚀");
});

// Endpoint to create a Teams meeting
app.post("/createMeeting", async (req, res) => {
  const { clientId, clientSecret, tenantId, userEmail, startDateTime, endDateTime, subject } = req.body;

  if (!clientId || !clientSecret || !tenantId || !userEmail) {
    return res.status(400).json({ error: "Missing required parameters" });
  }

  try {
    // Step 1: Get Access Token
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const accessToken = tokenResponse.data.access_token;

    // Step 2: Create Teams Meeting
    const meetingResponse = await axios.post(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userEmail)}/onlineMeetings`,
      {
        startDateTime,
        endDateTime,
        subject
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        }
      }
    );

    res.json(meetingResponse.data);
  } catch (error) {
    console.error("Error:", error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
