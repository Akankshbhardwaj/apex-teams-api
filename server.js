import express from "express";
import axios from "axios";
import bodyParser from "body-parser";

const app = express();
app.use(bodyParser.json());

app.get("/", (req, res) => {
  res.send("Microsoft Graph API - APEX Bridge is running 🚀");
});

app.post("/createMeeting", async (req, res) => {
  try {
    const tenantId = "5fbf0bf3-08f7-4648-b12b-ee3b3de59636";
    const clientId = "a5255c20-3bbd-4f4d-b996-e9d54e5d2077";
    const clientSecret = process.env.CLIENT_SECRET; // <-- Secret loaded from env

    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        scope: "https://graph.microsoft.com/.default",
        client_secret: clientSecret,
        grant_type: "client_credentials",
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const accessToken = tokenResponse.data.access_token;

    const meetingResponse = await axios.post(
      "https://graph.microsoft.com/v1.0/me/onlineMeetings",
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
    console.error("Error creating meeting:", error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

app.listen(10000, () => console.log("Server running on port 10000"));
