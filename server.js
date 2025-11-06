import express from "express";
import axios from "axios";

const app = express();
app.use(express.json());

// Root test
app.get("/", (req, res) => {
  res.send("APEX → Render API → Microsoft Graph working 🚀");
});

// Endpoint to create a Teams meeting
app.post("/createMeeting", async (req, res) => {
  const { accessToken, startDateTime, endDateTime, subject, userEmail } = req.body;

  if (!accessToken || !userEmail) {
    return res.status(400).json({ error: "Missing required parameters" });
  }

  try {
    const response = await axios.post(
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

    res.json(response.data);
  } catch (error) {
    console.error("Error creating meeting:", error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
