const express = require('express');
const axios = require('axios');
const app = express();

app.use(express.json());

app.post('/createMeeting', async (req, res) => {
  try {
    const { access_token, meeting } = req.body;

    const response = await axios.post(
      'https://graph.microsoft.com/v1.0/users/user@faramondtechnologies.com/onlineMeetings',
      meeting,
      {
        headers: {
          Authorization: `Bearer ${access_token}`,
          'Content-Type': 'application/json'
        }
      }
    );

    res.json(response.data);
  } catch (err) {
    res.status(500).json({ error: err.response?.data || err.message });
  }
});

app.get('/', (req, res) => {
  res.send('APEX Teams API is running!');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
