const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// ✅ VERY IMPORTANT: Serve public folder correctly
app.use(express.static(path.join(__dirname, 'public')));

// ✅ ROOT ROUTE → FORCE LOAD LOGIN PAGE
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// ✅ TEST ROUTE (to confirm server works)
app.get('/test', (req, res) => {
  res.send("Server is working");
});

// START SERVER
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});