require("dotenv").config();
const express = require("express");
const path    = require("path");

const app = express();
app.use(express.json({ limit: "50mb" }));

// ── Em produção: serve o build do Vite ──────────────────────────────────────
if (process.env.NODE_ENV === "production") {
  app.use(express.static(path.join(__dirname, "dist")));
}

// ── Endpoint que faz o proxy seguro para a API da Anthropic ─────────────────
app.post("/api/parse", async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: { message: "ANTHROPIC_API_KEY não configurada no .env" } });
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type":    "application/json",
        "x-api-key":       apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(req.body),
    });

    const data = await response.json();
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: { message: err.message } });
  }
});

// ── SPA fallback (produção) ──────────────────────────────────────────────────
if (process.env.NODE_ENV === "production") {
  app.get("*", (_req, res) => {
    res.sendFile(path.join(__dirname, "dist", "index.html"));
  });
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n✅ Servidor rodando em http://localhost:${PORT}`);
  if (process.env.NODE_ENV !== "production") {
    console.log(`   Frontend (dev): http://localhost:3000\n`);
  }
});
