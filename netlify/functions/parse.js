// netlify/functions/parse.js
exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: "GEMINI_API_KEY não configurada" }) };
  }

  try {
    const { type, content, data: pdfData, prompt } = JSON.parse(event.body);

    const parts = type === "pdf"
      ? [{ inline_data: { mime_type: "application/pdf", data: pdfData } }, { text: prompt }]
      : [{ text: content }];

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`,
      { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ contents: [{ parts }] }) }
    );

    const data = await response.json();
    if (data.error) return { statusCode: 500, body: JSON.stringify({ error: data.error.message }) };

    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }) };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
