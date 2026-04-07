// api/gemini.js (o reemplaza claude.js)
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const { messages } = req.body;
  // Extraemos el contenido que enviaste (el prompt y el base64)
  const userContent = messages[0].content;
  const docPart = userContent.find(p => p.type === 'document' || p.type === 'image');
  const textPart = userContent.find(p => p.type === 'text');

  const API_KEY = process.env.GEMINI_API_KEY; // Agrégala en las variables de entorno de Vercel

  const geminiPayload = {
    contents: [{
      parts: [
        { text: textPart.text },
        {
          inline_data: {
            mime_type: docPart.source.media_type,
            data: docPart.source.data
          }
        }
      ]
    }],
    generationConfig: {
      response_mime_type: "application/json",
      temperature: 0.1
    }
  };

  try {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${API_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(geminiPayload)
      }
    );

    const data = await response.json();
    
    // Gemini devuelve el texto JSON dentro de esta estructura
    const resultText = data.candidates[0].content.parts[0].text;
    
    res.status(200).json({ content: [{ text: resultText }] });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
}
