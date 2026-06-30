export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "Falta configurar ANTHROPIC_API_KEY en Vercel" });
  }

  try {
    const { base64, mediaType } = req.body;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        messages: [
          {
            role: "user",
            content: [
              { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
              {
                type: "text",
                text: `Esto es un comprobante de pago/transferencia de alquiler en Argentina. Extraé estos datos y respondé ÚNICAMENTE con un JSON, sin texto adicional, sin markdown, sin backticks:
{
  "monto": (número, sin separadores de miles, solo dígitos y punto decimal),
  "moneda": ("ARS" o "USD"),
  "fecha": (formato YYYY-MM-DD, la fecha de la transferencia/pago. Si no se ve, usar null),
  "notas": (texto breve con cualquier referencia, número de operación, CBU, alias o nombre del emisor que aparezca)
}
Si algún dato no se puede leer con certeza, usá null en ese campo.`,
              },
            ],
          },
        ],
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({ error: data.error?.message || "Error de la API de Anthropic" });
    }

    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
