export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) {
    return res.status(500).json({ error: 'API Key no configurada' });
  }

  try {
    const response = await fetch('https://api.contifico.com/sistema/api/v1/producto/', {
      headers: { 'Authorization': API_KEY }
    });

    if (!response.ok) {
      return res.status(response.status).json({ error: `Error Contifico: ${response.status}` });
    }

    const data = await response.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error interno', detalle: err.message });
  }
}
