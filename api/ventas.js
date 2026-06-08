export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) {
    return res.status(500).json({ error: 'API Key no configurada. Agrégala en Vercel > Settings > Environment Variables' });
  }

  try {
    const { fecha_inicio, fecha_fin, tipo } = req.query;
    const tipoDoc = tipo || 'FAC';

    let url = `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=${tipoDoc}`;
    if (fecha_inicio) url += `&fecha_inicio=${fecha_inicio}`;
    if (fecha_fin)    url += `&fecha_fin=${fecha_fin}`;

    const response = await fetch(url, {
      headers: { 'Authorization': API_KEY }
    });

    if (!response.ok) {
      const text = await response.text();
      return res.status(response.status).json({ error: `Error Contifico: ${response.status}`, detalle: text });
    }

    const data = await response.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error interno', detalle: err.message });
  }
}
