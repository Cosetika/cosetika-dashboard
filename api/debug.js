export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  try {
    const controller = new AbortController();
    setTimeout(() => controller.abort(), 7000);
    
    const url = 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&fecha_inicio=01/06/26&fecha_fin=08/06/26';
    
    const response = await fetch(url, {
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' },
      signal: controller.signal
    });
    const text = await response.text();
    
    return res.status(200).json({ 
      status: response.status, 
      resp: text.substring(0, 500),
      url 
    });
  } catch(e) {
    return res.status(200).json({ error: e.message });
  }
}
