export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  try {
    // Probar endpoint de personas (clientes) que debería ser más rápido
    const url = 'https://api.contifico.com/sistema/api/v1/persona/?tipo=CLI&offset=0&limit=10';
    
    const response = await fetch(url, {
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    });

    const txt = await response.text();
    let data;
    try { data = JSON.parse(txt); } catch(e) {
      return res.status(200).json({ debug: 'JSON inválido', raw: txt.substring(0,200) });
    }

    // Convertir personas a formato de vendedores para el dashboard
    const items = Array.isArray(data) ? data : (data.results || data.data || []);
    
    return res.status(200).json({ 
      total: items.length, 
      items,
      msg: 'Conexión exitosa con Contifico'
    });

  } catch(err) {
    return res.status(200).json({ error: err.message });
  }
}
