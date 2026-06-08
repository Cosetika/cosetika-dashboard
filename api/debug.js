export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  const API_TOKEN = process.env.CONTIFICO_API_TOKEN;
  
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  // Probar diferentes combinaciones de autenticación
  const pruebas = [
    { 
      label: 'v1 solo API_KEY',
      url: 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC', 
      auth: API_KEY 
    },
    { 
      label: 'v1 API_KEY:API_TOKEN',
      url: 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC', 
      auth: `${API_KEY}:${API_TOKEN}` 
    },
    { 
      label: 'v2 solo API_KEY',
      url: 'https://api.contifico.com/sistema/api/v2/documento/?tipo_documento=FAC', 
      auth: API_KEY 
    },
    { 
      label: 'v1 marca (endpoint simple)',
      url: 'https://api.contifico.com/sistema/api/v1/marca/', 
      auth: API_KEY 
    },
    { 
      label: 'base demo solo API_KEY',
      url: 'https://base.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC', 
      auth: API_KEY 
    },
  ];

  const resultados = { 
    api_key: API_KEY ? API_KEY.substring(0, 8) + '...' : 'NO DEFINIDO',
    api_token: API_TOKEN ? API_TOKEN.substring(0, 8) + '...' : 'NO DEFINIDO'
  };

  for (const p of pruebas) {
    try {
      const response = await fetch(p.url, {
        headers: { 
          'Authorization': p.auth,
          'Accept': 'application/json'
        }
      });
      const text = await response.text();
      resultados[p.label] = {
        status: response.status,
        respuesta: text.substring(0, 200)
      };
    } catch(e) {
      resultados[p.label] = { error: e.message };
    }
  }

  return res.status(200).json(resultados);
}
