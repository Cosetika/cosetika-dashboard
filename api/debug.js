export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  const API_TOKEN = process.env.CONTIFICO_API_TOKEN;

  const pruebas = [
    { label: 'base v1 solo key', url: 'https://base.contifico.com/sistema/api/v1/marca/', auth: API_KEY },
    { label: 'base v1 key:token', url: 'https://base.contifico.com/sistema/api/v1/marca/', auth: `${API_KEY}:${API_TOKEN}` },
    { label: 'api v1 solo key', url: 'https://api.contifico.com/sistema/api/v1/marca/', auth: API_KEY },
    { label: 'api v1 key:token', url: 'https://api.contifico.com/sistema/api/v1/marca/', auth: `${API_KEY}:${API_TOKEN}` },
  ];

  const resultados = {
    key: API_KEY ? API_KEY.substring(0,8)+'...' : 'NO',
    token: API_TOKEN ? API_TOKEN.substring(0,8)+'...' : 'NO'
  };

  for (const p of pruebas) {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 8000);
      const response = await fetch(p.url, {
        headers: { 'Authorization': p.auth, 'Accept': 'application/json' },
        signal: controller.signal
      });
      clearTimeout(timeout);
      const text = await response.text();
      resultados[p.label] = { status: response.status, resp: text.substring(0, 150) };
    } catch(e) {
      resultados[p.label] = { error: e.message };
    }
  }

  return res.status(200).json(resultados);
}
