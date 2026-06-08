export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  const endpoints = [
    'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC',
    'https://api.contifico.com/sistema/api/v1/documento/',
    'https://api.contifico.com/sistema/api/v1/registro/documento/',
  ];

  const resultados = {};

  for (const url of endpoints) {
    try {
      const response = await fetch(url, {
        headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
      });
      const text = await response.text();
      resultados[url] = {
        status: response.status,
        primeros_200_chars: text.substring(0, 200)
      };
    } catch(e) {
      resultados[url] = { error: e.message };
    }
  }

  return res.status(200).json(resultados);
}
