export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  const RUC = process.env.CONTIFICO_RUC;
  
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  const endpoints = [
    `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC`,
    `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&ruc=${RUC}`,
    `https://api.contifico.com/sistema/api/v1/documento/?ruc=${RUC}`,
  ];

  const resultados = { ruc_usado: RUC };

  for (const url of endpoints) {
    try {
      const response = await fetch(url, {
        headers: { 
          'Authorization': API_KEY,
          'Accept': 'application/json',
          'X-RUC': RUC || ''
        }
      });
      const text = await response.text();
      resultados[url] = {
        status: response.status,
        respuesta: text.substring(0, 300)
      };
    } catch(e) {
      resultados[url] = { error: e.message };
    }
  }

  return res.status(200).json(resultados);
}
