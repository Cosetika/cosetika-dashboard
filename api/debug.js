export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  const pruebas = [
    { label: 'doc sin filtro', url: 'https://api.contifico.com/sistema/api/v1/documento/' },
    { label: 'doc tipo FAC', url: 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC' },
    { label: 'doc fecha junio', url: 'https://api.contifico.com/sistema/api/v1/documento/?fecha_inicio=01/06/26&fecha_fin=08/06/26' },
    { label: 'doc FAC + fecha', url: 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&fecha_inicio=01/06/26&fecha_fin=08/06/26' },
    { label: 'registro documento', url: 'https://api.contifico.com/sistema/api/v1/registro/documento/' },
  ];

  const resultados = { key: API_KEY.substring(0,8)+'...' };

  for (const p of pruebas) {
    try {
      const controller = new AbortController();
      setTimeout(() => controller.abort(), 8000);
      const response = await fetch(p.url, {
        headers: { 'Authorization': API_KEY, 'Accept': 'application/json' },
        signal: controller.signal
      });
      const text = await response.text();
      resultados[p.label] = { status: response.status, resp: text.substring(0, 200) };
    } catch(e) {
      resultados[p.label] = { error: e.message };
    }
  }

  return res.status(200).json(resultados);
}
