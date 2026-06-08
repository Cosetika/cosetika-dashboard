export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Sin API Key' });

  try {
    // Solo 10 documentos para probar velocidad
    const url = 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&offset=0&limit=10';
    
    const response = await fetch(url, {
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    });

    if (!response.ok) {
      const txt = await response.text();
      return res.status(200).json({ total: 0, documentos: [], debug: { status: response.status, msg: txt.substring(0,200) } });
    }

    const data = await response.json();
    const documentos = Array.isArray(data) ? data : (data.results || data.data || []);

    const procesados = documentos.map(doc => ({
      fecha: doc.fecha_emision || '',
      vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
      cliente: doc.cliente_razon_social || doc.razon_social || '',
      provincia: doc.provincia || '',
      total: parseFloat((doc.total || '0').toString().replace(',', '.')),
      estado: doc.estado || '',
      detalles: (doc.detalles || []).map(d => ({
        producto: d.producto_nombre || d.nombre || '',
        marca: d.adicional3 || d.marca || '',
        cantidad: parseFloat((d.cantidad || '0').toString().replace(',', '.')),
        total: parseFloat((d.base_gravable || '0').toString().replace(',', '.'))
      }))
    }));

    return res.status(200).json({ total: procesados.length, documentos: procesados });
  } catch(err) {
    return res.status(200).json({ total: 0, documentos: [], debug: { error: err.message } });
  }
}
