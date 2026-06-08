export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'API Key no configurada' });

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25000);

    const response = await fetch(
      'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC',
      {
        headers: { 'Authorization': API_KEY, 'Accept': 'application/json' },
        signal: controller.signal
      }
    );
    clearTimeout(timeout);

    const text = await response.text();
    let data;
    try { data = JSON.parse(text); } catch(e) {
      return res.status(500).json({ error: 'JSON inválido', raw: text.substring(0, 300) });
    }

    const documentos = Array.isArray(data) ? data : (data.results || data.data || []);

    const procesados = documentos.map(doc => {
      const detalles = (doc.detalles || []).map(d => ({
        producto: d.producto_nombre || d.nombre || d.descripcion || '',
        marca: d.adicional3 || d.marca || '',
        categoria: d.categoria || '',
        cantidad: parseFloat((d.cantidad||'0').toString().replace(',','.')),
        precio: parseFloat((d.precio||'0').toString().replace(',','.')),
        total: parseFloat((d.base_gravable||d.subtotal||'0').toString().replace(',','.'))
      }));

      return {
        fecha: doc.fecha_emision || '',
        numero: doc.numero || doc.secuencial || '',
        vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
        cliente: doc.cliente_razon_social || doc.razon_social || '',
        provincia: doc.provincia || '',
        total: parseFloat((doc.total||'0').toString().replace(',','.')),
        estado: doc.estado || '',
        detalles
      };
    });

    return res.status(200).json({ total: procesados.length, documentos: procesados });

  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
}
