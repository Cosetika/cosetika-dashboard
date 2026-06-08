export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) {
    return res.status(500).json({ error: 'API Key no configurada' });
  }

  try {
    // Sin filtro de estado para traer todas las facturas
    const url = `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC`;

    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    });

    const text = await response.text();

    if (!text || text.trim() === '') {
      return res.status(500).json({ 
        error: 'Contifico devolvió respuesta vacía',
        status: response.status
      });
    }

    let data;
    try {
      data = JSON.parse(text);
    } catch(e) {
      return res.status(500).json({ 
        error: 'Contifico no devolvió JSON válido',
        respuesta_raw: text.substring(0, 500),
        status: response.status
      });
    }

    // Mostrar estructura raw para debug
    const documentos = Array.isArray(data) ? data : (data.results || data.data || data.documentos || []);

    if (documentos.length === 0) {
      return res.status(200).json({ 
        total: 0, 
        documentos: [],
        debug_keys: Object.keys(data),
        debug_sample: JSON.stringify(data).substring(0, 300)
      });
    }

    // Mapear campos de Contifico
    const procesados = documentos.map(doc => {
      const detalles = (doc.detalles || []).map(d => ({
        producto: d.producto_nombre || d.nombre || d.descripcion || '',
        categoria: d.categoria || d.cat_producto || '',
        marca: d.adicional3 || d.marca || '',
        cantidad: parseFloat((d.cantidad || '0').toString().replace(',', '.')),
        precio: parseFloat((d.precio || '0').toString().replace(',', '.')),
        total: parseFloat((d.base_gravable || d.subtotal || '0').toString().replace(',', '.'))
      }));

      return {
        fecha: doc.fecha_emision || '',
        numero: doc.numero || doc.secuencial || '',
        vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
        cliente: doc.cliente_razon_social || doc.razon_social || '',
        provincia: doc.provincia || '',
        total: parseFloat((doc.total || '0').toString().replace(',', '.')),
        estado: doc.estado || '',
        detalles
      };
    });

    return res.status(200).json({
      total: procesados.length,
      documentos: procesados
    });

  } catch (err) {
    return res.status(500).json({ 
      error: 'Error interno', 
      detalle: err.message 
    });
  }
}
