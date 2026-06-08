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
    // Endpoint correcto de Contifico para documentos/facturas
    const url = `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&estado=C`;
 
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    });
 
    const text = await response.text();
 
    // Debug: mostrar qué devuelve Contifico
    if (!text || text.trim() === '') {
      return res.status(500).json({ 
        error: 'Contifico devolvió respuesta vacía',
        status: response.status,
        statusText: response.statusText
      });
    }
 
    // Intentar parsear JSON
    let data;
    try {
      data = JSON.parse(text);
    } catch(e) {
      return res.status(500).json({ 
        error: 'Contifico no devolvió JSON válido',
        respuesta_raw: text.substring(0, 500),
        status: response.status,
        url_usada: url
      });
    }
 
    // Si es array directo o tiene paginación
    const documentos = Array.isArray(data) ? data : (data.results || data.data || data.documentos || []);
 
    // Mapear campos exactos del CSV de Contifico al formato del dashboard
    const procesados = documentos.map(doc => {
      const detalles = (doc.detalles || []).map(d => ({
        producto: d.producto_nombre || d.nombre || d.descripcion || '',
        categoria: d.categoria || '',
        marca: d.adicional3 || d.marca || '',
        cantidad: parseFloat(d.cantidad || 0),
        precio: parseFloat(d.precio || 0),
        total: parseFloat(d.base_gravable || 0) + parseFloat(d.base_cero || 0)
      }));
 
      return {
        fecha: doc.fecha_emision || '',
        numero: doc.numero || doc.id || '',
        vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
        cliente: doc.cliente_razon_social || doc.razon_social || doc.cliente || '',
        cedula: doc.cliente_cedula || doc.cedula || '',
        provincia: doc.provincia || '',
        total: parseFloat(doc.total || 0),
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
      error: 'Error interno del servidor', 
      detalle: err.message 
    });
  }
}
