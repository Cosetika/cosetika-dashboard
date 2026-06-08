export const config = {
  runtime: 'edge',
};

export default async function handler(req) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json',
  };

  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers });
  }

  const API_KEY = process.env.CONTIFICO_API_KEY;
  if (!API_KEY) {
    return new Response(JSON.stringify({ error: 'API Key no configurada' }), { status: 500, headers });
  }

  try {
    const url = 'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&offset=0&limit=50';

    const response = await fetch(url, {
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    });

    const text = await response.text();
    let data;
    try { data = JSON.parse(text); } catch(e) {
      return new Response(JSON.stringify({ error: 'JSON inválido', raw: text.substring(0, 300) }), { status: 500, headers });
    }

    const documentos = Array.isArray(data) ? data : (data.results || data.data || []);

    const procesados = documentos.map(doc => {
      const detalles = (doc.detalles || []).map(d => ({
        producto: d.producto_nombre || d.nombre || d.descripcion || '',
        marca: d.adicional3 || d.marca || '',
        cantidad: parseFloat((d.cantidad||'0').toString().replace(',','.')),
        total: parseFloat((d.base_gravable||d.subtotal||'0').toString().replace(',','.'))
      }));

      return {
        fecha: doc.fecha_emision || '',
        vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
        cliente: doc.cliente_razon_social || doc.razon_social || '',
        provincia: doc.provincia || '',
        total: parseFloat((doc.total||'0').toString().replace(',','.')),
        estado: doc.estado || '',
        detalles
      };
    });

    return new Response(
      JSON.stringify({ total: procesados.length, documentos: procesados }),
      { status: 200, headers }
    );

  } catch(err) {
    return new Response(
      JSON.stringify({ error: err.message }),
      { status: 500, headers }
    );
  }
}
