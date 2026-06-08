const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

// Cache en memoria
let cache = { documentos: [], ultima_sync: null, sincronizando: false };

function fetchContifico(url) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method: 'GET',
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Timeout')); });
    req.end();
  });
}

async function sincronizar() {
  if (cache.sincronizando) return;
  cache.sincronizando = true;
  console.log('Sincronizando con Contifico...');
  try {
    const now = new Date();
    const y = String(now.getFullYear()).slice(-2);
    const url = `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&fecha_inicio=01/01/${y}&fecha_fin=${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${y}`;
    const r = await fetchContifico(url);
    const data = JSON.parse(r.body);
    const docs = Array.isArray(data) ? data : (data.results || data.data || []);
    cache.documentos = docs.map(doc => ({
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
    cache.ultima_sync = new Date().toISOString();
    console.log(`Sincronizado: ${cache.documentos.length} documentos`);
  } catch(e) {
    console.error('Error sync:', e.message);
  }
  cache.sincronizando = false;
}

// Sincronizar al arrancar y cada hora
sincronizar();
setInterval(sincronizar, 60 * 60 * 1000);

const MIME = { '.html':'text/html', '.js':'application/javascript', '.css':'text/css', '.json':'application/json' };

const server = http.createServer(async (req, res) => {
  const urlPath = req.url.split('?')[0];
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  if (urlPath === '/api/ventas') {
    res.writeHead(200, {'Content-Type': 'application/json'});
    res.end(JSON.stringify({
      total: cache.documentos.length,
      documentos: cache.documentos,
      ultima_sync: cache.ultima_sync,
      sincronizando: cache.sincronizando
    }));
    return;
  }

  if (urlPath === '/api/sync') {
    sincronizar(); // dispara sync en background
    res.writeHead(200, {'Content-Type': 'application/json'});
    res.end(JSON.stringify({ msg: 'Sincronización iniciada', ultima_sync: cache.ultima_sync }));
    return;
  }

  if (urlPath === '/api/debug') {
    try {
      const r = await fetchContifico('https://api.contifico.com/sistema/api/v1/marca/');
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ key: API_KEY.substring(0,8)+'...', status: r.status, resp: r.body.substring(0,300), cache_docs: cache.documentos.length, ultima_sync: cache.ultima_sync }));
    } catch(e) {
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // Static files
  let filePath = urlPath === '/' ? path.join(__dirname, 'index.html')
    : urlPath === '/login' ? path.join(__dirname, 'public', 'login.html')
    : path.join(__dirname, urlPath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    res.writeHead(200, {'Content-Type': MIME[path.extname(filePath)] || 'text/plain'});
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200, {'Content-Type': 'text/html'});
    fs.createReadStream(path.join(__dirname, 'index.html')).pipe(res);
  }
});

server.listen(PORT, '0.0.0.0', () => console.log(`Cosétika Dashboard running on port ${PORT}`));
