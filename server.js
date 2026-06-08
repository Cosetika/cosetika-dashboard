const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

function fetchContifico(url) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'GET',
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    };
    https.get(url, options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    }).on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const url = req.url.split('?')[0];

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  if (url === '/api/ventas') {
    if (!API_KEY) { res.writeHead(500, {'Content-Type':'application/json'}); res.end(JSON.stringify({error:'API Key no configurada'})); return; }
    try {
      const r = await fetchContifico('https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&offset=0&limit=50');
      let data;
      try { data = JSON.parse(r.body); } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'JSON inválido',raw:r.body.substring(0,200)})); return; }
      const docs = Array.isArray(data) ? data : (data.results||data.data||[]);
      const procesados = docs.map(doc => ({
        fecha: doc.fecha_emision||'',
        vendedor: doc.vendedor_nombre||doc.vendedor||'Sin vendedor',
        cliente: doc.cliente_razon_social||doc.razon_social||'',
        provincia: doc.provincia||'',
        total: parseFloat((doc.total||'0').toString().replace(',','.')),
        estado: doc.estado||'',
        detalles: (doc.detalles||[]).map(d => ({
          producto: d.producto_nombre||d.nombre||'',
          marca: d.adicional3||d.marca||'',
          cantidad: parseFloat((d.cantidad||'0').toString().replace(',','.')),
          total: parseFloat((d.base_gravable||'0').toString().replace(',','.'))
        }))
      }));
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({total:procesados.length,documentos:procesados}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  if (url === '/api/debug') {
    const r = await fetchContifico('https://api.contifico.com/sistema/api/v1/marca/');
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({key:API_KEY.substring(0,8)+'...',status:r.status,resp:r.body.substring(0,200)}));
    return;
  }

  // Servir archivos estáticos
  let filePath = url === '/' ? '/index.html' : url;
  if (filePath === '/login') filePath = '/public/login.html';
  
  const fullPath = path.join(__dirname, filePath.startsWith('/public') ? filePath : filePath);
  
  if (fs.existsSync(fullPath) && fs.statSync(fullPath).isFile()) {
    const ext = path.extname(fullPath);
    const types = {'.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json'};
    res.writeHead(200,{'Content-Type':types[ext]||'text/plain'});
    fs.createReadStream(fullPath).pipe(res);
  } else {
    // SPA fallback
    const index = path.join(__dirname, 'index.html');
    if (fs.existsSync(index)) {
      res.writeHead(200,{'Content-Type':'text/html'});
      fs.createReadStream(index).pipe(res);
    } else {
      res.writeHead(404); res.end('Not found');
    }
  }
});

server.listen(PORT, () => console.log(`Cosétika Dashboard running on port ${PORT}`));
