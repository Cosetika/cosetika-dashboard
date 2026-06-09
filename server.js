const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');
const { Pool } = require('pg');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

// PostgreSQL
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// Crear tabla si no existe
async function initDB() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS visitas (
        id SERIAL PRIMARY KEY,
        lugar VARCHAR(255) NOT NULL,
        tipo VARCHAR(50) NOT NULL,
        asesora VARCHAR(255) NOT NULL,
        fecha TIMESTAMP DEFAULT NOW(),
        notas TEXT
      )
    `);
    console.log('DB inicializada correctamente');
  } catch(e) {
    console.error('Error inicializando DB:', e.message);
  }
}

initDB();

const MIME = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon'
};

function bodyJSON(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); } catch(e) { reject(e); }
    });
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const urlPath = req.url.split('?')[0];

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // GET visitas
  if (urlPath === '/api/visitas' && req.method === 'GET') {
    try {
      const result = await pool.query(
        'SELECT * FROM visitas ORDER BY fecha DESC LIMIT 200'
      );
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result.rows));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // POST visita
  if (urlPath === '/api/visitas' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const { lugar, tipo, asesora, notas } = body;
      if (!lugar || !tipo || !asesora) throw new Error('Faltan campos');
      const result = await pool.query(
        'INSERT INTO visitas (lugar, tipo, asesora, notas) VALUES ($1, $2, $3, $4) RETURNING *',
        [lugar, tipo, asesora, notas || null]
      );
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result.rows[0]));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // DELETE visita
  if (urlPath.startsWith('/api/visitas/') && req.method === 'DELETE') {
    try {
      const id = urlPath.split('/').pop();
      await pool.query('DELETE FROM visitas WHERE id = $1', [id]);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // Debug
  if (urlPath === '/api/debug') {
    try {
      const dbCheck = await pool.query('SELECT COUNT(*) FROM visitas');
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        key: API_KEY.substring(0,8)+'...',
        db: 'conectada',
        visitas_count: dbCheck.rows[0].count
      }));
    } catch(e) {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ db_error: e.message }));
    }
    return;
  }


  // POST /api/chat - Chatbot con Claude
  if (urlPath === '/api/chat' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const { messages, system } = body;

      const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
      if (!ANTHROPIC_KEY) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'API Key de Anthropic no configurada' }));
        return;
      }

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': ANTHROPIC_KEY,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 1024,
          system: system || '',
          messages: messages || []
        })
      });

      const data = await response.json();
      res.writeHead(response.status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(data));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // Static files
  let filePath = urlPath === '/' ? path.join(__dirname, 'index.html')
    : urlPath === '/login' ? path.join(__dirname, 'public', 'login.html')
    : urlPath === '/bot' ? path.join(__dirname, 'bot.html')
    : path.join(__dirname, urlPath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    res.writeHead(200, { 'Content-Type': MIME[path.extname(filePath)] || 'text/plain' });
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    fs.createReadStream(path.join(__dirname, 'index.html')).pipe(res);
  }
});

server.listen(PORT, '0.0.0.0', () => console.log(`Cosétika Dashboard running on port ${PORT}`));
