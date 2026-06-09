const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');
const { Pool } = require('pg');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

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
      );
      CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(255) NOT NULL,
        usuario VARCHAR(100) UNIQUE NOT NULL,
        password VARCHAR(255) NOT NULL,
        rol VARCHAR(50) DEFAULT 'asesora',
        modulos TEXT DEFAULT 'ventas,visitas,kpis,inventario',
        activo BOOLEAN DEFAULT true,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Insertar usuarios iniciales si no existen
    const usuarios = [
      { nombre: 'Fernando Espíndola', usuario: 'Fernando', password: '1234', rol: 'admin', modulos: 'ventas,visitas,kpis,inventario,config' },
      { nombre: 'Giovanna Portilla', usuario: 'Giovanna', password: '1234', rol: 'jefa_ventas', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Daniela Villegas', usuario: 'Daniela', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Liseth Gavilanes', usuario: 'Liseth', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Karen Rebeca Mora', usuario: 'Karen', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'María Caridad Zea', usuario: 'Maria', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Nicole Yanira Leon', usuario: 'Nicole', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
    ];

    for (const u of usuarios) {
      await pool.query(`
        INSERT INTO usuarios (nombre, usuario, password, rol, modulos)
        VALUES ($1, $2, $3, $4, $5)
        ON CONFLICT (usuario) DO NOTHING
      `, [u.nombre, u.usuario, u.password, u.rol, u.modulos]);
    }
    console.log('DB inicializada');
  } catch(e) {
    console.error('Error DB:', e.message);
  }
}

initDB();

const MIME = {
  '.html': 'text/html', '.js': 'application/javascript',
  '.css': 'text/css', '.json': 'application/json',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.ico': 'image/x-icon'
};

function bodyJSON(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch(e) { reject(e); } });
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const urlPath = req.url.split('?')[0];
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // LOGIN
  if (urlPath === '/api/login' && req.method === 'POST') {
    try {
      const { usuario, password } = await bodyJSON(req);
      const result = await pool.query(
        'SELECT * FROM usuarios WHERE usuario = $1 AND password = $2 AND activo = true',
        [usuario, password]
      );
      if (result.rows.length === 0) {
        res.writeHead(401, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Usuario o contraseña incorrectos' }));
        return;
      }
      const u = result.rows[0];
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: true,
        usuario: { id: u.id, nombre: u.nombre, usuario: u.usuario, rol: u.rol, modulos: u.modulos }
      }));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // GET usuarios (solo admin)
  if (urlPath === '/api/usuarios' && req.method === 'GET') {
    try {
      const result = await pool.query('SELECT id, nombre, usuario, rol, modulos, activo FROM usuarios ORDER BY id');
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result.rows));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // PUT usuario
  if (urlPath.startsWith('/api/usuarios/') && req.method === 'PUT') {
    try {
      const id = urlPath.split('/').pop();
      const body = await bodyJSON(req);
      const { password, modulos, activo, rol } = body;
      if (password) {
        await pool.query('UPDATE usuarios SET password = $1 WHERE id = $2', [password, id]);
      }
      if (modulos !== undefined) {
        await pool.query('UPDATE usuarios SET modulos = $1 WHERE id = $2', [modulos, id]);
      }
      if (activo !== undefined) {
        await pool.query('UPDATE usuarios SET activo = $1 WHERE id = $2', [activo, id]);
      }
      if (rol !== undefined) {
        await pool.query('UPDATE usuarios SET rol = $1 WHERE id = $2', [rol, id]);
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // POST nuevo usuario
  if (urlPath === '/api/usuarios' && req.method === 'POST') {
    try {
      const { nombre, usuario, password, rol, modulos } = await bodyJSON(req);
      await pool.query(
        'INSERT INTO usuarios (nombre, usuario, password, rol, modulos) VALUES ($1, $2, $3, $4, $5)',
        [nombre, usuario, password || '1234', rol || 'asesora', modulos || 'ventas,visitas,kpis,inventario']
      );
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // GET visitas
  if (urlPath === '/api/visitas' && req.method === 'GET') {
    try {
      const result = await pool.query('SELECT * FROM visitas ORDER BY fecha DESC LIMIT 300');
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
      const { lugar, tipo, asesora, notas } = await bodyJSON(req);
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

  // Chat API
  if (urlPath === '/api/chat' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
      if (!ANTHROPIC_KEY) { res.writeHead(500, {'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Sin API Key'})); return; }
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_KEY, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 1024, system: body.system || '', messages: body.messages || [] })
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
    : urlPath === '/login' ? path.join(__dirname, 'login.html')
    : urlPath === '/bot' ? path.join(__dirname, 'bot.html')
    : urlPath === '/sofia.jpg' ? path.join(__dirname, 'sofia.jpg')
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
