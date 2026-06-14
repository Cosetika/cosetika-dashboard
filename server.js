const http = require('http');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// ─── CACHÉ v2 ────────────────────────────────────────────────
let cache = { documentos: [], ultima_sync: null, sincronizando: false };

// ─── CACHÉ CATÁLOGO PRODUCTOS ─────────────────────────────────
// { 'NOMBRE PRODUCTO': 'MARCA', ... }
let catalogoProductos = {};
let catalogoSyncedAt = null;

async function sincronizarCatalogo() {
  try {
    console.log('Sincronizando catálogo de productos desde Contifico...');
    let nuevosCatalogo = {};
    let nextUrl = 'https://api.contifico.com/sistema/api/v2/producto/?page_size=100';
    let paginas = 0;
    while (nextUrl && paginas < 50) {
      const resp = await fetch(nextUrl, {
        headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
      });
      if (!resp.ok) { console.error('Error catálogo HTTP:', resp.status); break; }
      const data = await resp.json();
      const results = data.results || [];
      results.forEach(p => {
        const id     = p.id || '';
        const nombre = (p.nombre || p.descripcion || '').trim();
        const marca  = (p.marca_nombre || p.marca || '').trim().toUpperCase();
        const codigo = (p.codigo || p.codigo_principal || '').trim();
        if (id) nuevosCatalogo[id] = { nombre, marca, codigo };
      });
      nextUrl = data.next || null;
      paginas++;
      console.log(`Catálogo pág ${paginas}: ${results.length} prods, total: ${Object.keys(nuevosCatalogo).length}`);
    }
    if (Object.keys(nuevosCatalogo).length > 0) {
      catalogoProductos = nuevosCatalogo;
      catalogoSyncedAt = new Date().toISOString();
      console.log(`✓ Catálogo: ${Object.keys(catalogoProductos).length} productos con marcas`);
    } else {
      console.warn('Catálogo vacío — verificar endpoint /api/v1/producto/');
    }
  } catch(e) {
    console.error('Error catálogo:', e.message);
  }
}

// Sincronizar catálogo al arrancar y cada 24 horas
sincronizarCatalogo().catch(e => console.error('Error catálogo inicial:', e.message));
setInterval(() => sincronizarCatalogo().catch(e => console.error('Error catálogo:', e.message)), 24 * 60 * 60 * 1000);

function fmtDateEC(d) {
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

async function sincronizarHoy() {
  if (cache.sincronizando) return;
  cache.sincronizando = true;
  try {
    const now = new Date();
    const fecha = fmtDateEC(now);
    const url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fecha}&fecha_final=${fecha}&page_size=100`;
    console.log('Sincronizando v2:', url);
    let todos = [];
    let nextUrl = url;
    let paginas = 0;
    while (nextUrl && paginas < 20) {
      const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await resp.json();
      todos = todos.concat(data.results || []);
      nextUrl = data.next || null;
      paginas++;
      console.log(`Página ${paginas}: ${(data.results||[]).length} docs, total: ${todos.length}`);
    }
    const clientes = todos.filter(d => d.tipo_registro === 'CLI');
    // Agregar cliente_nombre directo desde el objeto cliente
    clientes.forEach(d => {
      d.cliente_nombre = d.cliente?.razon_social || d.cliente?.nombre_comercial || d.persona_id || '—';
    });
    cache.documentos = clientes;
    cache.ultima_sync = new Date().toISOString();
    console.log(`✓ Sync: ${clientes.length} facturas de clientes hoy`);
  } catch(e) {
    console.error('Error sync:', e.message);
  }
  cache.sincronizando = false;
}

sincronizarHoy().catch(e => console.error('Error sync inicial:', e.message));
setInterval(() => sincronizarHoy().catch(e => console.error('Error sync:', e.message)), 60 * 60 * 1000);


// ─── GENERADOR DATA.JSON DESDE CONTIFICO ─────────────────────
// Usa producto_id y cliente.id como claves únicas → sin duplicados
// aunque cambien los nombres en Contifico

const EXCLUIR_VENDEDORES = ['Fernando Espíndola', 'Fernando Espindola'];

async function generarDataJson(fechaInicial, fechaFinal) {
  console.log(`Generando data.json: ${fechaInicial} → ${fechaFinal}`);
  // Estructura: { vendedor_id: { nombre, clientes: { cliente_id: {...} } } }
  const vendedores = {};

  let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fechaInicial}&fecha_final=${fechaFinal}&page_size=100`;
  let paginas = 0;

  while (nextUrl && paginas < 200) {
    const resp = await fetch(nextUrl, {
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    });
    if (!resp.ok) { console.error('Error generarData HTTP:', resp.status); break; }
    const data = await resp.json();
    const docs = (data.results || []).filter(d =>
      d.tipo_registro === 'CLI' &&
      !d.anulado &&
      d.vendedor &&
      !EXCLUIR_VENDEDORES.includes(d.vendedor.razon_social)
    );

    docs.forEach(doc => {
      const vendId   = doc.vendedor.id;
      const vendNom  = doc.vendedor.razon_social;
      const cliId    = doc.cliente?.id || doc.persona_id;
      const cliNom   = doc.cliente?.razon_social || doc.cliente?.nombre_comercial || '—';
      const cliRuc   = doc.cliente?.ruc || doc.cliente?.cedula || '';
      const cliProv  = doc.cliente?.adicional1_cliente || '';
      const mes      = parseInt((doc.fecha_emision || '').split('/')[1]) || 0;
      const totalDoc = parseFloat(doc.total || 0);
      const subDoc   = parseFloat(doc.subtotal || doc.subtotal_12 || 0);

      if (!cliId || totalDoc === 0) return;

      // Inicializar vendedor
      if (!vendedores[vendId]) vendedores[vendId] = { nombre: vendNom, clientes: {} };
      // Actualizar nombre por si cambió en Contifico
      vendedores[vendId].nombre = vendNom;

      // Inicializar cliente
      const vObj = vendedores[vendId].clientes;
      if (!vObj[cliId]) {
        vObj[cliId] = {
          id:         cliId,
          nombre:     cliNom,
          ruc:        cliRuc,
          total:      0,
          subtotal:   0,
          num_compras:0,
          provincia:  cliProv,
          marcas:     {},   // { marca: total }
          productos:  {},   // { producto_id: { nombre, codigo, marca, cantidad, total } }
          frecuencia: {}    // { mes: { total, subtotal, compras } }
        };
      }
      const cli = vObj[cliId];
      // Actualizar nombre por si cambió
      cli.nombre   = cliNom;
      cli.ruc      = cliRuc;
      cli.total    += totalDoc;
      cli.subtotal += subDoc;
      cli.num_compras++;

      // Frecuencia por mes
      if (!cli.frecuencia[mes]) cli.frecuencia[mes] = { total:0, subtotal:0, compras:0 };
      cli.frecuencia[mes].total    += totalDoc;
      cli.frecuencia[mes].subtotal += subDoc;
      cli.frecuencia[mes].compras++;

      // Productos y marcas del documento
      (doc.detalles || []).forEach(det => {
        const prodId   = det.producto_id || '';
        const prodNom  = det.producto_nombre || '';
        const cantidad = parseFloat(det.cantidad || 0);
        const base     = parseFloat(det.base_gravable || det.base_cero || 0);
        const pctDesc  = parseFloat(det.porcentaje_descuento || 0);
        if (!prodId || cantidad === 0 || base === 0) return;

        // Buscar en catálogo por producto_id
        const cat   = catalogoProductos[prodId] || {};
        const marca = cat.marca || '';
        const cod   = cat.codigo || '';
        const nom   = cat.nombre || prodNom; // nombre actualizado desde catálogo

        // Acumular producto
        if (!cli.productos[prodId]) {
          cli.productos[prodId] = { id: prodId, nombre: nom, codigo: cod, marca, cantidad: 0, total: 0 };
        }
        cli.productos[prodId].nombre   = nom; // siempre actualizar nombre
        cli.productos[prodId].cantidad += cantidad;
        cli.productos[prodId].total    += base;

        // Acumular marca
        if (marca) {
          cli.marcas[marca] = (cli.marcas[marca] || 0) + base;
        }
      });
    });

    nextUrl = data.next || null;
    paginas++;
    if (paginas % 10 === 0) console.log(`  Página ${paginas}, docs procesados...`);
  }

  // Convertir a formato compatible con el dashboard existente
  const resultado = {};
  Object.values(vendedores).forEach(vend => {
    const clientes = Object.values(vend.clientes).map(cli => ({
      id:          cli.id,
      nombre:      cli.nombre,
      ruc:         cli.ruc,
      total:       Math.round(cli.total * 100) / 100,
      subtotal:    Math.round(cli.subtotal * 100) / 100,
      num_compras: cli.num_compras,
      provincia:   cli.provincia,
      marcas:      Object.entries(cli.marcas).map(([marca, total]) => ({
        marca, total: Math.round(total * 100) / 100
      })).sort((a,b) => b.total - a.total),
      productos:   Object.values(cli.productos).map(p => ({
        id:       p.id,
        nombre:   p.nombre,
        codigo:   p.codigo,
        marca:    p.marca,
        cantidad: Math.round(p.cantidad),
        total:    Math.round(p.total * 100) / 100
      })).sort((a,b) => b.cantidad - a.cantidad),
      frecuencia:  Object.entries(cli.frecuencia).map(([mes, f]) => ({
        mes:      parseInt(mes),
        total:    Math.round(f.total * 100) / 100,
        subtotal: Math.round(f.subtotal * 100) / 100,
        compras:  f.compras
      })).sort((a,b) => a.mes - b.mes)
    })).sort((a,b) => b.total - a.total);

    resultado[vend.nombre] = clientes;
  });

  return resultado;
}

// Regenerar data.json automáticamente cada noche a las 2 AM
async function regenerarDataJsonAuto() {
  try {
    const ahora  = new Date();
    const inicio = new Date(ahora.getFullYear(), 0, 1); // 1 enero año actual
    const fi = fmtDateEC(inicio);
    const ff = fmtDateEC(ahora);
    console.log(`Regenerando data.json automáticamente: ${fi} → ${ff}`);
    const data = await generarDataJson(fi, ff);
    const ruta = require('path').join(__dirname, 'data.json');
    require('fs').writeFileSync(ruta, JSON.stringify(data, null, 2));
    console.log(`✓ data.json regenerado: ${Object.keys(data).length} vendedoras`);
  } catch(e) {
    console.error('Error regenerando data.json:', e.message);
  }
}

// Programar regeneración diaria a las 2 AM
function programarRegeneracion() {
  const ahora = new Date();
  const manana2am = new Date(ahora);
  manana2am.setDate(manana2am.getDate() + 1);
  manana2am.setHours(2, 0, 0, 0);
  const ms = manana2am - ahora;
  setTimeout(() => {
    regenerarDataJsonAuto();
    setInterval(regenerarDataJsonAuto, 24 * 60 * 60 * 1000);
  }, ms);
  console.log(`Próxima regeneración data.json en ${Math.round(ms/3600000)}h`);
}
programarRegeneracion();

// ─── DB ───────────────────────────────────────────────────────
async function initDB() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS visitas (
        id SERIAL PRIMARY KEY, lugar VARCHAR(255) NOT NULL,
        tipo VARCHAR(50) NOT NULL, asesora VARCHAR(255) NOT NULL,
        fecha TIMESTAMP DEFAULT NOW(), notas TEXT
      );
      CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY, nombre VARCHAR(255) NOT NULL,
        usuario VARCHAR(100) UNIQUE NOT NULL, password VARCHAR(255) NOT NULL,
        rol VARCHAR(50) DEFAULT 'asesora', modulos TEXT DEFAULT 'ventas,visitas,kpis,inventario',
        activo BOOLEAN DEFAULT true, created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS planificacion (
        id SERIAL PRIMARY KEY, asesora VARCHAR(255) NOT NULL,
        semana DATE NOT NULL, dia VARCHAR(20), sector VARCHAR(255),
        cliente VARCHAR(255), coordinado BOOLEAN DEFAULT false,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);
    const usuarios = [
      { nombre: 'Fernando Espíndola', usuario: 'Fernando', password: '1234', rol: 'admin', modulos: 'ventas,visitas,kpis,inventario,config' },
      { nombre: 'Giovanna Portilla', usuario: 'Giovanna', password: '1234', rol: 'jefa_ventas', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Daniela Villegas Chamorro', usuario: 'Daniela', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Liseth Gavilanes', usuario: 'Liseth', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Karen Rebeca Mora', usuario: 'Karen', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'María Caridad Zea', usuario: 'Maria', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
      { nombre: 'Nicole Yanira Leon', usuario: 'Nicole', password: '1234', rol: 'asesora', modulos: 'ventas,visitas,kpis,inventario' },
    ];
    for (const u of usuarios) {
      await pool.query(
        'INSERT INTO usuarios(nombre,usuario,password,rol,modulos) VALUES($1,$2,$3,$4,$5) ON CONFLICT(usuario) DO NOTHING',
        [u.nombre, u.usuario, u.password, u.rol, u.modulos]
      );
    }
    console.log('DB inicializada');
  } catch(e) { console.error('Error DB:', e.message); }
}
initDB();

const MIME = { '.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json','.png':'image/png','.jpg':'image/jpeg','.ico':'image/x-icon' };

function bodyJSON(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch(e) { reject(e); } });
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const urlObj = new URL(req.url, 'http://localhost');
  const urlPath = urlObj.pathname;
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // LOGIN
  if (urlPath === '/api/login' && req.method === 'POST') {
    try {
      const { usuario, password } = await bodyJSON(req);
      const r = await pool.query('SELECT * FROM usuarios WHERE usuario=$1 AND password=$2 AND activo=true', [usuario, password]);
      if (!r.rows.length) { res.writeHead(401,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Usuario o contraseña incorrectos'})); return; }
      const u = r.rows[0];
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:true, usuario:{id:u.id,nombre:u.nombre,usuario:u.usuario,rol:u.rol,modulos:u.modulos}}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // USUARIOS
  if (urlPath === '/api/usuarios' && req.method === 'GET') {
    try { const r = await pool.query('SELECT id,nombre,usuario,rol,modulos,activo FROM usuarios ORDER BY id'); res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows)); }
    catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/usuarios' && req.method === 'POST') {
    try {
      const {nombre,usuario,password,rol,modulos} = await bodyJSON(req);
      await pool.query('INSERT INTO usuarios(nombre,usuario,password,rol,modulos) VALUES($1,$2,$3,$4,$5)',[nombre,usuario,password||'1234',rol||'asesora',modulos||'ventas,visitas,kpis,inventario']);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath.startsWith('/api/usuarios/') && req.method === 'PUT') {
    try {
      const id = urlPath.split('/').pop();
      const body = await bodyJSON(req);
      if (body.password) await pool.query('UPDATE usuarios SET password=$1 WHERE id=$2',[body.password,id]);
      if (body.modulos!==undefined) await pool.query('UPDATE usuarios SET modulos=$1 WHERE id=$2',[body.modulos,id]);
      if (body.activo!==undefined) await pool.query('UPDATE usuarios SET activo=$1 WHERE id=$2',[body.activo,id]);
      if (body.rol!==undefined) await pool.query('UPDATE usuarios SET rol=$1 WHERE id=$2',[body.rol,id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath.startsWith('/api/usuarios/') && req.method === 'DELETE') {
    try {
      const id = urlPath.split('/').pop();
      await pool.query('DELETE FROM usuarios WHERE id=$1 AND rol!='admin'',[id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // VISITAS
  if (urlPath === '/api/visitas' && req.method === 'GET') {
    try { const r = await pool.query('SELECT * FROM visitas ORDER BY fecha DESC LIMIT 300'); res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows)); }
    catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/visitas' && req.method === 'POST') {
    try {
      const {lugar,tipo,asesora,notas} = await bodyJSON(req);
      const r = await pool.query('INSERT INTO visitas(lugar,tipo,asesora,notas) VALUES($1,$2,$3,$4) RETURNING *',[lugar,tipo,asesora,notas||null]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows[0]));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // PLANIFICACION
  if (urlPath === '/api/planificacion' && req.method === 'GET') {
    try {
      const asesora = urlObj.searchParams.get('asesora') || '';
      const semana = urlObj.searchParams.get('semana') || '';
      const r = await pool.query('SELECT * FROM planificacion WHERE asesora=$1 AND semana=$2 ORDER BY id',[asesora,semana]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/planificacion' && req.method === 'POST') {
    try {
      const {asesora,semana,filas} = await bodyJSON(req);
      if (!asesora||!semana||!filas) throw new Error('Faltan datos');
      await pool.query('DELETE FROM planificacion WHERE asesora=$1 AND semana=$2',[asesora,semana]);
      for (const fila of filas) {
        await pool.query('INSERT INTO planificacion(asesora,semana,dia,sector,cliente,coordinado) VALUES($1,$2,$3,$4,$5,$6)',
          [asesora,semana,fila.dia||'',fila.sector||'',fila.cliente||'',fila.coordinado||false]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true,filas:filas.length}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // VENTAS HOY (caché v2)
  if (urlPath === '/api/ventas-hoy' && req.method === 'GET') {
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ total: cache.documentos.length, ultima_sync: cache.ultima_sync, sincronizando: cache.sincronizando, documentos: cache.documentos }));
    return;
  }

  // SYNC MANUAL
  if (urlPath === '/api/sync' && req.method === 'GET') {
    sincronizarHoy().catch(e => console.error(e));
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({msg:'Sync iniciado', ultima_sync: cache.ultima_sync}));
    return;
  }

  // CATÁLOGO DE PRODUCTOS CON MARCAS
  // Devuelve { por_id: {id: {nombre,marca,codigo}}, por_nombre: {nombre: marca} }
  if (urlPath === '/api/productos-catalogo' && req.method === 'GET') {
    // Construir lookup por nombre también (para data.json histórico)
    const porNombre = {};
    Object.values(catalogoProductos).forEach(p => {
      if(p.nombre && p.marca) porNombre[p.nombre] = p.marca;
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      por_id:     catalogoProductos,
      por_nombre: porNombre,
      total:      Object.keys(catalogoProductos).length,
      synced_at:  catalogoSyncedAt
    }));
    return;
  }

  // SYNC MANUAL CATÁLOGO — ejecuta sincrónicamente y devuelve resultado detallado
  if (urlPath === '/api/sync-catalogo' && req.method === 'GET') {
    try {
      const url = 'https://api.contifico.com/sistema/api/v2/producto/?page_size=100';
      console.log('Probando catálogo URL:', url);
      const resp = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const txt = await resp.text();
      let parsed;
      try { parsed = JSON.parse(txt); } catch(e) { parsed = null; }
      
      const debug = {
        url, status: resp.status,
        api_key_presente: !!API_KEY,
        api_key_primeros_chars: API_KEY ? API_KEY.substring(0,20)+'...' : 'VACÍA',
        count: parsed?.count,
        results_length: parsed?.results?.length,
        raw_preview: txt.substring(0, 300),
        error_parse: parsed ? null : 'No se pudo parsear JSON'
      };

      if (parsed?.results?.length > 0) {
        // Procesar catálogo
        let nuevosCatalogo = {};
        parsed.results.forEach(p => {
          const id    = p.id || '';
          const nombre= (p.nombre || '').trim();
          const marca = (p.marca_nombre || p.marca || '').trim().toUpperCase();
          const codigo= (p.codigo || '').trim();
          if (id) nuevosCatalogo[id] = { nombre, marca, codigo };
        });
        // Cargar páginas restantes
        let nextUrl = parsed.next;
        while (nextUrl) {
          const r2 = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
          const d2 = await r2.json();
          (d2.results||[]).forEach(p => {
            if (p.id) nuevosCatalogo[p.id] = { nombre:(p.nombre||'').trim(), marca:(p.marca_nombre||p.marca||'').trim().toUpperCase(), codigo:(p.codigo||'').trim() };
          });
          nextUrl = d2.next;
        }
        catalogoProductos = nuevosCatalogo;
        catalogoSyncedAt = new Date().toISOString();
        debug.catalogoTotal = Object.keys(catalogoProductos).length;
        debug.muestra = Object.entries(catalogoProductos).slice(0,3).map(([id,p])=>({id,...p}));
        debug.resultado = 'OK';
      } else {
        debug.resultado = 'Sin resultados';
      }
      
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify(debug, null, 2));
    } catch(e) {
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ error: e.message, stack: e.stack?.substring(0,500) }));
    }
    return;
  }

  // DIAGNÓSTICO CATÁLOGO — prueba diferentes endpoints de productos
  if (urlPath === '/api/test-catalogo' && req.method === 'GET') {
    try {
      const resultados = {};
      // Solo v2 que funciona — mostrar item completo
      const url = 'https://api.contifico.com/sistema/api/v2/producto/?page_size=3';
      const r = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await r.json();
      resultados['v2'] = {
        status: r.status,
        count: data.count,
        campos_disponibles: data.results?.[0] ? Object.keys(data.results[0]) : [],
        items: data.results?.map(p => JSON.stringify(p, null, 2))
      };
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify(resultados, null, 2));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // REGENERAR DATA.JSON MANUAL — solo admin puede llamar esto
  if (urlPath === '/api/regenerar-data' && req.method === 'GET') {
    const fi = urlObj.searchParams.get('desde') || fmtDateEC(new Date(new Date().getFullYear(),0,1));
    const ff = urlObj.searchParams.get('hasta') || fmtDateEC(new Date());
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({msg:`Regenerando data.json del ${fi} al ${ff}...`, ok:true}));
    // Ejecutar en background
    generarDataJson(fi, ff).then(data => {
      const ruta = require('path').join(__dirname, 'data.json');
      require('fs').writeFileSync(ruta, JSON.stringify(data, null, 2));
      console.log(`✓ data.json regenerado manualmente: ${Object.keys(data).length} vendedoras`);
    }).catch(e => console.error('Error regenerar manual:', e.message));
    return;
  }

  // TEST V2 — muestra primer documento CLI con detalles completos
  if (urlPath === '/api/test-v2' && req.method === 'GET') {
    try {
      const fechaParam = urlObj.searchParams.get('fecha');
      const testFecha = fechaParam || fmtDateEC(new Date());
      const testUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${testFecha}&fecha_final=${testFecha}&page_size=100`;
      const response = await fetch(testUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await response.json();
      // Buscar primer documento de CLIENTE
      const cli = (data.results||[]).find(d => d.tipo_registro === 'CLI');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        count_total: data.count,
        count_resultados: (data.results||[]).length,
        primer_cli: cli ? JSON.stringify(cli, null, 2) : 'No hay documentos CLI en esta fecha'
      }));
    } catch(e) { res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // CHAT
  if (urlPath === '/api/chat' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
      if (!ANTHROPIC_KEY) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Sin API Key'})); return; }
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':ANTHROPIC_KEY,'anthropic-version':'2023-06-01'},
        body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:1024,system:body.system||'',messages:body.messages||[]})
      });
      const data = await response.json();
      res.writeHead(response.status,{'Content-Type':'application/json'}); res.end(JSON.stringify(data));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // STATIC FILES
  let filePath = urlPath==='/' ? path.join(__dirname,'index.html')
    : urlPath==='/login' ? path.join(__dirname,'login.html')
    : urlPath==='/bot' ? path.join(__dirname,'bot.html')
    : urlPath==='/sofia.jpg' ? path.join(__dirname,'sofia.jpg')
    : path.join(__dirname, urlPath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    res.writeHead(200,{'Content-Type':MIME[path.extname(filePath)]||'text/plain'});
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200,{'Content-Type':'text/html'});
    fs.createReadStream(path.join(__dirname,'index.html')).pipe(res);
  }
});

server.listen(PORT,'0.0.0.0',()=>console.log(`Cosétika Dashboard running on port ${PORT}`));
