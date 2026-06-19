const http = require('http');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

// Inferir provincia desde dirección — mapa basado en datos reales de clientes Cosétika
// Ordenado de más específico a más general (longer match first)
const CIUDAD_PROV_ENTRIES = [
  ['GUAYAQUIL KENNEDY NORTE','GUAYAS'],['GUAYAQUIL KENNEDY','GUAYAS'],['GUAYAQUIL CDLA','GUAYAS'],
  ['GUAYAQUIL.','GUAYAS'],['BASTIÓN POPULAR','GUAYAS'],['KENNEDY NORTE','GUAYAS'],
  ['REPUBLICA DEL SALVADOR','PICHINCHA'],['REPUBLICA DEL','PICHINCHA'],
  ['AV. ELOY ALFARO','PICHINCHA'],['ELOY ALFARO','PICHINCHA'],
  ['MARISCAL SUCRE Y','PICHINCHA'],['MARISCAL SUCRE','PICHINCHA'],
  ['QUITO VALLE DE','PICHINCHA'],['QUITO VALLE','PICHINCHA'],
  ['QUITO AV.','PICHINCHA'],['QUITO AV','PICHINCHA'],
  ['TUMBACO CALLE','PICHINCHA'],['VALLE DE LOS','PICHINCHA'],['VALLE DE','PICHINCHA'],
  ['AV. ELOY','PICHINCHA'],['AV DE LOS','PICHINCHA'],['AV DE','PICHINCHA'],
  ['SANTO DOMINGO DE LOS','SANTO DOMINGO'],['SANTO DOMINGO','SANTO DOMINGO'],
  ['LAGO AGRIO','SUCUMBÍOS'],['LA TRONCAL','CAÑAR'],['EL EMPALME','GUAYAS'],
  ['LUIS CORDERO','PICHINCHA'],['LAS CASAS','PICHINCHA'],
  ['AMBATO AV','TUNGURAHUA'],['AMBATO.','TUNGURAHUA'],
  ['MACHALA.','EL ORO'],['MANTA','MANABÍ'],['PORTOVIEJO','MANABÍ'],
  ['GYE VILLA','GUAYAS'],['GYE.','GUAYAS'],
  ['ALBORADA','GUAYAS'],['BASTIÓN','GUAYAS'],['CDLA.','GUAYAS'],
  ['URDESA','GUAYAS'],['KENNEDY','GUAYAS'],
  ['GUAYAQUIL','GUAYAS'],
  ['CUMBAYA','PICHINCHA'],['CUMBAYÁ','PICHINCHA'],['CONOCOTO','PICHINCHA'],
  ['QUITUMBE','PICHINCHA'],['POMASQUI','PICHINCHA'],['SANGOLQUI','PICHINCHA'],
  ['CARCELEN','PICHINCHA'],['TUMBACO','PICHINCHA'],['PUEMBO','PICHINCHA'],
  ['MACHACHI','PICHINCHA'],['SHYRIS Y','PICHINCHA'],['SHYRIS','PICHINCHA'],
  ['LLANO','PICHINCHA'],['QUITO','PICHINCHA'],
  ['CUENCA','AZUAY'],
  ['AMBATO','TUNGURAHUA'],['RIOBAMBA','CHIMBORAZO'],
  ['IBARRA','IMBABURA'],['OTAVALO','IMBABURA'],
  ['QUEVEDO','LOS RÍOS'],['LOS RIOS','LOS RÍOS'],
  ['LAGO','SUCUMBÍOS'],['SUCUMBIOS','SUCUMBÍOS'],
  ['CAYAMBE','PICHINCHA'],['PUJILI','COTOPAXI'],['LATACUNGA','COTOPAXI'],
  ['BABAHOYO','LOS RÍOS'],['MACHALA','EL ORO'],
  ['ESMERALDAS','ESMERALDAS'],['LOJA','LOJA'],
  ['TULCAN','CARCHI'],['TULCÁN','CARCHI'],
  ['GUARANDA','BOLÍVAR'],['AZOGUES','CAÑAR'],
  ['SANTA ELENA','SANTA ELENA'],['SALINAS','SANTA ELENA'],
  ['TENA','NAPO'],['PUYO','PASTAZA'],['MACAS','MORONA SANTIAGO'],
  ['ZAMORA','ZAMORA CHINCHIPE'],
  ['NUEVA LOJA','SUCUMBÍOS'],['COCA','ORELLANA'],
];

const PROVINCIAS_NOMBRE = ['PICHINCHA','GUAYAS','AZUAY','TUNGURAHUA','CHIMBORAZO','LOJA',
  'IMBABURA','CARCHI','COTOPAXI','LOS RÍOS','MANABÍ','EL ORO','ESMERALDAS',
  'SANTO DOMINGO','PASTAZA','NAPO','SUCUMBÍOS','ORELLANA','MORONA SANTIAGO',
  'ZAMORA CHINCHIPE','BOLÍVAR','CAÑAR','SANTA ELENA','GALÁPAGOS'];

function provinciaDesdeDir(dir){
  if(!dir) return '';
  const d = dir.toUpperCase();
  for(const [ciudad, prov] of CIUDAD_PROV_ENTRIES){
    if(d.startsWith(ciudad) || d.includes(' '+ciudad) || d.includes(','+ciudad)) return prov;
  }
  for(const prov of PROVINCIAS_NOMBRE){
    if(d.includes(prov)) return prov;
  }
  return '';
}

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// ─── CACHÉ v2 ────────────────────────────────────────────────
let cache = { documentos: [], ultima_sync: null, sincronizando: false };

// ─── CATÁLOGO PRODUCTOS ───────────────────────────────────────
let catalogoProductos = {};
let catalogoSyncedAt = null;

// ─── CATÁLOGO DE CLIENTES (id → provincia) ───────────────────
let catalogoClientes = {}; // { persona_id: provincia }
let catalogoClientesSyncedAt = null;

async function sincronizarCatalogoClientes() {
  try {
    console.log('Sincronizando catálogo de clientes...');
    let nuevoCatalogo = {};
    let nextUrl = 'https://api.contifico.com/sistema/api/v1/persona/?es_cliente=true&page_size=100';
    let paginas = 0;
    while (nextUrl && paginas < 50) {
      const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!resp.ok) break;
      const data = await resp.json();
      const results = Array.isArray(data) ? data : (data.results || []);
      results.forEach(p => {
        if (p.id && p.direccion) {
          nuevoCatalogo[p.id] = provinciaDesdeDir(p.direccion);
        }
      });
      nextUrl = Array.isArray(data) ? null : (data.next || null);
      paginas++;
    }
    if (Object.keys(nuevoCatalogo).length > 0) {
      catalogoClientes = nuevoCatalogo;
      catalogoClientesSyncedAt = new Date().toISOString();
      console.log('Catálogo clientes: ' + Object.keys(catalogoClientes).length + ' clientes con provincia');
    }
  } catch(e) { console.error('Error catálogo clientes:', e.message); }
}
sincronizarCatalogoClientes().catch(e => console.error(e));
setInterval(() => sincronizarCatalogoClientes().catch(e => console.error(e)), 24 * 60 * 60 * 1000);

async function sincronizarCatalogo() {
  try {
    let nuevosCatalogo = {};
    let nextUrl = 'https://api.contifico.com/sistema/api/v2/producto/?page_size=100';
    let paginas = 0;
    while (nextUrl && paginas < 50) {
      const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!resp.ok) break;
      const data = await resp.json();
      (data.results || []).forEach(p => {
        if (p.id) nuevosCatalogo[p.id] = {
          nombre: (p.nombre || '').trim(),
          marca:  (p.marca_nombre || p.marca || '').trim().toUpperCase(),
          codigo: (p.codigo || '').trim()
        };
      });
      nextUrl = data.next || null;
      paginas++;
    }
    if (Object.keys(nuevosCatalogo).length > 0) {
      catalogoProductos = nuevosCatalogo;
      catalogoSyncedAt = new Date().toISOString();
      console.log('Catálogo: ' + Object.keys(catalogoProductos).length + ' productos');
    }
  } catch(e) { console.error('Error catálogo:', e.message); }
}
sincronizarCatalogo().catch(e => console.error(e));
setInterval(() => sincronizarCatalogo().catch(e => console.error(e)), 24 * 60 * 60 * 1000);

// ─── GENERADOR DATA.JSON ──────────────────────────────────────
async function generarDataJson(fi, ff) {
  const vendedores = {};
  let nextUrl = 'https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=' + fi + '&fecha_final=' + ff + '&page_size=100';
  let paginas = 0;
  while (nextUrl && paginas < 200) {
    const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
    if (!resp.ok) break;
    const data = await resp.json();
    const docs = (data.results || []).filter(d => {
      if (d.tipo_registro !== 'CLI') return false;  // solo clientes
      if (d.anulado) return false;                   // excluir anulados
      if (d.tipo_documento === 'NC') return false;   // excluir notas de crédito
      // Si no hay objeto vendedor pero hay identificación, lo incluimos como "Sin asignar"
      if (!d.vendedor && !d.vendedor_id && !d.vendedor_identificacion) return false;
      // Excluir autoconsumo: facturas al cliente Corporación Cosétika (RUC 1793143660001)
      const cliRuc = (d.cliente?.ruc || d.cliente?.cedula || '').trim();
      if (cliRuc === '1793143660001') return false;
      return true;
    });
    docs.forEach(doc => {
      const vendId = doc.vendedor?.id || doc.vendedor_identificacion || 'sin_vendedor';
      const vendNom = doc.vendedor?.razon_social || ('Vendedor ' + (doc.vendedor_identificacion || 'Sin Asignar'));
      const cliId = doc.cliente && doc.cliente.id ? doc.cliente.id : doc.persona_id;
      const cliNom = (doc.cliente && (doc.cliente.razon_social || doc.cliente.nombre_comercial)) || '—';
      const cliRuc = (doc.cliente && (doc.cliente.ruc || doc.cliente.cedula)) || '';
      // Buscar provincia: primero en catálogo, luego inferir de dirección
      const cliProv = catalogoClientes[cliId] || provinciaDesdeDir(doc.cliente?.direccion || '');
      const mes = parseInt((doc.fecha_emision || '').split('/')[1]) || 0;
      const totalDoc = parseFloat(doc.total || 0);
      const subDoc = parseFloat(doc.subtotal || doc.subtotal_12 || 0);
      if (!cliId || totalDoc === 0) return;
      if (!vendedores[vendId]) vendedores[vendId] = { nombre: vendNom, clientes: {} };
      vendedores[vendId].nombre = vendNom;
      const vObj = vendedores[vendId].clientes;
      if (!vObj[cliId]) vObj[cliId] = { id: cliId, nombre: cliNom, ruc: cliRuc, total: 0, subtotal: 0, num_compras: 0, provincia: cliProv, marcas: {}, productos: {}, frecuencia: {} };
      const cli = vObj[cliId];
      cli.nombre = cliNom; cli.ruc = cliRuc;
      if(!cli.provincia && cliProv) cli.provincia = cliProv;
      cli.total += totalDoc; cli.subtotal += subDoc; cli.num_compras++;
      const anioDoc = parseInt((doc.fecha_emision || '').split('/')[2]) || new Date().getFullYear();
      const freqKey = `${anioDoc}-${String(mes).padStart(2,'0')}`;
      if (!cli.frecuencia[freqKey]) cli.frecuencia[freqKey] = { anio: anioDoc, mes, total: 0, subtotal: 0, compras: 0 };
      cli.frecuencia[freqKey].total += totalDoc; cli.frecuencia[freqKey].subtotal += subDoc; cli.frecuencia[freqKey].compras++;
      (doc.detalles || []).forEach(det => {
        const prodId = det.producto_id || '';
        const cantidad = parseFloat(det.cantidad || 0);
        const base = parseFloat(det.base_gravable || det.base_cero || 0);
        if (!prodId || cantidad === 0 || base === 0) return;
        const cat = catalogoProductos[prodId] || {};
        const marca = cat.marca || '';
        const nom = cat.nombre || det.producto_nombre || '';
        if (!cli.productos[prodId]) cli.productos[prodId] = { id: prodId, nombre: nom, codigo: cat.codigo || '', marca, cantidad: 0, total: 0 };
        cli.productos[prodId].nombre = nom;
        cli.productos[prodId].cantidad += cantidad;
        cli.productos[prodId].total += base;
        if (marca) cli.marcas[marca] = (cli.marcas[marca] || 0) + base;
      });
    });
    nextUrl = data.next || null;
    paginas++;
  }
  const resultado = {};
  Object.values(vendedores).forEach(vend => {
    resultado[vend.nombre] = Object.values(vend.clientes).map(cli => ({
      id: cli.id, nombre: cli.nombre, ruc: cli.ruc,
      total: Math.round(cli.total * 100) / 100,
      subtotal: Math.round(cli.subtotal * 100) / 100,
      num_compras: cli.num_compras, provincia: cli.provincia,
      marcas: Object.entries(cli.marcas).map(([m,t]) => ({ marca: m, total: Math.round(t*100)/100 })).sort((a,b) => b.total-a.total),
      productos: Object.values(cli.productos).map(p => ({ id: p.id, nombre: p.nombre, codigo: p.codigo, marca: p.marca, cantidad: Math.round(p.cantidad), total: Math.round(p.total*100)/100 })).sort((a,b) => b.cantidad-a.cantidad),
      frecuencia: Object.values(cli.frecuencia).map(f => ({ anio: f.anio, mes: f.mes, total: Math.round(f.total*100)/100, subtotal: Math.round(f.subtotal*100)/100, compras: f.compras })).sort((a,b) => a.anio!==b.anio ? a.anio-b.anio : a.mes-b.mes)
    })).sort((a,b) => b.total-a.total);
  });
  return resultado;
}

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
    const clientes = todos.filter(d => d.tipo_registro === 'CLI' && !d.anulado && d.tipo_documento !== 'NC');
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

// ─── DB ───────────────────────────────────────────────────────
// ─── CACHÉ DE DATA EN MEMORIA (cargada desde PostgreSQL al arrancar) ─────────
let DATA_CACHE = null;
let DATA_CACHE_TS = null;

async function cargarDataDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM ventas_data ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      DATA_CACHE = JSON.parse(r.rows[0].datos);
      DATA_CACHE_TS = new Date().toISOString();
      console.log('✓ DATA cargada desde PostgreSQL: ' + Object.keys(DATA_CACHE).length + ' vendedoras');
    } else {
      // Fallback: cargar desde data.json si existe
      try {
        const raw = fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8');
        DATA_CACHE = JSON.parse(raw);
        DATA_CACHE_TS = new Date().toISOString();
        console.log('✓ DATA cargada desde data.json: ' + Object.keys(DATA_CACHE).length + ' vendedoras');
      } catch(e) { console.log('Sin data.json, esperando regeneración'); DATA_CACHE = {}; }
    }
  } catch(e) { console.error('Error cargando DATA:', e.message); DATA_CACHE = {}; }
}

async function guardarDataEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO ventas_data (datos, actualizado_at) VALUES ($1, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, actualizado_at = NOW()
    `, [json]);
    DATA_CACHE = data;
    DATA_CACHE_TS = new Date().toISOString();
    console.log('✓ DATA guardada en PostgreSQL');
  } catch(e) { console.error('Error guardando DATA en DB:', e.message); }
}

async function initDB() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS ventas_data (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
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
initDB().then(() => cargarDataDesdeDB()).catch(e => console.error('Error init:', e.message));

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
      await pool.query("DELETE FROM usuarios WHERE id=$1 AND rol!='admin'",[id]);
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

  // TEST V2
  if (urlPath === '/api/test-v2' && req.method === 'GET') {
    try {
      const now = new Date();
      const testFecha = fmtDateEC(now);
      const testUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${testFecha}&fecha_final=${testFecha}`;
      console.log('Testing v2:', testUrl);
      const inicio = Date.now();
      const response = await fetch(testUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const tiempo = Date.now() - inicio;
      const texto = await response.text();
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ url_probada: testUrl, status: response.status, tiempo_ms: tiempo, tiempo_seg: (tiempo/1000).toFixed(1)+'s', respuesta_preview: texto.substring(0,500) }));
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

  // CATÁLOGO PRODUCTOS
  if (urlPath === '/api/productos-catalogo' && req.method === 'GET') {
    const porNombre = {};
    Object.values(catalogoProductos).forEach(p => { if(p.nombre && p.marca) porNombre[p.nombre] = p.marca; });
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ por_id: catalogoProductos, por_nombre: porNombre, total: Object.keys(catalogoProductos).length, synced_at: catalogoSyncedAt }));
    return;
  }

  // SYNC CATÁLOGO MANUAL
  if (urlPath === '/api/sync-catalogo' && req.method === 'GET') {
    sincronizarCatalogo().catch(e => console.error(e));
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ msg: 'Sync iniciado', total: Object.keys(catalogoProductos).length, synced_at: catalogoSyncedAt }));
    return;
  }

  // DIAGNÓSTICO DE TOTALES vs CONTIFICO
  if (urlPath === '/api/diagnostico-ventas' && req.method === 'GET') {
    try {
      const fi = urlObj.searchParams.get('desde') || '01/01/2026';
      const ff = urlObj.searchParams.get('hasta') || fmtDateEC(new Date());
      let totalCLI=0, totalNC=0, totalFernando=0, totalAnulado=0;
      let countCLI=0, countNC=0, countFernando=0, countAnulado=0;
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fi}&fecha_final=${ff}&page_size=100`;
      let paginas=0;
      while(nextUrl && paginas<200){
        const resp = await fetch(nextUrl, { headers:{'Authorization':API_KEY,'Accept':'application/json'} });
        const data = await resp.json();
        (data.results||[]).forEach(d=>{
          const total = parseFloat(d.total||0);
          if(d.anulado){ totalAnulado+=total; countAnulado++; return; }
          if(d.tipo_registro==='CLI'){
            const vend = d.vendedor?.razon_social||'';
            if(vend.includes('Fernando')||vend.includes('Espíndola')||vend.includes('Espindola')){
              totalFernando+=total; countFernando++;
            } else {
              totalCLI+=total; countCLI++;
            }
          } else if(d.tipo_documento==='NC'||d.tipo_registro==='NC'){
            totalNC+=total; countNC++;
          }
        });
        nextUrl=data.next||null; paginas++;
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        periodo: `${fi} → ${ff}`,
        ventas_clientes: {total: Math.round(totalCLI*100)/100, facturas: countCLI},
        ventas_fernando: {total: Math.round(totalFernando*100)/100, facturas: countFernando},
        notas_credito:   {total: Math.round(totalNC*100)/100, docs: countNC},
        anulados:        {total: Math.round(totalAnulado*100)/100, docs: countAnulado},
        neto_esperado:   Math.round((totalCLI)*100)/100
      },null,2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error:e.message}));
    }
    return;
  }

  // REGENERAR DATA.JSON
  if (urlPath === '/api/regenerar-data' && req.method === 'GET') {
    const fi = urlObj.searchParams.get('desde') || fmtDateEC(new Date(new Date().getFullYear(),0,1));
    const ff = urlObj.searchParams.get('hasta') || fmtDateEC(new Date());
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ msg: 'Regenerando data.json del ' + fi + ' al ' + ff, ok: true }));
    generarDataJson(fi, ff).then(async data => {
      await guardarDataEnDB(data);
      // También actualizar data.json como backup
      try {
        fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(data, null, 2));
      } catch(e) { /* OK si no se puede escribir */ }
      console.log('✓ Regeneración completada: ' + Object.keys(data).length + ' vendedoras');
    }).catch(e => console.error('Error regenerar:', e.message));
    return;
  }

  // DATA.JSON desde caché en memoria (PostgreSQL)
  if (urlPath === '/data.json') {
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify(DATA_CACHE || {}));
    return;
  }

  // VER VENDEDORES EN CONTIFICO
  if (urlPath === '/api/ver-vendedores' && req.method === 'GET') {
    try {
      const desde = urlObj.searchParams.get('desde') || '01/01/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(new Date());
      const url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      const resp = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await resp.json();
      const vendedores = {};
      (data.results||[]).filter(d=>d.tipo_registro==='CLI'&&!d.anulado&&d.tipo_documento!=='NC').forEach(d=>{
        const vNom = d.vendedor?.razon_social || 'SIN VENDEDOR';
        if(!vendedores[vNom]) vendedores[vNom]={facturas:0,total:0};
        vendedores[vNom].facturas++;
        vendedores[vNom].total+=parseFloat(d.total||0);
      });
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({periodo:`${desde}→${hasta}`,vendedores},null,2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error:e.message}));
    }
    return;
  }

  // BUSCAR CLIENTE O VENDEDOR EN CONTIFICO
  if (urlPath === '/api/buscar-cliente' && req.method === 'GET') {
    try {
      const nombre = urlObj.searchParams.get('q') || 'cosetika';
      const desde = urlObj.searchParams.get('desde') || '01/06/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(new Date());
      // Paginar para obtener más resultados
      let encontrados = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while(nextUrl && paginas < 5) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        const data = await resp.json();
        const filtrados = (data.results||[]).filter(d => {
          const cliNom = (d.cliente?.razon_social || d.cliente?.nombre_comercial || '').toLowerCase();
          const vendNom = (d.vendedor?.razon_social || '').toLowerCase();
          return cliNom.includes(nombre.toLowerCase()) || vendNom.includes(nombre.toLowerCase());
        }).map(d => ({
          documento: d.documento,
          tipo_registro: d.tipo_registro,
          tipo_doc: d.tipo_documento,
          fecha: d.fecha_emision,
          cliente: d.cliente?.razon_social,
          cliente_ruc: d.cliente?.ruc || d.cliente?.cedula,
          vendedor: d.vendedor?.razon_social,
          vendedor_obj: d.vendedor ? 'existe' : 'NULL',
          total: d.total,
          anulado: d.anulado
        }));
        encontrados = encontrados.concat(filtrados);
        nextUrl = data.next || null;
        paginas++;
      }
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ busqueda: nombre, encontrados }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // VER DATA DE UN VENDEDOR ESPECÍFICO
  if (urlPath === '/api/ver-vendedor' && req.method === 'GET') {
    const nombre = urlObj.searchParams.get('nombre') || 'Fernando';
    const encontrado = Object.entries(DATA_CACHE||{}).find(([k])=>k.toLowerCase().includes(nombre.toLowerCase()));
    res.writeHead(200,{'Content-Type':'application/json'});
    if(!encontrado){
      const vendedores = Object.keys(DATA_CACHE||{});
      res.end(JSON.stringify({error:'No encontrado', vendedores_disponibles: vendedores}));
    } else {
      const [nombre_real, clientes] = encontrado;
      const total2026 = clientes.reduce((a,c)=>a+c.frecuencia.filter(f=>f.anio===2026).reduce((s,f)=>s+f.total,0),0);
      const total2025 = clientes.reduce((a,c)=>a+c.frecuencia.filter(f=>f.anio===2025).reduce((s,f)=>s+f.total,0),0);
      // Clientes con ventas en 2026
      const clis2026 = clientes.filter(c=>c.frecuencia.some(f=>f.anio===2026&&f.total>0));
      res.end(JSON.stringify({
        vendedor: nombre_real,
        total_clientes: clientes.length,
        total_2026: Math.round(total2026*100)/100,
        total_2025: Math.round(total2025*100)/100,
        clientes_con_ventas_2026: clis2026.length,
        detalle_2026: clis2026.map(c=>({
          nombre: c.nombre,
          frecuencia_2026: c.frecuencia.filter(f=>f.anio===2026)
        }))
      }));
    }
    return;
  }

  // VER CAMPOS DE CLIENTE EN CONTIFICO
  if (urlPath === '/api/ver-cliente-campos' && req.method === 'GET') {
    try {
      // Buscar persona por cédula para ver campos provincia/canton
      const urls = [
        'https://api.contifico.com/sistema/api/v1/persona/?cedula=1207822287&page_size=1',
        'https://api.contifico.com/sistema/api/v2/persona/?cedula=1207822287&page_size=1',
        'https://api.contifico.com/sistema/api/v1/persona/BleXkGyPWij1JdrN/',
        'https://api.contifico.com/sistema/api/v2/persona/BleXkGyPWij1JdrN/',
      ];
      // Buscar todos los campos del endpoint de persona v1
      const url = 'https://api.contifico.com/sistema/api/v1/persona/?es_cliente=true&page_size=2';
      const respP = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const dataP = await respP.json();
      const primer = Array.isArray(dataP) ? dataP[0] : dataP.results?.[0] || dataP[0];
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        campos: primer ? Object.keys(primer) : [],
        primer_cliente: primer,
        total: dataP.count || (Array.isArray(dataP) ? dataP.length : '?')
      }, null, 2));
      return;
      const url2 = url;
      const resp = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await resp.json();
      const cli = (data.results||[]).find(d=>d.tipo_registro==='CLI'&&d.cliente);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        cliente_completo: cli?.cliente,
        campos_cliente: cli?.cliente ? Object.keys(cli.cliente) : []
      }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error:e.message}));
    }
    return;
  }

  // LISTAR VENDEDORES EXACTOS EN DATA_CACHE
  if (urlPath === '/api/lista-vendedores') {
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({
      vendedores: Object.keys(DATA_CACHE||{}),
      totales: Object.entries(DATA_CACHE||{}).map(([v,clientes])=>({
        vendedor: v,
        clientes: clientes.length,
        total_2026: Math.round(clientes.reduce((a,c)=>a+c.frecuencia.filter(f=>f.anio===2026).reduce((s,f)=>s+f.total,0),0)*100)/100
      }))
    }));
    return;
  }

  // ESTADO DE LA DATA
  if (urlPath === '/api/data-status') {
    const muestra = {};
    Object.entries(DATA_CACHE||{}).slice(0,2).forEach(([v,clientes])=>{
      muestra[v] = {
        clientes: clientes.length,
        ejemplo_frecuencia: clientes[0]?.frecuencia?.slice(0,3) || []
      };
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      vendedoras: Object.keys(DATA_CACHE||{}).length,
      actualizado: DATA_CACHE_TS,
      fuente: DATA_CACHE && Object.keys(DATA_CACHE).length > 0 ? 'postgresql' : 'vacia',
      muestra_estructura: muestra
    }));
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
