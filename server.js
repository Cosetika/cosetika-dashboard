const http = require('http');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');
const XLSX = require('xlsx');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
let webpush; try { webpush = require('web-push'); } catch(e) { console.log('web-push no instalado'); }

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';
const VAPID_PUBLIC_KEY = process.env.VAPID_PUBLIC_KEY || '';
const VAPID_PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY || '';
if (webpush && VAPID_PUBLIC_KEY && VAPID_PRIVATE_KEY) {
  webpush.setVapidDetails('mailto:info@cosetika.com', VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);
  console.log('✓ Web Push VAPID configurado');
}
// Credenciales de la casilla pedidos@cosetika.com — configuradas como variables de
// entorno en Railway, nunca hardcodeadas en el código.
const PEDIDOS_EMAIL_HOST = process.env.PEDIDOS_EMAIL_HOST || '';
const PEDIDOS_EMAIL_USER = process.env.PEDIDOS_EMAIL_USER || '';
const PEDIDOS_EMAIL_PASS = process.env.PEDIDOS_EMAIL_PASS || '';
const PEDIDOS_EMAIL_PORT = parseInt(process.env.PEDIDOS_EMAIL_PORT) || 993;
// Casilla de referidos (formulario web). Host/puerto heredan de la de pedidos si no se definen.
// .trim() elimina espacios/saltos de línea invisibles que se cuelan al copiar y pegar en Railway
const REFERIDOS_EMAIL_HOST = (process.env.REFERIDOS_EMAIL_HOST || process.env.PEDIDOS_EMAIL_HOST || '').trim();
const REFERIDOS_EMAIL_USER = (process.env.REFERIDOS_EMAIL_USER || '').trim();
const REFERIDOS_EMAIL_PASS = (process.env.REFERIDOS_EMAIL_PASS || '').trim();
const REFERIDOS_EMAIL_PORT = parseInt(process.env.REFERIDOS_EMAIL_PORT) || parseInt(process.env.PEDIDOS_EMAIL_PORT) || 993;


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

let categoriasContifico = {};
async function sincronizarCategorias() {
  try {
    const mapa = {};
    for (const v of ['v1','v2']) {
      let next = `https://api.contifico.com/sistema/api/${v}/categoria/?page_size=100`;
      let pg = 0;
      while (next && pg < 30) {
        const r = await fetch(next, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!r.ok) break;
        const d = await r.json();
        const lista = Array.isArray(d) ? d : (d.results || []);
        lista.forEach(c => { if (c && c.id) mapa[c.id] = String(c.nombre || c.descripcion || '').trim(); });
        next = (d && d.next) || null; pg++;
      }
      if (Object.keys(mapa).length) break;
    }
    if (Object.keys(mapa).length) { categoriasContifico = mapa; console.log('Categorías: ' + Object.keys(mapa).length); }
  } catch(e) { console.error('Error categorías:', e.message); }
}

async function sincronizarCatalogo() {
  try {
    await sincronizarCategorias();
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
          codigo: (p.codigo || '').trim(),
          // Precios de lista de Contifico (para las prefacturas) + IVA del producto
          pvp1: parseFloat(p.pvp1) || 0,
          pvp2: parseFloat(p.pvp2) || 0,
          pvp3: parseFloat(p.pvp3) || 0,
          pvp4: parseFloat(p.pvp4) || 0,
          iva:  (p.porcentaje_iva === 0 ? 0 : 15),
          // Costo: Contifico expone el campo con distintos nombres según la versión de la
          // ficha. Tomamos el primero con valor y dejamos anotado cuál fue, para poder
          // auditarlo desde el panel sin adivinar.
          costo: (function(){
            const cands = [['costo_promedio',p.costo_promedio],['costo',p.costo],['costo_produccion',p.costo_produccion],['costo_ultima_compra',p.costo_ultima_compra],['precio_costo',p.precio_costo]];
            for (const [k,v] of cands) { const n = parseFloat(v); if (n > 0) return n; }
            return 0;
          })(),
          costo_campo: (function(){
            const cands = [['costo_promedio',p.costo_promedio],['costo',p.costo],['costo_produccion',p.costo_produccion],['costo_ultima_compra',p.costo_ultima_compra],['precio_costo',p.precio_costo]];
            for (const [k,v] of cands) { if (parseFloat(v) > 0) return k; }
            return null;
          })(),
          categoria: String(p.categoria_nombre || categoriasContifico[p.categoria_id] || (p.categoria && (p.categoria.nombre || p.categoria)) || '').trim(),
          categoria_id: p.categoria_id || null,
          estado: String(p.estado || '').trim(),
          tipo: String(p.tipo || '').trim()
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
  // Reasignaciones manuales: el cliente (todo su historial) se agrupa bajo la vendedora destino
  const REASIG = {};
  try {
    const rr = await pool.query('SELECT cliente_ruc, vendedora_destino FROM clientes_reasignados');
    rr.rows.forEach(x => { const k = String(x.cliente_ruc||'').replace(/\D/g,''); if (k) REASIG[k] = x.vendedora_destino; });
    if (rr.rows.length) console.log(`Reasignaciones activas: ${rr.rows.length}`);
  } catch(e) { console.log('Sin reasignaciones:', e.message); }
  const documentosVistos = new Set(); // evita procesar el mismo documento dos veces (ej: si la API repite en paginación)
  let nextUrl = 'https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=' + fi + '&fecha_final=' + ff + '&page_size=100';
  let paginas = 0;
  let duplicadosOmitidos = 0;
  while (nextUrl && paginas < 500) {
    const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
    if (!resp.ok) break;
    const data = await resp.json();
    const docs = (data.results || []).filter(d => {
      if (d.tipo_registro !== 'CLI') return false;  // solo clientes
      if (d.anulado) return false;                   // excluir anulados
      // Las notas de crédito (NC) SÍ pasan: se restan del total (neteo, igual que Contifico)
      if (noEsVenta(d)) return false;  // cotizaciones, proformas y PREFACTURAS no son ventas
      // Si no hay objeto vendedor pero hay identificación, lo incluimos como "Sin asignar"
      if (!d.vendedor && !d.vendedor_id && !d.vendedor_identificacion) return false;
      // Excluir autoconsumo: facturas al cliente Corporación Cosétika (RUC 1793143660001)
      const cliRuc = (d.cliente?.ruc || d.cliente?.cedula || '').trim();
      if (cliRuc === '1793143660001') return false;
      // Evitar procesar el mismo documento dos veces
      const docKey = d.id || d.documento;
      if (documentosVistos.has(docKey)) { duplicadosOmitidos++; return false; }
      documentosVistos.add(docKey);
      return true;
    });
    docs.forEach(doc => {
      let vendId = doc.vendedor?.id || doc.vendedor_identificacion || 'sin_vendedor';
      let vendNom = doc.vendedor?.razon_social || ('Vendedor ' + (doc.vendedor_identificacion || 'Sin Asignar'));
      // Reasignación manual de cliente → vendedora (Configuración → Sincronización)
      const rucReasig = String((doc.cliente && (doc.cliente.ruc || doc.cliente.cedula)) || '').replace(/\D/g,'');
      if (rucReasig && REASIG[rucReasig]) { vendNom = REASIG[rucReasig]; vendId = 'reasig::' + vendNom; }
      const cliId = doc.cliente && doc.cliente.id ? doc.cliente.id : doc.persona_id;
      const cliNom = (doc.cliente && (doc.cliente.razon_social || doc.cliente.nombre_comercial)) || '—';
      const cliRuc = (doc.cliente && (doc.cliente.ruc || doc.cliente.cedula)) || '';
      // Buscar provincia con prioridad: override manual por RUC/Cédula (Excel) > catálogo
      // sincronizado de Contifico (por persona_id) > inferencia por dirección.
      const cliProv = resolverProvinciaCliente(cliRuc, cliId, doc.cliente?.direccion || '');
      const mes = parseInt((doc.fecha_emision || '').split('/')[1]) || 0;
      // Notas de crédito restan (neteo contra la factura original, igual que Contifico)
      const signoDoc = esNotaCredito(doc) ? -1 : 1;
      const totalDoc = signoDoc * parseFloat(doc.total || 0);
      const subDoc = signoDoc * parseFloat(doc.subtotal || doc.subtotal_12 || 0);
      if (!cliId || totalDoc === 0) return;
      if (!vendedores[vendId]) vendedores[vendId] = { nombre: vendNom, clientes: {} };
      vendedores[vendId].nombre = vendNom;
      const vObj = vendedores[vendId].clientes;
      if (!vObj[cliId]) vObj[cliId] = { id: cliId, nombre: cliNom, ruc: cliRuc, total: 0, subtotal: 0, saldo: 0, num_compras: 0, provincia: cliProv, marcas: {}, marcasPorAnio: {}, marcasPorMes: {}, productos: {}, frecuencia: {}, telefono: '', direccion: '' };
      const cli = vObj[cliId];
      cli.nombre = cliNom; cli.ruc = cliRuc;
      // Teléfono y dirección: se actualiza cada vez que aparece en una factura nueva
      const telCli = doc.cliente?.telefonos || doc.cliente?.telefono;
      if(telCli) cli.telefono = String(telCli);
      if(doc.cliente?.direccion) cli.direccion = doc.cliente.direccion;
      // La provincia se recalcula siempre con el valor más reciente de cliProv, en vez de
      // quedarse fija con el primer valor calculado. Esto es necesario porque el override
      // manual de provincias (Excel subido por Fernando) puede actualizarse en cualquier
      // momento, y de lo contrario un cliente que ya tenía un valor (aunque fuera "Sin
      // provincia" o uno inferido incorrectamente) nunca reflejaría la corrección.
      if(cliProv) cli.provincia = cliProv;
      cli.total += totalDoc; cli.subtotal += subDoc; if (signoDoc > 0) { cli.num_compras++; cli.saldo += parseFloat(doc.saldo || 0); }
      const anioDoc = parseInt((doc.fecha_emision || '').split('/')[2]) || new Date().getFullYear();
      const freqKey = `${anioDoc}-${String(mes).padStart(2,'0')}`;
      if (!cli.frecuencia[freqKey]) cli.frecuencia[freqKey] = { anio: anioDoc, mes, total: 0, subtotal: 0, compras: 0, saldo: 0 };
      cli.frecuencia[freqKey].total += totalDoc; cli.frecuencia[freqKey].subtotal += subDoc;
      if (signoDoc > 0) { cli.frecuencia[freqKey].compras++; cli.frecuencia[freqKey].saldo += parseFloat(doc.saldo || 0); }
      // Desglose exacto por día — para el gráfico "ventas del mes por día" (instantáneo, sin pegarle a Contifico en vivo)
      const diaDoc = parseInt((doc.fecha_emision || '').split('/')[0]) || 0;
      if (diaDoc) {
        const diaKey = `${anioDoc}-${String(mes).padStart(2,'0')}-${String(diaDoc).padStart(2,'0')}`;
        if (!cli.frecuenciaPorDia) cli.frecuenciaPorDia = {};
        if (!cli.frecuenciaPorDia[diaKey]) cli.frecuenciaPorDia[diaKey] = { anio: anioDoc, mes, dia: diaDoc, total: 0, subtotal: 0, compras: 0 };
        cli.frecuenciaPorDia[diaKey].total += totalDoc;
        cli.frecuenciaPorDia[diaKey].subtotal += subDoc;
        if (signoDoc > 0) cli.frecuenciaPorDia[diaKey].compras++;
      }
      (doc.detalles || []).forEach(det => {
        const prodId = det.producto_id || '';
        const cantidad = signoDoc * parseFloat(det.cantidad || 0);
        const base = signoDoc * parseFloat(det.base_gravable || det.base_cero || 0);
        // Solo se descarta si no hay producto identificable o si la cantidad es cero.
        // base===0 es válido (regalos, cortesías, descuento 100%): se cuenta la unidad
        // vendida/entregada, simplemente no aporta nada al total en $.
        if (!prodId || cantidad === 0) return;
        const cat = catalogoProductos[prodId] || {};
        const marca = cat.marca || '';
        const nom = cat.nombre || det.producto_nombre || '';
        if (!cli.productos[prodId]) cli.productos[prodId] = { id: prodId, nombre: nom, codigo: cat.codigo || '', marca, cantidad: 0, total: 0 };
        cli.productos[prodId].nombre = nom;
        cli.productos[prodId].cantidad += cantidad;
        cli.productos[prodId].total += base;
        // Desglose exacto por año y mes para gráficos mensuales por producto
        const pmKey = `${anioDoc}-${mes}-${prodId}`;
        if (!cli.productosPorMes) cli.productosPorMes = {};
        if (!cli.productosPorMes[pmKey]) cli.productosPorMes[pmKey] = { anio: anioDoc, mes, id: prodId, nombre: nom, marca, cantidad: 0, total: 0 };
        cli.productosPorMes[pmKey].cantidad += cantidad;
        cli.productosPorMes[pmKey].total += base;
        if (marca) {
          cli.marcas[marca] = (cli.marcas[marca] || 0) + base;
          // Desglose exacto por año y mes (sin necesidad de ratio en el frontend)
          const mkKey = `${anioDoc}-${marca}`;
          if (!cli.marcasPorAnio) cli.marcasPorAnio = {};
          if (!cli.marcasPorAnio[mkKey]) cli.marcasPorAnio[mkKey] = { anio: anioDoc, marca, total: 0 };
          cli.marcasPorAnio[mkKey].total += base;
          const mkMesKey = `${anioDoc}-${mes}-${marca}`;
          if (!cli.marcasPorMes) cli.marcasPorMes = {};
          if (!cli.marcasPorMes[mkMesKey]) cli.marcasPorMes[mkMesKey] = { anio: anioDoc, mes, marca, total: 0 };
          cli.marcasPorMes[mkMesKey].total += base;
        }
      });
    });
    nextUrl = data.next || null;
    paginas++;
  }
  const resultado = {};
  Object.values(vendedores).forEach(vend => {
    const listaCli = Object.values(vend.clientes).map(cli => ({
      id: cli.id, nombre: cli.nombre, ruc: cli.ruc,
      telefono: cli.telefono || '',
      direccion: cli.direccion || '',
      total: Math.round(cli.total * 100) / 100,
      subtotal: Math.round(cli.subtotal * 100) / 100,
      saldo: Math.round((cli.saldo || 0) * 100) / 100,
      num_compras: cli.num_compras, provincia: cli.provincia,
      marcas: Object.entries(cli.marcas).map(([m,t]) => ({ marca: m, total: Math.round(t*100)/100 })).sort((a,b) => b.total-a.total),
      marcas_anio: Object.values(cli.marcasPorAnio||{}).map(x => ({ anio: x.anio, marca: x.marca, total: Math.round(x.total*100)/100 })),
      marcas_mes: Object.values(cli.marcasPorMes||{}).map(x => ({ anio: x.anio, mes: x.mes, marca: x.marca, total: Math.round(x.total*100)/100 })),
      productos: Object.values(cli.productos).map(p => ({ id: p.id, nombre: p.nombre, codigo: p.codigo, marca: p.marca, cantidad: Math.round(p.cantidad), total: Math.round(p.total*100)/100 })).sort((a,b) => b.cantidad-a.cantidad),
      productos_mes: Object.values(cli.productosPorMes||{}).map(x => ({ anio: x.anio, mes: x.mes, id: x.id, nombre: x.nombre, marca: x.marca, cantidad: Math.round(x.cantidad*100)/100, total: Math.round(x.total*100)/100 })),
      frecuencia: Object.values(cli.frecuencia).map(f => ({ anio: f.anio, mes: f.mes, total: Math.round(f.total*100)/100, subtotal: Math.round(f.subtotal*100)/100, compras: f.compras, saldo: Math.round((f.saldo||0)*100)/100 })).sort((a,b) => a.anio!==b.anio ? a.anio-b.anio : a.mes-b.mes),
      frecuencia_dia: Object.values(cli.frecuenciaPorDia||{}).map(f => ({ anio: f.anio, mes: f.mes, dia: f.dia, total: Math.round(f.total*100)/100, subtotal: Math.round(f.subtotal*100)/100, compras: f.compras }))
    })).sort((a,b) => b.total-a.total);
    // Acumular (no sobreescribir): una vendedora puede aparecer con su id real y con
    // el id sintético 'reasig::' — ambos grupos se fusionan bajo el mismo nombre
    if (!resultado[vend.nombre]) resultado[vend.nombre] = [];
    resultado[vend.nombre].push(...listaCli);
  });
  Object.keys(resultado).forEach(k => resultado[k].sort((a,b) => b.total - a.total));
  console.log(`Generación completa. Duplicados omitidos: ${duplicadosOmitidos}`);
  return resultado;
}

function nowEC() {
  // Retorna un Date ajustado a la hora actual de Ecuador (UTC-5)
  return new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Guayaquil' }));
}

function fmtDateEC(d) {
  // Siempre usar hora de Ecuador (UTC-5) independientemente del timezone del servidor
  const ecDate = new Date(d.toLocaleString('en-US', { timeZone: 'America/Guayaquil' }));
  const dd = String(ecDate.getDate()).padStart(2,'0');
  const mm = String(ecDate.getMonth()+1).padStart(2,'0');
  const yyyy = ecDate.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// ─── INVENTARIO: parseo del Excel "Reporte de Saldos de Inventario por Bodega" ──
// Lee SKU (columna 'Código'), suma 'Bodega POS' + 'BODEGA CASA', e ignora el resto
// de bodegas/personas. Hace match contra catalogoProductos por código (SKU corto),
// que ya viene poblado desde la API de productos de Contifico (p.codigo).
function parsearExcelInventario(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Fecha de corte: buscar en las primeras filas una celda que diga "Fecha de Corte: YYYY-MM-DD"
  let fechaCorte = null;
  for (let i = 0; i < Math.min(6, filas.length); i++) {
    const celda = (filas[i] || []).find(c => typeof c === 'string' && c.includes('Fecha de Corte'));
    if (celda) {
      const m = /(\d{4}-\d{2}-\d{2})/.exec(celda);
      if (m) fechaCorte = m[1];
    }
  }
  if (!fechaCorte) fechaCorte = fmtDateEC(nowEC()).split('/').reverse().join('-'); // fallback: hoy

  // Encontrar la fila de encabezados: la que contiene 'Código' y 'Producto'
  let filaEncabezado = -1;
  for (let i = 0; i < Math.min(15, filas.length); i++) {
    const fila = filas[i] || [];
    if (fila.includes('Código') && fila.includes('Producto')) { filaEncabezado = i; break; }
  }
  if (filaEncabezado === -1) throw new Error('No se encontró la fila de encabezados (Código/Producto) en el Excel');

  const encabezados = filas[filaEncabezado];
  const idxCodigo = encabezados.indexOf('Código');
  const idxProducto = encabezados.indexOf('Producto');
  const idxMarca = encabezados.indexOf('Marca');
  const idxPOS = encabezados.indexOf('Bodega POS');
  const idxCasa = encabezados.findIndex(h => (h||'').toString().toUpperCase().trim() === 'BODEGA CASA');
  if (idxCodigo === -1) throw new Error('No se encontró la columna Código');
  if (idxPOS === -1 && idxCasa === -1) throw new Error('No se encontraron las columnas Bodega POS / BODEGA CASA');

  const filasProducto = [];
  for (let i = filaEncabezado + 1; i < filas.length; i++) {
    const fila = filas[i];
    if (!fila || fila[idxCodigo] === null || fila[idxCodigo] === undefined || fila[idxCodigo] === '') continue; // fila de totales u otra vacía
    const sku = String(fila[idxCodigo]).trim();
    const nombre = idxProducto !== -1 ? (fila[idxProducto]||'').toString().trim() : '';
    const marca = idxMarca !== -1 ? (fila[idxMarca]||'').toString().trim().toUpperCase() : '';
    const cantPOS = idxPOS !== -1 ? (parseFloat(fila[idxPOS]) || 0) : 0;
    const cantCasa = idxCasa !== -1 ? (parseFloat(fila[idxCasa]) || 0) : 0;
    filasProducto.push({ sku, nombre, marca, cantidad: cantPOS + cantCasa });
  }
  return { fechaCorte, filasProducto };
}

// Resuelve cada fila del Excel (por SKU) contra catalogoProductos (por p.codigo),
// devolviendo { productos: { [producto_id]: {cantidad, sku, nombre, marca} }, sinMatch: [...] }
function resolverInventarioContraCatalogo(filasProducto) {
  const skuAProductoId = {};
  Object.entries(catalogoProductos).forEach(([id, info]) => {
    const cod = (info.codigo||'').trim();
    if (cod) skuAProductoId[cod] = id;
  });
  const productos = {};
  const sinMatch = [];
  filasProducto.forEach(f => {
    const prodId = skuAProductoId[f.sku];
    if (prodId) {
      productos[prodId] = { cantidad: f.cantidad, sku: f.sku, nombre: f.nombre, marca: f.marca };
    } else {
      sinMatch.push(f);
    }
  });
  return { productos, sinMatch };
}

// Rotación mensual promedio.
//
// Por defecto usa los últimos meses CERRADOS, pero puede incluir el mes en curso
// proyectado a mes completo (regla de tres sobre los días transcurridos). Esto último
// hace falta para decidir promociones a mitad de mes: sin ello, a fines de agosto la
// rotación seguía mirando mayo, junio y julio, ignorando por completo lo que estaba
// pasando ahora — justo lo que cambia cuando se ajustan precios o entra una línea nueva.
//
// La proyección solo se aplica después del día 8: antes, dos o tres días de ventas
// multiplicados por diez producen cifras disparatadas.
function calcularRotacionMensual(fechaCorte, opciones) {
  opciones = opciones || {};
  const nMeses = [2,3,4,6].includes(parseInt(opciones.meses)) ? parseInt(opciones.meses) : 3;
  const incluirActual = !!opciones.incluirMesActual;
  const [anioCorte, mesCorte] = fechaCorte.split('-').map(Number);

  const hoyEcR = nowEC();
  const esMesEnCursoElCorte = (hoyEcR.getFullYear() === anioCorte && (hoyEcR.getMonth()+1) === mesCorte);
  const diaHoy = hoyEcR.getDate();
  const diasDelMes = new Date(anioCorte, mesCorte, 0).getDate();
  // Factor de la regla de tres: cuánto representaría el mes completo al ritmo actual
  const factorProyeccion = (diaHoy >= 8 ? diasDelMes / diaHoy : null);
  const usarActual = incluirActual && esMesEnCursoElCorte && factorProyeccion !== null;

  // Meses cerrados anteriores. Si se incluye el mes en curso, ocupa uno de los espacios.
  const cerrados = [];
  let a = anioCorte, m = mesCorte;
  for (let i = 0; i < (usarActual ? nMeses - 1 : nMeses); i++) {
    m -= 1;
    if (m === 0) { m = 12; a -= 1; }
    cerrados.push({ anio: a, mes: m });
  }

  const acumulado = {};
  Object.values(DATA_CACHE||{}).forEach(clientes => {
    (clientes||[]).forEach(cli => {
      (cli.productos_mes||[]).forEach(pm => {
        const key = pm.id || pm.nombre;
        if (cerrados.some(x => x.anio===pm.anio && x.mes===pm.mes)) {
          acumulado[key] = (acumulado[key]||0) + (pm.cantidad||0);
        } else if (usarActual && pm.anio===anioCorte && pm.mes===mesCorte) {
          acumulado[key] = (acumulado[key]||0) + (pm.cantidad||0) * factorProyeccion;
        }
      });
    });
  });

  const divisor = usarActual ? nMeses : cerrados.length;
  const rotacion = {};
  Object.entries(acumulado).forEach(([id, total]) => { rotacion[id] = total / divisor; });
  console.log(`Rotación: ${cerrados.map(x=>x.mes+'/'+x.anio).join(', ')}`
    + (usarActual ? ` + ${mesCorte}/${anioCorte} proyectado (día ${diaHoy} de ${diasDelMes}, ×${factorProyeccion.toFixed(2)})` : '')
    + ` · promedio de ${divisor}`);
  rotacion.__meta = { meses: divisor, incluye_mes_actual: usarActual,
    factor: usarActual ? Math.round(factorProyeccion*100)/100 : null, dia: diaHoy, dias_mes: diasDelMes };
  return rotacion;
}

// Mínimo de seguridad y umbral de alerta amarilla, por marca (en meses de cobertura)
const INVENTARIO_REGLAS_MARCA = {
  'BIOSKIN':   { minimo: 1, amarillo: 1.5 },
  'ZIAJA':     { minimo: 3, amarillo: 4 },
  'ZIAJA PRO': { minimo: 3, amarillo: 4 },
  'ERAYBA':    { minimo: 3, amarillo: 4 }
};

function calcularSemaforo(marca, coberturaMeses) {
  const reglas = INVENTARIO_REGLAS_MARCA[marca] || { minimo: 3, amarillo: 4 };
  if (coberturaMeses < reglas.minimo) return 'rojo';
  if (coberturaMeses < reglas.amarillo) return 'amarillo';
  return 'verde';
}

// Construye la lista completa de inventario por marca: todos los productos del catálogo
// de esa marca, con su inventario actual (0 si no está en el Excel cargado), rotación
// mensual, cobertura en meses, y semáforo.
function construirInventarioPorMarca(marcaFiltro, opciones) {
  if (!INVENTARIO_CACHE) return { fecha_corte: null, productos: [] };
  opciones = opciones || {};
  // La rotación se mide contra HOY cuando se pide incluir el mes en curso: el Excel de
  // inventario puede tener una fecha de corte anterior y no debe fijar el calendario.
  const refFecha = opciones.incluirMesActual
    ? new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'})
    : INVENTARIO_CACHE.fecha_corte;
  const rotacion = calcularRotacionMensual(refFecha, opciones);
  const meta = rotacion.__meta || {};

  // Meses que se muestran como columnas de ventas: los cerrados del cálculo, y si se
  // incluye el mes en curso, ese va al final marcado como proyectado.
  const [anioCorte, mesCorte] = refFecha.split('-').map(Number);
  const nCols = meta.meses || 3;
  const meses3 = [];
  let a = anioCorte, m = mesCorte;
  const cerradosAMostrar = meta.incluye_mes_actual ? nCols - 1 : nCols;
  for(let i = 0; i < cerradosAMostrar; i++){
    m -= 1; if(m === 0){ m = 12; a -= 1; }
    meses3.unshift({ anio: a, mes: m }); // orden cronológico
  }
  if (meta.incluye_mes_actual) meses3.push({ anio: anioCorte, mes: mesCorte, enCurso: true });

  // Acumular ventas por producto por mes
  const ventasMes = {}; // { prodId: { 'anio-mes': cantidad } }
  Object.values(DATA_CACHE||{}).forEach(clientes => {
    (clientes||[]).forEach(cli => {
      (cli.productos_mes||[]).forEach(pm => {
        if(!meses3.some(x => x.anio===pm.anio && x.mes===pm.mes)) return;
        const key = pm.id || pm.nombre;
        if(!ventasMes[key]) ventasMes[key] = {};
        const mk = `${pm.anio}-${pm.mes}`;
        ventasMes[key][mk] = (ventasMes[key][mk]||0) + (pm.cantidad||0);
      });
    });
  });

  const MESES_LABEL = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

  const productosDelCatalogo = Object.entries(catalogoProductos)
    .filter(([id, info]) => (info.marca||'').toUpperCase() === marcaFiltro)
    .filter(([id, info]) => !(info.nombre||'').trim().toUpperCase().startsWith('PROMO'))
    .filter(([id, info]) => !(info.nombre||'').trim().toUpperCase().startsWith('LÍNEA') && !(info.nombre||'').trim().toUpperCase().startsWith('LINEA'))
    .filter(([id, info]) => !/\bKITS?\b/.test((info.nombre||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase()));

  const lista = productosDelCatalogo.map(([id, info]) => {
    const inv = INVENTARIO_CACHE.productos[id];
    const stock = inv ? inv.cantidad : 0;
    const rotacionMensual = rotacion[id] || 0;
    const cobertura = rotacionMensual > 0 ? stock / rotacionMensual : (stock > 0 ? 99 : 0);
    const necesidad12Meses = (rotacionMensual * 12) - stock;

    // Ventas de cada uno de los 3 meses cerrados
    const vm = ventasMes[id] || {};
    const ventas3 = meses3.map(x => ({
      label: MESES_LABEL[x.mes] + (x.enCurso ? '*' : ''),
      en_curso: !!x.enCurso,
      cantidad: Math.round(vm[`${x.anio}-${x.mes}`] || 0)
    }));

    return {
      id,
      sku: info.codigo || (inv ? inv.sku : ''),
      nombre: info.nombre,
      marca: marcaFiltro,
      stock: Math.round(stock*100)/100,
      rotacion_mensual: Math.round(rotacionMensual*100)/100,
      cobertura_meses: Math.round(cobertura*10)/10,
      necesidad_12_meses: Math.round(necesidad12Meses),
      semaforo: calcularSemaforo(marcaFiltro, cobertura),
      ventas3
    };
  }).sort((a,b) => a.cobertura_meses - b.cobertura_meses);

  return { fecha_corte: INVENTARIO_CACHE.fecha_corte, productos: lista,
    meses3labels: meses3.map(x=>MESES_LABEL[x.mes] + (x.enCurso ? '*' : '')),
    rotacion_meta: meta };
}

async function sincronizarHoy() {
  if (cache.sincronizando) return;
  cache.sincronizando = true;
  try {
    const now = nowEC();
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
    const clientes = todos.filter(d => d.tipo_registro === 'CLI' && !d.anulado && !esNotaCredito(d) && !noEsVenta(d));
    // Notas de crédito del día — se restan en los gráficos (mismo neteo que la regeneración)
    cache.nc_documentos = todos.filter(d => d.tipo_registro === 'CLI' && !d.anulado && esNotaCredito(d));
    // Agregar cliente_nombre directo desde el objeto cliente
    clientes.forEach(d => {
      d.cliente_nombre = d.cliente?.razon_social || d.cliente?.nombre_comercial || d.persona_id || '—';
    });
    cache.documentos = clientes;
    cache.ultima_sync = new Date().toISOString();
    console.log(`✓ Sync: ${clientes.length} facturas de clientes hoy`);

    // Guardar el detalle de cada factura en la BD para tener historial real consultable
    // por fecha (la tabla en memoria `cache.documentos` se sobreescribe cada hora, así
    // que sin esto perderíamos el detalle de días anteriores al pasar la medianoche).
    try {
      for (const d of clientes) {
        const vendNom = d.vendedor?.razon_social || d.vendedor?.nombre || 'Sin asignar';
        await pool.query(
          `INSERT INTO facturas_detalle(documento_id, fecha, documento, cliente_nombre, vendedor_nombre, subtotal, total, cedula_ruc)
           VALUES($1,$2,$3,$4,$5,$6,$7,$8)
           ON CONFLICT (documento_id, fecha) DO UPDATE SET
             documento=$3, cliente_nombre=$4, vendedor_nombre=$5, subtotal=$6, total=$7, cedula_ruc=$8, actualizado_at=NOW()`,
          [
            String(d.id || d.documento),
            fechaParaSQL(fecha), // fecha en formato YYYY-MM-DD
            d.documento || '',
            d.cliente_nombre || '—',
            vendNom,
            parseFloat(d.subtotal || (d.total/1.15) || 0),
            parseFloat(d.total || 0),
            String(d.cliente?.cedula || d.cliente?.ruc || d.cliente?.identificacion || '').replace(/\D/g,'') || null
          ]
        );
      }
    } catch(eDb) {
      console.error('Error guardando facturas_detalle:', eDb.message);
    }
  } catch(e) {
    console.error('Error sync:', e.message);
  }
  cache.sincronizando = false;
}

// Correo que recibe la clienta con su número de guía.
function plantillaCorreoGuia(nombre, guia, ciudad){
  const esc = v => String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  return `<!DOCTYPE html><html lang="es"><body style="margin:0;padding:0;background:#f6f3ef;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f3ef;padding:28px 12px">
    <tr><td align="center">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,0.06)">
        <tr><td style="background:#1a1108;padding:22px 26px;text-align:center">
          <div style="color:#ffffff;font-size:22px;font-weight:800;letter-spacing:6px">COSÉTIKA</div>
        </td></tr>
        <tr><td style="padding:28px 26px 8px">
          <p style="margin:0 0 14px;font-size:16px;color:#2b2118">Hola ${esc(nombre)},</p>
          <p style="margin:0 0 20px;font-size:14px;line-height:1.6;color:#5b5248">
            Tu pedido ya salió de nuestra bodega y está en camino${ciudad ? ' hacia ' + esc(ciudad) : ''}.
            Puedes seguirlo con este número de guía:
          </p>
          <div style="background:#faf7f4;border:1px solid #e7dccd;border-radius:11px;padding:18px;text-align:center;margin-bottom:20px">
            <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;color:#8b8177;margin-bottom:6px">Número de guía · Servientrega</div>
            <div style="font-size:26px;font-weight:800;color:#8a5a3b;letter-spacing:1px">${esc(guia)}</div>
          </div>
          <p style="margin:0 0 22px;font-size:14px;line-height:1.6;color:#5b5248">
            Ingresa ese número en <a href="https://www.servientrega.com.ec/" style="color:#8a5a3b;font-weight:600">servientrega.com.ec</a> para ver dónde está tu envío.
          </p>
          <p style="margin:0 0 6px;font-size:14px;line-height:1.6;color:#5b5248">
            Gracias por confiar en nosotros. Cualquier novedad con tu pedido, respóndenos este correo.
          </p>
        </td></tr>
        <tr><td style="padding:18px 26px 26px;border-top:1px solid #f0eae3;text-align:center">
          <div style="font-size:12px;color:#8b8177">Corporación Cosétika S.A.S.</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}

// ─── CARTERA POR COBRAR (en vivo desde Contifico) ─────────────────────────────────
// Suma el saldo pendiente de las facturas y separa lo vencido de lo que aún está en plazo,
// usando los días de crédito de cada clienta. Se refresca dos veces al día.
let CARTERA = { total:0, vencida:0, por_vencer:0, docs:0, clientes:[], vencimientos:[], at:null, error:null };
let CARTERA_EN_CURSO = false;
async function sincronizarCartera(){
  if (!API_KEY) return;
  if (CARTERA_EN_CURSO) return;      // evita que dos peticiones disparen el mismo barrido
  CARTERA_EN_CURSO = true;
  try {
    const hoyK = nowEC();
    const desdeD = new Date(hoyK); desdeD.setMonth(desdeD.getMonth() - 6);
    const desde = fmtDateEC(desdeD), hasta = fmtDateEC(hoyK);
    let url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
    let pg = 0; const vistos = new Set(); const porCliente = {}; const porDia = {};
    let total = 0, vencida = 0, docs = 0;
    let fallos = 0;
    while (url && pg < 400) {
      // Contifico limita las peticiones seguidas. Antes se hacía `if (!r.ok) break;` y el
      // barrido se rendía en la primera página estrangulada, dejando la cartera con una
      // fracción del total. Ahora se reintenta con espera creciente y solo se abandona
      // después de tres intentos fallidos sobre la MISMA página.
      let d = null;
      for (let intento = 0; intento < 3 && !d; intento++) {
        if (intento) await new Promise(res => setTimeout(res, [0, 2000, 6000][intento]));
        try {
          const r = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
          if (r.ok) d = await r.json();
          else { fallos++; console.log(`Cartera: página ${pg+1} devolvió HTTP ${r.status} (intento ${intento+1})`); }
        } catch(e) { fallos++; console.log(`Cartera: error de red en página ${pg+1}: ${e.message}`); }
      }
      if (!d) { console.error(`✗ Cartera incompleta: la página ${pg+1} falló tres veces`); break; }
      (d.results || []).forEach(doc => {
        if (doc.tipo_registro !== 'CLI' || doc.anulado || noEsVenta(doc) || esNotaCredito(doc)) return;
        const k = doc.id || doc.documento;
        if (vistos.has(k)) return; vistos.add(k);
        const saldo = parseFloat(doc.saldo || 0);
        if (!(saldo > 0.01)) return;
        const ident = String((doc.cliente && (doc.cliente.ruc || doc.cliente.cedula)) || '').replace(/\D/g,'');
        if (ident === '1793143660001') return;   // autoconsumo
        const info = ident ? (CREDITO_CACHE[ident] || (ident.length===13?CREDITO_CACHE[ident.substring(0,10)]:null)) : null;
        // Regla del negocio: toda clienta con crédito paga a 30 días. Si Contifico tiene un
        // plazo distinto para alguien, manda el de Contifico.
        const dias = (info && info.dias > 0) ? info.dias : 30;
        const fe = String(doc.fecha_emision || '').split('/');
        const fEmis = fe.length===3 ? new Date(parseInt(fe[2]), parseInt(fe[1])-1, parseInt(fe[0])) : null;
        let estaVencida = false, diasAtraso = 0, claveVence = null;
        if (fEmis) {
          const vence = new Date(fEmis); vence.setDate(vence.getDate() + dias);
          const hoy0 = new Date(hoyK.getFullYear(), hoyK.getMonth(), hoyK.getDate());
          if (vence < hoy0) { estaVencida = true; diasAtraso = Math.round((hoy0 - vence)/86400000); }
          claveVence = vence.getFullYear()+'-'+String(vence.getMonth()+1).padStart(2,'0')+'-'+String(vence.getDate()).padStart(2,'0');
          if (!estaVencida) {
            if (!porDia[claveVence]) porDia[claveVence] = { fecha: claveVence, monto: 0, docs: 0 };
            porDia[claveVence].monto += saldo; porDia[claveVence].docs++;
          }
        }
        total += saldo; docs++;
        if (estaVencida) vencida += saldo;
        const nom = (doc.cliente && (doc.cliente.razon_social || doc.cliente.nombre_comercial)) || '—';
        if (!porCliente[nom]) porCliente[nom] = { nombre: nom, total:0, vencido:0, docs:0, atraso:0 };
        porCliente[nom].total += saldo; porCliente[nom].docs++;
        if (estaVencida) { porCliente[nom].vencido += saldo; porCliente[nom].atraso = Math.max(porCliente[nom].atraso, diasAtraso); }
      });
      url = d.next || null; pg++;
      if (url) await new Promise(res => setTimeout(res, 200));   // respiro entre páginas
    }
    const r2 = x => Math.round(x*100)/100;
    CARTERA = {
      total: r2(total), vencida: r2(vencida), por_vencer: r2(total - vencida), docs,
      clientes: Object.values(porCliente).sort((a,b)=>b.total-a.total).slice(0,25)
        .map(c=>({ ...c, total:r2(c.total), vencido:r2(c.vencido) })),
      // Calendario de vencimientos: qué se cobra cada día de aquí en adelante
      vencimientos: Object.values(porDia).sort((a,b)=>a.fecha.localeCompare(b.fecha))
        .map(d=>({ ...d, monto: r2(d.monto) })),
      paginas: pg, reintentos: fallos, completo: !url,
      at: new Date().toISOString(), error: null
    };
    // Persistir: tras un redeploy el panel muestra la última lectura conocida en vez de ceros
    try { await setConfigApp('cartera_cache', JSON.stringify(CARTERA)); } catch(e) {}
    console.log(`✓ Cartera: ${docs} facturas · total ${CARTERA.total} · vencida ${CARTERA.vencida} · ${pg} páginas${fallos?' ('+fallos+' reintentos)':''}`);
  } catch(e) { CARTERA.error = e.message; console.error('Error cartera:', e.message); }
  CARTERA_EN_CURSO = false;
}
// Al arrancar se recupera la última lectura guardada; el recálculo completo corre una vez
// al día (de madrugada) o cuando se pide a mano desde el panel.
setTimeout(async () => {
  try {
    const raw = await getConfigApp('cartera_cache', null);
    if (raw) { const c = JSON.parse(raw); if (c && c.at) { CARTERA = c; console.log(`✓ Cartera recuperada del caché: ${c.docs} facturas · ${c.at}`); } }
  } catch(e) {}
  if (!CARTERA.at) sincronizarCartera().catch(e=>console.error(e));
}, 60 * 1000);
setInterval(() => {
  const h = nowEC().getHours();
  if (h === 3) sincronizarCartera().catch(e=>console.error(e));   // una vez al día, 3 AM Ecuador
}, 60 * 60 * 1000);

// ─── FUENTE ÚNICA DE VERDAD ───────────────────────────────────────────────────────
// El data.json que recibe el navegador ya viene COMPLETO: caché histórica + el tramo que
// aún no ha entrado en ella, fusionado aquí con la MISMA función que genera todo el resto
// (generarDataJson). El navegador no vuelve a aplicar ninguna regla de negocio.
//
// Por qué importa: durante meses las reglas (qué es venta, qué se descarta, cómo se netean
// las notas de crédito) vivían escritas dos veces —una en el servidor y otra en el navegador—
// y bastaba que una se quedara atrás para que las cifras se separaran de Contifico. Con una
// sola implementación, esa clase de error deja de ser posible.
const _r2 = x => Math.round((x||0)*100)/100;
function _consolidar(lista, clave, sumar){
  const m = {};
  (lista||[]).forEach(x => {
    if (!x) return;
    const k = clave(x);
    if (!m[k]) m[k] = { ...x };
    else sumar(m[k], x);
  });
  return Object.values(m);
}
// REEMPLAZO DEL MES EN CURSO.
// No intenta adivinar hasta dónde llega el caché ni pegarle encima lo que falta: borra el
// mes completo de lo guardado y lo repone con lo que Contifico dice AHORA. Da igual si el
// caché venía corto, adelantado o con días repetidos — el mes queda exacto por construcción.
//
// Los totales del cliente se DERIVAN de su detalle mensual (frecuencia, marcas_mes,
// productos_mes) en vez de arrastrarse sumando y restando. Así no pueden acumular deriva.
function fusionMesActual(base, extra, anio, mes){
  const delMes = x => x && x.anio === anio && x.mes === mes;
  const out = {};

  // 1) Copiar el caché SIN el mes objetivo
  Object.entries(base || {}).forEach(([v, cls]) => {
    out[v] = (cls || []).map(c => ({
      ...c,
      frecuencia:     (c.frecuencia||[]).filter(x => !delMes(x)),
      frecuencia_dia: (c.frecuencia_dia||[]).filter(x => !delMes(x)),
      marcas_mes:     (c.marcas_mes||[]).filter(x => !delMes(x)),
      productos_mes:  (c.productos_mes||[]).filter(x => !delMes(x)),
      _saldoCache:    c.saldo || 0,
      _codigos:       Object.fromEntries((c.productos||[]).map(p => [p.id || p.nombre, p.codigo || '']))
    }));
  });

  // 2) Reponer el mes con los datos frescos
  Object.entries(extra || {}).forEach(([v, cls]) => {
    if (!out[v]) out[v] = [];
    const porId = {}; out[v].forEach(c => { porId[c.id] = c; });
    (cls || []).forEach(ce => {
      let c = porId[ce.id];
      if (!c) {
        c = { ...ce, frecuencia:[], frecuencia_dia:[], marcas_mes:[], productos_mes:[], _saldoCache:null,
              _codigos: Object.fromEntries((ce.productos||[]).map(p => [p.id || p.nombre, p.codigo || ''])) };
        out[v].push(c); porId[ce.id] = c;
      }
      c.nombre = ce.nombre || c.nombre;
      if (ce.ruc) c.ruc = ce.ruc;
      if (ce.telefono) c.telefono = ce.telefono;
      if (ce.direccion) c.direccion = ce.direccion;
      if (ce.provincia) c.provincia = ce.provincia;
      (ce.productos||[]).forEach(p => { if (p.codigo) c._codigos[p.id || p.nombre] = p.codigo; });
      c.frecuencia     = c.frecuencia.concat((ce.frecuencia||[]).filter(delMes));
      c.frecuencia_dia = c.frecuencia_dia.concat((ce.frecuencia_dia||[]).filter(delMes));
      c.marcas_mes     = c.marcas_mes.concat((ce.marcas_mes||[]).filter(delMes));
      c.productos_mes  = c.productos_mes.concat((ce.productos_mes||[]).filter(delMes));
    });
  });

  // 3) Derivar todos los agregados desde el detalle mensual
  Object.values(out).forEach(cls => cls.forEach(c => {
    c.frecuencia.sort((a,b) => a.anio!==b.anio ? a.anio-b.anio : a.mes-b.mes);
    c.total       = _r2(c.frecuencia.reduce((a,f) => a + (f.total||0), 0));
    c.subtotal    = _r2(c.frecuencia.reduce((a,f) => a + (f.subtotal||0), 0));
    c.num_compras = c.frecuencia.reduce((a,f) => a + (f.compras||0), 0);
    // Saldo: derivado si el caché ya trae saldo por mes; si es un caché viejo que no lo
    // tiene, se conserva el saldo guardado (se vuelve exacto tras la próxima regeneración).
    const traeSaldo = c.frecuencia.some(f => f.saldo !== undefined);
    c.saldo = traeSaldo ? _r2(c.frecuencia.reduce((a,f) => a + (f.saldo||0), 0))
                        : _r2(c._saldoCache || 0);
    c.marcas_anio = _consolidar(c.marcas_mes.map(x => ({ anio:x.anio, marca:x.marca, total:x.total })),
      x => x.anio+'|'+x.marca, (a,b) => { a.total = _r2(a.total + b.total); });
    c.marcas = _consolidar(c.marcas_mes.map(x => ({ marca:x.marca, total:x.total })),
      x => x.marca, (a,b) => { a.total = _r2(a.total + b.total); }).sort((a,b) => b.total - a.total);
    c.productos = _consolidar(c.productos_mes.map(x => ({ id:x.id, nombre:x.nombre, codigo:(c._codigos||{})[x.id||x.nombre]||'', marca:x.marca, cantidad:x.cantidad, total:x.total })),
      x => x.id || x.nombre, (a,b) => { a.cantidad = _r2(a.cantidad + b.cantidad); a.total = _r2(a.total + b.total); })
      .sort((a,b) => b.cantidad - a.cantidad);
    delete c._saldoCache; delete c._codigos;
  }));

  Object.keys(out).forEach(v => out[v].sort((a,b) => b.total - a.total));
  return out;
}

let DATA_SERVIDA = { clave:'', ts:0, data:null };
// Recalcular el mes en curso en segundo plano, un poco antes de que expire el memo (3 min).
// Así la petición de un usuario siempre encuentra el resultado listo en vez de esperar a
// que se recorra Contifico — era el principal motivo de que la app tardara en abrir.
setInterval(() => { dataCompleta().catch(e=>console.error('Precalentado data:', e.message)); }, 2 * 60 * 1000);
setTimeout(() => { dataCompleta().catch(e=>console.error('Precalentado data:', e.message)); }, 25 * 1000);
async function dataCompleta(){
  const cacheBase = DATA_CACHE || {};
  const hoy = nowEC();
  const anio = hoy.getFullYear();
  const mes = hoy.getMonth() + 1;
  const hastaStr = fmtDateEC(hoy);
  const clave = anio + '-' + mes + '|' + hastaStr + '|' + (DATA_CACHE_TS || '');
  if (DATA_SERVIDA.data && DATA_SERVIDA.clave === clave && (Date.now() - DATA_SERVIDA.ts) < 3 * 60 * 1000) {
    return DATA_SERVIDA.data;
  }
  try {
    // El mes en curso se recalcula ENTERO con las mismas reglas de siempre y reemplaza al
    // del caché. No se infiere nada sobre hasta dónde llegaba lo guardado.
    const desdeStr = '01/' + String(mes).padStart(2,'0') + '/' + anio;
    const mesFresco = await generarDataJson(desdeStr, hastaStr);
    const fusion = fusionMesActual(cacheBase, mesFresco, anio, mes);
    DATA_SERVIDA = { clave, ts: Date.now(), data: fusion };
    console.log(`✓ data.json: mes ${desdeStr}→${hastaStr} reemplazado con datos frescos`);
    return fusion;
  } catch(e) {
    console.error('Error recalculando el mes en curso:', e.message);
    return cacheBase;   // degradar antes que romper
  }
}

// ─── VENTAS PENDIENTES ────────────────────────────────────────────────────────────
// El data.json guardado llega hasta cierta fecha de corte (app_config.data_hasta). Todo
// lo emitido después vive solo en Contifico. Esta función trae ese tramo completo — no
// únicamente "hoy" — para que no exista forma de perder un día.
let PEND_CACHE = { ts: 0, clave: '', data: null };
// Último día con ventas presente en el caché guardado (usa el desglose diario del data.json)
function ultimoDiaEnCache(){
  let max = null;
  try {
    Object.values(DATA_CACHE || {}).forEach(clientes => (clientes || []).forEach(c => {
      (c.frecuencia_dia || []).forEach(f => {
        if (!f || !f.anio || !f.mes || !f.dia) return;
        const d = new Date(f.anio, f.mes - 1, f.dia);
        if (!max || d > max) max = d;
      });
    }));
  } catch(e) {}
  return max;
}
function _parseDDMMYYYY(str){
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(str||'').trim());
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
}
async function ventasPendientes(forzar){
  const hoyD = nowEC();
  const hasta = fmtDateEC(hoyD);
  let desde = hasta;
  // El corte REAL se deduce del propio caché: el último día con ventas registradas.
  // Es imposible que se solapen (nunca pedimos un día que el caché ya tenga) y es
  // inmune a que la regeneración nocturna falle, se atrase o se salte días.
  let corte = ultimoDiaEnCache();
  const corteStr = corte ? fmtDateEC(corte) : (await getConfigApp('data_hasta', '') || null);
  if (!corte) corte = _parseDDMMYYYY(corteStr);
  if (corte) {
    const sig = new Date(corte); sig.setDate(sig.getDate() + 1);
    // Tope de seguridad: nunca pedir más de 40 días hacia atrás
    const tope = new Date(hoyD); tope.setDate(tope.getDate() - 40);
    desde = fmtDateEC(sig < tope ? tope : sig);
  }
  const dDesde = _parseDDMMYYYY(desde);
  const dHasta = _parseDDMMYYYY(hasta);
  if (dDesde && dHasta && dDesde > dHasta) desde = hasta;  // el caché ya cubre hoy

  // El memo se invalida solo si cambia el rango. Sin esto, cuando la regeneración
  // avanzaba el corte del caché, el memo viejo seguía devolviendo días que el caché
  // ya había absorbido — y esos días se contaban dos veces.
  const claveRango = desde + '|' + hasta;
  if (!forzar && PEND_CACHE.data && PEND_CACHE.clave === claveRango && (Date.now() - PEND_CACHE.ts) < 5 * 60 * 1000) {
    return PEND_CACHE.data;
  }

  let todos = [];
  let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
  let pgs = 0;
  while (nextUrl && pgs < 60) {
    const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
    if (!resp.ok) break;
    const data = await resp.json();
    todos = todos.concat(data.results || []);
    nextUrl = data.next || null;
    pgs++;
  }
  const vistos = new Set();
  const base = todos.filter(d => {
    if (d.tipo_registro !== 'CLI' || d.anulado || noEsVenta(d)) return false;
    // Cinturón de seguridad: jamás devolver un día que el caché ya contiene
    if (corte) {
      const fe = String(d.fecha_emision || '').split('/');
      const fd = new Date(parseInt(fe[2]), parseInt(fe[1]) - 1, parseInt(fe[0]));
      if (!isNaN(fd) && fd <= corte) return false;
    }
    const k = d.id || d.documento;
    if (vistos.has(k)) return false;
    vistos.add(k);
    return true;
  });
  const documentos = base.filter(d => !esNotaCredito(d));
  const nc_documentos = base.filter(d => esNotaCredito(d));
  documentos.forEach(d => { d.cliente_nombre = d.cliente?.razon_social || d.cliente?.nombre_comercial || d.persona_id || '—'; });
  nc_documentos.forEach(d => { d.cliente_nombre = d.cliente?.razon_social || d.cliente?.nombre_comercial || d.persona_id || '—'; });
  const out = { total: documentos.length, desde, hasta, corte_cache: corteStr || null, documentos, nc_documentos, generado: new Date().toISOString() };
  PEND_CACHE = { ts: Date.now(), clave: claveRango, data: out };
  console.log(`✓ Ventas pendientes ${desde} → ${hasta}: ${documentos.length} facturas, ${nc_documentos.length} NC`);
  return out;
}
// Refrescar el tramo pendiente cada 10 minutos para que la app abra siempre al instante
setInterval(() => { ventasPendientes(true).catch(e => console.error('Pendientes:', e.message)); }, 10 * 60 * 1000);
setTimeout(() => { ventasPendientes(true).catch(e => console.error('Pendientes:', e.message)); }, 45 * 1000);

// Convierte fecha DD/MM/YYYY (formato Contifico) a YYYY-MM-DD (formato SQL)
function fechaParaSQL(fechaDDMMYYYY){
  const [d,m,y] = fechaDDMMYYYY.split('/');
  return `${y}-${m}-${d}`;
}

// ─── PEDIDOS WEB: lectura de correos "Nuevo pedido" de WooCommerce vía IMAP ──────
// Convierte el HTML del correo a texto plano preservando saltos de línea entre bloques,
// para que los regex de extracción no peguen palabras de celdas/párrafos distintos.
function stripHtmlParaPedido(html){
  return (html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<\/(p|div|tr|h[1-6]|li)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&#8217;|&rsquo;/g, "'")
    .replace(/&#36;|&dollar;/gi, '$')
    .replace(/&#44;/g, ',')
    .replace(/&#46;/g, '.')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n')
    .trim();
}

// Extrae los datos del pedido desde el HTML del correo de WooCommerce "Nuevo pedido".
// Plantilla estable de WooCommerce — ver ejemplo real en conversación con Fernando (pedido #16548).
function parsearPedidoWooCommerce(html, asuntoCorreo, fechaCorreo){
  const texto = stripHtmlParaPedido(html);

  // Número de pedido: preferimos el asunto "...#16548" (más confiable), si no del cuerpo
  let numeroPedido = null;
  const mAsunto = (asuntoCorreo || '').match(/#(\d+)/);
  if (mAsunto) numeroPedido = mAsunto[1];
  if (!numeroPedido) {
    const mCuerpo = texto.match(/n\.?º\s*(\d+)|#(\d+)/);
    if (mCuerpo) numeroPedido = mCuerpo[1] || mCuerpo[2];
  }
  if (!numeroPedido) return null; // sin número de pedido no podemos identificar el registro

  let cliente = null;
  const mCliente = texto.match(/Has recibido un nuevo pedido de\s+(.+?):/i);
  if (mCliente) cliente = mCliente[1].trim();

  let cedulaRuc = null;
  const mCedula = texto.match(/C[ée]dula o RUC:?\s*([0-9]{10,13})/i);
  if (mCedula) cedulaRuc = mCedula[1];

  let total = null;
  // Más tolerante: permite texto/entidades cortas entre "Total:" y el monto (ej. símbolo
  // de moneda en HTML separado, o markup residual que stripTags no limpió del todo)
  const mTotal = texto.match(/\bTotal:?\s*[^\d\n]{0,15}?\$?\s*([\d,]+\.\d{2})/i);
  if (mTotal) total = parseFloat(mTotal[1].replace(/,/g, ''));

  let subtotal = null;
  const mSubtotal = texto.match(/Subtotal:?\s*[^\d\n]{0,15}?\$?\s*([\d,]+\.\d{2})/i);
  if (mSubtotal) subtotal = parseFloat(mSubtotal[1].replace(/,/g, ''));

  const productos = [];
  const regexProducto = /(.+?)\s*\(#(\w+)\)\s*\n?×(\d+)\s*\n?\$?\s*([\d,]+\.\d{2})/g;
  let m;
  while ((m = regexProducto.exec(texto)) !== null) {
    productos.push({
      nombre: m[1].trim(),
      sku: m[2],
      cantidad: parseInt(m[3]),
      precio: parseFloat(m[4].replace(',', ''))
    });
  }

  let telefono = null;
  const mTel = texto.match(/\b(09\d{8})\b/);
  if (mTel) telefono = mTel[1];

  // Fecha del pedido: usamos la fecha del correo (más confiable que parsear "junio 30, 2026")
  const fecha = fechaCorreo
    ? `${fechaCorreo.getFullYear()}-${String(fechaCorreo.getMonth()+1).padStart(2,'0')}-${String(fechaCorreo.getDate()).padStart(2,'0')}`
    : null;

  return { numeroPedido, cliente, cedulaRuc, telefono, subtotal, total, productos, fecha };
}

// Conecta a la casilla pedidos@cosetika.com vía IMAP, revisa correos no leídos de
// "Nuevo pedido", los parsea y guarda en pedidos_web. Marca los correos como leídos
// para no reprocesarlos en la siguiente corrida.
async function enviarPushATodos(payload) {
  if (!webpush) { console.log('⚠️ web-push no disponible'); return; }
  if (!VAPID_PUBLIC_KEY) { console.log('⚠️ VAPID_PUBLIC_KEY no configurado'); return; }
  try {
    const r = await pool.query('SELECT * FROM push_subscriptions');
    const subs = r.rows;
    console.log(`📱 Enviando push a ${subs.length} dispositivos:`, payload.title);
    if (subs.length === 0) { console.log('⚠️ No hay suscripciones registradas'); return; }
    const badge = 0; // Badge se maneja desde el cliente
    const fullPayload = JSON.stringify({ ...payload, badge });
    await Promise.allSettled(subs.map(async sub => {
      const pushSub = { endpoint: sub.endpoint, keys: { p256dh: sub.p256dh, auth: sub.auth } };
      try {
        await webpush.sendNotification(pushSub, fullPayload);
        console.log(`✓ Push enviado a ${sub.usuario_nombre}`);
      } catch(e) {
        console.error(`✗ Error push a ${sub.usuario_nombre}:`, e.message, e.statusCode);
        if (e.statusCode === 410 || e.statusCode === 404) {
          await pool.query('DELETE FROM push_subscriptions WHERE endpoint=$1', [sub.endpoint]);
          console.log(`🗑️ Suscripción eliminada (expirada): ${sub.usuario_nombre}`);
        }
      }
    }));
  } catch(e) { console.error('Error enviando push:', e.message); }
}

async function sincronizarPedidosWeb(opciones){
  opciones = opciones || {};
  if (!PEDIDOS_EMAIL_HOST || !PEDIDOS_EMAIL_USER || !PEDIDOS_EMAIL_PASS) {
    console.log('⚠️ Pedidos web: variables de entorno de correo no configuradas, omitiendo sync');
    return { ok: false, error: 'Credenciales de correo no configuradas' };
  }

  let client;
  let procesados = 0;
  let errores = 0;
  try {
    client = new ImapFlow({
      host: PEDIDOS_EMAIL_HOST,
      port: PEDIDOS_EMAIL_PORT,
      secure: true,
      auth: { user: PEDIDOS_EMAIL_USER, pass: PEDIDOS_EMAIL_PASS },
      logger: false
    });
    await client.connect();

    const lock = await client.getMailboxLock('INBOX');
    try {
      // Normalmente solo correos no leídos; con incluirLeidos:true (resync manual)
      // se reprocesan TODOS, incluso los ya marcados como leídos.
      const mensajes = opciones.incluirLeidos
        ? await client.search({ all: true })
        : await client.search({ seen: false });
      for (const seq of (mensajes || [])) {
        try {
          const { content } = await client.download(seq, undefined, { uid: false });
          const parsed = await simpleParser(content);
          const asunto = parsed.subject || '';

          // Filtrar solo correos de "nuevo pedido" (evita procesar otros correos que
          // puedan llegar a esa casilla)
          if (!/nuevo pedido/i.test(asunto)) {
            await client.messageFlagsAdd(seq, ['\\Seen']);
            continue;
          }

          const html = parsed.html || parsed.textAsHtml || '';
          const pedido = parsearPedidoWooCommerce(html, asunto, parsed.date || new Date());

          if (pedido && pedido.numeroPedido) {
            // Verificar si ya existe ANTES de insertar para evitar push duplicados
            const yaExiste = await pool.query('SELECT id FROM pedidos_web WHERE numero_pedido=$1', [pedido.numeroPedido]);
            const esNuevo = yaExiste.rows.length === 0;

            await pool.query(
              `INSERT INTO pedidos_web(numero_pedido, fecha, cliente_nombre, cedula_ruc, telefono, subtotal, total, productos, email_uid, html_crudo)
               VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
               ON CONFLICT (numero_pedido) DO UPDATE SET
                 cliente_nombre=$3, cedula_ruc=$4, telefono=$5, subtotal=$6, total=$7, productos=$8, html_crudo=$10`,
              [
                pedido.numeroPedido, pedido.fecha, pedido.cliente || '—',
                pedido.cedulaRuc || null, pedido.telefono || null,
                pedido.subtotal || 0, pedido.total || 0,
                JSON.stringify(pedido.productos || []), String(seq), html
              ]
            );
            procesados++;
            // Solo enviar push si es un pedido GENUINAMENTE NUEVO
            if (esNuevo) {
              console.log(`🛒 Nuevo pedido #${pedido.numeroPedido} — enviando push`);
              // El VALOR va primero en el título: en el móvil el título se recorta y un
              // nombre largo dejaba el monto fuera de pantalla. El nombre pasa al cuerpo,
              // que sí admite dos líneas y no se pierde.
              enviarPushATodos({
                title: 'Nuevo pedido',
                body: `${String(pedido.cliente || '—').substring(0,90)}\n$${parseFloat(pedido.total||0).toFixed(2)}`,
                tag: `pedido-${pedido.numeroPedido}`,
                url: '/'
              }).catch(()=>{});
            } else {
              console.log(`ℹ️ Pedido #${pedido.numeroPedido} ya existía — sin push`);
            }
          } else {
            errores++;
            console.log('⚠️ No se pudo parsear pedido del correo:', asunto);
          }

          await client.messageFlagsAdd(seq, ['\\Seen']);
        } catch (eMsg) {
          errores++;
          console.error('Error procesando correo de pedido:', eMsg.message);
        }
      }
    } finally {
      lock.release();
    }

    await client.logout();
    console.log(`✓ Pedidos web sync: ${procesados} pedidos guardados, ${errores} errores`);
    // Si entraron pedidos nuevos, completar YA sus datos desde la API de WooCommerce
    // (cédula/RUC, teléfono, email, dirección y SKUs) — sin esperar el ciclo periódico
    if (procesados > 0) {
      completarPedidosDesdeWoo(Math.max(procesados, 5)).catch(e => console.error('Woo backfill:', e.message));
    }
    return { ok: true, procesados, errores };
  } catch (e) {
    console.error('Error conectando a pedidos@cosetika.com:', e.message);
    try { if (client) await client.logout(); } catch(e2){}
    return { ok: false, error: e.message };
  }
}

// ─── REFERIDOS: lectura de correos "Nuevo Formulario de referidos" vía IMAP ─────
// La casilla puede ser compartida (ej. info@cosetika.com), por eso la búsqueda se hace
// SOLO por asunto — nunca se tocan ni se marcan como leídos otros correos de la casilla.
function parsearReferido(texto){
  const t = String(texto||'').replace(/\s+/g,' ').trim();
  const m = /Recu[eé]rdanos tu nombre y apellido\s*:?\s*(.*?)\s*Nombre y apellido de tu referido\s*:?\s*(.*?)\s*N[uú]mero de tel[eé]fono de tu referido\s*:?\s*([+\d][\d\s()+.-]*)/i.exec(t);
  if(!m) return null;
  const limpiar = x => String(x||'').replace(/[*_#|]/g,'').trim().substring(0,490);
  return {
    cliente: limpiar(m[1]),
    referido: limpiar(m[2]),
    telefono: String(m[3]||'').replace(/[^\d+]/g,'').substring(0,90)
  };
}

async function sincronizarReferidos(opciones){
  opciones = opciones || {};
  if (!REFERIDOS_EMAIL_HOST || !REFERIDOS_EMAIL_USER || !REFERIDOS_EMAIL_PASS) {
    console.log('⚠️ Referidos: variables de entorno de correo no configuradas, omitiendo sync');
    return { ok: false, error: 'Credenciales de correo de referidos no configuradas (REFERIDOS_EMAIL_USER/PASS)' };
  }
  let client;
  let procesados = 0, nuevos = 0, errores = 0;
  try {
    client = new ImapFlow({
      host: REFERIDOS_EMAIL_HOST,
      port: REFERIDOS_EMAIL_PORT,
      secure: true,
      auth: { user: REFERIDOS_EMAIL_USER, pass: REFERIDOS_EMAIL_PASS },
      logger: false
    });
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');
    try {
      // Búsqueda SOLO por asunto. Normal: solo no leídos; incluirLeidos: todo el historial.
      const criterio = opciones.incluirLeidos
        ? { subject: 'Formulario de referidos' }
        : { subject: 'Formulario de referidos', seen: false };
      const mensajes = await client.search(criterio);
      for (const seq of (mensajes || [])) {
        try {
          const { content } = await client.download(seq, undefined, { uid: false });
          const parsed = await simpleParser(content);
          const texto = parsed.text || String(parsed.html||'').replace(/<[^>]+>/g,' ');
          const ref = parsearReferido(texto);
          if (ref && (ref.referido || ref.telefono)) {
            const msgId = (parsed.messageId || '').substring(0,290) || null;
            const fecha = parsed.date ? new Date(parsed.date).toISOString().substring(0,10) : new Date().toISOString().substring(0,10);
            if (msgId) {
              const r = await pool.query(
                `INSERT INTO referidos(cliente,referido,telefono,fecha,message_id)
                 VALUES($1,$2,$3,$4,$5) ON CONFLICT (message_id) DO NOTHING RETURNING id`,
                [ref.cliente, ref.referido, ref.telefono, fecha, msgId]
              );
              if (r.rows.length > 0) {
                nuevos++;
                // Push solo en el sync automático — no al importar historial (evita avalancha)
                if (!opciones.incluirLeidos) {
                  console.log(`💜 Nuevo referido: ${ref.referido} — enviando push`);
                  enviarPushATodos({
                    title: `Nuevo referido: ${ref.referido || '—'}`,
                    body: `Tel: ${ref.telefono || '—'} · Refiere: ${ref.cliente || '—'}`,
                    tag: `referido-${r.rows[0].id}`,
                    url: '/'
                  }).catch(()=>{});
                }
              }
            } else {
              // Sin message-id: dedup manual por contenido+fecha
              const ya = await pool.query(
                'SELECT id FROM referidos WHERE cliente=$1 AND referido=$2 AND telefono=$3 AND fecha=$4',
                [ref.cliente, ref.referido, ref.telefono, fecha]
              );
              if (ya.rows.length === 0) {
                await pool.query('INSERT INTO referidos(cliente,referido,telefono,fecha) VALUES($1,$2,$3,$4)',
                  [ref.cliente, ref.referido, ref.telefono, fecha]);
                nuevos++;
                if (!opciones.incluirLeidos) {
                  enviarPushATodos({
                    title: `Nuevo referido: ${ref.referido || '—'}`,
                    body: `Tel: ${ref.telefono || '—'} · Refiere: ${ref.cliente || '—'}`,
                    tag: `referido-nuevo-${Date.now()}`,
                    url: '/'
                  }).catch(()=>{});
                }
              }
            }
            procesados++;
          } else {
            errores++;
            console.log('⚠️ No se pudo parsear referido del correo:', parsed.subject);
          }
          await client.messageFlagsAdd(seq, ['\\Seen']);
        } catch (eMsg) { errores++; console.error('Error procesando correo de referido:', eMsg.message); }
      }
    } finally { lock.release(); }
    await client.logout();
    console.log(`✓ Referidos sync: ${procesados} procesados, ${nuevos} nuevos, ${errores} errores`);
    return { ok: true, procesados, nuevos, errores };
  } catch (e) {
    // Armar un mensaje descriptivo: ImapFlow lanza "Command failed" genérico,
    // pero trae el detalle en authenticationFailed/responseText/serverResponseCode
    let detalle = e.message || 'Error desconocido';
    if (e.authenticationFailed) {
      detalle = 'Login rechazado intentando como "' + REFERIDOS_EMAIL_USER + '" en ' + REFERIDOS_EMAIL_HOST + ':' + REFERIDOS_EMAIL_PORT
        + (e.responseText ? ' — el servidor dijo: ' + e.responseText : '')
        + '. Verifica la contraseña en Railway (sin espacios ni comillas) y prueba entrar en webmail.dreamhost.com con esas mismas credenciales.';
    } else if (e.responseText) {
      detalle = detalle + ' — respuesta del servidor: ' + e.responseText;
    } else if (e.code) {
      detalle = detalle + ' (código: ' + e.code + ', host: ' + REFERIDOS_EMAIL_HOST + ':' + REFERIDOS_EMAIL_PORT + ')';
    }
    console.error('Error conectando a casilla de referidos:', detalle);
    try { if (client) await client.logout(); } catch(e2){}
    return { ok: false, error: detalle };
  }
}

// Sync de referidos cada 20 segundos (igual que pedidos)
setTimeout(() => sincronizarReferidos().catch(e => console.error('Error sync referidos inicial:', e.message)), 20000);
setInterval(() => sincronizarReferidos().catch(e => console.error('Error sync referidos:', e.message)), 20 * 1000);

// ─── INSTITUTOS: alumnas marcadas en el campo adicional de Contifico ────────
async function getConfigApp(clave, porDefecto){
  try{
    const r = await pool.query('SELECT valor FROM app_config WHERE clave=$1',[clave]);
    return r.rows.length ? r.rows[0].valor : porDefecto;
  }catch(e){ return porDefecto; }
}
async function setConfigApp(clave, valor){
  await pool.query('INSERT INTO app_config(clave,valor) VALUES($1,$2) ON CONFLICT (clave) DO UPDATE SET valor=$2',[clave,valor]);
}

let INSTITUTOS_ULTIMA_SYNC = null;
const CAMPOS_ADICIONALES = ['adicional1_cliente','adicional2_cliente','adicional3_cliente','adicional4_cliente'];

// ─── SEGURIDAD: token de sesión firmado y guardia de administrador ──────────
const zlib = require('zlib');
const crypto = require('crypto');
const SESION_SECRET = process.env.SESION_SECRET || (process.env.CONTIFICO_API_KEY || 'cosetika') + '::sesion';
function firmarSesion(u){
  const payload = Buffer.from(JSON.stringify({ id: u.id, rol: u.rol, nombre: u.nombre })).toString('base64url');
  const firma = crypto.createHmac('sha256', SESION_SECRET).update(payload).digest('base64url');
  return payload + '.' + firma;
}
function leerSesion(req){
  try {
    let t = String(req.headers['x-sesion'] || '').trim();
    if (!t) {
      const ck = String(req.headers.cookie || '');
      const m = ck.match(/(?:^|;\s*)cosetika_ses=([^;]+)/);
      if (m) t = decodeURIComponent(m[1]);
    }
    if (!t || !t.includes('.')) return null;
    const [payload, firma] = t.split('.');
    const esperada = crypto.createHmac('sha256', SESION_SECRET).update(payload).digest('base64url');
    if (firma !== esperada) return null;
    return JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
  } catch(e) { return null; }
}
// Devuelve true si la petición NO es de un admin (y ya respondió 403)
function bloquearSiNoAdmin(req, res){
  const s = leerSesion(req);
  if (s && s.rol === 'admin') return false;
  res.writeHead(403, {'Content-Type':'application/json'});
  res.end(JSON.stringify({ ok:false, error:'Acceso restringido: solo el administrador puede ver esta información' }));
  return true;
}

// ─── PEDIDOS WEB: completar cédula/RUC y SKUs desde la API de WooCommerce ───
// El correo de WooCommerce a veces trae los datos de pago en vez de la cédula; la API
// siempre entrega el pedido completo, así que la usamos como fuente de verdad.
function cedulaDesdeOrdenWoo(orden){
  const bill = orden.billing || {};
  const candidatos = [];
  (orden.meta_data || []).forEach(m => {
    const k = String((m && m.key) || '').toLowerCase();
    if (/cedula|c\u00e9dula|ruc|identificacion|identificaci\u00f3n|dni|nit/.test(k)) candidatos.push(String((m && m.value) || ''));
  });
  Object.entries(bill).forEach(([k, v]) => {
    if (/cedula|ruc|identific|dni/.test(String(k).toLowerCase())) candidatos.push(String(v||''));
  });
  Object.entries(bill).forEach(([k, v]) => {
    const kl = String(k).toLowerCase();
    if (kl.includes('phone') || kl.includes('postcode') || kl.includes('telefono')) return;
    candidatos.push(String(v||''));
  });
  (orden.meta_data || []).forEach(m => {
    const k = String((m && m.key) || '').toLowerCase();
    if (k.includes('phone') || k.includes('telefono')) return;
    candidatos.push(String((m && m.value) || ''));
  });
  for (const c of candidatos) {
    const d = String(c).replace(/\D/g, '');
    if (d.length === 10 || d.length === 13) return d;
  }
  return '';
}

async function completarPedidosDesdeWoo(limite = 25){
  const WOO_URL0 = (process.env.WOO_URL || '').trim();
  const mU = WOO_URL0.match(/https?:\/\/[^\s"']+/);
  const WOO_URL = (mU ? mU[0] : WOO_URL0).replace(/\/+$/, '');
  const CK = process.env.WOO_CK || process.env.WC_CONSUMER_KEY || '';
  const CS = process.env.WOO_CS || process.env.WC_CONSUMER_SECRET || '';
  if (!WOO_URL || !CK || !CS) return { ok:false, error:'Faltan credenciales de WooCommerce' };
  try {
    // Solo pedidos de HOY en adelante (los históricos se dejan como están)
    const r = await pool.query(
      `SELECT numero_pedido FROM pedidos_web
       WHERE ((cedula_ruc IS NULL OR LENGTH(REGEXP_REPLACE(cedula_ruc,'\\D','','g')) NOT IN (10,13))
              OR direccion IS NULL OR direccion = '')
         AND fecha >= CURRENT_DATE ORDER BY id DESC LIMIT $1`, [limite]);
    let actualizados = 0;
    for (const row of r.rows) {
      const num = row.numero_pedido;
      try {
        const resp = await fetch(`${WOO_URL}/wp-json/wc/v3/orders/${encodeURIComponent(num)}?consumer_key=${encodeURIComponent(CK)}&consumer_secret=${encodeURIComponent(CS)}`, { headers:{'Accept':'application/json'} });
        if (!resp.ok) continue;
        const orden = await resp.json();
        const ced = cedulaDesdeOrdenWoo(orden);
        // Productos con SKU real desde la API
        const prods = (orden.line_items || []).map(it => ({
          sku: it.sku || '', nombre: it.name || '',
          cantidad: parseFloat(it.quantity || 0),
          total: Math.round((parseFloat(it.total || 0) + parseFloat(it.total_tax || 0)) * 100) / 100
        }));
        const b = orden.billing || {}; const sh = orden.shipping || {};
        const tel = String(b.phone || '').replace(/\D/g,'').substring(0,20);
        const dir = [sh.address_1 || b.address_1, sh.address_2 || b.address_2].filter(Boolean).join(' · ').substring(0,390);
        const ciu = String(sh.city || b.city || '').substring(0,140);
        const prov = String(sh.state || b.state || '').substring(0,140);
        const mail = String(b.email || '').substring(0,190);
        const nota = String(orden.customer_note || '').substring(0,900);
        await pool.query(
          `UPDATE pedidos_web SET cedula_ruc   = COALESCE(NULLIF($1,''), cedula_ruc),
                                  telefono     = COALESCE(NULLIF($2,''), telefono),
                                  productos    = CASE WHEN $3::text <> '[]' THEN $3 ELSE productos END,
                                  email        = COALESCE(NULLIF($5,''), email),
                                  direccion    = COALESCE(NULLIF($6,''), direccion),
                                  ciudad       = COALESCE(NULLIF($7,''), ciudad),
                                  provincia_env= COALESCE(NULLIF($8,''), provincia_env),
                                  nota_cliente = COALESCE(NULLIF($9,''), nota_cliente)
           WHERE numero_pedido = $4`,
          [ced, tel, JSON.stringify(prods), num, mail, dir, ciu, prov, nota]);
        if (ced) actualizados++;
      } catch(e) {}
    }
    if (actualizados) console.log(`✓ Cédulas completadas desde WooCommerce: ${actualizados} pedido(s)`);
    return { ok:true, revisados: r.rows.length, actualizados };
  } catch(e) { console.error('Error completando pedidos desde Woo:', e.message); return { ok:false, error:e.message }; }
}
setTimeout(() => completarPedidosDesdeWoo(20).catch(e=>console.error(e)), 60 * 1000);
setInterval(() => completarPedidosDesdeWoo(20).catch(e=>console.error(e)), 60 * 1000);

// ─── CRÉDITO DE CLIENTES (cupo y días) desde Contifico ──────────────────────
let CREDITO_CACHE = {};
let CREDITO_SYNC_AT = null;
let CREDITO_SYNC_LOG = { estado:'sin correr', paginas:0, personas:0, error:null };
async function sincronizarCreditos(){
  if (!API_KEY) { CREDITO_SYNC_LOG = { estado:'sin API_KEY', paginas:0, personas:0, error:'CONTIFICO_API_KEY no configurada' }; return; }
  try {
    const personas = [];
    let nextUrl = 'https://api.contifico.com/sistema/api/v2/persona/?page_size=100';
    let pag = 0;
    while (nextUrl && pag < 500) {
      const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!resp.ok) { CREDITO_SYNC_LOG = { estado:'HTTP '+resp.status, paginas:pag, personas:personas.length, error:(await resp.text().catch(()=>'')).slice(0,300) }; break; }
      const data = await resp.json();
      if (Array.isArray(data)) { personas.push(...data); nextUrl = null; }
      else { personas.push(...(data.results || [])); nextUrl = data.next || null; }
      pag++;
    }
    const mapa = {};
    personas.forEach(p => {
      const info = {
        cupo: parseFloat(p.cupo_credito) || 0,
        dias: parseInt(p.dias_credito) || 0,
        aplica: String(p.aplicar_cupo) === 'True' || p.aplicar_cupo === true,
        nombre: p.razon_social || ''
      };
      [p.cedula, p.ruc].forEach(v => {
        const d = String(v || '').replace(/\D/g, '');
        if (d) { mapa[d] = info; if (d.length === 13) mapa[d.substring(0,10)] = info; }
      });
    });
    if (Object.keys(mapa).length) {
      CREDITO_CACHE = mapa;
      CREDITO_SYNC_AT = new Date().toISOString();
      const conCupo = personas.filter(p => (parseFloat(p.cupo_credito)||0) > 0).length;
      CREDITO_SYNC_LOG = { estado:'ok', paginas:pag, personas:personas.length, con_cupo:conCupo, error:null };
      console.log(`✓ Créditos: ${personas.length} personas · ${conCupo} con cupo asignado`);
    } else if (CREDITO_SYNC_LOG.estado === 'sin correr') {
      CREDITO_SYNC_LOG = { estado:'respuesta vacía', paginas:pag, personas:personas.length, error:'Contifico no devolvió personas' };
    }
  } catch(e) { CREDITO_SYNC_LOG = { estado:'excepción', paginas:0, personas:0, error:e.message }; console.error('Error sincronizando créditos:', e.message); }
}
setTimeout(() => sincronizarCreditos(), 15 * 1000);
setInterval(() => sincronizarCreditos(), 60 * 60 * 1000);

// Consulta en vivo a Contifico el crédito de una cédula/RUC puntual (con TTL corto)
const CREDITO_VIVO = {};   // digitos -> { info, ts }
const CREDITO_VIVO_TTL = 10 * 60 * 1000;
async function creditoEnVivo(digitos) {
  const d = String(digitos || '').replace(/\D/g, '');
  if (!d || !API_KEY) return null;
  const ca = CREDITO_VIVO[d];
  if (ca && (Date.now() - ca.ts) < CREDITO_VIVO_TTL) return ca.info;
  let info = null;
  for (const campo of (d.length === 13 ? ['ruc','cedula'] : ['cedula','ruc'])) {
    try {
      const rr = await fetch(`https://api.contifico.com/sistema/api/v1/persona/?${campo}=${d}&page_size=5`, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!rr.ok) continue;
      const dd = await rr.json();
      const lista = Array.isArray(dd) ? dd : (dd.results || []);
      if (!lista.length) continue;
      // Si hay varias fichas de la misma persona, nos quedamos con la de mayor cupo
      const mejor = lista.slice().sort((a,b) => (parseFloat(b.cupo_credito)||0) - (parseFloat(a.cupo_credito)||0))[0];
      info = {
        cupo: parseFloat(mejor.cupo_credito) || 0,
        dias: parseInt(mejor.dias_credito) || 0,
        aplica: String(mejor.aplicar_cupo) === 'True' || mejor.aplicar_cupo === true,
        nombre: mejor.razon_social || ''
      };
      break;
    } catch(e) {}
  }
  CREDITO_VIVO[d] = { info, ts: Date.now() };
  if (info && info.cupo > 0) { CREDITO_CACHE[d] = info; if (d.length === 13) CREDITO_CACHE[d.substring(0,10)] = info; }
  return info;
}

async function sincronizarInstitutos(){
  if (!API_KEY) return { ok:false, error:'CONTIFICO_API_KEY no configurada' };
  const campo = await getConfigApp('instituto_campo', 'adicional2_cliente');
  // Leer TODAS las personas de Contifico (paginado v2, fallback v1)
  const personasApi = [];
  let nextUrl = 'https://api.contifico.com/sistema/api/v2/persona/?page_size=100';
  let paginas = 0;
  try {
    while (nextUrl && paginas < 500) {
      const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json();
      if (Array.isArray(data)) { personasApi.push(...data); nextUrl = null; }
      else { personasApi.push(...(data.results || [])); nextUrl = data.next || null; }
      paginas++;
    }
  } catch (e) {
    try {
      const resp = await fetch('https://api.contifico.com/sistema/api/v1/persona/?es_cliente=true', { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await resp.json();
      if (Array.isArray(data)) personasApi.push(...data); else throw e;
    } catch (e2) {
      console.error('Error leyendo personas de Contifico:', e.message);
      return { ok:false, error:'No se pudo leer personas de Contifico: ' + e.message };
    }
  }
  // Diagnóstico: cuántas personas tienen valor en cada campo adicional
  const conteo = {}; const ejemplos = {};
  CAMPOS_ADICIONALES.forEach(c => {
    conteo[c] = 0; ejemplos[c] = [];
    personasApi.forEach(pa => {
      const v = String(pa[c] || '').trim();
      if (v) { conteo[c]++; if (ejemplos[c].length < 5 && !ejemplos[c].includes(v)) ejemplos[c].push(v); }
    });
  });
  const alumnas = personasApi
    .map(pa => ({
      valor: String(pa[campo] || '').trim().toUpperCase(),
      cedula: String(pa.cedula || '').replace(/\D/g, ''),
      ruc: String(pa.ruc || '').replace(/\D/g, ''),
      nombre: String(pa.razon_social || pa.nombre_comercial || '').trim()
    }))
    .filter(a => a.valor);
  // Reset y re-aplicación (sync completo = fuente de verdad es Contifico)
  await pool.query('UPDATE personas SET instituto=NULL WHERE instituto IS NOT NULL');
  let actualizadas = 0, insertadas = 0;
  for (const a of alumnas) {
    let hecho = false;
    if (a.cedula) {
      const r = await pool.query('UPDATE personas SET instituto=$1 WHERE cedula=$2', [a.valor, a.cedula]);
      if (r.rowCount > 0) { actualizadas += r.rowCount; hecho = true; }
    }
    if (!hecho && a.ruc) {
      const r = await pool.query('UPDATE personas SET instituto=$1 WHERE ruc=$2', [a.valor, a.ruc]);
      if (r.rowCount > 0) { actualizadas += r.rowCount; hecho = true; }
    }
    if (!hecho && a.nombre) {
      const r = await pool.query('UPDATE personas SET instituto=$1 WHERE LOWER(razon_social)=LOWER($2)', [a.valor, a.nombre]);
      if (r.rowCount > 0) { actualizadas += r.rowCount; hecho = true; }
    }
    if (!hecho) {
      await pool.query("INSERT INTO personas(cedula,ruc,razon_social,instituto,origen) VALUES($1,$2,$3,$4,'institutos')",
        [a.cedula || null, a.ruc || null, a.nombre || '—', a.valor]);
      insertadas++;
    }
  }
  INSTITUTOS_ULTIMA_SYNC = new Date().toISOString();
  console.log(`✓ Institutos sync: ${alumnas.length} alumnas (campo ${campo}) de ${personasApi.length} personas — ${actualizadas} actualizadas, ${insertadas} insertadas`);
  return { ok:true, alumnas: alumnas.length, actualizadas, insertadas, campo, personas_leidas: personasApi.length, conteo_por_campo: conteo, ejemplos_por_campo: ejemplos };
}
// ─── BODEGAS: stock por producto de Bodega POS y Bodega Casa ────────────────
let BODEGAS_SYNC = { enCurso: false, procesados: 0, total: 0 };

// Solo las 4 marcas propias, sin combos/armados: se excluyen las "Promo ...",
// las "Línea ..." (líneas completas = combos) y el Kit Básico de Ziaja Pro
function filtrarProductosBodega(lista){
  const MARCAS_BODEGA = ['BIOSKIN','ERAYBA','ZIAJA','ZIAJAPRO'];
  return lista.filter(pr => {
    const marcaN = String(pr.marca||'').toUpperCase().replace(/\s+/g,'');
    if (!MARCAS_BODEGA.includes(marcaN)) return false;
    const nombreN = String(pr.nombre||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().trim();
    if (nombreN.includes('PROMO')) return false;       // Promo Kit / Promo Línea Completa (todas las marcas)
    if (nombreN.startsWith('LINEA ')) return false;    // Línea Acai Berry, Línea Completa Curls, etc. (combos)
    if (/\bKITS?\b/.test(nombreN)) return false;      // Kits (Ziaja Pro, Bioskin, etc.) — no son productos en sí
    return true;
  });
}

async function sincronizarBodegas(){
  if (BODEGAS_SYNC.enCurso) return { ok:false, error:'Sincronización ya en curso' };
  if (!API_KEY) return { ok:false, error:'CONTIFICO_API_KEY no configurada' };
  const ids = Object.keys(catalogoProductos || {});
  if (ids.length === 0) return { ok:false, error:'El catálogo de productos aún no está cargado — intenta en un minuto' };
  BODEGAS_SYNC = { enCurso: true, procesados: 0, total: ids.length };
  let guardados = 0, errores = 0;
  try {
    for (const pid of ids) {
      BODEGAS_SYNC.procesados++;
      try {
        const resp = await fetch(`https://api.contifico.com/sistema/api/v1/producto/${pid}/stock/`, {
          headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
        });
        if (!resp.ok) { errores++; continue; }
        const bodegas = await resp.json();
        if (!Array.isArray(bodegas)) { errores++; continue; }
        // Solo interesan Bodega POS y Bodega Casa
        const relevantes = bodegas.filter(b => /pos|casa/i.test(String(b.bodega_nombre||'')));
        const info = catalogoProductos[pid] || {};
        for (const b of relevantes) {
          await pool.query(
            `INSERT INTO stock_bodegas(producto_id, codigo, nombre, marca, bodega, cantidad, actualizado_at)
             VALUES($1,$2,$3,$4,$5,$6,NOW())
             ON CONFLICT (producto_id, bodega) DO UPDATE SET
               codigo=$2, nombre=$3, marca=$4, cantidad=$6, actualizado_at=NOW()`,
            [pid, info.codigo||'', info.nombre||'', info.marca||'', String(b.bodega_nombre||'').substring(0,190), parseFloat(b.cantidad||0)]
          );
          guardados++;
        }
      } catch(e) { errores++; }
    }
    await setConfigApp('bodegas_ultima_sync', new Date().toISOString());
    // La Proyección (Inventario) se alimenta de estas mismas cantidades: reconstruir
    await reconstruirInventarioDesdeBodegas();
    console.log(`✓ Bodegas sync: ${ids.length} productos, ${guardados} registros, ${errores} errores`);
    return { ok:true, productos: ids.length, guardados, errores };
  } finally {
    BODEGAS_SYNC.enCurso = false;
  }
}

// Reemplaza el antiguo Excel de saldos: el stock de la Proyección es la suma de
// Bodega POS + Bodega Casa sincronizadas desde Contifico, con corte = hoy.
async function reconstruirInventarioDesdeBodegas(){
  try {
    const r = await pool.query('SELECT producto_id, codigo, cantidad FROM stock_bodegas');
    if (r.rows.length === 0) return;
    const productos = {};
    r.rows.forEach(row => {
      if (!productos[row.producto_id]) productos[row.producto_id] = { cantidad: 0, sku: row.codigo || '' };
      productos[row.producto_id].cantidad += parseFloat(row.cantidad) || 0; // POS + Casa
    });
    const fechaCorte = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Guayaquil' });
    INVENTARIO_CACHE = { fecha_corte: fechaCorte, productos };
    INVENTARIO_CACHE_TS = new Date().toISOString();
    await pool.query(
      `INSERT INTO inventario_data(id_unico, fecha_corte, datos, actualizado_at) VALUES('principal',$1,$2,NOW())
       ON CONFLICT (id_unico) DO UPDATE SET fecha_corte=$1, datos=$2, actualizado_at=NOW()`,
      [fechaCorte, JSON.stringify(INVENTARIO_CACHE)]
    );
    console.log(`✓ Proyección reconstruida desde bodegas: ${Object.keys(productos).length} productos, corte ${fechaCorte}`);
  } catch(e) { console.error('Error reconstruyendo proyección desde bodegas:', e.message); }
}

// Sync diario: al arrancar (si los datos tienen más de 20h) y chequeo cada 6 horas
async function chequearSyncBodegas(){
  try {
    const ult = await getConfigApp('bodegas_ultima_sync', null);
    if (!ult || (Date.now() - new Date(ult).getTime()) > 20*60*60*1000) {
      sincronizarBodegas().catch(e => console.error('Error sync bodegas:', e.message));
    }
  } catch(e) {}
}
setTimeout(chequearSyncBodegas, 120000);
setInterval(chequearSyncBodegas, 6 * 60 * 60 * 1000);

setTimeout(() => sincronizarInstitutos().catch(e => console.error('Error sync institutos inicial:', e.message)), 40000);
setInterval(() => sincronizarInstitutos().catch(e => console.error('Error sync institutos:', e.message)), 12 * 60 * 60 * 1000);

sincronizarHoy().catch(e => console.error('Error sync inicial:', e.message));
// Cada 5 minutos: solo consulta los documentos de HOY (una o dos páginas), así que es
// barato. Antes corría cada hora y una factura emitida en Contifico podía tardar hasta
// 60 minutos en aparecer en la app.
setInterval(() => sincronizarHoy().catch(e => console.error('Error sync:', e.message)), 5 * 60 * 1000);

// Sync de pedidos web: revisa la casilla pedidos@cosetika.com cada 10 minutos
setTimeout(() => sincronizarPedidosWeb().catch(e => console.error('Error sync pedidos inicial:', e.message)), 15000);
setInterval(() => sincronizarPedidosWeb().catch(e => console.error('Error sync pedidos:', e.message)), 30 * 1000);

// ─── FUSIÓN INCREMENTAL: ventas del MES EN CURSO dentro de DATA_CACHE (cada 15 min) ──
// Recalcula desde cero el mes actual completo (rápido: solo ese mes, no 18 meses) y
// reemplaza limpiamente esa porción en cada cliente, dejando el resto del historial intacto.
function consolidarMarcasAnio(lista){
  const mapa = {};
  lista.forEach(x=>{
    const k = x.anio+'|'+x.marca;
    if(!mapa[k]) mapa[k] = { anio:x.anio, marca:x.marca, total:0 };
    mapa[k].total += x.total;
  });
  return Object.values(mapa).map(x=>({...x, total: Math.round(x.total*100)/100}));
}
function consolidarMarcasMes(lista){
  const mapa = {};
  lista.forEach(x=>{
    const k = x.anio+'|'+x.mes+'|'+x.marca;
    if(!mapa[k]) mapa[k] = { anio:x.anio, mes:x.mes, marca:x.marca, total:0 };
    mapa[k].total += x.total;
  });
  return Object.values(mapa).map(x=>({...x, total: Math.round(x.total*100)/100}));
}

function consolidarProductosMes(lista){
  const mapa = {};
  lista.forEach(x=>{
    const k = x.anio+'|'+x.mes+'|'+(x.id||x.nombre);
    if(!mapa[k]) mapa[k] = { anio:x.anio, mes:x.mes, id:x.id, nombre:x.nombre, marca:x.marca, cantidad:0, total:0 };
    mapa[k].cantidad += x.cantidad;
    mapa[k].total += x.total;
  });
  return Object.values(mapa).map(x=>({...x, cantidad: Math.round(x.cantidad*100)/100, total: Math.round(x.total*100)/100}));
}

// Nota de crédito: la API de Contifico puede reportar el tipo como 'NC', 'NCR', 'N/C'
// o el nombre completo — reconocer TODAS las variantes (un 'NC' exacto dejaba pasar
// notas de crédito como si fueran ventas positivas)
// Documentos que NO son ventas reales: cotización, proforma y PREFACTURA (las que crea
// la app desde los pedidos web). Sin esto, las prefacturas inflaban ventas y facturas.
function noEsVenta(d) {
  const t = String((d && d.tipo_documento) || '').toUpperCase().trim();
  const tr = String((d && d.tipo_registro) || '').toUpperCase().trim();
  // DAC = documento de anticipo de cliente. NO es una venta: el propio reporte de Contifico
  // lo lista aparte, en la línea "Anticipos", y no lo suma al total de ventas. Hasta ahora
  // quedaba fuera solo porque los anticipos no llevan vendedora asignada — un accidente, no
  // una regla. Bastaba que alguien le pusiera vendedora a uno para inflar las ventas del mes.
  return t === 'COT' || t === 'PRO' || t === 'PRE' || t === 'DAC'
      || tr === 'PRE' || tr === 'PRO' || tr === 'COT' || tr === 'DAC';
}

function esNotaCredito(d) {
  const t = String((d && d.tipo_documento) || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  // 'NCT' es el código real que usa Contifico; startsWith cubre NC/NCR/NCT y futuros
  return t === 'N/C' || t.startsWith('NC') || t.includes('CREDITO');
}

let regenerandoEnProceso = false;

async function fusionarMesActualEnCache() {
  if (!DATA_CACHE || Object.keys(DATA_CACHE).length === 0) return;
  if (regenerandoEnProceso) { console.log('Fusión incremental omitida: regeneración en curso'); return; }
  regenerandoEnProceso = true;
  try {
    const hoy = nowEC();
    const anioActual = hoy.getFullYear();
    const mesActual = hoy.getMonth() + 1;
    const desde = fmtDateEC(new Date(anioActual, hoy.getMonth(), 1));
    const hasta = fmtDateEC(hoy);

    // Obtener data del año completo hasta hoy desde Contifico
    const fi = `${anioActual}-01-01`;
    const dataMes = await generarDataJson(fi, hasta);

    // Reemplazar DATA_CACHE completamente con datos frescos del año actual
    // Esto evita cualquier acumulación o duplicación — siempre es la fuente de verdad
    Object.entries(dataMes).forEach(([vendNom, clientesFrescos]) => {
      DATA_CACHE[vendNom] = clientesFrescos;
    });

    // Reordenar
    Object.keys(DATA_CACHE).forEach(v => DATA_CACHE[v].sort((a,b)=>b.total-a.total));
    await guardarDataEnDB(DATA_CACHE);
    try { await setConfigApp('data_hasta', hasta); PEND_CACHE = { ts:0, clave:'', data:null }; } catch(e) {}
    console.log(`✓ Fusión completa (${fi} - ${hasta}): ${Object.keys(DATA_CACHE).length} vendedoras`);
  } catch(e) {
    console.error('Error en fusión:', e.message);
  }
  regenerandoEnProceso = false;
}
// DESACTIVADO: esta fusión cada 15 min chocaba con la regeneración nocturna (mismo candado
// regenerandoEnProceso) y hacía que la regen de las 2 AM se saltara aleatoriamente. Las ventas
// de hoy ya viven en el caché en vivo (sincronizarHoy) — este proceso era redundante.
// setInterval(() => fusionarMesActualEnCache().catch(e => console.error(e)), 15 * 60 * 1000);

// ─── FUSIÓN: AÑO EN CURSO completo dentro de DATA_CACHE (regeneración diaria 2 AM) ──
// Misma idea que fusionarMesActualEnCache, pero reemplaza el AÑO actual completo en vez de
// solo el mes en curso. Así la regeneración nocturna corrige cualquier dato retroactivo del
// año en curso (ej. una factura de marzo editada/anulada en Contifico) sin tener que volver
// a traer ni tocar años anteriores (2025 y previos), que ya están cerrados y no cambian.
async function fusionarAnioActualEnCache(anioActual, dataAnio) {
  // Paso 1: quitar de DATA_CACHE cualquier dato del año actual (será reemplazado limpio)
  Object.keys(DATA_CACHE).forEach(vendNom => {
    DATA_CACHE[vendNom].forEach(cli => {
      const freqAnioViejo = (cli.frecuencia||[]).filter(f=>f.anio===anioActual);
      const totalAnioViejo = freqAnioViejo.reduce((a,f)=>a+f.total,0);
      const subtotalAnioViejo = freqAnioViejo.reduce((a,f)=>a+f.subtotal,0);
      const comprasAnioViejo = freqAnioViejo.reduce((a,f)=>a+f.compras,0);
      cli.total = Math.round((cli.total - totalAnioViejo)*100)/100;
      cli.subtotal = Math.round((cli.subtotal - subtotalAnioViejo)*100)/100;
      cli.num_compras = Math.max(0, (cli.num_compras||0) - comprasAnioViejo);
      cli.frecuencia = (cli.frecuencia||[]).filter(f => f.anio!==anioActual);
      cli.frecuencia_dia = (cli.frecuencia_dia||[]).filter(f => f.anio!==anioActual);
      cli.marcas_anio = (cli.marcas_anio||[]).filter(ma => ma.anio!==anioActual);
      cli.marcas_mes = (cli.marcas_mes||[]).filter(x => x.anio!==anioActual);
      cli.productos_mes = (cli.productos_mes||[]).filter(x => x.anio!==anioActual);
      cli.saldo = 0; // se recalcula abajo con los datos frescos del año
    });
  });

  // Paso 2: insertar los datos frescos del año actual completo
  Object.entries(dataAnio).forEach(([vendNom, clientesAnio]) => {
    if (!DATA_CACHE[vendNom]) DATA_CACHE[vendNom] = [];
    const porId = {}; DATA_CACHE[vendNom].forEach(c => { porId[c.id] = c; });
    const nuevos = [];
    clientesAnio.forEach(cliAnio => {
      let cli = porId[cliAnio.id];
      if (!cli) { nuevos.push(cliAnio); return; }
      cli.total = Math.round((cli.total + cliAnio.total) * 100) / 100;
      cli.subtotal = Math.round((cli.subtotal + cliAnio.subtotal) * 100) / 100;
      cli.num_compras = (cli.num_compras||0) + cliAnio.num_compras;
      // La provincia también debe actualizarse aquí: generarDataJson ya la recalcula
      // correctamente con el override más reciente, pero esta función nunca la copiaba
      // de vuelta al cliente existente en DATA_CACHE (por eso un override subido después
      // de que el cliente ya existiera en caché nunca se reflejaba, aun regenerando).
      if(cliAnio.provincia) cli.provincia = cliAnio.provincia;
      if(cliAnio.telefono) cli.telefono = cliAnio.telefono;
      if(cliAnio.direccion) cli.direccion = cliAnio.direccion;
      cli.saldo = Math.round((parseFloat(cliAnio.saldo)||0)*100)/100;
      cli.frecuencia = (cli.frecuencia||[]).concat(cliAnio.frecuencia);
      cli.frecuencia_dia = (cli.frecuencia_dia||[]).concat(cliAnio.frecuencia_dia||[]);
      cli.marcas_anio = (cli.marcas_anio||[]).concat(cliAnio.marcas_anio);
      cli.marcas_mes = (cli.marcas_mes||[]).concat(cliAnio.marcas_mes);
      cli.productos_mes = consolidarProductosMes((cli.productos_mes||[]).concat(cliAnio.productos_mes||[]));
      // cli.marcas: reconstruir COMPLETO desde marcas_anio (ya correctamente filtrado/
      // reinsertado arriba), nunca acumular sobre el cli.marcas anterior.
      cli.marcas = consolidarMarcasAnio((cli.marcas_anio||[]).map(ma=>({anio:0,marca:ma.marca,total:ma.total})))
        .map(m=>({marca:m.marca,total:m.total})).sort((a,b)=>b.total-a.total);
      // Productos: reconstruir desde productos_mes (ya filtrado/concatenado arriba con los años
      // anteriores intactos + el año actual fresco) en vez de sumar incrementalmente sobre
      // cli.productos — así se evita cualquier riesgo de doble conteo, sin depender de un
      // campo "histórico" separado que podría quedar desincronizado entre regeneraciones.
      const prodMap = {};
      (cli.productos_mes||[]).forEach(pm=>{
        const k = pm.id||pm.nombre;
        if(!prodMap[k]) prodMap[k] = { id: pm.id, nombre: pm.nombre, codigo: '', marca: pm.marca, cantidad: 0, total: 0 };
        prodMap[k].cantidad += pm.cantidad;
        prodMap[k].total = Math.round((prodMap[k].total + pm.total)*100)/100;
      });
      // Conservar el "codigo" si ya existía en cli.productos (productos_mes no lo guarda)
      (cli.productos||[]).forEach(p=>{ const k=p.id||p.nombre; if(prodMap[k] && p.codigo) prodMap[k].codigo = p.codigo; });
      cli.productos = Object.values(prodMap).map(p=>({...p, cantidad: Math.round(p.cantidad)})).sort((a,b)=>b.cantidad-a.cantidad);
      // Reset de los acumuladores de la fusión mensual (15 min): tras una fusión anual,
      // el "histórico" para la próxima fusión mensual debe ser este productos recién calculado.
      cli.productos_historico = null;
      cli.productos_historico_anio = null;
      cli.productos_historico_mes = null;
    });
    DATA_CACHE[vendNom] = Object.values(porId).concat(nuevos);
  });

  Object.keys(DATA_CACHE).forEach(v => DATA_CACHE[v].sort((a,b)=>b.total-a.total));
}
setTimeout(() => fusionarMesActualEnCache().catch(e => console.error(e)), 20 * 1000);

// ─── DB ───────────────────────────────────────────────────────
// ─── CACHÉ DE DATA EN MEMORIA (cargada desde PostgreSQL al arrancar) ─────────
let DATA_CACHE = null;
let DATA_CACHE_TS = null;

// ─── CACHÉ DE INVENTARIO (snapshot de bodega, subido manualmente por Fernando) ──
// Estructura: { fecha_corte: 'YYYY-MM-DD', productos: { [producto_id]: cantidad } }
// productos_id es el mismo id que usa catalogoProductos (no el SKU corto del Excel,
// que solo se usa para hacer el match en el momento de la carga).
let INVENTARIO_CACHE = null;
let INVENTARIO_CACHE_TS = null;

// ─── OVERRIDE DE PROVINCIAS POR CLIENTE (subido manualmente por Fernando) ───
// Estructura: { [rucOCedula]: 'NOMBRE_PROVINCIA' }. Tiene prioridad máxima sobre
// provinciaDesdeDir (inferencia por palabras clave en la dirección, que se usa
// solo como respaldo cuando el cliente no aparece en el Excel) — ver resolverProvinciaCliente().
let PROVINCIAS_OVERRIDE = {};
let PROVINCIAS_OVERRIDE_TS = null;

async function cargarProvinciasOverrideDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM provincias_override ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      PROVINCIAS_OVERRIDE = JSON.parse(r.rows[0].datos);
      PROVINCIAS_OVERRIDE_TS = new Date().toISOString();
      console.log('✓ Override de provincias cargado desde PostgreSQL: ' + Object.keys(PROVINCIAS_OVERRIDE).length + ' clientes');
    } else {
      PROVINCIAS_OVERRIDE = {};
      console.log('Sin override de provincias todavía (esperando primera carga de Excel)');
    }
  } catch(e) { console.error('Error cargando override de provincias:', e.message); PROVINCIAS_OVERRIDE = {}; }
}

async function guardarProvinciasOverrideEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO provincias_override (datos, actualizado_at) VALUES ($1, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, actualizado_at = NOW()
    `, [json]);
    PROVINCIAS_OVERRIDE = data;
    PROVINCIAS_OVERRIDE_TS = new Date().toISOString();
    console.log('✓ Override de provincias guardado en PostgreSQL: ' + Object.keys(data).length + ' clientes');
  } catch(e) { console.error('Error guardando override de provincias:', e.message); }
}

// ─── CANTIDAD DE SKU POR MARCA (configurado manualmente por Fernando) ──────
// Estructura: { 'BIOSKIN': 39, 'ERAYBA': 23, 'ZIAJA': 39, 'ZIAJA PRO': N }.
// Guardado en el servidor (no localStorage) para que sea el mismo valor sin
// importar desde qué navegador/dispositivo se entre al dashboard.
let SKU_POR_MARCA = { 'BIOSKIN': 0, 'ERAYBA': 0, 'ZIAJA': 0, 'ZIAJA PRO': 0 };
let SKU_POR_MARCA_TS = null;

async function cargarSkuPorMarcaDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM sku_por_marca ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      SKU_POR_MARCA = { ...SKU_POR_MARCA, ...JSON.parse(r.rows[0].datos) };
      SKU_POR_MARCA_TS = new Date().toISOString();
      console.log('✓ SKU por marca cargado desde PostgreSQL:', SKU_POR_MARCA);
    }
  } catch(e) { console.error('Error cargando SKU por marca:', e.message); }
}

async function guardarSkuPorMarcaEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO sku_por_marca (datos, actualizado_at) VALUES ($1, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, actualizado_at = NOW()
    `, [json]);
    SKU_POR_MARCA = data;
    SKU_POR_MARCA_TS = new Date().toISOString();
    console.log('✓ SKU por marca guardado en PostgreSQL:', data);
  } catch(e) { console.error('Error guardando SKU por marca:', e.message); }
}

// ─── METAS DE INGRESO DE CLIENTES A MERCATELY (KPI manual, por asesora) ────
// MERCATELY_METAS: { 'Nombre completo asesora': metaMensual }. Editable en
// cualquier momento desde Configuración. Los registros mensuales reales (cuántos
// clientes entraron cada mes) viven en la tabla mercately_registros, separada,
// para mantener historial completo sin sobrescribir meses anteriores.
let MERCATELY_METAS = {};
let MERCATELY_METAS_TS = null;

async function cargarMercatelyMetasDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM mercately_metas ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      MERCATELY_METAS = JSON.parse(r.rows[0].datos);
      MERCATELY_METAS_TS = new Date().toISOString();
      console.log('✓ Metas Mercately cargadas desde PostgreSQL:', MERCATELY_METAS);
    } else {
      // Valores iniciales solicitados por Fernando, solo la primera vez (tabla vacía)
      MERCATELY_METAS = {
        'Giovanna Portilla': 0,
        'Liseth Gavilanes': 120,
        'Daniela Villegas Chamorro': 120,
        'María Caridad Zea Larrea': 120,
        'Karen Rebeca Mora Cedeño': 200,
        'Nicole Yanira Leon Marquez': 200
      };
      await guardarMercatelyMetasEnDB(MERCATELY_METAS);
      console.log('✓ Metas Mercately inicializadas con valores por defecto');
    }
  } catch(e) { console.error('Error cargando metas Mercately:', e.message); }
}

async function guardarMercatelyMetasEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO mercately_metas (datos, actualizado_at) VALUES ($1, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, actualizado_at = NOW()
    `, [json]);
    MERCATELY_METAS = data;
    MERCATELY_METAS_TS = new Date().toISOString();
    console.log('✓ Metas Mercately guardadas en PostgreSQL:', data);
  } catch(e) { console.error('Error guardando metas Mercately:', e.message); }
}

// ─── METAS DE CLIENTES NUEVOS EN CONTIFICO (KPI que reemplaza al "Clientes Nuevos"
// manual anterior). Se cuenta automáticamente al subir el Excel de Personas/Provincias
// — ver parsearExcelProvincias() y el endpoint /api/provincias/subir.
let CONTIFICO_CLIENTES_METAS = {};
let CONTIFICO_CLIENTES_METAS_TS = null;

async function cargarContificoClientesMetasDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM contifico_clientes_metas ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      CONTIFICO_CLIENTES_METAS = JSON.parse(r.rows[0].datos);
      CONTIFICO_CLIENTES_METAS_TS = new Date().toISOString();
      console.log('✓ Metas clientes Contifico cargadas desde PostgreSQL:', CONTIFICO_CLIENTES_METAS);
    } else {
      // Valores iniciales solicitados por Fernando, solo la primera vez (tabla vacía)
      CONTIFICO_CLIENTES_METAS = {
        'Giovanna Portilla': 0,
        'Liseth Gavilanes': 12,
        'Daniela Villegas Chamorro': 12,
        'María Caridad Zea Larrea': 8,
        'Karen Rebeca Mora Cedeño': 16,
        'Nicole Yanira Leon Marquez': 16
      };
      await guardarContificoClientesMetasEnDB(CONTIFICO_CLIENTES_METAS);
      console.log('✓ Metas clientes Contifico inicializadas con valores por defecto');
    }
  } catch(e) { console.error('Error cargando metas clientes Contifico:', e.message); }
}

async function guardarContificoClientesMetasEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO contifico_clientes_metas (datos, actualizado_at) VALUES ($1, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, actualizado_at = NOW()
    `, [json]);
    CONTIFICO_CLIENTES_METAS = data;
    CONTIFICO_CLIENTES_METAS_TS = new Date().toISOString();
    console.log('✓ Metas clientes Contifico guardadas en PostgreSQL:', data);
  } catch(e) { console.error('Error guardando metas clientes Contifico:', e.message); }
}

// Resuelve la provincia de un cliente con la prioridad: override por RUC/Cédula (Excel,
// subido manualmente por Fernando) > inferencia por palabras clave en la dirección
// (Contifico no expone un campo "provincia" directo en la API, solo dirección de texto).
function resolverProvinciaCliente(ruc, personaId, direccion){
  const rucLimpio = (ruc||'').trim();
  if(rucLimpio && PROVINCIAS_OVERRIDE[rucLimpio]) return PROVINCIAS_OVERRIDE[rucLimpio];
  return provinciaDesdeDir(direccion || '');
}

// Parsea el Excel de Personas/Clientes de Contifico (formato .xls o .xlsx), extrayendo
// RUC/Cédula + Provincia. Encabezados en la fila que contiene 'RUC' y 'Provincia'.
// Normaliza el texto de provincia que viene del Excel para que coincida exactamente
// con el nombre oficial usado en PROVINCIAS_NOMBRE (ej. "MANABI" sin tilde del Excel
// → "MANABÍ" con tilde, que es el estándar del resto del sistema). Si no encuentra
// coincidencia, devuelve el texto tal cual vino (mejor mostrar algo que perderlo).
function normalizarNombreProvincia(textoProvincia){
  if(!textoProvincia) return '';
  const sinTildes = s => (s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim();
  const buscado = sinTildes(textoProvincia);
  // 1) coincidencia exacta
  let match = PROVINCIAS_NOMBRE.find(p => sinTildes(p) === buscado);
  if(match) return match;
  // 2) una contiene a la otra (ej. Excel trae "SANTO DOMINGO DE LOS TSÁCHILAS",
  // el sistema usa solo "SANTO DOMINGO")
  match = PROVINCIAS_NOMBRE.find(p => { const pn=sinTildes(p); return buscado.includes(pn) || pn.includes(buscado); });
  return match || textoProvincia;
}

function parsearExcelProvincias(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let filaEncabezado = -1;
  for (let i = 0; i < Math.min(10, filas.length); i++) {
    const fila = filas[i] || [];
    if (fila.includes('RUC') && fila.includes('Provincia')) { filaEncabezado = i; break; }
  }
  if (filaEncabezado === -1) throw new Error('No se encontró la fila de encabezados (RUC/Provincia) en el Excel');

  const encabezados = filas[filaEncabezado];
  const idxRuc = encabezados.indexOf('RUC');
  const idxCedula = encabezados.indexOf('Cédula');
  const idxRazonSocial = encabezados.indexOf('Razón Social');
  const idxProvincia = encabezados.indexOf('Provincia');
  const idxVendedor = encabezados.indexOf('Vendedor Asignado');
  const idxEsCliente = encabezados.indexOf('Es Cliente');
  if (idxProvincia === -1) throw new Error('No se encontró la columna Provincia');

  const overrides = {};
  const clientesPorVendedor = {}; // { 'Nombre Vendedor': cantidadDeClientes }
  let filasConProvincia = 0, filasSinIdentificador = 0;
  for (let i = filaEncabezado + 1; i < filas.length; i++) {
    const fila = filas[i];
    if (!fila) continue;
    const razonSocial = idxRazonSocial !== -1 ? (fila[idxRazonSocial]||'').toString().trim() : '';
    if (!razonSocial) continue; // fila vacía o de otro tipo
    const ruc = idxRuc !== -1 ? (fila[idxRuc]||'').toString().trim() : '';
    const cedula = idxCedula !== -1 ? (fila[idxCedula]||'').toString().trim() : '';
    const identificador = ruc || cedula;
    const provinciaCruda = (fila[idxProvincia]||'').toString().trim().toUpperCase();
    const provincia = normalizarNombreProvincia(provinciaCruda);
    if (identificador && provincia) { overrides[identificador] = provincia; filasConProvincia++; }
    if (!identificador) filasSinIdentificador++;

    // Conteo de clientes por vendedor asignado, solo filas marcadas "Es Cliente" = Si/Sí
    // (si esa columna no existe en el Excel, se cuenta cualquier fila con vendedor asignado).
    const esCliente = idxEsCliente !== -1 ? (fila[idxEsCliente]||'').toString().trim().toUpperCase() : '';
    const cuentaComoCliente = idxEsCliente === -1 || esCliente === 'SI' || esCliente === 'SÍ' || esCliente === 'YES';
    if (cuentaComoCliente && idxVendedor !== -1) {
      const vendedor = (fila[idxVendedor]||'').toString().trim();
      if (vendedor && vendedor!=='N/A' && !vendedor.includes('Espíndola') && !vendedor.includes('Espindola')) {
        clientesPorVendedor[vendedor] = (clientesPorVendedor[vendedor]||0) + 1;
      }
    }
  }
  return { overrides, filasConProvincia, filasSinIdentificador, totalFilas: filas.length - filaEncabezado - 1, clientesPorVendedor };
}

// Parsea el Excel de testers (formato Cosétika: cada hoja = un cliente).
// Por cada hoja extrae el nombre real del cliente (fila CLIENTE si no es NUEVO),
// y por cada producto con fecha registrada genera un registro.
// Soporta hasta 3 columnas de fecha por producto (entregas repetidas).
function parsearExcelTesters(buffer, asesora) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const registros = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false, dateNF: 'yyyy-mm-dd' });

    let nombreReal = null;
    let datosInicio = null;

    for (let i = 0; i < filas.length; i++) {
      const fila = filas[i] || [];
      const celda0 = String(fila[0] || '').trim().toUpperCase();

      if (celda0 === 'CLIENTE') {
        const val = String(fila[1] || '').trim().replace(/^:/, '').trim();
        if (val.toUpperCase() !== 'NUEVO' && val !== '') {
          nombreReal = val;
        }
      }
      if (celda0 === 'CATEGORIA') {
        datosInicio = i + 1;
        break;
      }
    }

    if (datosInicio === null) continue;

    for (let i = datosInicio; i < filas.length; i++) {
      const fila = filas[i] || [];
      if (!fila[0]) continue;
      const categoria = String(fila[0] || '').trim();
      const producto  = String(fila[1] || '').trim();
      if (!categoria || !producto) continue;

      // Columnas 2, 3, 4 = hasta 3 fechas de entrega
      for (const colIdx of [2, 3, 4]) {
        const valFecha = fila[colIdx];
        if (!valFecha) continue;
        let fechaStr = null;
        if (typeof valFecha === 'string' && valFecha.match(/^\d{4}-\d{2}-\d{2}/)) {
          fechaStr = valFecha.substring(0, 10);
        } else if (typeof valFecha === 'string' && valFecha.trim()) {
          // Texto como "MAR" — guardamos null para la fecha pero sí registramos el tester
          fechaStr = null;
        }
        // Si hay algo en esa columna (incluye texto), registrar el tester
        registros.push({
          tab: sheetName.trim(),
          cliente: nombreReal || null,
          categoria,
          producto,
          fecha: fechaStr
        });
      }
    }
  }

  return registros;
}


async function cargarInventarioDesdeDB() {
  try {
    const r = await pool.query("SELECT datos, fecha_corte FROM inventario_data ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      INVENTARIO_CACHE = JSON.parse(r.rows[0].datos);
      INVENTARIO_CACHE_TS = new Date().toISOString();
      console.log('✓ Inventario cargado desde PostgreSQL: ' + Object.keys(INVENTARIO_CACHE.productos||{}).length + ' productos, corte ' + INVENTARIO_CACHE.fecha_corte);
    } else {
      INVENTARIO_CACHE = null;
      console.log('Sin inventario cargado todavía (esperando primera carga de Excel)');
    }
  } catch(e) { console.error('Error cargando inventario:', e.message); INVENTARIO_CACHE = null; }
}

async function guardarInventarioEnDB(data) {
  try {
    const json = JSON.stringify(data);
    await pool.query(`
      INSERT INTO inventario_data (datos, fecha_corte, actualizado_at) VALUES ($1, $2, NOW())
      ON CONFLICT (id_unico) DO UPDATE SET datos = $1, fecha_corte = $2, actualizado_at = NOW()
    `, [json, data.fecha_corte || null]);
    INVENTARIO_CACHE = data;
    INVENTARIO_CACHE_TS = new Date().toISOString();
    console.log('✓ Inventario guardado en PostgreSQL: ' + Object.keys(data.productos||{}).length + ' productos');
  } catch(e) { console.error('Error guardando inventario:', e.message); }
}

// ─── BACKUP SEMANAL DE BASE DE DATOS (descarga manual) ──────────────────────
// Exporta las tablas operativas que NO tienen respaldo en otro lugar (Contifico
// no las tiene): inventario, usuarios, visitas, planificación, zonas/provincias
// por asesora. ventas_data se incluye también por completitud, aunque en
// principio es recuperable regenerando desde Contifico si se perdiera.
// NOTA: el envío automático por correo (SMTP) no es posible en el plan actual
// de Railway, que bloquea todo tráfico SMTP saliente (puertos 25/465/587/2525)
// salvo en el plan Pro o superior. Por eso el backup es una descarga manual
// directa desde el dashboard, con un recordatorio visual si pasa de una semana.
// Lista de respaldo por si falla la consulta dinámica de tablas
const TABLAS_BACKUP = ['inventario_data', 'usuarios', 'visitas', 'planificacion', 'asesor_zonas', 'asesor_provincias', 'ventas_data'];

// ─── BACKUP AUTOMÁTICO A GOOGLE DRIVE ─────────────────────────────────────────────
// Railway no ofrece copias de la base en el plan actual y bloquea el SMTP saliente, así
// que la app se respalda sola: genera el volcado, lo comprime y lo sube a una carpeta de
// Drive. Se usa OAuth con refresh token — no caduca mientras no se revoque el permiso.
const GD_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GD_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const GD_REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN || '';
const GD_CARPETA = process.env.GOOGLE_DRIVE_CARPETA || 'Backups Cosétika App';
const GD_CONSERVAR = parseInt(process.env.GOOGLE_DRIVE_CONSERVAR || '12');   // cuántos backups guardar
let GD_TOKEN = { valor: null, expira: 0 };
let BACKUP_ESTADO = { ultimo: null, archivo: null, bytes: 0, error: null, subiendo: false };

function driveConfigurado(){ return !!(GD_CLIENT_ID && GD_CLIENT_SECRET && GD_REFRESH_TOKEN); }

async function driveToken(){
  if (GD_TOKEN.valor && Date.now() < GD_TOKEN.expira - 60000) return GD_TOKEN.valor;
  const body = new URLSearchParams({
    client_id: GD_CLIENT_ID, client_secret: GD_CLIENT_SECRET,
    refresh_token: GD_REFRESH_TOKEN, grant_type: 'refresh_token'
  });
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'}, body
  });
  const d = await r.json();
  if (!r.ok || !d.access_token) throw new Error('Google rechazó el refresh token: ' + (d.error_description || d.error || r.status));
  GD_TOKEN = { valor: d.access_token, expira: Date.now() + (d.expires_in||3600)*1000 };
  return GD_TOKEN.valor;
}

// Carpeta de destino: se crea una vez y su id queda guardado. Con el permiso drive.file
// la app solo ve lo que ella misma crea, así que no puede tocar el resto de tu Drive.
async function driveCarpeta(token){
  const guardada = await getConfigApp('drive_carpeta_id', null);
  if (guardada) {
    const chk = await fetch('https://www.googleapis.com/drive/v3/files/'+guardada+'?fields=id,trashed',
      { headers:{ Authorization:'Bearer '+token } });
    if (chk.ok) { const j = await chk.json(); if (!j.trashed) return guardada; }
  }
  const crear = await fetch('https://www.googleapis.com/drive/v3/files?fields=id', {
    method:'POST', headers:{ Authorization:'Bearer '+token, 'Content-Type':'application/json' },
    body: JSON.stringify({ name: GD_CARPETA, mimeType: 'application/vnd.google-apps.folder' })
  });
  const j = await crear.json();
  if (!crear.ok || !j.id) throw new Error('No se pudo crear la carpeta en Drive: ' + JSON.stringify(j).substring(0,200));
  await setConfigApp('drive_carpeta_id', j.id);
  console.log('✓ Carpeta de backups creada en Drive: ' + GD_CARPETA);
  return j.id;
}

async function subirBackupADrive(){
  if (!driveConfigurado()) { BACKUP_ESTADO.error = 'Faltan las credenciales de Google Drive'; return BACKUP_ESTADO; }
  if (BACKUP_ESTADO.subiendo) return BACKUP_ESTADO;
  BACKUP_ESTADO.subiendo = true;
  try {
    const token = await driveToken();
    const carpeta = await driveCarpeta(token);
    const backup = await generarBackupCompleto();
    const crudo = Buffer.from(JSON.stringify(backup), 'utf8');
    // Comprimir: el volcado en JSON pesa decenas de MB y en gzip baja alrededor de diez veces
    const comprimido = await new Promise((ok, err) => zlib.gzip(crudo, { level: 9 }, (e, b) => e ? err(e) : ok(b)));
    const fecha = fmtDateEC(nowEC()).split('/').reverse().join('-');
    const nombre = `backup_cosetika_${fecha}.json.gz`;

    const limite = '-----cosetika' + Date.now();
    const meta = JSON.stringify({ name: nombre, parents: [carpeta],
      description: 'Backup automático · ' + Object.keys(backup.tablas||{}).length + ' tablas · ' + new Date().toISOString() });
    const cuerpo = Buffer.concat([
      Buffer.from('--'+limite+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+meta+'\r\n'),
      Buffer.from('--'+limite+'\r\nContent-Type: application/gzip\r\n\r\n'),
      comprimido,
      Buffer.from('\r\n--'+limite+'--')
    ]);
    const up = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,size', {
      method:'POST',
      headers:{ Authorization:'Bearer '+token, 'Content-Type':'multipart/related; boundary='+limite },
      body: cuerpo
    });
    const res = await up.json();
    if (!up.ok || !res.id) throw new Error('Drive rechazó la subida: ' + JSON.stringify(res).substring(0,220));

    await registrarDescargaBackup();
    BACKUP_ESTADO = { ultimo: new Date().toISOString(), archivo: nombre, bytes: comprimido.length,
      crudo_bytes: crudo.length, error: null, subiendo: false, carpeta };
    console.log(`✓ Backup en Drive: ${nombre} (${(comprimido.length/1048576).toFixed(1)} MB comprimidos, ${(crudo.length/1048576).toFixed(1)} MB en crudo)`);
    await limpiarBackupsViejos(token, carpeta);
    return BACKUP_ESTADO;
  } catch(e) {
    BACKUP_ESTADO.error = e.message; BACKUP_ESTADO.subiendo = false;
    console.error('✗ Backup a Drive:', e.message);
    return BACKUP_ESTADO;
  }
}

// Deja solo los últimos N backups: sin esto la carpeta crece sin control
async function limpiarBackupsViejos(token, carpeta){
  try {
    const q = encodeURIComponent(`'${carpeta}' in parents and trashed=false`);
    const r = await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&orderBy=createdTime desc&fields=files(id,name,createdTime)&pageSize=100`,
      { headers:{ Authorization:'Bearer '+token } });
    const d = await r.json();
    const files = (d.files||[]);
    if (files.length <= GD_CONSERVAR) return;
    for (const f of files.slice(GD_CONSERVAR)) {
      await fetch('https://www.googleapis.com/drive/v3/files/'+f.id, { method:'DELETE', headers:{ Authorization:'Bearer '+token } });
      console.log('🗑️ Backup antiguo eliminado de Drive: ' + f.name);
    }
  } catch(e) { console.error('Limpieza de backups:', e.message); }
}

// ─── BACKUP A BACKBLAZE B2 ────────────────────────────────────────────────────────
// Alternativa a Drive sin OAuth: dos claves y listo. No hay pantalla de consentimiento
// ni permisos que caduquen, así que el respaldo no se rompe solo con el tiempo.
const B2_KEY_ID = process.env.B2_KEY_ID || '';
const B2_APP_KEY = process.env.B2_APP_KEY || '';
const B2_BUCKET = process.env.B2_BUCKET || '';           // nombre del bucket
const B2_CONSERVAR = parseInt(process.env.B2_CONSERVAR || '12');
function b2Configurado(){ return !!(B2_KEY_ID && B2_APP_KEY); }

async function b2Auth(){
  const basico = Buffer.from(B2_KEY_ID + ':' + B2_APP_KEY).toString('base64');
  const r = await fetch('https://api.backblazeb2.com/b2api/v3/b2_authorize_account', {
    headers: { Authorization: 'Basic ' + basico }
  });
  const d = await r.json();
  if (!r.ok || !d.authorizationToken) throw new Error('B2 rechazó las claves: ' + (d.message || r.status));
  const api = (d.apiInfo && d.apiInfo.storageApi) || d;
  return {
    token: d.authorizationToken,
    apiUrl: api.apiUrl,
    // Si la clave está restringida a un bucket, B2 ya lo indica y no hay que buscarlo
    bucketId: (api.bucketId || (d.allowed && d.allowed.bucketId)) || null,
    bucketName: (api.bucketName || (d.allowed && d.allowed.bucketName)) || null,
    accountId: d.accountId
  };
}

async function b2BucketId(a){
  if (a.bucketId) return a.bucketId;
  if (!B2_BUCKET) throw new Error('Falta B2_BUCKET (nombre del bucket) o una clave restringida a un bucket');
  const r = await fetch(a.apiUrl + '/b2api/v3/b2_list_buckets', {
    method:'POST', headers:{ Authorization: a.token, 'Content-Type':'application/json' },
    body: JSON.stringify({ accountId: a.accountId, bucketName: B2_BUCKET })
  });
  const d = await r.json();
  const b = (d.buckets || [])[0];
  if (!b) throw new Error('No se encontró el bucket "' + B2_BUCKET + '"');
  return b.bucketId;
}

async function subirBackupB2(){
  if (!b2Configurado()) { BACKUP_ESTADO.error = 'Faltan las claves de Backblaze'; return BACKUP_ESTADO; }
  if (BACKUP_ESTADO.subiendo) return BACKUP_ESTADO;
  BACKUP_ESTADO.subiendo = true;
  try {
    const a = await b2Auth();
    const bucketId = await b2BucketId(a);

    const backup = await generarBackupCompleto();
    const crudo = Buffer.from(JSON.stringify(backup), 'utf8');
    const comprimido = await new Promise((ok, err) => zlib.gzip(crudo, { level: 9 }, (e, b) => e ? err(e) : ok(b)));
    const fecha = fmtDateEC(nowEC()).split('/').reverse().join('-');
    const nombre = `backup_cosetika_${fecha}.json.gz`;
    const sha1 = crypto.createHash('sha1').update(comprimido).digest('hex');

    const ru = await fetch(a.apiUrl + '/b2api/v3/b2_get_upload_url', {
      method:'POST', headers:{ Authorization: a.token, 'Content-Type':'application/json' },
      body: JSON.stringify({ bucketId })
    });
    const du = await ru.json();
    if (!ru.ok || !du.uploadUrl) throw new Error('B2 no dio URL de subida: ' + (du.message || ru.status));

    const up = await fetch(du.uploadUrl, {
      method:'POST',
      headers:{
        Authorization: du.authorizationToken,
        'X-Bz-File-Name': encodeURIComponent(nombre),
        'Content-Type': 'application/gzip',
        'Content-Length': String(comprimido.length),
        'X-Bz-Content-Sha1': sha1
      },
      body: comprimido
    });
    const res = await up.json();
    if (!up.ok || !res.fileId) throw new Error('B2 rechazó el archivo: ' + (res.message || up.status));

    await registrarDescargaBackup();
    BACKUP_ESTADO = { destino:'Backblaze B2', ultimo: new Date().toISOString(), archivo: nombre,
      bytes: comprimido.length, crudo_bytes: crudo.length, error: null, subiendo: false };
    console.log(`✓ Backup en B2: ${nombre} (${(comprimido.length/1048576).toFixed(1)} MB de ${(crudo.length/1048576).toFixed(1)} MB)`);
    await limpiarBackupsB2(a, bucketId);
    return BACKUP_ESTADO;
  } catch(e) {
    BACKUP_ESTADO.error = e.message; BACKUP_ESTADO.subiendo = false;
    console.error('✗ Backup a B2:', e.message);
    return BACKUP_ESTADO;
  }
}

async function limpiarBackupsB2(a, bucketId){
  try {
    const r = await fetch(a.apiUrl + '/b2api/v3/b2_list_file_versions', {
      method:'POST', headers:{ Authorization: a.token, 'Content-Type':'application/json' },
      body: JSON.stringify({ bucketId, prefix: 'backup_cosetika_', maxFileCount: 200 })
    });
    const d = await r.json();
    const files = (d.files || []).sort((x,y) => String(y.fileName).localeCompare(String(x.fileName)));
    for (const f of files.slice(B2_CONSERVAR)) {
      await fetch(a.apiUrl + '/b2api/v3/b2_delete_file_version', {
        method:'POST', headers:{ Authorization: a.token, 'Content-Type':'application/json' },
        body: JSON.stringify({ fileName: f.fileName, fileId: f.fileId })
      });
      console.log('🗑️ Backup antiguo eliminado de B2: ' + f.fileName);
    }
  } catch(e) { console.error('Limpieza B2:', e.message); }
}

// Usa el destino que esté configurado: Backblaze si hay claves, si no Drive
async function respaldarAutomatico(){
  if (b2Configurado()) return subirBackupB2();
  if (driveConfigurado()) return subirBackupADrive();
  console.log('⚠️ Backup automático inactivo: no hay destino configurado');
  return { error: 'Sin destino configurado' };
}

// Respaldo semanal. No se ata a una hora exacta: si el servidor estaba reiniciándose
// justo a las 4 del domingo, esa ventana se perdería y habría que esperar otra semana.
// En vez de eso se comprueba cada hora si ya pasaron 7 días desde el último respaldo,
// leyendo la fecha de la BASE DE DATOS — así sobrevive a los redespliegues.
async function tocaRespaldar(){
  try {
    const r = await pool.query('SELECT ultima_descarga FROM backup_registro LIMIT 1');
    if (!r.rows.length || !r.rows[0].ultima_descarga) return true;   // nunca se hizo
    const dias = (Date.now() - new Date(r.rows[0].ultima_descarga).getTime()) / 86400000;
    return dias >= 7;
  } catch(e) { return false; }
}
setInterval(async () => {
  try {
    const h = nowEC();
    if (h.getHours() < 4) return;              // de madrugada en adelante, no a medianoche
    if (BACKUP_ESTADO.subiendo) return;
    if (!b2Configurado() && !driveConfigurado()) return;
    if (!(await tocaRespaldar())) return;
    console.log('🗄️ Respaldo semanal: han pasado 7 días o más desde el último');
    respaldarAutomatico().catch(e => console.error(e));
  } catch(e) { console.error('Programador de respaldo:', e.message); }
}, 60 * 60 * 1000);

// ─── RESTAURACIÓN DESDE UN BACKUP ─────────────────────────────────────────────────
// Devuelve la base al estado de un archivo de respaldo. Es una operación destructiva, así
// que está protegida por tres cosas: exige rol admin, exige escribir la palabra RESTAURAR,
// y saca un respaldo de seguridad ANTES de tocar nada, por si la restauración era un error.
//
// Reglas de compatibilidad, para que un backup viejo siga sirviendo aunque la app haya
// cambiado desde entonces:
//   · Una tabla del backup que ya no existe en la base se ignora.
//   · Una columna que ya no existe se ignora; una columna nueva queda con su valor por defecto.
//   · Los contadores (id automáticos) se reajustan al final para que no choquen los siguientes.
async function restaurarBackup(datos, opciones){
  opciones = opciones || {};
  const soloTablas = Array.isArray(opciones.tablas) && opciones.tablas.length ? opciones.tablas : null;
  const informe = { restauradas: [], ignoradas: [], errores: [], filas: 0 };

  const tablasBackup = Object.keys(datos.tablas || {});
  if (!tablasBackup.length) throw new Error('El archivo no contiene tablas');

  // Qué tablas y columnas existen HOY en la base
  const rT = await pool.query("SELECT tablename FROM pg_tables WHERE schemaname='public'");
  const existentes = new Set(rT.rows.map(x => x.tablename));
  const rC = await pool.query("SELECT table_name, column_name FROM information_schema.columns WHERE table_schema='public'");
  const columnasDe = {};
  rC.rows.forEach(x => { (columnasDe[x.table_name] = columnasDe[x.table_name] || new Set()).add(x.column_name); });

  const objetivo = tablasBackup.filter(t => {
    if (soloTablas && !soloTablas.includes(t)) return false;
    if (!existentes.has(t)) { informe.ignoradas.push({ tabla: t, motivo: 'ya no existe en la base' }); return false; }
    if (!Array.isArray(datos.tablas[t])) { informe.ignoradas.push({ tabla: t, motivo: 'el backup guardó un error para esta tabla' }); return false; }
    return true;
  });
  if (!objetivo.length) throw new Error('Ninguna tabla del archivo se puede restaurar');

  const cli = await pool.connect();
  try {
    await cli.query('BEGIN');
    // Vaciar primero TODAS las tablas objetivo: así no importa el orden ni las relaciones
    await cli.query('TRUNCATE ' + objetivo.map(t => '"'+t+'"').join(', ') + ' RESTART IDENTITY CASCADE');

    for (const tabla of objetivo) {
      const filas = datos.tablas[tabla];
      if (!filas.length) { informe.restauradas.push({ tabla, filas: 0 }); continue; }
      const cols = Object.keys(filas[0]).filter(c => columnasDe[tabla] && columnasDe[tabla].has(c));
      if (!cols.length) { informe.ignoradas.push({ tabla, motivo: 'ninguna columna coincide' }); continue; }

      const LOTE = 200;
      for (let i = 0; i < filas.length; i += LOTE) {
        const trozo = filas.slice(i, i + LOTE);
        const valores = [];
        const marcadores = trozo.map((f, fi) => '(' + cols.map((c, ci) => {
          let v = f[c];
          if (v && typeof v === 'object') v = JSON.stringify(v);   // json/jsonb y arrays
          valores.push(v === undefined ? null : v);
          return '$' + (fi*cols.length + ci + 1);
        }).join(',') + ')').join(',');
        await cli.query(
          'INSERT INTO "'+tabla+'" (' + cols.map(c => '"'+c+'"').join(',') + ') VALUES ' + marcadores,
          valores
        );
      }
      informe.restauradas.push({ tabla, filas: filas.length });
      informe.filas += filas.length;
    }

    // Reajustar los contadores automáticos para que los próximos ids no choquen
    for (const tabla of objetivo) {
      try {
        await cli.query(`SELECT setval(pg_get_serial_sequence('"${tabla}"','id'),
          GREATEST((SELECT COALESCE(MAX(id),0) FROM "${tabla}"), 1))
          WHERE pg_get_serial_sequence('"${tabla}"','id') IS NOT NULL`);
      } catch(e) { /* la tabla no tiene id serial */ }
    }

    await cli.query('COMMIT');
  } catch(e) {
    try { await cli.query('ROLLBACK'); } catch(e2) {}
    throw e;
  } finally { cli.release(); }

  // Recargar en memoria lo que la app mantiene cacheado
  try {
    const rd = await pool.query("SELECT datos FROM ventas_data WHERE id_unico='principal'");
    if (rd.rows.length) { DATA_CACHE = JSON.parse(rd.rows[0].datos); DATA_SERVIDA = { clave:'', ts:0, data:null }; }
  } catch(e) {}

  return informe;
}

async function generarBackupCompleto() {
  const backup = { generado_en: new Date().toISOString(), version: 2, tablas: {} };
  // Lista DINÁMICA: todas las tablas públicas de la BD — así el backup nunca se
  // queda desactualizado cuando se agregan módulos nuevos
  let tablas = [];
  try {
    const rt = await pool.query("SELECT tablename FROM pg_tables WHERE schemaname='public' ORDER BY tablename");
    tablas = rt.rows.map(x => x.tablename);
  } catch(e) { tablas = TABLAS_BACKUP; }
  for (const tabla of tablas) {
    try {
      let r;
      if (tabla === 'documentos') {
        // Los PDF (BYTEA) van en base64: restaurables y mucho más compactos que el
        // formato Buffer que produciría JSON.stringify sobre el binario crudo
        r = await pool.query("SELECT id, nombre, tamano, subido_por, created_at, encode(archivo,'base64') AS archivo_base64 FROM documentos");
      } else {
        r = await pool.query(`SELECT * FROM ${tabla}`);
      }
      backup.tablas[tabla] = r.rows;
    } catch(e) {
      backup.tablas[tabla] = { error: e.message };
    }
  }
  // Resumen de conteos para verificar el backup de un vistazo
  backup.resumen = {};
  Object.entries(backup.tablas).forEach(([t, rows]) => {
    backup.resumen[t] = Array.isArray(rows) ? rows.length : 'ERROR';
  });
  return backup;
}

async function registrarDescargaBackup() {
  try {
    await pool.query(`
      INSERT INTO backup_registro (ultima_descarga) VALUES (NOW())
      ON CONFLICT (id_unico) DO UPDATE SET ultima_descarga = NOW()
    `);
  } catch(e) { console.error('Error registrando fecha de backup:', e.message); }
}

async function obtenerEstadoBackup() {
  try {
    const r = await pool.query('SELECT ultima_descarga FROM backup_registro LIMIT 1');
    const ultima = r.rows.length > 0 ? r.rows[0].ultima_descarga : null;
    const diasDesde = ultima ? Math.floor((Date.now() - new Date(ultima).getTime()) / (1000*60*60*24)) : null;
    return {
      ultima_descarga: ultima,
      dias_desde_ultima_descarga: diasDesde,
      necesita_backup: diasDesde === null || diasDesde >= 7
    };
  } catch(e) {
    return { ultima_descarga: null, dias_desde_ultima_descarga: null, necesita_backup: true, error: e.message };
  }
}


async function cargarDataDesdeDB() {
  try {
    const r = await pool.query("SELECT datos FROM ventas_data ORDER BY actualizado_at DESC LIMIT 1");
    if (r.rows.length > 0) {
      DATA_CACHE = JSON.parse(r.rows[0].datos);
      DATA_CACHE_TS = new Date().toISOString();
      // Consolidar productos_mes al cargar para eliminar duplicados acumulados
      // (bug histórico donde fusionarAnioActualEnCache concatenaba sin consolidar)
      Object.values(DATA_CACHE).forEach(clientes => {
        (clientes||[]).forEach(cli => {
          if(cli.productos_mes && cli.productos_mes.length > 0) {
            cli.productos_mes = consolidarProductosMes(cli.productos_mes);
          }
        });
      });
      console.log(`✓ DATA cargada desde PostgreSQL: ` + Object.keys(DATA_CACHE).length + ` vendedoras`);
    // Auto-fix: asegurar que Giovanna y Ana tengan editar_planificacion
    pool.query(
      `UPDATE usuarios SET modulos = CASE
         WHEN modulos IS NULL OR modulos = '' THEN 'editar_planificacion'
         WHEN modulos NOT LIKE '%editar_planificacion%' THEN modulos || ',editar_planificacion'
         ELSE modulos
       END WHERE nombre ILIKE '%giovanna%' OR nombre ILIKE '%ana%'`
    ).catch(()=>{});
      // Disparar regeneración del año actual en background para limpiar datos desde Contifico
      setTimeout(() => regenerarDataAutomatico(), 5000);
    } else {
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

// ─── REGENERACIÓN AUTOMÁTICA DIARIA (madrugada, hora Ecuador) ───────────────
// Trae desde el 1 de enero del año EN CURSO hasta hoy (no años anteriores: esos ya están
// cerrados contablemente y no cambian, así que regenerarlos cada noche sería trabajo
// desperdiciado — solo agrega tiempo de ejecución y carga sobre Contifico sin beneficio).
// Si en algún momento se necesita corregir datos de años anteriores, usar el botón manual
// de Configuración indicando el rango de fechas que corresponda.
// ─── VENTAS POR DÍA POR MARCA: cálculo y caché en segundo plano ─────────────
global._vdmCache = global._vdmCache || {};
global._vdmCalculando = global._vdmCalculando || {};
async function calcularVentasDiaMarca(anio, mes) {
  const cacheKey = anio + '-' + mes;
  if (global._vdmCalculando[cacheKey]) return;
  global._vdmCalculando[cacheKey] = true;
  try {
    const mm = String(mes).padStart(2,'0');
    const ultimoDia = new Date(anio, mes, 0).getDate();
    const MARCAS_VDM = ['ZIAJA','BIOSKIN','ZIAJA PRO','ERAYBA'];
    const porMarca = {}; MARCAS_VDM.forEach(m => porMarca[m] = {});
    const vistos = new Set();
    let nextU = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=01/${mm}/${anio}&fecha_final=${String(ultimoDia).padStart(2,'0')}/${mm}/${anio}&page_size=100`;
    let pags = 0;
    while (nextU && pags < 60) {
      const resp = await fetch(nextU, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      if (!resp.ok) break;
      const data = await resp.json();
      (data.results || []).forEach(d => {
        if (d.tipo_registro !== 'CLI' || d.anulado) return;
        if (noEsVenta(d)) return;
        if (String(d.cliente?.ruc || d.cliente?.cedula || '').trim() === '1793143660001') return;
        const dk = d.id || d.documento;
        if (vistos.has(dk)) return; vistos.add(dk);
        const signo = esNotaCredito(d) ? -1 : 1;
        const dia = parseInt(String(d.fecha_emision||'').split('/')[0]) || 0;
        if (!dia) return;
        (d.detalles || []).forEach(det => {
          const info = catalogoProductos[det.producto_id];
          if (!info) return;
          let mN = String(info.marca||'').toUpperCase().trim();
          if (mN.replace(/\s+/g,'') === 'ZIAJAPRO') mN = 'ZIAJA PRO';
          if (!porMarca[mN]) return;
          const base = signo * parseFloat(det.base_gravable || det.base_cero || 0);
          porMarca[mN][dia] = (porMarca[mN][dia] || 0) + base;
        });
      });
      nextU = data.next || null;
      pags++;
    }
    const totales = {};
    MARCAS_VDM.forEach(m => {
      Object.keys(porMarca[m]).forEach(dv => { porMarca[m][dv] = Math.round(porMarca[m][dv]*100)/100; });
      totales[m] = Math.round(Object.values(porMarca[m]).reduce((a,b)=>a+b,0)*100)/100;
    });
    global._vdmCache[cacheKey] = { ts: Date.now(), data: { ok:true, anio, mes, marcas: porMarca, totales } };
    console.log(`✓ Ventas día×marca ${cacheKey} calculadas (${pags} páginas)`);
  } catch(e) { console.error('Error calculando ventas día×marca:', e.message); }
  global._vdmCalculando[cacheKey] = false;
}
// Mantener SIEMPRE caliente el mes actual: al arrancar y cada 10 minutos
setTimeout(() => { const h = nowEC(); calcularVentasDiaMarca(h.getFullYear(), h.getMonth()+1); }, 90 * 1000);
setInterval(() => { const h = nowEC(); calcularVentasDiaMarca(h.getFullYear(), h.getMonth()+1); }, 10 * 60 * 1000);

async function regenerarDataAutomatico() {
  if (regenerandoEnProceso) {
    console.log('Regeneración automática diaria pospuesta: otra regeneración en curso — reintento en 20 min');
    setTimeout(regenerarDataAutomatico, 20 * 60 * 1000);
    return;
  }
  regenerandoEnProceso = true;
  try {
    const hoy = nowEC();
    // RECONSTRUCCIÓN COMPLETA desde 2022: la antigua "fusión" del año actual sobre el caché
    // duplicaba datos sutilmente (el fantasma de +$725 de Liseth). Reconstruir todo desde
    // cero tarda unos minutos más pero es matemáticamente incapaz de duplicar.
    const fi = '01/01/2022';
    // Hasta AYER: lo de hoy vive en el caché en vivo y el frontend lo suma aparte
    const ayerR = new Date(hoy); ayerR.setDate(hoy.getDate()-1);
    const ff = fmtDateEC(ayerR);
    console.log(`⏰ Regeneración automática diaria (histórico completo): ${fi} al ${ff}`);
    const dataAnio = await generarDataJson(fi, ff);
    await guardarDataEnDB(dataAnio);
    // Marca hasta qué día llega el caché. Todo lo posterior lo completa /api/ventas-pendientes,
    // así ningún día puede quedar en tierra de nadie si la regeneración se atrasa o falla.
    try { await setConfigApp('data_hasta', ff); PEND_CACHE = { ts:0, clave:'', data:null }; } catch(e) {}
    try { fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(dataAnio, null, 2)); } catch(e) {}
    console.log('✓ Regeneración automática completada (histórico completo reconstruido)');
  } catch(e) { console.error('Error en regeneración automática:', e.message); }
  regenerandoEnProceso = false;
}
// Programar para correr a las 2:00 AM hora Ecuador (UTC-5) cada día
// SEGURO: si un deploy/reinicio se comió la regeneración de las 2 AM, el data.json
// amanece sin el día anterior. Al arrancar (y cada 2 horas) se verifica la frescura:
// si la última regeneración tiene más de 26 horas, se regenera automáticamente.
async function verificarFrescuraData(){
  try {
    if (regenerandoEnProceso) return;
    const r = await pool.query("SELECT actualizado_at FROM ventas_data WHERE id_unico='principal'");
    const ts = r.rows.length ? new Date(r.rows[0].actualizado_at).getTime() : 0;
    const horas = (Date.now() - ts) / 3600000;
    // Además del umbral de 26h: si HOY (Ecuador) aún no ha corrido ninguna regeneración
    // y ya pasaron las 2:30 AM, regenerar — garantiza que cada mañana esté el día anterior
    const fechaUltEC = ts ? new Date(ts).toLocaleDateString('en-CA', { timeZone: 'America/Guayaquil' }) : '';
    const ahoraEC = nowEC();
    const hoyEC = ahoraEC.toLocaleDateString('en-CA');
    const noCorrioHoy = fechaUltEC !== hoyEC && (ahoraEC.getHours() > 2 || (ahoraEC.getHours() === 2 && ahoraEC.getMinutes() >= 30));
    // Tercer chequeo, el más confiable: si el último día CON VENTAS del caché quedó a más
    // de 2 días de ayer, el caché está rezagado sin importar qué diga el timestamp.
    let rezagado = false;
    try {
      const ult = ultimoDiaEnCache();
      if (ult) {
        const ayer = new Date(ahoraEC); ayer.setDate(ayer.getDate() - 1);
        const diasAtras = Math.floor((ayer - ult) / 86400000);
        if (diasAtras > 2) { rezagado = true; console.log(`⚠️ Caché rezagado: último día con ventas ${fmtDateEC(ult)} (${diasAtras} días atrás)`); }
      }
    } catch(e) {}
    if (horas > 26 || noCorrioHoy || rezagado) {
      console.log(`⚠️ data.json desactualizado (última regen: ${fechaUltEC || 'nunca'} · hace ${Math.round(horas)}h) — regenerando ahora...`);
      regenerarDataAutomatico();
    }
  } catch(e) { console.error('Error verificando frescura de data:', e.message); }
}
setTimeout(verificarFrescuraData, 4 * 60 * 1000);       // al arrancar (tras cargar el catálogo)
setInterval(verificarFrescuraData, 60 * 60 * 1000); // y cada hora como respaldo

function programarRegeneracionDiaria() {
  const ahora = new Date();
  const proxima = new Date(ahora);
  proxima.setUTCHours(7, 0, 0, 0); // 2:00 AM Ecuador = 7:00 AM UTC
  if (proxima <= ahora) proxima.setUTCDate(proxima.getUTCDate() + 1);
  const msHastaProxima = proxima - ahora;
  setTimeout(() => {
    regenerarDataAutomatico();
    setInterval(regenerarDataAutomatico, 24 * 60 * 60 * 1000);
  }, msHastaProxima);
  console.log(`⏰ Próxima regeneración automática programada: ${proxima.toISOString()}`);
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
      CREATE TABLE IF NOT EXISTS inventario_data (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        fecha_corte DATE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS backup_registro (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        ultima_descarga TIMESTAMP
      );
      CREATE TABLE IF NOT EXISTS provincias_override (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS sku_por_marca (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS mercately_metas (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS mercately_registros (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        cantidad INTEGER NOT NULL DEFAULT 0, -- ACUMULADO TOTAL de clientes en Mercately a fin de ese mes (no el "++" mensual, que se calcula restando el acumulado del mes anterior)
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, anio, mes)
      );
      CREATE TABLE IF NOT EXISTS contifico_clientes_metas (
        id SERIAL PRIMARY KEY,
        id_unico VARCHAR(10) DEFAULT 'principal' UNIQUE,
        datos TEXT NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS contifico_clientes_registros (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        cantidad INTEGER NOT NULL DEFAULT 0, -- ACUMULADO TOTAL de clientes asignados en Contifico a fin de ese mes (contado automáticamente al subir el Excel de Personas)
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, anio, mes)
      );
      CREATE TABLE IF NOT EXISTS casa_abierta_registros (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        nombre_estetica VARCHAR(500) NOT NULL,
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, anio, mes)
      );
      CREATE TABLE IF NOT EXISTS visitas (
        id SERIAL PRIMARY KEY, lugar VARCHAR(255) NOT NULL,
        tipo VARCHAR(50) NOT NULL, asesora VARCHAR(255) NOT NULL,
        fecha TIMESTAMP DEFAULT NOW(), notas TEXT,
        inversion NUMERIC(10,2) DEFAULT 0
      );
      ALTER TABLE visitas ADD COLUMN IF NOT EXISTS inversion NUMERIC(10,2) DEFAULT 0;
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
        visitado_at VARCHAR(20),
        created_at TIMESTAMP DEFAULT NOW()
      );
      ALTER TABLE planificacion ADD COLUMN IF NOT EXISTS visitado_at VARCHAR(20);
      ALTER TABLE planificacion ADD COLUMN IF NOT EXISTS primera_compra_at VARCHAR(20);
      ALTER TABLE planificacion ADD COLUMN IF NOT EXISTS recompra_at VARCHAR(20);
      ALTER TABLE planificacion ADD COLUMN IF NOT EXISTS observaciones TEXT;
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS facturado_ayer BOOLEAN DEFAULT false;
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS fecha_control DATE;
      -- Migración única: pedidos movidos antes (cambiaban fecha) → restaurar fecha original
      UPDATE pedidos_web SET fecha_control = fecha, fecha = fecha - INTERVAL '1 day'
        WHERE facturado_ayer = true AND fecha_control IS NULL;
      CREATE TABLE IF NOT EXISTS capacitaciones (
        id SERIAL PRIMARY KEY,
        fecha DATE NOT NULL,
        ciudad VARCHAR(255),
        tema VARCHAR(500),
        direccion TEXT,
        horario VARCHAR(100),
        valor NUMERIC(10,2),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_cap_fecha ON capacitaciones(fecha);
      CREATE TABLE IF NOT EXISTS institutos (
        id SERIAL PRIMARY KEY,
        fecha DATE NOT NULL,
        nombre_instituto VARCHAR(500),
        actividad VARCHAR(100),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_inst_fecha ON institutos(fecha);
      CREATE TABLE IF NOT EXISTS giras (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255),
        fecha DATE NOT NULL,
        ciudad VARCHAR(255),
        valor_viaticos NUMERIC(10,2),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_giras_fecha ON giras(fecha);
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS valor_hotel NUMERIC(10,2) DEFAULT 0;
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS valor_mercately NUMERIC(10,2) DEFAULT 0;
      CREATE TABLE IF NOT EXISTS casas_abiertas (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255),
        fecha DATE NOT NULL,
        cliente VARCHAR(500),
        consignacion TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_casas_fecha ON casas_abiertas(fecha);
      CREATE TABLE IF NOT EXISTS asesor_zonas (
        id SERIAL PRIMARY KEY, asesora VARCHAR(255) NOT NULL,
        zona VARCHAR(255) NOT NULL, sector VARCHAR(255) NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, zona, sector)
      );
      CREATE TABLE IF NOT EXISTS asesor_provincias (
        id SERIAL PRIMARY KEY, asesora VARCHAR(255) NOT NULL,
        provincia VARCHAR(255) NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, provincia)
      );
      CREATE TABLE IF NOT EXISTS clientes_provincias (
        id SERIAL PRIMARY KEY,
        usuario_id INTEGER NOT NULL,
        provincia VARCHAR(100) NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(usuario_id, provincia)
      );
      CREATE TABLE IF NOT EXISTS envios_servientrega (
        id SERIAL PRIMARY KEY,
        guia VARCHAR(20) NOT NULL,
        fecha DATE NOT NULL,
        destinatario VARCHAR(500),
        razon_social VARCHAR(500),
        ciudad VARCHAR(255),
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(guia, fecha)
      );
      ALTER TABLE envios_servientrega ADD COLUMN IF NOT EXISTS email_destino VARCHAR(255);
      ALTER TABLE envios_servientrega ADD COLUMN IF NOT EXISTS email_enviado_at TIMESTAMP;
      ALTER TABLE envios_servientrega ADD COLUMN IF NOT EXISTS email_error TEXT;
      CREATE INDEX IF NOT EXISTS idx_envios_fecha ON envios_servientrega(fecha);
      CREATE INDEX IF NOT EXISTS idx_envios_destinatario ON envios_servientrega(LOWER(destinatario));
      CREATE INDEX IF NOT EXISTS idx_envios_razon_social ON envios_servientrega(LOWER(razon_social));
      CREATE TABLE IF NOT EXISTS facturas_detalle (
        id SERIAL PRIMARY KEY,
        documento_id VARCHAR(100) NOT NULL,
        fecha DATE NOT NULL,
        documento VARCHAR(100),
        cliente_nombre VARCHAR(500),
        vendedor_nombre VARCHAR(255),
        subtotal NUMERIC(12,2) DEFAULT 0,
        total NUMERIC(12,2) DEFAULT 0,
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(documento_id, fecha)
      );
      CREATE INDEX IF NOT EXISTS idx_facturas_fecha ON facturas_detalle(fecha);
      CREATE TABLE IF NOT EXISTS pedidos_web (
        id SERIAL PRIMARY KEY,
        numero_pedido VARCHAR(50) NOT NULL UNIQUE,
        fecha DATE NOT NULL,
        cliente_nombre VARCHAR(500),
        cedula_ruc VARCHAR(20),
        telefono VARCHAR(20),
        subtotal NUMERIC(12,2) DEFAULT 0,
        total NUMERIC(12,2) DEFAULT 0,
        productos TEXT,
        email_uid VARCHAR(100),
        facturado BOOLEAN DEFAULT false,
        documento_factura VARCHAR(100),
        html_crudo TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_pedidos_fecha ON pedidos_web(fecha);
      CREATE INDEX IF NOT EXISTS idx_pedidos_cedula ON pedidos_web(cedula_ruc);
      CREATE INDEX IF NOT EXISTS idx_pedidos_cliente ON pedidos_web(LOWER(cliente_nombre));
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS html_crudo TEXT;
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS email VARCHAR(200);
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS direccion VARCHAR(400);
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS ciudad VARCHAR(150);
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS provincia_env VARCHAR(150);
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS nota_cliente TEXT;
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS prefactura_doc VARCHAR(100);
      ALTER TABLE pedidos_web ADD COLUMN IF NOT EXISTS prefactura_at TIMESTAMP;
      CREATE TABLE IF NOT EXISTS metas_visitas (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL UNIQUE,
        meta INTEGER NOT NULL DEFAULT 30,
        updated_at TIMESTAMP DEFAULT NOW()
      );
      ALTER TABLE capacitaciones ADD COLUMN IF NOT EXISTS tipo VARCHAR(20);
      ALTER TABLE capacitaciones ADD COLUMN IF NOT EXISTS provincia VARCHAR(50);
      ALTER TABLE capacitaciones ADD COLUMN IF NOT EXISTS modalidad VARCHAR(20);
      ALTER TABLE capacitaciones ADD COLUMN IF NOT EXISTS realizada BOOLEAN DEFAULT false;
      ALTER TABLE institutos ADD COLUMN IF NOT EXISTS tipo_actividad VARCHAR(20);
      ALTER TABLE institutos ADD COLUMN IF NOT EXISTS realizada BOOLEAN DEFAULT false;
      ALTER TABLE institutos ADD COLUMN IF NOT EXISTS aperturado BOOLEAN DEFAULT false;
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS coordinada BOOLEAN DEFAULT false;
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS realizada BOOLEAN DEFAULT false;
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS ciudad_visita VARCHAR(100);
      ALTER TABLE giras ADD COLUMN IF NOT EXISTS fecha_fin DATE;
      ALTER TABLE casas_abiertas ADD COLUMN IF NOT EXISTS coordinada BOOLEAN DEFAULT false;
      ALTER TABLE casas_abiertas ADD COLUMN IF NOT EXISTS realizada BOOLEAN DEFAULT false;
      CREATE TABLE IF NOT EXISTS revisiones_lunes (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL,
        semana DATE NOT NULL,
        revisado BOOLEAN DEFAULT false,
        UNIQUE(asesora, semana)
      );
      CREATE TABLE IF NOT EXISTS kpi_metas (
        id SERIAL PRIMARY KEY,
        clave VARCHAR(50) NOT NULL UNIQUE,
        meta VARCHAR(100) NOT NULL,
        updated_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS equipos (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(100) NOT NULL UNIQUE,
        lider VARCHAR(255),
        activo BOOLEAN DEFAULT true,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS equipo_miembros (
        id SERIAL PRIMARY KEY,
        equipo_id INTEGER NOT NULL REFERENCES equipos(id) ON DELETE CASCADE,
        usuario_nombre VARCHAR(255) NOT NULL,
        UNIQUE(equipo_id, usuario_nombre)
      );
      CREATE TABLE IF NOT EXISTS push_subscriptions (
        id SERIAL PRIMARY KEY,
        usuario_id INTEGER,
        usuario_nombre VARCHAR(255),
        endpoint TEXT NOT NULL UNIQUE,
        p256dh TEXT,
        auth TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS testers (
        id SERIAL PRIMARY KEY,
        cliente_id VARCHAR(100) NOT NULL,
        cliente_nombre VARCHAR(500) NOT NULL,
        cliente_cedula VARCHAR(50),
        categoria VARCHAR(255),
        producto TEXT NOT NULL,
        codigo VARCHAR(50),
        fecha_entrega DATE,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_testers_cliente ON testers(cliente_id);
      CREATE INDEX IF NOT EXISTS idx_testers_nombre ON testers(LOWER(cliente_nombre));
      CREATE INDEX IF NOT EXISTS idx_testers_cedula ON testers(cliente_cedula);
      ALTER TABLE testers ADD COLUMN IF NOT EXISTS cliente_cedula VARCHAR(50);
      CREATE TABLE IF NOT EXISTS documentos (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(500) NOT NULL,
        tamano INTEGER,
        subido_por VARCHAR(255),
        archivo BYTEA NOT NULL,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS nsos (
        id SERIAL PRIMARY KEY,
        marca VARCHAR(100),
        nombre VARCHAR(500) NOT NULL,
        nso VARCHAR(100),
        created_at TIMESTAMP DEFAULT NOW()
      );
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS etiqueta TEXT;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS etiqueta_pdf BYTEA;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS etq_nombre VARCHAR(400);
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS etq_bytes INTEGER;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS etq_subido_at TIMESTAMP;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS certificado BYTEA;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS cert_nombre VARCHAR(400);
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS cert_bytes INTEGER;
      ALTER TABLE nsos ADD COLUMN IF NOT EXISTS cert_subido_at TIMESTAMP;
      CREATE TABLE IF NOT EXISTS referidos (
        id SERIAL PRIMARY KEY,
        cliente VARCHAR(500),
        referido VARCHAR(500),
        telefono VARCHAR(100),
        fecha DATE,
        message_id VARCHAR(300),
        contactado BOOLEAN DEFAULT false,
        contactado_at VARCHAR(30),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE UNIQUE INDEX IF NOT EXISTS idx_referidos_msgid ON referidos(message_id);
      ALTER TABLE referidos ADD COLUMN IF NOT EXISTS bono BOOLEAN DEFAULT false;
      ALTER TABLE referidos ADD COLUMN IF NOT EXISTS bono_at VARCHAR(30);
      ALTER TABLE referidos ADD COLUMN IF NOT EXISTS primera_compra BOOLEAN DEFAULT false;
      ALTER TABLE referidos ADD COLUMN IF NOT EXISTS primera_compra_at VARCHAR(30);
      CREATE TABLE IF NOT EXISTS personas (
        id SERIAL PRIMARY KEY,
        cedula VARCHAR(50),
        ruc VARCHAR(50),
        razon_social VARCHAR(500) NOT NULL,
        telefono VARCHAR(200),
        direccion TEXT,
        email VARCHAR(255),
        vendedor VARCHAR(255),
        updated_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_personas_cedula ON personas(cedula);
      CREATE INDEX IF NOT EXISTS idx_personas_ruc ON personas(ruc);
      CREATE INDEX IF NOT EXISTS idx_personas_nombre ON personas(LOWER(razon_social));
      ALTER TABLE personas ALTER COLUMN telefono TYPE VARCHAR(200);
      ALTER TABLE personas ALTER COLUMN cedula TYPE VARCHAR(50);
      ALTER TABLE personas ADD COLUMN IF NOT EXISTS instituto VARCHAR(200);
      ALTER TABLE personas ADD COLUMN IF NOT EXISTS origen VARCHAR(20);
      UPDATE personas SET origen='institutos' WHERE origen IS NULL AND instituto IS NOT NULL
        AND (vendedor IS NULL OR vendedor='') AND (telefono IS NULL OR telefono='') AND (email IS NULL OR email='');
      ALTER TABLE facturas_detalle ADD COLUMN IF NOT EXISTS cedula_ruc VARCHAR(50);
      CREATE TABLE IF NOT EXISTS caja_saldos (
        id SERIAL PRIMARY KEY,
        mes_key VARCHAR(7) NOT NULL,
        tipo VARCHAR(10) NOT NULL DEFAULT 'banco',   -- 'banco' | 'cobro'
        nombre VARCHAR(200) NOT NULL,
        monto NUMERIC(14,2) NOT NULL DEFAULT 0,
        orden INTEGER DEFAULT 0,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_caja_saldos_mes ON caja_saldos(mes_key);
      CREATE TABLE IF NOT EXISTS caja_pagos (
        id SERIAL PRIMARY KEY,
        mes_key VARCHAR(7) NOT NULL,
        dia INTEGER NOT NULL DEFAULT 1,
        concepto VARCHAR(300) NOT NULL,
        monto NUMERIC(14,2) NOT NULL DEFAULT 0,
        fuente VARCHAR(120),
        pagado BOOLEAN DEFAULT false,
        recurrente BOOLEAN DEFAULT false,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_caja_pagos_mes ON caja_pagos(mes_key);
      CREATE TABLE IF NOT EXISTS finanzas_reportes (
        id SERIAL PRIMARY KEY,
        tipo VARCHAR(20) NOT NULL,          -- 'pyg' | 'balance'
        anio INTEGER NOT NULL,
        datos TEXT NOT NULL,
        archivo VARCHAR(300),
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(tipo, anio)
      );
      CREATE TABLE IF NOT EXISTS producto_costos (
        codigo VARCHAR(100) PRIMARY KEY,
        nombre VARCHAR(500),
        costo NUMERIC(14,6) NOT NULL DEFAULT 0,
        fuente VARCHAR(120),
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS stock_bodegas (
        id SERIAL PRIMARY KEY,
        producto_id VARCHAR(30) NOT NULL,
        codigo VARCHAR(100),
        nombre VARCHAR(500),
        marca VARCHAR(100),
        bodega VARCHAR(200) NOT NULL,
        cantidad NUMERIC(13,2) DEFAULT 0,
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(producto_id, bodega)
      );
      CREATE TABLE IF NOT EXISTS kpis_control_notas (
        id SERIAL PRIMARY KEY,
        mes_key VARCHAR(7) NOT NULL,
        panel VARCHAR(30) NOT NULL,
        texto TEXT,
        actualizado_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(mes_key, panel)
      );
      CREATE TABLE IF NOT EXISTS viaticos_tarifas (
        id SERIAL PRIMARY KEY,
        provincia VARCHAR(100) NOT NULL,
        ciudad VARCHAR(200) NOT NULL,
        desayuno NUMERIC(10,2) DEFAULT 0,
        almuerzo NUMERIC(10,2) DEFAULT 0,
        cena NUMERIC(10,2) DEFAULT 0,
        hotel NUMERIC(10,2) DEFAULT 0,
        transporte NUMERIC(10,2) DEFAULT 0,
        taxi NUMERIC(10,2) DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW()
      );
      ALTER TABLE viaticos_tarifas ADD COLUMN IF NOT EXISTS dias NUMERIC(6,2) DEFAULT 0;
      ALTER TABLE viaticos_tarifas ADD COLUMN IF NOT EXISTS googlemaps NUMERIC(10,2) DEFAULT 0;
      CREATE TABLE IF NOT EXISTS nomina_detalle (
        id SERIAL PRIMARY KEY,
        mes_key VARCHAR(7) NOT NULL,
        cedula VARCHAR(20),
        empleado VARCHAR(300),
        cargo VARCHAR(200),
        dias NUMERIC(6,2) DEFAULT 0,
        sueldo NUMERIC(12,2) DEFAULT 0,
        horas_extra NUMERIC(12,2) DEFAULT 0,
        movilizacion NUMERIC(12,2) DEFAULT 0,
        comisiones NUMERIC(12,2) DEFAULT 0,
        bonificaciones NUMERIC(12,2) DEFAULT 0,
        decimo_tercero NUMERIC(12,2) DEFAULT 0,
        decimo_cuarto NUMERIC(12,2) DEFAULT 0,
        fondo_reserva NUMERIC(12,2) DEFAULT 0,
        total_ingresos NUMERIC(12,2) DEFAULT 0,
        iess_personal NUMERIC(12,2) DEFAULT 0,
        retencion NUMERIC(12,2) DEFAULT 0,
        anticipos NUMERIC(12,2) DEFAULT 0,
        otros_egresos NUMERIC(12,2) DEFAULT 0,
        total_egresos NUMERIC(12,2) DEFAULT 0,
        total_recibir NUMERIC(12,2) DEFAULT 0,
        aportes_patronales NUMERIC(12,2) DEFAULT 0,
        valor_ccc NUMERIC(12,2) DEFAULT 0,
        vacaciones NUMERIC(12,2) DEFAULT 0,
        prov_decimo_tercero NUMERIC(12,2) DEFAULT 0,
        prov_decimo_cuarto NUMERIC(12,2) DEFAULT 0,
        prov_fondo_reserva NUMERIC(12,2) DEFAULT 0,
        costo_empresa NUMERIC(12,2) DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(mes_key, cedula)
      );
      CREATE TABLE IF NOT EXISTS nomina_extras (
        id SERIAL PRIMARY KEY,
        mes_key VARCHAR(7) NOT NULL,
        cedula VARCHAR(20),
        empleado VARCHAR(300),
        concepto VARCHAR(120) NOT NULL,
        valor NUMERIC(12,2) DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(mes_key, cedula, concepto)
      );
      CREATE TABLE IF NOT EXISTS nomina_meses (
        mes_key VARCHAR(7) PRIMARY KEY,
        archivo VARCHAR(300),
        subido_por VARCHAR(255),
        subido_at TIMESTAMP DEFAULT NOW(),
        empleados INTEGER DEFAULT 0,
        costo_total NUMERIC(14,2) DEFAULT 0
      );
      CREATE TABLE IF NOT EXISTS visitas_excepciones (
        id SERIAL PRIMARY KEY,
        asesora VARCHAR(255) NOT NULL,
        semana DATE NOT NULL,
        motivo VARCHAR(100) DEFAULT 'Vacaciones',
        creado_por VARCHAR(255),
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(asesora, semana)
      );
      CREATE TABLE IF NOT EXISTS seguimiento_contactos (
        id SERIAL PRIMARY KEY,
        cliente_key VARCHAR(300) NOT NULL UNIQUE,
        cliente_nombre VARCHAR(500),
        asesora VARCHAR(255),
        contactado_at TIMESTAMPTZ DEFAULT NOW()
      );
      ALTER TABLE seguimiento_contactos ADD COLUMN IF NOT EXISTS comentario TEXT DEFAULT '';
      CREATE TABLE IF NOT EXISTS articulos (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(300) NOT NULL,
        categoria VARCHAR(100) DEFAULT '',
        activo BOOLEAN DEFAULT true,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS articulos_movimientos (
        id SERIAL PRIMARY KEY,
        articulo_id INTEGER NOT NULL,
        tipo VARCHAR(20) NOT NULL,
        cantidad NUMERIC(12,2) NOT NULL,
        nota VARCHAR(500) DEFAULT '',
        usuario VARCHAR(255) DEFAULT '',
        fecha TIMESTAMPTZ DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS googlemaps_historial (
        id SERIAL PRIMARY KEY,
        tarifa_id INTEGER NOT NULL,
        valor NUMERIC(10,2) DEFAULT 0,
        fecha TIMESTAMPTZ DEFAULT NOW()
      );
      INSERT INTO googlemaps_historial(tarifa_id, valor)
      SELECT v.id, v.googlemaps FROM viaticos_tarifas v
      WHERE v.googlemaps > 0 AND NOT EXISTS (SELECT 1 FROM googlemaps_historial h WHERE h.tarifa_id = v.id);
      CREATE TABLE IF NOT EXISTS testers_asesoras (
        id SERIAL PRIMARY KEY,
        producto_id VARCHAR(30) NOT NULL,
        codigo VARCHAR(100),
        nombre VARCHAR(500),
        marca VARCHAR(100),
        asesora VARCHAR(255) NOT NULL,
        entregado_at DATE NOT NULL DEFAULT CURRENT_DATE,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_testers_ase_prod ON testers_asesoras(producto_id, asesora);
      CREATE TABLE IF NOT EXISTS lotes (
        id SERIAL PRIMARY KEY,
        producto_id VARCHAR(30) NOT NULL,
        codigo VARCHAR(100),
        nombre VARCHAR(500),
        marca VARCHAR(100),
        lote VARCHAR(100) NOT NULL,
        fecha_caducidad DATE NOT NULL,
        cantidad NUMERIC(13,2) NOT NULL,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE INDEX IF NOT EXISTS idx_lotes_producto ON lotes(producto_id);
      CREATE TABLE IF NOT EXISTS stock_minimos (
        producto_id VARCHAR(30) PRIMARY KEY,
        minimo NUMERIC(13,2) DEFAULT 0,
        actualizado_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS clientes_reasignados (
        id SERIAL PRIMARY KEY,
        cliente_ruc VARCHAR(50) NOT NULL,
        cliente_nombre VARCHAR(500),
        vendedora_destino VARCHAR(255) NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(cliente_ruc)
      );
      CREATE TABLE IF NOT EXISTS app_config (
        clave VARCHAR(50) PRIMARY KEY,
        valor TEXT
      );
      ALTER TABLE personas ALTER COLUMN ruc TYPE VARCHAR(50);
    `);
    const usuarios = [
      { nombre: 'Fernando Espíndola', usuario: 'Fernando', password: '1234', rol: 'admin', modulos: 'ventas,visitas,kpis,inventario,config' },
      { nombre: 'Giovanna Portilla', usuario: 'Giovanna', password: '1234', rol: 'jefa_ventas', modulos: 'ventas,visitas,kpis,inventario,editar_visitas' },
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
initDB().then(() => cargarDataDesdeDB()).then(() => cargarInventarioDesdeDB()).then(() => cargarProvinciasOverrideDesdeDB()).then(() => cargarSkuPorMarcaDesdeDB()).then(() => cargarMercatelyMetasDesdeDB()).then(() => cargarContificoClientesMetasDesdeDB()).catch(e => console.error('Error init:', e.message));
programarRegeneracionDiaria();

const MIME = { '.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json','.png':'image/png','.jpg':'image/jpeg','.ico':'image/x-icon' };

function bodyJSON(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch(e) { reject(e); } });
    req.on('error', reject);
  });
}

function bodyBuffer(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

// Parser mínimo de multipart/form-data: extrae el primer archivo subido (campo 'file')
// como Buffer, usando el boundary del header Content-Type. No depende de librerías externas.
function parseMultipartFile(buffer, contentType) {
  const boundaryMatch = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || '');
  if (!boundaryMatch) return null;
  const boundary = '--' + (boundaryMatch[1] || boundaryMatch[2]).trim();
  const boundaryBuf = Buffer.from(boundary);
  const parts = [];
  let start = buffer.indexOf(boundaryBuf, 0);
  while (start !== -1) {
    const next = buffer.indexOf(boundaryBuf, start + boundaryBuf.length);
    if (next === -1) break;
    parts.push(buffer.slice(start + boundaryBuf.length, next));
    start = next;
  }
  for (const part of parts) {
    const headerEnd = part.indexOf('\r\n\r\n');
    if (headerEnd === -1) continue;
    const headerStr = part.slice(0, headerEnd).toString('utf8');
    if (!/name="file"/i.test(headerStr)) continue;
    // El contenido va desde después de los headers hasta antes del \r\n final de la parte
    let content = part.slice(headerEnd + 4);
    if (content.slice(-2).toString() === '\r\n') content = content.slice(0, -2);
    const filenameMatch = /filename="([^"]*)"/i.exec(headerStr);
    return { buffer: content, filename: filenameMatch ? filenameMatch[1] : 'archivo' };
  }
  return null;
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
      const token = firmarSesion(u);
      // Cookie firmada: el navegador la envía en cada petición, así el server puede
      // verificar el rol real sin depender del frontend (no se puede falsificar)
      res.writeHead(200,{
        'Content-Type':'application/json',
        'Set-Cookie': `cosetika_ses=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${60*60*24*30}`
      });
      res.end(JSON.stringify({ok:true, token, usuario:{id:u.id,nombre:u.nombre,usuario:u.usuario,rol:u.rol,modulos:u.modulos}}));
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
  // Renombrar usuaria propagando el nombre a todas las tablas (reemplazo de personal)
  if (/^\/api\/usuarios\/\d+\/renombrar$/.test(urlPath) && req.method === 'POST') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const { nuevo_nombre } = await bodyJSON(req);
      const nuevo = String(nuevo_nombre||'').trim().substring(0,250);
      if (!nuevo) throw new Error('El nombre nuevo está vacío');
      const rU = await pool.query('SELECT nombre FROM usuarios WHERE id=$1',[id]);
      if (!rU.rows.length) throw new Error('Usuario no encontrado');
      const viejo = rU.rows[0].nombre;
      if (viejo === nuevo) { res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, sin_cambios:true})); return; }
      await pool.query('UPDATE usuarios SET nombre=$1 WHERE id=$2',[nuevo,id]);
      const TABLAS_NOMBRE = [
        ['visitas','asesora'], ['planificacion','asesora'], ['giras','asesora'],
        ['casas_abiertas','asesora'], ['asesor_zonas','asesora'], ['asesor_provincias','asesora'],
        ['metas_visitas','asesora'], ['revisiones_lunes','asesora'], ['testers_asesoras','asesora'],
        ['mercately_registros','asesora'], ['contifico_clientes_registros','asesora'],
        ['casa_abierta_registros','asesora'], ['equipo_miembros','usuario_nombre'],
        ['push_subscriptions','usuario_nombre'], ['pedidos_web','asesora']
      ];
      const detalle = {};
      for (const [t, c] of TABLAS_NOMBRE) {
        try {
          const r = await pool.query(`UPDATE ${t} SET ${c}=$1 WHERE ${c}=$2`, [nuevo, viejo]);
          if (r.rowCount > 0) detalle[t] = r.rowCount;
        } catch(e) { console.log(`Renombrar: no se pudo actualizar ${t}: ${e.message}`); }
      }
      // Configs JSON con el nombre como clave
      try {
        const raw = await getConfigApp('meta_ventas', null);
        if (raw) {
          const cfg = JSON.parse(raw);
          if (cfg.metas && (viejo in cfg.metas)) { cfg.metas[nuevo] = cfg.metas[viejo]; delete cfg.metas[viejo]; await setConfigApp('meta_ventas', JSON.stringify(cfg)); detalle.meta_ventas = 1; }
        }
      } catch(e) {}
      try {
        const rM = await pool.query("SELECT datos FROM contifico_clientes_metas WHERE id_unico='principal'");
        if (rM.rows.length) {
          const d = JSON.parse(rM.rows[0].datos);
          if (d && typeof d === 'object' && (viejo in d)) { d[nuevo] = d[viejo]; delete d[viejo]; await pool.query("UPDATE contifico_clientes_metas SET datos=$1, actualizado_at=NOW() WHERE id_unico='principal'", [JSON.stringify(d)]); detalle.metas_clientes = 1; }
        }
      } catch(e) {}
      // Presupuesto de ventas: renombrar claves (pct, overrides y snapshots)
      try {
        const rawP = await getConfigApp('presupuesto_ventas', null);
        if (rawP) {
          const cfgP = JSON.parse(rawP);
          let cambioP = false;
          ['pct','overrides'].forEach(sec => {
            if (cfgP[sec] && (viejo in cfgP[sec])) { cfgP[sec][nuevo] = cfgP[sec][viejo]; delete cfgP[sec][viejo]; cambioP = true; }
          });
          Object.values(cfgP.snapshots || {}).forEach(sn => {
            if (sn && (viejo in sn)) { sn[nuevo] = sn[viejo]; delete sn[viejo]; cambioP = true; }
          });
          if (cambioP) { await setConfigApp('presupuesto_ventas', JSON.stringify(cfgP)); detalle.presupuesto = 1; }
        }
      } catch(e) {}
      console.log(`👤 Usuaria renombrada: "${viejo}" → "${nuevo}"`, detalle);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, viejo, nuevo, detalle}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
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
      const id = urlPath.split('/')[3]; // /api/usuarios/{id} — evita capturar query params con .pop()
      const solicitanteId = urlObj.searchParams.get('solicitante');
      if (solicitanteId && String(solicitanteId) === String(id)) {
        res.writeHead(400,{'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok:false, error: 'No puedes eliminar tu propia cuenta' }));
        return;
      }
      const r = await pool.query("DELETE FROM usuarios WHERE id=$1", [id]);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, eliminado: r.rowCount>0 }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // VISITAS
// ─── PUSH NOTIFICATIONS ──────────────────────────────────────────────────────

  // GET /api/push/vapid-key → clave pública VAPID para el cliente
  if (urlPath === '/api/push/vapid-key' && req.method === 'GET') {
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ publicKey: VAPID_PUBLIC_KEY }));
    return;
  }
  // GET /api/push/status → diagnóstico de suscripciones
  if (urlPath === '/api/push/status' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT id, usuario_nombre, created_at FROM push_subscriptions');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, subs: r.rows, vapid: !!VAPID_PUBLIC_KEY, webpush: !!webpush }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/push/test → enviar push de prueba
  if (urlPath === '/api/push/test' && req.method === 'POST') {
    await enviarPushATodos({ title: '🔔 Cosétika — Prueba', body: 'Las notificaciones funcionan correctamente', tag: 'test' });
    res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    return;
  }
  // Patch: agregar editar_planificacion a Giovanna y Ana si no lo tienen
  if (urlPath === '/api/fix-planificacion-perms' && req.method === 'POST') {
    try {
      const r = await pool.query(
        `UPDATE usuarios SET modulos = CASE
           WHEN modulos IS NULL OR modulos = '' THEN 'editar_planificacion'
           WHEN modulos NOT LIKE '%editar_planificacion%' THEN modulos || ',editar_planificacion'
           ELSE modulos
         END
         WHERE nombre ILIKE '%giovanna%' OR nombre ILIKE '%ana%'
         RETURNING nombre, modulos`
      );
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:true, updated: r.rows}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/push/subscribe → registrar suscripción
  if (urlPath === '/api/push/subscribe' && req.method === 'POST') {
    try {
      const { subscription, usuarioNombre } = await bodyJSON(req);
      await pool.query(
        `INSERT INTO push_subscriptions(endpoint, p256dh, auth, usuario_nombre)
         VALUES($1,$2,$3,$4)
         ON CONFLICT(endpoint) DO UPDATE SET p256dh=$2, auth=$3, usuario_nombre=$4`,
        [subscription.endpoint, subscription.keys.p256dh, subscription.keys.auth, usuarioNombre||'']
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // DELETE /api/push/unsubscribe → eliminar suscripción
  if (urlPath === '/api/push/unsubscribe' && req.method === 'POST') {
    try {
      const { endpoint } = await bodyJSON(req);
      await pool.query('DELETE FROM push_subscriptions WHERE endpoint=$1', [endpoint]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // CAPACITACIONES
  if (urlPath === '/api/capacitaciones' && req.method === 'GET') {
    try {
      const mes = urlObj.searchParams.get('mes');
      let r;
      if (mes) {
        r = await pool.query(
          `SELECT id, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha, ciudad, tema, direccion, horario, valor, tipo, provincia, modalidad, realizada FROM capacitaciones WHERE TO_CHAR(fecha,'YYYY-MM')=$1 ORDER BY fecha ASC`,
          [mes]
        );
      } else {
        r = await pool.query(`SELECT id, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha, ciudad, tema, direccion, horario, valor, tipo, provincia, modalidad, realizada FROM capacitaciones ORDER BY fecha DESC LIMIT 100`);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e){
      console.error('Error GET capacitaciones:', e.message);
      res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message}));
    }
    return;
  }
  if (urlPath === '/api/capacitaciones' && req.method === 'POST') {
    try {
      const { fecha, ciudad, tema, direccion, horario, valor, tipo, provincia, modalidad } = await bodyJSON(req);
      const r = await pool.query(
        `INSERT INTO capacitaciones(fecha,ciudad,tema,direccion,horario,valor,tipo,provincia,modalidad) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING *`,
        [fecha, ciudad||null, tema||null, direccion||null, horario||null, valor||null, tipo||null, provincia||null, modalidad||null]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, row:r.rows[0]}));
    } catch(e){
      console.error('Error INSERT capacitacion:', e.message);
      res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message}));
    }
    return;
  }
  if (urlPath.match(/^\/api\/capacitaciones\/\d+$/) && req.method === 'PUT') {
    try {
      const id = urlPath.split('/').pop();
      const body = await bodyJSON(req);
      // UPDATE dinámico: solo las columnas presentes en el body (permite marcar
      // "realizada" sin borrar el resto de campos)
      const permitidas = ['fecha','ciudad','tema','direccion','horario','valor','tipo','provincia','modalidad','realizada'];
      const cols = Object.keys(body).filter(k => permitidas.includes(k));
      if(cols.length > 0){
        const sets = cols.map((k,i)=>`${k}=$${i+1}`).join(',');
        const vals = [...cols.map(k=>body[k]), id];
        await pool.query(`UPDATE capacitaciones SET ${sets} WHERE id=$${cols.length+1}`, vals);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath.match(/^\/api\/capacitaciones\/\d+$/) && req.method === 'DELETE') {
    try {
      const id = urlPath.split('/').pop();
      await pool.query('DELETE FROM capacitaciones WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ── INSTITUTOS / GIRAS / CASAS ABIERTAS — CRUD explícito ─────────────
  for(const tabla of ['institutos','giras','casas_abiertas']){
    if(urlPath !== `/api/${tabla}` && !urlPath.match(new RegExp(`^/api/${tabla}/\\d+$`))) continue;

    if(req.method === 'GET'){
      try{
        const mes = urlObj.searchParams.get('mes');
        const anio = urlObj.searchParams.get('anio');
        let r;
        if(mes) r = await pool.query(`SELECT *, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha_str FROM ${tabla} WHERE TO_CHAR(fecha,'YYYY-MM')=$1 ORDER BY fecha ASC`,[mes]);
        else if(anio) r = await pool.query(`SELECT *, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha_str FROM ${tabla} WHERE TO_CHAR(fecha,'YYYY')=$1 ORDER BY fecha ASC`,[anio]);
        else r = await pool.query(`SELECT *, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha_str FROM ${tabla} ORDER BY fecha DESC LIMIT 100`);
        // Normalizar: usar fecha_str como fecha; fecha_fin (giras multi-día) a YYYY-MM-DD
        const rows = r.rows.map(row => {
          const out = {...row, fecha: row.fecha_str||row.fecha};
          if (out.fecha_fin) out.fecha_fin = new Date(out.fecha_fin).toISOString().substring(0,10);
          return out;
        });
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(rows));
      }catch(e){
        console.error(`Error GET ${tabla}:`, e.message);
        res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({error:e.message}));
      }
      return;
    }
    if(req.method === 'POST' && urlPath === `/api/${tabla}`){
      try{
        const body = await bodyJSON(req);
        const cols = Object.keys(body).filter(k=>body[k]!==undefined && body[k]!==null);
        const vals = cols.map(k=>body[k]);
        const placeholders = cols.map((_,i)=>`$${i+1}`).join(',');
        const r = await pool.query(
          `INSERT INTO ${tabla}(${cols.join(',')}) VALUES(${placeholders}) RETURNING *`,vals
        );
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true,row:r.rows[0]}));
      }catch(e){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({ok:false,error:e.message}));}
      return;
    }
    const matchId = urlPath.match(new RegExp(`^/api/${tabla}/(\\d+)$`));
    if(matchId && req.method === 'PUT'){
      try{
        const id = matchId[1];
        const body = await bodyJSON(req);
        const cols = Object.keys(body).filter(k=>body[k]!==undefined);
        const sets = cols.map((k,i)=>`${k}=$${i+1}`).join(',');
        const vals = [...cols.map(k=>body[k]), id];
        await pool.query(`UPDATE ${tabla} SET ${sets} WHERE id=$${cols.length+1}`,vals);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
      }catch(e){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({ok:false,error:e.message}));}
      return;
    }
    if(matchId && req.method === 'DELETE'){
      try{
        await pool.query(`DELETE FROM ${tabla} WHERE id=$1`,[matchId[1]]);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
      }catch(e){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({ok:false,error:e.message}));}
      return;
    }
  }

  if (urlPath === '/api/visitas' && req.method === 'GET') {
    try { const r = await pool.query('SELECT * FROM visitas ORDER BY fecha DESC LIMIT 300'); res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows)); }
    catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // GET /api/inversiones?anio=2026&mes=6  → viáticos por visita a provincia (tipo='provincia')
  if (urlPath === '/api/inversiones' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || nowEC().getFullYear();
      const mes  = parseInt(urlObj.searchParams.get('mes'))  || (nowEC().getMonth()+1);
      const r = await pool.query(
        `SELECT id, lugar, asesora,
                COALESCE(inversion,0) AS inversion,
                notas,
                fecha::date AS fecha
         FROM visitas
         WHERE tipo='provincia'
           AND EXTRACT(YEAR  FROM fecha AT TIME ZONE 'America/Guayaquil') = $1
           AND EXTRACT(MONTH FROM fecha AT TIME ZONE 'America/Guayaquil') = $2
           AND COALESCE(inversion,0) > 0
         ORDER BY fecha DESC, asesora`,
        [anio, mes]
      );
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // PUT /api/inversiones/:id  → editar inversión de una visita ya registrada
  if (urlPath.startsWith('/api/inversiones/') && req.method === 'PUT') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const { inversion, notas } = await bodyJSON(req);
      await pool.query('UPDATE visitas SET inversion=$1, notas=$2 WHERE id=$3', [parseFloat(inversion)||0, notas||null, id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // PUT /api/visitas/:id  → editar lugar, fecha e inversión de una visita
  if (urlPath.startsWith('/api/visitas/') && req.method === 'PUT') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const { lugar, fecha, inversion } = await bodyJSON(req);
      await pool.query(
        'UPDATE visitas SET lugar=$1, fecha=$2, inversion=$3 WHERE id=$4',
        [lugar, fecha, parseFloat(inversion)||0, id]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/visitas' && req.method === 'POST') {
    try {
      const {lugar,tipo,asesora,notas,inversion} = await bodyJSON(req);
      const r = await pool.query('INSERT INTO visitas(lugar,tipo,asesora,notas,inversion) VALUES($1,$2,$3,$4,$5) RETURNING *',[lugar,tipo,asesora,notas||null,parseFloat(inversion)||0]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows[0]));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ENVÍOS SERVIENTREGA — manifiestos diarios de guías
  // ─── AVISO DE GUÍA POR CORREO ─────────────────────────────────────────────────
  // Railway bloquea el SMTP saliente en este plan, así que el envío va por API HTTPS.
  const MAIL_API_KEY = process.env.RESEND_API_KEY || '';
  const MAIL_REMITENTE = process.env.EMAIL_REMITENTE || '';        // ej. envios@cosetika.com
  const MAIL_NOMBRE = process.env.EMAIL_NOMBRE || 'Cosétika';
  const MAIL_COPIA = process.env.EMAIL_COPIA || '';                 // opcional, copia interna
  function correoConfigurado(){ return !!(MAIL_API_KEY && MAIL_REMITENTE); }

  const normNom = x => String(x||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase().replace(/[^A-Z0-9 ]/g,' ').replace(/\s+/g,' ').trim();

  // Busca el correo de la persona a la que va la guía.
  //
  // El manifiesto de Servientrega NO trae cédula, solo el nombre del destinatario — y lo
  // escribe completo y en otro orden del que usa la clienta al comprar ("APELLIDO APELLIDO
  // NOMBRE" contra "Nombre Apellido"). Por eso la coincidencia exacta fallaba en la mitad
  // de los casos.
  //
  // Estrategia: comparar por PALABRAS. Se prioriza los pedidos web recientes, donde hay
  // pocos candidatos (los de las últimas semanas) y por tanto dos apellidos coincidentes
  // ya son prueba suficiente. Después el directorio completo de Contifico, donde al haber
  // miles de registros se exige más coincidencia para no equivocarse de persona.
  const _normN = x => String(x||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase().replace(/[^A-Z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
  const _tokens = x => _normN(x).split(' ').filter(t => t.length >= 3);

  let _cacheCand = null, _cacheCandTs = 0;
  async function candidatosCorreo(){
    if (_cacheCand && (Date.now() - _cacheCandTs) < 60000) return _cacheCand;
    const rp = await pool.query(
      `SELECT cliente_nombre AS nombre, email, cedula_ruc FROM pedidos_web
       WHERE email IS NOT NULL AND email <> '' AND fecha >= (CURRENT_DATE - 30)
       ORDER BY id DESC`);
    const rc = await pool.query(
      `SELECT razon_social AS nombre, email, COALESCE(cedula, ruc) AS cedula_ruc FROM personas
       WHERE email IS NOT NULL AND email <> ''`);
    _cacheCand = {
      pedidos: rp.rows.map(x => ({ ...x, toks: _tokens(x.nombre) })),
      personas: rc.rows.map(x => ({ ...x, toks: _tokens(x.nombre) }))
    };
    _cacheCandTs = Date.now();
    return _cacheCand;
  }

  function _mejorPorTokens(lista, toks, minCoincidencias){
    if (!toks.length) return null;
    let mejor = null, mejorPuntaje = 0, empatados = 0;
    for (const c of lista) {
      if (!c.toks.length) continue;
      let p = 0;
      for (const t of toks) if (c.toks.includes(t)) p++;
      if (p > mejorPuntaje) { mejorPuntaje = p; mejor = c; empatados = 1; }
      else if (p === mejorPuntaje && p > 0) empatados++;
    }
    // Si dos personas distintas empatan con el mismo puntaje, no se arriesga: mejor pedir
    // el correo a mano que mandarle la guía a quien no es.
    if (!mejor || mejorPuntaje < minCoincidencias || empatados > 1) return null;
    return { registro: mejor, puntaje: mejorPuntaje };
  }

  async function buscarCorreoDe(destinatario, razonSocial){
    const cand = await candidatosCorreo();
    const nombres = [destinatario, razonSocial].filter(Boolean);
    for (const nom of nombres) {
      const toks = _tokens(nom);
      if (!toks.length) continue;

      // 1) Pedidos web recientes: pocos candidatos, basta con 2 palabras coincidentes
      const h1 = _mejorPorTokens(cand.pedidos, toks, 2);
      if (h1) return { email: h1.registro.email, fuente: 'Pedido web',
        coincide: h1.registro.nombre, cedula: h1.registro.cedula_ruc || null };

      // 2) Directorio de Contifico: exacto primero
      const nEx = _normN(nom);
      const ex = cand.personas.find(p => _normN(p.nombre) === nEx);
      if (ex) return { email: ex.email, fuente: 'Contifico', coincide: ex.nombre, cedula: ex.cedula_ruc || null };

      // 3) Directorio de Contifico por palabras: al ser miles, se exige más evidencia
      const minC = toks.length >= 4 ? 3 : 2;
      const h3 = _mejorPorTokens(cand.personas, toks, minC);
      if (h3) return { email: h3.registro.email, fuente: 'Contifico (por nombre)',
        coincide: h3.registro.nombre, cedula: h3.registro.cedula_ruc || null };
    }
    return null;
  }

  // Previsualización: qué se enviaría y a quién. NO envía nada.
  if (urlPath === '/api/envios/preparar-correos' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const fecha = urlObj.searchParams.get('fecha');
      if (!fecha) throw new Error('Falta la fecha');
      const r = await pool.query(
        `SELECT guia, destinatario, razon_social, ciudad, email_destino, email_enviado_at
         FROM envios_servientrega WHERE fecha=$1 ORDER BY id ASC`, [fecha]);
      const lista = [];
      for (const e of r.rows) {
        let email = e.email_destino || null, fuente = e.email_destino ? 'guardado' : null, coincide = null;
        let cedula = null;
        if (!email) {
          const hit = await buscarCorreoDe(e.destinatario, e.razon_social);
          if (hit) { email = hit.email; fuente = hit.fuente; coincide = hit.coincide; cedula = hit.cedula; }
        }
        lista.push({ guia: e.guia, destinatario: e.destinatario, razon_social: e.razon_social,
          ciudad: e.ciudad, email, fuente, coincide, cedula, ya_enviado: !!e.email_enviado_at,
          enviado_at: e.email_enviado_at });
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, configurado: correoConfigurado(), remitente: MAIL_REMITENTE, fecha, envios: lista }));
    } catch(e) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // Envío real. Recibe la lista ya revisada por la persona.
  if (urlPath === '/api/envios/enviar-correos' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      if (!correoConfigurado()) throw new Error('Faltan RESEND_API_KEY o EMAIL_REMITENTE en las variables de Railway');
      const { fecha, envios } = await bodyJSON(req);
      if (!fecha || !Array.isArray(envios)) throw new Error('Faltan datos');
      const resultados = [];
      for (const e of envios) {
        const email = String(e.email||'').trim();
        if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
          resultados.push({ guia: e.guia, ok:false, error:'Correo inválido o vacío' });
          continue;
        }
        const nombre = String(e.destinatario || e.razon_social || 'Hola').split(' ').slice(0,2).join(' ');
        const html = plantillaCorreoGuia(nombre, e.guia, e.ciudad);
        try {
          const r = await fetch('https://api.resend.com/emails', {
            method:'POST',
            headers:{ Authorization: 'Bearer ' + MAIL_API_KEY, 'Content-Type':'application/json' },
            body: JSON.stringify({
              from: MAIL_NOMBRE + ' <' + MAIL_REMITENTE + '>',
              to: [email],
              ...(MAIL_COPIA ? { bcc: [MAIL_COPIA] } : {}),
              subject: 'Tu pedido está en camino · Guía ' + e.guia,
              html
            })
          });
          const d = await r.json();
          if (!r.ok) throw new Error(d.message || ('HTTP ' + r.status));
          await pool.query(
            'UPDATE envios_servientrega SET email_destino=$1, email_enviado_at=NOW(), email_error=NULL WHERE guia=$2 AND fecha=$3',
            [email, e.guia, fecha]);
          resultados.push({ guia: e.guia, ok:true, email });
        } catch(err) {
          await pool.query('UPDATE envios_servientrega SET email_destino=$1, email_error=$2 WHERE guia=$3 AND fecha=$4',
            [email, String(err.message).substring(0,400), e.guia, fecha]);
          resultados.push({ guia: e.guia, ok:false, email, error: err.message });
        }
      }
      const bien = resultados.filter(x=>x.ok).length;
      console.log(`✉️ Guías enviadas por correo: ${bien}/${resultados.length} (${fecha})`);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, enviados: bien, total: resultados.length, resultados }));
    } catch(e) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // GET /api/envios?fecha=YYYY-MM-DD  → envíos de un día específico
  if (urlPath === '/api/envios' && req.method === 'GET') {
    try {
      const fecha = urlObj.searchParams.get('fecha');
      if(fecha){
        const r = await pool.query('SELECT guia, fecha, destinatario, razon_social, ciudad FROM envios_servientrega WHERE fecha=$1 ORDER BY id ASC', [fecha]);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
      } else {
        // Sin fecha: devolver lista de fechas disponibles con conteo
        const r = await pool.query('SELECT fecha, COUNT(*) as total FROM envios_servientrega GROUP BY fecha ORDER BY fecha ASC');
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
      }
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // GET /api/envios/buscar?q=nombre → busca en destinatario y razon_social en TODAS las fechas
  if (urlPath === '/api/envios/buscar' && req.method === 'GET') {
    try {
      const q = (urlObj.searchParams.get('q')||'').trim();
      if(!q || q.length < 2){
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify([]));
        return;
      }
      const r = await pool.query(
        `SELECT guia, fecha, destinatario, razon_social, ciudad FROM envios_servientrega
         WHERE LOWER(destinatario) LIKE LOWER($1) OR LOWER(razon_social) LIKE LOWER($1)
         ORDER BY fecha DESC LIMIT 200`,
        [`%${q}%`]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/envios  body: { fecha: 'YYYY-MM-DD', envios: [{guia,destinatario,razonSocial,ciudad}] }
  // Reemplaza todos los envíos de esa fecha (un manifiesto = una subida = el día completo)
  if (urlPath === '/api/envios' && req.method === 'POST') {
    try {
      const { fecha, envios } = await bodyJSON(req);
      if(!fecha || !Array.isArray(envios)){
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'fecha y envios son requeridos'}));
        return;
      }
      await pool.query('DELETE FROM envios_servientrega WHERE fecha=$1', [fecha]);
      for(const e of envios){
        await pool.query(
          `INSERT INTO envios_servientrega(guia, fecha, destinatario, razon_social, ciudad)
           VALUES($1,$2,$3,$4,$5)
           ON CONFLICT (guia, fecha) DO UPDATE SET destinatario=$3, razon_social=$4, ciudad=$5`,
          [e.guia, fecha, e.destinatario||'', e.razonSocial||'', e.ciudad||'']
        );
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, fecha, total: envios.length }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // DELETE /api/envios?fecha=YYYY-MM-DD
  if (urlPath === '/api/envios' && req.method === 'DELETE') {
    try {
      const fecha = urlObj.searchParams.get('fecha');
      if(!fecha){
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'fecha es requerida'}));
        return;
      }
      await pool.query('DELETE FROM envios_servientrega WHERE fecha=$1', [fecha]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // PEDIDOS WEB — pedidos recibidos por correo desde la tienda WooCommerce
  // GET /api/pedidos-web?fecha=YYYY-MM-DD  → pedidos de un día específico
  // GET /api/pedidos-web?dias=7            → pedidos de los últimos N días
  if (urlPath === '/api/pedidos-web' && req.method === 'GET') {
    try {
      const fecha = urlObj.searchParams.get('fecha');
      const fechaControl = urlObj.searchParams.get('fecha_control');
      const dias = parseInt(urlObj.searchParams.get('dias')) || null;
      let r;
      if (fechaControl) {
        // Vista Control: usa fecha de despacho (fecha_control si el pedido fue movido)
        r = await pool.query('SELECT * FROM pedidos_web WHERE COALESCE(fecha_control, fecha)=$1 ORDER BY id DESC', [fechaControl]);
      } else if (fecha) {
        r = await pool.query('SELECT * FROM pedidos_web WHERE fecha=$1 ORDER BY id DESC', [fecha]);
      } else if (dias) {
        r = await pool.query(`SELECT * FROM pedidos_web WHERE fecha >= (CURRENT_DATE - $1::int) ORDER BY fecha DESC, id DESC`, [dias]);
      } else {
        r = await pool.query('SELECT * FROM pedidos_web ORDER BY fecha DESC, id DESC LIMIT 200');
      }
      // Índice por nombre para pedidos sin cédula (o con cédula que no matchea)
      const CRED_POR_NOMBRE = (() => {
        const m = {};
        const nk = x => String(x||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
        Object.values(CREDITO_CACHE).forEach(v => { const k = nk(v.nombre); if (k && v.cupo > 0) m[k] = v; });
        return { m, nk };
      })();
      const pedidos = r.rows.map(row => ({
        ...row,
        credito: (() => {
          const d = String(row.cedula_ruc || '').replace(/\D/g, '');
          let c = d ? (CREDITO_CACHE[d] || (d.length === 13 ? CREDITO_CACHE[d.substring(0,10)] : null) || CREDITO_CACHE[d + '001'] || null) : null;
          if (!c || !(c.cupo > 0)) {
            // Fallback por nombre del cliente (cubre pedidos sin cédula o con cédula distinta a la de Contifico)
            const kn = CRED_POR_NOMBRE.nk(row.cliente_nombre);
            const alt = kn ? (CRED_POR_NOMBRE.m[kn] || null) : null;
            if (alt) c = alt;
          }
          if (!d && !c) return null;
          return c ? { cupo: c.cupo, dias: c.dias, aplica: c.aplica } : { cupo: 0, dias: 0, aplica: false };
        })(),
        productos: (() => { try { return JSON.parse(row.productos || '[]'); } catch(e){ return []; } })()
      }));
      // Enriquecer con la asesora asignada en el directorio (personas). Sin match → cliente nueva
      try {
        if (pedidos.length) {
          const rp = await pool.query("SELECT cedula, ruc, razon_social, vendedor FROM personas WHERE vendedor IS NOT NULL AND vendedor <> ''");
          const normP = x => String(x||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
          const porCed = {}; const porNom = [];
          rp.rows.forEach(pe => {
            [pe.cedula, pe.ruc].forEach(v => {
              const d = String(v||'').replace(/\D/g,'');
              if (d) { porCed[d] = pe.vendedor; if (d.length === 13) porCed[d.substring(0,10)] = pe.vendedor; }
            });
            const n = normP(pe.razon_social);
            if (n) porNom.push({ n, ancla: n.split(' ').slice(0,2).join(' '), v: pe.vendedor });
          });
          pedidos.forEach(px => {
            let ase = null;
            const d = String(px.cedula_ruc||'').replace(/\D/g,'');
            if (d) ase = porCed[d] || (d.length === 13 ? porCed[d.substring(0,10)] : null);
            if (!ase) {
              const n = normP(px.cliente_nombre);
              if (n) {
                const anclaP = n.split(' ').slice(0,2).join(' ');
                const m = porNom.find(x => x.n === n || (anclaP.length >= 7 && (x.n.startsWith(anclaP) || n.startsWith(x.ancla))));
                if (m) ase = m.v;
              }
            }
            px.asesora = ase || null;
          });
        }
      } catch(eA) { /* si falla, los pedidos salen sin asesora */ }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(pedidos));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // GET /api/pedidos-web/buscar?q=nombre → búsqueda por cliente en todas las fechas
  if (urlPath === '/api/pedidos-web/buscar' && req.method === 'GET') {
    try {
      const q = (urlObj.searchParams.get('q')||'').trim();
      if (!q || q.length < 2) {
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify([]));
        return;
      }
      const r = await pool.query(
        `SELECT * FROM pedidos_web WHERE LOWER(cliente_nombre) LIKE LOWER($1) OR cedula_ruc LIKE $1
         ORDER BY fecha DESC LIMIT 200`,
        [`%${q}%`]
      );
      const pedidos = r.rows.map(row => ({
        ...row,
        productos: (() => { try { return JSON.parse(row.productos || '[]'); } catch(e){ return []; } })()
      }));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(pedidos));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // GET /api/pedidos-web/sync → fuerza una sincronización manual con la casilla de correo
  if (urlPath === '/api/pedidos-web/sync' && req.method === 'GET') {
    try {
      const resultado = await sincronizarPedidosWeb();
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(resultado));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/pedidos-web/mover-siguiente-dia  body: { numeroPedido }
  // Mueve el pedido al día siguiente y lo marca como "facturado ayer" — para
  // pedidos facturados después de que el courier ya pasó (despachan al día siguiente)
  if (urlPath === '/api/pedidos-web/mover-siguiente-dia' && req.method === 'POST') {
    try {
      const { numeroPedido } = await bodyJSON(req);
      // Solo se mueve la fecha de CONTROL (despacho) — la fecha original del pedido no cambia
      await pool.query(
        `UPDATE pedidos_web SET fecha_control = COALESCE(fecha_control, fecha) + INTERVAL '1 day', facturado_ayer = true WHERE numero_pedido=$1`,
        [numeroPedido]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // POST /api/pedidos-web/marcar-facturado  body: { numeroPedido, documentoFactura }
  // Permite marcar manualmente un pedido como facturado (por si el cruce automático no
  // lo detecta, ej. el nombre en Contifico es muy distinto al de la web)
  // CREAR PREFACTURA EN CONTIFICO desde un pedido web (María José solo revisa y factura)
  if (/^\/api\/pedidos-web\/[^/]+\/prefactura$/.test(urlPath) && req.method === 'POST') {
    try {
      let WOO_URL = (process.env.WOO_URL || '').trim();
      const mUrl = WOO_URL.match(/https?:\/\/[^\s"']+/); // tolera valores tipo "WOO_URL = https://..."
      WOO_URL = (mUrl ? mUrl[0] : WOO_URL).replace(/\/+$/, '');
      const WOO_CK = process.env.WOO_CK || process.env.WC_CONSUMER_KEY || '';
      const WOO_CS = process.env.WOO_CS || process.env.WC_CONSUMER_SECRET || '';
      if (!WOO_URL || !WOO_CK || !WOO_CS) throw new Error('Faltan credenciales de WooCommerce en Railway (WOO_URL + WC_CONSUMER_KEY/WC_CONSUMER_SECRET)');
      if (!API_KEY) throw new Error('CONTIFICO_API_KEY no configurada');
      const numero = decodeURIComponent(urlPath.split('/')[3]);
      const rP = await pool.query('SELECT * FROM pedidos_web WHERE numero_pedido=$1', [numero]);
      if (!rP.rows.length) throw new Error('Pedido no encontrado: ' + numero);
      const ped = rP.rows[0];
      if (ped.prefactura_doc) { res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, ya_existia:true, prefactura: ped.prefactura_doc})); return; }

      // 1) Pedido completo desde WooCommerce (SKU, cédula y montos confiables)
      const wooResp = await fetch(`${WOO_URL}/wp-json/wc/v3/orders/${encodeURIComponent(numero)}?consumer_key=${encodeURIComponent(WOO_CK)}&consumer_secret=${encodeURIComponent(WOO_CS)}`, { headers: { 'Accept': 'application/json' } });
      if (!wooResp.ok) throw new Error('WooCommerce respondió ' + wooResp.status + ' al pedir el pedido ' + numero);
      const orden = await wooResp.json();

      // 2) Cédula/RUC: campos de facturación + meta_data (10 o 13 dígitos)
      let cedula = '';
      const bill = orden.billing || {};
      // 1º prioridad: campos cuyo NOMBRE indique cédula/RUC (checkout ecuatoriano)
      const candidatos = [];
      (orden.meta_data || []).forEach(m => {
        const k = String((m && m.key) || '').toLowerCase();
        if (/cedula|c\u00e9dula|ruc|identificacion|identificaci\u00f3n|dni|nit/.test(k)) candidatos.push(String((m && m.value) || ''));
      });
      Object.entries(bill).forEach(([k, v]) => {
        if (/cedula|ruc|identific|dni/.test(String(k).toLowerCase())) candidatos.push(String(v||''));
      });
      // 2º prioridad: resto de campos, EXCLUYENDO teléfono y código postal (10 dígitos engañosos)
      Object.entries(bill).forEach(([k, v]) => {
        const kl = String(k).toLowerCase();
        if (kl.includes('phone') || kl.includes('postcode') || kl.includes('telefono')) return;
        candidatos.push(String(v||''));
      });
      (orden.meta_data || []).forEach(m => {
        const k = String((m && m.key) || '').toLowerCase();
        if (k.includes('phone') || k.includes('telefono')) return;
        candidatos.push(String((m && m.value) || ''));
      });
      for (const c of candidatos) {
        const d = c.replace(/\D/g, '');
        if (d.length === 10 || d.length === 13) { cedula = d; break; }
      }
      if (!cedula && ped.cedula_ruc) cedula = String(ped.cedula_ruc).replace(/\D/g, '');
      const nombreCli = `${bill.first_name||''} ${bill.last_name||''}`.trim() || ped.cliente_nombre || 'CONSUMIDOR FINAL';

      // 3) Buscar la persona en Contifico por cédula/RUC; crearla si no existe
      let persona = null;
      let clienteCreado = false;
      if (cedula) {
        for (const param of (cedula.length === 13 ? ['ruc','cedula'] : ['cedula','ruc'])) {
          try {
            const rB = await fetch(`https://api.contifico.com/sistema/api/v1/persona/?${param}=${cedula}&page_size=5`, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
            if (!rB.ok) continue;
            const dB = await rB.json();
            const lista = Array.isArray(dB) ? dB : (dB.results || []);
            persona = lista.find(p => String(p.cedula||'').replace(/\D/g,'') === cedula || String(p.ruc||'').replace(/\D/g,'') === cedula) || lista[0] || null;
            if (persona) break;
          } catch(e) {}
        }
      }
      if (!persona && cedula) {
        // Crear la clienta nueva en Contifico con los datos del checkout
        const PVP_WEB_NEW = (process.env.CONTIFICO_PVP_WEB || 'pvp1').toLowerCase().replace(/[^a-z0-9]/g,'');
        const cuerpoP = {
          tipo: 'N', es_cliente: true,
          pvp_default: PVP_WEB_NEW,
          razon_social: nombreCli.toUpperCase(),
          telefonos: String(bill.phone || ped.telefono || ''),
          email: String(bill.email || ''),
          direccion: [bill.address_1, bill.city, bill.state].filter(Boolean).join(', ')
        };
        if (cedula.length === 13) cuerpoP.ruc = cedula; else cuerpoP.cedula = cedula;
        try {
          const rC = await fetch('https://api.contifico.com/sistema/api/v1/persona/', {
            method: 'POST', headers: { 'Authorization': API_KEY, 'Content-Type': 'application/json' },
            body: JSON.stringify(cuerpoP)
          });
          const dC = await rC.json();
          if (rC.ok && dC && dC.id) { persona = dC; clienteCreado = true; }
          else console.log('No se pudo crear persona:', JSON.stringify(dC).substring(0,300));
        } catch(e) { console.log('Error creando persona:', e.message); }
      }

      // 4) Cruzar productos por SKU contra el catálogo de Contifico
      const porSku = {};
      Object.entries(catalogoProductos || {}).forEach(([id, info]) => {
        const c = String(info.codigo || '').trim().toUpperCase();
        if (c) porSku[c] = { id, nombre: info.nombre };
      });
      // PVP que corresponde a la clienta en Contifico (pvp1..pvp4); por defecto pvp1
      // PVP de la clienta; si no tiene (clienta nueva o consumidor final) → el configurado
      // en Railway con CONTIFICO_PVP_WEB (por defecto pvp1)
      const PVP_WEB = (process.env.CONTIFICO_PVP_WEB || 'pvp1').toLowerCase().replace(/[^a-z0-9]/g,'');
      const pvpPersona = String((persona && persona.pvp_default) || '').toLowerCase().replace(/[^a-z0-9]/g,'');
      const pvpKey = /^pvp[1-4]$/.test(pvpPersona) ? pvpPersona : PVP_WEB;
      const detalles = []; const noCruzados = []; const sinPrecio = [];
      (orden.line_items || []).forEach(it => {
        const sku = String(it.sku || '').trim().toUpperCase();
        const qty = parseFloat(it.quantity || 0);
        const totalLinea = parseFloat(it.total || 0);
        if (!qty) return;
        const match = sku ? porSku[sku] : null;
        if (!match) { noCruzados.push(`${it.name || sku || '?'} (SKU: ${sku || 'sin SKU'}) × ${qty}`); return; }
        const info = catalogoProductos[match.id] || {};
        // Precio de LISTA de Contifico según el PVP de la clienta.
        // Los pvp de Contifico ya vienen SIN IVA → se envían tal cual (no dividir por 1.15)
        const ivaProd = (info.iva === 0) ? 0 : 15;
        const precioBase = Math.round((info[pvpKey] || info.pvp1 || info.pvp2 || info.pvp3 || info.pvp4 || 0) * 10000) / 10000;
        if (!precioBase) { sinPrecio.push(`${info.nombre || sku} (sin PVP en Contifico)`); }
        // Regalos / promos: el pedido trae $0 → va el precio de lista con 100% de descuento
        const esRegalo = totalLinea <= 0.009;
        const desc = esRegalo ? 100 : 0;
        const baseGrav = esRegalo ? 0 : Math.round(precioBase * qty * 100) / 100;
        const det = {
          producto_id: match.id,
          cantidad: qty,
          precio: precioBase,
          porcentaje_iva: ivaProd,
          porcentaje_descuento: desc,
          base_cero: 0,
          base_no_gravable: 0
        };
        if (ivaProd === 0) { det.base_cero = baseGrav; det.base_gravable = 0; }
        else { det.base_gravable = baseGrav; }
        detalles.push(det);
      });
      if (detalles.length === 0) throw new Error('Ningún producto del pedido cruzó por SKU con Contifico. Sin cruzar: ' + noCruzados.join(' · '));

      // 5) Crear la prefactura (PRE) en Contifico
      const baseTotal = Math.round(detalles.reduce((a,d)=>a+(d.base_gravable||0),0) * 100) / 100;
      const baseCeroTotal = Math.round(detalles.reduce((a,d)=>a+(d.base_cero||0),0) * 100) / 100;
      const ivaTotal = Math.round(baseTotal * 0.15 * 100) / 100;
      const hoyEC2 = nowEC();
      const fechaDoc = `${String(hoyEC2.getDate()).padStart(2,'0')}/${String(hoyEC2.getMonth()+1).padStart(2,'0')}/${hoyEC2.getFullYear()}`;
      const POS_ID = (process.env.CONTIFICO_POS || process.env.CONTIFICO_POS_ID || '').trim();
      if (!POS_ID) throw new Error('Falta la variable CONTIFICO_POS en Railway: es el "Token POS" que Contifico te dio junto a la API Key (Configuración → Integraciones/API)');
      const cuerpoDoc = {
        pos: POS_ID,
        fecha_emision: fechaDoc,
        tipo_documento: 'PRE',
        documento: '',
        estado: 'P',
        electronico: false,
        descripcion: `Pedido web #${numero}` + (persona ? '' : ` — cliente: ${nombreCli} (${cedula || 'sin cédula'}) NO ENCONTRADO, asignar manualmente`),
        subtotal_0: baseCeroTotal,
        subtotal_12: baseTotal,
        iva: ivaTotal,
        total: Math.round((baseTotal + baseCeroTotal + ivaTotal) * 100) / 100,
        detalles
      };
      if (persona) {
        cuerpoDoc.cliente = {
          id: persona.id,
          tipo: persona.tipo || (String(persona.ruc||'').replace(/\D/g,'').length === 13 ? 'J' : 'N'),
          ruc: persona.ruc || undefined,
          cedula: persona.cedula || undefined,
          razon_social: persona.razon_social,
          telefonos: persona.telefonos || undefined,
          direccion: persona.direccion || undefined,
          email: persona.email || undefined
        };
      } else {
        cuerpoDoc.cliente = { cedula: '9999999999999', razon_social: 'CONSUMIDOR FINAL', tipo: 'N' };
      }
      const rD = await fetch('https://api.contifico.com/sistema/api/v1/documento/', {
        method: 'POST', headers: { 'Authorization': API_KEY, 'Content-Type': 'application/json' },
        body: JSON.stringify(cuerpoDoc)
      });
      const dD = await rD.json().catch(()=>({}));
      if (!rD.ok || !dD || (!dD.id && !dD.documento)) {
        throw new Error('Contifico rechazó la prefactura: ' + JSON.stringify(dD).substring(0, 400));
      }
      const docNum = dD.documento || dD.id;
      await pool.query('UPDATE pedidos_web SET prefactura_doc=$1, prefactura_at=NOW() WHERE numero_pedido=$2', [String(docNum).substring(0,90), numero]);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, prefactura: docNum, cliente: persona ? persona.razon_social : 'CONSUMIDOR FINAL (asignar)', cliente_creado: clienteCreado, productos_cruzados: detalles.length, no_cruzados: noCruzados, sin_precio: sinPrecio, pvp_usado: pvpKey }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false, error:e.message})); }
    return;
  }

  // Reiniciar la marca de prefactura (para volver a probar) — ?fecha=YYYY-MM-DD o ?numero=17355
  if (urlPath === '/api/pedidos-web/reset-prefactura' && req.method === 'POST') {
    try {
      const numeroR = urlObj.searchParams.get('numero');
      const fechaR = urlObj.searchParams.get('fecha') || nowEC().toLocaleDateString('en-CA');
      let r;
      if (numeroR) r = await pool.query('UPDATE pedidos_web SET prefactura_doc=NULL, prefactura_at=NULL WHERE numero_pedido=$1', [numeroR]);
      else r = await pool.query('UPDATE pedidos_web SET prefactura_doc=NULL, prefactura_at=NULL WHERE fecha=$1', [fechaR]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, reiniciados:r.rowCount, criterio: numeroR ? ('pedido '+numeroR) : ('fecha '+fechaR)}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (urlPath === '/api/pedidos-web/marcar-facturado' && req.method === 'POST') {
    try {
      const { numeroPedido, documentoFactura } = await bodyJSON(req);
      await pool.query(
        'UPDATE pedidos_web SET facturado=true, documento_factura=$2 WHERE numero_pedido=$1',
        [numeroPedido, documentoFactura || null]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // DEBUG TEMPORAL: ver el HTML crudo guardado de un pedido para ajustar el parser
  // GET /api/pedidos-web/debug-html?numero=16605
  if (urlPath === '/api/pedidos-web/debug-html' && req.method === 'GET') {
    try {
      const numero = urlObj.searchParams.get('numero');
      const r = await pool.query('SELECT numero_pedido, html_crudo FROM pedidos_web WHERE numero_pedido=$1', [numero]);
      if (r.rows.length === 0) {
        res.writeHead(404,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Pedido no encontrado'}));
        return;
      }
      res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});
      res.end(`<pre>${(r.rows[0].html_crudo||'').replace(/</g,'&lt;')}</pre>`);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // GET /api/pedidos-web/resync-todos → reprocesa TODOS los correos de "nuevo pedido" en
  // la bandeja (leídos o no), útil para recapturar el html_crudo de pedidos ya procesados
  // o para corregir datos tras un ajuste al parser.
  if (urlPath === '/api/pedidos-web/resync-todos' && req.method === 'GET') {
    try {
      const resultado = await sincronizarPedidosWeb({ incluirLeidos: true });
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(resultado));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ACTUALIZAR PERMISO editar_visitas PARA FERNANDO Y GIOVANNA (una sola vez)
  if (urlPath === '/api/fix-permisos-visitas' && req.method === 'GET') {
    try {
      const r1 = await pool.query("UPDATE usuarios SET modulos = modulos || ',editar_visitas' WHERE usuario IN ('Fernando','Giovanna') AND modulos NOT LIKE '%editar_visitas%' RETURNING nombre, modulos");
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ actualizados: r1.rows }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ASESOR ZONAS/SECTORES Y PROVINCIAS — overrides editables sobre ASESOR_DATA
  if (urlPath === '/api/asesor-config' && req.method === 'GET') {
    try {
      const zonas = await pool.query('SELECT asesora, zona, sector FROM asesor_zonas ORDER BY id');
      const provincias = await pool.query('SELECT asesora, provincia FROM asesor_provincias ORDER BY id');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ zonas: zonas.rows, provincias: provincias.rows }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/asesor-zona' && req.method === 'POST') {
    try {
      const {asesora, zona, sector} = await bodyJSON(req);
      await pool.query('INSERT INTO asesor_zonas(asesora,zona,sector) VALUES($1,$2,$3) ON CONFLICT DO NOTHING',[asesora,zona,sector]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/asesor-zona' && req.method === 'DELETE') {
    try {
      const {asesora, zona, sector} = await bodyJSON(req);
      if (sector) {
        await pool.query('DELETE FROM asesor_zonas WHERE asesora=$1 AND zona=$2 AND sector=$3',[asesora,zona,sector]);
      } else {
        await pool.query('DELETE FROM asesor_zonas WHERE asesora=$1 AND zona=$2',[asesora,zona]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/asesor-provincia' && req.method === 'POST') {
    try {
      const {asesora, provincia} = await bodyJSON(req);
      await pool.query('INSERT INTO asesor_provincias(asesora,provincia) VALUES($1,$2) ON CONFLICT DO NOTHING',[asesora,provincia]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/asesor-provincia' && req.method === 'DELETE') {
    try {
      const {asesora, provincia} = await bodyJSON(req);
      await pool.query('DELETE FROM asesor_provincias WHERE asesora=$1 AND provincia=$2',[asesora,provincia]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // EQUIPOS
  if (urlPath === '/api/equipos' && req.method === 'GET') {
    try {
      const eR = await pool.query('SELECT * FROM equipos WHERE activo=true ORDER BY id');
      const mR = await pool.query('SELECT equipo_id, usuario_nombre FROM equipo_miembros');
      const equipos = eR.rows.map(e => ({
        id: e.id, nombre: e.nombre, lider: e.lider,
        miembros: mR.rows.filter(m=>m.equipo_id===e.id).map(m=>m.usuario_nombre)
      }));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(equipos));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/equipos' && req.method === 'POST') {
    try {
      const { nombre, lider, miembros } = await bodyJSON(req);
      if(!nombre){ res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Nombre requerido'})); return; }
      const r = await pool.query('INSERT INTO equipos(nombre,lider) VALUES($1,$2) RETURNING id', [nombre, lider||null]);
      const eid = r.rows[0].id;
      for(const m of (miembros||[])) await pool.query('INSERT INTO equipo_miembros(equipo_id,usuario_nombre) VALUES($1,$2) ON CONFLICT DO NOTHING', [eid, m]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:eid}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  const equipoMatch = urlPath.match(/^\/api\/equipos\/(\d+)$/);
  if (equipoMatch && req.method === 'PUT') {
    try {
      const eid = equipoMatch[1];
      const { nombre, lider, miembros } = await bodyJSON(req);
      if(nombre!==undefined) await pool.query('UPDATE equipos SET nombre=$1 WHERE id=$2',[nombre,eid]);
      if(lider!==undefined) await pool.query('UPDATE equipos SET lider=$1 WHERE id=$2',[lider,eid]);
      if(Array.isArray(miembros)){
        await pool.query('DELETE FROM equipo_miembros WHERE equipo_id=$1',[eid]);
        for(const m of miembros) await pool.query('INSERT INTO equipo_miembros(equipo_id,usuario_nombre) VALUES($1,$2) ON CONFLICT DO NOTHING',[eid,m]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (equipoMatch && req.method === 'DELETE') {
    try {
      await pool.query('UPDATE equipos SET activo=false WHERE id=$1',[equipoMatch[1]]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // KPI CUMPLIMIENTO VISITAS (semana + mes por asesora vs meta)
  if (urlPath === '/api/kpi-cumplimiento-visitas' && req.method === 'GET') {
    try {
      const semana = urlObj.searchParams.get('semana'); // YYYY-MM-DD (lunes)
      const mes = urlObj.searchParams.get('mes'); // YYYY-MM
      const uR = await pool.query(`SELECT nombre FROM usuarios WHERE rol IN ('asesora','jefa_ventas') AND activo=true`);
      const asesoras = uR.rows.map(r=>r.nombre);
      const mR = await pool.query('SELECT asesora, meta FROM metas_visitas');
      const metas = {}; mR.rows.forEach(r=>{ metas[r.asesora]=parseInt(r.meta)||30; });
      const sR = semana ? await pool.query(
        `SELECT asesora, COUNT(*) FILTER (WHERE coordinado) AS visitadas FROM planificacion WHERE semana=$1 GROUP BY asesora`, [semana]
      ) : {rows:[]};
      const mesR = mes ? await pool.query(
        `SELECT asesora, COUNT(*) FILTER (WHERE coordinado) AS visitadas FROM planificacion WHERE TO_CHAR(semana,'YYYY-MM')=$1 GROUP BY asesora`, [mes]
      ) : {rows:[]};
      // Desglose por semana del mes
      const semR = mes ? await pool.query(
        `SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, COUNT(*) FILTER (WHERE coordinado) AS visitadas
         FROM planificacion
         WHERE semana >= (TO_DATE($1,'YYYY-MM') - INTERVAL '10 days')
           AND semana <  (TO_DATE($1,'YYYY-MM') + INTERVAL '1 month' + INTERVAL '10 days')
         GROUP BY asesora, semana ORDER BY semana`, [mes]
      ) : {rows:[]};
      const vSem = {}; sR.rows.forEach(r=>{ vSem[r.asesora]=parseInt(r.visitadas)||0; });
      const vMes = {}; mesR.rows.forEach(r=>{ vMes[r.asesora]=parseInt(r.visitadas)||0; });
      const detalle = asesoras.map(a=>({
        asesora: a,
        meta: metas[a]||30,
        visitadas_semana: vSem[a]||0,
        visitadas_mes: vMes[a]||0,
        cumple_semana: (vSem[a]||0) >= (metas[a]||30)
      }));
      const cumplen = detalle.filter(d=>d.cumple_semana).length;
      const pct = asesoras.length ? Math.round(cumplen/asesoras.length*100) : 0;
      // semanas: [{asesora, semana, visitadas}]
      const semanas = semR.rows.map(r=>({ asesora: r.asesora, semana: r.semana, visitadas: parseInt(r.visitadas)||0 }));
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ detalle, semanas, pct_semana: pct, total_asesoras: asesoras.length, cumplen_semana: cumplen }));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // REVISIONES LUNES (Giovanna marca revisión semanal por asesora)
  if (urlPath === '/api/revisiones-lunes' && req.method === 'GET') {
    try {
      const semana = urlObj.searchParams.get('semana');
      const mes = urlObj.searchParams.get('mes'); // YYYY-MM para el consolidado
      let r;
      if (semana) r = await pool.query(`SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, revisado FROM revisiones_lunes WHERE semana=$1`, [semana]);
      else if (mes) r = await pool.query(`SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, revisado FROM revisiones_lunes WHERE TO_CHAR(semana,'YYYY-MM')=$1`, [mes]);
      else r = await pool.query(`SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, revisado FROM revisiones_lunes ORDER BY semana DESC LIMIT 200`);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/revisiones-lunes' && req.method === 'POST') {
    try {
      const { asesora, semana, revisado } = await bodyJSON(req);
      await pool.query(
        `INSERT INTO revisiones_lunes(asesora,semana,revisado) VALUES($1,$2,$3)
         ON CONFLICT(asesora,semana) DO UPDATE SET revisado=$3`,
        [asesora, semana, !!revisado]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // KPI METAS (configurables por admin)
  if (urlPath === '/api/kpi-metas' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT clave, meta FROM kpi_metas');
      const defaults = {
        cap_oficina: '1', cap_provincia: '1', cap_virtuales: '1', inst_aperturas: '1', inst_visitas: '4',
        revisiones: 'Todos los lunes', visitas_cumplimiento: '100', giras_pct: '100', casas_pct: '100'
      };
      const metas = {...defaults};
      r.rows.forEach(row => { metas[row.clave] = row.meta; });
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(metas));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/kpi-metas' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      for (const [clave, meta] of Object.entries(body)) {
        await pool.query(
          `INSERT INTO kpi_metas(clave,meta) VALUES($1,$2)
           ON CONFLICT(clave) DO UPDATE SET meta=$2, updated_at=NOW()`,
          [clave, String(meta)]
        );
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // METAS VISITAS
  if (urlPath === '/api/metas-visitas' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT asesora, meta FROM metas_visitas ORDER BY asesora');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/metas-visitas' && req.method === 'POST') {
    try {
      const {asesora, meta} = await bodyJSON(req);
      await pool.query(
        `INSERT INTO metas_visitas(asesora,meta) VALUES($1,$2)
         ON CONFLICT(asesora) DO UPDATE SET meta=$2, updated_at=NOW()`,
        [asesora, parseInt(meta)||30]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e){ res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // PLANIFICACION
  if (urlPath === '/api/planificacion' && req.method === 'GET') {
    try {
      const asesora = urlObj.searchParams.get('asesora') || '';
      const semana = urlObj.searchParams.get('semana') || '';
      let r = await pool.query('SELECT * FROM planificacion WHERE asesora=$1 AND semana=$2 ORDER BY id',[asesora,semana]);
      // Si no hay coincidencia exacta, probar con coincidencia parcial (nombre guardado
      // puede ser más corto o más largo, ej. "Karen Rebeca Mora" vs "Karen Rebeca Mora
      // Cedeño") — usa las dos primeras palabras (nombre + primer apellido) como ancla.
      if (r.rows.length === 0 && asesora) {
        const ancla = asesora.trim().split(' ').slice(0,2).join(' ');
        r = await pool.query("SELECT * FROM planificacion WHERE asesora ILIKE $1 AND semana=$2 ORDER BY id", [ancla+'%', semana]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/planificacion' && req.method === 'POST') {
    const cli = await pool.connect();
    try {
      const {asesora,semana,filas} = await bodyJSON(req);
      if (!asesora||!semana||!filas) throw new Error('Faltan datos');

      // Quitar filas repetidas antes de guardar. Un mismo cliente puede aparecer dos veces
      // legítimamente (dos visitas el mismo día), pero NO con la misma hora de visita: eso
      // solo pasa cuando el registro se duplicó. Las filas sin visitar se respetan todas.
      const vistas = new Set();
      const limpias = [];
      let quitadas = 0;
      for (const f of filas) {
        if (f.visitado_at) {
          const k = [f.dia||'', f.sector||'', f.cliente||'', f.visitado_at].join('|');
          if (vistas.has(k)) { quitadas++; continue; }
          vistas.add(k);
        }
        limpias.push(f);
      }

      // Transacción: borrar e insertar tiene que ser una sola operación indivisible. Sin
      // esto, dos guardados simultáneos (doble clic o reintento de red) se entrelazaban —
      // ambos borraban y después ambos insertaban, dejando la planificación por duplicado.
      await cli.query('BEGIN');
      await cli.query('DELETE FROM planificacion WHERE asesora=$1 AND semana=$2',[asesora,semana]);
      for (const fila of limpias) {
        await cli.query('INSERT INTO planificacion(asesora,semana,dia,sector,cliente,coordinado,visitado_at,primera_compra_at,recompra_at,observaciones) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
          [asesora,semana,fila.dia||'',fila.sector||'',fila.cliente||'',fila.coordinado||false,fila.visitado_at||null,fila.primera_compra_at||null,fila.recompra_at||null,fila.observaciones||'']);
      }
      await cli.query('COMMIT');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true,filas:limpias.length,duplicados_quitados:quitadas}));
    } catch(e) {
      try { await cli.query('ROLLBACK'); } catch(e2) {}
      res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message}));
    } finally { cli.release(); }
    return;
  }

  // Limpieza de duplicados ya guardados: deja una sola fila por (asesora, semana, día,
  // sector, cliente, hora de visita). Solo toca filas con hora registrada, que son las
  // que no pueden repetirse de forma legítima.
  if (urlPath === '/api/planificacion/limpiar-duplicados' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const r = await pool.query(`
        DELETE FROM planificacion p
        WHERE visitado_at IS NOT NULL AND visitado_at <> ''
          AND EXISTS (
            SELECT 1 FROM planificacion q
            WHERE q.asesora = p.asesora AND q.semana = p.semana
              AND COALESCE(q.dia,'') = COALESCE(p.dia,'')
              AND COALESCE(q.sector,'') = COALESCE(p.sector,'')
              AND COALESCE(q.cliente,'') = COALESCE(p.cliente,'')
              AND q.visitado_at = p.visitado_at
              AND q.id < p.id
          )
        RETURNING id`);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, eliminadas: r.rowCount }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // VENTAS HOY (caché v2)
  // VENTAS PENDIENTES: todos los documentos desde el día siguiente al corte del caché
  // hasta hoy. Reemplaza al viejo "solo hoy", que perdía días completos cada vez que la
  // regeneración nocturna se atrasaba o fallaba.
  if (urlPath === '/api/ventas-pendientes' && req.method === 'GET') {
    try {
      const out = await ventasPendientes(urlObj.searchParams.get('forzar') === '1');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify(out));
    } catch(e) {
      // Ante cualquier fallo, devolver al menos el caché de hoy para no dejar la app sin datos
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ total: cache.documentos.length, documentos: cache.documentos, nc_documentos: cache.nc_documentos || [], degradado: true, error: e.message }));
    }
    return;
  }

  // ─── ESTADOS FINANCIEROS (PyG y Balance descargados de Contifico) ───────────────
  // El formato de Contifico es estable: fila de meses, columna A con el código jerárquico
  // de la cuenta, columna B con el nombre y una columna por mes. Se guarda el año completo
  // y cada carga nueva lo reemplaza, así subir el archivo de un mes actualiza todo.
  if (urlPath === '/api/finanzas/subir' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró el archivo (campo "file")'})); return; }
      const wb = XLSX.read(archivo.buffer, { type:'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
      const txt = x => String(x==null?'':x).trim();

      // Tipo y periodo salen de las primeras filas del encabezado
      let titulo = '', periodo = '';
      for (let i = 0; i < Math.min(6, filas.length); i++) {
        (filas[i]||[]).forEach(c => {
          const t = txt(c);
          if (/estado de resultados|situaci[oó]n financiera/i.test(t)) titulo = t;
          if (/^(desde el|hasta el)/i.test(t)) periodo = t;
        });
      }
      const tipo = /situaci[oó]n financiera/i.test(titulo) ? 'balance' : 'pyg';
      const mAnio = /(\d{4})\s*$/.exec(periodo) || /(20\d{2})/.exec(periodo);
      const anio = mAnio ? parseInt(mAnio[1]) : nowEC().getFullYear();

      // Fila de meses
      const MESES_N = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
      let filaMeses = -1, colIni = -1, meses = [];
      for (let i = 0; i < Math.min(15, filas.length); i++) {
        const f = filas[i] || [];
        const idx = f.findIndex(c => txt(c).toUpperCase() === 'ENERO');
        if (idx !== -1) {
          filaMeses = i; colIni = idx;
          for (let c = idx; c < f.length; c++) {
            const t = txt(f[c]).toUpperCase();
            if (MESES_N.includes(t)) meses.push(t.charAt(0) + t.slice(1).toLowerCase());
            else break;
          }
          break;
        }
      }
      if (filaMeses === -1) throw new Error('No se encontró la fila de meses (Enero, Febrero, ...) en el archivo');

      const num = v => { const n = parseFloat(String(v==null?'':v).replace(/[^0-9.\-]/g,'')); return isFinite(n) ? n : 0; };
      const cuentas = []; let resultado = null;
      for (let i = filaMeses + 1; i < filas.length; i++) {
        const f = filas[i]; if (!f) continue;
        const cod = txt(f[0]), nom = txt(f[1]);
        if (!nom) continue;
        const valores = meses.map((_,j) => num(f[colIni + j]));
        if (!cod) { resultado = { nombre: nom, valores }; continue; }
        cuentas.push({ codigo: cod, nombre: nom, nivel: cod.split('.').length, valores });
      }
      if (!cuentas.length) throw new Error('No se leyó ninguna cuenta del archivo');

      const datos = { tipo, anio, titulo, periodo, meses, cuentas, resultado };
      await pool.query(
        `INSERT INTO finanzas_reportes(tipo, anio, datos, archivo, actualizado_at) VALUES($1,$2,$3,$4,NOW())
         ON CONFLICT (tipo, anio) DO UPDATE SET datos=$3, archivo=$4, actualizado_at=NOW()`,
        [tipo, anio, JSON.stringify(datos), (archivo.filename||'').substring(0,290)]
      );
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, tipo, anio, cuentas: cuentas.length, meses: meses.length, titulo }));
    } catch(e) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // Diagnóstico de cartera: qué trae realmente el listado de documentos de Contifico
  if (urlPath === '/api/cartera/debug' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      // Por defecto una muestra corta: alcanza para ver la forma de los datos sin esperar
      const meses = parseInt(urlObj.searchParams.get('meses')) || 2;
      const maxPg = parseInt(urlObj.searchParams.get('paginas')) || 12;
      const hoyK = nowEC();
      const desdeD = new Date(hoyK); desdeD.setMonth(desdeD.getMonth() - meses);
      let url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fmtDateEC(desdeD)}&fecha_final=${fmtDateEC(hoyK)}&page_size=100`;
      let pg = 0;
      const stats = { docs:0, cli:0, fac:0, conSaldoCampo:0, saldoNull:0, saldoMayorCero:0,
        sumaSaldo:0, sumaTotalNoPagados:0, estados:{}, tipos:{}, campos:null, ejemplos:[] };
      while (url && pg < maxPg) {
        const r = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!r.ok) { stats.error = 'HTTP '+r.status; break; }
        const d = await r.json();
        (d.results || []).forEach(doc => {
          stats.docs++;
          if (!stats.campos) stats.campos = Object.keys(doc).sort();
          const t = String(doc.tipo_documento||'?');
          stats.tipos[t] = (stats.tipos[t]||0)+1;
          if (doc.tipo_registro !== 'CLI' || doc.anulado || noEsVenta(doc) || esNotaCredito(doc)) return;
          stats.cli++; stats.fac++;
          const est = String(doc.estado==null?'(null)':doc.estado);
          stats.estados[est] = (stats.estados[est]||0)+1;
          if (doc.saldo === undefined) return;
          stats.conSaldoCampo++;
          if (doc.saldo === null) { stats.saldoNull++; return; }
          const sal = parseFloat(doc.saldo)||0;
          if (sal > 0.01) { stats.saldoMayorCero++; stats.sumaSaldo += sal; }
          if (est !== 'P' && est !== 'C') stats.sumaTotalNoPagados += parseFloat(doc.total)||0;
          if (stats.ejemplos.length < 3 && sal > 0.01) stats.ejemplos.push({
            documento: doc.documento, tipo: doc.tipo_documento, fecha: doc.fecha_emision,
            estado: doc.estado, total: doc.total, saldo: doc.saldo,
            cliente: (doc.cliente && doc.cliente.razon_social) || '—',
            cobros: doc.cobros ? (Array.isArray(doc.cobros)? doc.cobros.length : 'obj') : 'sin campo'
          });
        });
        url = d.next || null; pg++;
      }
      stats.sumaSaldo = Math.round(stats.sumaSaldo*100)/100;
      stats.sumaTotalNoPagados = Math.round(stats.sumaTotalNoPagados*100)/100;
      stats.paginas = pg;
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify(stats, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // Qué versión está realmente desplegada. Sirve para saber si Railway tomó el último
  // push o sigue sirviendo un build anterior, sin tener que adivinar.
  if (urlPath === '/api/version' && req.method === 'GET') {
    try {
      const info = {};
      ['index.html','server.js','manifest.json'].forEach(f => {
        try {
          const st = fs.statSync(path.join(__dirname, f));
          info[f] = { modificado: new Date(st.mtime).toISOString(), bytes: st.size };
        } catch(e) { info[f] = 'no existe'; }
      });
      // Marcadores de funciones recientes: si dicen false, el archivo desplegado es viejo
      let idx = '';
      try { idx = fs.readFileSync(path.join(__dirname,'index.html'),'utf8'); } catch(e) {}
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        ok: true,
        arranque_servidor: new Date(Date.now() - Math.round(process.uptime()*1000)).toISOString(),
        uptime_minutos: Math.round(process.uptime()/60),
        archivos: info,
        tiene: {
          boton_imprimir_pedido: idx.includes('imprimirPedido'),
          flujo_de_caja: idx.includes('renderCaja'),
          panel_pyg: idx.includes('renderPyG'),
          boton_reiniciar_prefacturas: idx.includes('Reiniciar marcas de prefactura')
        }
      }, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (urlPath === '/api/cartera' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    // Nunca bloquear la respuesta: recorrer 12 meses de facturas toma minutos y dejaría el
    // panel colgado. Se dispara en segundo plano y el frontend reintenta.
    if (urlObj.searchParams.get('forzar') === '1' || !CARTERA.at) {
      sincronizarCartera().catch(e=>console.error(e));
    }
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok:true, sincronizando: CARTERA_EN_CURSO, ...CARTERA }));
    return;
  }

  // ─── FLUJO DE CAJA ──────────────────────────────────────────────────────────────
  if (urlPath === '/api/caja' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const hoyC = nowEC();
      const mk = String(urlObj.searchParams.get('mes') || (hoyC.getFullYear()+'-'+String(hoyC.getMonth()+1).padStart(2,'0')));
      const sal = await pool.query('SELECT * FROM caja_saldos WHERE mes_key=$1 ORDER BY tipo, orden, id', [mk]);
      const pag = await pool.query('SELECT * FROM caja_pagos WHERE mes_key=$1 ORDER BY dia, id', [mk]);
      const mes = await pool.query('SELECT DISTINCT mes_key FROM caja_pagos UNION SELECT DISTINCT mes_key FROM caja_saldos ORDER BY 1 DESC');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, mes_key:mk, saldos:sal.rows, pagos:pag.rows, meses:mes.rows.map(x=>x.mes_key) }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/caja/saldo' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const b = await bodyJSON(req);
      if (b.id) await pool.query('UPDATE caja_saldos SET nombre=$1, monto=$2, actualizado_at=NOW() WHERE id=$3', [String(b.nombre||'').substring(0,190), parseFloat(b.monto)||0, b.id]);
      else await pool.query('INSERT INTO caja_saldos(mes_key,tipo,nombre,monto,orden) VALUES($1,$2,$3,$4,$5)',
        [b.mes_key, b.tipo==='cobro'?'cobro':'banco', String(b.nombre||'').substring(0,190), parseFloat(b.monto)||0, parseInt(b.orden)||0]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/caja/pago' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const b = await bodyJSON(req);
      if (b.id) {
        const campos = [], vals = []; let i = 1;
        ['dia','concepto','monto','fuente','pagado','recurrente'].forEach(k=>{
          if (b[k] !== undefined) {
            campos.push(k+'=$'+i);
            vals.push(k==='monto' ? (parseFloat(b[k])||0) : k==='dia' ? (parseInt(b[k])||1) : k==='pagado'||k==='recurrente' ? !!b[k] : String(b[k]||'').substring(0,290));
            i++;
          }
        });
        if (campos.length) { vals.push(b.id); await pool.query(`UPDATE caja_pagos SET ${campos.join(',')}, actualizado_at=NOW() WHERE id=$${i}`, vals); }
      } else {
        await pool.query('INSERT INTO caja_pagos(mes_key,dia,concepto,monto,fuente,pagado,recurrente) VALUES($1,$2,$3,$4,$5,$6,$7)',
          [b.mes_key, parseInt(b.dia)||1, String(b.concepto||'').substring(0,290), parseFloat(b.monto)||0, String(b.fuente||'').substring(0,110), !!b.pagado, !!b.recurrente]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/caja\/(pago|saldo)\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const partes = urlPath.split('/');
      const tabla = partes[3] === 'pago' ? 'caja_pagos' : 'caja_saldos';
      await pool.query(`DELETE FROM ${tabla} WHERE id=$1`, [parseInt(partes[4])]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // Copiar el mes anterior: los pagos se traen sin marcar y los bancos con su último saldo,
  // que es lo que evita volver a teclear la misma lista cada mes.
  if (urlPath === '/api/caja/copiar' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const b = await bodyJSON(req);
      const mk = String(b.mes_key||'');
      const [a,m] = mk.split('-').map(Number);
      const prevD = new Date(a, m-2, 1);
      const prev = prevD.getFullYear()+'-'+String(prevD.getMonth()+1).padStart(2,'0');
      const yaP = await pool.query('SELECT COUNT(*)::int AS n FROM caja_pagos WHERE mes_key=$1', [mk]);
      const yaS = await pool.query('SELECT COUNT(*)::int AS n FROM caja_saldos WHERE mes_key=$1', [mk]);
      let pagos = 0, saldos = 0;
      if (yaP.rows[0].n === 0) {
        const r = await pool.query('INSERT INTO caja_pagos(mes_key,dia,concepto,monto,fuente,pagado,recurrente) SELECT $1,dia,concepto,monto,fuente,false,recurrente FROM caja_pagos WHERE mes_key=$2 RETURNING id', [mk, prev]);
        pagos = r.rowCount;
      }
      if (yaS.rows[0].n === 0) {
        const r = await pool.query('INSERT INTO caja_saldos(mes_key,tipo,nombre,monto,orden) SELECT $1,tipo,nombre,monto,orden FROM caja_saldos WHERE mes_key=$2 RETURNING id', [mk, prev]);
        saldos = r.rowCount;
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, desde:prev, pagos, saldos}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // Ajustes simples clave/valor (reutilizable). Solo admin.
  if (urlPath === '/api/ajuste' && (req.method === 'GET' || req.method === 'POST')) {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      if (req.method === 'GET') {
        const clave = String(urlObj.searchParams.get('clave')||'');
        const valor = await getConfigApp(clave, null);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, clave, valor}));
      } else {
        const b = await bodyJSON(req);
        await setConfigApp(String(b.clave||''), String(b.valor==null?'':b.valor));
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
      }
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (urlPath === '/api/finanzas' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const tipo = String(urlObj.searchParams.get('tipo') || 'pyg');
      const anioQ = parseInt(urlObj.searchParams.get('anio')) || null;
      const r = anioQ
        ? await pool.query('SELECT * FROM finanzas_reportes WHERE tipo=$1 AND anio=$2', [tipo, anioQ])
        : await pool.query('SELECT * FROM finanzas_reportes WHERE tipo=$1 ORDER BY anio DESC LIMIT 1', [tipo]);
      const rAnios = await pool.query('SELECT anio FROM finanzas_reportes WHERE tipo=$1 ORDER BY anio DESC', [tipo]);
      if (!r.rows.length) { res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, vacio:true, anios: rAnios.rows.map(x=>x.anio)})); return; }
      const fila = r.rows[0];
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, anios: rAnios.rows.map(x=>x.anio), actualizado: fila.actualizado_at, archivo: fila.archivo, ...JSON.parse(fila.datos) }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // SUBIR EXCEL DE COSTOS (el inventario valorizado que se descarga de Contifico).
  // Detecta sola la columna de código y la de costo, sin importar el orden ni el nombre exacto.
  if (urlPath === '/api/costos/subir' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró el archivo (campo "file")'})); return; }
      const wb = XLSX.read(archivo.buffer, { type:'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
      const norm = x => String(x==null?'':x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim();

      // Fila de encabezados: la primera que tenga algo parecido a "código"
      let fh = -1, encab = null;
      for (let i = 0; i < Math.min(25, filas.length); i++) {
        const f = (filas[i]||[]).map(norm);
        if (f.some(h => h === 'CODIGO' || h === 'COD' || h === 'SKU' || h.startsWith('CODIGO'))) { fh = i; encab = f; break; }
      }
      if (fh === -1) throw new Error('No se encontró una columna de Código en el Excel');

      const idxCod = encab.findIndex(h => h === 'CODIGO' || h === 'COD' || h === 'SKU' || h.startsWith('CODIGO'));
      const idxNom = encab.findIndex(h => h === 'PRODUCTO' || h === 'NOMBRE' || h.startsWith('DESCRIPCION'));
      // Costo unitario: preferir "costo unitario/promedio/prom"; nunca el costo total
      const cands = [];
      encab.forEach((h,i) => {
        if (!h || !h.includes('COSTO')) return;
        let score = 0;
        if (h.includes('UNIT')) score += 10;
        if (h.includes('PROM')) score += 8;
        if (h.includes('ULTIM')) score += 6;
        if (h === 'COSTO') score += 4;
        if (h.includes('TOTAL') || h.includes('VALOR')) score -= 12;
        cands.push({ i, h, score });
      });
      cands.sort((a,b) => b.score - a.score);
      const idxCosto = cands.length ? cands[0].i : -1;
      if (idxCosto === -1) throw new Error('No se encontró ninguna columna de Costo. Encabezados leídos: ' + encab.filter(Boolean).join(' | '));

      const vistos = {};
      for (let i = fh+1; i < filas.length; i++) {
        const f = filas[i];
        if (!f) continue;
        const cod = String(f[idxCod]==null?'':f[idxCod]).trim();
        if (!cod) continue;
        const costo = parseFloat(String(f[idxCosto]==null?'':f[idxCosto]).replace(/[^0-9.\-]/g,''));
        if (!isFinite(costo) || costo <= 0) continue;
        vistos[cod] = { costo, nombre: idxNom !== -1 ? String(f[idxNom]||'').trim() : '' };
      }
      const codigos = Object.keys(vistos);
      if (!codigos.length) throw new Error('No se leyó ningún costo mayor a cero. Columna usada: ' + (encab[idxCosto]||'?'));
      for (const cod of codigos) {
        await pool.query(
          `INSERT INTO producto_costos(codigo, nombre, costo, fuente, actualizado_at) VALUES($1,$2,$3,$4,NOW())
           ON CONFLICT (codigo) DO UPDATE SET nombre=$2, costo=$3, fuente=$4, actualizado_at=NOW()`,
          [cod, vistos[cod].nombre, vistos[cod].costo, 'Excel: ' + (encab[idxCosto]||'costo')]
        );
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, cargados: codigos.length, columna_costo: encab[idxCosto], columna_codigo: encab[idxCod] }));
    } catch(e) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // CATÁLOGO CON COSTOS — insumo del simulador de promociones (solo admin)
  if (urlPath === '/api/productos-costos' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    // Costos subidos por Excel (Contifico no los expone por API)
    const COSTOS = {}; let costosAt = null;
    try {
      const rc = await pool.query('SELECT codigo, costo, fuente, actualizado_at FROM producto_costos');
      rc.rows.forEach(x => { COSTOS[String(x.codigo).trim().toUpperCase()] = parseFloat(x.costo)||0; if (!costosAt || x.actualizado_at > costosAt) costosAt = x.actualizado_at; });
    } catch(e) {}
    const lista = Object.entries(catalogoProductos || {}).map(([id, p]) => ({
      id, codigo: p.codigo, nombre: p.nombre, marca: p.marca, categoria: p.categoria || '',
      costo: COSTOS[String(p.codigo||'').trim().toUpperCase()] || p.costo || 0,
      costo_campo: COSTOS[String(p.codigo||'').trim().toUpperCase()] ? 'excel' : (p.costo_campo || null),
      pvp1: p.pvp1 || 0, pvp2: p.pvp2 || 0, pvp3: p.pvp3 || 0, pvp4: p.pvp4 || 0,
      iva: p.iva, estado: p.estado || '', tipo: p.tipo || ''
    }));
    const conCosto = lista.filter(x => x.costo > 0).length;
    const campos = {};
    lista.forEach(x => { if (x.costo_campo) campos[x.costo_campo] = (campos[x.costo_campo]||0) + 1; });
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok:true, total: lista.length, con_costo: conCosto, campos_de_costo: campos, sincronizado: catalogoSyncedAt, costos_actualizado: costosAt, costos_cargados: Object.keys(COSTOS).length, productos: lista }));
    return;
  }

  // SONDEO: ficha cruda de un producto en Contifico, por todas las vías disponibles.
  // Sirve para descubrir cómo se llaman realmente los campos de costo y categoría.
  if (urlPath === '/api/producto-raw' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const q = String(urlObj.searchParams.get('codigo') || '').trim();
      const salida = {};
      const pedir = async (etiqueta, url) => {
        try {
          const r = await fetch(url, { headers:{'Authorization':API_KEY,'Accept':'application/json'} });
          const txt = await r.text();
          let d; try { d = JSON.parse(txt); } catch(e) { salida[etiqueta] = { http:r.status, crudo: txt.slice(0,300) }; return null; }
          const lista = Array.isArray(d) ? d : (d.results || (d.id ? [d] : []));
          salida[etiqueta] = { http:r.status, encontrados: lista.length, campos: lista.length ? Object.keys(lista[0]).sort() : [], ficha: lista[0] || null };
          return lista[0] || null;
        } catch(e) { salida[etiqueta] = { error: e.message }; return null; }
      };
      const enV2 = await pedir('v2_lista', `https://api.contifico.com/sistema/api/v2/producto/?codigo=${encodeURIComponent(q)}&page_size=3`);
      await pedir('v1_lista', `https://api.contifico.com/sistema/api/v1/producto/?codigo=${encodeURIComponent(q)}&page_size=3`);
      // El detalle por id suele traer bastante más que el listado
      const pid = enV2 && enV2.id;
      if (pid) {
        await pedir('v1_detalle', `https://api.contifico.com/sistema/api/v1/producto/${pid}/`);
        await pedir('v2_detalle', `https://api.contifico.com/sistema/api/v2/producto/${pid}/`);
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, codigo:q, resultados: salida }, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (urlPath === '/api/ventas-hoy' && req.method === 'GET') {
    // Si el caché tiene más de 2 minutos, refrescarlo antes de responder: es una consulta
    // corta y garantiza que una factura recién emitida se vea al abrir el panel.
    const edad = cache.ultima_sync ? (Date.now() - new Date(cache.ultima_sync).getTime()) : Infinity;
    if (edad > 2 * 60 * 1000 && !cache.sincronizando) {
      try { await sincronizarHoy(); } catch(e) { console.error('sync hoy on-demand:', e.message); }
    }
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ total: cache.documentos.length, ultima_sync: cache.ultima_sync, sincronizando: cache.sincronizando, documentos: cache.documentos, nc_documentos: cache.nc_documentos || [] }));
    return;
  }

  // VENTAS DE UNA FECHA ESPECÍFICA (histórico, hasta 1 semana atrás) — consulta directa
  // a Contifico igual que sincronizarHoy() pero para cualquier día solicitado. No usa el
  // caché de "hoy" porque ese se sobreescribe constantemente; cada llamada aquí trae el
  // detalle real de facturas de ese día puntual.
  if (urlPath === '/api/ventas-fecha' && req.method === 'GET') {
    try {
      const fechaParam = urlObj.searchParams.get('fecha'); // YYYY-MM-DD
      if(!fechaParam){
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'fecha es requerida (YYYY-MM-DD)'}));
        return;
      }

      // 1) Intentar leer desde la BD (historial guardado por sincronizarHoy cada hora) —
      // instantáneo y disponible para cualquier día ya sincronizado, sin pegarle a Contifico.
      const rDb = await pool.query(
        `SELECT documento_id, documento, cliente_nombre, vendedor_nombre, subtotal, total
         FROM facturas_detalle WHERE fecha=$1 ORDER BY id ASC`,
        [fechaParam]
      );

      if(rDb.rows.length > 0){
        const documentos = rDb.rows.map(row => ({
          id: row.documento_id,
          documento: row.documento,
          cliente_nombre: row.cliente_nombre,
          vendedor: { razon_social: row.vendedor_nombre },
          subtotal: parseFloat(row.subtotal),
          total: parseFloat(row.total)
        }));
        res.writeHead(200,{'Content-Type':'application/json'});
        res.end(JSON.stringify({ fecha: fechaParam, total: documentos.length, documentos, fuente: 'bd' }));
        return;
      }

      // 2) Fallback: no hay nada guardado en la BD para ese día (ej. antes de implementar
      // este historial, o un día que el servidor estuvo caído) — consultar Contifico en vivo.
      const [y,m,d] = fechaParam.split('-');
      const fechaEC = `${d}/${m}/${y}`; // formato que usa Contifico (igual que fmtDateEC)
      const url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fechaEC}&fecha_final=${fechaEC}&page_size=100`;
      let todos = [];
      let nextUrl = url;
      let paginas = 0;
      while (nextUrl && paginas < 20) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        const data = await resp.json();
        todos = todos.concat(data.results || []);
        nextUrl = data.next || null;
        paginas++;
      }
      const clientes = todos.filter(doc => doc.tipo_registro === 'CLI' && !doc.anulado && !esNotaCredito(doc) && !noEsVenta(doc));
      clientes.forEach(doc => {
        doc.cliente_nombre = doc.cliente?.razon_social || doc.cliente?.nombre_comercial || doc.persona_id || '—';
      });
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ fecha: fechaParam, total: clientes.length, documentos: clientes, fuente: 'contifico' }));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // BACKFILL: rellena facturas_detalle con el histórico de los últimos N días (por defecto 7)
  // consultando Contifico día por día. Se usa una sola vez para poblar los días anteriores
  // a que este historial existiera; después de eso sincronizarHoy() lo mantiene solo.
  if (urlPath === '/api/facturas-backfill' && req.method === 'GET') {
    try {
      const dias = parseInt(urlObj.searchParams.get('dias')) || 7;
      const resultado = [];
      const hoy = nowEC();
      for (let i = 0; i < dias; i++) {
        const d = new Date(hoy);
        d.setDate(d.getDate() - i);
        const fechaEC = fmtDateEC(d); // DD/MM/YYYY
        const fechaSQL = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

        const url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${fechaEC}&fecha_final=${fechaEC}&page_size=100`;
        let todos = [];
        let nextUrl = url;
        let paginas = 0;
        while (nextUrl && paginas < 20) {
          const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
          const data = await resp.json();
          todos = todos.concat(data.results || []);
          nextUrl = data.next || null;
          paginas++;
        }
        const clientes = todos.filter(doc => doc.tipo_registro === 'CLI' && !doc.anulado && !esNotaCredito(doc) && !noEsVenta(doc));

        for (const doc of clientes) {
          const cliNom = doc.cliente?.razon_social || doc.cliente?.nombre_comercial || doc.persona_id || '—';
          const vendNom = doc.vendedor?.razon_social || doc.vendedor?.nombre || 'Sin asignar';
          await pool.query(
            `INSERT INTO facturas_detalle(documento_id, fecha, documento, cliente_nombre, vendedor_nombre, subtotal, total, cedula_ruc)
             VALUES($1,$2,$3,$4,$5,$6,$7,$8)
             ON CONFLICT (documento_id, fecha) DO UPDATE SET
               documento=$3, cliente_nombre=$4, vendedor_nombre=$5, subtotal=$6, total=$7, cedula_ruc=$8, actualizado_at=NOW()`,
            [
              String(doc.id || doc.documento),
              fechaSQL,
              doc.documento || '',
              cliNom,
              vendNom,
              parseFloat(doc.subtotal || (doc.total/1.15) || 0),
              parseFloat(doc.total || 0),
              String(doc.cliente?.cedula || doc.cliente?.ruc || doc.cliente?.identificacion || '').replace(/\D/g,'') || null
            ]
          );
        }
        resultado.push({ fecha: fechaSQL, facturas: clientes.length });
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, dias_procesados: resultado }));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // VENTAS POR DÍA DE UN MES (para gráfico de líneas día 1 al último día del mes,
  // o hasta hoy si es el mes en curso). Acepta ?anio= y ?mes= opcionales; por defecto
  // usa el mes/año actuales del servidor (comportamiento original).
  // Lee de DATA_CACHE (ya mantenido por fusionarMesActualEnCache cada 15 min) — instantáneo,
  // sin pegarle a Contifico en vivo cada vez que alguien abre la pestaña Facturas.
  if (urlPath === '/api/ventas-mes-actual' && req.method === 'GET') {
    try {
      const ahora = nowEC();
      const anio = parseInt(urlObj.searchParams.get('anio')) || ahora.getFullYear();
      const mes = parseInt(urlObj.searchParams.get('mes')) || (ahora.getMonth() + 1); // 1-indexed, igual que frecuencia_dia
      const porDia = {}; // { dia: {total, subtotal} }
      const FUENTE = await dataCompleta();   // misma fuente única que /data.json
      Object.values(FUENTE||{}).forEach(clientes => {
        (clientes||[]).forEach(cli => {
          (cli.frecuencia_dia||[]).forEach(f => {
            if (f.anio !== anio || f.mes !== mes) return;
            if (!porDia[f.dia]) porDia[f.dia] = { total: 0, subtotal: 0 };
            porDia[f.dia].total += f.total;
            porDia[f.dia].subtotal += f.subtotal;
          });
        });
      });
      // Ya no hay nada que sumar aparte: dataCompleta() incluye el tramo pendiente.
      const diasArr = Object.keys(porDia).map(d=>parseInt(d)).sort((a,b)=>a-b).map(d=>({
        dia: d,
        total: Math.round(porDia[d].total*100)/100,
        subtotal: Math.round(porDia[d].subtotal*100)/100
      }));
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ anio, mes, dias: diasArr }));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // VENTAS POR DÍA POR MARCA (sin IVA) — sirve del caché precalculado
  if (urlPath === '/api/ventas-dia-marca' && req.method === 'GET') {
    try {
      const ahora = nowEC();
      const anio = parseInt(urlObj.searchParams.get('anio')) || ahora.getFullYear();
      const mes = parseInt(urlObj.searchParams.get('mes')) || (ahora.getMonth() + 1);
      const cacheKey = anio + '-' + mes;
      const esMesActual = (anio === ahora.getFullYear() && mes === ahora.getMonth() + 1);
      const hit = global._vdmCache[cacheKey];
      if (hit) {
        // Refrescar en segundo plano si está viejo (10 min mes actual, 6 h meses cerrados)
        const ttl = esMesActual ? 10 * 60 * 1000 : 6 * 60 * 60 * 1000;
        if (Date.now() - hit.ts > ttl) calcularVentasDiaMarca(anio, mes);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(hit.data)); return;
      }
      calcularVentasDiaMarca(anio, mes);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:false, pending:true }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
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
      const now = nowEC();
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
      const ff = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
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
          } else if(esNotaCredito(d)||d.tipo_registro==='NC'){
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

  // COMPARADOR DE DIAGNÓSTICO: app vs Contifico por vendedora/mes, clienta por clienta
  // ─── CUADRE CONTIFICO: explica documento por documento cualquier diferencia ───
  if (urlPath === '/api/cuadre' && req.method === 'GET') {
    try {
      const anioC = parseInt(urlObj.searchParams.get('anio')) || nowEC().getFullYear();
      const mesC = parseInt(urlObj.searchParams.get('mes')) || (nowEC().getMonth() + 1);
      const mmC = String(mesC).padStart(2,'0');
      const hoyEC = nowEC();
      const esMesActual = (anioC === hoyEC.getFullYear() && mesC === (hoyEC.getMonth()+1));
      const ultC = esMesActual ? hoyEC.getDate() : new Date(anioC, mesC, 0).getDate();
      const desdeC = `01/${mmC}/${anioC}`;
      const hastaC = `${String(ultC).padStart(2,'0')}/${mmC}/${anioC}`;

      const razones = {};
      const add = (k, doc, monto) => {
        if (!razones[k]) razones[k] = { docs: 0, total: 0, ejemplos: [] };
        razones[k].docs++; razones[k].total += monto;
        if (razones[k].ejemplos.length < 20) razones[k].ejemplos.push({ doc: doc.documento, tipo: doc.tipo_documento, fecha: doc.fecha_emision, cliente: (doc.cliente && (doc.cliente.razon_social||doc.cliente.nombre_comercial)) || '—', identificacion: (doc.cliente && (doc.cliente.ruc||doc.cliente.cedula)) || '', vendedor: (doc.vendedor && doc.vendedor.razon_social) || null, total: Math.round(parseFloat(doc.total||0)*100)/100 });
      };

      let nC = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desdeC}&fecha_final=${hastaC}&page_size=100`;
      let pgC = 0; const vistosC = new Set();
      let contTotal = 0, contFac = 0, contNC = 0, brutoFac = 0, subFac = 0, brutoNC = 0, subNC = 0;
      let aceptTotal = 0, aceptSub = 0, aceptDocs = 0;
      while (nC && pgC < 200) {
        const rC = await fetch(nC, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!rC.ok) { razones['error_api'] = { docs:0, total:0, ejemplos:[{ doc:'HTTP '+rC.status }] }; break; }
        const dC = await rC.json();
        (dC.results || []).forEach(d => {
          const tot = parseFloat(d.total || 0);
          const sub = parseFloat(d.subtotal || 0);
          const esNC = esNotaCredito(d);
          // Panorama de lo que hay en Contifico (facturas vivas)
          if (d.tipo_registro === 'CLI' && !noEsVenta(d) && !d.anulado) {
            contTotal++;
            if (esNC) { contNC++; brutoNC += tot; subNC += sub; } else { contFac++; brutoFac += tot; subFac += sub; }
          }
          // Ahora, las mismas reglas del pipeline, anotando por qué se descarta
          if (d.tipo_registro !== 'CLI') return add('proveedor_o_no_cliente', d, tot);
          if (d.anulado) return add('anulado', d, tot);
          if (noEsVenta(d)) return add(String(d.tipo_documento||'').toUpperCase()==='DAC' ? 'anticipo_de_cliente' : 'cotizacion_proforma_o_prefactura', d, tot);
          if (!d.vendedor && !d.vendedor_id && !d.vendedor_identificacion) return add('SIN_VENDEDOR_ASIGNADO', d, tot);
          const idCli = String((d.cliente && (d.cliente.ruc || d.cliente.cedula)) || '').trim();
          if (idCli === '1793143660001') return add('autoconsumo_cosetika', d, tot);
          const dk = d.id || d.documento;
          if (vistosC.has(dk)) return add('duplicado_en_paginacion', d, tot);
          vistosC.add(dk);
          const cid = (d.cliente && d.cliente.id) ? d.cliente.id : d.persona_id;
          if (!cid) return add('SIN_ID_DE_CLIENTE', d, tot);
          if (tot === 0) return add('total_cero', d, tot);
          const sg = esNC ? -1 : 1;
          aceptDocs++; aceptTotal += sg * tot; aceptSub += sg * sub;
        });
        nC = dC.next || null; pgC++;
      }

      // Lado app: exactamente lo que recibe el navegador (caché + tramo pendiente fusionado)
      let appSub = 0, appTot = 0, appCompras = 0;
      const FUENTE_APP = await dataCompleta();
      Object.values(FUENTE_APP || {}).forEach(clientes => {
        (clientes || []).forEach(c => {
          (c.frecuencia || []).forEach(f => {
            if (f.mes !== mesC) return;
            if (f.anio && f.anio !== anioC) return;
            appSub += (f.subtotal || 0); appTot += (f.total || 0); appCompras += (f.compras || 0);
          });
        });
      });

      const r2 = x => Math.round(x*100)/100;
      Object.keys(razones).forEach(k => { razones[k].total = r2(razones[k].total); });
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        ok: true, rango: { desde: desdeC, hasta: hastaC }, paginas_leidas: pgC,
        contifico_vivo: { facturas: contFac, notas_credito: contNC, documentos: contTotal,
          total_facturas_con_iva: r2(brutoFac), subtotal_facturas: r2(subFac),
          total_nc_con_iva: r2(brutoNC), subtotal_nc: r2(subNC),
          neto_con_iva: r2(brutoFac - brutoNC), neto_subtotal: r2(subFac - subNC) },
        pipeline_app: { documentos_aceptados: aceptDocs, neto_con_iva: r2(aceptTotal), neto_subtotal: r2(aceptSub) },
        lo_que_ve_la_app: { subtotal: r2(appSub), total_con_iva: r2(appTot), facturas: appCompras,
          ultimo_dia_en_cache: (function(){ const u = ultimoDiaEnCache(); return u ? fmtDateEC(u) : null; })() },
        diferencia_pipeline_vs_contifico: { con_iva: r2(aceptTotal - (brutoFac - brutoNC)), subtotal: r2(aceptSub - (subFac - subNC)) },
        DIFERENCIA_APP_VS_CONTIFICO: { subtotal: r2(appSub - (subFac - subNC)), con_iva: r2(appTot - (brutoFac - brutoNC - 0)), facturas: appCompras - contFac },
        cuadrado: Math.abs(appSub - (subFac - subNC)) < 0.5,
        descartados_por_razon: razones
      }, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false, error:e.message})); }
    return;
  }

  if (urlPath === '/api/debug-vendedora-mes' && req.method === 'GET') {
    try {
      const vendQ = (urlObj.searchParams.get('vendedora') || 'Liseth').toUpperCase();
      const anioQ = parseInt(urlObj.searchParams.get('anio')) || nowEC().getFullYear();
      const mesQ = parseInt(urlObj.searchParams.get('mes')) || (nowEC().getMonth() + 1);
      const mmQ = String(mesQ).padStart(2,'0');
      const ultQ = new Date(anioQ, mesQ, 0).getDate();
      // 1. Lado Contifico: todos los documentos del mes de esa vendedora
      const liveCli = {}; const docsVend = [];
      const vistosQ = new Set();
      let nextQ = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=01/${mmQ}/${anioQ}&fecha_final=${String(ultQ).padStart(2,'0')}/${mmQ}/${anioQ}&page_size=100`;
      let pagsQ = 0;
      while (nextQ && pagsQ < 60) {
        const respQ = await fetch(nextQ, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!respQ.ok) break;
        const dataQ = await respQ.json();
        (dataQ.results || []).forEach(d => {
          if (d.tipo_registro !== 'CLI') return;
          if (noEsVenta(d)) return;
          if (String(d.cliente?.ruc || d.cliente?.cedula || '').trim() === '1793143660001') return;
          const vN = String(d.vendedor?.razon_social || '').toUpperCase();
          if (!vN.includes(vendQ)) return;
          const dk = d.id || d.documento;
          if (vistosQ.has(dk)) return; vistosQ.add(dk);
          const signoQ = esNotaCredito(d) ? -1 : 1;
          const subQ = signoQ * parseFloat(d.subtotal || 0);
          const cliN = (d.cliente && (d.cliente.razon_social || d.cliente.nombre_comercial)) || '—';
          docsVend.push({ doc: d.documento, tipo: d.tipo_documento, fecha: d.fecha_emision, cliente: cliN, subtotal: Math.round(subQ*100)/100, anulado: !!d.anulado, estado: d.estado, autorizado_sri: d.autorizado_sri, firmado: d.firmado, electronico: d.electronico });
          if (d.anulado) return; // los anulados se listan arriba pero no suman
          if (!liveCli[cliN]) liveCli[cliN] = 0;
          liveCli[cliN] += subQ;
        });
        nextQ = dataQ.next || null;
        pagsQ++;
      }
      // 2. Lado app: DATA_CACHE de esa vendedora, frecuencia del mes (filtro tolerante, igual que los gráficos)
      const appCli = {};
      Object.entries(DATA_CACHE || {}).forEach(([vend, clientes]) => {
        if (!String(vend).toUpperCase().includes(vendQ)) return;
        (clientes || []).forEach(c => {
          (c.frecuencia || []).forEach(f => {
            if (f.mes !== mesQ) return;
            if (f.anio && f.anio !== anioQ) return;
            appCli[c.nombre] = (appCli[c.nombre] || 0) + (f.subtotal || 0);
          });
        });
      });
      // 3. Diferencias por clienta
      const nombres = new Set([...Object.keys(liveCli), ...Object.keys(appCli)]);
      const difs = [];
      nombres.forEach(n => {
        const lv = Math.round((liveCli[n] || 0)*100)/100;
        const ap = Math.round((appCli[n] || 0)*100)/100;
        if (Math.abs(lv - ap) > 0.5) difs.push({ cliente: n, contifico: lv, app: ap, diferencia: Math.round((ap - lv)*100)/100 });
      });
      difs.sort((a,b) => Math.abs(b.diferencia) - Math.abs(a.diferencia));
      const totLive = Math.round(Object.values(liveCli).reduce((a,b)=>a+b,0)*100)/100;
      const totApp = Math.round(Object.values(appCli).reduce((a,b)=>a+b,0)*100)/100;
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, vendedora: vendQ, mes: `${anioQ}-${mmQ}`,
        total_contifico: totLive, total_app: totApp, diferencia_total: Math.round((totApp-totLive)*100)/100,
        clientas_con_diferencia: difs,
        total_docs_validos: docsVend.filter(d=>!d.anulado).length,
        docs_sin_autorizacion_sri: docsVend.filter(d=>!d.anulado && d.autorizado_sri === false),
        docs_anulados_del_mes: docsVend.filter(d=>d.anulado),
        docs_del_mes: docsVend.filter(d=>!d.anulado).sort((a,b)=>String(a.doc).localeCompare(String(b.doc))).map(d=>`${d.doc} · ${d.fecha} · ${d.cliente} · $${d.subtotal}${d.tipo!=='FAC'?' ('+d.tipo+')':''}${d.autorizado_sri===false?' (SIN AUT. SRI)':''}`) }, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // DIAGNÓSTICO: sondear rutas de la API de Contifico que podrían exponer el token del POS
  if (urlPath === '/api/debug-contifico-pos' && req.method === 'GET') {
    try {
      const rutas = ['pos/', 'caja/', 'punto-venta/', 'punto_venta/', 'puntoventa/', 'punto-emision/', 'establecimiento/', 'empresa/', 'configuracion/'];
      const resultados = {};
      for (const ruta of rutas) {
        try {
          const rr = await fetch('https://api.contifico.com/sistema/api/v1/' + ruta, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
          const texto = await rr.text();
          resultados[ruta] = { status: rr.status, body: texto.substring(0, 600) };
        } catch(e) { resultados[ruta] = { error: e.message }; }
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(resultados, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // REGENERAR DATA.JSON
  if (urlPath === '/api/regenerar-data' && req.method === 'GET') {
    const desdeParam = urlObj.searchParams.get('desde');
    const anioActual = nowEC().getFullYear();
    const fi = desdeParam || '01/01/2022';
    const ff = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
    // El rango solicitado empieza en el año en curso (o después) → fusión segura, no toca
    // años anteriores. Si el rango pedido incluye años anteriores (ej. desde 2025), se
    // interpreta como intención deliberada de corregir histórico y se reemplaza todo el rango.
    // Siempre reemplazo completo — la fusión parcial quedó desactivada por duplicar datos
    const usarSoloAnioActual = false;
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({
      msg: usarSoloAnioActual
        ? `Regenerando año ${anioActual} (${fi} al ${ff}) — años anteriores no se tocan`
        : `Regenerando data.json del ${fi} al ${ff} (rango completo, reemplaza todo)`,
      ok: true
    }));
    generarDataJson(fi, ff).then(async data => {
      regenerandoEnProceso = true;
      if (usarSoloAnioActual && DATA_CACHE && Object.keys(DATA_CACHE).length > 0) {
        await fusionarAnioActualEnCache(anioActual, data);
        await guardarDataEnDB(DATA_CACHE);
        try { fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(DATA_CACHE, null, 2)); } catch(e) {}
        console.log(`✓ Regeneración (solo año ${anioActual}) completada — años anteriores intactos`);
      } else {
        await guardarDataEnDB(data);
        try { fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(data, null, 2)); } catch(e) {}
        console.log('✓ Regeneración completa (rango total) completada: ' + Object.keys(data).length + ' vendedoras');
      }
      regenerandoEnProceso = false;
    }).catch(e => { console.error('Error regenerar:', e.message); regenerandoEnProceso = false; });
    return;
  }

  // DATA.JSON desde caché en memoria (PostgreSQL)
  if (urlPath === '/data.json') {
    const completo = await dataCompleta();
    const cuerpo = JSON.stringify(completo || {});
    // Comprimir: el histórico son varios MB de JSON y el gzip lo reduce cerca de diez veces,
    // que es la mayor parte del tiempo de espera al abrir la app desde el móvil.
    if (/gzip/.test(req.headers['accept-encoding'] || '')) {
      zlib.gzip(cuerpo, (err, buf) => {
        if (err) { res.writeHead(200, {'Content-Type':'application/json'}); res.end(cuerpo); return; }
        res.writeHead(200, {'Content-Type':'application/json', 'Content-Encoding':'gzip', 'Vary':'Accept-Encoding'});
        res.end(buf);
      });
    } else {
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(cuerpo);
    }
    return;
  }

  // VER VENDEDORES EN CONTIFICO
  if (urlPath === '/api/ver-vendedores' && req.method === 'GET') {
    try {
      const desde = urlObj.searchParams.get('desde') || '01/01/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      const url = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      const resp = await fetch(url, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
      const data = await resp.json();
      const vendedores = {};
      (data.results||[]).filter(d=>d.tipo_registro==='CLI'&&!d.anulado&&!esNotaCredito(d)).forEach(d=>{
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

  // VER DESCRIPCIÓN DE FACTURAS DE FERNANDO A ASESORAS
  if (urlPath === '/api/ver-facturas-fernando-desc' && req.method === 'GET') {
    try {
      const desde = urlObj.searchParams.get('desde') || '01/01/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let encontrados = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while(nextUrl && paginas < 10) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        const data = await resp.json();
        const filtrados = (data.results||[]).filter(d => {
          const vendNom = (d.vendedor?.razon_social || '').toLowerCase();
          return vendNom.includes('fernando') && d.tipo_registro === 'CLI' && !d.anulado;
        }).map(d => ({
          documento: d.documento,
          tipo_doc: d.tipo_documento,
          fecha: d.fecha_emision,
          cliente: d.cliente?.razon_social,
          descripcion: d.descripcion,
          referencia: d.referencia,
          total: d.total,
          detalles: (d.detalles||[]).map(det=>det.producto_nombre)
        }));
        encontrados = encontrados.concat(filtrados);
        nextUrl = data.next || null;
        paginas++;
      }
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ total: encontrados.length, encontrados }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // VER FACTURAS DE UN VENDEDOR A UN CLIENTE ESPECÍFICO (rápido, usa caché)
  if (urlPath === '/api/ver-facturas-fernando-daniela' && req.method === 'GET') {
    try {
      const clientes = DATA_CACHE['Fernando Espíndola'] || [];
      const daniela = clientes.find(c => c.nombre.includes('Daniela Villegas'));
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({
        encontrado: !!daniela,
        detalle: daniela || null
      }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error:e.message}));
    }
    return;
  }

  // BUSCAR DOCUMENTO EXACTO POR NÚMERO (diagnóstico puntual)
  if (urlPath === '/api/ver-documento' && req.method === 'GET') {
    try {
      const numDoc = urlObj.searchParams.get('numero') || '';
      const desde = urlObj.searchParams.get('desde') || '01/06/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let encontrados = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while(nextUrl && paginas < 50) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        const data = await resp.json();
        const filtrados = (data.results||[]).filter(d => (d.documento||'').includes(numDoc));
        encontrados = encontrados.concat(filtrados.map(d => ({
          documento: d.documento,
          tipo_doc: d.tipo_documento,
          cliente: d.cliente?.razon_social,
          anulado: d.anulado,
          total: d.total,
          subtotal: d.subtotal,
          subtotal_0: d.subtotal_0,
          subtotal_12: d.subtotal_12,
          subtotal_15: d.subtotal_15,
          iva: d.iva,
          descuento: d.descuento,
          detalles: (d.detalles||[]).map(det=>({
            producto: det.producto_nombre,
            cantidad: det.cantidad,
            precio: det.precio,
            porcentaje_iva: det.porcentaje_iva,
            base_gravable: det.base_gravable,
            base_cero: det.base_cero,
            base_no_objeto: det.base_no_objeto
          })),
          campos_raiz_disponibles: Object.keys(d)
        })));
        nextUrl = data.next || null;
        paginas++;
        if(encontrados.length>0 && paginas>5) break; // ya encontramos, no seguir innecesariamente
      }
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ numero: numDoc, paginas_revisadas: paginas, encontrados }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // DIAGNÓSTICO TEMPORAL: comparar totales/conteos del dashboard vs Contifico para un rango
  if (urlPath === '/api/diagnostico-mes' && req.method === 'GET') {
    try {
      const desde = urlObj.searchParams.get('desde') || '01/06/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let todos = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while (nextUrl && paginas < 500) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!resp.ok) break;
        const data = await resp.json();
        todos = todos.concat(data.results || []);
        nextUrl = data.next || null;
        paginas++;
      }

      const porTipo = {}; // conteo y suma cruda por tipo_documento, sin filtrar nada
      let totalAnulados = 0, sumaAnulados = 0;
      todos.forEach(d => {
        const t = d.tipo_documento || '—';
        if (!porTipo[t]) porTipo[t] = { count: 0, total: 0, subtotal: 0 };
        porTipo[t].count++;
        porTipo[t].total += parseFloat(d.total || 0);
        porTipo[t].subtotal += parseFloat(d.subtotal || d.subtotal_12 || 0);
        if (d.anulado) { totalAnulados++; sumaAnulados += parseFloat(d.total||0); }
      });

      // Aplicar EXACTAMENTE el mismo filtro que generarDataJson
      const documentosVistos = new Set();
      let duplicados = 0;
      let sinVendedor = 0, sumaSinVendedor = 0;
      let cosetikaExcluidos = 0, sumaCosetika = 0;
      let usaronFallbackSubtotal12 = 0;
      const filtrados = todos.filter(d => {
        if (d.tipo_registro !== 'CLI') return false;
        if (d.anulado) return false;
        if (esNotaCredito(d)) return false;
        if (noEsVenta(d)) return false;
        if (!d.vendedor && !d.vendedor_id && !d.vendedor_identificacion) { sinVendedor++; sumaSinVendedor += parseFloat(d.total||0); return false; }
        const cliRuc = (d.cliente?.ruc || d.cliente?.cedula || '').trim();
        if (cliRuc === '1793143660001') { cosetikaExcluidos++; sumaCosetika += parseFloat(d.total||0); return false; }
        const docKey = d.id || d.documento;
        if (documentosVistos.has(docKey)) { duplicados++; return false; }
        documentosVistos.add(docKey);
        if (!d.subtotal && d.subtotal_12) usaronFallbackSubtotal12++;
        return true;
      });

      const sumaTotalFiltrado = filtrados.reduce((a,d)=>a+parseFloat(d.total||0),0);
      const sumaSubtotalFiltrado = filtrados.reduce((a,d)=>a+parseFloat(d.subtotal||d.subtotal_12||0),0);

      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        rango: { desde, hasta },
        total_documentos_crudos_api: todos.length,
        por_tipo_documento_crudo: porTipo,
        anulados_en_crudo: { count: totalAnulados, suma_total: Math.round(sumaAnulados*100)/100 },
        despues_de_filtros: {
          count: filtrados.length,
          suma_total_con_iva: Math.round(sumaTotalFiltrado*100)/100,
          suma_subtotal_sin_iva: Math.round(sumaSubtotalFiltrado*100)/100
        },
        excluidos_por_filtro: {
          duplicados_omitidos: duplicados,
          sin_vendedor: { count: sinVendedor, suma_total: Math.round(sumaSinVendedor*100)/100 },
          cosetika_autoconsumo: { count: cosetikaExcluidos, suma_total: Math.round(sumaCosetika*100)/100 }
        },
        documentos_que_usaron_fallback_subtotal_12: usaronFallbackSubtotal12
      }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }


  // DIAGNÓSTICO POR PRODUCTO: compara unidades crudas de Contifico vs. las que sobreviven
  // el filtro de generarDataJson (cantidad===0 || base===0), para un nombre de producto dado.
  if (urlPath === '/api/diagnostico-producto' && req.method === 'GET') {
    try {
      const nombreBuscado = (urlObj.searchParams.get('nombre') || '').toUpperCase().trim().replace(/\s+/g,' ');
      const desde = urlObj.searchParams.get('desde') || '01/01/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let todos = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while (nextUrl && paginas < 500) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!resp.ok) break;
        const data = await resp.json();
        todos = todos.concat(data.results || []);
        nextUrl = data.next || null;
        paginas++;
      }
      // Mismo filtro de documento que generarDataJson (sin excluir por base/cantidad todavía)
      const documentosVistos = new Set();
      const docsFiltrados = todos.filter(d => {
        if (d.tipo_registro !== 'CLI') return false;
        if (d.anulado) return false;
        if (esNotaCredito(d)) return false;
        if (noEsVenta(d)) return false;
        if (!d.vendedor && !d.vendedor_id && !d.vendedor_identificacion) return false;
        const cliRuc = (d.cliente?.ruc || d.cliente?.cedula || '').trim();
        if (cliRuc === '1793143660001') return false;
        const docKey = d.id || d.documento;
        if (documentosVistos.has(docKey)) return false;
        documentosVistos.add(docKey);
        return true;
      });

      // Resolver el/los producto_id correspondientes al nombre buscado, usando el MISMO
      // catálogo que usa generarDataJson (no el nombre crudo de la línea de detalle, que
      // puede venir vacío o distinto — el nombre real que se muestra en el dashboard sale
      // de catalogoProductos[producto_id].nombre).
      const idsCoincidentes = Object.entries(catalogoProductos)
        .filter(([id, info]) => {
          const nombreCat = (info.nombre||'').toUpperCase().trim().replace(/\s+/g,' ');
          return nombreCat.includes(nombreBuscado) || nombreBuscado.includes(nombreCat);
        })
        .map(([id, info]) => ({ id, nombre: info.nombre, marca: info.marca, codigo: info.codigo }));

      let cantidadTotalCruda = 0, cantidadConFiltroNuevo = 0, cantidadExcluidaPorCantidadCero = 0;
      let lineasCrudas = 0, lineasExcluidas = 0;
      const productIdsVistos = new Set();
      const ejemplosExcluidos = [];
      const porMesCrudo = {}, porMesFiltrado = {};
      const docsConEsteProducto = new Set();
      const idsBuscados = new Set(idsCoincidentes.map(x=>x.id));
      // Líneas que mencionan este nombre por texto (det.producto_nombre) pero cuyo producto_id
      // NO está en el catálogo resuelto arriba — estas se pierden silenciosamente en generarDataJson,
      // que también resuelve el nombre vía catalogoProductos[producto_id], no por texto crudo.
      let cantidadPorNombreSinIdEnCatalogo = 0;
      const ejemplosPorNombreSinIdEnCatalogo = [];
      docsFiltrados.forEach(doc => {
        const mes = parseInt((doc.fecha_emision || '').split('/')[1]) || 0;
        (doc.detalles || []).forEach(det => {
          const nombreDetNorm = (det.producto_nombre||'').toUpperCase().trim().replace(/\s+/g,' ');
          const coincidePorId = idsBuscados.has(det.producto_id);
          const coincidePorNombre = nombreDetNorm === nombreBuscado;
          if (!coincidePorId && coincidePorNombre) {
            cantidadPorNombreSinIdEnCatalogo += parseFloat(det.cantidad||0);
            if (ejemplosPorNombreSinIdEnCatalogo.length < 5) {
              ejemplosPorNombreSinIdEnCatalogo.push({ doc: doc.documento||doc.id, fecha: doc.fecha_emision, cantidad: det.cantidad, producto_id: det.producto_id||null, producto_nombre_crudo: det.producto_nombre });
            }
          }
          if (!coincidePorId) return;
          lineasCrudas++;
          const cantidad = parseFloat(det.cantidad || 0);
          const base = parseFloat(det.base_gravable || det.base_cero || 0);
          productIdsVistos.add(det.producto_id || '(sin id)');
          cantidadTotalCruda += cantidad;
          porMesCrudo[mes] = (porMesCrudo[mes]||0) + cantidad;
          docsConEsteProducto.add(doc.documento || doc.id);
          // Filtro ACTUAL (ya corregido): solo se excluye si no hay producto_id o cantidad===0
          if (!det.producto_id || cantidad === 0) {
            lineasExcluidas++;
            cantidadExcluidaPorCantidadCero += cantidad;
            if (ejemplosExcluidos.length < 8) {
              ejemplosExcluidos.push({ doc: doc.documento || doc.id, fecha: doc.fecha_emision, cantidad, base_gravable: det.base_gravable, base_cero: det.base_cero, producto_id: det.producto_id||null });
            }
          } else {
            cantidadConFiltroNuevo += cantidad;
            porMesFiltrado[mes] = (porMesFiltrado[mes]||0) + cantidad;
          }
        });
      });

      // Comparar contra lo que HOY tiene DATA_CACHE para este mismo producto (post-fusión),
      // para ver si la pérdida ocurre en generarDataJson o después (fusión incremental/anual).
      let cantidadEnCacheActual = 0;
      const anioConsulta = parseInt(desde.split('/')[2]) || new Date().getFullYear();
      Object.values(DATA_CACHE||{}).forEach(clientes=>{
        (clientes||[]).forEach(cli=>{
          (cli.productos_mes||[]).forEach(pm=>{
            const nombrePmNorm = (pm.nombre||'').toUpperCase().trim().replace(/\s+/g,' ');
            if (pm.anio===anioConsulta && nombrePmNorm===nombreBuscado) {
              cantidadEnCacheActual += pm.cantidad||0;
            }
          });
        });
      });

      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        producto_buscado: nombreBuscado,
        productos_en_catalogo_que_coinciden: idsCoincidentes,
        rango: { desde, hasta },
        producto_ids_distintos_encontrados: [...productIdsVistos],
        documentos_distintos_con_este_producto: docsConEsteProducto.size,
        lineas_de_detalle_encontradas: lineasCrudas,
        cantidad_total_cruda_sin_filtrar: cantidadTotalCruda,
        cantidad_que_SOBREVIVE_filtro_actual_en_vivo: cantidadConFiltroNuevo,
        cantidad_excluida_por_cantidad_cero_o_sin_id: cantidadExcluidaPorCantidadCero,
        lineas_excluidas_count: lineasExcluidas,
        ejemplos_de_lineas_excluidas: ejemplosExcluidos,
        cantidad_actualmente_en_DATA_CACHE_productos_mes: cantidadEnCacheActual,
        diferencia_entre_calculo_en_vivo_y_DATA_CACHE: Math.round((cantidadConFiltroNuevo - cantidadEnCacheActual)*100)/100,
        cantidad_con_nombre_coincidente_pero_SIN_id_en_catalogo: cantidadPorNombreSinIdEnCatalogo,
        ejemplos_nombre_coincidente_sin_id_en_catalogo: ejemplosPorNombreSinIdEnCatalogo,
        por_mes_cantidad_cruda: porMesCrudo,
        por_mes_cantidad_que_sobrevive_filtro: porMesFiltrado
      }, null, 2));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }


  // BUSCAR CLIENTE O VENDEDOR EN CONTIFICO
  if (urlPath === '/api/buscar-cliente' && req.method === 'GET') {
    try {
      const nombre = urlObj.searchParams.get('q') || 'cosetika';
      const desde = urlObj.searchParams.get('desde') || '01/06/2026';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      // Paginar para obtener más resultados
      let encontrados = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while(nextUrl && paginas < 30) {
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
  // DIAGNÓSTICO TEMPORAL: sumar lo que HAY GUARDADO en DATA_CACHE para un mes/año (sin llamar a Contifico)
  if (urlPath === '/api/diagnostico-cache' && req.method === 'GET') {
    const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
    const mes = parseInt(urlObj.searchParams.get('mes')) || (new Date().getMonth()+1);
    let totalConIva = 0, totalSinIva = 0, totalCompras = 0;
    const porVendedora = {};
    Object.entries(DATA_CACHE||{}).forEach(([vendNom, clientes]) => {
      let vConIva=0, vSinIva=0, vCompras=0;
      clientes.forEach(cli => {
        (cli.frecuencia||[]).forEach(f => {
          if (f.anio===anio && f.mes===mes) {
            vConIva += f.total||0; vSinIva += f.subtotal||0; vCompras += f.compras||0;
          }
        });
      });
      if (vConIva>0 || vSinIva>0) {
        porVendedora[vendNom] = { con_iva: Math.round(vConIva*100)/100, sin_iva: Math.round(vSinIva*100)/100, compras: vCompras };
        totalConIva += vConIva; totalSinIva += vSinIva; totalCompras += vCompras;
      }
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      mes, anio,
      total_en_cache: {
        con_iva: Math.round(totalConIva*100)/100,
        sin_iva: Math.round(totalSinIva*100)/100,
        compras: totalCompras
      },
      por_vendedora: porVendedora,
      cache_actualizado: DATA_CACHE_TS
    }, null, 2));
    return;
  }

  if (urlPath === '/api/data-status') {
    const muestra = {};
    Object.entries(DATA_CACHE||{}).slice(0,2).forEach(([v,clientes])=>{
      muestra[v] = {
        clientes: clientes.length,
        ejemplo_frecuencia: clientes[0]?.frecuencia?.slice(0,3) || []
      };
    });
    const anioActual = nowEC().getFullYear();
    let totalLineasProductosMes = 0, totalLineasProductosMesAnioActual = 0, clientesConProductosMes = 0;
    Object.values(DATA_CACHE||{}).forEach(clientes=>{
      (clientes||[]).forEach(cli=>{
        const pm = cli.productos_mes||[];
        if (pm.length>0) clientesConProductosMes++;
        totalLineasProductosMes += pm.length;
        totalLineasProductosMesAnioActual += pm.filter(x=>x.anio===anioActual).length;
      });
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      vendedoras: Object.keys(DATA_CACHE||{}).length,
      actualizado: DATA_CACHE_TS,
      fuente: DATA_CACHE && Object.keys(DATA_CACHE).length > 0 ? 'postgresql' : 'vacia',
      regenerando_en_proceso_AHORA: regenerandoEnProceso,
      total_lineas_productos_mes_TODOS_los_anios: totalLineasProductosMes,
      total_lineas_productos_mes_anio_actual: totalLineasProductosMesAnioActual,
      clientes_con_al_menos_una_linea_productos_mes: clientesConProductosMes,
      muestra_estructura: muestra
    }));
    return;
  }

  // SUBIR EXCEL DE INVENTARIO (multipart/form-data, campo 'file')
  if (urlPath === '/api/inventario/subir' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok:false, error: 'No se encontró el archivo en la solicitud (campo "file")' }));
        return;
      }
      const { fechaCorte, filasProducto } = parsearExcelInventario(archivo.buffer);
      const { productos, sinMatch } = resolverInventarioContraCatalogo(filasProducto);
      const data = { fecha_corte: fechaCorte, productos };
      await guardarInventarioEnDB(data);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        ok: true,
        fecha_corte: fechaCorte,
        productos_cargados: Object.keys(productos).length,
        productos_sin_match: sinMatch.length,
        ejemplos_sin_match: sinMatch.slice(0, 10)
      }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // SUBIR EXCEL DE PROVINCIAS POR CLIENTE (multipart/form-data, campo 'file')
  if (urlPath === '/api/provincias/subir' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok:false, error: 'No se encontró el archivo en la solicitud (campo "file")' }));
        return;
      }
      const { overrides, filasConProvincia, filasSinIdentificador, totalFilas, clientesPorVendedor } = parsearExcelProvincias(archivo.buffer);
      await guardarProvinciasOverrideEnDB(overrides);

      // Guardar el conteo de clientes por asesora como el acumulado del mes EN CURSO
      // (fecha real del servidor al momento de subir el Excel) — mismo patrón que
      // mercately_registros: cada subida reemplaza el acumulado de este mes, nunca de
      // meses anteriores ya cerrados.
      const ahora = nowEC();
      const anioActual = ahora.getFullYear(), mesActual = ahora.getMonth()+1;
      const asesorasGuardadas = [];
      for (const [asesora, cantidad] of Object.entries(clientesPorVendedor||{})) {
        await pool.query(`
          INSERT INTO contifico_clientes_registros (asesora, anio, mes, cantidad, actualizado_at) VALUES ($1, $2, $3, $4, NOW())
          ON CONFLICT (asesora, anio, mes) DO UPDATE SET cantidad = $4, actualizado_at = NOW()
        `, [asesora, anioActual, mesActual, cantidad]);
        asesorasGuardadas.push({ asesora, cantidad });
      }

      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        ok: true,
        clientes_cargados: filasConProvincia,
        filas_sin_identificador: filasSinIdentificador,
        total_filas_excel: totalFilas,
        clientes_por_asesora_guardados: asesorasGuardadas,
        mes_actualizado: `${mesActual}/${anioActual}`
      }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // TESTERS — registro de testers entregados a clientes
  // PERSONAS — directorio de clientes cargado desde el Excel mensual de Contifico
  // GET /api/personas/buscar?q=nombre_o_cedula → busca por nombre, cédula o RUC
  if (urlPath === '/api/personas/buscar' && req.method === 'GET') {
    try {
      const q = (urlObj.searchParams.get('q')||'').trim();
      if (!q || q.length < 2) {
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify([])); return;
      }
      const r = await pool.query(
        `SELECT razon_social, cedula, ruc, telefono, direccion, email, vendedor
         FROM personas
         WHERE LOWER(razon_social) LIKE LOWER($1) OR cedula LIKE $1 OR ruc LIKE $1
         LIMIT 5`,
        [`%${q}%`]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/personas/subir → carga masiva del Excel Personas.xls
  if (urlPath === '/api/personas/telefonos' && req.method === 'GET') {
    try {
      const r = await pool.query("SELECT cedula, ruc, telefono FROM personas WHERE telefono IS NOT NULL AND telefono <> ''");
      const mapa = {};
      r.rows.forEach(p => {
        [p.cedula, p.ruc].forEach(v => {
          const d = String(v||'').replace(/\D/g,'');
          if (d) { mapa[d] = p.telefono; if (d.length === 13) mapa[d.substring(0,10)] = p.telefono; }
        });
      });
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(mapa));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({})); }
    return;
  }
  if (urlPath === '/api/personas/subir' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró archivo'})); return;
      }
      // Parsear con xlsx (xlrd no disponible en Node — usamos xlsx que soporta .xls)
      const wb = XLSX.read(archivo.buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      // Buscar la fila que contiene 'Razón Social' — puede estar en índice 2 o variar
      let iHdr = -1;
      for(let i = 0; i < Math.min(10, filas.length); i++){
        const fila = filas[i] || [];
        if(fila.some(c => String(c||'').includes('Raz') && String(c||'').includes('Social'))){
          iHdr = i; break;
        }
      }
      if(iHdr === -1){
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró fila de encabezados con Razón Social'})); return;
      }
      const hdrs = (filas[iHdr]||[]).map(h => String(h||'').trim());
      const iRuc   = hdrs.findIndex(h => h === 'RUC');
      const iCed   = hdrs.findIndex(h => h.includes('dula'));
      const iNom   = hdrs.findIndex(h => h.includes('Raz') && h.includes('Social'));
      const iTel   = hdrs.findIndex(h => h.includes('fono'));
      const iDir   = hdrs.findIndex(h => h.includes('irecci'));
      const iEmail = hdrs.findIndex(h => h.toLowerCase().includes('email') || h.toLowerCase().includes('correo'));
      const iVend  = hdrs.findIndex(h => h.includes('Vendedor'));
      if(iNom === -1){
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró columna Razón Social'})); return;
      }
      // NO destructivo: la lista SOLO actualiza datos (provincia/dirección/teléfono).
      // La base histórica del conteo de clientas nuevas NUNCA se toca:
      //  - clientas ya existentes → se actualizan sus datos (conservan su origen)
      //  - clientas que no estaban → se insertan como 'excel_nuevo', que NO forma
      //    parte de la base del conteo (así siguen contando como nuevas)
      const rEx = await pool.query('SELECT id, cedula, ruc, razon_social FROM personas');
      const normP = x => String(x||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
      const porCedEx = {}; const porNomEx = {};
      rEx.rows.forEach(p => {
        [p.cedula, p.ruc].forEach(v => { const d = String(v||'').replace(/\D/g,''); if (d) { porCedEx[d] = p.id; if (d.length === 13) porCedEx[d.substring(0,10)] = p.id; } });
        const nN = normP(p.razon_social); if (nN && !porNomEx[nN]) porNomEx[nN] = p.id;
      });
      let insertados = 0, actualizados = 0;
      const datos = filas.slice(iHdr + 1).filter(r => r && r[iNom]);
      for (const r of datos) {
        const ced = String(r[iCed]||'').trim() || null;
        const ruc = String(r[iRuc]||'').trim() || null;
        const nom = String(r[iNom]||'').trim();
        const tel = String(r[iTel]||'').trim() || null;
        const dir = String(r[iDir]||'').trim() || null;
        const email = String(r[iEmail]||'').trim() || null;
        const vend = String(r[iVend]||'').trim() || null;
        const dCed = String(ced||'').replace(/\D/g,'');
        const dRuc = String(ruc||'').replace(/\D/g,'');
        let idEx = (dCed && porCedEx[dCed]) || (dRuc && porCedEx[dRuc]) || (dRuc.length === 13 && porCedEx[dRuc.substring(0,10)]) || porNomEx[normP(nom)] || null;
        if (idEx) {
          await pool.query(
            'UPDATE personas SET cedula=COALESCE($1,cedula), ruc=COALESCE($2,ruc), razon_social=$3, telefono=COALESCE($4,telefono), direccion=COALESCE($5,direccion), email=COALESCE($6,email), vendedor=COALESCE($7,vendedor) WHERE id=$8',
            [ced, ruc, nom, tel, dir, email, vend, idEx]);
          actualizados++;
        } else {
          await pool.query(
            "INSERT INTO personas(cedula,ruc,razon_social,telefono,direccion,email,vendedor,origen) VALUES($1,$2,$3,$4,$5,$6,$7,'excel_nuevo')",
            [ced, ruc, nom, tel, dir, email, vend]);
          insertados++;
        }
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, insertados, actualizados}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // GET /api/testers/resumen → clientes agrupados con conteo y última entrega (rápido)
  if (urlPath === '/api/testers/resumen' && req.method === 'GET') {
    try {
      const r = await pool.query(
        `SELECT cliente_id, MAX(cliente_nombre) AS cliente_nombre, MAX(cliente_cedula) AS cliente_cedula,
                COUNT(*) AS total,
                TO_CHAR(MAX(fecha_entrega), 'YYYY-MM-DD') AS ultima_entrega
         FROM testers
         GROUP BY cliente_id
         ORDER BY MAX(cliente_nombre)`
      );
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/testers/bulk → carga masiva en lotes de 200 filas por INSERT
  if (urlPath === '/api/testers/bulk' && req.method === 'POST') {
    try {
      const { registros, limpiar } = await bodyJSON(req);
      if (!Array.isArray(registros)) {
        res.writeHead(400,{'Content-Type':'application/json'});
        res.end(JSON.stringify({ok:false,error:'registros debe ser array'}));
        return;
      }
      if (limpiar) await pool.query('TRUNCATE TABLE testers RESTART IDENTITY');
      const LOTE = 200;
      let insertados = 0;
      for (let i = 0; i < registros.length; i += LOTE) {
        const lote = registros.slice(i, i + LOTE);
        const valores = [];
        const params = [];
        lote.forEach((r, j) => {
          const b = j * 6;
          valores.push(`($${b+1},$${b+2},$${b+3},$${b+4},$${b+5},$${b+6})`);
          params.push(
            r.clienteNombre||r.cliente_nombre||'',
            r.clienteNombre||r.cliente_nombre||'',
            r.categoria||null,
            r.producto||'',
            r.codigo||null,
            r.fechaEntrega||r.fecha_entrega||null
          );
        });
        await pool.query(
          `INSERT INTO testers(cliente_id,cliente_nombre,categoria,producto,codigo,fecha_entrega) VALUES ${valores.join(',')}`,
          params
        );
        insertados += lote.length;
      }
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:true, insertados}));
    } catch(e) {
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:false,error:e.message}));
    }
    return;
  }
  // GET /api/testers                      → todos los testers
  // GET /api/testers?clienteId=XXX        → testers de un cliente específico
  if (urlPath === '/api/testers' && req.method === 'GET') {
    try {
      const clienteId  = urlObj.searchParams.get('clienteId');
      const nombre     = urlObj.searchParams.get('nombre');
      const cedula     = urlObj.searchParams.get('cedula');
      let r;
      if (cedula) {
        // Búsqueda por cédula/RUC — más precisa que por nombre
        r = await pool.query(
          `SELECT * FROM testers WHERE cliente_cedula=$1 ORDER BY fecha_entrega DESC NULLS LAST`,
          [cedula]
        );
      } else if (clienteId || nombre) {
        const palabras = (nombre||clienteId||'')
          .toUpperCase().replace(/[^A-ZÁÉÍÓÚÑ ]/gi,'').split(' ').filter(p=>p.length>=4);
        if (palabras.length > 0) {
          const conds = palabras.map((_,i)=>`UPPER(cliente_id) LIKE $${i+2} OR UPPER(cliente_nombre) LIKE $${i+2}`);
          r = await pool.query(
            `SELECT * FROM testers WHERE cliente_id=$1 OR ${conds.join(' OR ')} ORDER BY fecha_entrega DESC NULLS LAST`,
            [clienteId||'', ...palabras.map(p=>`%${p}%`)]
          );
        } else {
          r = await pool.query(`SELECT * FROM testers WHERE cliente_id=$1 ORDER BY fecha_entrega DESC NULLS LAST`,[clienteId||'']);
        }
      } else {
        r = await pool.query('SELECT * FROM testers ORDER BY cliente_nombre, fecha_entrega DESC NULLS LAST');
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/testers → registrar un nuevo tester entregado
  if (urlPath === '/api/testers' && req.method === 'POST') {
    try {
      const { clienteId, clienteNombre, clienteCedula, categoria, producto, codigo, fechaEntrega } = await bodyJSON(req);
      if (!clienteId || !producto) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Faltan datos'})); return;
      }
      // Si viene con cédula, actualizar la cédula en todos los registros previos del cliente
      if (clienteCedula) {
        await pool.query(
          `UPDATE testers SET cliente_cedula=$1 WHERE (cliente_id=$2 OR LOWER(cliente_nombre)=LOWER($2)) AND cliente_cedula IS NULL`,
          [clienteCedula, clienteId]
        );
      }
      await pool.query(
        `INSERT INTO testers(cliente_id, cliente_nombre, cliente_cedula, categoria, producto, codigo, fecha_entrega) VALUES($1,$2,$3,$4,$5,$6,$7)`,
        [clienteId, clienteNombre||'', clienteCedula||null, categoria||null, producto, codigo||null, fechaEntrega||null]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/testers/enlazar -> enlaza testers huerfanos con un cliente real por cedula/RUC
  if (urlPath === '/api/testers/enlazar' && req.method === 'POST') {
    try {
      const { clienteId, clienteNombre, cedula } = await bodyJSON(req);
      const digits = String(cedula||'').replace(/\D/g,'');
      if ((digits.length !== 10 && digits.length !== 13) || (!clienteId && !clienteNombre)) {
        res.writeHead(400,{'Content-Type':'application/json'});
        res.end(JSON.stringify({ok:false,error:'Se requiere cliente y cedula (10 digitos) o RUC (13)'}));
        return;
      }
      // Candidatos: tal cual, base del RUC (10 primeros) o RUC derivado (cedula+001)
      const candidatos = [digits];
      if (digits.length === 13) candidatos.push(digits.substring(0,10));
      if (digits.length === 10) candidatos.push(digits + '001');
      const rp = await pool.query(
        'SELECT razon_social, cedula, ruc FROM personas WHERE cedula = ANY($1) OR ruc = ANY($1) LIMIT 1',
        [candidatos]
      );
      const persona = rp.rows[0] || null;
      const ru = await pool.query(
        'UPDATE testers SET cliente_cedula=$1 WHERE cliente_id=$2 OR LOWER(cliente_nombre)=LOWER($3)',
        [digits, clienteId||'', clienteNombre||clienteId||'']
      );
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:true, actualizados: ru.rowCount, persona}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // DELETE /api/testers/:id → eliminar un tester específico
  if (urlPath.startsWith('/api/testers/') && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM testers WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }


  // ─── DOCUMENTOS (PDFs guardados en PostgreSQL) ─────────────────────────────
  // GET /api/documentos → lista de documentos (sin el binario)
  if (urlPath === '/api/documentos' && req.method === 'GET') {
    try {
      const r = await pool.query(
        `SELECT id, nombre, tamano, subido_por, TO_CHAR(created_at,'DD/MM/YYYY HH24:MI') AS subido_el
         FROM documentos ORDER BY created_at DESC`
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/documentos?usuario=X → subir un PDF (multipart, campo 'file')
  if (urlPath === '/api/documentos' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo || !archivo.buffer || archivo.buffer.length === 0) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró archivo'})); return;
      }
      const esPdf = archivo.buffer.slice(0,5).toString('latin1').startsWith('%PDF') || /\.pdf$/i.test(archivo.filename||'');
      if (!esPdf) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Solo se permiten archivos PDF'})); return;
      }
      if (archivo.buffer.length > 20*1024*1024) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'El archivo supera el máximo de 20 MB'})); return;
      }
      const usuario = urlObj.searchParams.get('usuario') || '';
      const nombre = (archivo.filename || 'documento.pdf').substring(0, 490);
      const r = await pool.query(
        'INSERT INTO documentos(nombre,tamano,subido_por,archivo) VALUES($1,$2,$3,$4) RETURNING id',
        [nombre, archivo.buffer.length, usuario, archivo.buffer]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:r.rows[0].id}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // GET /api/documentos/:id/archivo → ver/descargar el PDF
  if (/^\/api\/documentos\/\d+\/archivo$/.test(urlPath) && req.method === 'GET') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const r = await pool.query('SELECT nombre, archivo FROM documentos WHERE id=$1', [id]);
      if (!r.rows.length) { res.writeHead(404,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'No existe'})); return; }
      const nombreAscii = String(r.rows[0].nombre||'documento.pdf').replace(/[^\x20-\x7E]/g,'_').replace(/["\\]/g,'');
      const forzarDescarga = urlObj.searchParams.get('descargar') === '1';
      res.writeHead(200,{
        'Content-Type':'application/pdf',
        'Content-Disposition':(forzarDescarga?'attachment':'inline')+'; filename="'+nombreAscii+'"',
        'Cache-Control':'no-cache'
      });
      res.end(r.rows[0].archivo);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // DELETE /api/documentos/:id → eliminar documento
  if (/^\/api\/documentos\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM documentos WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── NSOs (Notificaciones Sanitarias Obligatorias) ─────────────────────────
  // GET /api/nsos → lista completa
  if (urlPath === '/api/nsos' && req.method === 'GET') {
    try {
      const r = await pool.query(`SELECT id, marca, nombre, nso, cert_nombre, cert_bytes, etiqueta,
        (certificado IS NOT NULL) AS tiene_certificado,
        (etiqueta_pdf IS NOT NULL) AS tiene_etiqueta_pdf, etq_nombre, etq_bytes
        FROM nsos ORDER BY marca, nombre`);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // ─── ETIQUETA: comentario y PDF por producto ──────────────────────────────────
  // La etiqueta es otro documento distinto del certificado de NSO, así que vive aparte.
  if (/^\/api\/nsos\/\d+\/etiqueta$/.test(urlPath) && req.method === 'POST') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const b = await bodyJSON(req);
      await pool.query('UPDATE nsos SET etiqueta=$1 WHERE id=$2', [String(b.etiqueta||'').substring(0,2000), id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (/^\/api\/nsos\/\d+\/etiqueta-pdf$/.test(urlPath) && req.method === 'POST') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) throw new Error('No se encontró el archivo (campo "file")');
      const cab = archivo.buffer.slice(0, 1024).toString('latin1');
      if (!cab.includes('%PDF') && !/\.pdf$/i.test(archivo.filename||'')) throw new Error('El archivo no parece un PDF');
      await pool.query('UPDATE nsos SET etiqueta_pdf=$1, etq_nombre=$2, etq_bytes=$3, etq_subido_at=NOW() WHERE id=$4',
        [archivo.buffer, String(archivo.filename||'etiqueta.pdf').substring(0,390), archivo.buffer.length, id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, bytes: archivo.buffer.length}));
    } catch(e) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (/^\/api\/nsos\/\d+\/etiqueta-pdf$/.test(urlPath) && req.method === 'GET') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const r = await pool.query('SELECT etiqueta_pdf, etq_nombre, nombre FROM nsos WHERE id=$1', [id]);
      const f = r.rows[0];
      if (!f || !f.etiqueta_pdf) { res.writeHead(404,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Sin etiqueta'})); return; }
      res.writeHead(200, {
        'Content-Type':'application/pdf',
        'Content-Disposition': (urlObj.searchParams.get('descargar') ? 'attachment' : 'inline') +
          '; filename="' + String(f.etq_nombre || ('etiqueta - ' + (f.nombre||'') + '.pdf')).replace(/"/g,'') + '"',
        'Content-Length': f.etiqueta_pdf.length
      });
      res.end(f.etiqueta_pdf);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  if (/^\/api\/nsos\/\d+\/etiqueta-pdf$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      await pool.query('UPDATE nsos SET etiqueta_pdf=NULL, etq_nombre=NULL, etq_bytes=NULL, etq_subido_at=NULL WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ─── CERTIFICADOS DE NSO EN PDF ───────────────────────────────────────────────
  // Del nombre del archivo se deducen el producto y el código de NSO. El formato que usa
  // Cosétika es "NOMBRE DEL PRODUCTO - NSOC12345-26EC.pdf", pero también se acepta el
  // código en cualquier parte del nombre.
  function datosDelCertificado(nombreArchivo){
    const limpio = String(nombreArchivo||'').replace(/\.pdf$/i,'').trim();
    const mNso = /(NSOC?\s*[-–]?\s*\d{3,}\s*[-–]\s*\d{2}[A-Z]{2})/i.exec(limpio)
              || /(NSO[A-Z0-9\-–]{5,})/i.exec(limpio);
    const nso = mNso ? mNso[1].replace(/\s+/g,'').replace(/–/g,'-').toUpperCase() : '';
    let nombre = limpio;
    if (nso) {
      // Quitar el código y el separador que lo une al nombre
      nombre = limpio.replace(mNso[1], '').replace(/[\s\-–_]+$/,'').replace(/^[\s\-–_]+/,'').trim();
    }
    return { nombre: nombre || limpio, nso };
  }

  // Sube (o reemplaza) el certificado. Si el NSO no existe todavía, lo crea en la lista.
  if (urlPath === '/api/nsos/certificado' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) throw new Error('No se encontró el archivo (campo "file")');
      // La firma %PDF no siempre está en el byte 0: el formato permite basura previa y
      // algunos generadores dejan saltos de línea al inicio. Se busca en el primer KB, y
      // si no aparece pero la extensión es .pdf, se acepta igual.
      const cabecera = archivo.buffer.slice(0, 1024).toString('latin1');
      const pareceP = cabecera.includes('%PDF');
      const extP = /\.pdf$/i.test(archivo.filename || '');
      if (!pareceP && !extP) {
        throw new Error('El archivo no parece un PDF (empieza con "' +
          cabecera.slice(0,12).replace(/[^\x20-\x7e]/g,'·') + '")');
      }

      const det = datosDelCertificado(archivo.filename);
      const nsoParam = String(urlObj.searchParams.get('nso') || det.nso || '').trim().toUpperCase();
      const nombreParam = String(urlObj.searchParams.get('nombre') || det.nombre || '').trim();
      const marcaParam = String(urlObj.searchParams.get('marca') || '').trim();
      const idParam = parseInt(urlObj.searchParams.get('id')) || null;
      if (!nsoParam && !idParam) throw new Error('No se pudo leer el código de NSO del nombre del archivo. Renómbralo como "PRODUCTO - NSOC12345-26EC.pdf" o elige el producto de la lista.');

      // Buscar a qué registro pertenece: por id, o por código de NSO
      let fila = null;
      if (idParam) {
        const r = await pool.query('SELECT id, nombre, nso, marca FROM nsos WHERE id=$1', [idParam]);
        fila = r.rows[0] || null;
      }
      if (!fila && nsoParam) {
        const r = await pool.query("SELECT id, nombre, nso, marca FROM nsos WHERE UPPER(REPLACE(nso,' ','')) = $1", [nsoParam]);
        fila = r.rows[0] || null;
      }

      let creado = false;
      if (!fila) {
        // NSO nueva: entra a la lista con lo que dice el nombre del archivo
        const ins = await pool.query(
          'INSERT INTO nsos(marca, nombre, nso) VALUES($1,$2,$3) RETURNING id, nombre, nso, marca',
          [marcaParam || null, nombreParam || nsoParam, nsoParam]
        );
        fila = ins.rows[0]; creado = true;
      }

      await pool.query(
        'UPDATE nsos SET certificado=$1, cert_nombre=$2, cert_bytes=$3, cert_subido_at=NOW() WHERE id=$4',
        [archivo.buffer, String(archivo.filename||'certificado.pdf').substring(0,390), archivo.buffer.length, fila.id]
      );

      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, creado, id: fila.id, nombre: fila.nombre, nso: fila.nso,
        marca: fila.marca, bytes: archivo.buffer.length }));
    } catch(e) {
      res.writeHead(400,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // Ver el certificado en el navegador (o descargarlo con ?descargar=1)
  if (/^\/api\/nsos\/\d+\/certificado$/.test(urlPath) && req.method === 'GET') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const r = await pool.query('SELECT certificado, cert_nombre, nombre, nso FROM nsos WHERE id=$1', [id]);
      const f = r.rows[0];
      if (!f || !f.certificado) { res.writeHead(404,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:'Sin certificado'})); return; }
      const nombreArch = f.cert_nombre || ((f.nombre||'certificado') + ' - ' + (f.nso||'') + '.pdf');
      res.writeHead(200, {
        'Content-Type': 'application/pdf',
        'Content-Disposition': (urlObj.searchParams.get('descargar') ? 'attachment' : 'inline') + '; filename="' + nombreArch.replace(/"/g,'') + '"',
        'Content-Length': f.certificado.length
      });
      res.end(f.certificado);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // Quitar el certificado sin borrar el registro de NSO
  if (/^\/api\/nsos\/\d+\/certificado$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      await pool.query('UPDATE nsos SET certificado=NULL, cert_nombre=NULL, cert_bytes=NULL, cert_subido_at=NULL WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // POST /api/nsos/bulk → carga masiva {registros:[{marca,nombre,nso}], limpiar}
  if (urlPath === '/api/nsos/bulk' && req.method === 'POST') {
    try {
      const { registros, limpiar } = await bodyJSON(req);
      if (!Array.isArray(registros)) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'registros debe ser array'})); return;
      }
      if (limpiar) await pool.query('TRUNCATE TABLE nsos RESTART IDENTITY');
      let insertados = 0;
      const LOTE = 200;
      for (let i = 0; i < registros.length; i += LOTE) {
        const lote = registros.slice(i, i + LOTE);
        const valores = []; const params = [];
        lote.forEach((r, j) => {
          const b = j * 3;
          valores.push(`($${b+1},$${b+2},$${b+3})`);
          params.push(r.marca||'', r.nombre||'', r.nso||'');
        });
        await pool.query(`INSERT INTO nsos(marca,nombre,nso) VALUES ${valores.join(',')}`, params);
        insertados += lote.length;
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, insertados}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/nsos → agregar una NSO manualmente
  if (urlPath === '/api/nsos' && req.method === 'POST') {
    try {
      const { marca, nombre, nso } = await bodyJSON(req);
      if (!nombre || !String(nombre).trim()) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'El nombre del producto es obligatorio'})); return;
      }
      const r = await pool.query(
        'INSERT INTO nsos(marca,nombre,nso) VALUES($1,$2,$3) RETURNING id',
        [String(marca||'').trim().substring(0,90), String(nombre).trim().substring(0,490), String(nso||'').trim().substring(0,90)]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:r.rows[0].id}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // DELETE /api/nsos/:id → eliminar una NSO
  if (/^\/api\/nsos\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM nsos WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/nsos/subir → reemplaza todo con un Excel (multipart campo 'file')
  if (urlPath === '/api/nsos/subir' && req.method === 'POST') {
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontró archivo'})); return;
      }
      const wb = XLSX.read(archivo.buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      // Detectar fila de encabezado y columnas por nombre (tolera plantillas con o sin
      // encabezado en la columna de marca)
      let headerIdx = -1, nombreIdx = -1, nsoIdx = -1;
      for (let i = 0; i < Math.min(filas.length, 10); i++) {
        const f = filas[i] || [];
        const nI = f.findIndex(c => typeof c === 'string' && /nombre/i.test(c));
        const sI = f.findIndex(c => typeof c === 'string' && /nso/i.test(c));
        if (nI !== -1 && sI !== -1) { headerIdx = i; nombreIdx = nI; nsoIdx = sI; break; }
      }
      if (headerIdx === -1) { headerIdx = 0; nombreIdx = 2; nsoIdx = 3; }
      const marcaIdx = Math.max(0, nombreIdx - 1);
      const registros = [];
      for (let i = headerIdx + 1; i < filas.length; i++) {
        const f = filas[i] || [];
        const nombre = f[nombreIdx] != null ? String(f[nombreIdx]).trim() : '';
        if (!nombre) continue;
        registros.push({
          marca: f[marcaIdx] != null ? String(f[marcaIdx]).trim() : '',
          nombre,
          nso: f[nsoIdx] != null ? String(f[nsoIdx]).trim() : ''
        });
      }
      if (registros.length === 0) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'No se encontraron registros en el Excel'})); return;
      }
      // Conservar los certificados ya cargados: se vuelven a asociar por código de NSO
      const certs = await pool.query(`SELECT nso, certificado, cert_nombre, cert_bytes, cert_subido_at,
        etiqueta, etiqueta_pdf, etq_nombre, etq_bytes, etq_subido_at
        FROM nsos WHERE certificado IS NOT NULL OR etiqueta_pdf IS NOT NULL OR (etiqueta IS NOT NULL AND etiqueta <> '')`);
      await pool.query('TRUNCATE TABLE nsos RESTART IDENTITY');
      const LOTE = 200;
      for (let i = 0; i < registros.length; i += LOTE) {
        const lote = registros.slice(i, i + LOTE);
        const valores = []; const params = [];
        lote.forEach((r, j) => {
          const b = j * 3;
          valores.push(`($${b+1},$${b+2},$${b+3})`);
          params.push(r.marca, r.nombre, r.nso);
        });
        await pool.query(`INSERT INTO nsos(marca,nombre,nso) VALUES ${valores.join(',')}`, params);
      }
      // Devolver cada certificado a su NSO. Sin esto, subir el Excel consolidado borraba
      // todos los PDFs cargados, que es un trabajo que no se puede rehacer solo.
      let recuperados = 0;
      for (const c of certs.rows) {
        if (!c.nso) continue;
        const up = await pool.query(
          `UPDATE nsos SET certificado=$1, cert_nombre=$2, cert_bytes=$3, cert_subido_at=$4,
             etiqueta=$5, etiqueta_pdf=$6, etq_nombre=$7, etq_bytes=$8, etq_subido_at=$9
           WHERE UPPER(REPLACE(nso,' ','')) = $10`,
          [c.certificado, c.cert_nombre, c.cert_bytes, c.cert_subido_at,
           c.etiqueta, c.etiqueta_pdf, c.etq_nombre, c.etq_bytes, c.etq_subido_at,
           String(c.nso).replace(/\s/g,'').toUpperCase()]
        );
        recuperados += up.rowCount;
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, insertados: registros.length, certificados_conservados: recuperados}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // GET /api/nsos/descargar → genera el Excel consolidado
  if (urlPath === '/api/nsos/descargar' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT marca, nombre, nso FROM nsos ORDER BY marca, nombre');
      const filas = [['#','MARCA','Nombre (*)','NSO']];
      r.rows.forEach((row, i) => filas.push([i+1, row.marca, row.nombre, row.nso]));
      const ws = XLSX.utils.aoa_to_sheet(filas);
      ws['!cols'] = [{wch:5},{wch:14},{wch:50},{wch:18}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.writeHead(200,{
        'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':'attachment; filename="ETIQUETAS_NSOS_CONSOLIDADO.xlsx"',
        'Cache-Control':'no-cache'
      });
      res.end(buf);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // GET /api/sidebar-badges → conteo de pedidos/referidos nuevos desde una fecha dada
  if (urlPath === '/api/sidebar-badges' && req.method === 'GET') {
    try {
      const pd = urlObj.searchParams.get('pedidos_desde');
      const rd = urlObj.searchParams.get('referidos_desde');
      let pedidos = 0, referidos = 0;
      if (pd) {
        const r1 = await pool.query('SELECT COUNT(*) AS n FROM pedidos_web WHERE created_at > $1', [pd]);
        pedidos = parseInt(r1.rows[0].n) || 0;
      }
      if (rd) {
        const r2 = await pool.query('SELECT COUNT(*) AS n FROM referidos WHERE created_at > $1', [rd]);
        referidos = parseInt(r2.rows[0].n) || 0;
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({pedidos, referidos}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ─── PROVINCIAS VISIBLES EN CLIENTES → PROVINCIA (asignadas por asesora) ────
  if (urlPath === '/api/clientes-provincias' && req.method === 'GET') {
    try {
      const uid = parseInt(urlObj.searchParams.get('usuario_id')) || null;
      const r = uid
        ? await pool.query('SELECT usuario_id, provincia FROM clientes_provincias WHERE usuario_id=$1 ORDER BY provincia', [uid])
        : await pool.query('SELECT usuario_id, provincia FROM clientes_provincias ORDER BY usuario_id, provincia');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/clientes-provincias' && req.method === 'POST') {
    try {
      const { usuario_id, provincia } = await bodyJSON(req);
      if (!usuario_id || !provincia) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Faltan datos'})); return; }
      await pool.query('INSERT INTO clientes_provincias(usuario_id,provincia) VALUES($1,$2) ON CONFLICT DO NOTHING', [usuario_id, provincia]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/clientes-provincias' && req.method === 'DELETE') {
    try {
      const { usuario_id, provincia } = await bodyJSON(req);
      await pool.query('DELETE FROM clientes_provincias WHERE usuario_id=$1 AND provincia=$2', [usuario_id, provincia]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── BODEGAS (stock POS y Casa) ─────────────────────────────────────────────
  // GET /api/bodegas-stock → productos con cantidad por bodega + estado del sync
  if (urlPath === '/api/bodegas-stock' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT producto_id, codigo, nombre, marca, bodega, cantidad FROM stock_bodegas ORDER BY marca, nombre');
      const porProducto = {};
      r.rows.forEach(row => {
        if (!porProducto[row.producto_id]) porProducto[row.producto_id] = {
          producto_id: row.producto_id, codigo: row.codigo, nombre: row.nombre, marca: row.marca, pos: null, casa: null
        };
        const p = porProducto[row.producto_id];
        if (/pos/i.test(row.bodega)) p.pos = parseFloat(row.cantidad);
        else if (/casa/i.test(row.bodega)) p.casa = parseFloat(row.cantidad);
      });
      const productosFiltrados = filtrarProductosBodega(Object.values(porProducto));
      // Mínimo AUTOMÁTICO: rotación mensual × (días de provisión ÷ 30) para cada bodega
      const diasProv = parseInt(await getConfigApp('bodegas_dias_provision', '15')) || 15;
      let rotBod = {};
      try {
        const fcB = (INVENTARIO_CACHE && INVENTARIO_CACHE.fecha_corte) || new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'});
        rotBod = calcularRotacionMensual(fcB);
      } catch(e) {}
      productosFiltrados.forEach(pr => {
        const rot = rotBod[pr.producto_id] || 0;
        pr.rotacion = Math.round(rot*100)/100;
        pr.minimo = rot > 0 ? Math.ceil(rot * diasProv / 30) : 0;
      });
      const ultima = await getConfigApp('bodegas_ultima_sync', null);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ productos: productosFiltrados, ultima_sync: ultima, sync: BODEGAS_SYNC, dias_provision: diasProv }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/bodegas/config → días de provisión para el mínimo automático (5/10/15)
  if (urlPath === '/api/bodegas/config' && req.method === 'POST') {
    try {
      const { dias } = await bodyJSON(req);
      const d = parseInt(dias);
      if (![5,10,15].includes(d)) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Días inválidos (5, 10 o 15)'})); return; }
      await setConfigApp('bodegas_dias_provision', String(d));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, dias:d}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // PUT /api/bodegas/minimo → guardar el mínimo de un producto
  if (urlPath === '/api/bodegas/minimo' && req.method === 'PUT') {
    try {
      const { producto_id, minimo } = await bodyJSON(req);
      if (!producto_id) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Falta producto_id'})); return; }
      const m = Math.max(0, parseFloat(minimo) || 0);
      await pool.query(
        `INSERT INTO stock_minimos(producto_id, minimo, actualizado_at) VALUES($1,$2,NOW())
         ON CONFLICT (producto_id) DO UPDATE SET minimo=$2, actualizado_at=NOW()`,
        [String(producto_id).substring(0,29), m]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, minimo:m}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/bodegas/sync → dispara la sincronización en segundo plano
  if (urlPath === '/api/bodegas/sync' && req.method === 'POST') {
    try {
      if (BODEGAS_SYNC.enCurso) {
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, iniciado:false, sync:BODEGAS_SYNC})); return;
      }
      sincronizarBodegas().catch(e => console.error('Error sync bodegas manual:', e.message));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, iniciado:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // GET /api/bodegas-excel?marca=TODAS|BIOSKIN|... → Excel del stock (solo cantidades)
  if (urlPath === '/api/bodegas-excel' && req.method === 'GET') {
    try {
      const marcaSel = (urlObj.searchParams.get('marca') || 'TODAS').toUpperCase();
      const r = await pool.query('SELECT producto_id, codigo, nombre, marca, bodega, cantidad FROM stock_bodegas ORDER BY marca, nombre');
      const porProducto = {};
      r.rows.forEach(row => {
        if (!porProducto[row.producto_id]) porProducto[row.producto_id] = {
          producto_id: row.producto_id, codigo: row.codigo, nombre: row.nombre, marca: row.marca, pos: null, casa: null
        };
        const pr = porProducto[row.producto_id];
        if (/pos/i.test(row.bodega)) pr.pos = parseFloat(row.cantidad);
        else if (/casa/i.test(row.bodega)) pr.casa = parseFloat(row.cantidad);
      });
      let lista = filtrarProductosBodega(Object.values(porProducto));
      if (marcaSel !== 'TODAS') {
        lista = lista.filter(pr => String(pr.marca||'').toUpperCase() === marcaSel || String(pr.marca||'').toUpperCase().replace(/\s+/g,'') === marcaSel.replace(/\s+/g,''));
      }
      const diasProvX = parseInt(await getConfigApp('bodegas_dias_provision', '15')) || 15;
      let rotX = {};
      try {
        const fcX = (INVENTARIO_CACHE && INVENTARIO_CACHE.fecha_corte) || new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'});
        rotX = calcularRotacionMensual(fcX);
      } catch(e) {}
      const filas = [['Código','Producto','Marca','Rotación /mes',`Mínimo (${diasProvX}d)`,'Bodega POS','Bodega Casa']];
      lista.forEach(pr => {
        const rot = rotX[pr.producto_id] || 0;
        const minA = rot > 0 ? Math.ceil(rot * diasProvX / 30) : 0;
        filas.push([pr.codigo||'', pr.nombre||'', pr.marca||'', Math.round(rot*100)/100, minA, pr.pos!=null?pr.pos:'', pr.casa!=null?pr.casa:'']);
      });
      const ws = XLSX.utils.aoa_to_sheet(filas);
      ws['!cols'] = [{wch:12},{wch:55},{wch:14},{wch:12},{wch:12},{wch:12},{wch:12}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Stock');
      const buf = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
      const nombreArchivo = 'STOCK_BODEGAS' + (marcaSel!=='TODAS' ? '_'+marcaSel.replace(/[^A-Z0-9]/g,'_') : '') + '.xlsx';
      res.writeHead(200,{
        'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':'attachment; filename="'+nombreArchivo+'"',
        'Cache-Control':'no-cache'
      });
      res.end(buf);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  // ─── NOTAS DE KPIs CONTROL (una por panel y por mes) ────────────────────────
  if (urlPath === '/api/kpis-control-notas' && req.method === 'GET') {
    try {
      const mesK = urlObj.searchParams.get('mes') || '';
      const r = await pool.query('SELECT panel, texto FROM kpis_control_notas WHERE mes_key=$1', [mesK]);
      const notas = {};
      r.rows.forEach(x => { notas[x.panel] = x.texto || ''; });
      // Herencia: los paneles sin nota propia en este mes muestran la última nota escrita
      // en cualquier mes anterior (las reglas siguen vigentes hasta que se cambien)
      try {
        const rPrev = await pool.query(
          `SELECT DISTINCT ON (panel) panel, texto FROM kpis_control_notas
           WHERE mes_key < $1 AND texto <> '' ORDER BY panel, mes_key DESC`, [mesK]);
        rPrev.rows.forEach(x => { if (!notas[x.panel]) notas[x.panel] = x.texto || ''; });
      } catch(e) {}
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, notas }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/kpis-control-notas' && req.method === 'POST') {
    try {
      const { mes, panel, texto } = await bodyJSON(req);
      if (!mes || !panel) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Faltan datos'})); return; }
      await pool.query(
        `INSERT INTO kpis_control_notas(mes_key, panel, texto, actualizado_at) VALUES($1,$2,$3,NOW())
         ON CONFLICT (mes_key, panel) DO UPDATE SET texto=$3, actualizado_at=NOW()`,
        [String(mes).substring(0,7), String(panel).substring(0,29), String(texto||'').substring(0,2000)]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── VIÁTICOS: tarifario por provincia/ciudad (solo admin edita) ────────────
  if (urlPath === '/api/viaticos-tarifas' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT id, provincia, ciudad, googlemaps, dias, desayuno, almuerzo, cena, hotel, transporte, taxi FROM viaticos_tarifas ORDER BY provincia, ciudad');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/viaticos-tarifas' && req.method === 'POST') {
    try {
      const b = await bodyJSON(req);
      if (!String(b.provincia||'').trim() || !String(b.ciudad||'').trim()) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Provincia y ciudad son obligatorias'})); return;
      }
      const n = v => Math.max(0, parseFloat(v) || 0);
      const r = await pool.query(
        `INSERT INTO viaticos_tarifas(provincia,ciudad,googlemaps,dias,desayuno,almuerzo,cena,hotel,transporte,taxi)
         VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) RETURNING id`,
        [String(b.provincia).trim().substring(0,90), String(b.ciudad).trim().substring(0,190),
         n(b.googlemaps), n(b.dias), n(b.desayuno), n(b.almuerzo), n(b.cena), n(b.hotel), n(b.transporte), n(b.taxi)]
      );
      if (n(b.googlemaps) > 0) await pool.query('INSERT INTO googlemaps_historial(tarifa_id,valor) VALUES($1,$2)', [r.rows[0].id, n(b.googlemaps)]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:r.rows[0].id}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/viaticos-tarifas\/\d+$/.test(urlPath) && req.method === 'PUT') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const b = await bodyJSON(req);
      const CAMPOS = ['provincia','ciudad','googlemaps','dias','desayuno','almuerzo','cena','hotel','transporte','taxi'];
      const sets = []; const params = []; let i = 1;
      CAMPOS.forEach(c => {
        if (c in b) {
          sets.push(c+'=$'+(i++));
          params.push(['provincia','ciudad'].includes(c) ? String(b[c]).trim().substring(0,190) : Math.max(0, parseFloat(b[c])||0));
        }
      });
      if (sets.length === 0) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Sin campos'})); return; }
      // Historial GoogleMaps: registrar solo si el valor cambia
      let gmapsNuevo = null;
      if ('googlemaps' in b) {
        const nv = Math.max(0, parseFloat(b.googlemaps) || 0);
        const cur = await pool.query('SELECT googlemaps FROM viaticos_tarifas WHERE id=$1', [id]);
        if (cur.rows.length && parseFloat(cur.rows[0].googlemaps || 0) !== nv) gmapsNuevo = nv;
      }
      params.push(id);
      await pool.query('UPDATE viaticos_tarifas SET '+sets.join(', ')+' WHERE id=$'+i, params);
      if (gmapsNuevo !== null) await pool.query('INSERT INTO googlemaps_historial(tarifa_id,valor) VALUES($1,$2)', [id, gmapsNuevo]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/googlemaps-historial\/\d+$/.test(urlPath) && req.method === 'GET') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const r = await pool.query(
        "SELECT valor, TO_CHAR(fecha AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY HH24:MI') AS fecha, TO_CHAR(fecha AT TIME ZONE 'America/Guayaquil','DD/MM') AS fecha_corta FROM googlemaps_historial WHERE tarifa_id=$1 ORDER BY id ASC",
        [id]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (/^\/api\/viaticos-tarifas\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM viaticos_tarifas WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── TESTERS ENTREGADOS A ASESORAS (con histórico) ──────────────────────────
  if (urlPath === '/api/testers-asesoras' && req.method === 'GET') {
    try {
      const r = await pool.query(
        `SELECT id, producto_id, asesora, TO_CHAR(entregado_at,'YYYY-MM-DD') AS fecha
         FROM testers_asesoras ORDER BY entregado_at ASC, id ASC`);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/testers-asesoras' && req.method === 'POST') {
    try {
      const { producto_id, asesora } = await bodyJSON(req);
      if (!producto_id || !String(asesora||'').trim()) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Faltan datos'})); return;
      }
      const info = catalogoProductos[producto_id] || {};
      const hoyE = new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'});
      const r = await pool.query(
        `INSERT INTO testers_asesoras(producto_id, codigo, nombre, marca, asesora, entregado_at)
         VALUES($1,$2,$3,$4,$5,$6) RETURNING id, TO_CHAR(entregado_at,'YYYY-MM-DD') AS fecha`,
        [String(producto_id).substring(0,29), info.codigo||'', info.nombre||'', info.marca||'',
         String(asesora).trim().substring(0,250), hoyE]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:r.rows[0].id, fecha:r.rows[0].fecha}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/testers-asesoras\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM testers_asesoras WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── LOTES (caducidades con baja FIFO automática contra el stock real) ──────
  // GET /api/lotes → lotes por producto con restante FIFO, meses para agotar y semáforo
  if (urlPath === '/api/lotes' && req.method === 'GET') {
    try {
      const rl = await pool.query(
        `SELECT id, producto_id, codigo, nombre, marca, lote, TO_CHAR(fecha_caducidad,'YYYY-MM-DD') AS caduca, cantidad
         FROM lotes ORDER BY producto_id, fecha_caducidad ASC, id ASC`);
      // Stock actual por producto (POS + Casa)
      const rs = await pool.query('SELECT producto_id, cantidad FROM stock_bodegas');
      const stockPor = {};
      rs.rows.forEach(x => { stockPor[x.producto_id] = (stockPor[x.producto_id]||0) + (parseFloat(x.cantidad)||0); });
      // Rotación mensual (misma que Proyección)
      let rotacion = {};
      try {
        // Para caducidades el corte relevante es HOY, no la fecha del Excel de inventario:
        // se está decidiendo qué promocionar ahora.
        const fc = new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'});
        rotacion = calcularRotacionMensual(fc, {
          meses: parseInt(urlObj.searchParams.get('meses')) || 3,
          incluirMesActual: urlObj.searchParams.get('proyectar') === '1'
        });
      } catch(e) {}
      const hoyMs = new Date(new Date().toLocaleDateString('en-CA',{timeZone:'America/Guayaquil'}) + 'T12:00:00').getTime();
      // Agrupar lotes por producto
      const porLote = {};
      rl.rows.forEach(l => { (porLote[l.producto_id] = porLote[l.producto_id] || []).push(l); });
      // Base: TODOS los productos del catálogo de las 4 marcas (mismo filtro que Bodegas),
      // tengan o no lotes registrados
      const baseCatalogo = filtrarProductosBodega(
        Object.entries(catalogoProductos || {}).map(([id, info]) => ({
          producto_id: id, codigo: info.codigo||'', nombre: info.nombre||'', marca: info.marca||''
        }))
      );
      // Incluir también productos con lotes que ya no estén en el catálogo (no perder datos)
      const idsBase = new Set(baseCatalogo.map(b => b.producto_id));
      Object.keys(porLote).forEach(pid => {
        if (!idsBase.has(pid)) {
          const l0 = porLote[pid][0];
          baseCatalogo.push({ producto_id: pid, codigo: l0.codigo, nombre: l0.nombre, marca: l0.marca });
        }
      });
      const productos = baseCatalogo.map(base => {
        const pid = base.producto_id;
        const lts = porLote[pid] || [];
        const stock = Math.max(0, stockPor[pid] || 0);
        const rot = rotacion[pid] || 0;
        // FIFO: el stock actual se asigna del lote MÁS NUEVO hacia el más viejo;
        // lo que sobra tras llenar los nuevos es lo que queda del lote viejo
        let porAsignar = stock;
        const restantes = new Array(lts.length).fill(0);
        for (let i = lts.length - 1; i >= 0; i--) {
          const asig = Math.min(parseFloat(lts[i].cantidad)||0, porAsignar);
          restantes[i] = asig;
          porAsignar -= asig;
        }
        // porAsignar > 0 = hay stock sin lote registrado (se informa aparte)
        let acumulado = 0;
        const MARGEN_COMERCIAL_MESES = 4; // los productos deben venderse 4 meses antes de caducar
        const lotes = lts.map((l, i) => {
          const restante = restantes[i];
          acumulado += restante;
          const fechaCad = new Date(l.caduca + 'T12:00:00');
          const mesesCaducidad = (fechaCad.getTime() - hoyMs) / (30.44*24*60*60*1000);
          // Límite de venta: 4 meses antes de la caducidad
          const fechaLimite = new Date(fechaCad); fechaLimite.setMonth(fechaLimite.getMonth() - MARGEN_COMERCIAL_MESES);
          const mesesLimite = mesesCaducidad - MARGEN_COMERCIAL_MESES;
          const mesesAgotar = restante === 0 ? 0 : (rot > 0 ? acumulado / rot : 99);
          let estado = 'verde';
          if (restante === 0) estado = 'agotado';
          else if (mesesCaducidad < 0) estado = 'caducado';
          else if (mesesAgotar > mesesLimite) estado = 'rojo';       // no alcanza antes del límite de venta
          else if (mesesAgotar > mesesLimite * 0.75) estado = 'amarillo';
          return {
            id: l.id, lote: l.lote, caduca: l.caduca,
            limite_venta: fechaLimite.toISOString().substring(0,10),
            cantidad: parseFloat(l.cantidad)||0,
            restante: Math.round(restante*100)/100,
            meses_agotar: Math.round(mesesAgotar*10)/10,
            meses_caducidad: Math.round(mesesCaducidad*10)/10,
            meses_limite: Math.round(mesesLimite*10)/10,
            estado
          };
        });
        return {
          producto_id: pid, codigo: base.codigo, nombre: base.nombre, marca: base.marca,
          stock, rotacion: Math.round(rot*100)/100,
          sin_lote: lts.length ? Math.round(porAsignar*100)/100 : 0,
          lotes
        };
      }).sort((a,b) => String(a.nombre||'').localeCompare(String(b.nombre||'')));
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ productos, rotacion_meta: rotacion.__meta || null }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  function normalizarFechaCaducidad(f){
    const str = String(f||'').trim();
    if (/^\d{4}-\d{2}$/.test(str)) {
      const [a, m] = str.split('-').map(Number);
      const ultimoDia = new Date(a, m, 0).getDate();
      return `${a}-${String(m).padStart(2,'0')}-${ultimoDia}`;
    }
    return str;
  }
  // POST /api/lotes → registrar un lote
  if (urlPath === '/api/lotes' && req.method === 'POST') {
    try {
      const { producto_id, lote, fecha_caducidad, cantidad } = await bodyJSON(req);
      const cant = parseFloat(cantidad) || 0;
      if (!producto_id || !String(lote||'').trim() || !fecha_caducidad || cant <= 0) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Se requiere producto, número de lote, fecha de caducidad y cantidad mayor a 0'})); return;
      }
      const info = catalogoProductos[producto_id] || {};
      await pool.query(
        `INSERT INTO lotes(producto_id, codigo, nombre, marca, lote, fecha_caducidad, cantidad)
         VALUES($1,$2,$3,$4,$5,$6,$7)`,
        [String(producto_id).substring(0,29), info.codigo||'', info.nombre||'', info.marca||'',
         String(lote).trim().substring(0,90), normalizarFechaCaducidad(fecha_caducidad), cant]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // PUT /api/lotes/:id → editar lote (número, caducidad, cantidad — dinámico)
  if (/^\/api\/lotes\/\d+$/.test(urlPath) && req.method === 'PUT') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const body = await bodyJSON(req);
      const sets = []; const params = []; let i = 1;
      if ('lote' in body) {
        if (!String(body.lote||'').trim()) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'El número de lote es obligatorio'})); return; }
        sets.push('lote=$'+(i++)); params.push(String(body.lote).trim().substring(0,90));
      }
      if ('fecha_caducidad' in body && body.fecha_caducidad) { sets.push('fecha_caducidad=$'+(i++)); params.push(normalizarFechaCaducidad(body.fecha_caducidad)); }
      if ('cantidad' in body) {
        const c = parseFloat(body.cantidad)||0;
        if (c <= 0) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Cantidad inválida'})); return; }
        sets.push('cantidad=$'+(i++)); params.push(c);
      }
      if (sets.length === 0) { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Sin campos'})); return; }
      params.push(id);
      await pool.query('UPDATE lotes SET '+sets.join(', ')+' WHERE id=$'+i, params);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // DELETE /api/lotes/:id
  if (/^\/api\/lotes\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM lotes WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── CLIENTES REASIGNADOS (historial bajo otra vendedora) ───────────────────
  if (urlPath === '/api/clientes-reasignados' && req.method === 'GET') {
    try {
      const r = await pool.query('SELECT id, cliente_ruc, cliente_nombre, vendedora_destino FROM clientes_reasignados ORDER BY id DESC');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/clientes-reasignados' && req.method === 'POST') {
    try {
      const { cliente_ruc, cliente_nombre, vendedora_destino } = await bodyJSON(req);
      const digits = String(cliente_ruc||'').replace(/\D/g,'');
      if ((digits.length !== 10 && digits.length !== 13) || !vendedora_destino) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Se requiere cédula (10) o RUC (13) y vendedora destino'})); return;
      }
      await pool.query(
        `INSERT INTO clientes_reasignados(cliente_ruc, cliente_nombre, vendedora_destino) VALUES($1,$2,$3)
         ON CONFLICT (cliente_ruc) DO UPDATE SET cliente_nombre=$2, vendedora_destino=$3`,
        [digits, String(cliente_nombre||'').substring(0,490), String(vendedora_destino).substring(0,250)]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/clientes-reasignados\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM clientes_reasignados WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── INSTITUTOS (compras de alumnas) ────────────────────────────────────────
  // GET /api/institutos-compras?anio=&mes= → compras del mes de las alumnas con instituto
  if (urlPath === '/api/institutos-compras' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      const mes = parseInt(urlObj.searchParams.get('mes')) || (new Date().getMonth() + 1);
      const inicio = `${anio}-${String(mes).padStart(2,'0')}-01`;
      const ra = await pool.query("SELECT cedula, ruc, razon_social, instituto FROM personas WHERE instituto IS NOT NULL AND instituto<>''");
      const alumnas = ra.rows;
      const rf = await pool.query(
        `SELECT TO_CHAR(fecha,'YYYY-MM-DD') AS fecha_dia, documento, cliente_nombre, cedula_ruc, subtotal
         FROM facturas_detalle
         WHERE fecha >= $1::date AND fecha < ($1::date + INTERVAL '1 month')
         ORDER BY fecha DESC, id DESC`, [inicio]);
      const norm = x => String(x||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
      const porCed = {}; const porNombre = [];
      alumnas.forEach(a => {
        const ced = String(a.cedula||'').replace(/\D/g,'');
        const ruc = String(a.ruc||'').replace(/\D/g,'');
        if (ced) porCed[ced] = a;
        if (ruc) { porCed[ruc] = a; if (ruc.length === 13) porCed[ruc.substring(0,10)] = a; }
        const n = norm(a.razon_social);
        if (n) porNombre.push({ n, ancla: n.split(' ').slice(0,2).join(' '), a });
      });
      const compras = [];
      rf.rows.forEach(f => {
        let alum = null;
        const ced = String(f.cedula_ruc||'').replace(/\D/g,'');
        if (ced) alum = porCed[ced] || (ced.length === 13 ? porCed[ced.substring(0,10)] : null);
        if (!alum) {
          const n = norm(f.cliente_nombre);
          if (n) {
            const anclaF = n.split(' ').slice(0,2).join(' ');
            const m = porNombre.find(x => x.n === n || (anclaF.length >= 7 && (x.n.startsWith(anclaF) || n.startsWith(x.ancla))));
            if (m) alum = m.a;
          }
        }
        if (alum) compras.push({ fecha: f.fecha_dia, documento: f.documento, cliente: f.cliente_nombre, instituto: alum.instituto, subtotal: parseFloat(f.subtotal||0) });
      });
      const totales = {};
      compras.forEach(c => { if(!totales[c.instituto]) totales[c.instituto] = { total:0, compras:0 }; totales[c.instituto].total += c.subtotal; totales[c.instituto].compras++; });
      const campo = await getConfigApp('instituto_campo', 'adicional2_cliente');
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ alumnas: alumnas.length, compras, totales, campo, ultima_sync: INSTITUTOS_ULTIMA_SYNC }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/institutos/sync → sincronizar alumnas desde Contifico ahora
  if (urlPath === '/api/institutos/sync' && req.method === 'POST') {
    try {
      const resultado = await sincronizarInstitutos();
      res.writeHead(resultado.ok ? 200 : 500,{'Content-Type':'application/json'}); res.end(JSON.stringify(resultado));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/institutos/campo {campo} → configurar cuál adicional*_cliente es "Instituto"
  if (urlPath === '/api/institutos/campo' && req.method === 'POST') {
    try {
      const { campo } = await bodyJSON(req);
      if (!CAMPOS_ADICIONALES.includes(campo)) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Campo inválido'})); return;
      }
      await setConfigApp('instituto_campo', campo);
      const resultado = await sincronizarInstitutos();
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, ...resultado}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── ARTÍCULOS Y SUMINISTROS DE OFICINA (kardex simple) ──
  if (urlPath === '/api/articulos' && req.method === 'GET') {
    try {
      const r = await pool.query(`
        SELECT a.id, a.nombre, a.categoria,
          COALESCE(SUM(m.cantidad),0) AS stock,
          TO_CHAR(MAX(CASE WHEN m.tipo='ajuste' THEN m.fecha END) AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY') AS ultima_revision
        FROM articulos a
        LEFT JOIN articulos_movimientos m ON m.articulo_id = a.id
        WHERE a.activo = true
        GROUP BY a.id, a.nombre, a.categoria
        ORDER BY a.categoria, a.nombre`);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (urlPath === '/api/articulos' && req.method === 'POST') {
    try {
      const b = await bodyJSON(req);
      const nombre = String(b.nombre||'').trim().substring(0,290);
      if (!nombre) throw new Error('El nombre es obligatorio');
      const r = await pool.query('INSERT INTO articulos(nombre, categoria) VALUES($1,$2) RETURNING id',
        [nombre, String(b.categoria||'').trim().substring(0,90)]);
      const inicial = parseFloat(b.stock_inicial);
      if (!isNaN(inicial) && inicial > 0) {
        await pool.query("INSERT INTO articulos_movimientos(articulo_id,tipo,cantidad,nota,usuario) VALUES($1,'entrada',$2,'Stock inicial',$3)",
          [r.rows[0].id, inicial, String(b.usuario||'').substring(0,250)]);
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, id:r.rows[0].id}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/articulos\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('UPDATE articulos SET activo=false WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/articulos\/\d+\/movimientos$/.test(urlPath) && req.method === 'GET') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const r = await pool.query(
        "SELECT tipo, cantidad, nota, usuario, TO_CHAR(fecha AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY HH24:MI') AS fecha FROM articulos_movimientos WHERE articulo_id=$1 ORDER BY id DESC LIMIT 100", [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  if (/^\/api\/articulos\/\d+\/movimiento$/.test(urlPath) && req.method === 'POST') {
    try {
      const id = parseInt(urlPath.split('/')[3]);
      const b = await bodyJSON(req);
      const cant = parseFloat(b.cantidad);
      if (isNaN(cant) || cant <= 0) throw new Error('Cantidad inválida');
      await pool.query("INSERT INTO articulos_movimientos(articulo_id,tipo,cantidad,nota,usuario) VALUES($1,'entrada',$2,$3,$4)",
        [id, cant, String(b.nota||'').substring(0,490), String(b.usuario||'').substring(0,250)]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/articulos/revision' && req.method === 'POST') {
    try {
      const b = await bodyJSON(req);
      const conteos = b.conteos || {};
      const usuario = String(b.usuario||'').substring(0,250);
      let ajustes = 0;
      for (const [idStr, cantStr] of Object.entries(conteos)) {
        const id = parseInt(idStr);
        const conteo = parseFloat(cantStr);
        if (isNaN(id) || isNaN(conteo) || conteo < 0) continue;
        const rs = await pool.query('SELECT COALESCE(SUM(cantidad),0) AS stock FROM articulos_movimientos WHERE articulo_id=$1', [id]);
        const stock = parseFloat(rs.rows[0].stock) || 0;
        const dif = Math.round((conteo - stock) * 100) / 100;
        await pool.query("INSERT INTO articulos_movimientos(articulo_id,tipo,cantidad,nota,usuario) VALUES($1,'ajuste',$2,$3,$4)",
          [id, dif, `Revisión física: contado ${conteo}, teórico ${stock}` + (dif !== 0 ? ` (diferencia ${dif > 0 ? '+' : ''}${dif})` : ' (sin diferencia)'), usuario]);
        ajustes++;
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, ajustes}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/articulos-excel' && req.method === 'GET') {
    try {
      const r = await pool.query(`
        SELECT a.nombre, a.categoria, COALESCE(SUM(m.cantidad),0) AS stock,
          TO_CHAR(MAX(CASE WHEN m.tipo='ajuste' THEN m.fecha END) AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY') AS ultima_revision
        FROM articulos a LEFT JOIN articulos_movimientos m ON m.articulo_id = a.id
        WHERE a.activo = true GROUP BY a.id, a.nombre, a.categoria ORDER BY a.categoria, a.nombre`);
      const filas = [['Artículo','Categoría','Stock actual','Última revisión','Conteo físico']];
      r.rows.forEach(x => filas.push([x.nombre, x.categoria||'', parseFloat(x.stock), x.ultima_revision||'', '']));
      const ws = XLSX.utils.aoa_to_sheet(filas);
      ws['!cols'] = [{wch:45},{wch:18},{wch:12},{wch:14},{wch:14}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Articulos');
      const buf = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
      res.writeHead(200,{
        'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':'attachment; filename="ARTICULOS_OFICINA.xlsx"',
        'Cache-Control':'no-cache'
      });
      res.end(buf);
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }

  if (urlPath === '/api/pedidos-web/completar-cedulas' && req.method === 'POST') {
    try {
      const lim = parseInt(urlObj.searchParams.get('limite')) || 50;
      const r = await completarPedidosDesdeWoo(lim);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // Mapa ligero de cupos por cédula/RUC, para mostrar el crédito actual en la ficha
  // de cada clienta sin tener que consultar Contifico una por una.
  if (urlPath === '/api/creditos-mapa' && req.method === 'GET') {
    const mapa = {};
    Object.entries(CREDITO_CACHE || {}).forEach(([k, v]) => {
      if (v && v.cupo > 0) mapa[k] = { cupo: v.cupo, dias: v.dias || 0 };
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok:true, total: Object.keys(mapa).length, mapa }));
    return;
  }

  if (urlPath === '/api/creditos/sync') {
    sincronizarCreditos().catch(e=>console.error(e));
    res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, msg:'Sincronizando créditos desde Contifico'}));
    return;
  }
  // Diagnóstico de crédito: ?cedula=... o ?nombre=...
  if (urlPath === '/api/creditos/debug' && req.method === 'GET') {
    try {
      const ced = String(urlObj.searchParams.get('cedula')||'').replace(/\D/g,'');
      const nom = String(urlObj.searchParams.get('nombre')||'').toUpperCase().trim();
      if (urlObj.searchParams.get('sync') === '1' || !Object.keys(CREDITO_CACHE).length) { await sincronizarCreditos(); }
      const enCache = ced ? (CREDITO_CACHE[ced] || (ced.length===13?CREDITO_CACHE[ced.substring(0,10)]:null) || CREDITO_CACHE[ced+'001'] || null) : null;
      // Consultar en vivo a Contifico
      let vivo = [];
      if (ced) {
        for (const p of (ced.length===13 ? ['ruc','cedula'] : ['cedula','ruc'])) {
          try {
            const rr = await fetch(`https://api.contifico.com/sistema/api/v1/persona/?${p}=${ced}&page_size=5`, { headers:{'Authorization':API_KEY,'Accept':'application/json'} });
            const dd = await rr.json();
            const lista = Array.isArray(dd) ? dd : (dd.results||[]);
            lista.forEach(x => vivo.push({ id:x.id, razon_social:x.razon_social, cedula:x.cedula, ruc:x.ruc, cupo_credito:x.cupo_credito, dias_credito:x.dias_credito, aplicar_cupo:x.aplicar_cupo }));
            if (vivo.length) break;
          } catch(e) {}
        }
      }
      const porNombre = nom ? Object.entries(CREDITO_CACHE).filter(([k,v]) => String(v.nombre||'').toUpperCase().includes(nom)).slice(0,5).map(([k,v])=>({clave:k, ...v})) : [];
      const conCupo = Object.values(CREDITO_CACHE).filter(v=>v.cupo>0).length;
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, sincronizacion: CREDITO_SYNC_LOG, ultima_sync: CREDITO_SYNC_AT, personas_en_cache: Object.keys(CREDITO_CACHE).length,
        claves_con_cupo: conCupo, en_cache: enCache, en_contifico_ahora: vivo, coincidencias_por_nombre: porNombre }, null, 2));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── NÓMINA: subida de roles mensuales y consulta ──────────────────────────
  if (urlPath === '/api/nomina/subir' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) throw new Error('No se encontró archivo');
      const mesParam = urlObj.searchParams.get('mes') || '';
      const wb = XLSX.read(archivo.buffer, { type: 'buffer' });

      // ── ¿Es una planilla de UTILIDADES / pago extra? (hoja PLANILLA con "PAGO ACUMULADO")
      const hojaPlan = wb.SheetNames.find(n => n.toUpperCase().includes('PLANILLA'));
      if (hojaPlan && !wb.SheetNames.some(n => n.toUpperCase().includes('ROL'))) {
        const fp = XLSX.utils.sheet_to_json(wb.Sheets[hojaPlan], { header: 1, defval: null });
        let iH = -1;
        for (let i = 0; i < Math.min(10, fp.length); i++) {
          if ((fp[i]||[]).some(c => String(c||'').toUpperCase().includes('CEDULA') || String(c||'').toUpperCase().includes('CÉDULA'))) { iH = i; break; }
        }
        if (iH === -1) throw new Error('No se encontró la fila de encabezados en la planilla');
        const HP = (fp[iH]||[]).map(h => String(h||'').trim().toUpperCase());
        const idxCed = HP.findIndex(h => h.includes('CEDULA') || h.includes('CÉDULA'));
        const idxNom = HP.findIndex(h => h.includes('EMPLEADO') || h.includes('CARGO'));
        const idxTot = HP.findIndex(h => h.includes('ACUMULADO') || h.includes('TOTAL'));
        // Fecha de pago (busca "FECHA DE PAGO" en las primeras filas)
        let fechaPago = null;
        for (let i = 0; i < Math.min(10, fp.length); i++) {
          const fila = fp[i] || [];
          const j = fila.findIndex(c => String(c||'').toUpperCase().includes('FECHA DE PAGO'));
          if (j >= 0) { fechaPago = fila.slice(j+1).find(v => v); break; }
        }
        let mesPago = mesParam;
        if (!/^\d{4}-\d{2}$/.test(mesPago) && fechaPago) {
          let d = null;
          if (fechaPago instanceof Date) d = fechaPago;
          else if (typeof fechaPago === 'number' && fechaPago > 20000) {
            // Excel guarda las fechas como número de serie desde el 30/12/1899
            d = new Date(Date.UTC(1899, 11, 30) + fechaPago * 86400000);
          } else {
            const t = String(fechaPago).trim();
            const m1 = t.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/); // dd/mm/aaaa
            if (m1) {
              let anio = parseInt(m1[3]); if (anio < 100) anio += 2000;
              d = new Date(anio, parseInt(m1[2]) - 1, parseInt(m1[1]));
            } else {
              const dd = new Date(t);
              if (!isNaN(dd) && dd.getFullYear() > 2000) d = dd;
            }
          }
          if (d && !isNaN(d) && d.getFullYear() > 2000) mesPago = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
        }
        if (!/^\d{4}-\d{2}$/.test(mesPago)) throw new Error('No se pudo determinar el mes de pago de la planilla');
        const conceptoP = String(archivo.filename||'').toUpperCase().includes('DECIMO') ? 'Décimo cuarto' : 'Utilidades';
        // Limpiar cargas previas con fecha mal interpretada (ej. 1970)
        await pool.query("DELETE FROM nomina_extras WHERE mes_key < '2015-01'");
        const numP = v => { const x = parseFloat(v); return isNaN(x) ? 0 : x; };
        let n2 = 0;
        for (let i = iH + 1; i < fp.length; i++) {
          const f = fp[i] || [];
          const ced = String(f[idxCed] || '').replace(/\D/g,'');
          const val = numP(f[idxTot]);
          if (ced.length < 10 || !val) continue;
          await pool.query(
            `INSERT INTO nomina_extras(mes_key,cedula,empleado,concepto,valor) VALUES($1,$2,$3,$4,$5)
             ON CONFLICT (mes_key,cedula,concepto) DO UPDATE SET valor=$5, empleado=$3`,
            [mesPago, ced, String(f[idxNom]||'').substring(0,290), conceptoP, val]);
          n2++;
        }
        if (!n2) throw new Error('No se encontraron pagos en la planilla');
        res.writeHead(200,{'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok:true, tipo:'extra', concepto: conceptoP, mes: mesPago, empleados: n2 }));
        return;
      }

      const hojaRol = wb.SheetNames.find(n => n.toUpperCase().includes('ROL')) || wb.SheetNames[0];
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[hojaRol], { header: 1, defval: null });
      // Encabezados en la fila que contiene 'Empleado'
      let iHdr = -1;
      for (let i = 0; i < Math.min(8, filas.length); i++) {
        if ((filas[i]||[]).some(c => String(c||'').trim().toLowerCase() === 'empleado')) { iHdr = i; break; }
      }
      if (iHdr === -1) throw new Error('No se encontró la fila de encabezados (columna "Empleado")');
      const H = (filas[iHdr]||[]).map(h => String(h||'').trim().toLowerCase());
      const col = (...claves) => {
        for (const k of claves) {
          const idx = H.findIndex(h => h.includes(k));
          if (idx >= 0) return idx;
        }
        return -1;
      };
      const C = {
        cedula: col('cédula','cedula'), empleado: col('empleado'), cargo: col('cargo'),
        dias: col('días trabajados','dias trabajados'), sueldo: col('sueldo'),
        he100: col('horas extra 100'), he50: col('horas extra 50'),
        movApo: col('movilización aportable','movilizacion aportable'), comis: col('comisiones'),
        bonif: col('bonificaciones'), mov: col('movilización','movilizacion'), alim: col('alimentación','alimentacion'),
        d3m: col('décimo tercero men','decimo tercero men'), d4m: col('décimo cuarto men','decimo cuarto men'),
        frm: col('fondo de reserva men'), totIng: col('total ingresos'),
        iess: col('iess aporte'), ret: col('retencion renta','retención renta'),
        faltas: col('descuento faltas'), anticipo: col('anticipo sueldo'), consumos: col('descuentos consumos'),
        totEgr: col('total egresos'), totRec: col('total a recibir'),
        d3ac: col('decimo tercer sueldo acumulado','décimo tercer sueldo acumulado'),
        d4ac: col('decimo cuarto sueldo acumulado','décimo cuarto sueldo acumulado'),
        patron: col('aportes patronales'), ccc: col('valor ccc'), vac: col('vacaciones'), fr: col('fondos de reserva')
      };
      if (C.empleado < 0 || C.totRec < 0) throw new Error('El archivo no tiene el formato del rol (faltan columnas Empleado / Total a recibir)');
      // Cargos que se muestran siempre así, sin importar lo que diga el Excel
      const CARGOS_FIJOS = { '1722165089': 'JEFE DE MARCA / CAPACITADORA' };
      const num = (fila, idx) => { if (idx < 0) return 0; const v = parseFloat(fila[idx]); return isNaN(v) ? 0 : v; };
      const registros = [];
      for (let i = iHdr + 1; i < filas.length; i++) {
        const f = filas[i] || [];
        const nom = String(f[C.empleado] || '').trim();
        const ced = String(f[C.cedula] || '').trim();
        if (!nom || !ced) continue;                       // salta totales y filas vacías
        if (/^total/i.test(nom)) continue;
        const ingresos = num(f, C.totIng) || (num(f,C.sueldo)+num(f,C.he100)+num(f,C.he50)+num(f,C.movApo)+num(f,C.comis)+num(f,C.bonif)+num(f,C.mov)+num(f,C.alim)+num(f,C.d3m)+num(f,C.d4m)+num(f,C.frm));
        const patron = num(f, C.patron), ccc = num(f, C.ccc), vac = num(f, C.vac);
        const p3 = num(f, C.d3ac), p4 = num(f, C.d4ac), pfr = num(f, C.fr);
        const costo = Math.round((ingresos + patron + ccc + vac + p3 + p4 + pfr) * 100) / 100;
        registros.push({
          cedula: ced, empleado: nom, cargo: CARGOS_FIJOS[ced.replace(/\D/g,'')] || String(f[C.cargo]||'').trim(), dias: num(f,C.dias),
          sueldo: num(f,C.sueldo), horas_extra: num(f,C.he100)+num(f,C.he50),
          movilizacion: num(f,C.movApo)+num(f,C.mov)+num(f,C.alim), comisiones: num(f,C.comis),
          bonificaciones: num(f,C.bonif), decimo_tercero: num(f,C.d3m), decimo_cuarto: num(f,C.d4m),
          fondo_reserva: num(f,C.frm), total_ingresos: ingresos, iess_personal: num(f,C.iess),
          retencion: num(f,C.ret), anticipos: num(f,C.anticipo), otros_egresos: num(f,C.faltas)+num(f,C.consumos),
          total_egresos: num(f,C.totEgr), total_recibir: num(f,C.totRec),
          aportes_patronales: patron, valor_ccc: ccc, vacaciones: vac,
          prov_decimo_tercero: p3, prov_decimo_cuarto: p4, prov_fondo_reserva: pfr, costo_empresa: costo
        });
      }
      if (!registros.length) throw new Error('No se encontraron empleados en el archivo');
      // Mes: del parámetro o deducido del nombre del archivo
      let mesKey = mesParam;
      if (!/^\d{4}-\d{2}$/.test(mesKey)) {
        const MES_NOM = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
        const nombreArch = String(archivo.filename || '').toLowerCase();
        const mIdx = MES_NOM.findIndex(m => nombreArch.includes(m));
        const anioM = (nombreArch.match(/20\d{2}/) || [])[0] || (nombreArch.match(/\b(\d{2})\b/) ? '20'+nombreArch.match(/(\d{2})(?!.*\d)/)[1] : String(nowEC().getFullYear()));
        if (mIdx < 0) throw new Error('No se pudo deducir el mes del archivo — indícalo al subir');
        mesKey = `${anioM}-${String(mIdx+1).padStart(2,'0')}`;
      }
      await pool.query('DELETE FROM nomina_detalle WHERE mes_key=$1', [mesKey]);
      for (const r2 of registros) {
        await pool.query(
          `INSERT INTO nomina_detalle(mes_key,cedula,empleado,cargo,dias,sueldo,horas_extra,movilizacion,comisiones,bonificaciones,
             decimo_tercero,decimo_cuarto,fondo_reserva,total_ingresos,iess_personal,retencion,anticipos,otros_egresos,
             total_egresos,total_recibir,aportes_patronales,valor_ccc,vacaciones,prov_decimo_tercero,prov_decimo_cuarto,prov_fondo_reserva,costo_empresa)
           VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)`,
          [mesKey, r2.cedula, r2.empleado, r2.cargo, r2.dias, r2.sueldo, r2.horas_extra, r2.movilizacion, r2.comisiones, r2.bonificaciones,
           r2.decimo_tercero, r2.decimo_cuarto, r2.fondo_reserva, r2.total_ingresos, r2.iess_personal, r2.retencion, r2.anticipos, r2.otros_egresos,
           r2.total_egresos, r2.total_recibir, r2.aportes_patronales, r2.valor_ccc, r2.vacaciones, r2.prov_decimo_tercero, r2.prov_decimo_cuarto, r2.prov_fondo_reserva, r2.costo_empresa]);
      }
      const costoTotal = Math.round(registros.reduce((a,x)=>a+x.costo_empresa,0)*100)/100;
      await pool.query(
        `INSERT INTO nomina_meses(mes_key,archivo,subido_por,subido_at,empleados,costo_total) VALUES($1,$2,$3,NOW(),$4,$5)
         ON CONFLICT (mes_key) DO UPDATE SET archivo=$2, subido_por=$3, subido_at=NOW(), empleados=$4, costo_total=$5`,
        [mesKey, String(archivo.filename||'').substring(0,290), String(urlObj.searchParams.get('usuario')||'').substring(0,250), registros.length, costoTotal]);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, mes: mesKey, empleados: registros.length, costo_total: costoTotal }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/nomina' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      await pool.query("UPDATE nomina_detalle SET cargo='JEFE DE MARCA / CAPACITADORA' WHERE REGEXP_REPLACE(cedula,'\\D','','g')='1722165089' AND cargo <> 'JEFE DE MARCA / CAPACITADORA'");
      const r = await pool.query('SELECT * FROM nomina_detalle ORDER BY mes_key, empleado');
      const m = await pool.query("SELECT mes_key, archivo, empleados, costo_total, TO_CHAR(subido_at,'DD/MM/YYYY') AS subido FROM nomina_meses ORDER BY mes_key");
      const ex = await pool.query('SELECT mes_key, cedula, empleado, concepto, valor FROM nomina_extras ORDER BY mes_key');
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, detalle: r.rows, meses: m.rows, extras: ex.rows }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (/^\/api\/nomina\/\d{4}-\d{2}$/.test(urlPath) && req.method === 'DELETE') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const mk = urlPath.split('/').pop();
      await pool.query('DELETE FROM nomina_detalle WHERE mes_key=$1', [mk]);
      await pool.query('DELETE FROM nomina_meses WHERE mes_key=$1', [mk]);
      await pool.query('DELETE FROM nomina_extras WHERE mes_key=$1', [mk]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── VISITAS: semanas justificadas (vacaciones, etc.) cuentan como meta cumplida ──
  if (urlPath === '/api/visitas-excepciones' && req.method === 'GET') {
    try {
      const mesE = urlObj.searchParams.get('mes'); // YYYY-MM
      const r = mesE
        ? await pool.query("SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, motivo FROM visitas_excepciones WHERE TO_CHAR(semana,'YYYY-MM')=$1", [mesE])
        : await pool.query("SELECT asesora, TO_CHAR(semana,'YYYY-MM-DD') AS semana, motivo FROM visitas_excepciones");
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify([])); }
    return;
  }
  if (urlPath === '/api/visitas-excepciones' && req.method === 'POST') {
    try {
      const b = await bodyJSON(req);
      const ase = String(b.asesora||'').trim().substring(0,250);
      const sem = String(b.semana||'').substring(0,10);
      if (!ase || !sem) throw new Error('Faltan asesora o semana');
      if (b.activo === false) {
        await pool.query('DELETE FROM visitas_excepciones WHERE asesora=$1 AND semana=$2', [ase, sem]);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, activo:false})); return;
      }
      await pool.query(
        `INSERT INTO visitas_excepciones(asesora, semana, motivo, creado_por) VALUES($1,$2,$3,$4)
         ON CONFLICT (asesora, semana) DO UPDATE SET motivo=$3, creado_por=$4`,
        [ase, sem, String(b.motivo||'Vacaciones').substring(0,90), String(b.creado_por||'').substring(0,250)]
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, activo:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── SEGUIMIENTO: marcar clienta como contactada ──
  if (urlPath === '/api/seguimiento-contactos' && req.method === 'GET') {
    try {
      const r = await pool.query("SELECT cliente_key, cliente_nombre, asesora, comentario, contactado_at, TO_CHAR(contactado_at AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY') AS fecha FROM seguimiento_contactos");
      const mapa = {};
      r.rows.forEach(x => { mapa[x.cliente_key] = { fecha: x.contactado_at ? x.fecha : '', asesora: x.asesora, comentario: x.comentario || '', contactado: !!x.contactado_at }; });
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(mapa));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({})); }
    return;
  }
  if (urlPath === '/api/seguimiento-contactos' && req.method === 'POST') {
    try {
      const b = await bodyJSON(req);
      const key = String(b.cliente_key || '').substring(0,290);
      if (!key) throw new Error('Falta cliente_key');
      const nom = String(b.cliente_nombre||'').substring(0,490);
      const ase = String(b.asesora||'').substring(0,250);
      if ('comentario' in b && !('contactado' in b)) {
        // Solo guardar/actualizar el comentario (sin tocar la marca de contactada)
        await pool.query(
          `INSERT INTO seguimiento_contactos(cliente_key, cliente_nombre, asesora, comentario, contactado_at)
           VALUES($1,$2,$3,$4,NULL)
           ON CONFLICT (cliente_key) DO UPDATE SET comentario=$4, cliente_nombre=$2`,
          [key, nom, ase, String(b.comentario||'').substring(0,900)]
        );
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, comentario:String(b.comentario||'')})); return;
      }
      if (b.contactado === false) {
        // Desmarcar: se conserva el comentario
        await pool.query('UPDATE seguimiento_contactos SET contactado_at=NULL WHERE cliente_key=$1', [key]);
        res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, contactado:false})); return;
      }
      await pool.query(
        `INSERT INTO seguimiento_contactos(cliente_key, cliente_nombre, asesora, contactado_at)
         VALUES($1,$2,$3,NOW())
         ON CONFLICT (cliente_key) DO UPDATE SET cliente_nombre=$2, asesora=$3, contactado_at=NOW()`,
        [key, nom, ase]
      );
      const rf = await pool.query("SELECT TO_CHAR(contactado_at AT TIME ZONE 'America/Guayaquil','DD/MM/YYYY') AS fecha FROM seguimiento_contactos WHERE cliente_key=$1", [key]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, contactado:true, fecha: rf.rows[0]?.fecha || ''}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── ASIGNACIÓN DE PROVINCIAS PARA GIRAS (Visitas → Asignación) ──
  if (urlPath === '/api/asignacion-giras' && req.method === 'GET') {
    try {
      const raw = await getConfigApp('asignacion_giras', null);
      let asig = null;
      if (raw) { try { asig = JSON.parse(raw); } catch(e) {} }
      if (!asig) {
        asig = {
          'Liseth Gavilanes': ['Imbabura','Santo Domingo','Manabí'],
          'Daniela Villegas': ['Tungurahua','Chimborazo','Cotopaxi','Manabí'],
          'Karen Mora': ['Cantones de Guayas','Babahoyo'],
          'Nicole León': ['Azuay','El Oro','Santa Elena','Los Ríos']
        };
      }
      // Fusionar claves duplicadas ("Karen Mora" vs "Karen Rebeca Mora Cedeño") hacia el nombre del usuario
      try {
        const uR = await pool.query("SELECT nombre FROM usuarios WHERE rol IN ('asesora','jefa_ventas') AND activo=true");
        const nombresU = uR.rows.map(r => r.nombre);
        const normA = v => String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim().split(/\s+/).filter(Boolean);
        const mismoA = (a,b) => { const wa=normA(a), wb=normA(b); if(!wa.length||!wb.length||wa[0]!==wb[0]) return false; if(wa.length<2||wb.length<2) return true; return wa.slice(1).some(w=>wb.includes(w))||wb.slice(1).some(w=>wa.includes(w)); };
        const fusion = {}; let cambioA = false;
        Object.entries(asig).forEach(([k, provs]) => {
          const u = nombresU.find(n => mismoA(n, k));
          const destino = u || k;
          if (destino !== k) cambioA = true;
          if (!fusion[destino]) fusion[destino] = [];
          (provs||[]).forEach(p => {
            if (!fusion[destino].some(x => String(x).toUpperCase() === String(p).toUpperCase())) fusion[destino].push(p);
            else cambioA = true;
          });
        });
        asig = fusion;
        if (cambioA) await setConfigApp('asignacion_giras', JSON.stringify(asig));
      } catch(e) {}
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, asignacion: asig }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/asignacion-giras' && req.method === 'POST') {
    try {
      const { asignacion } = await bodyJSON(req);
      if (!asignacion || typeof asignacion !== 'object') { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Datos inválidos'})); return; }
      await setConfigApp('asignacion_giras', JSON.stringify(asignacion));
      // Sincronizar hacia Visitas → Provincias y Sectores (solo agrega, nunca quita)
      try {
        for (const [ase, provs] of Object.entries(asignacion)) {
          for (const p of (provs||[])) {
            await pool.query('INSERT INTO asesor_provincias(asesora,provincia) VALUES($1,$2) ON CONFLICT DO NOTHING',
              [String(ase).trim().substring(0,250), String(p).trim().substring(0,250)]);
          }
        }
      } catch(e) { console.log('Sync asesor_provincias:', e.message); }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── PRESUPUESTO DE VENTAS (Reportes → Presupuesto) ──
  if (urlPath === '/api/presupuesto-config' && req.method === 'GET') {
    // Lectura abierta a cualquier usuario con sesión: son metas de venta del equipo, no
    // información sensible, y el panel de objetivos por equipo las necesita. Escribirlas
    // (POST) sigue siendo exclusivo de admin.
    try {
      const raw = await getConfigApp('presupuesto_ventas', null);
      let cfg = null;
      if (raw) { try { cfg = JSON.parse(raw); } catch(e) {} }
      // Migración: la escalera de la nueva asesora ERAYBA vive bajo 'Nueva Asesora'
      // (el nombre real se pondrá renombrando a la usuaria cuando entre)
      if (cfg && cfg.overrides && cfg.overrides['Mayra Taipe']) {
        cfg.overrides['Nueva Asesora'] = Object.assign({}, cfg.overrides['Mayra Taipe'], cfg.overrides['Nueva Asesora'] || {});
        delete cfg.overrides['Mayra Taipe'];
        if (cfg.pct && 'Mayra Taipe' in cfg.pct) { delete cfg.pct['Mayra Taipe']; }
        await setConfigApp('presupuesto_ventas', JSON.stringify(cfg));
      }
      if (!cfg) {
        // Defaults iniciales: ultimátum de Nicole (ago 2026) + escalera de Mayra (sep 2026 → sep 2027)
        const escalera = {};
        let v = 4000, a = 2026, m = 9;
        for (let i = 0; i < 13; i++) {
          escalera[`${a}-${String(m).padStart(2,'0')}`] = v;
          v += 1000; m++; if (m > 12) { m = 1; a++; }
        }
        cfg = {
          pct: {},
          overrides: {
            'Nicole Yanira Leon Marquez': { '2026-08': 4000 },
            'Nueva Asesora': escalera
          },
          snapshots: {}
        };
      }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, config: cfg }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/presupuesto-config' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const { config } = await bodyJSON(req);
      if (!config || typeof config !== 'object') { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Config inválida'})); return; }
      await setConfigApp('presupuesto_ventas', JSON.stringify(config));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── PESOS DE KPIS para el pago de comisiones (Resumen del mes) ──
  if (urlPath === '/api/kpis-pesos' && req.method === 'GET') {
    try {
      const raw = await getConfigApp('kpis_pesos', null);
      let pesos = { visitas: 33.33, provincia: 33.33, casa: 33.33, nuevos: 0, seguimiento: 0 };
      if (raw) { try { pesos = Object.assign(pesos, JSON.parse(raw)); } catch(e) {} }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, pesos }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/kpis-pesos' && req.method === 'POST') {
    try {
      const { pesos } = await bodyJSON(req);
      if (!pesos || typeof pesos !== 'object') { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Pesos inválidos'})); return; }
      const limpio = {};
      ['visitas','provincia','casa','nuevos','seguimiento'].forEach(k => { limpio[k] = Math.max(0, Math.min(100, parseFloat(pesos[k]) || 0)); });
      await setConfigApp('kpis_pesos', JSON.stringify(limpio));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── META DE VENTAS por asesora (monto, mes objetivo y leyenda de recompensa) ──
  if (urlPath === '/api/meta-ventas' && req.method === 'GET') {
    try {
      const raw = await getConfigApp('meta_ventas', null);
      const porDefecto = {
        metas: {
          'Daniela Villegas Chamorro': 13500,
          'Giovanna Portilla': 17600,
          'Karen Rebeca Mora Cedeño': 8100,
          'Liseth Gavilanes': 14800,
          'Nicole Yanira Leon Marquez': 4949,
          'María Caridad': 3000
        },
        mes_meta: 9,
        leyenda: 'Recompensa: Zapatos OnCloud y $200 de bono'
      };
      let config = porDefecto;
      if (raw) { try { config = JSON.parse(raw); } catch(e) {} }
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ ok:true, config }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  if (urlPath === '/api/meta-ventas' && req.method === 'POST') {
    try {
      const { config } = await bodyJSON(req);
      if (!config || typeof config !== 'object') { res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Config inválida'})); return; }
      await setConfigApp('meta_ventas', JSON.stringify(config));
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── KPI CLIENTES NUEVOS (automático, sin Excel mensual) ────────────────────
  // Línea base = tabla personas (último Excel; las alumnas insertadas por el sync de
  // institutos NO cuentan como base — deben contar como nuevas para su asesora).
  // Nueva = cliente en ventas de Contifico cuya cédula/RUC no está en la base, contada
  // en el mes de su PRIMERA factura para la vendedora de esa primera factura.
  if (urlPath === '/api/kpi-clientes-nuevos' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      const mes = parseInt(urlObj.searchParams.get('mes')) || (new Date().getMonth() + 1);
      const rp = await pool.query("SELECT cedula, ruc, razon_social FROM personas WHERE origen IS DISTINCT FROM 'institutos' AND origen IS DISTINCT FROM 'excel_nuevo'");
      const normK = x => String(x||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
      const baseCed = new Set(); const baseNom = new Set();
      rp.rows.forEach(pr => {
        [pr.cedula, pr.ruc].forEach(v => {
          const d = String(v||'').replace(/\D/g,'');
          if (d) { baseCed.add(d); if (d.length === 13) baseCed.add(d.substring(0,10)); }
        });
        const n = normK(pr.razon_social); if (n) baseNom.add(n);
      });
      // Primera compra global por cliente en TODO el histórico de DATA
      const porCliente = {};
      Object.entries(DATA_CACHE || {}).forEach(([vend, clientes]) => {
        (clientes || []).forEach(c => {
          const ced = String(c.ruc||'').replace(/\D/g,'');
          const key = ced || ('nom:' + normK(c.nombre));
          (c.frecuencia || []).forEach(f => {
            if (!f.anio || !f.mes || !((f.compras||0) > 0 || (f.total||0) > 0)) return;
            const cur = porCliente[key];
            if (!cur || f.anio < cur.anio || (f.anio === cur.anio && f.mes < cur.mes)) {
              porCliente[key] = { anio: f.anio, mes: f.mes, vend, nombre: c.nombre, ced };
            }
          });
        });
      });
      const asesoras = {}; let total = 0;
      Object.values(porCliente).forEach(pc => {
        if (pc.anio !== anio || pc.mes !== mes) return;
        if (pc.ced) {
          if (baseCed.has(pc.ced) || (pc.ced.length === 13 && baseCed.has(pc.ced.substring(0,10)))) return;
        } else if (baseNom.has(normK(pc.nombre))) return;
        if (!asesoras[pc.vend]) asesoras[pc.vend] = { cantidad: 0, nombres: [] };
        asesoras[pc.vend].cantidad++;
        asesoras[pc.vend].nombres.push(pc.nombre);
        total++;
      });
      Object.values(asesoras).forEach(a => a.nombres.sort());
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, anio, mes, asesoras, total }));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  // ─── REFERIDOS ──────────────────────────────────────────────────────────────
  // GET /api/referidos → lista completa
  if (urlPath === '/api/referidos' && req.method === 'GET') {
    try {
      const r = await pool.query(
        `SELECT id, cliente, referido, telefono, TO_CHAR(fecha,'YYYY-MM-DD') AS fecha_dia,
                contactado, contactado_at, bono, bono_at, primera_compra, primera_compra_at
         FROM referidos ORDER BY fecha DESC NULLS LAST, id DESC`
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify(r.rows));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({error:e.message})); }
    return;
  }
  // POST /api/referidos/sync → resincronización manual (?historial=1 incluye ya leídos)
  if (urlPath === '/api/referidos/sync' && req.method === 'POST') {
    try {
      const incluirLeidos = urlObj.searchParams.get('historial') === '1';
      const resultado = await sincronizarReferidos({ incluirLeidos });
      res.writeHead(resultado.ok ? 200 : 500,{'Content-Type':'application/json'}); res.end(JSON.stringify(resultado));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // PUT /api/referidos/:id → actualizar contactado / bono / primera_compra (dinámico:
  // solo toca los campos enviados en el body)
  if (/^\/api\/referidos\/\d+$/.test(urlPath) && req.method === 'PUT') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      const body = await bodyJSON(req);
      const sets = []; const params = []; let i = 1;
      [['contactado','contactado_at'],['bono','bono_at'],['primera_compra','primera_compra_at']].forEach(([flag, ts])=>{
        if (flag in body) {
          sets.push(flag+'=$'+(i++)); params.push(!!body[flag]);
          sets.push(ts+'=$'+(i++)); params.push(body[ts]||null);
        }
      });
      if (sets.length === 0) {
        res.writeHead(400,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:'Sin campos para actualizar'})); return;
      }
      params.push(id);
      await pool.query('UPDATE referidos SET '+sets.join(', ')+' WHERE id=$'+i, params);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // POST /api/referidos/marcar-todos → marca todos los pendientes como contactados (histórico)
  if (urlPath === '/api/referidos/marcar-todos' && req.method === 'POST') {
    try {
      const r = await pool.query(
        "UPDATE referidos SET contactado=true, contactado_at='histórico' WHERE contactado=false OR contactado IS NULL"
      );
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true, actualizados: r.rowCount}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }
  // DELETE /api/referidos/:id
  if (/^\/api\/referidos\/\d+$/.test(urlPath) && req.method === 'DELETE') {
    try {
      const id = parseInt(urlPath.split('/').pop());
      await pool.query('DELETE FROM referidos WHERE id=$1', [id]);
      res.writeHead(200,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:true}));
    } catch(e) { res.writeHead(500,{'Content-Type':'application/json'}); res.end(JSON.stringify({ok:false,error:e.message})); }
    return;
  }

  if (urlPath === '/api/provincias/diagnostico-cliente' && req.method === 'GET') {
    try {
      const nombreBuscado = (urlObj.searchParams.get('nombre') || '').toUpperCase().trim();
      const desde = urlObj.searchParams.get('desde') || '01/01/2025';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let todos = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while (nextUrl && paginas < 500) {
        const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
        if (!resp.ok) break;
        const data = await resp.json();
        todos = todos.concat(data.results || []);
        nextUrl = data.next || null;
        paginas++;
      }
      const ejemplos = [];
      const rucsVistos = new Set();
      todos.forEach(d => {
        const cliNom = ((d.cliente && (d.cliente.razon_social || d.cliente.nombre_comercial)) || '').toUpperCase().trim();
        if (!cliNom.includes(nombreBuscado)) return;
        const cliRuc = (d.cliente && (d.cliente.ruc || d.cliente.cedula)) || '';
        const cliId = d.cliente && d.cliente.id ? d.cliente.id : d.persona_id;
        const key = cliRuc + '|' + cliId;
        if (rucsVistos.has(key)) return;
        rucsVistos.add(key);
        if (ejemplos.length < 5) {
          ejemplos.push({
            documento: d.documento || d.id,
            cliente_nombre: d.cliente?.razon_social || d.cliente?.nombre_comercial,
            cliente_ruc_crudo: d.cliente?.ruc,
            cliente_cedula_cruda: d.cliente?.cedula,
            ruc_o_cedula_usado: cliRuc,
            ruc_longitud: cliRuc.length,
            ruc_tiene_espacios: cliRuc !== cliRuc.trim(),
            cliente_id: cliId,
            existe_en_override: PROVINCIAS_OVERRIDE.hasOwnProperty(cliRuc.trim()),
            resultado_resolverProvinciaCliente: resolverProvinciaCliente(cliRuc, cliId, d.cliente?.direccion || '')
          });
        }
      });
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ nombre_buscado: nombreBuscado, documentos_encontrados: rucsVistos.size, ejemplos }, null, 2));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // DIAGNÓSTICO: comparar el total de una marca en vivo (recalculado desde Contifico)
  // contra lo que hay actualmente en DATA_CACHE (marcas_anio), para detectar pérdida
  // de datos en la fusión incremental/anual.
  if (urlPath === '/api/diagnostico-marca-total' && req.method === 'GET') {
    try {
      const marcaBuscada = (urlObj.searchParams.get('marca') || '').toUpperCase().trim();
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      const desde = urlObj.searchParams.get('desde') || `01/01/${anio}`;
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());

      // 1) Calcular EN VIVO desde Contifico (rehace generarDataJson para el rango)
      const dataEnVivo = await generarDataJson(desde, hasta);
      let totalEnVivo = 0;
      Object.values(dataEnVivo).forEach(clientes => {
        clientes.forEach(cli => {
          (cli.marcas_anio||[]).filter(x=>x.marca===marcaBuscada && x.anio===anio).forEach(x=>{ totalEnVivo += x.total; });
        });
      });

      // 2) Leer lo que HAY ACTUALMENTE en DATA_CACHE
      let totalEnCache = 0;
      let entradasConTotalCero = 0;
      let clientesConLaMarca = 0;
      Object.values(DATA_CACHE||{}).forEach(clientes => {
        (clientes||[]).forEach(cli => {
          const entradas = (cli.marcas_anio||[]).filter(x=>x.marca===marcaBuscada && x.anio===anio);
          if (entradas.length>0) clientesConLaMarca++;
          entradas.forEach(x=>{
            totalEnCache += x.total;
            if (x.total<=0) entradasConTotalCero++;
          });
        });
      });

      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({
        marca_buscada: marcaBuscada,
        anio,
        rango_consultado_en_vivo: { desde, hasta },
        total_EN_VIVO_recalculado_desde_Contifico: Math.round(totalEnVivo*100)/100,
        total_actual_en_DATA_CACHE: Math.round(totalEnCache*100)/100,
        diferencia: Math.round((totalEnVivo-totalEnCache)*100)/100,
        clientes_con_esta_marca_en_cache: clientesConLaMarca,
        entradas_con_total_cero_o_negativo_en_cache: entradasConTotalCero,
        data_cache_actualizado_en: DATA_CACHE_TS
      }, null, 2));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({error: e.message}));
    }
    return;
  }

  // DIAGNÓSTICO: probar resolverProvinciaCliente con un RUC/Cédula específico
  if (urlPath === '/api/provincias/diagnostico' && req.method === 'GET') {
    const identificador = (urlObj.searchParams.get('id')||'').trim();
    // Buscar al cliente real en DATA_CACHE por su RUC, para ver qué provincia tiene
    // GUARDADA ahí (lo que realmente usa el frontend), no solo lo que calcularía la función.
    const clientesEncontrados = [];
    Object.entries(DATA_CACHE||{}).forEach(([vendedora, clientes])=>{
      (clientes||[]).forEach(cli=>{
        if ((cli.ruc||'').trim() === identificador) {
          clientesEncontrados.push({ vendedora, nombre: cli.nombre, id: cli.id, ruc: cli.ruc, provincia_guardada_en_DATA_CACHE: cli.provincia });
        }
      });
    });
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      identificador_buscado: identificador,
      existe_en_override: PROVINCIAS_OVERRIDE.hasOwnProperty(identificador),
      valor_en_override: PROVINCIAS_OVERRIDE[identificador] || null,
      total_claves_en_override: Object.keys(PROVINCIAS_OVERRIDE).length,
      resultado_resolverProvinciaCliente: resolverProvinciaCliente(identificador, null, ''),
      clientes_encontrados_en_DATA_CACHE: clientesEncontrados,
      data_cache_actualizado_en: DATA_CACHE_TS,
      override_cache_ts: PROVINCIAS_OVERRIDE_TS
    }, null, 2));
    return;
  }

  // ESTADO DEL OVERRIDE DE PROVINCIAS
  if (urlPath === '/api/provincias/estado' && req.method === 'GET') {
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({
      ok: true,
      clientes_con_override: Object.keys(PROVINCIAS_OVERRIDE).length,
      actualizado_en: PROVINCIAS_OVERRIDE_TS
    }));
    return;
  }

  // SKU POR MARCA: GET para consultar, POST para guardar
  if (urlPath === '/api/sku-por-marca' && req.method === 'GET') {
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok: true, datos: SKU_POR_MARCA, actualizado_en: SKU_POR_MARCA_TS }));
    return;
  }
  if (urlPath === '/api/sku-por-marca' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const nuevo = {};
      ['BIOSKIN','ERAYBA','ZIAJA','ZIAJA PRO'].forEach(m => { nuevo[m] = parseInt(body[m]) || 0; });
      await guardarSkuPorMarcaEnDB(nuevo);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, datos: SKU_POR_MARCA }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // METAS DE MERCATELY: GET para consultar, POST para guardar (objeto completo {asesora: meta})
  if (urlPath === '/api/mercately/metas' && req.method === 'GET') {
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok: true, metas: MERCATELY_METAS, actualizado_en: MERCATELY_METAS_TS }));
    return;
  }
  if (urlPath === '/api/mercately/metas' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const nuevo = {};
      Object.keys(body).forEach(asesora => { nuevo[asesora] = parseInt(body[asesora]) || 0; });
      await guardarMercatelyMetasEnDB(nuevo);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, metas: MERCATELY_METAS }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // REGISTROS MENSUALES DE MERCATELY: acumulado total de clientes en Mercately a fin
  // de cada mes, por asesora. GET ?anio=2026 devuelve, además del acumulado, el "++"
  // (nuevos ese mes = acumulado actual - acumulado del mes anterior, consultando
  // diciembre del año previo si el mes es enero) y el DIF (++ menos la meta).
  // POST guarda/actualiza un registro puntual {asesora, anio, mes, cantidad=acumulado}.
  if (urlPath === '/api/mercately/registros' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      // Se trae también diciembre del año anterior, necesario para calcular el "++" de enero.
      const r = await pool.query(
        'SELECT asesora, anio, mes, cantidad FROM mercately_registros WHERE anio=$1 OR (anio=$2 AND mes=12)',
        [anio, anio-1]
      );
      const porAsesoraMes = {}; // "asesora|anio|mes" -> acumulado
      r.rows.forEach(row => { porAsesoraMes[row.asesora+'|'+row.anio+'|'+row.mes] = row.cantidad; });

      const registros = r.rows.filter(row => row.anio===anio).map(row => {
        const mesAnteriorAnio = row.mes===1 ? anio-1 : anio;
        const mesAnteriorMes = row.mes===1 ? 12 : row.mes-1;
        const acumuladoAnterior = porAsesoraMes[row.asesora+'|'+mesAnteriorAnio+'|'+mesAnteriorMes];
        const nuevos = (acumuladoAnterior!==undefined) ? (row.cantidad - acumuladoAnterior) : null;
        const meta = MERCATELY_METAS[row.asesora] || 0;
        const dif = (nuevos!==null) ? (nuevos - meta) : null;
        return { asesora: row.asesora, anio: row.anio, mes: row.mes, acumulado: row.cantidad, nuevos, dif };
      });
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, registros }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }
  if (urlPath === '/api/mercately/registros' && req.method === 'POST') {
    try {
      const { asesora, anio, mes, cantidad } = await bodyJSON(req);
      if (!asesora || !anio || !mes) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok: false, error: 'Faltan asesora, anio o mes' }));
        return;
      }
      await pool.query(`
        INSERT INTO mercately_registros (asesora, anio, mes, cantidad, actualizado_at) VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (asesora, anio, mes) DO UPDATE SET cantidad = $4, actualizado_at = NOW()
      `, [asesora, anio, mes, parseInt(cantidad)||0]);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // METAS DEL KPI "CLIENTES NUEVOS" (contado automáticamente desde el Excel de
  // Personas de Contifico) — reemplaza al antiguo cálculo manual de "Base mes
  // anterior / Cerrar mes". Mismo patrón que /api/mercately/metas.
  if (urlPath === '/api/contifico-clientes/metas' && req.method === 'GET') {
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok: true, metas: CONTIFICO_CLIENTES_METAS, actualizado_en: CONTIFICO_CLIENTES_METAS_TS }));
    return;
  }
  if (urlPath === '/api/contifico-clientes/metas' && req.method === 'POST') {
    try {
      const body = await bodyJSON(req);
      const nuevo = {};
      Object.keys(body).forEach(asesora => { nuevo[asesora] = parseInt(body[asesora]) || 0; });
      await guardarContificoClientesMetasEnDB(nuevo);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, metas: CONTIFICO_CLIENTES_METAS }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // REGISTROS MENSUALES DEL KPI DE CLIENTES (Contifico): acumulado total de clientes
  // asignados a cada asesora a fin de cada mes (se llena automáticamente al subir el
  // Excel de Personas, ver /api/provincias/subir). Mismo cálculo de ++ /DIF que Mercately.
  if (urlPath === '/api/contifico-clientes/registros' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      const r = await pool.query(
        'SELECT asesora, anio, mes, cantidad FROM contifico_clientes_registros WHERE anio=$1 OR (anio=$2 AND mes=12)',
        [anio, anio-1]
      );
      const porAsesoraMes = {};
      r.rows.forEach(row => { porAsesoraMes[row.asesora+'|'+row.anio+'|'+row.mes] = row.cantidad; });

      const registros = r.rows.filter(row => row.anio===anio).map(row => {
        const mesAnteriorAnio = row.mes===1 ? anio-1 : anio;
        const mesAnteriorMes = row.mes===1 ? 12 : row.mes-1;
        const acumuladoAnterior = porAsesoraMes[row.asesora+'|'+mesAnteriorAnio+'|'+mesAnteriorMes];
        const nuevos = (acumuladoAnterior!==undefined) ? (row.cantidad - acumuladoAnterior) : null;
        const meta = CONTIFICO_CLIENTES_METAS[row.asesora] || 0;
        const dif = (nuevos!==null) ? (nuevos - meta) : null;
        return { asesora: row.asesora, anio: row.anio, mes: row.mes, acumulado: row.cantidad, nuevos, dif };
      });
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, registros }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }
  // POST manual: permite ajustar un mes puntual a mano si hace falta corregir algo,
  // aparte de la actualización automática vía subida de Excel.
  if (urlPath === '/api/contifico-clientes/registros' && req.method === 'POST') {
    try {
      const { asesora, anio, mes, cantidad } = await bodyJSON(req);
      if (!asesora || !anio || !mes) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok: false, error: 'Faltan asesora, anio o mes' }));
        return;
      }
      await pool.query(`
        INSERT INTO contifico_clientes_registros (asesora, anio, mes, cantidad, actualizado_at) VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (asesora, anio, mes) DO UPDATE SET cantidad = $4, actualizado_at = NOW()
      `, [asesora, anio, mes, parseInt(cantidad)||0]);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // CASA ABIERTA: nombre de la estética registrada cada mes, por asesora (en el
  // servidor, no localStorage, para que se vea igual desde cualquier dispositivo).
  // GET ?anio=2026 devuelve el histórico del año completo (todas las asesoras).
  if (urlPath === '/api/casa-abierta/registros' && req.method === 'GET') {
    try {
      const anio = parseInt(urlObj.searchParams.get('anio')) || new Date().getFullYear();
      const r = await pool.query('SELECT asesora, anio, mes, nombre_estetica FROM casa_abierta_registros WHERE anio=$1', [anio]);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true, registros: r.rows }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }
  if (urlPath === '/api/casa-abierta/registros' && req.method === 'POST') {
    try {
      const { asesora, anio, mes, nombre_estetica } = await bodyJSON(req);
      if (!asesora || !anio || !mes || !nombre_estetica) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok: false, error: 'Faltan asesora, anio, mes o nombre_estetica' }));
        return;
      }
      await pool.query(`
        INSERT INTO casa_abierta_registros (asesora, anio, mes, nombre_estetica, actualizado_at) VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (asesora, anio, mes) DO UPDATE SET nombre_estetica = $4, actualizado_at = NOW()
      `, [asesora, anio, mes, String(nombre_estetica).trim()]);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: true }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }


  // CONSULTAR INVENTARIO POR MARCA: /api/inventario?marca=ZIAJA
  if (urlPath === '/api/inventario' && req.method === 'GET') {
    try {
      const marca = (urlObj.searchParams.get('marca')||'').toUpperCase().trim();
      if (!marca) {
        res.writeHead(400, {'Content-Type':'application/json'});
        res.end(JSON.stringify({ ok:false, error: 'Falta el parámetro marca' }));
        return;
      }
      const resultado = construirInventarioPorMarca(marca, {
        meses: parseInt(urlObj.searchParams.get('meses')) || 3,
        incluirMesActual: urlObj.searchParams.get('proyectar') === '1'
      });
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, marca, ...resultado }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // DESCARGAR BACKUP DIRECTO (sin correo, ya que Railway bloquea SMTP en este plan)
  // Estado y disparo manual del backup a Drive
  // RESTAURAR desde un archivo de respaldo (.json o .json.gz). Destructivo.
  if (urlPath === '/api/backup/restaurar' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    try {
      const buf = await bodyBuffer(req);
      const archivo = parseMultipartFile(buf, req.headers['content-type']);
      if (!archivo) throw new Error('No se encontró el archivo (campo "file")');
      if (String(urlObj.searchParams.get('confirmar')||'').toUpperCase() !== 'RESTAURAR') {
        throw new Error('Falta la confirmación: hay que escribir RESTAURAR para continuar');
      }

      // Descomprimir si viene en .gz
      let crudo = archivo.buffer;
      if (crudo[0] === 0x1f && crudo[1] === 0x8b) {
        crudo = await new Promise((ok, err) => zlib.gunzip(crudo, (e,b) => e ? err(e) : ok(b)));
      }
      let datos;
      try { datos = JSON.parse(crudo.toString('utf8')); }
      catch(e) { throw new Error('El archivo no es un backup válido (no se pudo leer como JSON)'); }
      if (!datos.tablas) throw new Error('El archivo no tiene la estructura de un backup de esta app');

      // Red de seguridad: respaldar el estado ACTUAL antes de sobrescribirlo
      let seguridad = null;
      if (b2Configurado() || driveConfigurado()) {
        try { const r = await respaldarAutomatico(); seguridad = r.archivo || null; } catch(e) {}
      }

      const tablasParam = String(urlObj.searchParams.get('tablas')||'').split(',').map(x=>x.trim()).filter(Boolean);
      const informe = await restaurarBackup(datos, { tablas: tablasParam });

      console.log(`⚠️ RESTAURACIÓN completada: ${informe.filas} filas en ${informe.restauradas.length} tablas`);
      res.writeHead(200,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, generado_en: datos.generado_en || null, respaldo_previo: seguridad, ...informe }));
    } catch(e) {
      res.writeHead(400,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  if (urlPath === '/api/backup/nube' && req.method === 'GET') {
    if (bloquearSiNoAdmin(req, res)) return;
    res.writeHead(200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok:true,
      destino: b2Configurado() ? 'Backblaze B2' : (driveConfigurado() ? 'Google Drive' : null),
      configurado: b2Configurado() || driveConfigurado(),
      conservar: b2Configurado() ? B2_CONSERVAR : GD_CONSERVAR, ...BACKUP_ESTADO }));
    return;
  }
  if (urlPath === '/api/backup/nube' && req.method === 'POST') {
    if (bloquearSiNoAdmin(req, res)) return;
    if (!b2Configurado() && !driveConfigurado()) {
      res.writeHead(400,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error:'No hay destino configurado. Faltan las variables B2_KEY_ID y B2_APP_KEY en Railway.' }));
      return;
    }
    const r = await respaldarAutomatico();
    res.writeHead(r.error?500:200,{'Content-Type':'application/json'});
    res.end(JSON.stringify({ ok: !r.error, ...r }));
    return;
  }

  if (urlPath === '/api/backup/descargar' && req.method === 'GET') {
    try {
      const backup = await generarBackupCompleto();
      const json = JSON.stringify(backup);
      const fechaStr = fmtDateEC(nowEC()).split('/').join('-');
      await registrarDescargaBackup();
      res.writeHead(200, {
        'Content-Type': 'application/json',
        'Content-Disposition': `attachment; filename="backup_cosetika_${fechaStr}.json"`
      });
      res.end(json);
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // ESTADO DEL BACKUP: cuándo fue la última descarga y si ya toca hacer una nueva
  if (urlPath === '/api/backup/estado' && req.method === 'GET') {
    try {
      const estado = await obtenerEstadoBackup();
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, ...estado }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // STATIC FILES

  let filePath = urlPath==='/' ? path.join(__dirname,'index.html')
    : urlPath==='/login' ? path.join(__dirname,'login.html')
    : urlPath==='/bot' ? path.join(__dirname,'bot.html')
    : urlPath==='/sofia.jpg' ? path.join(__dirname,'sofia.jpg')
    : path.join(__dirname, urlPath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const ext = path.extname(filePath);
    const headers = {'Content-Type': MIME[ext]||'text/plain'};
    // HTML siempre fresco — evita que el navegador sirva versiones viejas tras un deploy
    if(ext === '.html' || ext === '.js') headers['Cache-Control'] = 'no-cache, must-revalidate';
    res.writeHead(200, headers);
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200,{'Content-Type':'text/html','Cache-Control':'no-cache, must-revalidate'});
    fs.createReadStream(path.join(__dirname,'index.html')).pipe(res);
  }
});

server.listen(PORT,'0.0.0.0',()=>console.log(`Cosétika Dashboard running on port ${PORT}`));
