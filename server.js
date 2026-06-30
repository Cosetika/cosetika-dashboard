const http = require('http');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');
const XLSX = require('xlsx');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';
// Credenciales de la casilla pedidos@cosetika.com — configuradas como variables de
// entorno en Railway, nunca hardcodeadas en el código.
const PEDIDOS_EMAIL_HOST = process.env.PEDIDOS_EMAIL_HOST || '';
const PEDIDOS_EMAIL_USER = process.env.PEDIDOS_EMAIL_USER || '';
const PEDIDOS_EMAIL_PASS = process.env.PEDIDOS_EMAIL_PASS || '';
const PEDIDOS_EMAIL_PORT = parseInt(process.env.PEDIDOS_EMAIL_PORT) || 993;


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
  const documentosVistos = new Set(); // evita procesar el mismo documento dos veces (ej: si la API repite en paginación)
  let nextUrl = 'https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=' + fi + '&fecha_final=' + ff + '&page_size=100';
  let paginas = 0;
  let duplicadosOmitidos = 0;
  while (nextUrl && paginas < 200) {
    const resp = await fetch(nextUrl, { headers: { 'Authorization': API_KEY, 'Accept': 'application/json' } });
    if (!resp.ok) break;
    const data = await resp.json();
    const docs = (data.results || []).filter(d => {
      if (d.tipo_registro !== 'CLI') return false;  // solo clientes
      if (d.anulado) return false;                   // excluir anulados
      if (d.tipo_documento === 'NC') return false;   // excluir notas de crédito
      if (d.tipo_documento === 'COT') return false;  // excluir cotizaciones (no son ventas reales)
      if (d.tipo_documento === 'PRO') return false;  // excluir proformas
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
      const vendId = doc.vendedor?.id || doc.vendedor_identificacion || 'sin_vendedor';
      const vendNom = doc.vendedor?.razon_social || ('Vendedor ' + (doc.vendedor_identificacion || 'Sin Asignar'));
      const cliId = doc.cliente && doc.cliente.id ? doc.cliente.id : doc.persona_id;
      const cliNom = (doc.cliente && (doc.cliente.razon_social || doc.cliente.nombre_comercial)) || '—';
      const cliRuc = (doc.cliente && (doc.cliente.ruc || doc.cliente.cedula)) || '';
      // Buscar provincia con prioridad: override manual por RUC/Cédula (Excel) > catálogo
      // sincronizado de Contifico (por persona_id) > inferencia por dirección.
      const cliProv = resolverProvinciaCliente(cliRuc, cliId, doc.cliente?.direccion || '');
      const mes = parseInt((doc.fecha_emision || '').split('/')[1]) || 0;
      const totalDoc = parseFloat(doc.total || 0);
      const subDoc = parseFloat(doc.subtotal || doc.subtotal_12 || 0);
      if (!cliId || totalDoc === 0) return;
      if (!vendedores[vendId]) vendedores[vendId] = { nombre: vendNom, clientes: {} };
      vendedores[vendId].nombre = vendNom;
      const vObj = vendedores[vendId].clientes;
      if (!vObj[cliId]) vObj[cliId] = { id: cliId, nombre: cliNom, ruc: cliRuc, total: 0, subtotal: 0, num_compras: 0, provincia: cliProv, marcas: {}, marcasPorAnio: {}, marcasPorMes: {}, productos: {}, frecuencia: {} };
      const cli = vObj[cliId];
      cli.nombre = cliNom; cli.ruc = cliRuc;
      // La provincia se recalcula siempre con el valor más reciente de cliProv, en vez de
      // quedarse fija con el primer valor calculado. Esto es necesario porque el override
      // manual de provincias (Excel subido por Fernando) puede actualizarse en cualquier
      // momento, y de lo contrario un cliente que ya tenía un valor (aunque fuera "Sin
      // provincia" o uno inferido incorrectamente) nunca reflejaría la corrección.
      if(cliProv) cli.provincia = cliProv;
      cli.total += totalDoc; cli.subtotal += subDoc; cli.num_compras++;
      const anioDoc = parseInt((doc.fecha_emision || '').split('/')[2]) || new Date().getFullYear();
      const freqKey = `${anioDoc}-${String(mes).padStart(2,'0')}`;
      if (!cli.frecuencia[freqKey]) cli.frecuencia[freqKey] = { anio: anioDoc, mes, total: 0, subtotal: 0, compras: 0 };
      cli.frecuencia[freqKey].total += totalDoc; cli.frecuencia[freqKey].subtotal += subDoc; cli.frecuencia[freqKey].compras++;
      // Desglose exacto por día — para el gráfico "ventas del mes por día" (instantáneo, sin pegarle a Contifico en vivo)
      const diaDoc = parseInt((doc.fecha_emision || '').split('/')[0]) || 0;
      if (diaDoc) {
        const diaKey = `${anioDoc}-${String(mes).padStart(2,'0')}-${String(diaDoc).padStart(2,'0')}`;
        if (!cli.frecuenciaPorDia) cli.frecuenciaPorDia = {};
        if (!cli.frecuenciaPorDia[diaKey]) cli.frecuenciaPorDia[diaKey] = { anio: anioDoc, mes, dia: diaDoc, total: 0, subtotal: 0, compras: 0 };
        cli.frecuenciaPorDia[diaKey].total += totalDoc;
        cli.frecuenciaPorDia[diaKey].subtotal += subDoc;
        cli.frecuenciaPorDia[diaKey].compras++;
      }
      (doc.detalles || []).forEach(det => {
        const prodId = det.producto_id || '';
        const cantidad = parseFloat(det.cantidad || 0);
        const base = parseFloat(det.base_gravable || det.base_cero || 0);
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
    resultado[vend.nombre] = Object.values(vend.clientes).map(cli => ({
      id: cli.id, nombre: cli.nombre, ruc: cli.ruc,
      total: Math.round(cli.total * 100) / 100,
      subtotal: Math.round(cli.subtotal * 100) / 100,
      num_compras: cli.num_compras, provincia: cli.provincia,
      marcas: Object.entries(cli.marcas).map(([m,t]) => ({ marca: m, total: Math.round(t*100)/100 })).sort((a,b) => b.total-a.total),
      marcas_anio: Object.values(cli.marcasPorAnio||{}).map(x => ({ anio: x.anio, marca: x.marca, total: Math.round(x.total*100)/100 })),
      marcas_mes: Object.values(cli.marcasPorMes||{}).map(x => ({ anio: x.anio, mes: x.mes, marca: x.marca, total: Math.round(x.total*100)/100 })),
      productos: Object.values(cli.productos).map(p => ({ id: p.id, nombre: p.nombre, codigo: p.codigo, marca: p.marca, cantidad: Math.round(p.cantidad), total: Math.round(p.total*100)/100 })).sort((a,b) => b.cantidad-a.cantidad),
      productos_mes: Object.values(cli.productosPorMes||{}).map(x => ({ anio: x.anio, mes: x.mes, id: x.id, nombre: x.nombre, marca: x.marca, cantidad: Math.round(x.cantidad*100)/100, total: Math.round(x.total*100)/100 })),
      frecuencia: Object.values(cli.frecuencia).map(f => ({ anio: f.anio, mes: f.mes, total: Math.round(f.total*100)/100, subtotal: Math.round(f.subtotal*100)/100, compras: f.compras })).sort((a,b) => a.anio!==b.anio ? a.anio-b.anio : a.mes-b.mes),
      frecuencia_dia: Object.values(cli.frecuenciaPorDia||{}).map(f => ({ anio: f.anio, mes: f.mes, dia: f.dia, total: Math.round(f.total*100)/100, subtotal: Math.round(f.subtotal*100)/100, compras: f.compras }))
    })).sort((a,b) => b.total-a.total);
  });
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

// Rotación mensual: promedio de unidades vendidas en los 6 meses anteriores a fechaCorte
// (sin incluir el mes de corte, que normalmente está incompleto), usando productos_mes
// ya calculado en DATA_CACHE. Devuelve { [producto_id]: rotacionPromedioMensual }
function calcularRotacionMensual(fechaCorte) {
  const [anioCorte, mesCorte] = fechaCorte.split('-').map(Number); // YYYY-MM-DD
  // Construir lista de los 6 (anio,mes) anteriores al mes de corte (sin incluir el de corte)
  const mesesAtras = [];
  let a = anioCorte, m = mesCorte;
  for (let i = 0; i < 6; i++) {
    m -= 1;
    if (m === 0) { m = 12; a -= 1; }
    mesesAtras.push({ anio: a, mes: m });
  }
  const acumulado = {}; // producto_id -> total unidades en esos 6 meses
  Object.values(DATA_CACHE||{}).forEach(clientes => {
    (clientes||[]).forEach(cli => {
      (cli.productos_mes||[]).forEach(pm => {
        const matchMes = mesesAtras.some(x => x.anio===pm.anio && x.mes===pm.mes);
        if (!matchMes) return;
        const key = pm.id || pm.nombre;
        acumulado[key] = (acumulado[key]||0) + (pm.cantidad||0);
      });
    });
  });
  const rotacion = {};
  Object.entries(acumulado).forEach(([id, total]) => { rotacion[id] = total / 6; });
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
function construirInventarioPorMarca(marcaFiltro) {
  if (!INVENTARIO_CACHE) return { fecha_corte: null, productos: [] };
  const rotacion = calcularRotacionMensual(INVENTARIO_CACHE.fecha_corte);
  const productosDelCatalogo = Object.entries(catalogoProductos)
    .filter(([id, info]) => (info.marca||'').toUpperCase() === marcaFiltro)
    // Excluir "PROMOS": son combos armados a partir de otros productos, no tienen
    // stock propio ni rotación real — no aplica pedir reabastecimiento de ellos.
    .filter(([id, info]) => !(info.nombre||'').trim().toUpperCase().startsWith('PROMO'))
    // Excluir "Línea completa/Línea Facial/...": son agrupadores de catálogo, no
    // productos físicos con stock propio en bodega.
    .filter(([id, info]) => !(info.nombre||'').trim().toUpperCase().startsWith('LÍNEA') && !(info.nombre||'').trim().toUpperCase().startsWith('LINEA'));

  const lista = productosDelCatalogo.map(([id, info]) => {
    const inv = INVENTARIO_CACHE.productos[id];
    const stock = inv ? inv.cantidad : 0;
    const rotacionMensual = rotacion[id] || 0;
    const cobertura = rotacionMensual > 0 ? stock / rotacionMensual : (stock > 0 ? 99 : 0);
    // Cobertura 12 meses: cuántas unidades faltan (o sobran, si es negativo) para
    // tener cubierto todo el año a partir de la rotación actual.
    const necesidad12Meses = (rotacionMensual * 12) - stock;
    return {
      id,
      sku: info.codigo || (inv ? inv.sku : ''),
      nombre: info.nombre,
      marca: marcaFiltro,
      stock: Math.round(stock*100)/100,
      rotacion_mensual: Math.round(rotacionMensual*100)/100,
      cobertura_meses: Math.round(cobertura*10)/10,
      necesidad_12_meses: Math.round(necesidad12Meses),
      semaforo: calcularSemaforo(marcaFiltro, cobertura)
    };
  }).sort((a,b) => a.cobertura_meses - b.cobertura_meses);

  return { fecha_corte: INVENTARIO_CACHE.fecha_corte, productos: lista };
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
    const clientes = todos.filter(d => d.tipo_registro === 'CLI' && !d.anulado && d.tipo_documento !== 'NC' && d.tipo_documento !== 'COT' && d.tipo_documento !== 'PRO');
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
          `INSERT INTO facturas_detalle(documento_id, fecha, documento, cliente_nombre, vendedor_nombre, subtotal, total)
           VALUES($1,$2,$3,$4,$5,$6,$7)
           ON CONFLICT (documento_id, fecha) DO UPDATE SET
             documento=$3, cliente_nombre=$4, vendedor_nombre=$5, subtotal=$6, total=$7, actualizado_at=NOW()`,
          [
            String(d.id || d.documento),
            fechaParaSQL(fecha), // fecha en formato YYYY-MM-DD
            d.documento || '',
            d.cliente_nombre || '—',
            vendNom,
            parseFloat(d.subtotal || (d.total/1.15) || 0),
            parseFloat(d.total || 0)
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
            await pool.query(
              `INSERT INTO pedidos_web(numero_pedido, fecha, cliente_nombre, cedula_ruc, telefono, subtotal, total, productos, email_uid, html_crudo)
               VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
               ON CONFLICT (numero_pedido) DO UPDATE SET
                 fecha=$2, cliente_nombre=$3, cedula_ruc=$4, telefono=$5, subtotal=$6, total=$7, productos=$8, html_crudo=$10`,
              [
                pedido.numeroPedido,
                pedido.fecha,
                pedido.cliente || '—',
                pedido.cedulaRuc || null,
                pedido.telefono || null,
                pedido.subtotal || 0,
                pedido.total || 0,
                JSON.stringify(pedido.productos || []),
                String(seq),
                html
              ]
            );
            procesados++;
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
    return { ok: true, procesados, errores };
  } catch (e) {
    console.error('Error conectando a pedidos@cosetika.com:', e.message);
    try { if (client) await client.logout(); } catch(e2){}
    return { ok: false, error: e.message };
  }
}

sincronizarHoy().catch(e => console.error('Error sync inicial:', e.message));
setInterval(() => sincronizarHoy().catch(e => console.error('Error sync:', e.message)), 60 * 60 * 1000);

// Sync de pedidos web: revisa la casilla pedidos@cosetika.com cada 10 minutos
setTimeout(() => sincronizarPedidosWeb().catch(e => console.error('Error sync pedidos inicial:', e.message)), 15000);
setInterval(() => sincronizarPedidosWeb().catch(e => console.error('Error sync pedidos:', e.message)), 10 * 60 * 1000);

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

let regenerandoEnProceso = false;

async function fusionarMesActualEnCache() {
  if (!DATA_CACHE || Object.keys(DATA_CACHE).length === 0) return; // esperar a que haya data base cargada
  if (regenerandoEnProceso) { console.log('Fusión incremental omitida: regeneración manual en curso'); return; }
  regenerandoEnProceso = true;
  try {
    const hoy = nowEC();
    const anioActual = hoy.getFullYear();
    const mesActual = hoy.getMonth() + 1;
    const desde = fmtDateEC(new Date(anioActual, hoy.getMonth(), 1));
    const hasta = fmtDateEC(hoy);
    const dataMes = await generarDataJson(desde, hasta); // solo el mes en curso, rápido

    // Paso 1: quitar de DATA_CACHE cualquier dato del mes/año actual (será reemplazado limpio)
    Object.keys(DATA_CACHE).forEach(vendNom => {
      DATA_CACHE[vendNom].forEach(cli => {
        const freqMesViejo = (cli.frecuencia||[]).find(f=>f.anio===anioActual&&f.mes===mesActual);
        if (freqMesViejo) {
          cli.total = Math.round((cli.total - freqMesViejo.total)*100)/100;
          cli.subtotal = Math.round((cli.subtotal - freqMesViejo.subtotal)*100)/100;
          cli.num_compras = Math.max(0, (cli.num_compras||0) - freqMesViejo.compras);
        }
        cli.frecuencia = (cli.frecuencia||[]).filter(f => !(f.anio===anioActual && f.mes===mesActual));
        // frecuencia_dia: igual que frecuencia, se quita la porción del mes actual (será reemplazada limpia en el Paso 2)
        cli.frecuencia_dia = (cli.frecuencia_dia||[]).filter(f => !(f.anio===anioActual && f.mes===mesActual));
        // Restar del año actual lo que correspondía al mes actual (para no perder otros meses del mismo año)
        const marcasMesViejo = (cli.marcas_mes||[]).filter(x=>x.anio===anioActual&&x.mes===mesActual);
        cli.marcas_anio = (cli.marcas_anio||[]).map(ma=>{
          if(ma.anio!==anioActual) return ma;
          const aRestar = marcasMesViejo.find(m=>m.marca===ma.marca);
          return aRestar ? {...ma, total: Math.round((ma.total-aRestar.total)*100)/100} : ma;
        }).filter(ma=>ma.total>0 || ma.anio!==anioActual);
        cli.marcas_mes = (cli.marcas_mes||[]).filter(x => !(x.anio===anioActual && x.mes===mesActual));
        // productos_mes: igual que marcas_mes, se quita la porción del mes actual (será reemplazada limpia en el Paso 2)
        cli.productos_mes = (cli.productos_mes||[]).filter(x => !(x.anio===anioActual && x.mes===mesActual));
      });
    });

    // Paso 2: insertar los datos frescos del mes en curso
    Object.entries(dataMes).forEach(([vendNom, clientesMes]) => {
      if (!DATA_CACHE[vendNom]) DATA_CACHE[vendNom] = [];
      const porId = {}; DATA_CACHE[vendNom].forEach(c => { porId[c.id] = c; });
      const nuevos = [];
      clientesMes.forEach(cliMes => {
        let cli = porId[cliMes.id];
        if (!cli) { nuevos.push(cliMes); return; }
        cli.total = Math.round((cli.total + cliMes.total) * 100) / 100;
        cli.subtotal = Math.round((cli.subtotal + cliMes.subtotal) * 100) / 100;
        cli.num_compras = (cli.num_compras||0) + cliMes.num_compras;
        // Igual que en la fusión anual: copiar la provincia recién calculada (que ya
        // respeta el override más reciente) de vuelta al cliente existente en DATA_CACHE.
        if(cliMes.provincia) cli.provincia = cliMes.provincia;
        cli.frecuencia = (cli.frecuencia||[]).concat(cliMes.frecuencia);
        cli.frecuencia_dia = (cli.frecuencia_dia||[]).concat(cliMes.frecuencia_dia||[]);
        cli.marcas_anio = consolidarMarcasAnio((cli.marcas_anio||[]).concat(cliMes.marcas_anio));
        cli.marcas_mes = consolidarMarcasMes((cli.marcas_mes||[]).concat(cliMes.marcas_mes));
        cli.productos_mes = consolidarProductosMes((cli.productos_mes||[]).concat(cliMes.productos_mes||[]));
        // cli.marcas: reconstruir COMPLETO desde marcas_anio (que ya está bien mantenido,
        // con resta/suma correcta del mes actual arriba) — NUNCA acumular sobre el cli.marcas
        // anterior, porque ese campo no tenía el mismo tratamiento y se duplicaba cada 15 min.
        cli.marcas = consolidarMarcasAnio((cli.marcas_anio||[]).map(ma=>({anio:0,marca:ma.marca,total:ma.total})))
          .map(m=>({marca:m.marca,total:m.total})).sort((a,b)=>b.total-a.total);
        // Productos: NO se pueden sumar incrementalmente como antes (eso causaba doble conteo
        // cada 15 min, acumulando el mismo mes sobre sí mismo). En vez de eso, se reconstruye
        // cli.productos = productos_historico (todo excepto el mes en curso) + productos del
        // mes actual recién calculado desde cero. productos_historico se actualiza solo cuando
        // cambia el mes (ver más abajo), nunca durante el mes en curso.
        if (!cli.productos_historico_anio || !cli.productos_historico_mes ||
            cli.productos_historico_anio !== anioActual || cli.productos_historico_mes !== mesActual) {
          // Cambió el mes (o es la primera fusión tras un deploy/regeneración):
          // todo lo que había en cli.productos hasta ahora pasa a ser histórico.
          cli.productos_historico = (cli.productos||[]).map(p=>({...p}));
          cli.productos_historico_anio = anioActual;
          cli.productos_historico_mes = mesActual;
        }
        const prodMap = {}; (cli.productos_historico||[]).forEach(p=>{ prodMap[p.id||p.nombre] = {...p}; });
        (cliMes.productos||[]).forEach(p=>{
          const k = p.id||p.nombre;
          if(prodMap[k]){ prodMap[k] = {...prodMap[k], cantidad: prodMap[k].cantidad + p.cantidad, total: Math.round((prodMap[k].total+p.total)*100)/100}; }
          else prodMap[k] = {...p};
        });
        cli.productos = Object.values(prodMap).sort((a,b)=>b.cantidad-a.cantidad);
      });
      DATA_CACHE[vendNom] = Object.values(porId).concat(nuevos);
    });

    // Reordenar por total descendente
    Object.keys(DATA_CACHE).forEach(v => DATA_CACHE[v].sort((a,b)=>b.total-a.total));
    // Persistir en PostgreSQL para que sobreviva deploys/reinicios
    await guardarDataEnDB(DATA_CACHE);

    console.log(`✓ Fusión incremental del mes en curso completada (${desde} - ${hasta})`);
  } catch(e) {
    console.error('Error fusionando mes actual:', e.message);
  }
  regenerandoEnProceso = false;
}
setInterval(() => fusionarMesActualEnCache().catch(e => console.error(e)), 15 * 60 * 1000);

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
      cli.frecuencia = (cli.frecuencia||[]).concat(cliAnio.frecuencia);
      cli.frecuencia_dia = (cli.frecuencia_dia||[]).concat(cliAnio.frecuencia_dia||[]);
      cli.marcas_anio = (cli.marcas_anio||[]).concat(cliAnio.marcas_anio);
      cli.marcas_mes = (cli.marcas_mes||[]).concat(cliAnio.marcas_mes);
      cli.productos_mes = (cli.productos_mes||[]).concat(cliAnio.productos_mes||[]);
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
const TABLAS_BACKUP = ['inventario_data', 'usuarios', 'visitas', 'planificacion', 'asesor_zonas', 'asesor_provincias', 'ventas_data'];

async function generarBackupCompleto() {
  const backup = { generado_en: new Date().toISOString(), tablas: {} };
  for (const tabla of TABLAS_BACKUP) {
    try {
      const r = await pool.query(`SELECT * FROM ${tabla}`);
      backup.tablas[tabla] = r.rows;
    } catch(e) {
      backup.tablas[tabla] = { error: e.message };
    }
  }
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

// ─── REGENERACIÓN AUTOMÁTICA DIARIA (madrugada, hora Ecuador) ───────────────
// Trae desde el 1 de enero del año EN CURSO hasta hoy (no años anteriores: esos ya están
// cerrados contablemente y no cambian, así que regenerarlos cada noche sería trabajo
// desperdiciado — solo agrega tiempo de ejecución y carga sobre Contifico sin beneficio).
// Si en algún momento se necesita corregir datos de años anteriores, usar el botón manual
// de Configuración indicando el rango de fechas que corresponda.
async function regenerarDataAutomatico() {
  if (regenerandoEnProceso) { console.log('Regeneración automática diaria omitida: otra regeneración en curso'); return; }
  regenerandoEnProceso = true;
  try {
    const hoy = nowEC();
    const anioActual = hoy.getFullYear();
    const fi = `01/01/${anioActual}`;
    const ff = fmtDateEC(hoy);
    console.log(`⏰ Regeneración automática diaria: ${fi} al ${ff}`);
    const dataAnio = await generarDataJson(fi, ff);
    if (!DATA_CACHE || Object.keys(DATA_CACHE).length === 0) {
      // No hay caché previo (primer arranque): usar el resultado tal cual, sin fusión
      await guardarDataEnDB(dataAnio);
    } else {
      await fusionarAnioActualEnCache(anioActual, dataAnio);
      await guardarDataEnDB(DATA_CACHE);
    }
    try { fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(DATA_CACHE, null, 2)); } catch(e) {}
    console.log('✓ Regeneración automática completada (año ' + anioActual + ' refrescado, años anteriores intactos)');
  } catch(e) { console.error('Error en regeneración automática:', e.message); }
  regenerandoEnProceso = false;
}
// Programar para correr a las 2:00 AM hora Ecuador (UTC-5) cada día
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

  // ENVÍOS SERVIENTREGA — manifiestos diarios de guías
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
      const dias = parseInt(urlObj.searchParams.get('dias')) || null;
      let r;
      if (fecha) {
        r = await pool.query('SELECT * FROM pedidos_web WHERE fecha=$1 ORDER BY id DESC', [fecha]);
      } else if (dias) {
        r = await pool.query(`SELECT * FROM pedidos_web WHERE fecha >= (CURRENT_DATE - $1::int) ORDER BY fecha DESC, id DESC`, [dias]);
      } else {
        r = await pool.query('SELECT * FROM pedidos_web ORDER BY fecha DESC, id DESC LIMIT 200');
      }
      const pedidos = r.rows.map(row => ({
        ...row,
        productos: (() => { try { return JSON.parse(row.productos || '[]'); } catch(e){ return []; } })()
      }));
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
  // POST /api/pedidos-web/marcar-facturado  body: { numeroPedido, documentoFactura }
  // Permite marcar manualmente un pedido como facturado (por si el cruce automático no
  // lo detecta, ej. el nombre en Contifico es muy distinto al de la web)
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
      const clientes = todos.filter(doc => doc.tipo_registro === 'CLI' && !doc.anulado && doc.tipo_documento !== 'NC' && doc.tipo_documento !== 'COT' && doc.tipo_documento !== 'PRO');
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
        const clientes = todos.filter(doc => doc.tipo_registro === 'CLI' && !doc.anulado && doc.tipo_documento !== 'NC' && doc.tipo_documento !== 'COT' && doc.tipo_documento !== 'PRO');

        for (const doc of clientes) {
          const cliNom = doc.cliente?.razon_social || doc.cliente?.nombre_comercial || doc.persona_id || '—';
          const vendNom = doc.vendedor?.razon_social || doc.vendedor?.nombre || 'Sin asignar';
          await pool.query(
            `INSERT INTO facturas_detalle(documento_id, fecha, documento, cliente_nombre, vendedor_nombre, subtotal, total)
             VALUES($1,$2,$3,$4,$5,$6,$7)
             ON CONFLICT (documento_id, fecha) DO UPDATE SET
               documento=$3, cliente_nombre=$4, vendedor_nombre=$5, subtotal=$6, total=$7, actualizado_at=NOW()`,
            [
              String(doc.id || doc.documento),
              fechaSQL,
              doc.documento || '',
              cliNom,
              vendNom,
              parseFloat(doc.subtotal || (doc.total/1.15) || 0),
              parseFloat(doc.total || 0)
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
      Object.values(DATA_CACHE||{}).forEach(clientes => {
        (clientes||[]).forEach(cli => {
          (cli.frecuencia_dia||[]).forEach(f => {
            if (f.anio !== anio || f.mes !== mes) return;
            if (!porDia[f.dia]) porDia[f.dia] = { total: 0, subtotal: 0 };
            porDia[f.dia].total += f.total;
            porDia[f.dia].subtotal += f.subtotal;
          });
        });
      });
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
    const desdeParam = urlObj.searchParams.get('desde');
    const anioActual = nowEC().getFullYear();
    const fi = desdeParam || fmtDateEC(new Date(anioActual,0,1));
    const ff = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
    // El rango solicitado empieza en el año en curso (o después) → fusión segura, no toca
    // años anteriores. Si el rango pedido incluye años anteriores (ej. desde 2025), se
    // interpreta como intención deliberada de corregir histórico y se reemplaza todo el rango.
    const anioInicioSolicitado = parseInt(fi.split('/')[2]) || anioActual;
    const usarSoloAnioActual = anioInicioSolicitado >= anioActual;
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
    res.writeHead(200, {'Content-Type':'application/json'});
    res.end(JSON.stringify(DATA_CACHE || {}));
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
      while (nextUrl && paginas < 200) {
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
        if (d.tipo_documento === 'NC') return false;
        if (d.tipo_documento === 'COT') return false;
        if (d.tipo_documento === 'PRO') return false;
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
      while (nextUrl && paginas < 200) {
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
        if (d.tipo_documento === 'NC') return false;
        if (d.tipo_documento === 'COT') return false;
        if (d.tipo_documento === 'PRO') return false;
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

  // DIAGNÓSTICO: ver el RUC/Cédula real que trae Contifico para un cliente por nombre
  if (urlPath === '/api/provincias/diagnostico-cliente' && req.method === 'GET') {
    try {
      const nombreBuscado = (urlObj.searchParams.get('nombre') || '').toUpperCase().trim();
      const desde = urlObj.searchParams.get('desde') || '01/01/2025';
      const hasta = urlObj.searchParams.get('hasta') || fmtDateEC(nowEC());
      let todos = [];
      let nextUrl = `https://api.contifico.com/sistema/api/v2/documento/?fecha_inicial=${desde}&fecha_final=${hasta}&page_size=100`;
      let paginas = 0;
      while (nextUrl && paginas < 200) {
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
      const resultado = construirInventarioPorMarca(marca);
      res.writeHead(200, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:true, marca, ...resultado }));
    } catch(e) {
      res.writeHead(500, {'Content-Type':'application/json'});
      res.end(JSON.stringify({ ok:false, error: e.message }));
    }
    return;
  }

  // DESCARGAR BACKUP DIRECTO (sin correo, ya que Railway bloquea SMTP en este plan)
  if (urlPath === '/api/backup/descargar' && req.method === 'GET') {
    try {
      const backup = await generarBackupCompleto();
      const json = JSON.stringify(backup, null, 2);
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
    res.writeHead(200,{'Content-Type':MIME[path.extname(filePath)]||'text/plain'});
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200,{'Content-Type':'text/html'});
    fs.createReadStream(path.join(__dirname,'index.html')).pipe(res);
  }
});

server.listen(PORT,'0.0.0.0',()=>console.log(`Cosétika Dashboard running on port ${PORT}`));
