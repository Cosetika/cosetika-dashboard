// capacitaciones-rutas.js
//
// CÓMO INTEGRAR (manual, sin tocar el resto de tu app):
//
// 1. Copia este archivo a la carpeta de tus rutas.
// 2. En tu server.js (o donde montas tus rutas), agrega:
//
//      const capacitacionesRutas = require('./capacitaciones-rutas');
//      app.use('/api/capacitaciones', capacitacionesRutas(pool));
//
//    donde "pool" es tu instancia existente de `pg.Pool` (la misma que usas
//    para votación u otras rutas).
//
// 3. Define la variable de entorno CAPACITACIONES_EDIT_CODE en Railway con el
//    código que Giovanna usará para entrar en modo edición. Si no la defines,
//    se usa 'GIOVANNA2026' por defecto (cámbialo antes de ir a producción).
//
// 4. Corre tablas_capacitaciones.sql una vez contra tu base de datos.

const express = require('express');

module.exports = function (pool) {
  const router = express.Router();
  const EDIT_CODE = process.env.CAPACITACIONES_EDIT_CODE || 'GIOVANNA2026';

  // Middleware simple: exige el código de edición para crear/editar/borrar.
  // Esto NO es un sistema de usuarios — es un candado compartido, como
  // conversamos. Si más adelante quieres login real por usuario, este
  // middleware es el lugar para reemplazarlo.
  function checkEditCode(req, res, next) {
    const code = req.headers['x-edit-code'] || req.body.codigo;
    if (code !== EDIT_CODE) {
      return res.status(403).json({ error: 'Código de edición inválido' });
    }
    next();
  }

  // GET /api/capacitaciones?year=2026&month=7
  // Devuelve todas las capacitaciones de ese mes, ordenadas por fecha.
  router.get('/', async (req, res) => {
    const { year, month } = req.query;
    if (!year || !month) {
      return res.status(400).json({ error: 'year y month son requeridos' });
    }
    try {
      const result = await pool.query(
        `SELECT * FROM capacitaciones
         WHERE EXTRACT(YEAR FROM fecha) = $1 AND EXTRACT(MONTH FROM fecha) = $2
         ORDER BY fecha ASC`,
        [year, month]
      );
      res.json(result.rows);
    } catch (err) {
      console.error('Error GET /api/capacitaciones:', err);
      res.status(500).json({ error: 'Error al obtener capacitaciones' });
    }
  });

  // POST /api/capacitaciones
  // Crea una fila nueva. Requiere código de edición.
  router.post('/', checkEditCode, async (req, res) => {
    const { fecha, ciudad, tema, direccion, horario, valor } = req.body;
    if (!fecha) {
      return res.status(400).json({ error: 'fecha es requerida' });
    }
    try {
      const result = await pool.query(
        `INSERT INTO capacitaciones (fecha, ciudad, tema, direccion, horario, valor)
         VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
        [fecha, ciudad || '', tema || '', direccion || '', horario || '', valor || 0]
      );
      res.status(201).json(result.rows[0]);
    } catch (err) {
      console.error('Error POST /api/capacitaciones:', err);
      res.status(500).json({ error: 'Error al crear capacitación' });
    }
  });

  // PUT /api/capacitaciones/:id
  // Actualiza uno o varios campos de una fila. Requiere código de edición.
  router.put('/:id', checkEditCode, async (req, res) => {
    const { id } = req.params;
    const { fecha, ciudad, tema, direccion, horario, valor } = req.body;
    try {
      const result = await pool.query(
        `UPDATE capacitaciones SET
           fecha = COALESCE($1, fecha),
           ciudad = COALESCE($2, ciudad),
           tema = COALESCE($3, tema),
           direccion = COALESCE($4, direccion),
           horario = COALESCE($5, horario),
           valor = COALESCE($6, valor),
           updated_at = NOW()
         WHERE id = $7 RETURNING *`,
        [fecha, ciudad, tema, direccion, horario, valor, id]
      );
      if (result.rows.length === 0) {
        return res.status(404).json({ error: 'Capacitación no encontrada' });
      }
      res.json(result.rows[0]);
    } catch (err) {
      console.error('Error PUT /api/capacitaciones/:id:', err);
      res.status(500).json({ error: 'Error al actualizar capacitación' });
    }
  });

  // DELETE /api/capacitaciones/:id
  // Elimina una fila. Requiere código de edición.
  router.delete('/:id', checkEditCode, async (req, res) => {
    const { id } = req.params;
    try {
      await pool.query('DELETE FROM capacitaciones WHERE id = $1', [id]);
      res.status(204).end();
    } catch (err) {
      console.error('Error DELETE /api/capacitaciones/:id:', err);
      res.status(500).json({ error: 'Error al eliminar capacitación' });
    }
  });

  return router;
};
