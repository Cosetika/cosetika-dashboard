-- Tabla para el planificador de capacitaciones
-- Ejecutar una sola vez contra tu base de datos PostgreSQL existente

CREATE TABLE IF NOT EXISTS capacitaciones (
  id SERIAL PRIMARY KEY,
  fecha DATE NOT NULL,
  ciudad VARCHAR(120) NOT NULL DEFAULT '',
  tema VARCHAR(255) NOT NULL DEFAULT '',
  direccion VARCHAR(255) NOT NULL DEFAULT '',
  horario VARCHAR(120) NOT NULL DEFAULT '',
  valor NUMERIC(10,2) DEFAULT 0,
  created_at TIMESTAMP NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMP NOT NULL DEFAULT NOW()
);

-- Acelera las consultas por mes (WHERE EXTRACT(YEAR/MONTH FROM fecha) = ...)
CREATE INDEX IF NOT EXISTS idx_capacitaciones_fecha ON capacitaciones(fecha);
