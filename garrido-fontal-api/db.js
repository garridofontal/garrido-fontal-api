// ── Base de datos del Organizador de Trabajos (Postgres en Railway) ──────────
const { Pool } = require('pg');

const pool = process.env.DATABASE_URL
  ? new Pool({ connectionString: process.env.DATABASE_URL, max: 5 })
  : null;

async function q(text, params) {
  if (!pool) throw new Error('DATABASE_URL non configurada');
  return pool.query(text, params);
}

async function migrate() {
  if (!pool) { console.warn('Organizador: sen DATABASE_URL, non se crea a base de datos'); return; }
  await q(`
    CREATE TABLE IF NOT EXISTS clientes (
      id          SERIAL PRIMARY KEY,
      nombre      TEXT NOT NULL,
      tipo        TEXT NOT NULL DEFAULT 'Particular',
      cif         TEXT NOT NULL DEFAULT '',
      direccion   TEXT NOT NULL DEFAULT '',
      cp          TEXT NOT NULL DEFAULT '',
      concello    TEXT NOT NULL DEFAULT '',
      telefono    TEXT NOT NULL DEFAULT '',
      email       TEXT NOT NULL DEFAULT '',
      notas       TEXT NOT NULL DEFAULT '',
      avisos      JSONB NOT NULL DEFAULT '[]',
      alias       JSONB NOT NULL DEFAULT '[]',
      origen      TEXT NOT NULL DEFAULT 'app',
      creado      TIMESTAMPTZ NOT NULL DEFAULT now(),
      actualizado TIMESTAMPTZ NOT NULL DEFAULT now(),
      UNIQUE (cif, nombre)
    );
    CREATE TABLE IF NOT EXISTS trabajos (
      id                  SERIAL PRIMARY KEY,
      codigo              TEXT UNIQUE NOT NULL,
      cliente_id          INTEGER NOT NULL REFERENCES clientes(id),
      tipo                TEXT NOT NULL DEFAULT 'Velux',
      descripcion         TEXT NOT NULL DEFAULT '',
      origen              TEXT NOT NULL DEFAULT '',
      direccion_obra      TEXT NOT NULL DEFAULT '',
      cp_obra             TEXT NOT NULL DEFAULT '',
      concello_obra       TEXT NOT NULL DEFAULT '',
      estado              SMALLINT NOT NULL DEFAULT 0,
      cancelado_en        TIMESTAMPTZ,
      presupuesto_enviado DATE,
      visita_en           TIMESTAMPTZ,
      montaje_en          TIMESTAMPTZ,
      notas               TEXT NOT NULL DEFAULT '',
      creado              TIMESTAMPTZ NOT NULL DEFAULT now(),
      actualizado         TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS trabajo_eventos (
      id         SERIAL PRIMARY KEY,
      trabajo_id INTEGER NOT NULL REFERENCES trabajos(id) ON DELETE CASCADE,
      tipo       TEXT NOT NULL,
      estado     SMALLINT,
      texto      TEXT NOT NULL DEFAULT '',
      fecha      TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS documentos (
      id          SERIAL PRIMARY KEY,
      tipo        TEXT NOT NULL,
      numero      TEXT,
      fecha       DATE,
      base        NUMERIC(12,2) NOT NULL DEFAULT 0,
      total       NUMERIC(12,2) NOT NULL DEFAULT 0,
      origen      TEXT NOT NULL DEFAULT 'app',
      cliente_id  INTEGER REFERENCES clientes(id),
      trabajo_id  INTEGER REFERENCES trabajos(id) ON DELETE SET NULL,
      cobrada     BOOLEAN NOT NULL DEFAULT false,
      cobrada_en  DATE,
      creado      TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS archivos (
      id         SERIAL PRIMARY KEY,
      trabajo_id INTEGER REFERENCES trabajos(id) ON DELETE CASCADE,
      cliente_id INTEGER REFERENCES clientes(id),
      tipo       TEXT NOT NULL DEFAULT 'foto',
      nombre     TEXT NOT NULL DEFAULT '',
      mime       TEXT NOT NULL DEFAULT 'application/octet-stream',
      tamano     INTEGER NOT NULL DEFAULT 0,
      datos      BYTEA NOT NULL,
      miniatura  BYTEA,
      creado     TIMESTAMPTZ NOT NULL DEFAULT now()
    );
    CREATE INDEX IF NOT EXISTS archivos_trabajo ON archivos (trabajo_id);
    ALTER TABLE documentos ADD COLUMN IF NOT EXISTS lineas JSONB;
    ALTER TABLE documentos ADD COLUMN IF NOT EXISTS iva_pct NUMERIC(5,2);
    ALTER TABLE documentos ADD COLUMN IF NOT EXISTS datos JSONB;
    ALTER TABLE documentos ADD COLUMN IF NOT EXISTS archivo_id INTEGER REFERENCES archivos(id) ON DELETE SET NULL;
    CREATE UNIQUE INDEX IF NOT EXISTS documentos_tipo_numero ON documentos (tipo, numero) WHERE numero IS NOT NULL;
    CREATE INDEX IF NOT EXISTS trabajos_estado ON trabajos (estado);
    CREATE INDEX IF NOT EXISTS eventos_trabajo ON trabajo_eventos (trabajo_id);
    CREATE INDEX IF NOT EXISTS documentos_cliente ON documentos (cliente_id);
    ALTER TABLE trabajos ADD COLUMN IF NOT EXISTS cerrado_en TIMESTAMPTZ;
    ALTER TABLE documentos ADD COLUMN IF NOT EXISTS pdf_id INTEGER REFERENCES archivos(id) ON DELETE SET NULL;
    CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT, fecha TIMESTAMPTZ NOT NULL DEFAULT now());
  `);
  await migrarFases();
  console.log('Organizador: base de datos lista');
}

// Fases nuevas: 0 Contacto · 1 Visita · 2 Presupuesto · 3 Contrato · 4 Instalación · 5 Cobro (+ «cobrado» = trabajo cerrado)
// Antes: 0 Contacto · 1 Pendiente de ver · 2 Pendiente de presupuesto · 3 Pendiente de hacer · 4 Pendiente de cobrar · 5 Terminado
// Tipos nuevos: Velux · Mantenimiento · Accesorios · Reparación. Se ejecuta una sola vez.
// Nombre nuevo que corresponde a cada número de fase antiguo (para reescribir el historial)
const NOMBRES_ANTIGUOS = ['Contacto', 'Visita', 'Presupuesto', 'Contrato', 'Cobro', 'Cobrado'];
async function migrarFases() {
  if ((await q(`SELECT 1 FROM meta WHERE clave='fases_v2'`)).rows.length) return;
  const cli = await pool.connect();
  try {
    await cli.query('BEGIN');
    await cli.query(`UPDATE trabajos SET cerrado_en = coalesce(actualizado, now()) WHERE estado = 5 AND cerrado_en IS NULL`);
    await cli.query(`UPDATE trabajo_eventos SET texto = 'Estado cambiado – ' || ($1::text[])[estado + 1] WHERE tipo = 'estado' AND estado BETWEEN 0 AND 5`, [NOMBRES_ANTIGUOS]);
    await cli.query(`UPDATE trabajos SET estado = 5 WHERE estado = 4`);
    await cli.query(`UPDATE trabajo_eventos SET estado = 5 WHERE tipo = 'estado' AND estado = 4`);
    await cli.query(`UPDATE trabajos SET tipo = CASE tipo WHEN 'Cubierta' THEN 'Mantenimiento' WHEN 'Gotera' THEN 'Reparación' ELSE 'Reparación' END
                     WHERE tipo NOT IN ('Velux', 'Mantenimiento', 'Accesorios', 'Reparación')`);
    await cli.query(`INSERT INTO meta (clave, valor) VALUES ('fases_v2', 'ok')`);
    await cli.query('COMMIT');
    console.log('Organizador: fases y tipos migrados');
  } catch (e) { await cli.query('ROLLBACK'); throw e; } finally { cli.release(); }
}

module.exports = { pool, q, migrate };
