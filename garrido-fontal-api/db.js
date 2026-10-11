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
  `);
  console.log('Organizador: base de datos lista');
}

module.exports = { pool, q, migrate };
