-- ============================================================
-- CONFERÊNCIA REDE — Schema Supabase
-- Cole e execute no SQL Editor do seu projeto Supabase
-- ============================================================

-- 1. Importações (um registro por arquivo importado)
CREATE TABLE imports (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  filename          TEXT NOT NULL,
  imported_at       TIMESTAMPTZ DEFAULT NOW(),
  transaction_count INTEGER DEFAULT 0,
  gross_total       NUMERIC(12,2) DEFAULT 0,
  net_total         NUMERIC(12,2) DEFAULT 0,
  storage_path      TEXT
);

-- 2. Transações (todas acumuladas de todos os extratos)
CREATE TABLE transactions (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  import_id       UUID REFERENCES imports(id) ON DELETE CASCADE,
  date            DATE NOT NULL,
  description     TEXT,
  type            TEXT NOT NULL,
  gross_amount    NUMERIC(12,2) NOT NULL,
  net_amount      NUMERIC(12,2) NOT NULL,
  taxa_pct        NUMERIC(5,2),
  prazo_dias      INTEGER,
  settlement_date DATE,
  card_brand      TEXT,
  nsu             TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Conferência (o que efetivamente caiu na conta por dia)
CREATE TABLE reconciliations (
  settlement_date DATE PRIMARY KEY,
  actual_amount   NUMERIC(12,2),
  notes           TEXT,
  updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- Índices para performance
CREATE INDEX idx_tx_import_id       ON transactions(import_id);
CREATE INDEX idx_tx_settlement_date ON transactions(settlement_date);
CREATE INDEX idx_tx_date            ON transactions(date);
CREATE INDEX idx_tx_type            ON transactions(type);

-- ============================================================
-- STORAGE: crie o bucket manualmente no painel do Supabase
-- Storage → New bucket → nome: "extratos" → Private
-- ============================================================
