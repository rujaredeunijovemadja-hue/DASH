-- ═══════════════════════════════════════════════════════════════
-- RUJA — MIGRAÇÃO PARA SUPABASE
-- Execute este script no SQL Editor do Supabase
-- Execute bloco por bloco para evitar truncamento
-- ═══════════════════════════════════════════════════════════════

-- ─── BLOCO 1: Tabela de Departamentos ───
CREATE TABLE IF NOT EXISTS ruja_departamentos (
  id          TEXT PRIMARY KEY,
  nome        TEXT NOT NULL,
  icone       TEXT DEFAULT '🏛',
  lider       TEXT,
  capacidade  INTEGER DEFAULT 0,
  descricao   TEXT DEFAULT '',
  criado_em   TIMESTAMPTZ DEFAULT NOW(),
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_departamentos ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_departamentos FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 2: Tabela de Líderes ───
CREATE TABLE IF NOT EXISTS ruja_lideres (
  id          TEXT PRIMARY KEY,
  nome        TEXT NOT NULL,
  contato     TEXT DEFAULT '',
  departamento TEXT DEFAULT '',
  funcao      TEXT DEFAULT '',
  data_nasc   TEXT DEFAULT '',
  criado_em   TIMESTAMPTZ DEFAULT NOW(),
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_lideres ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_lideres FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 3: Tabela de Jovens ───
CREATE TABLE IF NOT EXISTS ruja_jovens (
  id            TEXT PRIMARY KEY,
  nome          TEXT NOT NULL,
  idade         INTEGER DEFAULT 0,
  contato       TEXT DEFAULT '',
  instagram     TEXT DEFAULT '',
  endereco      TEXT DEFAULT '',
  departamento  TEXT DEFAULT '',
  lider         TEXT DEFAULT '',
  status        TEXT DEFAULT 'Em Risco',
  entrada       TEXT DEFAULT '',
  batizado      TEXT DEFAULT 'nao',
  data_batismo  TEXT DEFAULT '',
  data_nasc     TEXT DEFAULT '',
  obs           TEXT DEFAULT '',
  criado_em     TIMESTAMPTZ DEFAULT NOW(),
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_jovens ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_jovens FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 4: Tabela de Frequências ───
CREATE TABLE IF NOT EXISTS ruja_frequencias (
  id        TEXT PRIMARY KEY,
  jovem_id  TEXT NOT NULL REFERENCES ruja_jovens(id) ON DELETE CASCADE,
  data      TEXT NOT NULL,
  evento    TEXT DEFAULT '',
  presenca  TEXT DEFAULT 'falta',
  obs       TEXT DEFAULT '',
  criado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_frequencias ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_frequencias FOR ALL TO authenticated USING (true) WITH CHECK (true);
CREATE INDEX IF NOT EXISTS idx_freq_jovem ON ruja_frequencias(jovem_id);
CREATE INDEX IF NOT EXISTS idx_freq_data  ON ruja_frequencias(data);

-- ─── BLOCO 5: Tabela de Recuperações ───
CREATE TABLE IF NOT EXISTS ruja_recuperacoes (
  id            TEXT PRIMARY KEY,
  jovem_id      TEXT NOT NULL REFERENCES ruja_jovens(id) ON DELETE CASCADE,
  data_inicio   TEXT DEFAULT '',
  lider_resp    TEXT DEFAULT '',
  motivo        TEXT DEFAULT '',
  status        TEXT DEFAULT 'ativo',
  obs           TEXT DEFAULT '',
  criado_em     TIMESTAMPTZ DEFAULT NOW(),
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_recuperacoes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_recuperacoes FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 6: Tabela de Histórico Mensal ───
CREATE TABLE IF NOT EXISTS ruja_historico_mensal (
  id              SERIAL PRIMARY KEY,
  mes             TEXT NOT NULL UNIQUE,
  ativos_depto    INTEGER DEFAULT 0,
  batizados_depto INTEGER DEFAULT 0,
  total           INTEGER DEFAULT 0,
  criado_em       TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_historico_mensal ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_historico_mensal FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 7: Tabela de Configurações Gerais (JSONB) ───
CREATE TABLE IF NOT EXISTS ruja_configuracoes (
  chave       TEXT PRIMARY KEY,
  valor_json  JSONB NOT NULL,
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_configuracoes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_configuracoes FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- Insere configs padrão (só se não existirem)
INSERT INTO ruja_configuracoes (chave, valor_json) VALUES
  ('regras',        '{"ativo":75,"oscilando":40,"risco":3}'),
  ('metas',         '{"ativosDepto":20,"batizadosDepto":10}'),
  ('lider_supremo', '{"nome":"","contato":"","instagram":"","foto":"","descricao":"","dataPosseLider":"","versiculoLider":"","visao":"","tempoNaRuja":""}'),
  ('database_mode', '"supabase"'),
  ('backup_mode',   '"google_sheets"')
ON CONFLICT (chave) DO NOTHING;

-- Atualiza config existente para registrar modo Supabase
INSERT INTO config (chave, valor) VALUES ('database_mode','supabase')
ON CONFLICT (chave) DO UPDATE SET valor = 'supabase';

INSERT INTO config (chave, valor) VALUES ('backup_mode','google_sheets')
ON CONFLICT (chave) DO UPDATE SET valor = 'google_sheets';

-- ─── BLOCO 8: Tabela de Logs de Migração ───
CREATE TABLE IF NOT EXISTS migration_logs (
  id                SERIAL PRIMARY KEY,
  tabela            TEXT NOT NULL,
  registros_migrados INTEGER DEFAULT 0,
  data_execucao     TIMESTAMPTZ DEFAULT NOW(),
  status            TEXT DEFAULT 'pendente',
  observacao        TEXT DEFAULT ''
);
ALTER TABLE migration_logs ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON migration_logs FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 9: Tabela de Audit Logs ───
CREATE TABLE IF NOT EXISTS ruja_audit_logs (
  id          SERIAL PRIMARY KEY,
  usuario_id  UUID REFERENCES auth.users(id),
  acao        TEXT NOT NULL,
  tabela      TEXT NOT NULL,
  registro_id TEXT,
  dados_antes JSONB,
  dados_depois JSONB,
  criado_em   TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE ruja_audit_logs ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all" ON ruja_audit_logs FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ─── BLOCO 10: Índices de performance ───
CREATE INDEX IF NOT EXISTS idx_jovens_status ON ruja_jovens(status);
CREATE INDEX IF NOT EXISTS idx_jovens_depto  ON ruja_jovens(departamento);
CREATE INDEX IF NOT EXISTS idx_rec_status    ON ruja_recuperacoes(status);
