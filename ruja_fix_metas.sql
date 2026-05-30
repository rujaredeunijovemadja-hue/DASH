-- ════════════════════════════════════════════════════════════════
-- RUJA — CORREÇÃO EMERGENCIAL DAS METAS NO SUPABASE
-- Execute no SQL Editor do Supabase (dashboard.supabase.com)
-- ════════════════════════════════════════════════════════════════

-- 1. Ver o valor atual (diagnóstico)
SELECT chave, valor_json, pg_typeof(valor_json)
FROM ruja_configuracoes
WHERE chave = 'metas';

-- 2. Corrigir — sobrescrever com valores válidos
--    Ajuste os números conforme a meta real da RUJA
INSERT INTO ruja_configuracoes (chave, valor_json, atualizado_em)
VALUES (
  'metas',
  '{"ativosDepto": 20, "batizadosDepto": 10}'::jsonb,
  NOW()
)
ON CONFLICT (chave) DO UPDATE
  SET valor_json    = '{"ativosDepto": 20, "batizadosDepto": 10}'::jsonb,
      atualizado_em = NOW();

-- 3. Confirmar resultado
SELECT chave, valor_json
FROM ruja_configuracoes
WHERE chave IN ('metas', 'regras', 'lider_supremo');
