-- ═══════════════════════════════════════════════════════════════
-- RUJA — FOTOS DE JOVENS
-- Execute BLOCO 1 e depois BLOCO 2 separadamente
-- ═══════════════════════════════════════════════════════════════

-- ─── BLOCO 1: Colunas na tabela de jovens ───
ALTER TABLE ruja_jovens ADD COLUMN IF NOT EXISTS foto_path TEXT DEFAULT '';
ALTER TABLE ruja_jovens ADD COLUMN IF NOT EXISTS foto_url  TEXT DEFAULT '';

-- ─── BLOCO 2: Bucket + políticas de acesso ───
-- Cria bucket privado
INSERT INTO storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
VALUES (
  'ruja-jovens-fotos',
  'ruja-jovens-fotos',
  false,
  2097152,
  ARRAY['image/jpeg','image/jpg','image/png','image/webp']
)
ON CONFLICT (id) DO NOTHING;

-- Políticas RLS do Storage
CREATE POLICY "ruja_foto_select" ON storage.objects
  FOR SELECT TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');

CREATE POLICY "ruja_foto_insert" ON storage.objects
  FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'ruja-jovens-fotos');

CREATE POLICY "ruja_foto_update" ON storage.objects
  FOR UPDATE TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');

CREATE POLICY "ruja_foto_delete" ON storage.objects
  FOR DELETE TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');
