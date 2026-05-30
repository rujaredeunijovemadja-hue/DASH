# REGRAS_DE_ACESSO_RUJA.md

Políticas de acesso, autenticação e segurança do sistema RUJA.

---

## Autenticação

**Provider:** Supabase Auth (email + senha)
**Método:** `supabase.auth.signInWithPassword({ email, password })`
**Recuperação de senha:** `supabase.auth.resetPasswordForEmail(email)`
**Sessão:** mantida via `supabase.auth.getSession()` no carregamento
**Alterar senha:** `supabase.auth.updateUser({ password })`

### Fluxo de login
1. Usuário informa e-mail e senha
2. Supabase valida credenciais
3. Se válido → `mostrarApp(nome, funcao)` exibe o painel
4. Se inválido → mensagem de erro no campo `#loginError`
5. A sessão persiste via cookie/localStorage gerenciado pelo Supabase SDK

---

## Papéis (Roles)

Atualmente o sistema não implementa diferenciação de papéis. Todos os usuários autenticados têm acesso total (CRUD) a todos os dados.

**Papéis planejados (não implementados):**
| Papel | Acesso |
|-------|--------|
| Admin | Tudo — incluindo configurações, exclusão, migração |
| Líder | Leitura total, edição de jovens/frequências do seu departamento |
| Visualizador | Apenas leitura |

---

## Row Level Security (RLS)

**Status:** habilitado em todas as tabelas ✅

**Política atual (todas as tabelas):**
```sql
CREATE POLICY "auth_all" ON <tabela>
  FOR ALL TO authenticated
  USING (true)
  WITH CHECK (true);
```

**Interpretação:** qualquer usuário autenticado pode ler, inserir, atualizar e excluir qualquer registro de qualquer tabela. Não há isolamento por usuário ou departamento.

**Limitação conhecida:** adequado para grupo pequeno e confiável. Para múltiplas igrejas ou líderes com acesso restrito, implementar políticas por `auth.uid()`.

---

## Storage

**Bucket:** `ruja-jovens-fotos` (privado)
**Acesso:** apenas usuários autenticados

**Políticas:**
```sql
-- SELECT (download/view)
CREATE POLICY "ruja_foto_select" ON storage.objects
  FOR SELECT TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');

-- INSERT (upload)
CREATE POLICY "ruja_foto_insert" ON storage.objects
  FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'ruja-jovens-fotos');

-- UPDATE
CREATE POLICY "ruja_foto_update" ON storage.objects
  FOR UPDATE TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');

-- DELETE
CREATE POLICY "ruja_foto_delete" ON storage.objects
  FOR DELETE TO authenticated
  USING (bucket_id = 'ruja-jovens-fotos');
```

---

## Dados Sensíveis

| Dado | Armazenamento | Observação |
|------|---------------|-----------|
| Credenciais de acesso | Supabase Auth | Hashed — nunca acessível |
| Nome, contato, endereço dos jovens | `ruja_jovens` (Supabase) + `localStorage` (cache) | localStorage sem criptografia |
| Fotos | Supabase Storage (privado) | URL pública requer auth |
| Chave Supabase anon | `index.html` (hardcoded) | Publishable key — seguro expor no cliente |
| Service role key | **não exposta** ✅ | Nunca deve aparecer no código cliente |

---

## Regras de Negócio de Acesso

1. **Sem autenticação → nenhum dado visível** — o app exige login antes de exibir qualquer tela
2. **Login falhou → mensagem de erro** — sem detalhes técnicos expostos ao usuário
3. **Sessão expirada → tela de login** — gerenciado pelo SDK do Supabase
4. **Exclusão de jovem** → cascata para `ruja_frequencias` e `ruja_recuperacoes` (ON DELETE CASCADE no banco)
5. **Exclusão de líder** → sincroniza campo `lider` nos jovens do departamento antes de excluir

---

## Riscos e Recomendações

| Risco | Severidade | Recomendação |
|-------|-----------|--------------|
| RLS sem isolamento por papel | Médio | Implementar policies por `auth.uid()` se expandir para múltiplas igrejas |
| `localStorage` com dados pessoais | Baixo | Implementar limpeza ao fazer logout |
| Sem timeout de sessão explícito | Baixo | Configurar no dashboard Supabase: Auth → Settings → JWT expiry |
| Sem 2FA | Baixo | Habilitar no Supabase Auth se necessário |
| `ruja_audit_logs` não utilizada | Baixo | Implementar registro de alterações sensíveis |

---

*Última atualização: 2026-05-30*
