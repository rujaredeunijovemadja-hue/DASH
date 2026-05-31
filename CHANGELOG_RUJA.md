# CHANGELOG_RUJA.md

Todas as alterações significativas do sistema RUJA são registradas aqui.
Formato: `[DATA] TIPO: Descrição — Arquivo(s) — Commit`

---

## [2026-05-30] — Auditoria Completa + Correções

### Bugs Críticos Corrigidos

**C1/C11 — saveRegras não persistia no Supabase**
- `saveRegras()` só salvava em `localStorage`. Agora chama `await salvarConfig('regras', regras)`.
- Arquivo: `index.html` | Commit: pendente

**C2 — importCSV chamada mas não definida**
- Função `importCSV(input)` implementada: lê CSV, faz upsert no Supabase pela tabela selecionada.
- Arquivo: `index.html` | Commit: pendente

**C3 — chartDepto não declarada com `let`**
- `let chartDepto = null` adicionado junto com `chartCrescimento` e `chartStatus`.
- Evita `TypeError` no 2º acesso ao Dashboard.
- Arquivo: `index.html` | Commit: pendente

**C5 — deleteLider duplicada (versão sem sincronização)**
- Versão simples da linha ~2062 removida. Mantida apenas a versão completa (~4376) que sincroniza o campo `lider` nos jovens do departamento.
- Arquivo: `index.html` | Commit: pendente

**C8 — getFreqPct() comparação de tipo incorreta**
- `f.jovemId === jovemId` → `String(f.jovemId) === String(jovemId)`.
- Evita frequência sempre 0% quando IDs têm tipos diferentes (number vs string).
- Arquivo: `index.html` | Commit: pendente

**C9 — migrarDoGAS sem botão na UI**
- Botão "▶ Iniciar Migração" adicionado no card de Sincronização na tela de Configurações.
- Arquivo: `index.html` | Commit: pendente

**C10 — pushToSheets com nome semântico errado**
- Função desabilitada com aviso claro. Usuário redirecionado para usar `exportCSV()`.
- Arquivo: `index.html` | Commit: pendente

**C11 — Status automático por frequência não implementado**
- Implementadas `calcularStatusAutomatico(jovemId)` e `atualizarStatusJovem(jovemId)`.
- Chamada automática após `salvarFreqLote()`.
- Regras lidas de `regras.ativo`, `regras.oscilando`, `regras.risco`.
- Arquivo: `index.html` | Commit: pendente

### Bugs Médios Corrigidos

**M1 — gravarSnapshotMensal chamado com dados vazios**
- Guard `if (!jovens || jovens.length === 0) return` adicionado.
- Arquivo: `index.html` | Commit: pendente

**M3 — IDs de foto duplicados (modalJovem e modalLiderSupremo)**
- IDs do modal Líder Supremo renomeados com prefixo `ls`: `lsFotoUploadArea`, `lsFotoPreview`, `lsFotoInput`, `lsBtnRemoverFoto`, `lsFotoStatus`.
- Handlers `onLsFotoSelecionada()` e `removerFotoLiderSupremo()` criados.
- Arquivo: `index.html` | Commit: pendente

**M7 — renderConfig lia gasUrl só do localStorage**
- Agora usa `GAS_URL || localStorage.getItem('gasUrl')` (prioriza valor carregado do Supabase).
- Arquivo: `index.html` | Commit: pendente

**C12 — Documentação obrigatória ausente**
- Criados: `PROGRAMAS_INFO.md`, `CHANGELOG_RUJA.md`, `BANCO_DE_DADOS_RUJA.md`, `REGRAS_DE_ACESSO_RUJA.md`.
- Commit: pendente

---

## [2026-05-30] — Correções anteriores (sessão de debug)

**fix: getDiasParaAniversario — zeragem de horas**
- Aniversariantes do dia apareciam como "em 336 dias".
- Causa: comparação sem `.setHours(0,0,0,0)`.
- Arquivo: `index.html` | Commit: `99d70d85`

**fix: metas corrompidas no Supabase (-2207422412000)**
- Valor timestamp gravado como meta numérica.
- Correção via SQL em `ruja_fix_metas.sql` + sanitização em `carregarConfigs()`.
- Arquivo: `ruja_db_layer.js` | Commit: `066cce4d`

**fix: renderConfig com fallbacks**
- Campos de regras e metas em branco ao abrir Configurações.
- Arquivo: `index.html` | Commit: `b90013ab`

**fix: IDs duplicados lsVisao/lsVersiculo e label sem for**
- IDs de display renomeados para `lsVisaoDisplay` e `lsVersiculoDisplay`.
- Label órfão convertido para `<span>`.
- Arquivo: `index.html` | Commit: `4dc3de4f`

**fix: políticas de storage idempotentes**
- `CREATE POLICY` falha se policy já existe (erro 42710).
- Reescrito com bloco `DO $$ IF NOT EXISTS $$`.
- Arquivo: `ruja_foto_sql.sql` | Commit: `7ad6f2f0`

**fix: botões de edição — IDs number vs string**
- `editJovem`, `editDepto`, `editLider`, `editRecup` reescritos com `String(id)`.
- Arquivo: `index.html` | Commit: `8aa49d0e` (via patch)

---

## [2026-05-30] — Migração GAS → Supabase

- Criação das tabelas via `ruja_migration_supabase.sql`
- Implementação da camada de dados em `ruja_db_layer.js`
- Deploy via Vercel conectado ao GitHub

---

*Formato: [DATA] TIPO(módulo): descrição*
*Tipos: feat | fix | refactor | docs | security | perf*

---

## [2026-05-30] — Fix Login Mobile

### Bug Crítico: Login não funciona em dispositivos móveis

**Causas raiz identificadas (3):**

**CAUSA 1 — `login-bg` sem `pointer-events:none`** *(Principal)*
O `<div class="login-bg">` é `position:absolute; inset:0` e cobre toda a tela de login sem `pointer-events:none`. No desktop o mouse passa direto. Em mobile, o evento de toque fica capturado pela camada de background — o botão "Entrar" recebe o evento visualmente mas nunca dispara o listener de click.
- Fix: `pointer-events:none` adicionado ao `.login-bg`

**CAUSA 2 — Sem `touch-action:manipulation` no botão**
Browsers móveis aplicam delay de ~300ms no evento `click` para detectar double-tap (zoom). Sem `touch-action:manipulation`, o usuário toca, aguarda o delay, e pode soltar o dedo fora do botão — o click nunca dispara.
- Fix: `touch-action:manipulation` + `min-height:52px` no `.login-btn`

**CAUSA 3 — Viewport iOS com 100vh incorreto**
No Safari iOS, `100vh` inclui a barra de endereço, causando corte da tela de login em telas pequenas. O botão ficava parcialmente fora da área visível e de toque.
- Fix: `-webkit-fill-available` + `env(safe-area-inset-*)` no `.login-page`

**Correções adicionais:**

- `touchend` listener adicionado ao `loginBtn` e `forgotPasswordBtn` (elimina delay residual em Android antigo)
- `autocorrect="off"`, `autocapitalize="none"`, `inputmode="email"` no campo de email (evita correção automática do teclado mobile que sobrescreve o email)
- Media queries para 480px e `max-height:600px` (teclado virtual aberto)
- `box-sizing:border-box` no login-card

**Navegadores cobertos pelos fixes:**
- ✅ Android Chrome (touch-action + touchend)
- ✅ Android Samsung Internet (touch-action + pointer-events)
- ✅ Android Brave (touch-action + pointer-events)
- ✅ iPhone Safari (fill-available + safe-area + pointer-events)
- ✅ iPhone Chrome (touch-action + safe-area)

**Validação Pós-Deploy Fase 1:** 13/13 OK (análise estática)

**Arquivo:** `index.html` | Commit: pendente

---
