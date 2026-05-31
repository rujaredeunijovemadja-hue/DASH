# PROGRAMAS_INFO.md — RUJA Sistema de Gestão

> **REGRA:** Este arquivo deve ser atualizado a cada alteração de módulo, dependência ou risco.
> Nunca remover, ignorar ou criar documentação paralela em substituição.

---

## Visão Geral

**Sistema:** RUJA — Painel de Gestão da Rede UniJovem ADJA
**Deploy:** Vercel (via GitHub `rujaredeunijovemadja-hue/DASH`)
**Banco de dados:** Supabase (PostgreSQL + Auth + Storage)
**Arquivo principal:** `index.html` (Single Page Application, ~5200 linhas)

---

## Arquitetura

```
Browser
  └── index.html (SPA — HTML + CSS + JS em arquivo único)
        ├── ruja_db_layer.js    (embutido no HTML — camada Supabase)
        ├── Supabase (banco principal)
        │     ├── ruja_jovens
        │     ├── ruja_lideres
        │     ├── ruja_departamentos
        │     ├── ruja_frequencias
        │     ├── ruja_recuperacoes
        │     ├── ruja_historico_mensal
        │     ├── ruja_configuracoes
        │     ├── ruja_audit_logs
        │     ├── migration_logs
        │     └── config (legada)
        └── Google Sheets (backup opcional — não fonte principal)
```

---

## Módulos

| Módulo | Página | Funções principais | Status |
|--------|--------|-------------------|--------|
| Autenticação | `loginPage` | `signInWithPassword`, `signOut`, `resetPasswordForEmail` | ✅ |
| Dashboard | `page-dashboard` | `renderDashboard` | ✅ |
| Jovens | `page-jovens` | `saveJovem`, `editJovem`, `deleteJovemUI`, `renderJovens` | ✅ |
| Frequência | `page-frequencia` | `salvarFreqLote`, `getFreqPct`, `calcularStatusAutomatico` | ✅ |
| Recuperação | `page-recuperacao` | `saveRecup`, `editRecup`, `deleteRecup`, `renderRecuperacao` | ✅ |
| Departamentos | `page-departamentos` | `saveDepto`, `editDepto`, `deleteDepto`, `renderDepartamentos` | ✅ |
| Líderes | `page-lideres` | `saveLider`, `editLider`, `deleteLider`, `renderLideres` | ✅ |
| Metas | `page-metas` | `saveMetas`, `renderMetas` | ✅ |
| Aniversariantes | `page-aniversarios` | `getDiasParaAniversario`, `renderAniversarios` | ✅ |
| Configurações | `page-config` | `saveRegras`, `renderConfig`, `importCSV`, `exportCSV` | ✅ |
| Líder Supremo | `page-lidersupremo` | `saveLiderSupremo`, `renderLiderSupremo` | ✅ |
| Alertas | `page-alertas` | `renderAlertas` | ✅ |
| Storage (Fotos) | — | `uploadFotoJovem`, `removerFotoJovem` | ✅ |

---

## Dependências Externas

| Dependência | Versão | Uso |
|-------------|--------|-----|
| Supabase JS | 2.x (CDN) | Banco de dados, Auth, Storage |
| Chart.js | 4.4.1 (CDN) | Gráficos do Dashboard |
| Google Fonts | Barlow, Barlow Condensed | Tipografia |
| Google Sheets (GAS) | opcional | Backup / migração histórica |

---

## Variáveis de Ambiente / Configuração

| Variável | Onde | Descrição |
|----------|------|-----------|
| `_SB_URL` | `ruja_db_layer.js` | URL do projeto Supabase |
| `_SB_KEY` | `ruja_db_layer.js` | Publishable (anon) key do Supabase |
| `FOTO_BUCKET` | `index.html` | Nome do bucket de storage: `ruja-jovens-fotos` |
| `GAS_URL` | `ruja_configuracoes` (chave `gas_url`) | URL do Google Apps Script (backup) |
| `window._agendaUrl` | `ruja_configuracoes` (chave `agenda_url`) | URL da agenda |

---

## Regras de Negócio Críticas

- **Status automático:** calculado por `calcularStatusAutomatico()` após cada frequência salva
  - `Ativo`: frequência ≥ `regras.ativo`% (padrão 75%)
  - `Oscilando`: frequência ≥ `regras.oscilando`% (padrão 40%)
  - `Em Risco`: ≥ `regras.risco` faltas consecutivas (padrão 3)
  - `Ocioso`: demais casos
- **Metas:** armazenadas em `ruja_configuracoes` (chave `metas`, JSONB)
- **Regras:** armazenadas em `ruja_configuracoes` (chave `regras`, JSONB)
- **Fonte de dados:** Supabase é a fonte principal. Google Sheets é backup/importação
- **Snapshot mensal:** `gravarSnapshotMensal()` chamado a cada `saveData()` — só grava se `jovens.length > 0`

---

## Riscos Conhecidos

| Risco | Severidade | Mitigação |
|-------|-----------|-----------|
| IDs como `String(Date.now())` — colisão em multi-usuário | Médio | Migrar para UUID em versão futura |
| Datas armazenadas como TEXT | Médio | Funcional, mas limita queries de range |
| RLS sem isolamento por papel | Médio | Aceitável para uso interno; adicionar roles se necessário |
| `ruja_audit_logs` não utilizada | Baixo | Implementar triggers ou chamadas manuais |
| `localStorage` com dados pessoais | Baixo | Cache local apenas; dados reais no Supabase |

---

## Arquivos no Repositório

| Arquivo | Função |
|---------|--------|
| `index.html` | Aplicação completa (SPA) |
| `ruja_db_layer.js` | Camada de dados Supabase (referência — integrado ao index.html) |
| `ruja_migration_supabase.sql` | Script de criação das tabelas |
| `ruja_foto_sql.sql` | Script de storage e políticas de foto |
| `ruja_fix_metas.sql` | Script de correção emergencial de metas corrompidas |
| `INSTRUCOES_MIGRACAO_RUJA.md` | Guia de migração GAS → Supabase |
| `PROGRAMAS_INFO.md` | Este arquivo |
| `CHANGELOG_RUJA.md` | Histórico de alterações |
| `BANCO_DE_DADOS_RUJA.md` | Schema detalhado do banco |
| `REGRAS_DE_ACESSO_RUJA.md` | Políticas de acesso e RLS |

---

*Última atualização: 2026-05-30 — Auditoria completa e correção de 12 bugs críticos*

---

## Histórico de Atualizações deste Arquivo

| Data | Alteração |
|------|-----------|
| 2026-05-30 | Criação inicial — auditoria completa |
| 2026-05-30 | Adicionados: pullFromSheets (M6), bottom nav (M8), gasUrlInput (B3), _auditLog (B5), login mobile (MOB) |
| 2026-05-30 | Módulos atualizados: Configurações (gasUrlInput editável), Mobile (bottom nav expandido) |

## Bugs Resolvidos (resumo)

| ID | Descrição | Status |
|----|-----------|--------|
| C1-C12 | 12 bugs críticos | ✅ Corrigidos |
| M1-M8 | 8 bugs médios aplicáveis | ✅ Corrigidos |
| B3, B5 | Bugs baixos priorizados | ✅ Corrigidos |
| MOB | Login mobile não funcionava | ✅ Corrigido |

## Riscos Atualizados

| Risco | Severidade | Status |
|-------|-----------|--------|
| IDs como String(Date.now()) | Médio | Aberto — migrar para UUID futuramente |
| Datas como TEXT | Médio | Aberto — funcional mas limita queries |
| RLS sem isolamento por papel | Médio | Aberto — aceitável para uso interno |
| localStorage com dados pessoais | Baixo | Aberto — cache apenas, dados no Supabase |
| ruja_audit_logs cobertura parcial | Baixo | Parcialmente resolvido — B5 expandiu |
