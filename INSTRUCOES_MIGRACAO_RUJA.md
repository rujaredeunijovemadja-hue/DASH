# RUJA — Guia de Migração para Supabase

## Objetivo
Migrar o banco de dados do sistema RUJA de localStorage + Google Apps Script para Supabase como banco principal.

---

## ETAPA 1 — Criar tabelas no Supabase

1. Acesse o **Supabase Dashboard** → seu projeto RUJA
2. Vá em **SQL Editor**
3. Abra o arquivo `ruja_migration_supabase.sql`
4. Execute **um bloco por vez** (cada bloco está separado por `─── BLOCO N ───`)
5. Verifique se cada bloco retornou sem erro antes de continuar

> ⚠️ Execute separadamente para evitar truncamento do SQL Editor

---

## ETAPA 2 — Substituir a camada de dados no index.html

### 2a. Substituir bloco CONFIG & GAS SYNC

No `index.html`, encontre as linhas **1276 a 1590** (do comentário `// CONFIG & GAS SYNC` até o final do DATA STORE).

Substitua **todo esse bloco** pelo conteúdo de `ruja_db_layer.js`.

### 2b. Substituir funções CRUD

Substitua as seguintes funções pelo conteúdo de `ruja_crud_patch.js`:

| Função original          | Localização aprox. |
|--------------------------|--------------------|
| `saveJovem()`            | linha 2935         |
| `deleteJovem(id)`        | linha 2977         |
| `saveDepto()`            | linha 3180         |
| `excluirDepto(id)`       | linha 3200         |
| `saveLider()`            | linha 3475         |
| `deleteLider(id)`        | linha 3536         |
| `saveLiderSupremo()`     | linha 3375         |
| `saveMetas()`            | linha 1809         |
| `saveRegras()`           | buscar no config   |
| `saveRecup()`            | linha 3104         |
| `deleteRecup(id)`        | linha 3140         |

### 2c. Adicionar botão de migração nas Configurações

Adicione este HTML na página de Configurações (`page-config`):

```html
<div class="card" style="margin-top:20px;border-color:rgba(245,158,11,0.3)">
  <div class="card-header">
    <span class="card-title">🔄 Migração do Google Sheets</span>
  </div>
  <p style="font-size:13px;color:var(--text2);margin-bottom:16px">
    Importa todos os dados existentes do Google Sheets para o Supabase.
    Execute apenas uma vez. Após validar, o GAS será usado apenas como backup.
  </p>
  <button class="btn btn-warning" onclick="migrarDoGAS()">
    ▶ Iniciar Migração
  </button>
  <div id="migrationStatus" style="margin-top:12px;font-size:12px;color:var(--text3)"></div>
</div>
```

---

## ETAPA 3 — Executar a migração

1. Faça push do `index.html` atualizado para o GitHub
2. Aguarde o deploy no Vercel
3. Acesse o sistema e faça login
4. Vá em **Configurações** → clique em **▶ Iniciar Migração**
5. Aguarde o processo (pode levar 30–60s dependendo do volume)

---

## ETAPA 4 — Validar

No Supabase Dashboard → **Table Editor**, verifique:

| Tabela              | Qtd esperada         |
|---------------------|----------------------|
| `ruja_jovens`       | = jovens no GAS      |
| `ruja_frequencias`  | = registros no GAS   |
| `ruja_departamentos`| = deptos no GAS      |
| `ruja_lideres`      | = líderes no GAS     |
| `ruja_recuperacoes` | = planos no GAS      |
| `migration_logs`    | registros da migração|

Se os números baterem ✅, o GAS virou backup.

---

## Arquitetura após migração

```
Browser
  └── index.html
        ├── Supabase (banco principal)
        │     ├── ruja_jovens
        │     ├── ruja_frequencias
        │     ├── ruja_departamentos
        │     ├── ruja_lideres
        │     ├── ruja_recuperacoes
        │     ├── ruja_historico_mensal
        │     ├── ruja_configuracoes
        │     └── migration_logs
        │
        └── Google Sheets (backup opcional)
              └── gas_url na tabela config
```

---

## Ganhos reais

- **Multi-usuário real**: dois líderes podem editar ao mesmo tempo
- **Sem perda de dados**: localStorage era por navegador/device
- **Histórico preservado**: `ruja_audit_logs` registra alterações
- **Offline-first**: cache no localStorage como fallback
- **GAS como backup**: exportação para planilha ainda disponível
