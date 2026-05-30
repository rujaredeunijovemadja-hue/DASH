# BANCO_DE_DADOS_RUJA.md

Documentação completa do schema do banco de dados Supabase do projeto RUJA.

---

## Tabelas

### `ruja_jovens`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | TEXT PK | — | ID único (String(Date.now())) |
| `nome` | TEXT NOT NULL | — | Nome completo |
| `idade` | INTEGER | 0 | Idade (calculada ou manual) |
| `contato` | TEXT | '' | WhatsApp |
| `instagram` | TEXT | '' | @usuario |
| `endereco` | TEXT | '' | Endereço |
| `departamento` | TEXT | '' | Departamentos separados por `;` |
| `lider` | TEXT | '' | Nome do líder responsável |
| `status` | TEXT | 'Em Risco' | Ativo/Oscilando/Ocioso/Em Risco |
| `entrada` | TEXT | '' | Data de entrada (YYYY-MM-DD) |
| `batizado` | TEXT | 'nao' | 'sim' ou 'nao' |
| `data_batismo` | TEXT | '' | Data do batismo (YYYY-MM-DD) |
| `data_nasc` | TEXT | '' | Data de nascimento (YYYY-MM-DD) |
| `obs` | TEXT | '' | Observações |
| `foto_path` | TEXT | '' | Caminho no Storage |
| `foto_url` | TEXT | '' | URL pública da foto |
| `criado_em` | TIMESTAMPTZ | NOW() | — |
| `atualizado_em` | TIMESTAMPTZ | NOW() | — |

**Índices:** `idx_jovens_status (status)`, `idx_jovens_depto (departamento)`
**RLS:** habilitado — policy `auth_all` para `authenticated`

---

### `ruja_lideres`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | TEXT PK | — | ID único |
| `nome` | TEXT NOT NULL | — | Nome completo |
| `contato` | TEXT | '' | WhatsApp |
| `departamento` | TEXT | '' | Departamento principal |
| `funcao` | TEXT | '' | Função/cargo |
| `data_nasc` | TEXT | '' | Data de nascimento (YYYY-MM-DD) |
| `criado_em` | TIMESTAMPTZ | NOW() | — |
| `atualizado_em` | TIMESTAMPTZ | NOW() | — |

**RLS:** habilitado — policy `auth_all`

---

### `ruja_departamentos`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | TEXT PK | — | ID único |
| `nome` | TEXT NOT NULL | — | Nome do departamento |
| `icone` | TEXT | '🏛' | Emoji ícone |
| `lider` | TEXT | — | Nome do líder |
| `capacidade` | INTEGER | 0 | Capacidade máxima |
| `descricao` | TEXT | '' | Descrição |
| `criado_em` | TIMESTAMPTZ | NOW() | — |
| `atualizado_em` | TIMESTAMPTZ | NOW() | — |

**RLS:** habilitado — policy `auth_all`

---

### `ruja_frequencias`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | TEXT PK | — | ID composto `{jovemId}_{data}_{evento}` |
| `jovem_id` | TEXT NOT NULL FK | — | → `ruja_jovens.id` ON DELETE CASCADE |
| `data` | TEXT NOT NULL | — | Data do evento (YYYY-MM-DD) |
| `evento` | TEXT | '' | Nome do evento |
| `presenca` | TEXT | 'falta' | 'presente' ou 'falta' |
| `obs` | TEXT | '' | Observações |
| `criado_em` | TIMESTAMPTZ | NOW() | — |

**Índices:** `idx_freq_jovem (jovem_id)`, `idx_freq_data (data)`
**RLS:** habilitado — policy `auth_all`

---

### `ruja_recuperacoes`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | TEXT PK | — | ID único |
| `jovem_id` | TEXT NOT NULL FK | — | → `ruja_jovens.id` ON DELETE CASCADE |
| `data_inicio` | TEXT | '' | Data de início do plano |
| `lider_resp` | TEXT | '' | Líder responsável |
| `motivo` | TEXT | '' | Motivo da recuperação |
| `status` | TEXT | 'ativo' | 'ativo' ou 'concluido' |
| `obs` | TEXT | '' | Observações |
| `criado_em` | TIMESTAMPTZ | NOW() | — |
| `atualizado_em` | TIMESTAMPTZ | NOW() | — |

**Índices:** `idx_rec_status (status)`
**RLS:** habilitado — policy `auth_all`

---

### `ruja_historico_mensal`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `id` | SERIAL PK | — | Auto-incremento |
| `mes` | TEXT NOT NULL UNIQUE | — | Formato YYYY-MM |
| `ativos_depto` | INTEGER | 0 | Jovens ativos com departamento |
| `batizados_depto` | INTEGER | 0 | Batizados ativos com departamento |
| `total` | INTEGER | 0 | Total de jovens cadastrados |
| `criado_em` | TIMESTAMPTZ | NOW() | — |

**RLS:** habilitado — policy `auth_all`

---

### `ruja_configuracoes`
| Campo | Tipo | Default | Descrição |
|-------|------|---------|-----------|
| `chave` | TEXT PK | — | Nome da configuração |
| `valor_json` | JSONB NOT NULL | — | Valor estruturado |
| `atualizado_em` | TIMESTAMPTZ | NOW() | — |

**Chaves conhecidas:**

| Chave | Estrutura | Descrição |
|-------|-----------|-----------|
| `regras` | `{ativo:75, oscilando:40, risco:3}` | Thresholds de status |
| `metas` | `{ativosDepto:20, batizadosDepto:10}` | Metas do ministério |
| `lider_supremo` | `{nome, contato, instagram, ...}` | Perfil do líder |
| `database_mode` | `"supabase"` | Modo de banco ativo |
| `backup_mode` | `"google_sheets"` | Modo de backup |

**RLS:** habilitado — policy `auth_all`

---

### `config` (legada)
Tabela legada usada antes da migração. Ainda consultada para `agenda_url` e `gas_url`.

| Campo | Tipo |
|-------|------|
| `chave` | TEXT PK |
| `valor` | TEXT |

---

### `ruja_audit_logs` (não utilizada)
Estrutura criada mas sem escrita pelo código atual.

| Campo | Tipo |
|-------|------|
| `id` | SERIAL PK |
| `usuario_id` | UUID FK → auth.users |
| `acao` | TEXT |
| `tabela` | TEXT |
| `registro_id` | TEXT |
| `dados_antes` | JSONB |
| `dados_depois` | JSONB |
| `criado_em` | TIMESTAMPTZ |

---

### `migration_logs`
Logs de execução da migração GAS → Supabase.

| Campo | Tipo |
|-------|------|
| `id` | SERIAL PK |
| `tabela` | TEXT |
| `registros_migrados` | INTEGER |
| `data_execucao` | TIMESTAMPTZ |
| `status` | TEXT |
| `observacao` | TEXT |

---

## Storage

**Bucket:** `ruja-jovens-fotos`
- Acesso: privado (autenticado)
- Limite de arquivo: 2MB
- Tipos aceitos: image/jpeg, image/jpg, image/png, image/webp
- Caminho: `jovens/{jovem_id}/perfil.webp`

---

## Observações de Design

- IDs como `TEXT` com valor `String(Date.now())` — funcional para uso single-user/pequeno grupo; em multi-usuário concorrente pode ter colisão
- Datas armazenadas como `TEXT` no formato `YYYY-MM-DD` — evitar queries `BETWEEN` sem CAST
- `departamento` em `ruja_jovens` pode conter múltiplos valores separados por `;` (ex: `"Teens;Simply;Up"`)

*Última atualização: 2026-05-30*
