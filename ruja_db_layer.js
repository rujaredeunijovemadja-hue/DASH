// ═══════════════════════════════════════════════════════════════
// RUJA — CAMADA DE DADOS SUPABASE
// Substitui o bloco "CONFIG & GAS SYNC" e "DATA STORE" do index.html
// Cole este conteúdo no lugar das linhas 1276-1590 do index.html
// ═══════════════════════════════════════════════════════════════

// ─── SUPABASE CLIENT ───────────────────────────────────────────
const _SB_URL = 'https://wjdnemfifgquuonwwaua.supabase.co';
const _SB_KEY = 'sb_publishable_Qcg3wmK3zXaNYvwt90QUkw_IbUerToD';
const sb = window.supabase.createClient(_SB_URL, _SB_KEY);

// ─── DATA STORE (in-memory) ────────────────────────────────────
let jovens          = [];
let frequencias     = [];
let recuperacoes    = [];
let departamentos   = [];
let lideres         = [];
let liderSupremo    = { nome:'', contato:'', instagram:'', foto:'', descricao:'', dataPosseLider:'', versiculoLider:'', visao:'', tempoNaRuja:'' };
let regras          = { ativo:75, oscilando:40, risco:3 };
let metas           = { ativosDepto:20, batizadosDepto:10 };
let historicoMensal = [];

// GAS como backup opcional
let GAS_URL = '';

// ─── SESSÃO ────────────────────────────────────────────────────
(async function() {
  const { data: { session } } = await sb.auth.getSession();
  if (session) {
    const user = session.user;
    const meta = user.user_metadata || {};
    const nome = meta.nome || user.email.split('@')[0];
    const funcao = meta.funcao || 'Líder';
    mostrarApp(nome, funcao);
  }
})();

// ─── HELPERS ──────────────────────────────────────────────────
function toggleSenha() {
  const inp = document.getElementById('loginPassword');
  const btn = document.getElementById('toggleSenhaBtn');
  inp.type = inp.type === 'password' ? 'text' : 'password';
  btn.textContent = inp.type === 'password' ? '👁️' : '🙈';
}

function mostrarErroLogin(msg) {
  const el = document.getElementById('loginError');
  if (el) { el.textContent = msg; el.style.display = 'block'; }
}

function ocultarErroLogin() {
  const el = document.getElementById('loginError');
  if (el) el.style.display = 'none';
}

// ─── MOSTRAR APP (carrega dados do Supabase) ───────────────────
async function mostrarApp(nome, funcao) {
  const loading    = document.getElementById('loadingScreen');
  const loadingMsg = document.getElementById('loadingMsg');
  const setMsg = (msg) => { if (loadingMsg) loadingMsg.textContent = msg; };

  const forcarAbertura = () => {
    if (loading) loading.style.display = 'none';
    document.getElementById('appShell').style.display = 'flex';
  };

  const emergencia = setTimeout(forcarAbertura, 15000);

  try {
    document.getElementById('loginPage').style.display = 'none';
    loading.style.display = 'flex';
    document.getElementById('userName').textContent    = nome;
    document.getElementById('userRole').textContent    = funcao;
    document.getElementById('userAvatar').textContent  = nome.charAt(0).toUpperCase();

    // 1. Configs globais (agenda_url, gas_url de backup, metas, regras, lider_supremo)
    setMsg('Carregando configurações...');
    await carregarConfigs();

    // 2. Carrega todos os dados do Supabase em paralelo
    setMsg('Carregando dados...');
    await carregarTodos();

    setMsg('Preparando painel...');
    try { renderAll(); } catch(e) { console.warn('renderAll erro:', e); }

    clearTimeout(emergencia);
    loading.style.display = 'none';
    document.getElementById('appShell').style.display = 'flex';
    showSyncBadge('online');

  } catch(e) {
    console.error('mostrarApp erro crítico:', e);
    clearTimeout(emergencia);
    if (loading) loading.style.display = 'none';
    document.getElementById('appShell').style.display = 'flex';
    showSyncBadge('error');
    try { renderAll(); } catch(_) {}
  }
}

// ─── CARREGAR CONFIGS ──────────────────────────────────────────
async function carregarConfigs() {
  try {
    // config antiga (agenda_url, gas_url)
    const { data: configAntiga } = await sb.from('config').select('chave, valor');
    if (configAntiga) {
      configAntiga.forEach(c => {
        if (c.chave === 'agenda_url') window._agendaUrl = c.valor;
        if (c.chave === 'gas_url' && c.valor) GAS_URL = c.valor; // apenas backup
      });
    }

    // configs novas (ruja_configuracoes)
    const { data: configs } = await sb.from('ruja_configuracoes').select('chave, valor_json');
    if (configs) {
      configs.forEach(c => {
        if (c.chave === 'regras')        regras       = c.valor_json;
        if (c.chave === 'metas')         metas        = c.valor_json;
        if (c.chave === 'lider_supremo') liderSupremo = c.valor_json;
      });
    }
  } catch(e) {
    console.warn('Erro ao carregar configs:', e.message);
  }
}

// ─── CARREGAR TODOS OS DADOS ───────────────────────────────────
async function carregarTodos() {
  const [
    rJovens, rFreq, rRec, rDeptos, rLideres, rHist
  ] = await Promise.allSettled([
    sb.from('ruja_jovens').select('*').order('nome'),
    sb.from('ruja_frequencias').select('*').order('data', { ascending: false }),
    sb.from('ruja_recuperacoes').select('*').order('data_inicio', { ascending: false }),
    sb.from('ruja_departamentos').select('*').order('nome'),
    sb.from('ruja_lideres').select('*').order('nome'),
    sb.from('ruja_historico_mensal').select('*').order('mes'),
  ]);

  // Mapeia resultados (snake_case → camelCase para compatibilidade com UI existente)
  if (rJovens.status === 'fulfilled' && rJovens.value.data)
    jovens = rJovens.value.data.map(mapJovem);

  if (rFreq.status === 'fulfilled' && rFreq.value.data)
    frequencias = rFreq.value.data.map(mapFreq);

  if (rRec.status === 'fulfilled' && rRec.value.data)
    recuperacoes = rRec.value.data.map(mapRecup);

  if (rDeptos.status === 'fulfilled' && rDeptos.value.data)
    departamentos = rDeptos.value.data.map(mapDepto);

  if (rLideres.status === 'fulfilled' && rLideres.value.data)
    lideres = rLideres.value.data.map(mapLider);

  if (rHist.status === 'fulfilled' && rHist.value.data)
    historicoMensal = rHist.value.data.map(h => ({
      mes: h.mes,
      ativosDepto:    h.ativos_depto,
      batizadosDepto: h.batizados_depto,
      total:          h.total,
    }));

  console.log(`✅ Dados carregados: ${jovens.length} jovens, ${frequencias.length} freq, ${departamentos.length} deptos`);
}

// ─── MAPPERS snake → camelCase ─────────────────────────────────
const mapJovem = r => ({
  id: r.id, nome: r.nome, idade: r.idade, contato: r.contato,
  instagram: r.instagram, endereco: r.endereco,
  departamento: r.departamento, lider: r.lider, status: r.status,
  entrada: r.entrada, batizado: r.batizado, dataBatismo: r.data_batismo,
  dataNasc: r.data_nasc, obs: r.obs,
});

const mapFreq = r => ({
  id: r.id, jovemId: r.jovem_id, data: r.data,
  evento: r.evento, presenca: r.presenca, obs: r.obs,
});

const mapRecup = r => ({
  id: r.id, jovemId: r.jovem_id, dataInicio: r.data_inicio,
  liderResp: r.lider_resp, motivo: r.motivo, status: r.status, obs: r.obs,
});

const mapDepto = r => ({
  id: r.id, nome: r.nome, icone: r.icone,
  lider: r.lider, capacidade: r.capacidade, desc: r.descricao,
});

const mapLider = r => ({
  id: r.id, nome: r.nome, contato: r.contato,
  departamento: r.departamento, funcao: r.funcao, dataNasc: r.data_nasc,
});

// ─── SAVE DATA (Supabase como primário, localStorage como cache) ──
async function saveData() {
  // Limpa inválidos
  jovens        = jovens.filter(j => j && j.id && j.nome && j.nome.trim());
  lideres       = lideres.filter(l => l && l.id && l.nome && l.nome.trim());
  departamentos = departamentos.filter(d => d && d.id && d.nome && d.nome.trim());
  frequencias   = frequencias.filter(f => f && f.id && f.jovemId);
  recuperacoes  = recuperacoes.filter(r => r && r.id && r.jovemId);

  // Cache local para offline
  localStorage.setItem('ruja_cache_jovens',        JSON.stringify(jovens));
  localStorage.setItem('ruja_cache_departamentos',  JSON.stringify(departamentos));

  // Snapshot mensal
  await gravarSnapshotMensal();
}

// ─── UPSERT JOVEM ─────────────────────────────────────────────
async function upsertJovem(j) {
  const { error } = await sb.from('ruja_jovens').upsert({
    id: j.id, nome: j.nome, idade: parseInt(j.idade)||0,
    contato: j.contato||'', instagram: j.instagram||'',
    endereco: j.endereco||'', departamento: j.departamento||'',
    lider: j.lider||'', status: j.status||'Em Risco',
    entrada: j.entrada||'', batizado: j.batizado||'nao',
    data_batismo: j.dataBatismo||'', data_nasc: j.dataNasc||'', obs: j.obs||'',
    atualizado_em: new Date().toISOString(),
  });
  if (error) throw error;
}

// ─── DELETE JOVEM ─────────────────────────────────────────────
async function deleteJovem(id) {
  const { error } = await sb.from('ruja_jovens').delete().eq('id', id);
  if (error) throw error;
  jovens = jovens.filter(j => j.id !== id);
}

// ─── UPSERT FREQUÊNCIA ────────────────────────────────────────
async function upsertFrequencia(f) {
  const { error } = await sb.from('ruja_frequencias').upsert({
    id: f.id, jovem_id: f.jovemId, data: f.data,
    evento: f.evento||'', presenca: f.presenca||'falta', obs: f.obs||'',
  });
  if (error) throw error;
}

// ─── DELETE FREQUÊNCIA ────────────────────────────────────────
async function deleteFrequencia(id) {
  const { error } = await sb.from('ruja_frequencias').delete().eq('id', id);
  if (error) throw error;
  frequencias = frequencias.filter(f => f.id !== id);
}

// ─── UPSERT RECUPERAÇÃO ───────────────────────────────────────
async function upsertRecuperacao(r) {
  const { error } = await sb.from('ruja_recuperacoes').upsert({
    id: r.id, jovem_id: r.jovemId, data_inicio: r.dataInicio||'',
    lider_resp: r.liderResp||'', motivo: r.motivo||'',
    status: r.status||'ativo', obs: r.obs||'',
    atualizado_em: new Date().toISOString(),
  });
  if (error) throw error;
}

// ─── DELETE RECUPERAÇÃO ───────────────────────────────────────
async function deleteRecuperacao(id) {
  const { error } = await sb.from('ruja_recuperacoes').delete().eq('id', id);
  if (error) throw error;
  recuperacoes = recuperacoes.filter(r => r.id !== id);
}

// ─── UPSERT DEPARTAMENTO ──────────────────────────────────────
async function upsertDepartamento(d) {
  const { error } = await sb.from('ruja_departamentos').upsert({
    id: d.id, nome: d.nome, icone: d.icone||'🏛',
    lider: d.lider||'', capacidade: parseInt(d.capacidade)||0,
    descricao: d.desc||'', atualizado_em: new Date().toISOString(),
  });
  if (error) throw error;
}

// ─── DELETE DEPARTAMENTO ──────────────────────────────────────
async function deleteDepartamento(id) {
  const { error } = await sb.from('ruja_departamentos').delete().eq('id', id);
  if (error) throw error;
  departamentos = departamentos.filter(d => d.id !== id);
}

// ─── UPSERT LÍDER ─────────────────────────────────────────────
async function upsertLider(l) {
  const { error } = await sb.from('ruja_lideres').upsert({
    id: l.id, nome: l.nome, contato: l.contato||'',
    departamento: l.departamento||'', funcao: l.funcao||'',
    data_nasc: l.dataNasc||'', atualizado_em: new Date().toISOString(),
  });
  if (error) throw error;
}

// ─── DELETE LÍDER ─────────────────────────────────────────────
async function deleteLider(id) {
  const { error } = await sb.from('ruja_lideres').delete().eq('id', id);
  if (error) throw error;
  lideres = lideres.filter(l => l.id !== id);
}

// ─── SALVAR CONFIGURAÇÃO ──────────────────────────────────────
async function salvarConfig(chave, valor) {
  const { error } = await sb.from('ruja_configuracoes').upsert({
    chave, valor_json: valor, atualizado_em: new Date().toISOString(),
  });
  if (error) throw error;
}

// ─── SNAPSHOT MENSAL ──────────────────────────────────────────
async function gravarSnapshotMensal() {
  const mesAtual       = new Date().toISOString().slice(0, 7);
  const ativosDepto    = jovens.filter(j => j.status === 'Ativo' && j.departamento).length;
  const batizadosDepto = jovens.filter(j => j.status === 'Ativo' && j.departamento && j.batizado === 'sim').length;
  const total          = jovens.length;

  const snap = { mes: mesAtual, ativosDepto, batizadosDepto, total };

  // Atualiza in-memory
  const idx = historicoMensal.findIndex(h => h.mes === mesAtual);
  if (idx !== -1) historicoMensal[idx] = snap; else historicoMensal.push(snap);
  historicoMensal.sort((a, b) => a.mes.localeCompare(b.mes));

  // Persiste no Supabase
  try {
    await sb.from('ruja_historico_mensal').upsert({
      mes: mesAtual, ativos_depto: ativosDepto,
      batizados_depto: batizadosDepto, total,
    }, { onConflict: 'mes' });
  } catch(e) { console.warn('Erro ao gravar snapshot:', e.message); }
}

// ─── SYNC BADGE (visual) ──────────────────────────────────────
function showSyncBadge(estado) {
  const el = document.getElementById('syncBadge');
  if (!el) return;
  const map = {
    online:  { txt: '✅ Supabase Online',  cor: '#22C55E' },
    syncing: { txt: '🔄 Sincronizando...',  cor: '#F59E0B' },
    error:   { txt: '❌ Erro de conexão',   cor: '#D42B2B' },
    offline: { txt: '📴 Offline',           cor: '#6B7280' },
  };
  const s = map[estado] || map.offline;
  el.textContent = s.txt;
  el.style.color = s.cor;
}

// schedulePush mantido vazio (GAS agora é só backup opcional)
function schedulePush()   { /* GAS = backup opcional */ }
function marcarDadosSujos() { /* não necessário com Supabase em tempo real */ }
function pararAutoSync()  { /* não necessário */ }

// ─── MIGRAÇÃO DO GAS → SUPABASE (executa uma vez) ────────────
async function migrarDoGAS() {
  if (!GAS_URL) { showToast('GAS URL não configurada.'); return; }

  showToast('Iniciando migração do Google Sheets...');
  const btn = document.querySelector('[onclick="migrarDoGAS()"]');
  if (btn) btn.disabled = true;

  try {
    const controller = new AbortController();
    setTimeout(() => controller.abort(), 20000);
    const r   = await fetch(GAS_URL, { signal: controller.signal });
    const d   = await r.json();

    const entidades = [
      { chave: 'jovens',          tabela: 'ruja_jovens',          dados: d.jovens,          mapFn: j => ({
          id: j.id, nome: j.nome, idade: parseInt(j.idade)||0,
          contato: j.contato||'', instagram: j.instagram||'',
          endereco: j.endereco||'', departamento: j.departamento||'',
          lider: j.lider||'', status: j.status||'Em Risco',
          entrada: j.entrada||'', batizado: j.batizado||'nao',
          data_batismo: j.dataBatismo||'', data_nasc: j.dataNasc||'', obs: j.obs||'',
        })
      },
      { chave: 'departamentos',   tabela: 'ruja_departamentos',   dados: d.departamentos,   mapFn: x => ({
          id: x.id, nome: x.nome, icone: x.icone||'🏛',
          lider: x.lider||'', capacidade: parseInt(x.capacidade)||0, descricao: x.desc||'',
        })
      },
      { chave: 'lideres',         tabela: 'ruja_lideres',         dados: d.lideres,         mapFn: x => ({
          id: x.id, nome: x.nome, contato: x.contato||'',
          departamento: x.departamento||'', funcao: x.funcao||'', data_nasc: x.dataNasc||'',
        })
      },
      { chave: 'frequencias',     tabela: 'ruja_frequencias',     dados: d.frequencias,     mapFn: x => ({
          id: x.id, jovem_id: x.jovemId, data: x.data||'',
          evento: x.evento||'', presenca: x.presenca||'falta', obs: x.obs||'',
        })
      },
      { chave: 'recuperacoes',    tabela: 'ruja_recuperacoes',    dados: d.recuperacoes,    mapFn: x => ({
          id: x.id, jovem_id: x.jovemId, data_inicio: x.dataInicio||'',
          lider_resp: x.liderResp||'', motivo: x.motivo||'',
          status: x.status||'ativo', obs: x.obs||'',
        })
      },
    ];

    for (const ent of entidades) {
      if (!ent.dados || !ent.dados.length) {
        await registrarLogMigracao(ent.tabela, 0, 'vazio', 'Sem dados no GAS');
        continue;
      }
      try {
        const payload = ent.dados.filter(x => x && x.id).map(ent.mapFn);
        const { error } = await sb.from(ent.tabela).upsert(payload, { onConflict: 'id' });
        if (error) throw error;
        await registrarLogMigracao(ent.tabela, payload.length, 'sucesso', '');
        showToast(`✅ ${ent.chave}: ${payload.length} registros migrados`);
      } catch(e) {
        await registrarLogMigracao(ent.tabela, 0, 'erro', e.message);
        console.error(`Erro ao migrar ${ent.tabela}:`, e);
      }
    }

    // Recarrega dados do Supabase após migração
    await carregarTodos();
    renderAll();
    showToast('🎉 Migração concluída! Dados agora no Supabase.');

  } catch(e) {
    showToast('Erro na migração: ' + e.message);
    console.error(e);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function registrarLogMigracao(tabela, qtd, status, obs) {
  try {
    await sb.from('migration_logs').insert({
      tabela, registros_migrados: qtd, status, observacao: obs,
    });
  } catch(_) {}
}
