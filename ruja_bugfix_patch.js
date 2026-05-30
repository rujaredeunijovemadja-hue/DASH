// ═══════════════════════════════════════════════════════════════════════
// RUJA — PATCH DE CORREÇÕES CRÍTICAS
// Versão: 2025-01
// Instruções: cole este bloco no index.html antes da tag </body>
// OU substitua as funções listadas no index.html pelas versões corrigidas
// ═══════════════════════════════════════════════════════════════════════

// ───────────────────────────────────────────────────────────────────────
// BUG 1 — ANIVERSÁRIOS: getDiasParaAniversario() não detecta "hoje"
// CAUSA: (aniv < hoje) usa comparação de objetos Date sem zeragem de horas.
//   Quando aniv == hoje (mesmo dia), (aniv < hoje) pode ser false porque
//   `hoje` inclui horas/minutos/segundos, então aniv fica no passado e
//   avança para o PRÓXIMO ANO → exibe "em 336 dias" em vez de "Hoje".
// FIX: zerar horas em ambos e comparar via getTime().
// ───────────────────────────────────────────────────────────────────────
function getDiasParaAniversario(dataNasc) {
  if (!dataNasc) return 999;

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); // ← zera horas

  const partes = dataNasc.split('-');
  const mes    = parseInt(partes[1]) - 1;
  const dia    = parseInt(partes[2]);

  const aniv = new Date(hoje.getFullYear(), mes, dia);
  aniv.setHours(0, 0, 0, 0); // ← zera horas

  if (aniv.getTime() === hoje.getTime()) return 0; // ← HOJE

  if (aniv < hoje) {
    // já passou este ano → calcular para o próximo
    aniv.setFullYear(hoje.getFullYear() + 1);
  }

  return Math.ceil((aniv - hoje) / (1000 * 60 * 60 * 24));
}

// Também corrigir a exibição nos cards de aniversário
// Substituir a função cardAniversario existente (dentro de renderAniversarios):
// Procure o trecho `const diasLabel = ...` e substitua pela lógica abaixo.
// Se preferir, substitua renderAniversarios inteira com esta versão segura:
function _diasLabel(dias) {
  if (dias === 0)  return '🎉 Hoje!';
  if (dias === 1)  return 'Amanhã';
  if (dias <= 7)   return `em ${dias} dias`;
  if (dias <= 30)  return `em ${dias} dias`;
  return `em ${dias} dias`;
}

// ───────────────────────────────────────────────────────────────────────
// BUG 2 — METAS COM VALOR NEGATIVO/ABSURDO (-2207422412000)
// CAUSA: Supabase retorna `valor_json` como objeto JS (JSONB). Quando a
//   tabela ruja_configuracoes não tem a chave 'metas', carregarConfigs()
//   não define `metas`, mantendo o valor inicial. Porém se algum código
//   antigo (index.html pré-patch) ler de `config` (tabela legada), pode
//   haver um campo chamado "meta_ativos_departamento" gravado como DATE
//   ou TIMESTAMP (epoch em ms), gerando o número negativo.
//
// FIX duplo:
//   1. Garantir parsing robusto ao ler de ruja_configuracoes
//   2. Sanitizar metas ao carregar (nunca aceitar valores < 0 ou > 10000)
// ───────────────────────────────────────────────────────────────────────
async function carregarConfigs() {
  try {
    // Config legada (agenda_url, gas_url)
    const { data: configAntiga } = await sb.from('config').select('chave, valor');
    if (configAntiga) {
      configAntiga.forEach(c => {
        if (c.chave === 'agenda_url') window._agendaUrl = c.valor;
        if (c.chave === 'gas_url' && c.valor) GAS_URL = c.valor;
      });
    }

    // Configs novas (ruja_configuracoes)
    const { data: configs } = await sb.from('ruja_configuracoes').select('chave, valor_json');
    if (configs) {
      configs.forEach(c => {
        // valor_json já é objeto JS quando vem de JSONB — nenhum JSON.parse necessário
        const val = (typeof c.valor_json === 'string') ? JSON.parse(c.valor_json) : c.valor_json;

        if (c.chave === 'regras') {
          regras = {
            ativo:     _sanitizarInteiro(val.ativo,    75, 0, 100),
            oscilando: _sanitizarInteiro(val.oscilando, 40, 0, 100),
            risco:     _sanitizarInteiro(val.risco,     3,  0, 100),
          };
        }
        if (c.chave === 'metas') {
          metas = {
            ativosDepto:    _sanitizarInteiro(val.ativosDepto,    20, 1, 9999),
            batizadosDepto: _sanitizarInteiro(val.batizadosDepto, 10, 0, 9999),
          };
        }
        if (c.chave === 'lider_supremo') liderSupremo = val;
      });
    }

    console.log('✅ Configs carregadas — metas:', metas, '| regras:', regras);
  } catch(e) {
    console.warn('⚠️ Erro ao carregar configs:', e.message);
  }
}

function _sanitizarInteiro(val, padrao, min, max) {
  const n = parseInt(val);
  if (isNaN(n) || n < min || n > max) return padrao;
  return n;
}

// ───────────────────────────────────────────────────────────────────────
// BUG 3 — BOTÕES DE EDIÇÃO: IDs numéricos vs string causam .find() falhar
// CAUSA: O Supabase retorna `id` como string (TEXT). Porém o código faz:
//   const j = jovens.find(x => x.id === id)
//   Mas no HTML o onclick passa: editJovem(${j.id})
//   Se j.id veio como número (parseInt em algum mapper antigo) e o
//   parâmetro id chega como string (do onclick), a comparação === falha.
//
// FIX: normalizar todos os .find() para comparação tolerante de tipo.
// Substituir todas as funções de edição com versões seguras:
// ───────────────────────────────────────────────────────────────────────
function editJovem(id) {
  const sid = String(id);
  const j = jovens.find(x => String(x.id) === sid);
  if (!j) { console.warn('editJovem: jovem não encontrado, id=', id); return; }
  editingId = sid;
  openModal('modalJovem');
  setTimeout(() => {
    document.getElementById('jNome').value         = j.nome || '';
    document.getElementById('jIdade').value        = j.idade || '';
    document.getElementById('jContato').value      = j.contato || '';
    document.getElementById('jInstagram').value    = j.instagram || '';
    document.getElementById('jEndereco').value     = j.endereco || '';
    try { setDeptoCheckboxes(j.departamento || ''); } catch(e) {
      try { document.getElementById('jDepartamento').value = j.departamento || ''; } catch(_) {}
    }
    document.getElementById('jLider').value        = j.lider || '';
    document.getElementById('jStatus').value       = j.status || '';
    document.getElementById('jEntrada').value      = dateFromISO(j.entrada) || j.entrada || '';
    document.getElementById('jObs').value          = j.obs || '';
    document.getElementById('jBatizado').value     = j.batizado || 'nao';
    document.getElementById('jDataBatismo').value  = dateFromISO(j.dataBatismo) || j.dataBatismo || '';
    document.getElementById('jDataNasc').value     = dateFromISO(j.dataNasc) || j.dataNasc || '';
    try { document.getElementById('modalJovemTitle').textContent = 'Editar Jovem'; } catch(_) {}
    console.log('[editJovem] aberto para id=' + sid);
  }, 60);
}

function editDepto(id) {
  const sid = String(id);
  const d = departamentos.find(x => String(x.id) === sid);
  if (!d) { console.warn('editDepto: departamento não encontrado, id=', id); return; }
  editingDeptoId = sid;
  openModal('modalDepto');
  setTimeout(() => {
    document.getElementById('dNome').value      = d.nome || '';
    document.getElementById('dIcone').value     = d.icone || '🏛';
    document.getElementById('dLider').value     = d.lider || '';
    document.getElementById('dCapacidade').value = d.capacidade || '';
    document.getElementById('dDesc').value      = d.desc || '';
    try { document.getElementById('modalDeptoTitle').textContent = 'Editar Departamento'; } catch(_) {}
    console.log('[editDepto] aberto para id=' + sid);
  }, 60);
}

function editLider(id) {
  const sid = String(id);
  const l = lideres.find(x => String(x.id) === sid);
  if (!l) { console.warn('editLider: líder não encontrado, id=', id); return; }
  editingLiderId = sid;
  openModal('modalLider');
  setTimeout(() => {
    document.getElementById('lNome').value         = l.nome || '';
    document.getElementById('lContato').value      = l.contato || '';
    document.getElementById('lDepartamento').value = l.departamento || '';
    document.getElementById('lFuncao').value       = l.funcao || '';
    document.getElementById('lDataNasc').value     = dateFromISO(l.dataNasc) || l.dataNasc || '';
    try { document.getElementById('modalLiderTitle').textContent = 'Editar Líder'; } catch(_) {}
    console.log('[editLider] aberto para id=' + sid);
  }, 60);
}

function editRecup(id) {
  const sid = String(id);
  const r = recuperacoes.find(x => String(x.id) === sid);
  if (!r) { console.warn('editRecup: recuperação não encontrada, id=', id); return; }
  editingRecupId = sid;
  openModal('modalRecup');
  setTimeout(() => {
    try { document.getElementById('rJovem').value      = r.jovemId || ''; } catch(_) {}
    try { document.getElementById('rDataInicio').value = dateFromISO(r.dataInicio) || r.dataInicio || ''; } catch(_) {}
    try { document.getElementById('rLiderResp').value  = r.liderResp || ''; } catch(_) {}
    try { document.getElementById('rMotivo').value     = r.motivo || ''; } catch(_) {}
    try { document.getElementById('rStatus').value     = r.status || 'ativo'; } catch(_) {}
    try { document.getElementById('rObs').value        = r.obs || ''; } catch(_) {}
    try { document.getElementById('modalRecupTitle').textContent = 'Editar Plano'; } catch(_) {}
    console.log('[editRecup] aberto para id=' + sid);
  }, 60);
}

// ───────────────────────────────────────────────────────────────────────
// BUG 4 — SAVE JOVEM: editingId pode ser número mas Supabase espera string
// FIX: normalizar editingId ao comparar e ao fazer upsert
// ───────────────────────────────────────────────────────────────────────
async function saveJovem() {
  const val = {
    nome:         (document.getElementById('jNome')?.value || '').trim(),
    idade:        parseInt(document.getElementById('jIdade')?.value) || null,
    contato:      (document.getElementById('jContato')?.value || '').trim(),
    instagram:    (document.getElementById('jInstagram')?.value || '').trim(),
    endereco:     (document.getElementById('jEndereco')?.value || '').trim(),
    departamento: document.getElementById('jDepartamento')?.value || '',
    lider:        document.getElementById('jLider')?.value || '',
    status:       document.getElementById('jStatus')?.value || 'Em Risco',
    entrada:      dateToISO(document.getElementById('jEntrada')?.value) || document.getElementById('jEntrada')?.value || '',
    batizado:     document.getElementById('jBatizado')?.value || 'nao',
    dataBatismo:  dateToISO(document.getElementById('jDataBatismo')?.value) || document.getElementById('jDataBatismo')?.value || '',
    dataNasc:     dateToISO(document.getElementById('jDataNasc')?.value) || document.getElementById('jDataNasc')?.value || '',
    obs:          (document.getElementById('jObs')?.value || '').trim(),
  };

  // Tentar ler departamento via checkboxes (se existir o método)
  try {
    const deptos = [];
    document.querySelectorAll('.depto-checkbox:checked').forEach(cb => deptos.push(cb.value));
    if (deptos.length) val.departamento = deptos.join(';');
  } catch(_) {}

  if (!val.nome || !val.contato) return alert('Preencha Nome e WhatsApp.');

  if (editingId) {
    const sid = String(editingId);
    const i = jovens.findIndex(j => String(j.id) === sid);
    val.id = sid;
    if (i !== -1) jovens[i] = { ...jovens[i], ...val };
    else jovens.push(val);
  } else {
    val.id = String(Date.now());
    jovens.push(val);
  }

  try {
    console.log('[saveJovem] salvando:', val);
    const { error } = await sb.from('ruja_jovens').upsert({
      id: val.id, nome: val.nome, idade: parseInt(val.idade) || 0,
      contato: val.contato, instagram: val.instagram || '',
      endereco: val.endereco || '', departamento: val.departamento || '',
      lider: val.lider || '', status: val.status || 'Em Risco',
      entrada: val.entrada || '', batizado: val.batizado || 'nao',
      data_batismo: val.dataBatismo || '', data_nasc: val.dataNasc || '',
      obs: val.obs || '', atualizado_em: new Date().toISOString(),
    });
    if (error) throw error;
    console.log('[saveJovem] ✅ salvo no Supabase');
    closeModal('modalJovem');
    showToast('Jovem salvo!');
    renderJovens();
  } catch(e) {
    showToast('❌ Erro ao salvar: ' + e.message);
    console.error('[saveJovem] erro:', e);
  }
}

// ───────────────────────────────────────────────────────────────────────
// BUG 5 — FILTROS: status inconsistente ("Recebido" vs "recebido")
// FIX: normalizar comparações de status para lowercase
// ───────────────────────────────────────────────────────────────────────
function getStatusJovem(j) {
  // Compatibilidade com valores antigos
  const s = (j.status || '').toLowerCase().trim();
  if (s === 'ativo')     return 'Ativo';
  if (s === 'oscilando') return 'Oscilando';
  if (s === 'ocioso')    return 'Ocioso';
  if (s === 'em risco')  return 'Em Risco';
  if (s === 'recebido')  return 'Ativo'; // legado → normaliza
  return j.status || 'Em Risco';
}

// ───────────────────────────────────────────────────────────────────────
// BUG 6 — GOOGLE SHEETS COMO FONTE: garantir que GAS nunca seja chamado
//   automaticamente como fonte de dados (apenas como importação manual)
// FIX: sobrescrever funções de sync automático para serem no-op
// ───────────────────────────────────────────────────────────────────────
function schedulePush()    { /* GAS = backup opcional. Não sincronizar automaticamente. */ }
function marcarDadosSujos(){ /* Não necessário com Supabase em tempo real. */ }
function pararAutoSync()   { /* Não necessário. */ }

// Bloquear qualquer chamada de leitura direta do GAS como fonte
const _gasReadBlocked = true;
async function lerDadosGAS() {
  console.warn('⛔ lerDadosGAS() bloqueado. O app lê do Supabase. Use "Importar do Sheets" para migração manual.');
  return null;
}

// ───────────────────────────────────────────────────────────────────────
// BOTÕES DE SINCRONIZAÇÃO — feedback visual melhorado
// ───────────────────────────────────────────────────────────────────────
async function forcarPullSupabase() {
  showToast('🔄 Buscando dados do Supabase...');
  try {
    await carregarTodos();
    renderAll();
    showToast('✅ Dados atualizados do Supabase!');
    console.log('[forcarPull] dados recarregados do Supabase');
  } catch(e) {
    showToast('❌ Erro ao buscar dados: ' + e.message);
    console.error('[forcarPull]', e);
  }
}

async function importarSheetsParaSupabase() {
  if (!GAS_URL) {
    showToast('⚠️ URL do Google Sheets não configurada.');
    return;
  }
  showToast('📥 Importando do Sheets...');
  try {
    const resp = await fetch(GAS_URL + '?acao=exportar');
    if (!resp.ok) throw new Error('Falha na requisição: ' + resp.status);
    const dados = await resp.json();
    // Aqui seria feita a migração completa via migrarDoGAS()
    await migrarDoGAS(dados);
    showToast('✅ Importado com sucesso!');
  } catch(e) {
    showToast('❌ Erro ao importar: ' + e.message);
    console.error('[importarSheets]', e);
  }
}

// ───────────────────────────────────────────────────────────────────────
// LOGS DE DIAGNÓSTICO (temporários — remover após estabilizar)
// ───────────────────────────────────────────────────────────────────────
const _origOpenModal = window.openModal;
if (typeof _origOpenModal === 'function') {
  window.openModal = function(id) {
    console.log('[openModal] chamado:', id, '| editingId:', editingId, '| editingDeptoId:', editingDeptoId, '| editingLiderId:', editingLiderId);
    return _origOpenModal(id);
  };
}

console.log('✅ RUJA Bugfix Patch carregado — aniversários, metas, botões de edição e fonte de dados corrigidos.');
