// ═══════════════════════════════════════════════════════════════
// RUJA — PATCH DAS FUNÇÕES CRUD (substituir funções existentes)
// Cada função abaixo substitui a versão localStorage/GAS equivalente
// ═══════════════════════════════════════════════════════════════

// ─── SAVE / DELETE JOVEM ──────────────────────────────────────
async function saveJovem() {
  const val = {
    nome:         document.getElementById('jNome').value.trim(),
    idade:        parseInt(document.getElementById('jIdade').value)||null,
    contato:      document.getElementById('jContato').value.trim(),
    instagram:    document.getElementById('jInstagram').value.trim(),
    endereco:     document.getElementById('jEndereco').value.trim(),
    departamento: document.getElementById('jDepartamento').value,
    lider:        document.getElementById('jLider').value,
    status:       document.getElementById('jStatus').value,
    entrada:      dateToISO(document.getElementById('jEntrada').value)||document.getElementById('jEntrada').value,
    batizado:     document.getElementById('jBatizado').value,
    dataBatismo:  dateToISO(document.getElementById('jDataBatismo').value)||document.getElementById('jDataBatismo').value,
    dataNasc:     dateToISO(document.getElementById('jDataNasc').value)||document.getElementById('jDataNasc').value,
    obs:          document.getElementById('jObs').value.trim(),
  };
  if (!val.nome || !val.contato) return alert('Preencha Nome e WhatsApp.');

  if (editingId) {
    const i = jovens.findIndex(j => j.id === editingId);
    if (i !== -1) jovens[i] = { ...jovens[i], ...val };
    val.id = editingId;
  } else {
    val.id = String(Date.now());
    jovens.push(val);
  }

  try {
    await upsertJovem(val);
    await saveData();
    closeModal('modalJovem');
    showToast('Jovem salvo!');
    renderJovens();
  } catch(e) {
    showToast('Erro ao salvar: ' + e.message);
    console.error(e);
  }
}

async function deleteJovem(id) {
  if (!confirm('Excluir este jovem?')) return;
  try {
    await deleteJovem(id); // remove do Supabase e do array
    frequencias  = frequencias.filter(f => f.jovemId !== id);
    recuperacoes = recuperacoes.filter(r => r.jovemId !== id);
    await saveData();
    renderJovens();
    showToast('Jovem excluído');
  } catch(e) {
    showToast('Erro ao excluir: ' + e.message);
  }
}

// ─── SAVE / DELETE DEPARTAMENTO ───────────────────────────────
async function saveDepto() {
  const val = {
    id:         editingDeptoId || String(Date.now()),
    nome:       document.getElementById('dNome').value.trim(),
    icone:      document.getElementById('dIcone').value.trim() || '🏛',
    lider:      document.getElementById('dLider').value,
    capacidade: parseInt(document.getElementById('dCapacidade').value)||0,
    desc:       document.getElementById('dDesc').value.trim(),
  };
  if (!val.nome) return alert('Preencha o nome do departamento.');

  if (editingDeptoId) {
    const i = departamentos.findIndex(d => d.id === editingDeptoId);
    if (i !== -1) departamentos[i] = { ...departamentos[i], ...val };
  } else {
    departamentos.push(val);
  }

  try {
    await upsertDepartamento(val);
    await saveData();
    closeModal('modalDepto');
    renderDepartamentos();
    showToast('Departamento salvo!');
  } catch(e) {
    showToast('Erro ao salvar: ' + e.message);
  }
}

async function excluirDepto(id) {
  if (!confirm('Excluir este departamento?')) return;
  try {
    await deleteDepartamento(id);
    await saveData();
    renderDepartamentos();
    showToast('Departamento excluído');
  } catch(e) {
    showToast('Erro ao excluir: ' + e.message);
  }
}

// ─── SAVE / DELETE LÍDER ──────────────────────────────────────
async function saveLider() {
  const val = {
    id:           editingLiderId || String(Date.now()),
    nome:         document.getElementById('lNome').value.trim(),
    contato:      document.getElementById('lContato').value.trim(),
    departamento: document.getElementById('lDepartamento').value,
    funcao:       document.getElementById('lFuncao').value.trim(),
    dataNasc:     dateToISO(document.getElementById('lDataNasc').value)||document.getElementById('lDataNasc').value,
  };
  if (!val.nome) return alert('Preencha o nome do líder.');

  if (editingLiderId) {
    const i = lideres.findIndex(l => l.id === editingLiderId);
    if (i !== -1) lideres[i] = { ...lideres[i], ...val };
  } else {
    lideres.push(val);
  }

  try {
    await upsertLider(val);
    await saveData();
    closeModal('modalLider');
    renderLideres();
    showToast('Líder salvo!');
  } catch(e) {
    showToast('Erro ao salvar: ' + e.message);
  }
}

async function deleteLider(id) {
  if (!confirm('Excluir este líder?')) return;
  try {
    await deleteLider(id);
    await saveData();
    renderLideres();
    showToast('Líder excluído');
  } catch(e) {
    showToast('Erro ao excluir: ' + e.message);
  }
}

// ─── SAVE LÍDER SUPREMO ───────────────────────────────────────
async function saveLiderSupremo() {
  liderSupremo = {
    nome:           document.getElementById('lsNome').value.trim(),
    contato:        document.getElementById('lsContato').value.trim(),
    instagram:      document.getElementById('lsInstagram').value.trim(),
    foto:           document.getElementById('lsFoto').value.trim(),
    descricao:      document.getElementById('lsDescricao').value.trim(),
    dataPosseLider: dateToISO(document.getElementById('lsDataPosse').value)||document.getElementById('lsDataPosse').value,
    versiculoLider: document.getElementById('lsVersiculo').value.trim(),
    visao:          document.getElementById('lsVisao').value.trim(),
    tempoNaRuja:    document.getElementById('lsTempoNaRuja').value.trim(),
  };
  try {
    await salvarConfig('lider_supremo', liderSupremo);
    await saveData();
    renderLiderSupremo();
    showToast('Líder Supremo salvo!');
  } catch(e) {
    showToast('Erro ao salvar: ' + e.message);
  }
}

// ─── SAVE METAS ───────────────────────────────────────────────
async function saveMetas() {
  metas.ativosDepto    = parseInt(document.getElementById('metaAtivosDepto').value)||20;
  metas.batizadosDepto = parseInt(document.getElementById('metaBatizadosDepto').value)||10;
  try {
    await salvarConfig('metas', metas);
    await saveData();
    showToast('Metas salvas!');
    if (currentPage === 'metas') renderMetas();
  } catch(e) {
    showToast('Erro ao salvar metas: ' + e.message);
  }
}

// ─── SAVE REGRAS ──────────────────────────────────────────────
async function saveRegras() {
  regras.ativo    = parseInt(document.getElementById('regraAtivo').value)||75;
  regras.oscilando= parseInt(document.getElementById('regraOscilando').value)||40;
  regras.risco    = parseInt(document.getElementById('regraRisco').value)||3;
  try {
    await salvarConfig('regras', regras);
    showToast('Regras salvas!');
  } catch(e) {
    showToast('Erro ao salvar regras: ' + e.message);
  }
}

// ─── SALVAR FREQUÊNCIA INLINE ─────────────────────────────────
async function salvarFreqInline(jovemId, data, evento, presenca) {
  const id = `${jovemId}_${data}_${evento}`.replace(/\s/g,'_');
  const reg = { id, jovemId, data, evento, presenca, obs: '' };

  const idx = frequencias.findIndex(f => f.id === id);
  if (idx !== -1) frequencias[idx] = reg;
  else frequencias.push(reg);

  try {
    await upsertFrequencia(reg);
    await saveData();
  } catch(e) {
    console.error('Erro ao salvar frequência:', e);
    showToast('Erro ao salvar frequência');
  }
}

// ─── SAVE RECUPERAÇÃO ────────────────────────────────────────
async function saveRecup() {
  const val = {
    id:          editingRecupId || String(Date.now()),
    jovemId:     document.getElementById('rJovem').value,
    dataInicio:  dateToISO(document.getElementById('rDataInicio').value)||document.getElementById('rDataInicio').value,
    liderResp:   document.getElementById('rLiderResp').value,
    motivo:      document.getElementById('rMotivo').value.trim(),
    status:      document.getElementById('rStatus').value,
    obs:         document.getElementById('rObs').value.trim(),
  };
  if (!val.jovemId) return alert('Selecione um jovem.');

  if (editingRecupId) {
    const i = recuperacoes.findIndex(r => r.id === editingRecupId);
    if (i !== -1) recuperacoes[i] = { ...recuperacoes[i], ...val };
  } else {
    recuperacoes.push(val);
  }

  try {
    await upsertRecuperacao(val);
    await saveData();
    closeModal('modalRecup');
    renderRecuperacao();
    showToast('Plano salvo!');
  } catch(e) {
    showToast('Erro ao salvar: ' + e.message);
  }
}

async function deleteRecup(id) {
  if (!confirm('Excluir este plano?')) return;
  try {
    await deleteRecuperacao(id);
    await saveData();
    renderRecuperacao();
    showToast('Plano excluído');
  } catch(e) {
    showToast('Erro ao excluir: ' + e.message);
  }
}
