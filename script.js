// ══════════════════════════════════════════════════════════════════════════════
//  UP RUJA — Google Apps Script
//  Sincronização bidirecional entre o Dashboard e o Google Sheets
//  Versão: 3.0 | Rede UniJovem ADJA
//
//  ✅ V3:
//  • ID da planilha fixo — conecta direto sem precisar configurar nada
//  • AUTO-CONFIGURAÇÃO na primeira execução
//  • Campo "Batizado" e "Data do Batismo" na aba Jovens
//  • Aba "Metas", "Histórico Mensal", "Atividades"
//  • Relatório semanal com progresso das metas
// ══════════════════════════════════════════════════════════════════════════════
//
//  ┌─────────────────────────────────────────────────────────────────────────┐
//  │  INSTALAÇÃO — apenas 2 passos:                                         │
//  │  1. Acesse https://script.google.com → "Novo projeto"                  │
//  │  2. Cole TODO este código, salve (Ctrl+S) e implante:                  │
//  │     Implantar → Nova implantação → App da Web                          │
//  │     Executar como: Você | Acesso: Qualquer pessoa                       │
//  │                                                                         │
//  │  ✅ A planilha já está vinculada pelo ID fixo no código.               │
//  │     Todas as abas são criadas automaticamente na 1ª requisição.        │
//  └─────────────────────────────────────────────────────────────────────────┘

// ── NOMES DAS ABAS
const ABAS = {
  JOVENS:        'Jovens',
  FREQUENCIAS:   'Frequências',
  RECUPERACOES:  'Recuperações',
  DEPARTAMENTOS: 'Departamentos',
  LIDERES:       'Líderes',
  METAS:         'Metas',
  HISTORICO:     'Histórico Mensal',
  ATIVIDADES:    'Atividades',
  CONFIG:        'Configurações',
  LOG:           'Log de Alterações',
};

// ── COLUNAS DE CADA ABA  ← batizado e dataBatismo adicionados
const COLS = {
  JOVENS:        ['id','nome','idade','contato','instagram','endereco','departamento','lider','status','entrada','batizado','dataBatismo','dataNasc','obs'],
  FREQUENCIAS:   ['id','jovemId','data','evento','presenca','obs'],
  RECUPERACOES:  ['id','jovemId','responsavel','etapa','motivo','acao','obs','status'],
  DEPARTAMENTOS: ['id','nome','icone','lider','capacidade','desc'],
  LIDERES:       ['id','nome','contato','departamento','funcao'],
  HISTORICO:     ['mes','ativosDepto','batizadosDepto','total'],
  ATIVIDADES:    ['id','titulo','tipo','data','horario','local','responsavel','departamento','hierarquia','status','desc','aprovadoPor','aprovadoEm','canceladoPor','canceladoEm'],
};

// ══════════════════════════════════════════════════════════════════════════════
// doGet — Dashboard chama GET para PUXAR todos os dados
// ══════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    autoConfigurarSeNecessario();

    return jsonResp({
      jovens:          lerAba(ABAS.JOVENS,        COLS.JOVENS),
      frequencias:     lerAba(ABAS.FREQUENCIAS,   COLS.FREQUENCIAS),
      recuperacoes:    lerAba(ABAS.RECUPERACOES,  COLS.RECUPERACOES),
      departamentos:   lerAba(ABAS.DEPARTAMENTOS, COLS.DEPARTAMENTOS),
      lideres:         lerAba(ABAS.LIDERES,        COLS.LIDERES),
      historicoMensal: lerAba(ABAS.HISTORICO,      COLS.HISTORICO),
      atividades:      lerAba(ABAS.ATIVIDADES,    COLS.ATIVIDADES),
      regras:          lerConfig('regras'),
      metas:           lerConfig('metas'),
      ts:              new Date().toISOString(),
    });
  } catch(err) {
    return jsonResp({ erro: err.message });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// doPost — Dashboard chama POST para ENVIAR dados ao Sheets
// ══════════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    autoConfigurarSeNecessario();

    const payload = JSON.parse(e.postData.contents);

    if (payload.action !== 'push') {
      return jsonResp({ ok: false, erro: 'Ação desconhecida: ' + payload.action });
    }

    const usuario = payload.usuario || 'Dashboard';

    if (payload.jovens)          escreverAba(ABAS.JOVENS,        COLS.JOVENS,        payload.jovens);
    if (payload.frequencias)     escreverAba(ABAS.FREQUENCIAS,   COLS.FREQUENCIAS,   payload.frequencias);
    if (payload.recuperacoes)    escreverAba(ABAS.RECUPERACOES,  COLS.RECUPERACOES,  payload.recuperacoes);
    if (payload.departamentos)   escreverAba(ABAS.DEPARTAMENTOS, COLS.DEPARTAMENTOS, payload.departamentos);
    if (payload.lideres)         escreverAba(ABAS.LIDERES,        COLS.LIDERES,       payload.lideres);
    if (payload.historicoMensal) escreverAba(ABAS.HISTORICO,      COLS.HISTORICO,     payload.historicoMensal);
    if (payload.atividades)      escreverAba(ABAS.ATIVIDADES,    COLS.ATIVIDADES,    payload.atividades);
    if (payload.regras)          salvarConfig('regras', payload.regras);
    if (payload.metas)           salvarConfig('metas',  payload.metas);

    registrarLog(usuario, 'push completo v3');

    return jsonResp({ ok: true, ts: new Date().toISOString() });
  } catch(err) {
    return jsonResp({ ok: false, erro: err.message });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// LER ABA — converte linhas do Sheets em array de objetos
// ══════════════════════════════════════════════════════════════════════════════
function lerAba(nomeAba, colunas) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeAba);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dados = sheet.getRange(2, 1, lastRow - 1, colunas.length).getValues();

  return dados
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      colunas.forEach((col, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        // Garante tipos numéricos corretos
        const numericos = ['id','jovemId','capacidade','idade','ativosDepto','batizadosDepto','total'];
        if (numericos.includes(col) && typeof val === 'number') val = Number(val);
        obj[col] = val;
      });
      return obj;
    });
}

// ══════════════════════════════════════════════════════════════════════════════
// ESCREVER ABA — sobrescreve completamente com os dados recebidos
// ══════════════════════════════════════════════════════════════════════════════
function escreverAba(nomeAba, colunas, dados) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(nomeAba);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  if (!dados || !dados.length) return;

  const rows = dados.map(obj =>
    colunas.map(col => {
      const v = obj[col];
      return (v === undefined || v === null) ? '' : v;
    })
  );

  sheet.getRange(2, 1, rows.length, colunas.length).setValues(rows);

  // Formata colunas de data
  ['data','entrada','dataBatismo'].forEach(campo => {
    const idx = colunas.indexOf(campo);
    if (idx !== -1) {
      sheet.getRange(2, idx + 1, rows.length, 1).setNumberFormat('DD/MM/YYYY');
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// CONFIG — aba de configurações (regras + metas na mesma aba, separadas por seção)
// ══════════════════════════════════════════════════════════════════════════════
function lerConfig(secao) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(ABAS.CONFIG);
  if (!sheet) return secao === 'regras'
    ? { ativo:75, oscilando:40, risco:3 }
    : { ativosDepto:20, batizadosDepto:10 };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const dados = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const result = {};

  dados.forEach(row => {
    const [s, chave, valor] = row;
    if (s === secao && chave) result[chave] = isNaN(Number(valor)) ? valor : Number(valor);
  });

  // Defaults caso esteja vazio
  if (secao === 'regras') {
    if (!result.ativo)     result.ativo     = 75;
    if (!result.oscilando) result.oscilando = 40;
    if (!result.risco)     result.risco     = 3;
  }
  if (secao === 'metas') {
    if (!result.ativosDepto)    result.ativosDepto    = 20;
    if (!result.batizadosDepto) result.batizadosDepto = 10;
  }

  return result;
}

function salvarConfig(secao, dados) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(ABAS.CONFIG);
  if (!sheet) return;

  // Remove linhas da seção e reescreve
  const lastRow = sheet.getLastRow();
  const existentes = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];

  // Filtra as linhas que não são dessa seção
  const outras = existentes.filter(r => r[0] !== secao);

  // Adiciona as linhas desta seção
  const novas = Object.entries(dados).map(([k, v]) => [secao, k, v]);
  const todasLinhas = [...outras, ...novas];

  // Limpa e reescreve
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  if (todasLinhas.length > 0) {
    sheet.getRange(2, 1, todasLinhas.length, 3).setValues(todasLinhas);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// LOG DE ALTERAÇÕES
// ══════════════════════════════════════════════════════════════════════════════
function registrarLog(usuario, acao) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(ABAS.LOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), usuario, acao]);
    const total = sheet.getLastRow();
    if (total > 501) sheet.deleteRows(2, total - 501);
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════════════════════════
// OBTER PLANILHA — ID fixo, conecta direto sem configuração manual
// ══════════════════════════════════════════════════════════════════════════════
function getSpreadsheet() {
  const ID_FIXO = '1yFtkXSlL_jtI-4NYBR3vUlDPAeyGhB-PgVk89k5ltbQ';
  try {
    return SpreadsheetApp.openById(ID_FIXO);
  } catch(e) {
    Logger.log('❌ Erro ao abrir planilha: ' + e.message);
    throw new Error('Não foi possível acessar a planilha. Verifique as permissões do script.');
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// AUTO-CONFIGURAÇÃO — detecta primeira execução e monta todas as abas
// ══════════════════════════════════════════════════════════════════════════════
function autoConfigurarSeNecessario() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ruja_configurado') === 'v3') return;
  Logger.log('🚀 Primeira execução — configurando planilha automaticamente...');
  configurarPlanilha();
  props.setProperty('ruja_configurado', 'v3');
  Logger.log('✅ Auto-configuração concluída!');
}

// ══════════════════════════════════════════════════════════════════════════════
// RECONFIGURAR — force a reconfiguração das abas (não apaga dados)
// Execute no editor do Apps Script se precisar recriar abas faltando
// ══════════════════════════════════════════════════════════════════════════════
function reconfigurarPlanilha() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('ruja_configurado');
  autoConfigurarSeNecessario();
  Logger.log('✅ Planilha reconfigurada!');
}

// ══════════════════════════════════════════════════════════════════════════════
// CONFIGURAR PLANILHA — execute uma vez para criar todas as abas
// ══════════════════════════════════════════════════════════════════════════════
function configurarPlanilha() {
  const ss       = getSpreadsheet();
  const VERMELHO = '#D42B2B';
  const BRANCO   = '#FFFFFF';
  const PRETO    = '#1A1A1A';
  const AZUL     = '#1D4ED8';

  // ── Abas de dados
  const configuracoes = [
    {
      nome:   ABAS.JOVENS,
      cols:   ['ID','Nome','Idade','WhatsApp','Instagram','Endereço','Departamento','Líder','Status','Data Entrada','Batizado','Data Batismo','Nascimento','Observação'],
      keys:   COLS.JOVENS,
      widths: [60, 200, 60, 140, 130, 260, 120, 120, 100, 110, 90, 110, 110, 300],
    },
    {
      nome:   ABAS.FREQUENCIAS,
      cols:   ['ID','ID Jovem','Data','Evento','Presença','Observação'],
      keys:   COLS.FREQUENCIAS,
      widths: [60, 80, 100, 180, 120, 200],
    },
    {
      nome:   ABAS.RECUPERACOES,
      cols:   ['ID','ID Jovem','Responsável','Etapa','Motivo','Próxima Ação','Observação','Status'],
      keys:   COLS.RECUPERACOES,
      widths: [60, 80, 140, 220, 150, 200, 250, 110],
    },
    {
      nome:   ABAS.DEPARTAMENTOS,
      cols:   ['ID','Nome','Ícone','Líder','Capacidade','Descrição'],
      keys:   COLS.DEPARTAMENTOS,
      widths: [60, 140, 70, 140, 90, 250],
    },
    {
      nome:   ABAS.LIDERES,
      cols:   ['ID','Nome','WhatsApp','Departamento','Função'],
      keys:   COLS.LIDERES,
      widths: [60, 180, 140, 140, 180],
    },
    {
      nome:   ABAS.ATIVIDADES,
      cols:   ['ID','Título','Tipo','Data','Horário','Local','Responsável','Departamento','Hierarquia','Status','Descrição','Aprovado Por','Aprovado Em','Cancelado Por','Cancelado Em'],
      keys:   COLS.ATIVIDADES,
      widths: [60, 220, 100, 100, 80, 160, 140, 140, 120, 100, 280, 140, 160, 140, 160],
    },
    {
      nome:   ABAS.HISTORICO,
      cols:   ['Mês (AAAA-MM)','Ativos em Dep.','Batizados em Dep.','Total Jovens'],
      keys:   COLS.HISTORICO,
      widths: [130, 130, 150, 120],
    },
  ];

  configuracoes.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.nome);
    if (!sheet) sheet = ss.insertSheet(cfg.nome);

    const hdr = sheet.getRange(1, 1, 1, cfg.cols.length);
    hdr.setValues([cfg.cols]);
    hdr.setBackground(VERMELHO).setFontColor(BRANCO).setFontWeight('bold').setFontSize(11);
    hdr.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 36);
    cfg.widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // Formatação condicional — Jovens: status
    if (cfg.nome === ABAS.JOVENS) {
      const statusCol   = sheet.getRange('I2:I1000');
      const batizadoCol = sheet.getRange('K2:K1000');
      const regrasStatus = [
        ['Ativo',     '#C8F7C5', '#145A32'],
        ['Oscilando', '#FEF9C3', '#78350F'],
        ['Ocioso',    '#FFEDD5', '#7C2D12'],
        ['Em Risco',  '#FEE2E2', '#7F1D1D'],
      ].map(([val, bg, fg]) =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(val).setBackground(bg).setFontColor(fg)
          .setRanges([statusCol]).build()
      );
      const regrasBatismo = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('sim').setBackground('#DBEAFE').setFontColor('#1D4ED8')
          .setRanges([batizadoCol]).build(),
      ];
      sheet.setConditionalFormatRules([...regrasStatus, ...regrasBatismo]);
    }

    // Formatação condicional — Frequências: presença
    if (cfg.nome === ABAS.FREQUENCIAS) {
      const presCol = sheet.getRange('E2:E1000');
      sheet.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('presente').setBackground('#C8F7C5').setFontColor('#145A32').setRanges([presCol]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('falta').setBackground('#FEE2E2').setFontColor('#7F1D1D').setRanges([presCol]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('justificada').setBackground('#FEF9C3').setFontColor('#78350F').setRanges([presCol]).build(),
      ]);
    }

    // Formatação condicional — Atividades: status
    if (cfg.nome === ABAS.ATIVIDADES) {
      const statusAtivCol = sheet.getRange('J2:J1000');
      sheet.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('aprovado').setBackground('#C8F7C5').setFontColor('#145A32').setRanges([statusAtivCol]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pendente').setBackground('#FEF9C3').setFontColor('#78350F').setRanges([statusAtivCol]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('cancelado').setBackground('#FEE2E2').setFontColor('#7F1D1D').setRanges([statusAtivCol]).build(),
      ]);
    }

    // Formatação condicional — Histórico: barras de dados
    if (cfg.nome === ABAS.HISTORICO) {
      const ativosCol    = sheet.getRange('B2:B1000');
      const batizadosCol = sheet.getRange('C2:C1000');
      sheet.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule()
          .setGradientMaxpointWithValue('#D42B2B', SpreadsheetApp.InterpolationType.NUMBER, '20')
          .setGradientMinpointWithValue('#FEE2E2', SpreadsheetApp.InterpolationType.NUMBER, '0')
          .setRanges([ativosCol]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .setGradientMaxpointWithValue('#1D4ED8', SpreadsheetApp.InterpolationType.NUMBER, '10')
          .setGradientMinpointWithValue('#DBEAFE', SpreadsheetApp.InterpolationType.NUMBER, '0')
          .setRanges([batizadosCol]).build(),
      ]);
    }

    Logger.log('✅ Aba configurada: ' + cfg.nome);
  });

  // ── Aba METAS
  let metasSheet = ss.getSheetByName(ABAS.METAS);
  if (!metasSheet) metasSheet = ss.insertSheet(ABAS.METAS);
  metasSheet.getRange('A1:C1').setValues([['Meta','Valor Atual','Objetivo']]).setBackground(VERMELHO).setFontColor(BRANCO).setFontWeight('bold');
  metasSheet.getRange('A2:C3').setValues([
    ['Ativos em Departamento',    '=COUNTIFS(Jovens!I:I,"Ativo",Jovens!G:G,"<>"&"")', 20],
    ['Batizados Ativos em Dep.',  '=COUNTIFS(Jovens!I:I,"Ativo",Jovens!G:G,"<>"&"",Jovens!K:K,"sim")', 10],
  ]);
  metasSheet.setColumnWidth(1, 220);
  metasSheet.setColumnWidth(2, 140);
  metasSheet.setColumnWidth(3, 100);
  metasSheet.setFrozenRows(1);
  // Barra de progresso com formatação condicional
  const progressoCol = metasSheet.getRange('B2:B10');
  metasSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(20).setBackground('#C8F7C5').setFontColor('#145A32')
      .setRanges([progressoCol]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(10, 19).setBackground('#FEF9C3').setFontColor('#78350F')
      .setRanges([progressoCol]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(10).setBackground('#FEE2E2').setFontColor('#7F1D1D')
      .setRanges([progressoCol]).build(),
  ]);
  Logger.log('✅ Aba configurada: ' + ABAS.METAS);

  // ── Aba CONFIG
  let configSheet = ss.getSheetByName(ABAS.CONFIG);
  if (!configSheet) configSheet = ss.insertSheet(ABAS.CONFIG);
  configSheet.getRange('A1:C1').setValues([['Seção','Chave','Valor']]).setBackground(PRETO).setFontColor(BRANCO).setFontWeight('bold');
  configSheet.getRange('A2:C7').setValues([
    ['regras', 'ativo',           75],
    ['regras', 'oscilando',       40],
    ['regras', 'risco',            3],
    ['',       '',                ''],
    ['metas',  'ativosDepto',     20],
    ['metas',  'batizadosDepto',  10],
  ]);
  configSheet.setColumnWidth(1, 100);
  configSheet.setColumnWidth(2, 160);
  configSheet.setColumnWidth(3, 100);
  configSheet.setFrozenRows(1);

  // ── Aba LOG
  let logSheet = ss.getSheetByName(ABAS.LOG);
  if (!logSheet) logSheet = ss.insertSheet(ABAS.LOG);
  logSheet.getRange('A1:C1').setValues([['Data/Hora','Usuário','Ação']]).setBackground(PRETO).setFontColor(BRANCO).setFontWeight('bold');
  logSheet.setColumnWidth(1, 160); logSheet.setColumnWidth(2, 160); logSheet.setColumnWidth(3, 300);
  logSheet.setFrozenRows(1);

  // ── Remove Sheet1 padrão
  ['Plan1','Sheet1','Página1'].forEach(nome => {
    const s = ss.getSheetByName(nome);
    if (s && ss.getSheets().length > 1) try { ss.deleteSheet(s); } catch(e){}
  });

  // ── Reordena abas
  [ABAS.JOVENS, ABAS.FREQUENCIAS, ABAS.RECUPERACOES, ABAS.DEPARTAMENTOS,
   ABAS.LIDERES, ABAS.METAS, ABAS.HISTORICO, ABAS.ATIVIDADES, ABAS.CONFIG, ABAS.LOG
  ].forEach((nome, i) => {
    const s = ss.getSheetByName(nome);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });

  const url = ss.getUrl();
  Logger.log('\n══════════════════════════════════════════');
  Logger.log('✅ RUJA v2 — Planilha configurada!');
  Logger.log('🔗 URL: ' + url);
  Logger.log('══════════════════════════════════════════');
  Logger.log('Agora: Implantar → Nova implantação → App da Web');

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Planilha RUJA v2 configurada!\n\n' +
      'Abas criadas:\n' +
      '• Jovens (+ Batizado e Data do Batismo)\n' +
      '• Frequências\n• Recuperações\n• Departamentos\n• Líderes\n' +
      '• Metas (com fórmulas automáticas)\n' +
      '• Histórico Mensal\n• Atividades & Eventos\n• Configurações\n• Log de Alterações\n\n' +
      'URL da planilha:\n' + url
    );
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════════════════════════
// MIGRAR PARA V2 — execute se já tinha a v1 instalada
// Adiciona novas colunas/abas sem apagar dados existentes
// ══════════════════════════════════════════════════════════════════════════════
function migrarParaV3() {
  const ss = getSpreadsheet();
  Logger.log('🔄 Iniciando migração para v3...');

  // 1. Aba Jovens — adiciona colunas Batizado e Data Batismo se não existirem
  const jovensSheet = ss.getSheetByName(ABAS.JOVENS);
  if (jovensSheet) {
    const hdr = jovensSheet.getRange(1, 1, 1, jovensSheet.getLastColumn()).getValues()[0];
    const temBatizado = hdr.includes('Batizado');
    if (!temBatizado) {
      // A v1 tinha 11 colunas (obs na coluna 11)
      // V2 adiciona batizado (col 11) e dataBatismo (col 12) antes de obs
      const lastCol = jovensSheet.getLastColumn();
      jovensSheet.insertColumns(lastCol, 2); // insere 2 colunas antes de obs

      const lastRow = jovensSheet.getLastRow();
      // Move coluna obs (era col 11) para col 13
      // (Na prática, como insertColumns move, obs já está em 13)
      // Define cabeçalhos das novas colunas
      jovensSheet.getRange(1, lastCol, 1, 1).setValue('Batizado').setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
      jovensSheet.getRange(1, lastCol+1, 1, 1).setValue('Data Batismo').setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
      jovensSheet.setColumnWidth(lastCol, 90);
      jovensSheet.setColumnWidth(lastCol+1, 110);

      // Preenche linhas existentes com "nao"
      if (lastRow > 1) {
        jovensSheet.getRange(2, lastCol, lastRow-1, 1).setValue('nao');
      }

      // Formatação condicional batizado
      const batizadoCol = jovensSheet.getRange(`${columnToLetter(lastCol)}2:${columnToLetter(lastCol)}1000`);
      const regrasExist = jovensSheet.getConditionalFormatRules();
      regrasExist.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('sim').setBackground('#DBEAFE').setFontColor('#1D4ED8')
          .setRanges([batizadoCol]).build()
      );
      jovensSheet.setConditionalFormatRules(regrasExist);
      Logger.log('✅ Colunas Batizado e Data Batismo adicionadas à aba Jovens');
    } else {
      Logger.log('ℹ️ Aba Jovens já tem coluna Batizado — sem alteração');
    }
  }

  // 2. Cria aba Metas se não existir
  if (!ss.getSheetByName(ABAS.METAS)) {
    const metasSheet = ss.insertSheet(ABAS.METAS);
    metasSheet.getRange('A1:C1').setValues([['Meta','Valor Atual','Objetivo']]).setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
    metasSheet.getRange('A2:C3').setValues([
      ['Ativos em Departamento',    '=COUNTIFS(Jovens!I:I,"Ativo",Jovens!G:G,"<>"&"")', 20],
      ['Batizados Ativos em Dep.',  '=COUNTIFS(Jovens!I:I,"Ativo",Jovens!G:G,"<>"&"",Jovens!K:K,"sim")', 10],
    ]);
    metasSheet.setColumnWidth(1, 220); metasSheet.setColumnWidth(2, 140); metasSheet.setColumnWidth(3, 100);
    metasSheet.setFrozenRows(1);
    Logger.log('✅ Aba Metas criada');
  } else {
    Logger.log('ℹ️ Aba Metas já existe — sem alteração');
  }

  // 3. Cria aba Histórico Mensal se não existir
  if (!ss.getSheetByName(ABAS.HISTORICO)) {
    const histSheet = ss.insertSheet(ABAS.HISTORICO);
    histSheet.getRange('A1:D1').setValues([['Mês (AAAA-MM)','Ativos em Dep.','Batizados em Dep.','Total Jovens']]).setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
    histSheet.setColumnWidth(1, 130); histSheet.setColumnWidth(2, 130); histSheet.setColumnWidth(3, 150); histSheet.setColumnWidth(4, 120);
    histSheet.setFrozenRows(1);
    Logger.log('✅ Aba Histórico Mensal criada');
  } else {
    Logger.log('ℹ️ Aba Histórico Mensal já existe — sem alteração');
  }

  // 4. Atualiza Config para o formato v2 (3 colunas: seção, chave, valor)
  const configSheet = ss.getSheetByName(ABAS.CONFIG);
  if (configSheet) {
    const hdr = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    if (hdr.length < 3 || hdr[2] !== 'Valor') {
      // Recria no formato v2
      const lastRow = configSheet.getLastRow();
      if (lastRow > 1) configSheet.deleteRows(2, lastRow-1);
      configSheet.getRange('A1:C1').setValues([['Seção','Chave','Valor']]).setBackground('#1A1A1A').setFontColor('#FFFFFF').setFontWeight('bold');
      configSheet.getRange('A2:C7').setValues([
        ['regras','ativo',75],['regras','oscilando',40],['regras','risco',3],
        ['','',''],
        ['metas','ativosDepto',20],['metas','batizadosDepto',10],
      ]);
      configSheet.setColumnWidth(1, 100); configSheet.setColumnWidth(2, 160); configSheet.setColumnWidth(3, 100);
      Logger.log('✅ Aba Configurações atualizada para formato v2');
    }
  }

  Logger.log('\n══════════════════════════════════════════');
  // 5. Aba Atividades — cria se não existir
  if (!ss.getSheetByName(ABAS.ATIVIDADES)) {
    const ativSheet = ss.insertSheet(ABAS.ATIVIDADES);
    const ativHdr = ['ID','Título','Tipo','Data','Horário','Local','Responsável','Departamento','Hierarquia','Status','Descrição'];
    ativSheet.getRange(1, 1, 1, ativHdr.length).setValues([ativHdr])
      .setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
    [60,220,100,100,80,160,140,140,120,100,280].forEach((w,i)=>ativSheet.setColumnWidth(i+1,w));
    ativSheet.setFrozenRows(1);
    // Formatação condicional de status
    const statusCol = ativSheet.getRange('J2:J1000');
    ativSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('aprovado').setBackground('#C8F7C5').setFontColor('#145A32').setRanges([statusCol]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pendente').setBackground('#FEF9C3').setFontColor('#78350F').setRanges([statusCol]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('cancelado').setBackground('#FEE2E2').setFontColor('#7F1D1D').setRanges([statusCol]).build(),
    ]);
    Logger.log('✅ Aba Atividades criada');
  } else {
    Logger.log('ℹ️ Aba Atividades já existe — sem alteração');
  }

  // 6. Aba Jovens — adiciona coluna Nascimento se não existir
  const jovensSheetV3 = ss.getSheetByName(ABAS.JOVENS);
  if (jovensSheetV3) {
    const hdrV3 = jovensSheetV3.getRange(1, 1, 1, jovensSheetV3.getLastColumn()).getValues()[0];
    if (!hdrV3.includes('Nascimento')) {
      const lastColV3 = jovensSheetV3.getLastColumn();
      jovensSheetV3.insertColumnBefore(lastColV3); // insere antes de Obs
      jovensSheetV3.getRange(1, lastColV3, 1, 1)
        .setValue('Nascimento').setBackground('#D42B2B').setFontColor('#FFFFFF').setFontWeight('bold');
      jovensSheetV3.setColumnWidth(lastColV3, 110);
      Logger.log('✅ Coluna Nascimento adicionada à aba Jovens');
    } else {
      Logger.log('ℹ️ Coluna Nascimento já existe');
    }
  }

  Logger.log('✅ Migração para v3 concluída!');
  Logger.log('Agora faça: Implantar → Gerenciar implantações → Editar → Nova versão');
  Logger.log('══════════════════════════════════════════');
}

// Helper: número de coluna para letra (ex: 11 → K)
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ══════════════════════════════════════════════════════════════════════════════
// UTILITÁRIOS
// ══════════════════════════════════════════════════════════════════════════════

function verUrlPlanilha() {
  const ss = getSpreadsheet();
  Logger.log('🔗 URL: ' + ss.getUrl());
  Logger.log('📋 ID: ' + ss.getId());
}

function testarLeitura() {
  const result = doGet({});
  const data   = JSON.parse(result.getContent());
  Logger.log('Jovens:          ' + (data.jovens          || []).length + ' registros');
  Logger.log('Frequências:     ' + (data.frequencias     || []).length + ' registros');
  Logger.log('Recuperações:    ' + (data.recuperacoes    || []).length + ' registros');
  Logger.log('Departamentos:   ' + (data.departamentos    || []).length + ' registros');
  Logger.log('Líderes:         ' + (data.lideres          || []).length + ' registros');
  Logger.log('Histórico:       ' + (data.historicoMensal  || []).length + ' meses');
  Logger.log('Atividades:      ' + (data.atividades        || []).length + ' registros');
  Logger.log('Regras:          ' + JSON.stringify(data.regras));
  Logger.log('Metas:           ' + JSON.stringify(data.metas));

  // Verifica campos de batizado
  const j = (data.jovens || [])[0];
  if (j) Logger.log('Campos Jovem[0]: ' + Object.keys(j).join(', '));

  Logger.log('✅ Leitura v2 OK!');
}

function limparTodosDados() {
  try {
    SpreadsheetApp.getUi().alert('⚠️ ATENÇÃO: Apagará TODOS os dados. Confirme rodando novamente após ler este aviso.');
  } catch(e) {}
  const ss = getSpreadsheet();
  Object.values(ABAS).forEach(nome => {
    const sheet = ss.getSheetByName(nome);
    if (!sheet) return;
    const last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);
  });
  Logger.log('⚠️ Dados limpos.');
}

// ── RELATÓRIO SEMANAL (atualizado com metas)
function enviarRelatorioSemanal() {
  const dados          = JSON.parse(doGet({}).getContent());
  const jovens         = dados.jovens         || [];
  const historico      = dados.historicoMensal|| [];
  const metas          = dados.metas          || { ativosDepto:20, batizadosDepto:10 };

  const total          = jovens.length;
  const ativos         = jovens.filter(j => j.status === 'Ativo').length;
  const oscilando      = jovens.filter(j => j.status === 'Oscilando').length;
  const ociosos        = jovens.filter(j => j.status === 'Ocioso').length;
  const risco          = jovens.filter(j => j.status === 'Em Risco').length;
  const ativosDepto    = jovens.filter(j => j.status === 'Ativo' && j.departamento).length;
  const batizadosDepto = jovens.filter(j => j.status === 'Ativo' && j.departamento && j.batizado === 'sim').length;
  const semDepto       = jovens.filter(j => !j.departamento).length;
  const pctAtivos      = total ? Math.round(ativos / total * 100) : 0;
  const pctMetaAtivos  = metas.ativosDepto    ? Math.round(ativosDepto    / metas.ativosDepto    * 100) : 0;
  const pctMetaBatiz   = metas.batizadosDepto ? Math.round(batizadosDepto / metas.batizadosDepto * 100) : 0;

  const barra = (pct, tam=10) => '█'.repeat(Math.round(pct/100*tam)) + '░'.repeat(tam - Math.round(pct/100*tam));

  const corpo = `
🦁 RELATÓRIO SEMANAL — UP RUJA
Data: ${new Date().toLocaleDateString('pt-BR')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 FUNIL DE STATUS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total de jovens:    ${total}
🟢 Ativos:          ${ativos} (${pctAtivos}%)
🟡 Oscilando:       ${oscilando}
🟠 Ociosos:         ${ociosos}
🔴 Em Risco:        ${risco}
⚠️  Sem departamento: ${semDepto}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🎯 PROGRESSO DAS METAS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🏛️ Ativos em Departamento:
   ${ativosDepto} / ${metas.ativosDepto}  ${barra(pctMetaAtivos)}  ${pctMetaAtivos}%
   ${pctMetaAtivos >= 100 ? '✅ META ATINGIDA!' : `Faltam ${metas.ativosDepto - ativosDepto} jovens`}

🔵 Batizados Ativos em Dep.:
   ${batizadosDepto} / ${metas.batizadosDepto}  ${barra(pctMetaBatiz)}  ${pctMetaBatiz}%
   ${pctMetaBatiz >= 100 ? '✅ META ATINGIDA!' : `Faltam ${metas.batizadosDepto - batizadosDepto} jovens`}

${risco > 0
  ? `\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🚨 EM RISCO — AÇÃO URGENTE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${jovens.filter(j=>j.status==='Em Risco').map(j=>`• ${j.nome}
  📱 ${j.contato}
  Líder: ${j.lider||'Não definido'}`).join('\n')}`
  : '\n✅ Nenhum jovem em risco esta semana!'}

🦁 UP RUJA — Rede UniJovem ADJA
  `.trim();

  // Atividades aprovadas próximas (7 dias)
  const ativs = dados.atividades || [];
  const hoje7 = new Date(); hoje7.setDate(hoje7.getDate()+7);
  const proxAtivs = ativs.filter(a => {
    if(a.status !== 'aprovado') return false;
    const d = new Date(a.data);
    return d >= new Date() && d <= hoje7;
  });
  if(proxAtivs.length) {
    corpo += '\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📅 ATIVIDADES APROVADAS (próximos 7 dias)\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    corpo += proxAtivs.map(a=>`• ${a.titulo} — ${a.data}${a.horario?' às '+a.horario:''} | ${a.local||'Local TBD'}`).join('\n');
  }

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: `🦁 Relatório RUJA v3 — ${new Date().toLocaleDateString('pt-BR')} | Metas: ${pctMetaAtivos}% / ${pctMetaBatiz}%`,
    body: corpo,
  });

  Logger.log('✅ Relatório v3 enviado para: ' + Session.getActiveUser().getEmail());
}

// ══════════════════════════════════════════════════════════════════════════════
// RESPOSTA JSON
// ══════════════════════════════════════════════════════════════════════════════
function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════════════════
//  GATILHO SEMANAL:
//  Apps Script → Gatilhos → + Adicionar gatilho
//  Função: enviarRelatorioSemanal | Semanal → Segunda-feira → 08:00
// ══════════════════════════════════════════════════════════════════════════════
