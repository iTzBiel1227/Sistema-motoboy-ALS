Código.gs

function doGet() {

  /*
   * Tenta garantir o gatilho automático de status.
   *
   * Caso a conta ainda não tenha autorização para criar
   * o gatilho automaticamente, o restante do sistema
   * continua funcionando normalmente.
   *
   * Também existe a função configurarAutomacaoSistema()
   * para execução manual uma única vez.
   */

  try {
    garantirGatilhoStatus();
  }
  catch (erro) {
    // Não impede o carregamento do sistema.
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('ALS Transportes | Sistema de Corridas');

}


/* ============================================================
   CONFIGURAÇÃO AUTOMÁTICA DO SISTEMA
   ============================================================ */

function configurarAutomacaoSistema() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Respostas ao formulário 1'
    );

  if (!aba) {
    throw new Error(
      'A aba "Respostas ao formulário 1" não foi encontrada.'
    );
  }

  garantirEstruturaCorridas(
    aba
  );

  ajustarCadastrosLocais();

  atualizarStatusAutomatico();

  garantirGatilhoStatus();

  return {
    sucesso: true,
    mensagem:
      'Automação configurada com sucesso.'
  };

}


/* ============================================================
   GARANTIR GATILHO AUTOMÁTICO DE STATUS
   ============================================================ */

function garantirGatilhoStatus() {

  const triggers =
    ScriptApp.getProjectTriggers();

  const existe =
    triggers.some(
      function(trigger) {

        return (
          trigger.getHandlerFunction() ===
          'atualizarStatusAutomatico'
        );

      }
    );

  if (existe) {
    return;
  }

  ScriptApp
    .newTrigger(
      'atualizarStatusAutomatico'
    )
    .timeBased()
    .everyMinutes(5)
    .create();

}


/* ============================================================
   GARANTIR ESTRUTURA DA ABA DE CORRIDAS
   ============================================================ */

function garantirEstruturaCorridas(
  aba
) {

  if (
    aba.getMaxColumns() < 13
  ) {

    aba.insertColumnsAfter(
      aba.getMaxColumns(),
      13 - aba.getMaxColumns()
    );

  }

  const cabecalhos = [
    'STATUS',
    'ALTERADO EM',
    'HISTÓRICO'
  ];

  for (
    let i = 0;
    i < cabecalhos.length;
    i++
  ) {

    const coluna =
      11 + i;

    const atual =
      aba
        .getRange(
          1,
          coluna
        )
        .getValue();

    if (!atual) {

      aba
        .getRange(
          1,
          coluna
        )
        .setValue(
          cabecalhos[i]
        );

    }

  }

}


/* ============================================================
   REGISTRAR CORRIDA
   ============================================================ */

function registrarCorrida(
  dados
) {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Respostas ao formulário 1'
    );

  if (!aba) {

    throw new Error(
      'A aba "Respostas ao formulário 1" não foi encontrada.'
    );

  }

  if (
    !dados.solicitante ||
    !dados.motoboy ||
    !dados.corrida
  ) {

    throw new Error(
      'Preencha todos os campos obrigatórios.'
    );

  }

  if (
    dados.valor === undefined ||
    dados.valor === null ||
    Number(dados.valor) < 0
  ) {

    throw new Error(
      'Informe um valor válido para a corrida.'
    );

  }

  garantirEstruturaCorridas(
    aba
  );

  const numeroCorrida =
    gerarNumeroCorrida(
      aba
    );

  const data =
    new Date(
      dados.data +
      'T' +
      dados.horario
    );

  const carimbo =
    new Date();

  const dataEmissao =
    new Date();

  const status =
    'Solicitada';

  const historico =
    formatarDataHoraServidor(
      carimbo
    ) +
    ' — Corrida solicitada';

  aba.appendRow([
    carimbo,
    dados.motoboy,
    dados.corrida,
    dados.solicitante,
    Number(dados.valor),
    data,
    dados.horario,
    dados.observacao || '',
    numeroCorrida,
    dataEmissao,
    status,
    carimbo,
    historico
  ]);

  SpreadsheetApp.flush();

  const ultimaLinha =
    aba.getLastRow();

  aba
    .getRange(
      ultimaLinha,
      9
    )
    .setNumberFormat('@');

  aba
    .getRange(
      ultimaLinha,
      10
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );

  aba
    .getRange(
      ultimaLinha,
      12
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );

  SpreadsheetApp.flush();

  return {

    sucesso: true,

    numeroCorrida:
      numeroCorrida,

    status:
      status

  };

}


/* ============================================================
   GERAR NÚMERO DA CORRIDA
   ============================================================ */

function gerarNumeroCorrida(
  aba
) {

  const hoje =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd'
    );

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {

    return (
      'COR-' +
      hoje +
      '-0001'
    );

  }

  const quantidadeLinhas =
    ultimaLinha - 1;

  const valores =
    aba
      .getRange(
        2,
        9,
        quantidadeLinhas,
        1
      )
      .getValues()
      .flat();

  let maiorNumero =
    0;

  valores.forEach(
    function(valor) {

      if (!valor) {
        return;
      }

      const texto =
        String(
          valor
        ).trim();

      const partes =
        texto.split('-');

      if (partes.length !== 3) {
        return;
      }

      const numero =
        parseInt(
          partes[2],
          10
        );

      if (
        !isNaN(numero) &&
        numero > maiorNumero
      ) {

        maiorNumero =
          numero;

      }

    }
  );

  return (
    'COR-' +
    hoje +
    '-' +
    String(
      maiorNumero + 1
    ).padStart(
      4,
      '0'
    )
  );

}


/* ============================================================
   ATUALIZAR STATUS AUTOMATICAMENTE
   ============================================================ */

function atualizarStatusAutomatico() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Respostas ao formulário 1'
    );

  if (!aba) {
    return;
  }

  garantirEstruturaCorridas(
    aba
  );

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const quantidadeLinhas =
    ultimaLinha - 1;

  const dados =
    aba
      .getRange(
        2,
        1,
        quantidadeLinhas,
        13
      )
      .getValues();

  const extras =
    aba
      .getRange(
        2,
        11,
        quantidadeLinhas,
        3
      )
      .getValues();

  const agora =
    new Date();

  const duasHoras =
    2 * 60 * 60 * 1000;

  let houveAlteracao =
    false;

  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const possuiDados =
      dados[i].some(
        function(valor) {

          return (
            valor !== '' &&
            valor !== null
          );

        }
      );

    if (!possuiDados) {
      continue;
    }

    const carimbo =
      dados[i][0];

    const statusAtual =
      String(
        extras[i][0] || ''
      ).trim();

    const historicoAtual =
      String(
        extras[i][2] || ''
      );

    if (
      !(carimbo instanceof Date) ||
      isNaN(
        carimbo.getTime()
      )
    ) {

      if (!statusAtual) {

        extras[i][0] =
          'Solicitada';

        extras[i][1] =
          agora;

        extras[i][2] =
          'Corrida solicitada';

        houveAlteracao =
          true;

      }

      continue;

    }


    /*
     * Corrida antiga sem status.
     */

    if (!statusAtual) {

      const passouDuasHoras =
        (
          agora.getTime() -
          carimbo.getTime()
        ) >= duasHoras;

      if (passouDuasHoras) {

        extras[i][0] =
          'Concluída';

        extras[i][1] =
          agora;

        extras[i][2] =
          formatarDataHoraServidor(
            carimbo
          ) +
          ' — Corrida solicitada\n' +
          formatarDataHoraServidor(
            agora
          ) +
          ' — Corrida concluída automaticamente após 2 horas';

      }

      else {

        extras[i][0] =
          'Solicitada';

        extras[i][1] =
          carimbo;

        extras[i][2] =
          formatarDataHoraServidor(
            carimbo
          ) +
          ' — Corrida solicitada';

      }

      houveAlteracao =
        true;

      continue;

    }


    /*
     * Corrida solicitada que atingiu 2 horas.
     */

    if (
      statusAtual === 'Solicitada' &&
      (
        agora.getTime() -
        carimbo.getTime()
      ) >= duasHoras
    ) {

      extras[i][0] =
        'Concluída';

      extras[i][1] =
        agora;

      extras[i][2] =
        adicionarEventoHistorico(
          historicoAtual,
          formatarDataHoraServidor(
            agora
          ) +
          ' — Corrida concluída automaticamente após 2 horas'
        );

      houveAlteracao =
        true;

    }

  }


  if (houveAlteracao) {

    aba
      .getRange(
        2,
        11,
        quantidadeLinhas,
        3
      )
      .setValues(
        extras
      );

    SpreadsheetApp.flush();

  }

}


/* ============================================================
   ALTERAR STATUS MANUALMENTE
   ============================================================ */

function alterarStatusCorrida(
  linhaPlanilha,
  novoStatus
) {

  if (
    !linhaPlanilha ||
    Number(
      linhaPlanilha
    ) < 2
  ) {

    throw new Error(
      'Linha da corrida inválida.'
    );

  }

  if (
    novoStatus !== 'Cancelada'
  ) {

    throw new Error(
      'Status manual permitido: Cancelada.'
    );

  }

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Respostas ao formulário 1'
    );

  if (!aba) {

    throw new Error(
      'A aba "Respostas ao formulário 1" não foi encontrada.'
    );

  }

  garantirEstruturaCorridas(
    aba
  );

  const numeroLinha =
    Number(
      linhaPlanilha
    );

  const numeroCorrida =
    aba
      .getRange(
        numeroLinha,
        9
      )
      .getDisplayValue();

  if (!numeroCorrida) {

    throw new Error(
      'A corrida não foi encontrada.'
    );

  }

  const dadosStatus =
    aba
      .getRange(
        numeroLinha,
        11,
        1,
        3
      )
      .getValues()[0];

  const statusAtual =
    String(
      dadosStatus[0] || ''
    ).trim();

  const historicoAtual =
    String(
      dadosStatus[2] || ''
    );

  if (
    statusAtual === 'Cancelada'
  ) {

    return {

      sucesso: true,

      status:
        'Cancelada',

      numeroCorrida:
        numeroCorrida,

      alteradoEm:
        formatarDataHoraServidor(
          dadosStatus[1]
        ),

      historico:
        historicoAtual

    };

  }

  const agora =
    new Date();

  const novoHistorico =
    adicionarEventoHistorico(

      historicoAtual,

      formatarDataHoraServidor(
        agora
      ) +
      ' — Corrida cancelada manualmente'

    );

  aba
    .getRange(
      numeroLinha,
      11,
      1,
      3
    )
    .setValues([[
      'Cancelada',
      agora,
      novoHistorico
    ]]);

  SpreadsheetApp.flush();

  return {

    sucesso: true,

    status:
      'Cancelada',

    numeroCorrida:
      numeroCorrida,

    alteradoEm:
      formatarDataHoraServidor(
        agora
      ),

    historico:
      novoHistorico

  };

}


/* ============================================================
   ADICIONAR EVENTO AO HISTÓRICO
   ============================================================ */

function adicionarEventoHistorico(
  historico,
  evento
) {

  const textoAtual =
    String(
      historico || ''
    ).trim();

  if (!textoAtual) {
    return evento;
  }

  return (
    textoAtual +
    '\n' +
    evento
  );

}


/* ============================================================
   NORMALIZAR TEXTO
   ============================================================ */

function normalizarTextoCadastro(
  texto
) {

  return String(
    texto || ''
  )
  .trim()
  .toLowerCase()
  .normalize('NFD')
  .replace(
    /[\u0300-\u036f]/g,
    ''
  );

}


/* ============================================================
   CHAVE DA COMBINAÇÃO
   ============================================================ */

function gerarChaveCombinacao(
  combinacao
) {

  return String(
    combinacao || ''
  )
  .split('+')
  .map(
    function(parte) {

      return normalizarTextoCadastro(
        parte
      );

    }
  )
  .filter(
    function(parte) {

      return parte !== '';

    }
  )
  .sort()
  .join('|');

}


/* ============================================================
   GARANTIR COMBINAÇÃO ESPECIAL
   ============================================================ */

function garantirCombinacaoEspecial(
  aba,
  nomeCombinacao,
  valorRetorno,
  valorFixo
) {

  const chaveDesejada =
    gerarChaveCombinacao(
      nomeCombinacao
    );

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha >= 2) {

    const dados =
      aba
        .getRange(
          2,
          9,
          ultimaLinha - 1,
          3
        )
        .getValues();

    for (
      let i = 0;
      i < dados.length;
      i++
    ) {

      const combinacaoAtual =
        String(
          dados[i][0] || ''
        ).trim();

      if (!combinacaoAtual) {
        continue;
      }

      const chaveAtual =
        gerarChaveCombinacao(
          combinacaoAtual
        );

      if (
        chaveAtual ===
        chaveDesejada
      ) {

        const linhaPlanilha =
          i + 2;

        aba
          .getRange(
            linhaPlanilha,
            10
          )
          .setValue(
            valorRetorno
          );

        aba
          .getRange(
            linhaPlanilha,
            11
          )
          .setValue(
            valorFixo
          );

        return;

      }

    }

  }

  const novaLinha =
    aba.getLastRow() + 1;

  aba
    .getRange(
      novaLinha,
      9
    )
    .setValue(
      nomeCombinacao
    );

  aba
    .getRange(
      novaLinha,
      10
    )
    .setValue(
      valorRetorno
    );

  aba
    .getRange(
      novaLinha,
      11
    )
    .setValue(
      valorFixo
    );

}


/* ============================================================
   AJUSTAR CADASTROS
   ============================================================ */

function ajustarCadastrosLocais() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Cadastros'
    );

  if (!aba) {

    throw new Error(
      'A aba "Cadastros" não foi encontrada.'
    );

  }

  if (aba.getMaxColumns() < 11) {

    aba.insertColumnsAfter(
      aba.getMaxColumns(),
      11 - aba.getMaxColumns()
    );

  }

  aba
    .getRange('I1')
    .setValue(
      'COMBINAÇÃO ESPECIAL'
    );

  aba
    .getRange('J1')
    .setValue(
      'VALOR DO RETORNO'
    );

  aba
    .getRange('K1')
    .setValue(
      'VALOR FIXO'
    );


  /* ========================================================
     MARINEX -> MARIMEX
     ======================================================== */

  const ultimaLinhaInicial =
    aba.getLastRow();

  if (
    ultimaLinhaInicial >= 2
  ) {

    const dados =
      aba
        .getRange(
          2,
          3,
          ultimaLinhaInicial - 1,
          2
        )
        .getValues();

    dados.forEach(
      function(linha, indice) {

        const nome =
          String(
            linha[0] || ''
          )
          .trim()
          .toLowerCase();

        if (
          nome === 'marinex'
        ) {

          aba
            .getRange(
              indice + 2,
              3
            )
            .setValue(
              'Marimex'
            );

        }

      }
    );

  }


  /* ========================================================
     CORRIGIR MARINEX NAS COMBINAÇÕES
     ======================================================== */

  const ultimaLinhaEspecial =
    aba.getLastRow();

  if (
    ultimaLinhaEspecial >= 2
  ) {

    const especiais =
      aba
        .getRange(
          2,
          9,
          ultimaLinhaEspecial - 1,
          3
        )
        .getValues();

    especiais.forEach(
      function(linha, indice) {

        const combinacao =
          String(
            linha[0] || ''
          );

        if (
          combinacao
            .toLowerCase()
            .includes('marinex')
        ) {

          aba
            .getRange(
              indice + 2,
              9
            )
            .setValue(
              combinacao.replace(
                /marinex/gi,
                'Marimex'
              )
            );

        }

      }
    );

  }


  /* ========================================================
     GARANTIR BARRACA DA JÔ
     ======================================================== */

  const ultimaLinhaAtual =
    aba.getLastRow();

  let existeBarraca =
    false;

  if (
    ultimaLinhaAtual >= 2
  ) {

    const locais =
      aba
        .getRange(
          2,
          3,
          ultimaLinhaAtual - 1,
          2
        )
        .getValues();

    locais.forEach(
      function(linha, indice) {

        const nomeOriginal =
          String(
            linha[0] || ''
          ).trim();

        const nome =
          nomeOriginal.toLowerCase();

        if (
          nome === 'barraca da jô' ||
          nome === 'barraca da jo'
        ) {

          existeBarraca =
            true;

          if (
            nomeOriginal !==
            'Barraca da Jô'
          ) {

            aba
              .getRange(
                indice + 2,
                3
              )
              .setValue(
                'Barraca da Jô'
              );

          }

        }

      }
    );

  }

  if (!existeBarraca) {

    const novaLinha =
      aba.getLastRow() + 1;

    aba
      .getRange(
        novaLinha,
        3
      )
      .setValue(
        'Barraca da Jô'
      );

    aba
      .getRange(
        novaLinha,
        4
      )
      .setValue(
        80
      );

  }


  /* ========================================================
     COMBINAÇÃO ESPECIAL 1
     POSTO VIADUTO + MARIMEX
     ======================================================== */

  garantirCombinacaoEspecial(
    aba,
    'Posto Viaduto + Marimex',
    0,
    150
  );


  /* ========================================================
     COMBINAÇÃO ESPECIAL 2
     POSTO VIADUTO + CARGA PESADA + BARRACA DA JÔ
     ======================================================== */

  garantirCombinacaoEspecial(
    aba,
    'Posto Viaduto + Carga pesada + Barraca da Jô',
    0,
    170
  );

}


/* ============================================================
   OBTER CADASTROS
   ============================================================ */

function obterCadastros() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Cadastros'
    );

  if (!aba) {

    throw new Error(
      'A aba "Cadastros" não foi encontrada.'
    );

  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {

    return {

      motoboys: [],

      solicitantes: []

    };

  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        2
      )
      .getValues();

  const motoboys = [];
  const solicitantes = [];

  dados.forEach(
    function(linha) {

      if (linha[0]) {

        motoboys.push(
          String(
            linha[0]
          )
        );

      }

      if (linha[1]) {

        solicitantes.push(
          String(
            linha[1]
          )
        );

      }

    }
  );

  return {

    motoboys:
      motoboys,

    solicitantes:
      solicitantes

  };

}


/* ============================================================
   OBTER PREÇOS
   ============================================================ */

function obterPrecos() {

  ajustarCadastrosLocais();

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Cadastros'
    );

  if (!aba) {

    throw new Error(
      'A aba "Cadastros" não foi encontrada.'
    );

  }

  const ultimaLinha =
    aba.getLastRow();

  const quantidadeLinhas =
    Math.max(
      ultimaLinha - 1,
      1
    );

  const dadosCorridas =
    aba
      .getRange(
        2,
        3,
        quantidadeLinhas,
        2
      )
      .getValues();

  const corridas = {};

  dadosCorridas.forEach(
    function(linha) {

      const nome =
        linha[0];

      const valor =
        linha[1];

      if (nome) {

        corridas[
          String(nome)
        ] =
          Number(
            valor
          ) || 0;

      }

    }
  );

  const valorRetorno =
    Number(
      aba
        .getRange('G2')
        .getValue()
    ) || 0;


  const especiais =
    aba
      .getRange(
        2,
        9,
        quantidadeLinhas,
        3
      )
      .getValues();

  const combinacoesEspeciais = {};
  const valoresFixosEspeciais = {};

  especiais.forEach(
    function(linha) {

      const combinacao =
        linha[0];

      const valorRetorno =
        linha[1];

      const valorFixo =
        linha[2];

      if (combinacao) {

        const chave =
          String(
            combinacao
          );

        combinacoesEspeciais[
          chave
        ] =
          Number(
            valorRetorno
          ) || 0;

        if (
          valorFixo !== '' &&
          valorFixo !== null &&
          !isNaN(
            Number(
              valorFixo
            )
          )
        ) {

          valoresFixosEspeciais[
            chave
          ] =
            Number(
              valorFixo
            );

        }

      }

    }
  );

  return {

    corridas:
      corridas,

    valorRetorno:
      valorRetorno,

    combinacoesEspeciais:
      combinacoesEspeciais,

    valoresFixosEspeciais:
      valoresFixosEspeciais

  };

}


/* ============================================================
   OBTER CORRIDAS
   ============================================================ */

function obterCorridas() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    planilha.getSheetByName(
      'Respostas ao formulário 1'
    );

  if (!aba) {

    throw new Error(
      'A aba "Respostas ao formulário 1" não foi encontrada.'
    );

  }

  garantirEstruturaCorridas(
    aba
  );

  /*
   * Garante que os status antigos
   * estejam atualizados antes da leitura.
   */

  atualizarStatusAutomatico();

  SpreadsheetApp.flush();

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        13
      )
      .getValues();

  const corridas = [];

  dados.forEach(
    function(linha, indice) {

      const possuiDados =
        linha.some(
          function(valor) {

            return (
              valor !== '' &&
              valor !== null
            );

          }
        );

      if (!possuiDados) {
        return;
      }

      const carimbo =
        linha[0];

      const motoboy =
        linha[1] || '';

      const corrida =
        linha[2] || '';

      const solicitante =
        linha[3] || '';

      const valor =
        Number(
          linha[4]
        ) || 0;

      const data =
        linha[5];

      const horario =
        linha[6] || '';

      const observacao =
        linha[7] || '';

      const numeroCorrida =
        linha[8] || '';

      const emissao =
        linha[9];

      const status =
        String(
          linha[10] ||
          'Solicitada'
        ).trim();

      const statusAlteradoEm =
        linha[11];

      const historico =
        String(
          linha[12] || ''
        );


      if (
        !corrida &&
        !motoboy &&
        !solicitante
      ) {

        return;

      }


      let conclusaoPrevista =
        '';

      if (
        carimbo instanceof Date &&
        !isNaN(
          carimbo.getTime()
        )
      ) {

        conclusaoPrevista =
          formatarDataHoraServidor(
            new Date(
              carimbo.getTime() +
              (2 * 60 * 60 * 1000)
            )
          );

      }


      corridas.push({

        linhaPlanilha:
          indice + 2,

        carimbo:
          formatarDataHoraServidor(
            carimbo
          ),

        carimboISO:
          formatarDataHoraISOServidor(
            carimbo
          ),

        motoboy:
          String(
            motoboy
          ),

        corrida:
          String(
            corrida
          ),

        solicitante:
          String(
            solicitante
          ),

        valor:
          valor,

        data:
          formatarDataServidor(
            data
          ),

        dataISO:
          formatarDataISOServidor(
            data
          ),

        horario:
          formatarHorarioServidor(
            horario
          ),

        observacao:
          String(
            observacao
          ),

        numeroCorrida:
          String(
            numeroCorrida
          ),

        emissao:
          formatarDataHoraServidor(
            emissao
          ),

        status:
          status,

        statusAlteradoEm:
          formatarDataHoraServidor(
            statusAlteradoEm
          ),

        historico:
          historico,

        conclusaoPrevista:
          conclusaoPrevista

      });

    }
  );

  corridas.reverse();

  return corridas;

}


/* ============================================================
   FORMATAR DATA
   ============================================================ */

function formatarDataServidor(
  data
) {

  if (!data) {
    return '';
  }

  if (!(data instanceof Date)) {
    return String(data);
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );

}


/* ============================================================
   FORMATAR DATA ISO
   ============================================================ */

function formatarDataISOServidor(
  data
) {

  if (!data) {
    return '';
  }

  if (!(data instanceof Date)) {

    const texto =
      String(data);

    const partes =
      texto.split('/');

    if (
      partes.length === 3
    ) {

      return (
        partes[2] +
        '-' +
        partes[1].padStart(2, '0') +
        '-' +
        partes[0].padStart(2, '0')
      );

    }

    return '';

  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

}


/* ============================================================
   FORMATAR DATA/HORA ISO
   ============================================================ */

function formatarDataHoraISOServidor(
  data
) {

  if (!data) {
    return '';
  }

  if (!(data instanceof Date)) {
    return '';
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss"
  );

}


/* ============================================================
   FORMATAR HORÁRIO
   ============================================================ */

function formatarHorarioServidor(
  data
) {

  if (!data) {
    return '';
  }

  if (data instanceof Date) {

    return Utilities.formatDate(
      data,
      Session.getScriptTimeZone(),
      'HH:mm'
    );

  }

  const texto =
    String(data).trim();

  if (
    /^\d{1,2}:\d{2}$/.test(
      texto
    )
  ) {

    const partes =
      texto.split(':');

    return (
      String(
        Number(
          partes[0]
        )
      ).padStart(
        2,
        '0'
      ) +
      ':' +
      partes[1]
    );

  }

  const dataConvertida =
    new Date(texto);

  if (
    !isNaN(
      dataConvertida.getTime()
    )
  ) {

    return Utilities.formatDate(
      dataConvertida,
      Session.getScriptTimeZone(),
      'HH:mm'
    );

  }

  return texto;

}


/* ============================================================
   FORMATAR DATA/HORA
   ============================================================ */

function formatarDataHoraServidor(
  data
) {

  if (!data) {
    return '';
  }

  if (!(data instanceof Date)) {
    return String(data);
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm'
  );

}