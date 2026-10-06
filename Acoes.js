/**
 * Monitor de Integridade de Dados - Acoes do app
 *
 * O AppSheet registra a decisao do usuario na aba Inconsistencias
 * (Status = "Em processamento" + Acao_Solicitada). Este arquivo aplica a
 * decisao na base de origem e fecha a inconsistencia.
 *
 *   CORRIGIR -> grava o Valor_Sugerido no campo (ou cria o cadastro ausente)
 *   ALTERAR  -> grava o Novo_Valor informado pelo usuario
 *   EXCLUIR  -> remove o cadastro da base
 */

const STATUS_FINAL = { CORRIGIR: 'Corrigido', ALTERAR: 'Alterado', EXCLUIR: 'Excluido' };
const CAMPOS_EDITAVEIS = ['Matricula', 'CPF', 'Nome', 'Email', 'Setor'];

/** Gatilho de mudanca na planilha (disparado, por exemplo, pelo AppSheet). */
function aoAlterarPlanilha(e) {
  processarAcoes();
}

function processarAcoesManual() {
  const n = processarAcoes();
  SpreadsheetApp.getUi().alert(n + ' ação(ões) processada(s).');
}

/** Aplica todas as acoes pendentes. Retorna quantas foram processadas. */
function processarAcoes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 0;

  let processadas = 0;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(VARREDURA.abaInconsistencias);
    if (!aba || aba.getLastRow() < 2) return 0;

    const nCols = aba.getLastColumn();
    const c = indiceColunas_(aba.getRange(1, 1, 1, nCols).getValues()[0]);
    if (c.Acao_Solicitada === undefined || c.Novo_Valor === undefined) {
      throw new Error('Colunas Acao_Solicitada e Novo_Valor não encontradas na aba Inconsistencias.');
    }
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, nCols).getValues();

    dados.forEach(function (l, i) {
      const acao = txt_(l[c.Acao_Solicitada]).toUpperCase();
      if (!acao || txt_(l[c.Status]) !== 'Em processamento') return;

      const linha = i + 2;
      try {
        aplicarAcao_(ss, acao, {
          base: txt_(l[c.Base]),
          registroId: txt_(l[c.Registro_ID]),
          cpf: txt_(l[c.CPF]),
          tipo: txt_(l[c.Tipo]),
          campo: txt_(l[c.Campo]),
          sugerido: txt_(l[c.Valor_Sugerido]),
          novoValor: txt_(l[c.Novo_Valor]),
        });
        gravarCelulas_(aba, linha, c, {
          Status: STATUS_FINAL[acao],
          Resolvido_Em: new Date(),
          Resolvido_Por: txt_(l[c.Resolvido_Por]) || 'Não identificado',
        });
      } catch (erro) {
        // Devolve para a fila e deixa o motivo visivel para o usuario
        gravarCelulas_(aba, linha, c, {
          Status: 'Pendente',
          Acao_Solicitada: 'ERRO: ' + erro.message,
          Resolvido_Por: '',
        });
      }
      processadas++;
    });
  } finally {
    lock.releaseLock();
  }

  // Reavalia as bases: outras inconsistencias do mesmo cadastro podem ter sido resolvidas
  if (processadas > 0) executarVarredura();
  return processadas;
}

function aplicarAcao_(ss, acao, inc) {
  if (!STATUS_FINAL[acao]) throw new Error('Ação desconhecida: ' + acao);
  const aba = ss.getSheetByName(inc.base);
  if (!aba) throw new Error('Base não encontrada: ' + inc.base);
  const valores = aba.getDataRange().getValues();
  const col = indiceColunas_(valores[0]);

  if (inc.tipo === TIPO.AUSENTE) {
    if (acao !== 'CORRIGIR') throw new Error('Para colaborador ausente a única ação é criar o cadastro.');
    criarCadastroAusente_(ss, aba, valores, col, inc);
    return;
  }

  let linha = -1;
  for (let i = 1; i < valores.length; i++) {
    if (txt_(valores[i][col.ID]) === inc.registroId) { linha = i + 1; break; }
  }

  if (acao === 'EXCLUIR') {
    if (linha > 0) aba.deleteRow(linha); // se ja nao existe, o objetivo esta cumprido
    return;
  }

  if (linha < 0) throw new Error('Registro ' + inc.registroId + ' não encontrado na base ' + inc.base + '.');
  if (CAMPOS_EDITAVEIS.indexOf(inc.campo) < 0) throw new Error('Campo não editável: ' + inc.campo);

  const valor = acao === 'CORRIGIR' ? inc.sugerido : inc.novoValor;
  if (!valor) {
    throw new Error(acao === 'CORRIGIR'
      ? 'Não há valor sugerido para aplicar.'
      : 'Informe o novo valor antes de alterar.');
  }

  aba.getRange(linha, col[inc.campo] + 1).setNumberFormat('@').setValue(valor);
  if (col.Atualizado_Em !== undefined) aba.getRange(linha, col.Atualizado_Em + 1).setValue(new Date());
}

/** Cria na base de destino o cadastro que existe nas outras bases. */
function criarCadastroAusente_(ss, abaDestino, valoresDestino, colDestino, inc) {
  const cpf = soDigitos_(inc.cpf);
  let origem = null;

  VARREDURA.bases.some(function (base) {
    if (base === inc.base) return false;
    const aba = ss.getSheetByName(base);
    if (!aba) return false;
    const valores = aba.getDataRange().getValues();
    const col = indiceColunas_(valores[0]);
    for (let i = 1; i < valores.length; i++) {
      if (soDigitos_(valores[i][col.CPF]) === cpf) {
        origem = { linha: valores[i], col: col };
        return true;
      }
    }
    return false;
  });
  if (!origem) throw new Error('Cadastro de origem não encontrado para o CPF ' + inc.cpf + '.');

  // Proximo ID no padrao da base (ex.: SN-0017)
  let prefixo = inc.base + '-';
  let maior = 0;
  let digitos = 4;
  valoresDestino.slice(1).forEach(function (l) {
    const m = /^(.*-)(\d+)$/.exec(txt_(l[colDestino.ID]));
    if (!m) return;
    prefixo = m[1];
    digitos = m[2].length;
    maior = Math.max(maior, Number(m[2]));
  });
  const novoId = prefixo + String(maior + 1).padStart(digitos, '0');

  const cabecalho = valoresDestino[0];
  const novaLinha = cabecalho.map(function (nome) {
    const n = txt_(nome);
    if (n === 'ID') return novoId;
    if (n === 'Atualizado_Em') return new Date();
    return origem.col[n] !== undefined ? txt_(origem.linha[origem.col[n]]) : '';
  });

  const linha = abaDestino.getLastRow() + 1;
  cabecalho.forEach(function (nome, i) {
    if (txt_(nome) !== 'Atualizado_Em') abaDestino.getRange(linha, i + 1).setNumberFormat('@');
  });
  abaDestino.getRange(linha, 1, 1, cabecalho.length).setValues([novaLinha]);
}

/** Grava apenas as celulas informadas, para nao sobrescrever edicoes simultaneas do app. */
function gravarCelulas_(aba, linha, c, valores) {
  Object.keys(valores).forEach(function (nome) {
    if (c[nome] === undefined) return;
    aba.getRange(linha, c[nome] + 1).setValue(valores[nome]);
  });
}
