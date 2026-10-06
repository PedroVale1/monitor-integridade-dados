/**
 * Monitor de Integridade de Dados - Varredura
 *
 * Le as bases, cruza os colaboradores por CPF (com a matricula como chave
 * reserva) e grava o resultado nas abas Inconsistencias e Varreduras.
 *
 * Funcoes que voce executa:
 *   executarVarredura()            -> roda a varredura (usada tambem pelo gatilho)
 *   ativarVarreduraAutomatica()    -> cria o gatilho periodico
 *   desativarVarreduraAutomatica() -> remove o gatilho
 */

const VARREDURA = {
  bases: ['RM', 'SABER', 'ONBASE', 'ServiceNow'],
  camposObrigatorios: ['Matricula', 'CPF', 'Nome', 'Email', 'Setor'],
  camposComparados: ['Matricula', 'Nome', 'Email', 'Setor'],
  abaInconsistencias: 'Inconsistencias',
  abaVarreduras: 'Varreduras',
  intervaloHoras: 1,
};

const TIPO = {
  CAMPO_EM_BRANCO: 'CAMPO_EM_BRANCO',
  EMAIL_INVALIDO: 'EMAIL_INVALIDO',
  CPF_INVALIDO: 'CPF_INVALIDO',
  CPF_DUPLICADO: 'CPF_DUPLICADO',
  MATRICULA_DUPLICADA: 'MATRICULA_DUPLICADA',
  DIVERGENCIA: 'DIVERGENCIA_ENTRE_BASES',
  AUSENTE: 'AUSENTE_NA_BASE',
  ORFAO: 'REGISTRO_ORFAO',
};

// ---------- Menu na planilha ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Monitor de Integridade')
    .addItem('Executar varredura agora', 'executarVarreduraManual')
    .addSeparator()
    .addItem('Ativar varredura automática', 'ativarVarreduraAutomatica')
    .addItem('Desativar varredura automática', 'desativarVarreduraAutomatica')
    .addSeparator()
    .addItem('Recriar bases fictícias', 'criarBasesFicticias')
    .addToUi();
}

function executarVarreduraManual() {
  const r = executarVarredura();
  SpreadsheetApp.getUi().alert(
    'Varredura ' + r.id + ' concluída.\n\n' +
    'Registros analisados: ' + r.registros + '\n' +
    'Inconsistências pendentes: ' + r.pendentes + '\n' +
    'Novas nesta varredura: ' + r.novas + '\n' +
    'Resolvidas desde a última: ' + r.resolvidas
  );
}

// ---------- Gatilho periodico ----------

function ativarVarreduraAutomatica() {
  removerGatilhos_();
  ScriptApp.newTrigger('executarVarredura')
    .timeBased()
    .everyHours(VARREDURA.intervaloHoras)
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Varredura automática ativada (a cada ' + VARREDURA.intervaloHoras + 'h).');
}

function desativarVarreduraAutomatica() {
  removerGatilhos_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Varredura automática desativada.');
}

function removerGatilhos_() {
  ScriptApp.getProjectTriggers().forEach(function (g) {
    if (g.getHandlerFunction() === 'executarVarredura') ScriptApp.deleteTrigger(g);
  });
}

// ---------- Funcao principal ----------

function executarVarredura() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Já existe uma varredura em andamento.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inicio = new Date();
  const id = 'VAR-' + Utilities.formatDate(inicio, ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmmss');

  try {
    const registros = lerBases_(ss);
    const achados = detectarInconsistencias_(registros);
    const sync = sincronizarInconsistencias_(ss, achados, id, inicio);
    registrarVarredura_(ss, id, inicio, registros.length, sync.ativos, 'Sucesso');
    return {
      id: id,
      registros: registros.length,
      pendentes: sync.ativos.length,
      novas: sync.novas,
      resolvidas: sync.resolvidas,
    };
  } catch (e) {
    registrarVarredura_(ss, id, inicio, 0, [], 'Erro: ' + e.message);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ---------- Leitura das bases ----------

function lerBases_(ss) {
  const registros = [];
  VARREDURA.bases.forEach(function (base) {
    const aba = ss.getSheetByName(base);
    if (!aba) throw new Error('Aba da base não encontrada: ' + base);
    const valores = aba.getDataRange().getValues();
    const col = indiceColunas_(valores[0]);

    valores.slice(1).forEach(function (linha) {
      if (linha.join('').trim() === '') return; // ignora linha vazia
      const campos = {};
      VARREDURA.camposObrigatorios.forEach(function (c) { campos[c] = txt_(linha[col[c]]); });
      registros.push({ base: base, id: txt_(linha[col.ID]), campos: campos });
    });
  });
  return registros;
}

// ---------- Regras de deteccao ----------

function detectarInconsistencias_(registros) {
  const achados = [];
  function add(base, r, tipo, campo, valorAtual, sugerido, detalhe) {
    achados.push({
      base: base,
      registroId: r ? r.id : '',
      matricula: r ? r.campos.Matricula : '',
      cpf: r ? r.campos.CPF : '',
      tipo: tipo, campo: campo,
      valorAtual: valorAtual || '', sugerido: sugerido || '', detalhe: detalhe,
    });
  }

  // 1) Chave da pessoa: CPF valido; se nao houver, resolve pela matricula
  const cpfsPorMatricula = {};
  registros.forEach(function (r) {
    r.cpfValido = cpfValido_(r.campos.CPF);
    if (r.cpfValido && r.campos.Matricula) {
      (cpfsPorMatricula[r.campos.Matricula] = cpfsPorMatricula[r.campos.Matricula] || [])
        .push(soDigitos_(r.campos.CPF));
    }
  });
  const grupos = {};
  registros.forEach(function (r) {
    if (r.cpfValido) r.chave = soDigitos_(r.campos.CPF);
    else if (cpfsPorMatricula[r.campos.Matricula]) r.chave = moda_(cpfsPorMatricula[r.campos.Matricula]).valor;
    else r.chave = null;
    if (r.chave) (grupos[r.chave] = grupos[r.chave] || []).push(r);
  });

  // Valor mais comum do campo entre os OUTROS registros da mesma pessoa
  function sugestao(r, campo) {
    if (!r.chave) return '';
    if (campo === 'CPF') return formatarCpf_(r.chave);
    const valores = grupos[r.chave]
      .filter(function (o) { return o !== r && valorUtilizavel_(o, campo); })
      .map(function (o) { return o.campos[campo]; });
    return valores.length ? moda_(valores).valor : '';
  }

  // 2) Validacoes por registro
  registros.forEach(function (r) {
    VARREDURA.camposObrigatorios.forEach(function (campo) {
      if (!r.campos[campo]) {
        add(r.base, r, TIPO.CAMPO_EM_BRANCO, campo, '', sugestao(r, campo),
          'Campo obrigatório "' + campo + '" em branco.');
      }
    });
    if (r.campos.Email && !emailValido_(r.campos.Email)) {
      add(r.base, r, TIPO.EMAIL_INVALIDO, 'Email', r.campos.Email, sugestao(r, 'Email'),
        'E-mail fora do formato válido.');
    }
    if (r.campos.CPF && !r.cpfValido) {
      add(r.base, r, TIPO.CPF_INVALIDO, 'CPF', r.campos.CPF, sugestao(r, 'CPF'),
        'CPF com dígito verificador inválido.');
    }
  });

  // 3) Duplicidades dentro da mesma base
  function duplicados(campoChave, tipo, rotulo, normalizar) {
    const mapa = {};
    registros.forEach(function (r) {
      if (campoChave === 'CPF' && !r.cpfValido) return;
      const v = normalizar(r.campos[campoChave]);
      if (!v) return;
      const k = r.base + '|' + v;
      (mapa[k] = mapa[k] || []).push(r);
    });
    Object.keys(mapa).forEach(function (k) {
      const lista = mapa[k];
      if (lista.length < 2) return;
      const ids = lista.map(function (r) { return r.id; }).join(', ');
      lista.forEach(function (r) {
        add(r.base, r, tipo, campoChave, r.campos[campoChave], '',
          rotulo + ' repetido(a) na base ' + r.base + ' nos registros: ' + ids + '.');
      });
    });
  }
  duplicados('CPF', TIPO.CPF_DUPLICADO, 'CPF', soDigitos_);
  duplicados('Matricula', TIPO.MATRICULA_DUPLICADA, 'Matrícula', txt_);

  // 4) Comparacao entre bases, pessoa por pessoa
  Object.keys(grupos).forEach(function (chave) {
    const grupo = grupos[chave];

    VARREDURA.camposComparados.forEach(function (campo) {
      const validos = grupo.filter(function (r) { return valorUtilizavel_(r, campo); });
      if (validos.length < 2) return;
      const m = moda_(validos.map(function (r) { return comparavel_(r.campos[campo]); }));
      validos.forEach(function (r) {
        if (comparavel_(r.campos[campo]) === m.valor) return;
        const referencia = validos.filter(function (o) {
          return comparavel_(o.campos[campo]) === m.valor;
        })[0].campos[campo];
        add(r.base, r, TIPO.DIVERGENCIA, campo, r.campos[campo], m.empate ? '' : referencia,
          campo + ' diferente do valor encontrado nas outras bases para o mesmo CPF.');
      });
    });

    const presentes = {};
    grupo.forEach(function (r) { presentes[r.base] = true; });
    const basesPresentes = VARREDURA.bases.filter(function (b) { return presentes[b]; });
    const basesAusentes = VARREDURA.bases.filter(function (b) { return !presentes[b]; });
    const ref = grupo[0];

    if (basesPresentes.length === 1) {
      grupo.forEach(function (r) {
        add(r.base, r, TIPO.ORFAO, 'CPF', r.campos.CPF, '',
          'Colaborador existe apenas na base ' + r.base + '.');
      });
    } else {
      basesAusentes.forEach(function (b) {
        achados.push({
          base: b, registroId: '', matricula: ref.campos.Matricula, cpf: formatarCpf_(chave),
          tipo: TIPO.AUSENTE, campo: 'Registro', valorAtual: '', sugerido: '',
          detalhe: (ref.campos.Nome || 'Colaborador') + ' existe em ' + basesPresentes.join(', ') +
            ' mas não em ' + b + '.',
        });
      });
    }
  });

  return achados;
}

// ---------- Sincronizacao com a aba Inconsistencias ----------
// Nao duplica o que ja esta pendente, respeita o que foi ignorado e
// marca como corrigido o que deixou de aparecer.

function sincronizarInconsistencias_(ss, achados, varreduraId, agora) {
  const aba = ss.getSheetByName(VARREDURA.abaInconsistencias);
  if (!aba) throw new Error('Aba não encontrada: ' + VARREDURA.abaInconsistencias);
  const nCols = aba.getLastColumn();
  const cabecalho = aba.getRange(1, 1, 1, nCols).getValues()[0];
  const c = indiceColunas_(cabecalho);
  const dados = aba.getLastRow() > 1
    ? aba.getRange(2, 1, aba.getLastRow() - 1, nCols).getValues()
    : [];

  function chaveLinha(l) {
    return [txt_(l[c.Base]), txt_(l[c.Registro_ID]) || soDigitos_(txt_(l[c.CPF])),
      txt_(l[c.Tipo]), txt_(l[c.Campo])].join('|');
  }
  function chaveAchado(a) {
    return [a.base, a.registroId || soDigitos_(a.cpf), a.tipo, a.campo].join('|');
  }

  const pendentes = {};
  const ignorados = {};
  dados.forEach(function (l, i) {
    const status = txt_(l[c.Status]);
    if (status === 'Pendente') pendentes[chaveLinha(l)] = i;
    else if (status === 'Ignorado') ignorados[chaveLinha(l)] = true;
  });

  const vistos = {};
  const ativos = [];
  let novas = 0;

  achados.forEach(function (a) {
    const k = chaveAchado(a);
    if (vistos[k] || ignorados[k]) return;
    vistos[k] = true;
    ativos.push(a);

    if (k in pendentes) {
      const l = dados[pendentes[k]];
      l[c.Varredura_ID] = varreduraId;
      l[c.Valor_Atual] = a.valorAtual;
      l[c.Valor_Sugerido] = a.sugerido;
      l[c.Detalhe] = a.detalhe;
    } else {
      const l = cabecalho.map(function () { return ''; });
      l[c.ID] = 'INC-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      l[c.Varredura_ID] = varreduraId;
      l[c.Data_Deteccao] = agora;
      l[c.Base] = a.base;
      l[c.Registro_ID] = a.registroId;
      l[c.Matricula] = a.matricula;
      l[c.CPF] = a.cpf;
      l[c.Tipo] = a.tipo;
      l[c.Campo] = a.campo;
      l[c.Valor_Atual] = a.valorAtual;
      l[c.Valor_Sugerido] = a.sugerido;
      l[c.Detalhe] = a.detalhe;
      l[c.Status] = 'Pendente';
      dados.push(l);
      novas++;
    }
  });

  let resolvidas = 0;
  Object.keys(pendentes).forEach(function (k) {
    if (vistos[k]) return;
    const l = dados[pendentes[k]];
    l[c.Status] = 'Corrigido';
    l[c.Resolvido_Em] = agora;
    l[c.Resolvido_Por] = 'Varredura automática';
    resolvidas++;
  });

  if (dados.length > 0) {
    const colunasData = [c.Data_Deteccao, c.Resolvido_Em];
    cabecalho.forEach(function (_, i) {
      const formato = colunasData.indexOf(i) >= 0 ? 'dd/MM/yyyy HH:mm:ss' : '@';
      aba.getRange(2, i + 1, dados.length, 1).setNumberFormat(formato);
    });
    aba.getRange(2, 1, dados.length, nCols).setValues(dados);
  }

  return { ativos: ativos, novas: novas, resolvidas: resolvidas };
}

// ---------- Log da varredura ----------

function registrarVarredura_(ss, id, inicio, totalRegistros, ativos, status) {
  const aba = ss.getSheetByName(VARREDURA.abaVarreduras);
  if (!aba) throw new Error('Aba não encontrada: ' + VARREDURA.abaVarreduras);

  function contar(tipos) {
    return ativos.filter(function (a) { return tipos.indexOf(a.tipo) >= 0; }).length;
  }
  const duplicidades = contar([TIPO.CPF_DUPLICADO, TIPO.MATRICULA_DUPLICADA]);
  const emails = contar([TIPO.EMAIL_INVALIDO]);
  const brancos = contar([TIPO.CAMPO_EM_BRANCO]);
  // "Divergencias" agrupa os demais tipos: divergencia entre bases, CPF invalido, ausente e orfao
  const divergencias = ativos.length - duplicidades - emails - brancos;
  const duracao = Math.round((new Date() - inicio) / 100) / 10;

  aba.appendRow([id, inicio, totalRegistros, ativos.length, duplicidades, emails, brancos,
    divergencias, duracao, status]);
}

// ---------- Utilitarios ----------

function txt_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function soDigitos_(v) {
  return txt_(v).replace(/\D/g, '');
}

function comparavel_(v) {
  return txt_(v).toLowerCase();
}

function indiceColunas_(cabecalho) {
  const indice = {};
  cabecalho.forEach(function (nome, i) { indice[txt_(nome)] = i; });
  return indice;
}

function emailValido_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function cpfValido_(cpf) {
  const d = soDigitos_(cpf);
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  const n = d.split('').map(Number);
  function dv(qtd) {
    let soma = 0;
    for (let i = 0; i < qtd; i++) soma += n[i] * (qtd + 1 - i);
    const resto = (soma * 10) % 11;
    return resto === 10 ? 0 : resto;
  }
  return dv(9) === n[9] && dv(10) === n[10];
}

function formatarCpf_(digitos) {
  const d = soDigitos_(digitos);
  if (d.length !== 11) return txt_(digitos);
  return d.slice(0, 3) + '.' + d.slice(3, 6) + '.' + d.slice(6, 9) + '-' + d.slice(9);
}

/** O valor pode servir de referencia para comparar ou sugerir? */
function valorUtilizavel_(r, campo) {
  const v = r.campos[campo];
  if (!v) return false;
  if (campo === 'Email') return emailValido_(v);
  if (campo === 'CPF') return r.cpfValido;
  return true;
}

/** Valor mais frequente da lista; "empate" indica que nao ha maioria clara. */
function moda_(valores) {
  const contagem = {};
  const ordem = [];
  valores.forEach(function (v) {
    if (!(v in contagem)) { contagem[v] = 0; ordem.push(v); }
    contagem[v]++;
  });
  let melhor = ordem[0];
  ordem.forEach(function (v) { if (contagem[v] > contagem[melhor]) melhor = v; });
  const empate = ordem.filter(function (v) { return contagem[v] === contagem[melhor]; }).length > 1;
  return { valor: melhor, empate: empate };
}
