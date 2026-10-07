/**
 * Monitor de Integridade de Dados - Relatorio web
 *
 * Publicado como App da Web. Mostra o status da ultima varredura e o que
 * foi corrigido, alterado, excluido ou ignorado.
 */

const ROTULO_TIPO = {
  CAMPO_EM_BRANCO: 'Campo obrigatório em branco',
  EMAIL_INVALIDO: 'E-mail inválido',
  CPF_INVALIDO: 'CPF inválido',
  CPF_DUPLICADO: 'CPF duplicado',
  MATRICULA_DUPLICADA: 'Matrícula duplicada',
  DIVERGENCIA_ENTRE_BASES: 'Divergência entre bases',
  AUSENTE_NA_BASE: 'Colaborador ausente na base',
  REGISTRO_ORFAO: 'Cadastro só existe nesta base',
};

const ROTULO_STATUS = {
  Corrigido: 'Corrigida',
  Alterado: 'Alterada',
  Excluido: 'Cadastro excluído',
  Ignorado: 'Ignorada',
};

function doGet() {
  const modelo = HtmlService.createTemplateFromFile('RelatorioPagina');
  modelo.d = montarDadosRelatorio_();
  return modelo.evaluate()
    .setTitle('Monitor de Integridade de Dados')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function montarDadosRelatorio_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fuso = ss.getSpreadsheetTimeZone();
  function dataHora(v) {
    return v instanceof Date ? Utilities.formatDate(v, fuso, 'dd/MM/yyyy HH:mm') : txt_(v);
  }
  function plural(n, um, varios) { return n + ' ' + (n === 1 ? um : varios); }

  // Ambiente a partir do sufixo do nome da planilha (_DEV, _HML, _PROD)
  const sufixo = (/_(DEV|HML|PROD)$/i.exec(ss.getName()) || [])[1];
  const ambiente = { DEV: 'Desenvolvimento', HML: 'Homologação', PROD: 'Produção' }[(sufixo || '').toUpperCase()] || '';

  // Ultima varredura
  let ultima = null;
  const abaV = ss.getSheetByName(VARREDURA.abaVarreduras);
  if (abaV && abaV.getLastRow() > 1) {
    const nCols = abaV.getLastColumn();
    const cv = indiceColunas_(abaV.getRange(1, 1, 1, nCols).getValues()[0]);
    const l = abaV.getRange(abaV.getLastRow(), 1, 1, nCols).getValues()[0];
    ultima = {
      id: txt_(l[cv.ID]),
      quando: dataHora(l[cv.Data_Hora]),
      registros: Number(l[cv.Registros_Analisados]) || 0,
      duracao: String(l[cv.Duracao_Seg]).replace('.', ','),
      sucesso: txt_(l[cv.Status]) === 'Sucesso',
      status: txt_(l[cv.Status]),
    };
  }

  // Inconsistencias
  const contagem = { Pendente: 0, Corrigido: 0, Alterado: 0, Excluido: 0, Ignorado: 0 };
  const porBase = {};
  VARREDURA.bases.forEach(function (b) { porBase[b] = { base: b, pendentes: 0, tratadas: 0 }; });
  const porTipo = {};
  const decisoes = [];

  const abaI = ss.getSheetByName(VARREDURA.abaInconsistencias);
  if (abaI && abaI.getLastRow() > 1) {
    const nCols = abaI.getLastColumn();
    const c = indiceColunas_(abaI.getRange(1, 1, 1, nCols).getValues()[0]);
    abaI.getRange(2, 1, abaI.getLastRow() - 1, nCols).getValues().forEach(function (l) {
      let status = txt_(l[c.Status]);
      if (!status) return;
      if (status === 'Em processamento') status = 'Pendente';
      if (!(status in contagem)) return;
      contagem[status]++;

      const base = txt_(l[c.Base]);
      const tipo = txt_(l[c.Tipo]);
      if (!porBase[base]) porBase[base] = { base: base, pendentes: 0, tratadas: 0 };

      if (status === 'Pendente') {
        porBase[base].pendentes++;
        porTipo[tipo] = (porTipo[tipo] || 0) + 1;
      } else {
        porBase[base].tratadas++;
        const quando = l[c.Resolvido_Em];
        decisoes.push({
          ordem: quando instanceof Date ? quando.getTime() : 0,
          quando: dataHora(quando),
          base: base,
          registro: txt_(l[c.Registro_ID]) || 'CPF ' + txt_(l[c.CPF]),
          tipo: ROTULO_TIPO[tipo] || tipo,
          campo: txt_(l[c.Campo]),
          status: status,
          rotuloStatus: ROTULO_STATUS[status] || status,
          por: txt_(l[c.Resolvido_Por]) || 'Não identificado',
        });
      }
    });
  }

  const total = contagem.Pendente + contagem.Corrigido + contagem.Alterado + contagem.Excluido + contagem.Ignorado;
  const tratadas = total - contagem.Pendente;
  function pct(n) { return total ? Math.round((n / total) * 1000) / 10 : 0; }

  const maiorBase = Math.max.apply(null, Object.keys(porBase).map(function (b) {
    return porBase[b].pendentes + porBase[b].tratadas;
  }).concat([1]));
  const listaTipos = Object.keys(porTipo).map(function (t) {
    return { tipo: ROTULO_TIPO[t] || t, pendentes: porTipo[t] };
  }).sort(function (a, b) { return b.pendentes - a.pendentes; });
  const maiorTipo = listaTipos.length ? listaTipos[0].pendentes : 1;

  return {
    ambiente: ambiente,
    ultima: ultima,
    titulo: contagem.Pendente === 0
      ? 'Nenhuma inconsistência aguardando decisão'
      : plural(contagem.Pendente, 'inconsistência aguarda decisão', 'inconsistências aguardam decisão'),
    total: total,
    tratadas: tratadas,
    percentualTratado: String(pct(tratadas)).replace('.', ','),
    faixas: [
      { chave: 'corrigido', rotulo: 'Corrigidas', n: contagem.Corrigido, pct: pct(contagem.Corrigido) },
      { chave: 'alterado', rotulo: 'Alteradas', n: contagem.Alterado, pct: pct(contagem.Alterado) },
      { chave: 'excluido', rotulo: 'Cadastros excluídos', n: contagem.Excluido, pct: pct(contagem.Excluido) },
      { chave: 'ignorado', rotulo: 'Ignoradas', n: contagem.Ignorado, pct: pct(contagem.Ignorado) },
      { chave: 'pendente', rotulo: 'Pendentes', n: contagem.Pendente, pct: pct(contagem.Pendente) },
    ],
    porBase: Object.keys(porBase).map(function (b) {
      const x = porBase[b];
      return {
        base: x.base, pendentes: x.pendentes, tratadas: x.tratadas,
        pctPendentes: Math.round((x.pendentes / maiorBase) * 100),
        pctTratadas: Math.round((x.tratadas / maiorBase) * 100),
      };
    }),
    porTipo: listaTipos.map(function (t) {
      return { tipo: t.tipo, pendentes: t.pendentes, pct: Math.round((t.pendentes / maiorTipo) * 100) };
    }),
    decisoes: decisoes.sort(function (a, b) { return b.ordem - a.ordem; }).slice(0, 12),
    geradoEm: Utilities.formatDate(new Date(), fuso, 'dd/MM/yyyy HH:mm'),
  };
}
