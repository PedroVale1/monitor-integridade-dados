/**
 * Monitor de Integridade de Dados - Parte 1
 * Cria as bases ficticias (RM, SABER, ONBASE, ServiceNow) com erros plantados
 * e as abas de controle (Inconsistencias, Varreduras).
 *
 * Como usar: rode a funcao criarBasesFicticias() uma unica vez.
 * ATENCAO: ela apaga e recria o conteudo dessas abas.
 */

const BASES = {
  RM: 'RM',
  SABER: 'SAB',
  ONBASE: 'ONB',
  ServiceNow: 'SN',
};

const CABECALHO_BASE = ['ID', 'Matricula', 'CPF', 'Nome', 'Email', 'Setor', 'Atualizado_Em'];

const CABECALHO_INCONSISTENCIAS = [
  'ID', 'Varredura_ID', 'Data_Deteccao', 'Base', 'Registro_ID', 'Matricula', 'CPF',
  'Tipo', 'Campo', 'Valor_Atual', 'Valor_Sugerido', 'Detalhe',
  'Status', 'Resolvido_Em', 'Resolvido_Por',
];

const CABECALHO_VARREDURAS = [
  'ID', 'Data_Hora', 'Registros_Analisados', 'Inconsistencias_Encontradas',
  'Duplicidades', 'Emails_Invalidos', 'Campos_Em_Branco', 'Divergencias',
  'Duracao_Seg', 'Status',
];

const NOMES = [
  'Ana Souza', 'Bruno Lima', 'Carla Mendes', 'Diego Rocha', 'Elaine Castro',
  'Felipe Araújo', 'Gabriela Pinto', 'Hugo Martins', 'Isabela Freitas', 'João Barros',
  'Karina Lopes', 'Lucas Teixeira', 'Marina Queiroz', 'Nelson Farias', 'Olívia Moura',
];

const SETORES = ['TI', 'RH', 'Financeiro', 'Logística', 'Comercial'];

// ---------- Funcao principal ----------

function criarBasesFicticias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colaboradores = gerarColaboradores();

  Object.keys(BASES).forEach(function (base) {
    const linhas = colaboradores.map(function (c, i) {
      return [idRegistro(base, i + 1), c.matricula, c.cpf, c.nome, c.email, c.setor, new Date()];
    });
    plantarErros(base, linhas);
    escreverAba(ss, base, CABECALHO_BASE, linhas);
  });

  escreverAba(ss, 'Inconsistencias', CABECALHO_INCONSISTENCIAS, []);
  escreverAba(ss, 'Varreduras', CABECALHO_VARREDURAS, []);

  // Remove a aba padrao vazia, se existir
  ['Página1', 'Planilha1', 'Sheet1'].forEach(function (nome) {
    const aba = ss.getSheetByName(nome);
    if (aba && aba.getLastRow() === 0) ss.deleteSheet(aba);
  });

  SpreadsheetApp.getUi().alert('Bases fictícias criadas com sucesso!');
}

// ---------- Geracao dos dados ----------

function gerarColaboradores() {
  return NOMES.map(function (nome, i) {
    return {
      matricula: String(1001 + i),
      cpf: gerarCpf(i + 1),
      nome: nome,
      email: gerarEmail(nome),
      setor: SETORES[i % SETORES.length],
    };
  });
}

function gerarEmail(nome) {
  const semAcento = nome.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return semAcento.toLowerCase().replace(/\s+/g, '.') + '@3coracoes.com.br';
}

/** Gera um CPF ficticio com digitos verificadores validos. */
function gerarCpf(semente) {
  const base = String((semente * 12345679 + 100200300) % 1000000000).padStart(9, '0');
  const nums = base.split('').map(Number);
  const d1 = digitoVerificador(nums);
  const d2 = digitoVerificador(nums.concat(d1));
  const s = base + d1 + d2;
  return s.slice(0, 3) + '.' + s.slice(3, 6) + '.' + s.slice(6, 9) + '-' + s.slice(9);
}

function digitoVerificador(nums) {
  let peso = nums.length + 1;
  const soma = nums.reduce(function (acc, d) { return acc + d * peso--; }, 0);
  const resto = (soma * 10) % 11;
  return resto === 10 ? 0 : resto;
}

function idRegistro(base, n) {
  return BASES[base] + '-' + String(n).padStart(4, '0');
}

// ---------- Erros plantados (o que a varredura tera que achar) ----------
// Colunas: 0 ID | 1 Matricula | 2 CPF | 3 Nome | 4 Email | 5 Setor | 6 Atualizado_Em

function plantarErros(base, linhas) {
  if (base === 'RM') {
    // Cadastro duplicado: mesmo CPF com outra matricula
    const dup = linhas[2].slice();
    dup[0] = idRegistro(base, 16);
    dup[1] = '9001';
    linhas.push(dup);
    // Campo obrigatorio em branco
    linhas[5][5] = '';
  }

  if (base === 'SABER') {
    // E-mail invalido (sem @)
    linhas[1][4] = linhas[1][4].replace('@', '');
    // E-mail divergente das outras bases para o mesmo CPF
    linhas[7][4] = 'contato.' + linhas[7][4];
  }

  if (base === 'ONBASE') {
    // E-mail em branco
    linhas[3][4] = '';
    // Matricula duplicada para CPFs diferentes
    linhas[9][1] = linhas[8][1];
    // Nome em branco
    linhas[11][3] = '';
  }

  if (base === 'ServiceNow') {
    // E-mail invalido (com espaco)
    linhas[4][4] = linhas[4][4].replace('.', ' ');
    // CPF invalido (digito verificador errado)
    linhas[10][2] = '123.456.789-00';
    // Colaborador ausente nesta base
    linhas.splice(13, 1);
    // Colaborador que so existe nesta base (orfao)
    linhas.push([idRegistro(base, 16), '9999', gerarCpf(999), 'Usuário Teste Órfão',
      'usuario.teste@3coracoes.com.br', 'TI', new Date()]);
  }
}

// ---------- Escrita na planilha ----------

function escreverAba(ss, nome, cabecalho, linhas) {
  let aba = ss.getSheetByName(nome);
  if (!aba) aba = ss.insertSheet(nome);
  aba.clear();

  aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
  aba.setFrozenRows(1);

  if (linhas.length > 0) {
    // Formato texto nas colunas de dados para nao perder zeros nem virar numero
    aba.getRange(2, 1, linhas.length, cabecalho.length - 1).setNumberFormat('@');
    aba.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  }
  aba.autoResizeColumns(1, cabecalho.length);
}
