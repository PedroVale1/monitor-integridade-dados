# Monitor de Integridade de Dados

Aplicação que encontra e corrige inconsistências nos cadastros de colaboradores mantidos em quatro bases (RM, SABER, ONBASE e ServiceNow). Desenvolvida como estudo de caso, com bases fictícias.

- Uma rotina em **Google Apps Script** varre as bases a cada hora e registra as inconsistências.
- Um app em **AppSheet** mostra as pendências e permite corrigir, alterar ou excluir cadastros com um clique.
- Um **relatório web** mostra o status da última varredura e as decisões tomadas.

## Como funciona

As bases ficam em abas de uma planilha Google, junto com duas abas de controle: `Inconsistencias` (uma linha por problema) e `Varreduras` (registro de cada execução).

1. A varredura cruza os colaboradores por CPF, usando a matrícula como chave reserva, e grava o que encontra em `Inconsistencias`.
2. No app, a pessoa escolhe o que fazer com cada pendência. O app registra a decisão na linha da inconsistência.
3. O Apps Script aplica a decisão na base de origem e fecha a pendência.
4. Uma nova varredura roda em seguida, porque uma ação pode resolver outras pendências do mesmo cadastro.

Os botões do app não escrevem direto nas bases. Assim a regra de negócio fica toda neste repositório, versionada e testável.

## Inconsistências detectadas

| Tipo | O que detecta |
| --- | --- |
| `CAMPO_EM_BRANCO` | Matrícula, CPF, nome, e-mail ou setor vazio |
| `EMAIL_INVALIDO` | E-mail fora do formato |
| `CPF_INVALIDO` | Dígito verificador incorreto |
| `CPF_DUPLICADO` | Mesmo CPF em mais de um registro da mesma base |
| `MATRICULA_DUPLICADA` | Mesma matrícula em mais de um registro da mesma base |
| `DIVERGENCIA_ENTRE_BASES` | Valor diferente da maioria das bases para o mesmo colaborador |
| `AUSENTE_NA_BASE` | Colaborador existe em outras bases, mas falta em uma |
| `REGISTRO_ORFAO` | Colaborador existe em uma única base |

Sempre que possível, a varredura sugere a correção com o valor mais comum nas outras bases.

## Estrutura do repositório

| Arquivo | Papel |
| --- | --- |
| `CriarBases.js` | Gera as bases fictícias com os erros plantados |
| `Varredura.js` | Regras de detecção, sincronização das pendências, menu e gatilhos |
| `Acoes.js` | Aplica nas bases as decisões tomadas no app |
| `Relatorio.js` | Monta os dados do relatório web |
| `RelatorioPagina.html` | Página do relatório web |
| `appsscript.json` | Configuração do projeto Apps Script |

## Menu na planilha

O script adiciona o menu **Monitor de Integridade** à planilha:

| Opção | O que faz |
| --- | --- |
| Executar varredura agora | Roda a varredura na hora |
| Processar ações pendentes | Aplica as decisões registradas no app |
| Ativar automações | Cria os gatilhos da varredura periódica e do processamento das ações |
| Desativar automações | Remove os gatilhos |
| Recriar bases fictícias | Zera os dados e recria as bases de teste |

## Versionamento

O fluxo segue o Git Flow:

| Branch ou tag | Uso |
| --- | --- |
| `feature/...` | Uma branch por funcionalidade, integrada à `develop` por Pull Request |
| `develop` | Código em desenvolvimento |
| `release/v1.0.0` | Versão enviada para homologação (tag `v1.0.0-rc.1`) |
| `main` | Código de produção (tag `v1.0.0`) |

Cada ambiente tem a sua própria pasta no Google Drive, com planilha, script, app e relatório independentes. A versão de cada planilha é nomeada no histórico de versões com o mesmo nome da tag.

| Ambiente | Pasta no Drive | Versão |
| --- | --- | --- |
| Desenvolvimento | `01-Desenvolvimento` | `develop` |
| Homologação | `02-Homologacao` | `v1.0.0-rc.1` |
| Produção | `03-Producao` | `v1.0.0` |

## Como rodar em uma planilha nova

Pré-requisitos: Node.js, Git e o [clasp](https://github.com/google/clasp), com a API do Google Apps Script ativada na conta.

1. Crie uma planilha Google e abra **Extensões > Apps Script** para criar o projeto vinculado.
2. Copie o ID do script em **Configurações do projeto**.
3. Clone este repositório e faça login no clasp:

```bash
git clone https://github.com/PedroVale1/monitor-integridade-dados.git
cd monitor-integridade-dados
clasp login
```

4. No arquivo `.clasp.json`, troque o valor de `scriptId` pelo ID do seu script e envie o código:

```bash
clasp push
```

5. Recarregue a planilha e, pelo menu **Monitor de Integridade**, execute **Recriar bases fictícias**, **Executar varredura agora** e **Ativar automações**.
6. Para o relatório, publique o script em **Implantar > Nova implantação > App da Web**.

## Documentação

A documentação completa, com a arquitetura, o uso do app e os links de cada ambiente, está na pasta `04-Documentacao` do projeto no Google Drive.
