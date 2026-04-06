/**
 * Renderiza o arquivo HTML principal do sistema.
 * Essa é a porta de entrada do seu Web App.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Cliente');
}

/**
 * Abre a planilha principal pelo ID.
 *
 * IMPORTANTE:
 * - Esse ID precisa ser de uma planilha Google Sheets real
 * - Se for arquivo Excel (.xlsx) no Drive, pode dar erro
 */
function getPlanilha() {
  return SpreadsheetApp.openById("1jHZO6UdehPbtuvGZ6QlW_aEUF_rtG0p8PMRMYaNMLNg");
}

/**
 * Busca uma aba pelo nome e já valida se ela existe.
 *
 * Por que isso ajuda?
 * Antes, se a aba estivesse com nome errado ou apagada,
 * o código quebrava com erro pouco claro.
 * Agora a mensagem fica objetiva.
 */
function getAba(nome) {
  const aba = getPlanilha().getSheetByName(nome);

  if (!aba) {
    throw new Error(`A aba "${nome}" não foi encontrada.`);
  }

  return aba;
}

/**
 * Lê todas as abas necessárias e monta a lista de processos
 * no formato que o front-end precisa.
 *
 * O front espera cada processo com:
 * - ids
 * - campos do processo
 * - nomes resolvidos de cliente, produto e unidade
 *
 * Aqui acontece o "join manual" entre as abas.
 */
function buscarProcessos() {
  const processos = getAba("Processos").getDataRange().getDisplayValues();
  const clientes = getAba("Clientes").getDataRange().getDisplayValues();
  const produtos = getAba("Produtos").getDataRange().getDisplayValues();
  const unidades = getAba("Unidades").getDataRange().getDisplayValues();

  // Esses mapas servem para trocar ID por nome.
  // Exemplo:
  // clienteId "CLI_1" -> "Banco XPTO"
  const mapC = {};
  const mapP = {};
  const mapU = {};

  /**
   * Monta mapa de clientes
   * [0] = id
   * [1] = nome
   */
  for (let i = 1; i < clientes.length; i++) {
    mapC[clientes[i][0]] = clientes[i][1];
  }

  /**
   * Monta mapa de produtos
   * [0] = id
   * [2] = nome
   *
   * Aqui mantive sua estrutura original.
   * Se a planilha mudar de ordem, isso precisa ser ajustado.
   */
  for (let i = 1; i < produtos.length; i++) {
    mapP[produtos[i][0]] = produtos[i][2];
  }

  /**
   * Monta mapa de unidades
   * [0] = id
   * [2] = nome
   */
  for (let i = 1; i < unidades.length; i++) {
    mapU[unidades[i][0]] = unidades[i][2];
  }

  const lista = [];

  /**
   * Começa em 1 para ignorar o cabeçalho.
   * Cada linha da aba "Processos" vira um objeto JS.
   */
  for (let i = 1; i < processos.length; i++) {
    const l = processos[i];

    /**
     * Se a linha estiver vazia ou sem ID, ela é ignorada.
     * Isso evita trazer lixo de linhas em branco da planilha.
     */
    if (!l[0]) continue;

    lista.push({
      id: l[0],
      clienteId: l[1],
      produtoId: l[2],
      unidadeId: l[3],
      codigoContrato: l[4],
      valorContrato: l[5],
      tipoGarantia: l[6],
      andamento: l[7],
      dataCriacao: l[8],
      dataAtualizacao: l[9],

      /**
       * Aqui convertemos IDs para nomes amigáveis.
       * Se o ID não for encontrado no mapa, o código retorna o valor original.
       * Isso evita que o front quebre se faltar algum relacionamento.
       */
      clienteNome: mapC[l[1]] || l[1] || "",
      produtoNome: mapP[l[2]] || l[2] || "",
      unidadeNome: mapU[l[3]] || l[3] || ""
    });
  }

  return lista;
}

/**
 * Retorna as opções para preencher os selects do formulário.
 *
 * O front usa isso para popular:
 * - Cliente
 * - Produto
 * - Unidade
 *
 * Estrutura esperada:
 * {
 *   clientes: [{ id, nome }],
 *   produtos: [{ id, nome }],
 *   unidades: [{ id, nome }]
 * }
 */
function buscarOpcoesFormulario() {
  const clientes = getAba("Clientes").getDataRange().getDisplayValues();
  const produtos = getAba("Produtos").getDataRange().getDisplayValues();
  const unidades = getAba("Unidades").getDataRange().getDisplayValues();

  return {
    clientes: clientes
      .slice(1) // ignora cabeçalho
      .filter(l => l[0]) // ignora linhas sem ID
      .map(l => ({ id: l[0], nome: l[1] })),

    produtos: produtos
      .slice(1)
      .filter(l => l[0])
      .map(l => ({ id: l[0], nome: l[2] })),

    unidades: unidades
      .slice(1)
      .filter(l => l[0])
      .map(l => ({ id: l[0], nome: l[2] }))
  };
}

/**
 * Cria um novo processo na aba "Processos".
 *
 * Antes o ID era baseado em getLastRow().
 * Isso podia gerar repetição se linhas fossem apagadas.
 *
 * Agora o ID usa timestamp, o que reduz muito esse risco
 * sem mudar a estrutura do sistema.
 */
function criarProcesso(d) {
  const sheet = getAba("Processos");
  const id = "PROCESSO_" + new Date().getTime();

  sheet.appendRow([
    id,
    d.clienteId || "",
    d.produtoId || "",
    d.unidadeId || "",
    d.codigoContrato || "",
    d.valorContrato || "",
    d.tipoGarantia || "",
    d.andamento || "",
    d.dataCriacao || "",
    d.dataAtualizacao || ""
  ]);

  return "Criado!";
}

/**
 * Edita um processo existente.
 *
 * Regra:
 * - procura pelo ID na primeira coluna
 * - quando encontra, atualiza da coluna 2 até a 10
 *
 * Observação importante:
 * - o ID não é alterado
 * - por isso o range começa na coluna 2
 */
function editarProcesso(d) {
  const sheet = getAba("Processos");
  const dados = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === d.id) {
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        d.clienteId || "",
        d.produtoId || "",
        d.unidadeId || "",
        d.codigoContrato || "",
        d.valorContrato || "",
        d.tipoGarantia || "",
        d.andamento || "",
        d.dataCriacao || "",
        d.dataAtualizacao || ""
      ]]);

      return "Atualizado!";
    }
  }

  /**
   * Se não achar o ID, lança erro claro.
   * O front recebe isso no withFailureHandler.
   */
  throw new Error("Processo não encontrado para edição.");
}

/**
 * Exclui um processo pelo ID.
 *
 * Regra:
 * - procura a linha com o ID informado
 * - exclui a linha inteira
 */
function excluirProcesso(id) {
  const sheet = getAba("Processos");
  const dados = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === id) {
      sheet.deleteRow(i + 1);
      return "Excluído!";
    }
  }

  throw new Error("Processo não encontrado para exclusão.");
}