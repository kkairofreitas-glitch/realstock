if (process.env.NODE_ENV !== "production") {
  try {
    require("dotenv").config();
  } catch (erro) {
    console.warn("dotenv não carregado:", erro.message);
  }
}

const {
  testarConexao,
  criarTabelas,
  carregarUsuariosPostgres,
  salvarUsuariosPostgres,
  salvarProdutosPostgres,
  carregarProdutosPostgres,
  salvarEnderecamentosPostgres,
  carregarEnderecamentosPostgres,
  salvarContagensPostgres,
  carregarContagensPostgres,
  salvarConfiguracaoPostgres,
  carregarConfiguracaoPostgres,
  limparInventarioPostgres,
  salvarContagemSemBasePostgres,
  carregarContagemSemBasePostgres,
  salvarFinalizacoesSemBasePostgres,
  carregarFinalizacoesSemBasePostgres,
  carregarContagensWmsPostgres,
  salvarContagemWmsPostgres,
  excluirContagemWmsPostgres,
  salvarFinalizacaoWmsPostgres,
  carregarFinalizacoesWmsPostgres,
  salvarBaseEsperadaWmsPostgres,
  salvarBaseEsperadaWmsLotePostgres,
  buscarBaseEsperadaWmsPostgres,
  carregarBaseEsperadaWmsPostgres,
  buscarBaseEsperadaWmsPorIdPostgres,
  atualizarBaseEsperadaWmsPostgres,
  excluirBaseEsperadaWmsPostgres,
} = require("./db");

const express = require("express");
const fileUpload = require("express-fileupload");
const session = require("express-session");
const path = require("path");
const fs = require("fs");
const readline = require("readline");
const PdfPrinter = require("pdfmake");
const http = require("http");
const WebSocket = require("ws");
const ExcelJS = require("exceljs");
const sqlite3 = require("sqlite3").verbose();

const pastaBanco = path.join(__dirname, "data");

if (!fs.existsSync(pastaBanco)) {
  fs.mkdirSync(pastaBanco, { recursive: true });
}

const db = new sqlite3.Database(path.join(pastaBanco, "inventario.db"));
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS produtos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codigoBarras TEXT,
      codigo TEXT,
      descricao TEXT,
      categoria TEXT,
      custoUnitario REAL,
      qtdeCongelada REAL,
      qtdeContada REAL
    )
  `);
});

const fonts = {
  Helvetica: {
    normal: "Helvetica",
    bold: "Helvetica-Bold",
    italics: "Helvetica-Oblique",
    bolditalics: "Helvetica-BoldOblique",
  },
};

const app = express();
const port = process.env.PORT || 5000;
function caminhoPublico(nomeArquivo) {
  const caminhoRaiz = path.join(__dirname, nomeArquivo);
  const caminhoPublic = path.join(__dirname, "public", nomeArquivo);

  if (fs.existsSync(caminhoRaiz)) {
    return caminhoRaiz;
  }

  if (fs.existsSync(caminhoPublic)) {
    return caminhoPublic;
  }

  return caminhoRaiz;
}

let usuarios = [];
let inventario = [];
let historicoAlteracoes = [];
let enderecamentos = [];

let contagens = [];
let historicoAuditoriaItens = [];
let auditoriaImportacao = {
  totalImportadoBruto: 0,
  totalUnicosBruto: 0,
  duplicatasRemovidas: 0,
  itensZeradosIgnorados: 0,
};
let itemAuditoriaAtual = null;
let tipoUltimaImportacao = "--";
let ultimaImportacao = {
  arquivo: "--",
  tipo: "--",
  horario: null,
  status: "--",
  observacao: "Sem importações nesta sessão.",
};
let contagemSemBase = [];
let modoOperacao = "com-base";
let finalizacoesSemBase = [];
let layoutsSalvos = [];
let layoutsUniversais = [];
function normalizarModoOperacao(valor) {
  const modo = String(valor || "com-base").trim();

  if (modo === "sem-base") return "sem-base";
  if (modo === "wms") return "wms";

  return "com-base";
}

function obterEnderecosSemBaseConhecidos() {
  const numeros = new Set();

  (Array.isArray(contagemSemBase) ? contagemSemBase : []).forEach((item) => {
    const enderecos = Array.isArray(item?.enderecos)
      ? item.enderecos
      : [];

    enderecos.forEach((registro) => {
      const numero = String(
        registro?.enderecoNumero || ""
      ).trim();

      if (numero) {
        numeros.add(numero);
      }
    });
  });

  (
    Array.isArray(finalizacoesSemBase)
      ? finalizacoesSemBase
      : []
  ).forEach((item) => {
    const numero = String(
      item?.enderecoNumero || ""
    ).trim();

    if (numero) {
      numeros.add(numero);
    }
  });

  return numeros;
}

function resolverModoEventoEndereco(evento) {
  const modoGravado = String(
    evento?.modoOperacao || ""
  ).trim();

  if (
    modoGravado === "com-base" ||
    modoGravado === "sem-base" ||
    modoGravado === "wms"
  ) {
    return modoGravado;
  }

  /*
    Compatibilidade com registros antigos.

    Antes, algumas transmissões/finalizações eram salvas
    sem o campo modoOperacao.

    Se o endereço também existe nos dados persistidos do
    sem-base, tratamos esse evento legado como sem-base.
  */
  const numero = String(
    evento?.enderecoNumero || ""
  ).trim();

  if (
    numero &&
    obterEnderecosSemBaseConhecidos().has(numero)
  ) {
    return "sem-base";
  }

  return "com-base";
}

function eventoPertenceAoModo(
  evento,
  modoReferencia = modoOperacao
) {
  return (
    resolverModoEventoEndereco(evento) ===
    normalizarModoOperacao(modoReferencia)
  );
}

const dataDir = path.join(__dirname, "data");
const contagensPath = path.join(dataDir, "contagens.json");
const layoutTxtPath = path.join(dataDir, "layout-txt.json");
const layoutsTxtPath = path.join(dataDir, "layouts-txt.json");
const layoutsUniversaisPath = path.join(
  dataDir,
  "layouts-universais.json"
);
const layoutsExportacaoPath = path.join(
  dataDir,
  "layouts-exportacao.json"
);
const usuariosPath = path.join(dataDir, "usuarios.json");
const enderecamentosPath = path.join(dataDir, "enderecamentos.json");
const encerramentosDir = path.join(dataDir, "encerramentos");
const ultimoEncerramentoPath = path.join(
  encerramentosDir,
  "ultimo-encerramento.json"
);
const contagemSemBasePath = path.join(dataDir, "contagem-sem-base.json");
const configModoPath = path.join(dataDir, "config-modo.json");
const finalizacoesSemBasePath = path.join(
  dataDir,
  "finalizacoes-sem-base.json"
);

const FORMATOS_PREVIEW_UNIVERSAL = new Set([
  ".txt",
  ".csv",
  ".json",
  ".xlsx",
]);


function limparEventosDoModoNosEnderecamentos(
  modoReferencia
) {
  const modo = normalizarModoOperacao(
    modoReferencia
  );

  enderecamentos = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).map((endereco) => {
    const transmissoes = (
      Array.isArray(endereco?.transmissoes)
        ? endereco.transmissoes
        : []
    ).filter(
      (evento) =>
        !eventoPertenceAoModo(
          evento,
          modo
        )
    );

    const finalizacoes = (
      Array.isArray(endereco?.finalizacoes)
        ? endereco.finalizacoes
        : []
    ).filter(
      (evento) =>
        !eventoPertenceAoModo(
          evento,
          modo
        )
    );

    return {
      ...endereco,

      transmissoes,
      finalizacoes,

      contagensRecebidas:
        transmissoes.filter(
          (evento) =>
            evento?.tipo === "transmissao"
        ).length,

      finalizadoViaColetor:
        finalizacoes.length > 0,

      ultimaContagemEm:
        transmissoes
          .filter((evento) => evento?.data)
          .sort(
            (a, b) =>
              new Date(b.data) -
              new Date(a.data)
          )[0]?.data || null,

      finalizadoEm:
        finalizacoes
          .filter((evento) => evento?.data)
          .sort(
            (a, b) =>
              new Date(b.data) -
              new Date(a.data)
          )[0]?.data || null,

      atualizadoEm:
        new Date().toISOString(),
    };
  });
}

function haAtividadeNoModo(
  modoReferencia = modoOperacao
) {
  const modo = normalizarModoOperacao(
    modoReferencia
  );

  if (modo === "sem-base") {
    if (
      (
        Array.isArray(contagemSemBase) &&
        contagemSemBase.length > 0
      ) ||
      (
        Array.isArray(finalizacoesSemBase) &&
        finalizacoesSemBase.length > 0
      )
    ) {
      return true;
    }
  }

  if (modo === "com-base") {
    if (
      (
        Array.isArray(inventario) &&
        inventario.length > 0
      ) ||
      (
        Array.isArray(contagens) &&
        contagens.length > 0
      )
    ) {
      return true;
    }
  }

  return (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).some((endereco) => {
    const possuiTransmissao = (
      Array.isArray(endereco?.transmissoes)
        ? endereco.transmissoes
        : []
    ).some(
      (evento) =>
        !evento?.excluida &&
        eventoPertenceAoModo(
          evento,
          modo
        )
    );

    const possuiFinalizacao = (
      Array.isArray(endereco?.finalizacoes)
        ? endereco.finalizacoes
        : []
    ).some(
      (evento) =>
        !evento?.excluida &&
        eventoPertenceAoModo(
          evento,
          modo
        )
    );

    return (
      possuiTransmissao ||
      possuiFinalizacao
    );
  });
}


function limparNomeColunaUniversal(valor, indice) {
  const texto = String(valor ?? "").trim();

  return texto || `coluna_${indice + 1}`;
}

function criarNomesUnicosUniversal(colunas = []) {
  const ocorrencias = new Map();

  return colunas.map((coluna, indice) => {
    const nomeBase =
      limparNomeColunaUniversal(coluna, indice);

    const chave = nomeBase.toLowerCase();

    const quantidade =
      (ocorrencias.get(chave) || 0) + 1;

    ocorrencias.set(chave, quantidade);

    return quantidade === 1
      ? nomeBase
      : `${nomeBase}_${quantidade}`;
  });
}

function detectarDelimitadorUniversal(linha = "") {
  const candidatos = [";", ",", "\t", "|"];

  let melhor = null;
  let maiorQuantidade = 1;

  candidatos.forEach((delimitador) => {
    const quantidade =
      String(linha).split(delimitador).length;

    if (quantidade > maiorQuantidade) {
      maiorQuantidade = quantidade;
      melhor = delimitador;
    }
  });

  return melhor;
}

function separarLinhaDelimitadaUniversal(
  linha,
  delimitador
) {
  const texto = String(linha ?? "");

  if (!delimitador) {
    return [texto];
  }

  const valores = [];

  let valorAtual = "";
  let dentroDeAspas = false;

  for (
    let indice = 0;
    indice < texto.length;
    indice += 1
  ) {
    const caractere = texto[indice];
    const proximo = texto[indice + 1];

    if (caractere === '"') {
      if (dentroDeAspas && proximo === '"') {
        valorAtual += '"';
        indice += 1;
      } else {
        dentroDeAspas = !dentroDeAspas;
      }

      continue;
    }

    if (
      caractere === delimitador &&
      !dentroDeAspas
    ) {
      valores.push(valorAtual.trim());
      valorAtual = "";
      continue;
    }

    valorAtual += caractere;
  }

  valores.push(valorAtual.trim());

  return valores;
}

function converterMatrizUniversal(
  matriz = [],
  possuiCabecalho = true
) {
  const linhas = matriz.filter(
    (linha) =>
      Array.isArray(linha) &&
      linha.some(
        (valor) =>
          String(valor ?? "").trim() !== ""
      )
  );

  if (!linhas.length) {
    return {
      colunas: [],
      itens: [],
    };
  }

  const totalColunas = linhas.reduce(
    (maior, linha) =>
      Math.max(maior, linha.length),
    0
  );

  const colunasOriginais = possuiCabecalho
    ? Array.from(
        { length: totalColunas },
        (_, indice) => linhas[0]?.[indice]
      )
    : Array.from(
        { length: totalColunas },
        (_, indice) => `coluna_${indice + 1}`
      );

  const colunas =
    criarNomesUnicosUniversal(
      colunasOriginais
    );

  const indiceInicio =
    possuiCabecalho ? 1 : 0;

  const itens = linhas
    .slice(indiceInicio)
    .map((linha, indiceLinha) => {
      const item = {
        __linha:
          indiceLinha +
          indiceInicio +
          1,
      };

      colunas.forEach(
        (coluna, indiceColuna) => {
          item[coluna] =
            linha[indiceColuna] ?? "";
        }
      );

      return item;
    });

  return {
    colunas,
    itens,
  };
}

function interpretarTextoDelimitadoUniversal(
  conteudo,
  opcoes = {}
) {
  const linhas = String(conteudo || "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter(
      (linha) =>
        String(linha).trim() !== ""
    );

  if (!linhas.length) {
    return {
      colunas: [],
      itens: [],
      delimitador: null,
    };
  }

  const delimitadorInformado =
    String(opcoes.delimitador || "");

  const delimitador =
    delimitadorInformado === "\\t"
      ? "\t"
      : delimitadorInformado ||
        detectarDelimitadorUniversal(
          linhas[0]
        );

  if (!delimitador) {
    return {
      colunas: [],
      itens: [],
      delimitador: null,
      semDelimitador: true,
    };
  }

  const matriz = linhas.map((linha) =>
    separarLinhaDelimitadaUniversal(
      linha,
      delimitador
    )
  );

  return {
    ...converterMatrizUniversal(
      matriz,
      opcoes.possuiCabecalho !== false
    ),

    delimitador,
    semDelimitador: false,
  };
}

function interpretarTxtPosicaoFixaUniversal(
  conteudo
) {
  const linhas = String(conteudo || "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter(
      (linha) =>
        String(linha).trim() !== ""
    );

  const colunas = [
    "codigoBarras",
    "codigo",
    "descricao",
    "custoUnitario",
    "qtdeCongelada",
    "categoria",
    "tipo",
  ];

  const itens = linhas.map(
    (linha, indice) => ({
      __linha: indice + 1,

      codigoBarras:
        extrairCampoLinha(
          linha,
          "codigoBarras"
        ),

      codigo:
        extrairCampoLinha(
          linha,
          "codigo"
        ),

      descricao:
        extrairCampoLinha(
          linha,
          "descricao"
        ),

      custoUnitario:
        extrairCampoLinha(
          linha,
          "custoUnitario"
        ),

      qtdeCongelada:
        extrairCampoLinha(
          linha,
          "qtdeCongelada"
        ),

      categoria:
        extrairCampoLinha(
          linha,
          "categoria"
        ),

      tipo:
        extrairCampoLinha(
          linha,
          "tipo"
        ),
    })
  );

  return {
    colunas,
    itens,
    delimitador: null,
    tipoLeitura: "posicao-fixa",
  };
}

function interpretarJsonUniversal(conteudo) {
  const dados = JSON.parse(
    String(conteudo || "null")
  );

  let lista = [];

  if (Array.isArray(dados)) {
    lista = dados;
  } else if (
    dados &&
    typeof dados === "object"
  ) {
    const primeiraLista =
      Object.values(dados).find(
        Array.isArray
      );

    lista = primeiraLista || [dados];
  }

  const itensOriginais = lista.filter(
    (item) =>
      item &&
      typeof item === "object" &&
      !Array.isArray(item)
  );

  const conjuntoColunas = new Set();

  itensOriginais.forEach((item) => {
    Object.keys(item).forEach(
      (chave) =>
        conjuntoColunas.add(chave)
    );
  });

  const colunas =
    criarNomesUnicosUniversal(
      Array.from(conjuntoColunas)
    );

  const itens = itensOriginais.map(
    (item, indice) => {
      const novoItem = {
        __linha: indice + 1,
      };

      colunas.forEach((coluna) => {
        novoItem[coluna] =
          item[coluna] ?? "";
      });

      return novoItem;
    }
  );

  return {
    colunas,
    itens,
    delimitador: null,
    tipoLeitura: "json",
  };
}

function obterValorCelulaExcelUniversal(
  celula
) {
  const valor = celula?.value;

  if (valor === null || valor === undefined) {
    return "";
  }

  if (valor instanceof Date) {
    return valor.toISOString();
  }

  if (typeof valor !== "object") {
    return valor;
  }

  if (
    Object.prototype.hasOwnProperty.call(
      valor,
      "result"
    )
  ) {
    return valor.result ?? "";
  }

  if (
    Object.prototype.hasOwnProperty.call(
      valor,
      "text"
    )
  ) {
    return valor.text ?? "";
  }

  if (Array.isArray(valor.richText)) {
    return valor.richText
      .map((parte) => parte.text || "")
      .join("");
  }

  if (valor.hyperlink) {
    return valor.text || valor.hyperlink;
  }

  return String(valor);
}

async function interpretarExcelUniversal(
  buffer,
  possuiCabecalho = true
) {
  const workbook =
    new ExcelJS.Workbook();

  await workbook.xlsx.load(buffer);

  const planilha =
    workbook.worksheets[0];

  if (!planilha) {
    return {
      colunas: [],
      itens: [],
      planilha: null,
      tipoLeitura: "excel",
    };
  }

  const matriz = [];

  planilha.eachRow(
    {
      includeEmpty: false,
    },
    (linha) => {
      const valores = [];

      const totalCelulas = Math.max(
        linha.cellCount,
        linha.actualCellCount
      );

      for (
        let indice = 1;
        indice <= totalCelulas;
        indice += 1
      ) {
        valores.push(
          obterValorCelulaExcelUniversal(
            linha.getCell(indice)
          )
        );
      }

      matriz.push(valores);
    }
  );

  return {
    ...converterMatrizUniversal(
      matriz,
      possuiCabecalho
    ),

    planilha: planilha.name,
    tipoLeitura: "excel",
  };
}
async function interpretarArquivoUniversal(
  arquivo,
  opcoes = {}
) {
  if (!arquivo?.data) {
    throw new Error(
      "Arquivo inválido ou não recebido."
    );
  }

  const nomeArquivo = String(
    arquivo.name || "arquivo"
  );

  const extensao = path
    .extname(nomeArquivo)
    .toLowerCase();

  if (
    !FORMATOS_PREVIEW_UNIVERSAL.has(
      extensao
    )
  ) {
    throw new Error(
      "Formato não suportado. Use TXT, CSV, JSON ou XLSX."
    );
  }

  const possuiCabecalho =
    String(
      opcoes.possuiCabecalho ?? "true"
    ) !== "false";

  const delimitador = String(
    opcoes.delimitador || ""
  );

  const tipoLeitura = String(
    opcoes.tipoLeitura || "automatico"
  );

  let resultado = {
    colunas: [],
    itens: [],
    delimitador: null,
    planilha: null,
    tipoLeitura,
  };

  if (extensao === ".csv") {
    resultado =
      interpretarTextoDelimitadoUniversal(
        arquivo.data.toString("utf8"),
        {
          possuiCabecalho,
          delimitador,
        }
      );

    resultado.tipoLeitura =
      "delimitado";
  }

  if (extensao === ".txt") {
    const conteudo =
      arquivo.data.toString("utf8");

    if (
      tipoLeitura === "posicao-fixa"
    ) {
      resultado =
        interpretarTxtPosicaoFixaUniversal(
          conteudo
        );
    } else if (
      tipoLeitura === "delimitado"
    ) {
      resultado =
        interpretarTextoDelimitadoUniversal(
          conteudo,
          {
            possuiCabecalho,
            delimitador,
          }
        );

      resultado.tipoLeitura =
        "delimitado";
    } else {
      const delimitado =
        interpretarTextoDelimitadoUniversal(
          conteudo,
          {
            possuiCabecalho,
            delimitador,
          }
        );

      resultado =
        delimitado.semDelimitador
          ? interpretarTxtPosicaoFixaUniversal(
              conteudo
            )
          : {
              ...delimitado,
              tipoLeitura:
                "delimitado",
            };
    }
  }

  if (extensao === ".json") {
    resultado =
      interpretarJsonUniversal(
        arquivo.data.toString("utf8")
      );
  }

  if (extensao === ".xlsx") {
    resultado =
      await interpretarExcelUniversal(
        arquivo.data,
        possuiCabecalho
      );
  }

  return {
    nomeArquivo,
    extensao,
    formato:
      extensao.replace(".", ""),
    possuiCabecalho,
    ...resultado,
    colunas: Array.isArray(
      resultado?.colunas
    )
      ? resultado.colunas
      : [],
    itens: Array.isArray(
      resultado?.itens
    )
      ? resultado.itens
      : [],
  };
}
function normalizarTextoUniversal(valor) {
  return String(valor ?? "").trim();
}

function converterNumeroUniversal(valor) {
  if (
    typeof valor === "number" &&
    Number.isFinite(valor)
  ) {
    return valor;
  }

  let texto = String(valor ?? "")
    .trim();

  if (!texto) {
    return 0;
  }

  texto = texto
    .replace(/\s/g, "")
    .replace(/R\$/gi, "");

  const possuiVirgula =
    texto.includes(",");

  const possuiPonto =
    texto.includes(".");

  if (
    possuiVirgula &&
    possuiPonto
  ) {
    const ultimaVirgula =
      texto.lastIndexOf(",");

    const ultimoPonto =
      texto.lastIndexOf(".");

    if (ultimaVirgula > ultimoPonto) {
      texto = texto
        .replace(/\./g, "")
        .replace(",", ".");
    } else {
      texto = texto.replace(/,/g, "");
    }
  } else if (possuiVirgula) {
    texto = texto.replace(",", ".");
  }

  const numero = Number(texto);

  return Number.isFinite(numero)
    ? numero
    : NaN;
}

function aplicarMapeamentoUniversal(
  itemOrigem,
  mapeamento = {}
) {
  const resultado = {};

  Object.entries(mapeamento).forEach(
    ([campoDestino, colunaOrigem]) => {
      if (!colunaOrigem) {
        resultado[campoDestino] = "";
        return;
      }

      resultado[campoDestino] =
        itemOrigem?.[colunaOrigem] ?? "";
    }
  );

  return resultado;
}
/* ==========================================================
   TRANSFORMAÇÕES UNIVERSAIS
========================================================== */

const TIPOS_TRANSFORMACAO_UNIVERSAL =
  new Set([
    "trim",
    "maiusculas",
    "minusculas",
    "remover-acentos",
    "substituir",
    "remover-texto",
    "preencher-esquerda",
    "preencher-direita",
    "valor-fixo",
    "valor-padrao",
"numero",
"moeda",
"concatenar",
  ]);

function normalizarTransformacoesUniversais(
  transformacoes
) {
  if (
    !transformacoes ||
    typeof transformacoes !== "object" ||
    Array.isArray(transformacoes)
  ) {
    return {};
  }

  const resultado = {};

  Object.entries(transformacoes).forEach(
    ([campo, regras]) => {
      if (!Array.isArray(regras)) {
        return;
      }

      resultado[campo] = regras
        .filter(
          (regra) =>
            regra &&
            typeof regra === "object" &&
            TIPOS_TRANSFORMACAO_UNIVERSAL.has(
              String(regra.tipo || "")
            )
        )
        .map((regra) => ({
          tipo: String(
            regra.tipo || ""
          ),

          valor: String(
            regra.valor ?? ""
          ),

          procurar: String(
            regra.procurar ?? ""
          ),

          substituirPor: String(
            regra.substituirPor ?? ""
          ),

          tamanho:
            Math.max(
              0,
              Number.parseInt(
                regra.tamanho,
                10
              ) || 0
            ),

          caractere: String(
            regra.caractere ?? "0"
          ).slice(0, 1),

          separador: String(
            regra.separador ?? " "
          ),

          colunas:
            Array.isArray(
              regra.colunas
            )
              ? regra.colunas
                  .map((coluna) =>
                    String(
                      coluna || ""
                    ).trim()
                  )
                  .filter(Boolean)
              : [],
        }));
    }
  );

  return resultado;
}

function removerAcentosUniversal(valor) {
  return String(valor ?? "")
    .normalize("NFD")
    .replace(
      /[\u0300-\u036f]/g,
      ""
    );
}

function substituirTodasOcorrenciasUniversal(
  texto,
  procurar,
  substituirPor
) {
  if (!procurar) {
    return texto;
  }

  return String(texto).split(
    procurar
  ).join(substituirPor);
}

function aplicarRegraTransformacaoUniversal(
  valorAtual,
  regra,
  itemOrigem = {}
) {
  const tipo = String(
    regra?.tipo || ""
  );

  switch (tipo) {
    case "trim":
      return String(
        valorAtual ?? ""
      ).trim();

    case "maiusculas":
      return String(
        valorAtual ?? ""
      ).toUpperCase();

    case "minusculas":
      return String(
        valorAtual ?? ""
      ).toLowerCase();

    case "remover-acentos":
      return removerAcentosUniversal(
        valorAtual
      );

    case "substituir":
      return substituirTodasOcorrenciasUniversal(
        String(valorAtual ?? ""),
        String(
          regra.procurar ?? ""
        ),
        String(
          regra.substituirPor ?? ""
        )
      );

    case "remover-texto":
      return substituirTodasOcorrenciasUniversal(
        String(valorAtual ?? ""),
        String(
          regra.valor ?? ""
        ),
        ""
      );

    case "preencher-esquerda": {
      const tamanho =
        Math.max(
          0,
          Number.parseInt(
            regra.tamanho,
            10
          ) || 0
        );

      const caractere =
        String(
          regra.caractere ?? "0"
        ).slice(0, 1) || "0";

      return String(
        valorAtual ?? ""
      ).padStart(
        tamanho,
        caractere
      );
    }

    case "preencher-direita": {
      const tamanho =
        Math.max(
          0,
          Number.parseInt(
            regra.tamanho,
            10
          ) || 0
        );

      const caractere =
        String(
          regra.caractere ?? " "
        ).slice(0, 1) || " ";

      return String(
        valorAtual ?? ""
      ).padEnd(
        tamanho,
        caractere
      );
    }

    case "valor-fixo":
      return regra.valor ?? "";

    case "valor-padrao": {
      const vazio =
        valorAtual === null ||
        valorAtual === undefined ||
        String(valorAtual).trim() === "";

      return vazio
        ? regra.valor ?? ""
        : valorAtual;
    }

    case "numero": {
      /*
        converterNumeroUniversal mantém
        números negativos.

        Exemplos:
        -15,50 -> -15.5
        -8     -> -8
      */
      return converterNumeroUniversal(
        valorAtual
      );
    }
    case "moeda": {
      /*
        Converte diferentes representações
        monetárias para um número decimal.
    
        O banco deve armazenar número,
        não texto formatado com R$.
      */
      return converterNumeroUniversal(
        valorAtual
      );
    }
    case "concatenar": {
      const colunas =
        Array.isArray(
          regra.colunas
        )
          ? regra.colunas
          : [];

      const separador =
        String(
          regra.separador ?? " "
        );

      return colunas
        .map(
          (coluna) =>
            itemOrigem?.[coluna] ??
            ""
        )
        .map((parte) =>
          String(parte).trim()
        )
        .filter(
          (parte) => parte !== ""
        )
        .join(separador);
    }

    default:
      return valorAtual;
  }
}

function aplicarTransformacoesUniversais(
  itemMapeado,
  transformacoes = {},
  itemOrigem = {}
) {
  const resultado = {
    ...(itemMapeado || {}),
  };

  const regrasNormalizadas =
    normalizarTransformacoesUniversais(
      transformacoes
    );

  Object.entries(
    regrasNormalizadas
  ).forEach(([campo, regras]) => {
    let valorAtual =
      resultado[campo] ?? "";

    regras.forEach((regra) => {
      valorAtual =
        aplicarRegraTransformacaoUniversal(
          valorAtual,
          regra,
          itemOrigem
        );
    });

    resultado[campo] =
      valorAtual;
  });

  return resultado;
}
function validarItemUniversal(
  item,
  destino
) {
  const erros = [];

  if (destino === "base-principal") {
    const codigo =
      normalizarTextoUniversal(
        item.codigo
      );

    const descricao =
      normalizarTextoUniversal(
        item.descricao
      );

    const quantidade =
      converterNumeroUniversal(
        item.qtdeCongelada
      );

    if (!codigo) {
      erros.push(
        "Código interno não informado"
      );
    }

    if (!descricao) {
      erros.push(
        "Descrição não informada"
      );
    }

    if (!Number.isFinite(quantidade)) {
      erros.push(
        "Quantidade congelada inválida"
      );
    }
  }

  if (destino === "saldo-atual") {
    const codigo =
      normalizarTextoUniversal(
        item.codigo
      );

    const codigoBarras =
      normalizarTextoUniversal(
        item.codigoBarras
      );

    const quantidade =
      converterNumeroUniversal(
        item.qtdeCongelada
      );

    if (!codigo && !codigoBarras) {
      erros.push(
        "Código interno ou código de barras obrigatório"
      );
    }

    if (!Number.isFinite(quantidade)) {
      erros.push(
        "Saldo atual inválido"
      );
    }
  }

  if (destino === "complementar") {
    const codigo =
      normalizarTextoUniversal(
        item.codigo
      );

    const codigoBarras =
      normalizarTextoUniversal(
        item.codigoBarras
      );

    if (!codigo && !codigoBarras) {
      erros.push(
        "Código interno ou código de barras obrigatório"
      );
    }
  }

  if (destino === "base-wms") {
    const enderecoWms =
      normalizarTextoUniversal(
        item.enderecoWms
      );

    const codigo =
      normalizarTextoUniversal(
        item.codigo
      );

    const codigoBarras =
      normalizarTextoUniversal(
        item.codigoBarras
      );

    const quantidade =
      converterNumeroUniversal(
        item.quantidadeEsperada
      );

    if (!enderecoWms) {
      erros.push(
        "Endereço WMS não informado"
      );
    }

    if (!codigo && !codigoBarras) {
      erros.push(
        "Código interno ou código de barras obrigatório"
      );
    }

    if (!Number.isFinite(quantidade)) {
      erros.push(
        "Quantidade esperada inválida"
      );
    }
  }

  return erros;
}
const layoutTxtPadrao = {
  codigoBarras: { inicio: 0, fim: 13, tipo: "texto" },
  codigo: { inicio: 14, fim: 23, tipo: "texto" },
  descricao: { inicio: 24, fim: 75, tipo: "texto" },
  custoUnitario: { inicio: 75, fim: 83, tipo: "moeda" },
  qtdeCongelada: { inicio: 84, fim: 92, tipo: "quantidade" },
  categoria: { inicio: 93, fim: 126, tipo: "texto" },
  tipo: { inicio: 126, fim: 127, tipo: "texto" },
};

let layoutTxt = JSON.parse(JSON.stringify(layoutTxtPadrao));

app.use(
  session({
    secret: "inventario2025",
    resave: false,
    saveUninitialized: false,
  })
);

app.use(fileUpload());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, "public"), { index: false }));
app.use(express.static(__dirname, { index: false }));
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

function broadcastInventario() {
  const dados = JSON.stringify({ type: "inventario", inventario });
  wss.clients.forEach((client) => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(dados);
    }
  });
}

function autenticar(req, res, next) {
  if (req.session?.logado) {
    return next();
  }

  const aceitaJson =
    req.xhr ||
    req.headers.accept?.includes(
      "application/json"
    ) ||
    req.headers["x-requested-with"] ===
      "XMLHttpRequest" ||
    req.originalUrl?.startsWith(
      "/configurador-universal/"
    ) ||
    req.headers["content-type"]?.includes(
      "application/json"
    );

  if (aceitaJson) {
    return res.status(401).json({
      sucesso: false,
      erro:
        "Sessão expirada. Faça login novamente.",
    });
  }

  return res.redirect(
    `/login?redirect=${encodeURIComponent(
      req.originalUrl || "/"
    )}`
  );
}
function permitirSomenteLiderOuAdmin(req, res, next) {
  const aceitaJson =
    req.xhr ||
    req.headers.accept?.includes("application/json") ||
    req.headers["content-type"]?.includes("application/json");

  if (!req.session?.logado || !req.session?.usuario) {
    if (aceitaJson) {
      return res.status(401).json({
        erro: "Sessão expirada. Faça login novamente.",
      });
    }

    return res.redirect("/login");
  }

  const funcao = String(req.session.usuario.funcao || "")
    .trim()
    .toLowerCase();

  const permitido =
    funcao === "líder" ||
    funcao === "lider" ||
    funcao === "administrador";

  if (permitido) {
    return next();
  }

  if (aceitaJson) {
    return res.status(403).json({
      erro: "Usuário sem permissão para acessar esta informação.",
    });
  }

  return res.redirect("/coleta-mobile");
}

function permitirSomenteOperador(req, res, next) {
  if (!req.session?.logado || !req.session?.usuario) {
    return res.redirect('/login?redirect=/coleta-mobile');
  }

  const funcao = String(req.session.usuario.funcao || '').toLowerCase();

  if (funcao === 'operador') {
    return next();
  }

  return res.redirect('/');
}
/*
  ============================================================
  CONFIGURADOR UNIVERSAL — ROTAS PRIORITÁRIAS
  Estas rotas ficam antes das rotas de páginas e fallbacks.
  ============================================================
*/

app.get(
  "/configurador-universal/status",
  autenticar,
  (req, res) => {
    return res.json({
      sucesso: true,
      modulo: "configurador-universal",
      previewDisponivel: true,
      versao: "1.0.0",
    });
  }
);

app.post(
  "/configurador-universal/preview",
  autenticar,
  async (req, res) => {
    try {
      const arquivo = req.files?.arquivo;

      if (!arquivo) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Selecione um arquivo para gerar o preview.",
        });
      }

      const nomeArquivo = String(
        arquivo.name || "arquivo"
      );

      const extensao = path
        .extname(nomeArquivo)
        .toLowerCase();

      if (
        !FORMATOS_PREVIEW_UNIVERSAL.has(extensao)
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Formato não suportado. Use TXT, CSV, JSON ou XLSX.",
        });
      }

      const possuiCabecalho =
        String(
          req.body?.possuiCabecalho ?? "true"
        ) !== "false";

      const delimitador = String(
        req.body?.delimitador || ""
      );

      const tipoLeituraSolicitado = String(
        req.body?.tipoLeitura || "automatico"
      );

      const limitePreview = Math.min(
        100,
        Math.max(
          1,
          Number(req.body?.limitePreview) || 30
        )
      );

      let resultado = {
        colunas: [],
        itens: [],
        delimitador: null,
        planilha: null,
        tipoLeitura: tipoLeituraSolicitado,
      };

      if (extensao === ".csv") {
        resultado =
          interpretarTextoDelimitadoUniversal(
            arquivo.data.toString("utf8"),
            {
              possuiCabecalho,
              delimitador,
            }
          );

        resultado.tipoLeitura = "delimitado";
      }

      if (extensao === ".txt") {
        const conteudo =
          arquivo.data.toString("utf8");

        if (
          tipoLeituraSolicitado ===
          "posicao-fixa"
        ) {
          resultado =
            interpretarTxtPosicaoFixaUniversal(
              conteudo
            );
        } else if (
          tipoLeituraSolicitado ===
          "delimitado"
        ) {
          resultado =
            interpretarTextoDelimitadoUniversal(
              conteudo,
              {
                possuiCabecalho,
                delimitador,
              }
            );

          resultado.tipoLeitura = "delimitado";
        } else {
          const delimitado =
            interpretarTextoDelimitadoUniversal(
              conteudo,
              {
                possuiCabecalho,
                delimitador,
              }
            );

          if (delimitado.semDelimitador) {
            resultado =
              interpretarTxtPosicaoFixaUniversal(
                conteudo
              );
          } else {
            resultado = {
              ...delimitado,
              tipoLeitura: "delimitado",
            };
          }
        }
      }

      if (extensao === ".json") {
        resultado = interpretarJsonUniversal(
          arquivo.data.toString("utf8")
        );
      }

      if (extensao === ".xlsx") {
        resultado =
          await interpretarExcelUniversal(
            arquivo.data,
            possuiCabecalho
          );
      }

      const itens = Array.isArray(
        resultado?.itens
      )
        ? resultado.itens
        : [];

      const colunas = Array.isArray(
        resultado?.colunas
      )
        ? resultado.colunas
        : [];

      return res.status(200).json({
        sucesso: true,

        arquivo: {
          nome: nomeArquivo,
          extensao,
          tamanho:
            Number(arquivo.size) ||
            Number(arquivo.data?.length) ||
            0,
        },

        formato: extensao.slice(1),

        tipoLeitura:
          resultado.tipoLeitura ||
          tipoLeituraSolicitado,

        possuiCabecalho,

        delimitador:
          resultado.delimitador ?? null,

        planilha:
          resultado.planilha || null,

        totalLinhas: itens.length,
        totalColunas: colunas.length,

        colunas,

        preview: itens.slice(
          0,
          limitePreview
        ),
      });
    } catch (erro) {
      console.error(
        "Erro no preview universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          erro?.message ||
          "Não foi possível interpretar o arquivo.",
      });
    }
  }
);
app.post(
  "/configurador-universal/validar",
  autenticar,
  async (req, res) => {
    try {
      const arquivo =
        req.files?.arquivo;

      if (!arquivo) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Selecione um arquivo para validar.",
        });
      }

      const destino = String(
        req.body?.destino || ""
      ).trim();

      const destinosPermitidos =
        new Set([
          "base-principal",
          "saldo-atual",
          "complementar",
          "base-wms",
        ]);

      if (
        !destinosPermitidos.has(
          destino
        )
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Selecione um destino válido.",
        });
      }

      let mapeamento = {};

      try {
        mapeamento = JSON.parse(
          String(
            req.body?.mapeamento ||
            "{}"
          )
        );
      } catch (erro) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O mapeamento informado é inválido.",
        });
      }

      if (
        !mapeamento ||
        typeof mapeamento !==
          "object" ||
        Array.isArray(mapeamento)
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O mapeamento informado é inválido.",
        });
      }
      let transformacoes = {};

      try {
        transformacoes =
          normalizarTransformacoesUniversais(
            JSON.parse(
              String(
                req.body?.transformacoes ||
                "{}"
              )
            )
          );
      } catch (erro) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "As transformações informadas são inválidas.",
        });
      }
      const leitura =
        await interpretarArquivoUniversal(
          arquivo,
          {
            possuiCabecalho:
              req.body?.possuiCabecalho,

            delimitador:
              req.body?.delimitador,

            tipoLeitura:
              req.body?.tipoLeitura,
          }
        );

      const validos = [];
      const invalidos = [];

      leitura.itens.forEach(
        (itemOrigem, indice) => {
          const itemMapeado =
  aplicarMapeamentoUniversal(
    itemOrigem,
    mapeamento
  );

const itemTransformado =
  aplicarTransformacoesUniversais(
    itemMapeado,
    transformacoes,
    itemOrigem
  );

const erros =
  validarItemUniversal(
    itemTransformado,
    destino
  );

          const registro = {
            linha:
              Number(
                itemOrigem?.__linha
              ) ||
              indice + 1,

            origem: itemOrigem,
            mapeado: itemMapeado,
transformado: itemTransformado,
erros,
          };

          if (erros.length) {
            invalidos.push(registro);
          } else {
            validos.push(registro);
          }
        }
      );

      return res.json({
        sucesso: true,

        arquivo: {
          nome:
            leitura.nomeArquivo,
          formato:
            leitura.formato,
        },

        destino,

        resumo: {
          total:
            leitura.itens.length,

          validos:
            validos.length,

          invalidos:
            invalidos.length,

          percentualValido:
            leitura.itens.length > 0
              ? Number(
                  (
                    validos.length /
                    leitura.itens.length *
                    100
                  ).toFixed(2)
                )
              : 0,
        },

        amostraValidos:
          validos.slice(0, 20),

        amostraInvalidos:
          invalidos.slice(0, 50),
      });
    } catch (erro) {
      console.error(
        "Erro na validação universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          erro?.message ||
          "Não foi possível validar o arquivo.",
      });
    }
  }
);
/* ==========================================================
   PROCESSAMENTO UNIVERSAL
   TODOS OS DESTINOS
========================================================== */

app.post(
  "/configurador-universal/processar",
  autenticar,
  permitirSomenteLiderOuAdmin,
  async (req, res) => {
    try {
      const arquivo =
        req.files?.arquivo;

      if (!arquivo) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Selecione um arquivo para processar.",
        });
      }

      const destino = String(
        req.body?.destino || ""
      ).trim();

      const destinosPermitidos = [
        "base-principal",
        "saldo-atual",
        "complementar",
        "base-wms",
      ];

      if (
        !destinosPermitidos.includes(
          destino
        )
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O destino informado é inválido.",
        });
      }

      let mapeamento = {};

      try {
        mapeamento = JSON.parse(
          String(
            req.body?.mapeamento ||
            "{}"
          )
        );
      } catch (erro) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O mapeamento informado é inválido.",
        });
      }

      if (
        !mapeamento ||
        typeof mapeamento !==
          "object" ||
        Array.isArray(mapeamento)
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O mapeamento informado é inválido.",
        });
      }

      let transformacoes = {};

      try {
        transformacoes =
          normalizarTransformacoesUniversais(
            JSON.parse(
              String(
                req.body
                  ?.transformacoes ||
                "{}"
              )
            )
          );
      } catch (erro) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "As transformações informadas são inválidas.",
        });
      }

      const leitura =
        await interpretarArquivoUniversal(
          arquivo,
          {
            possuiCabecalho:
              req.body
                ?.possuiCabecalho,

            delimitador:
              req.body
                ?.delimitador,

            tipoLeitura:
              req.body
                ?.tipoLeitura,
          }
        );

      if (
        !Array.isArray(
          leitura.itens
        ) ||
        !leitura.itens.length
      ) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "O arquivo não possui registros para processar.",
        });
      }

      const registrosValidos = [];
      const registrosInvalidos = [];

      leitura.itens.forEach(
        (itemOrigem, indice) => {
          const itemMapeado =
            aplicarMapeamentoUniversal(
              itemOrigem,
              mapeamento
            );

          const itemTransformado =
            aplicarTransformacoesUniversais(
              itemMapeado,
              transformacoes,
              itemOrigem
            );

          const erros =
            validarItemUniversal(
              itemTransformado,
              destino
            );

          const linha =
            Number(
              itemOrigem?.__linha
            ) ||
            indice + 1;

          if (erros.length) {
            registrosInvalidos.push({
              linha,
              erros,
            });

            return;
          }

          registrosValidos.push(
            itemTransformado
          );
        }
      );

      /*
        Nenhum destino é gravado parcialmente.
        Havendo linha inválida, a importação
        inteira é bloqueada.
      */
      if (
        registrosInvalidos.length
      ) {
        return res.status(422).json({
          sucesso: false,

          erro:
            "A importação foi bloqueada porque existem registros inválidos.",

          resumo: {
            total:
              leitura.itens.length,

            validos:
              registrosValidos.length,

            invalidos:
              registrosInvalidos.length,
          },

          amostraInvalidos:
            registrosInvalidos.slice(
              0,
              50
            ),
        });
      }

      /*
        Normaliza os campos do inventário.
        converterNumeroUniversal preserva
        corretamente os números negativos.
      */
      const normalizarItemInventario =
        (item) => {
          const codigo =
            normalizarTextoUniversal(
              item?.codigo
            );

          const codigoBarras =
            normalizarTextoUniversal(
              item?.codigoBarras
            );

          const quantidade =
            converterNumeroUniversal(
              item?.qtdeCongelada
            );

          const custo =
            converterNumeroUniversal(
              item?.custoUnitario
            );

          return {
            codigoBarras,

            codigo,

            codigoInterno:
              codigo,

            descricao:
              normalizarTextoUniversal(
                item?.descricao
              ),

            categoria:
              normalizarTextoUniversal(
                item?.categoria
              ),

            tipo:
              normalizarTextoUniversal(
                item?.tipo
              ),

            custoUnitario:
              Number.isFinite(custo)
                ? custo
                : 0,

            qtdeCongelada:
              Number.isFinite(
                quantidade
              )
                ? quantidade
                : 0,

            qtdeContada: 0,
          };
        };

      /* ====================================================
         BASE PRINCIPAL
         Mantém o comportamento da rota /importar-txt:
         inicia um novo ciclo de inventário.
      ==================================================== */

      if (
        destino ===
        "base-principal"
      ) {
        const itensImportados =
          registrosValidos.map(
            normalizarItemInventario
          );

        const totalImportadoBruto =
          itensImportados.length;

        const itensUnicosImportados =
          removerDuplicadosPorCodigoInterno(
            itensImportados
          );

        const totalUnicosBruto =
          itensUnicosImportados.length;

        const duplicatasRemovidas =
          totalImportadoBruto -
          totalUnicosBruto;

        const itensZeradosIgnorados =
          itensUnicosImportados.filter(
            (item) =>
              parseQuantidade(
                item.qtdeCongelada
              ) <= 0
          ).length;

        auditoriaImportacao = {
          totalImportadoBruto,
          totalUnicosBruto,
          duplicatasRemovidas,
          itensZeradosIgnorados,
        };

        /*
          Regra existente da Base principal:
          reinicia o ciclo trabalhado.
        */
        limparContagensPersistidas();

        historicoAlteracoes = [];
        historicoAuditoriaItens = [];
        limparEventosDoModoNosEnderecamentos(
          "com-base"
        );
        itemAuditoriaAtual = null;

        inventario =
          itensUnicosImportados;

        await salvarProdutosNoBanco(
          inventario
        );

        modoOperacao =
          "com-base";

        await salvarModoOperacao();

        tipoUltimaImportacao =
          "Importação base";

        ultimaImportacao = {
          arquivo:
            leitura.nomeArquivo,

          tipo:
            "Importação base",

          horario:
            new Date().toISOString(),

          status:
            "Sucesso na importação",

          observacao:
            `Base principal importada com sucesso. ${inventario.length} itens carregados.`,
        };

        await salvarEnderecamentos();

        broadcastInventario();

        return res.json({
          sucesso: true,

          mensagem:
            "Arquivo base importado com sucesso.",

          destino,

          arquivo: {
            nome:
              leitura.nomeArquivo,

            formato:
              leitura.formato,
          },

          resultado: {
            registrosRecebidos:
              totalImportadoBruto,

            registrosSalvos:
              inventario.length,

            registrosIgnorados:
              duplicatasRemovidas,

            substituirBase: true,
          },

          modoOperacao,
          ultimaImportacao,
        });
      }

      if (
        destino ===
        "saldo-atual"
      ) {
        const totalImportadoBruto =
          registrosValidos.length;
      
        /*
          Normaliza códigos somente para localizar
          o produto antigo.
      
          Exemplos equivalentes para a busca:
          000038186
          38186
      
          O valor gravado continuará sendo exatamente
          o resultado da transformação configurada.
        */
        const normalizarChaveCodigoSaldo =
          (valor) => {
            const texto =
              normalizarTextoUniversal(
                valor
              );
      
            if (!texto) {
              return "";
            }
      
            if (/^\d+$/.test(texto)) {
              return (
                texto.replace(
                  /^0+(?=\d)/,
                  ""
                ) || "0"
              );
            }
      
            return texto.toUpperCase();
          };
      
        const normalizarChaveBarrasSaldo =
          (valor) =>
            normalizarTextoUniversal(
              valor
            );
      
        /*
          Índices da base atual utilizados apenas
          para recuperar descrição, custo, categoria,
          tipo, código de barras e outros dados
          quando eles não vierem no arquivo de saldo.
        */
        const inventarioPorCodigo =
          new Map();
      
        const inventarioPorCodigoBarras =
          new Map();
      
        inventario.forEach(
          (produtoAtual) => {
            const chaveCodigo =
              normalizarChaveCodigoSaldo(
                produtoAtual?.codigo ||
                produtoAtual?.codigoInterno
              );
      
            const chaveBarras =
              normalizarChaveBarrasSaldo(
                produtoAtual?.codigoBarras
              );
      
            if (chaveCodigo) {
              inventarioPorCodigo.set(
                chaveCodigo,
                produtoAtual
              );
            }
      
            if (chaveBarras) {
              inventarioPorCodigoBarras.set(
                chaveBarras,
                produtoAtual
              );
            }
          }
        );
      
        const itensAtualizados =
          registrosValidos.map(
            (itemImportado) => {
              const codigoImportado =
                normalizarTextoUniversal(
                  itemImportado?.codigo
                );
      
              const codigoBarrasImportado =
                normalizarTextoUniversal(
                  itemImportado?.codigoBarras
                );
      
              const chaveCodigoImportado =
                normalizarChaveCodigoSaldo(
                  codigoImportado
                );
      
              const chaveBarrasImportado =
                normalizarChaveBarrasSaldo(
                  codigoBarrasImportado
                );
      
              /*
                Prioriza o código de barras porque
                ele não é afetado pela retirada dos
                zeros à esquerda do código interno.
              */
              const produtoAtual =
                (
                  chaveBarrasImportado
                    ? inventarioPorCodigoBarras.get(
                        chaveBarrasImportado
                      )
                    : null
                ) ||
                (
                  chaveCodigoImportado
                    ? inventarioPorCodigo.get(
                        chaveCodigoImportado
                      )
                    : null
                ) ||
                null;
      
              const campoFoiInformado =
                (campo) =>
                  Object.prototype
                    .hasOwnProperty.call(
                      itemImportado,
                      campo
                    ) &&
                  itemImportado[campo] !==
                    null &&
                  itemImportado[campo] !==
                    undefined &&
                  String(
                    itemImportado[campo]
                  ).trim() !== "";
      
              const descricaoImportada =
                normalizarTextoUniversal(
                  itemImportado?.descricao
                );
      
              const categoriaImportada =
                normalizarTextoUniversal(
                  itemImportado?.categoria
                );
      
              const tipoImportado =
                normalizarTextoUniversal(
                  itemImportado?.tipo
                );
      
              const custoImportado =
                converterNumeroUniversal(
                  itemImportado?.custoUnitario
                );
      
              const quantidadeImportada =
                converterNumeroUniversal(
                  itemImportado?.qtdeCongelada
                );
      
              /*
                O código transformado é o que será
                efetivamente gravado.
      
                Portanto:
                000038186 transformado em número
                será salvo como 38186.
              */
              const codigoFinal =
                codigoImportado ||
                normalizarTextoUniversal(
                  produtoAtual?.codigo ||
                  produtoAtual?.codigoInterno
                );
      
              const codigoBarrasFinal =
                codigoBarrasImportado ||
                normalizarTextoUniversal(
                  produtoAtual?.codigoBarras
                );
      
              /*
                Mantém as contagens já realizadas.
      
                Nenhum registro de contagem é apagado.
              */
              const totalContado =
                contagens
                  .filter(
                    (contagem) => {
                      if (
                        !contagem ||
                        contagem.ativo === false
                      ) {
                        return false;
                      }
      
                      const barrasContagem =
                        normalizarTextoUniversal(
                          contagem.codigoBarras
                        );
      
                      return (
                        barrasContagem !== "" &&
                        barrasContagem ===
                          codigoBarrasFinal
                      );
                    }
                  )
                  .reduce(
                    (
                      total,
                      contagem
                    ) =>
                      total +
                      (
                        Number(
                          contagem.quantidade
                        ) || 0
                      ),
                    0
                  );
      
              return {
                codigoBarras:
                  codigoBarrasFinal,
      
                codigo:
                  codigoFinal,
      
                codigoInterno:
                  codigoFinal,
      
                descricao:
                  campoFoiInformado(
                    "descricao"
                  )
                    ? descricaoImportada
                    : normalizarTextoUniversal(
                        produtoAtual?.descricao
                      ),
      
                custoUnitario:
                  campoFoiInformado(
                    "custoUnitario"
                  ) &&
                  Number.isFinite(
                    custoImportado
                  )
                    ? custoImportado
                    : (
                        Number(
                          produtoAtual
                            ?.custoUnitario
                        ) || 0
                      ),
      
                qtdeCongelada:
                  Number.isFinite(
                    quantidadeImportada
                  )
                    ? quantidadeImportada
                    : 0,
      
                categoria:
                  campoFoiInformado(
                    "categoria"
                  )
                    ? categoriaImportada
                    : normalizarTextoUniversal(
                        produtoAtual?.categoria
                      ),
      
                tipo:
                  campoFoiInformado(
                    "tipo"
                  )
                    ? tipoImportado
                    : normalizarTextoUniversal(
                        produtoAtual?.tipo
                      ),
      
                qtdeContada:
                  totalContado,
              };
            }
          );
      
        /*
          O Saldo Atual substitui a lista de produtos
          pelo conteúdo do novo arquivo, igual à rota
          TXT original.
      
          Ele não mantém os produtos ausentes no
          arquivo novo.
        */
        const itensUnicosImportados =
          removerDuplicadosPorCodigoInterno(
            itensAtualizados
          );
      
        const totalUnicosBruto =
          itensUnicosImportados.length;
      
        const duplicatasRemovidas =
          totalImportadoBruto -
          totalUnicosBruto;
      
        const itensZeradosIgnorados =
          itensUnicosImportados.filter(
            (item) =>
              parseQuantidade(
                item.qtdeCongelada
              ) <= 0
          ).length;
      
        auditoriaImportacao = {
          totalImportadoBruto,
          totalUnicosBruto,
          duplicatasRemovidas,
          itensZeradosIgnorados,
        };
      
        /*
          Aqui está a correção principal.
      
          O inventário passa a conter somente os
          produtos existentes no arquivo de Saldo
          Atual recém-importado.
        */
        inventario =
          itensUnicosImportados;
      
        /*
          salvarProdutosNoBanco substitui apenas a
          tabela/lista de produtos.
      
          Não são apagados:
          - contagens;
          - endereçamentos;
          - auditorias;
          - históricos;
          - usuários;
          - WMS;
          - finalizações.
        */
        await salvarProdutosNoBanco(
          inventario
        );
      
        tipoUltimaImportacao =
          "Atualização de saldo";
      
        ultimaImportacao = {
          arquivo:
            leitura.nomeArquivo,
      
          tipo:
            "Atualização de saldo",
      
          horario:
            new Date().toISOString(),
      
          status:
            "Sucesso na importação",
      
          observacao:
            `Saldo atualizado com sucesso. ${inventario.length} itens carregados. Contagens, usuários, endereços, auditorias e históricos foram preservados.`,
        };
      
        broadcastInventario();
      
        return res.json({
          sucesso: true,
      
          mensagem:
            "Saldo atual importado com sucesso.",
      
          destino,
      
          arquivo: {
            nome:
              leitura.nomeArquivo,
      
            formato:
              leitura.formato,
          },
      
          resultado: {
            registrosRecebidos:
              totalImportadoBruto,
      
            registrosSalvos:
              inventario.length,
      
            registrosIgnorados:
              duplicatasRemovidas,
      
            substituirBase: true,
          },
      
          ultimaImportacao,
        });
      }
      /* ====================================================
         ARQUIVO COMPLEMENTAR
         Acrescenta os registros sem limpar o inventário
         nem os dados já trabalhados.
      ==================================================== */

      if (
        destino ===
        "complementar"
      ) {
        const itensComplementares =
          registrosValidos.map(
            (item) => {
              const normalizado =
                normalizarItemInventario(
                  item
                );

              const quantidadeContada =
                converterNumeroUniversal(
                  item?.qtdeContada
                );

              return {
                ...normalizado,

                qtdeContada:
                  Number.isFinite(
                    quantidadeContada
                  )
                    ? quantidadeContada
                    : 0,
              };
            }
          );

        inventario = [
          ...inventario,
          ...itensComplementares,
        ];

        await salvarProdutosNoBanco(
          inventario
        );

        tipoUltimaImportacao =
          "Arquivo complementar";

        ultimaImportacao = {
          arquivo:
            leitura.nomeArquivo,

          tipo:
            "Arquivo complementar",

          horario:
            new Date().toISOString(),

          status:
            "Sucesso na importação",

          observacao:
            `${itensComplementares.length} registros complementares adicionados. Os dados já trabalhados foram preservados.`,
        };

        broadcastInventario();

        return res.json({
          sucesso: true,

          mensagem:
            "Arquivo complementar importado com sucesso.",

          destino,

          arquivo: {
            nome:
              leitura.nomeArquivo,

            formato:
              leitura.formato,
          },

          resultado: {
            registrosRecebidos:
              registrosValidos.length,

            registrosSalvos:
              itensComplementares.length,

            registrosIgnorados: 0,

            substituirBase: false,
          },

          ultimaImportacao,
        });
      }

      /* ====================================================
         BASE ESPERADA WMS
         Mantém o processamento universal já existente.
      ==================================================== */

      if (
        destino ===
        "base-wms"
      ) {
        if (!usarPostgres) {
          return res.status(503).json({
            sucesso: false,

            erro:
              "O PostgreSQL precisa estar ativo para importar a Base Esperada WMS.",
          });
        }

        const itensWms =
          registrosValidos.map(
            (item) => ({
              enderecoWms:
                normalizarTextoUniversal(
                  item?.enderecoWms
                ),

              codigoBarras:
                normalizarTextoUniversal(
                  item?.codigoBarras
                ),

              codigo:
                normalizarTextoUniversal(
                  item?.codigo
                ),

              descricao:
                normalizarTextoUniversal(
                  item?.descricao
                ),

              quantidadeEsperada:
                converterNumeroUniversal(
                  item
                    ?.quantidadeEsperada
                ),
            })
          );

        const substituirBase =
          String(
            req.body
              ?.substituirBase ||
            "false"
          ) === "true";

        const resultado =
          await salvarBaseEsperadaWmsLotePostgres(
            itensWms,
            substituirBase
          );

        return res.json({
          sucesso: true,

          mensagem:
            substituirBase
              ? "Base Esperada WMS substituída com sucesso."
              : "Base Esperada WMS atualizada com sucesso.",

          destino,

          arquivo: {
            nome:
              leitura.nomeArquivo,

            formato:
              leitura.formato,
          },

          resultado: {
            registrosRecebidos:
              Number(
                resultado
                  ?.registrosRecebidos
              ) ||
              itensWms.length,

            registrosSalvos:
              Number(
                resultado
                  ?.registrosSalvos
              ) || 0,

            registrosIgnorados:
              Number(
                resultado
                  ?.registrosIgnorados
              ) || 0,

            substituirBase,
          },
        });
      }

      return res.status(400).json({
        sucesso: false,
        erro:
          "Não foi possível determinar o processamento do destino.",
      });
    } catch (erro) {
      console.error(
        "Erro no processamento universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,

        erro:
          erro?.message ||
          "Não foi possível processar o arquivo.",
      });
    }
  }
);
function registrarAlteracao(usuario, codigoBarras, campo, valorAntigo, valorNovo) {
  historicoAlteracoes.push({
    usuario,
    codigoBarras,
    campo,
    valorAntigo,
    valorNovo,
    data: new Date().toISOString(),
  });
}
function garantirPastaEncerramentos() {
  garantirPastaData();

  if (!fs.existsSync(encerramentosDir)) {
    fs.mkdirSync(encerramentosDir, { recursive: true });
  }
}

function gerarTimestampEncerramento() {
  const agora = new Date();
  const yyyy = agora.getFullYear();
  const mm = String(agora.getMonth() + 1).padStart(2, "0");
  const dd = String(agora.getDate()).padStart(2, "0");
  const hh = String(agora.getHours()).padStart(2, "0");
  const mi = String(agora.getMinutes()).padStart(2, "0");
  const ss = String(agora.getSeconds()).padStart(2, "0");

  return `${yyyy}-${mm}-${dd}_${hh}-${mi}-${ss}`;
}

function gerarCsvInventario(lista) {
  const cabecalho = [
    "codigoBarras",
    "codigo",
    "descricao",
    "categoria",
    "custoUnitario",
    "qtdeCongelada",
    "qtdeContada",
    "divergencia"
  ];

  const linhas = lista.map((item) => {
    const qtdeCongelada = Number(item.qtdeCongelada) || 0;
    const qtdeContada = Number(item.qtdeContada) || 0;
    const divergencia = qtdeContada - qtdeCongelada;

    return [
      item.codigoBarras ?? "",
      item.codigo ?? "",
      `"${String(item.descricao ?? "").replace(/"/g, '""')}"`,
      `"${String(item.categoria ?? "").replace(/"/g, '""')}"`,
      Number(item.custoUnitario) || 0,
      qtdeCongelada,
      qtdeContada,
      divergencia
    ].join(";");
  });

  return [cabecalho.join(";"), ...linhas].join("\n");
}
function montarSnapshotEncerramento(usuario) {
  return {
    encerradoEm: new Date().toISOString(),
    encerradoPor: usuario || "desconhecido",
    resumo: {
      totalItens: inventario.length,
      totalAlteracoes: historicoAlteracoes.length,
      totalAuditorias: historicoAuditoriaItens.length,
      totalContagens: contagens.length,
      totalEnderecamentos: enderecamentos.length
    },
    auditoriaImportacao,
    inventario,
    historicoAlteracoes,
    historicoAuditoriaItens,
    contagens,
    enderecamentos
  };
}

async function resetarSistemaAposEncerramento(
  modoReferencia = modoOperacao
) {
  const modo = normalizarModoOperacao(
    modoReferencia
  );

  /*
    Primeiro removemos os eventos de endereçamento
    pertencentes ao modo que está sendo encerrado.
  */
  limparEventosDoModoNosEnderecamentos(
    modo
  );

  /*
  ==========================================================
  MODO SEM BASE
  ==========================================================

  Encerrar o inventário significa encerrar
  todo o ciclo desse modo:

  - contagens
  - finalizações
  - transmissões
  - cadastros de endereços

  Endereços pertencentes aos outros modos
  permanecem intactos.
*/
if (modo === "sem-base") {

  /*
    1. Limpa dados operacionais
    específicos do sem base.
  */
  contagemSemBase = [];

  finalizacoesSemBase = [];


  /*
    2. Remove os CADASTROS DE ENDEREÇO
    pertencentes ao sem base.

    Exemplo:
    1, 2, 3, 4 e 5 deixam de existir
    depois do encerramento.
  */
  enderecamentos = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).filter(
    (endereco) =>
      normalizarModoOperacao(
        endereco?.modoOperacao
      ) !== "sem-base"
  );


  /*
    3. Persiste as limpezas.
  */
  await salvarContagemSemBase();

  await salvarFinalizacoesSemBase();


  /*
    permitirVazio = true é proposital.

    Se não existir nenhum endereço de
    outro modo, enderecamentos ficará [].

    Nesse caso o PostgreSQL DEVE ser
    realmente zerado.
  */
  await salvarEnderecamentos({
    permitirVazio: true,
  });


  console.log(
    "✅ Inventário SEM BASE encerrado: " +
    "contagens, finalizações e endereços removidos."
  );


  return;
}
  /*
    MODO COM BASE
  */
  if (modo === "com-base") {
    inventario = [];
    historicoAlteracoes = [];
    historicoAuditoriaItens = [];
    contagens = [];
    itemAuditoriaAtual = null;

    auditoriaImportacao = {
      totalImportadoBruto: 0,
      totalUnicosBruto: 0,
      duplicatasRemovidas: 0,
      itensZeradosIgnorados: 0,
    };

    await salvarProdutosNoBanco([]);
    await salvarContagens();
    await salvarEnderecamentos();

    return;
  }

  /*
    WMS:
    neste ponto não apagamos dados dos outros modos.
  */
  await salvarEnderecamentos();
}


function garantirPastaData() {
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }
}

async function carregarContagens() {
  try {
    garantirPastaData();

    if (usarPostgres) {
      const lidasBanco = await carregarContagensPostgres();

      if (Array.isArray(lidasBanco) && lidasBanco.length > 0) {
        contagens = lidasBanco;

        fs.writeFileSync(
          contagensPath,
          JSON.stringify(contagens, null, 2),
          "utf8"
        );

        console.log(`✅ Contagens carregadas do PostgreSQL: ${contagens.length}`);
        return;
      }
    }

    if (!fs.existsSync(contagensPath)) {
      fs.writeFileSync(contagensPath, JSON.stringify([], null, 2), "utf8");
      contagens = [];

      if (usarPostgres) {
        await salvarContagensPostgres(contagens);
      }

      return;
    }

    const conteudo = fs.readFileSync(contagensPath, "utf8");
    contagens = JSON.parse(conteudo || "[]");

    if (usarPostgres) {
      await salvarContagensPostgres(contagens);
      console.log(`✅ Contagens migradas do JSON para PostgreSQL: ${contagens.length}`);
    }
  } catch (erro) {
    console.error("Erro ao carregar contagens:", erro.message);
    contagens = [];
  }
}

async function salvarContagens() {
  try {
    garantirPastaData();

    fs.writeFileSync(
      contagensPath,
      JSON.stringify(Array.isArray(contagens) ? contagens : [], null, 2),
      "utf8"
    );

    if (usarPostgres) {
      await salvarContagensPostgres(contagens);
    }

    console.log("✅ Contagens salvas no JSON e PostgreSQL.");
  } catch (erro) {
    console.error("Erro ao salvar contagens:", erro.message);
  }
}
async function limparContagensPersistidas() {
  contagens = [];
  await salvarContagens();
}

function recalcularInventarioComBaseNasContagens() {
  inventario = inventario.map((produto) => {
    const codigoBarras = String(produto.codigoBarras || '').trim();

    const totalContado = contagens
      .filter((c) => {
        return (
          c &&
          c.ativo !== false &&                    // ✅ OK
          String(c.codigoBarras || '').trim() === codigoBarras
        );
      })
      .reduce((acc, c) => acc + (Number(c.quantidade) || 0), 0);

    return {
      ...produto,
      qtdeContada: totalContado,
    };
  });
}
function calcularTotalPosicoesEndereco(inicio, fim) {
  const nInicio = Number(inicio) || 0;
  const nFim = Number(fim) || 0;
  if (nFim < nInicio) return 0;
  return nFim - nInicio + 1;
}
function normalizarTextoEndereco(valor) {
  return String(valor || "").trim();
}
function normalizarEnderecoSalvo(item) {
  const transmissoes = Array.isArray(item?.transmissoes) ? item.transmissoes : [];
  const finalizacoes = Array.isArray(item?.finalizacoes) ? item.finalizacoes : [];

  const totalTransmissoes = transmissoes.filter((e) => e.tipo === "transmissao").length;
  const totalFinalizacoesEventos = transmissoes.filter((e) => e.tipo === "finalizacao").length;
  const totalFinalizacoes = finalizacoes.length || totalFinalizacoesEventos;

  const ultimoEvento =
    transmissoes.length > 0
      ? [...transmissoes].sort((a, b) => new Date(b.data) - new Date(a.data))[0]
      : null;

  return {
    ...item,

    modoOperacao:
  normalizarModoOperacao(
    item?.modoOperacao || "sem-base"
  ),

    id: Number(item?.id) || 0,
    tipo: item?.tipo || "",
    nome: item?.nome || "",
    inicio: Number(item?.inicio) || 0,
    fim: Number(item?.fim) || 0,
    sequencia: Number(item?.sequencia) || 0,
    status: item?.status || "pendente",
    observacoes: item?.observacoes || "",
    totalPosicoes:
      Number(item?.totalPosicoes) || calcularTotalPosicoesEndereco(item?.inicio, item?.fim),
    transmissoes,
    finalizacoes,
    posicoesConcluidas: Number(item?.posicoesConcluidas) || 0,
posicoesPendentes: Number(item?.posicoesPendentes) || 0,
posicoesEmContagem: Number(item?.posicoesEmContagem) || 0,
posicoesDuplicadas: Number(item?.posicoesDuplicadas) || 0,
    consolidacoesPorNumero: Array.isArray(item?.consolidacoesPorNumero) ? item.consolidacoesPorNumero : [],
    contagensRecebidas: Number(item?.contagensRecebidas) || totalTransmissoes,
    finalizadoViaColetor:
  typeof item?.finalizadoViaColetor === "boolean"
    ? item.finalizadoViaColetor
    : (totalFinalizacoes > 0),
    consolidadoNoSistema: !!item?.consolidadoNoSistema,
    ultimaContagemEm: item?.ultimaContagemEm || ultimoEvento?.data || null,
    finalizadoEm: item?.finalizadoEm || null,
    criadoEm: item?.criadoEm || new Date().toISOString(),
    atualizadoEm: item?.atualizadoEm || new Date().toISOString(),
  };
}
async function carregarEnderecamentos() {
  try {
    garantirPastaData();

    if (usarPostgres) {
      const lidosBanco = await carregarEnderecamentosPostgres();

      if (Array.isArray(lidosBanco) && lidosBanco.length > 0) {
        enderecamentos = lidosBanco.map((item) => {
          const normalizado = normalizarEnderecoSalvo(item);
          const resumoFaixa = recalcularStatusFaixa(normalizado);

          return {
            ...normalizado,
            status: resumoFaixa.status,
            totalPosicoes: resumoFaixa.totalPosicoes,
            posicoesConcluidas: resumoFaixa.concluidos,
            posicoesPendentes: resumoFaixa.pendentes,
            posicoesEmContagem: resumoFaixa.emContagem,
            posicoesDuplicadas: resumoFaixa.duplicados,
          };
        });

        fs.writeFileSync(
          enderecamentosPath,
          JSON.stringify(enderecamentos, null, 2),
          "utf8"
        );

        console.log(`✅ Endereçamentos carregados do PostgreSQL: ${enderecamentos.length}`);
        return;
      }
    }

    if (!fs.existsSync(enderecamentosPath)) {
      fs.writeFileSync(enderecamentosPath, JSON.stringify([], null, 2), "utf8");
      enderecamentos = [];

      if (usarPostgres) {
        await salvarEnderecamentosPostgres(enderecamentos);
      }

      return;
    }

    const conteudo = fs.readFileSync(enderecamentosPath, "utf8");
    const lidos = JSON.parse(conteudo || "[]");

    enderecamentos = Array.isArray(lidos)
      ? lidos.map((item) => {
          const normalizado = normalizarEnderecoSalvo(item);
          const resumoFaixa = recalcularStatusFaixa(normalizado);

          return {
            ...normalizado,
            status: resumoFaixa.status,
            totalPosicoes: resumoFaixa.totalPosicoes,
            posicoesConcluidas: resumoFaixa.concluidos,
            posicoesPendentes: resumoFaixa.pendentes,
            posicoesEmContagem: resumoFaixa.emContagem,
            posicoesDuplicadas: resumoFaixa.duplicados,
          };
        })
      : [];

    if (usarPostgres) {
      await salvarEnderecamentosPostgres(enderecamentos);
      console.log(`✅ Endereçamentos migrados do JSON para PostgreSQL: ${enderecamentos.length}`);
    }
  } catch (erro) {
    console.error("Erro ao carregar endereçamentos:", erro.message);
    enderecamentos = [];
  }
}

async function salvarEnderecamentos(
  opcoes = {}
) {
  try {
    garantirPastaData();


    /*
      JSON local sempre representa
      exatamente o estado atual da memória.
    */
    fs.writeFileSync(
      enderecamentosPath,
      JSON.stringify(
        enderecamentos,
        null,
        2
      ),
      "utf8"
    );


    /*
      PostgreSQL:
      lista vazia só poderá apagar a tabela
      quando isso tiver sido autorizado
      explicitamente.
    */
    if (usarPostgres) {
      await salvarEnderecamentosPostgres(
        enderecamentos,
        {
          permitirVazio:
            opcoes?.permitirVazio === true,
        }
      );
    }


    console.log(
      `✅ Endereçamentos salvos: ${
        Array.isArray(enderecamentos)
          ? enderecamentos.length
          : 0
      }`
    );

  } catch (erro) {
    console.error(
      "Erro ao salvar endereçamentos:",
      erro.message
    );

    throw erro;
  }
}

function gerarNovoIdEnderecamento() {
  const ids = enderecamentos.map((e) => Number(e.id) || 0);
  return ids.length ? Math.max(...ids) + 1 : 1;
}
function buscarEnderecoPorNumero(
  enderecoNumero,
  modoReferencia = modoOperacao
) {
  const numero = Number(
    enderecoNumero
  );

  if (
    !Number.isFinite(numero) ||
    numero <= 0
  ) {
    return null;
  }

  const modoAtual =
    normalizarModoOperacao(
      modoReferencia
    );

  const lista =
    Array.isArray(enderecamentos)
      ? enderecamentos
      : [];

  return (
    lista.find((item) => {
      const modoEndereco =
        normalizarModoOperacao(
          item?.modoOperacao
        );

      if (
        modoEndereco !==
        modoAtual
      ) {
        return false;
      }

      const inicio =
        Number(item?.inicio) || 0;

      const fim =
        Number(item?.fim) || 0;

      return (
        numero >= inicio &&
        numero <= fim
      );
    }) || null
  );
}

function existeFaixaDuplicadaOuSobreposta({ idIgnorar = null, tipo, inicio, fim }) {
  const nInicio = Number(inicio) || 0;
  const nFim = Number(fim) || 0;
  const tipoNormalizado = normalizarTextoEndereco(tipo).toLowerCase();

  return enderecamentos.find((item) => {
    if (idIgnorar !== null && Number(item.id) === Number(idIgnorar)) {
      return false;
    }


    const modoItem =
    normalizarModoOperacao(
      item?.modoOperacao
    );
  
  const modoAtual =
    normalizarModoOperacao(
      modoOperacao
    );
  
  if (modoItem !== modoAtual) {
    return false;
  }

    const itemTipo = normalizarTextoEndereco(item.tipo).toLowerCase();
    if (itemTipo !== tipoNormalizado) {
      return false;
    }

    const itemInicio = Number(item.inicio) || 0;
    const itemFim = Number(item.fim) || 0;

    const faixaDuplicada = itemInicio === nInicio && itemFim === nFim;
    const faixaSobreposta = nInicio <= itemFim && nFim >= itemInicio;

    return faixaDuplicada || faixaSobreposta;
  });
}


function obterConsolidacaoPorNumero(endereco, enderecoNumero) {
  const consolidacoes = Array.isArray(endereco?.consolidacoesPorNumero)
    ? endereco.consolidacoesPorNumero
    : [];

  return consolidacoes.find(
    (item) => Number(item.enderecoNumero) === Number(enderecoNumero)
  ) || null;
}
function recalcularStatusFaixa(endereco) {
  const inicio = Number(endereco?.inicio) || 0;
  const fim = Number(endereco?.fim) || 0;
  const eventos = Array.isArray(endereco?.transmissoes)
  ? endereco.transmissoes.filter(
      (evento) =>
        !evento?.excluida &&
        eventoPertenceAoModo(evento)
    )
  : [];

const finalizacoesAtivas = Array.isArray(
  endereco?.finalizacoes
)
  ? endereco.finalizacoes.filter(
      (finalizacao) =>
        !finalizacao?.excluida &&
        eventoPertenceAoModo(finalizacao)
    )
  : [];
  
  if (fim < inicio) {
    return {
      status: "pendente",
      totalPosicoes: 0,
      concluidos: 0,
      pendentes: 0,
      emContagem: 0,
      duplicados: 0,
    };
  }

  const mapaTransmissoes = new Map();
  const mapaFinalizacoes = new Map();

  for (let numero = inicio; numero <= fim; numero += 1) {
    mapaTransmissoes.set(numero, 0);
    mapaFinalizacoes.set(numero, 0);
  }

  eventos
    .filter((e) => e.tipo === "transmissao")
    .forEach((e) => {
      const numero = Number(e.enderecoNumero) || 0;
      if (mapaTransmissoes.has(numero)) {
        mapaTransmissoes.set(numero, (mapaTransmissoes.get(numero) || 0) + 1);
      }
    });

  finalizacoesAtivas.forEach((f) => {
    const numero = Number(f.enderecoNumero) || 0;
    if (mapaFinalizacoes.has(numero)) {
      mapaFinalizacoes.set(numero, (mapaFinalizacoes.get(numero) || 0) + 1);
    }
  });

  let concluidos = 0;
  let pendentes = 0;
  let emContagem = 0;
  let duplicados = 0;

  for (let numero = inicio; numero <= fim; numero += 1) {
    const qtdTrans = mapaTransmissoes.get(numero) || 0;
    const qtdFin = mapaFinalizacoes.get(numero) || 0;

    if (qtdFin > 1) {
      duplicados += 1;
      continue;
    }

    if (qtdFin === 1) {
      concluidos += 1;
      continue;
    }

    if (qtdTrans > 0) {
      emContagem += 1;
      continue;
    }

    pendentes += 1;
  }

  let status = "pendente";

  if (duplicados > 0) {
    status = "recontagem";
  } else if (concluidos > 0 && concluidos === (fim - inicio + 1)) {
    status = "concluido";
  } else if (concluidos > 0 || emContagem > 0) {
    status = "em-contagem";
  }

  return {
    status,
    totalPosicoes: fim - inicio + 1,
    concluidos,
    pendentes,
    emContagem,
    duplicados,
  };
}

function obterResumoContagensDoProduto(codigoBarras) {
  const codigo = String(codigoBarras || "").trim();

  if (!codigo) {
    return {
      qtdeContada: 0,
      enderecos: [],
      endereco: "",
      enderecoNumero: ""
    };
  }

  const registros = contagens.filter((c) => {
    return (
      c &&
      c.ativo !== false &&
      String(c.codigoBarras || "").trim() === codigo
    );
  });

  if (!registros.length) {
    return {
      qtdeContada: 0,
      enderecos: [],
      endereco: "",
      enderecoNumero: ""
    };
  }

  let qtdeContada = 0;
  const mapa = new Map();

  registros.forEach((registro) => {
    const qtd = Number(registro.quantidade) || 0;
    qtdeContada += qtd;

    const enderecoNumero = Number(registro.enderecoNumero) || 0;

    if (!enderecoNumero) {
      return;
    }

    const enderecoObj = registro.enderecoId
      ? enderecamentos.find((e) => Number(e.id) === Number(registro.enderecoId))
      : buscarEnderecoPorNumero(enderecoNumero);

    const nomeEndereco = enderecoObj?.nome || "ENDEREÇO";
    const chave = `${nomeEndereco}::${enderecoNumero}`;

    if (!mapa.has(chave)) {
      mapa.set(chave, {
        enderecoId: enderecoObj?.id || null,
        enderecoNumero,
        nome: nomeEndereco,
        quantidade: 0
      });
    }

    mapa.get(chave).quantidade += qtd;
  });

  const enderecos = Array.from(mapa.values()).map((e) => ({
    ...e,
    texto: `${e.nome} • ${e.enderecoNumero} (${e.quantidade})`
  }));

  return {
    qtdeContada,
    enderecos,
    endereco: enderecos.map((e) => e.texto).join(" | "),
    enderecoNumero: enderecos.map((e) => e.enderecoNumero).join(", ")
  };
}
function gerarIdFinalizacao(endereco) {
  return `FIN-${Number(endereco.id)}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}
async function registrarEventoEndereco(
  enderecoNumero,
  tipoEvento,
  usuario = "sistema",
  itensEvento = null
) {
  const endereco = buscarEnderecoPorNumero(enderecoNumero);
  if (!endereco) return null;

  if (!Array.isArray(endereco.transmissoes)) {
    endereco.transmissoes = [];
  }

  if (!Array.isArray(endereco.finalizacoes)) {
    endereco.finalizacoes = [];
  }

  const agoraIso = new Date().toISOString();

  const novoEvento = {
    tipo: tipoEvento,

    modoOperacao:
    normalizarModoOperacao(
      modoOperacao
    ),

    enderecoNumero:
      Number(enderecoNumero),
    usuario,
    data: agoraIso,
  
    itens:
      Array.isArray(itensEvento)
        ? itensEvento
        : [],
  };
  endereco.transmissoes.push(novoEvento);

  const totalTransmissoes = endereco.transmissoes.filter(
    (e) => e.tipo === "transmissao"
  ).length;

  const totalFinalizacoes = endereco.transmissoes.filter(
    (e) => e.tipo === "finalizacao"
  ).length;

  endereco.contagensRecebidas = totalTransmissoes;
  endereco.ultimaContagemEm = agoraIso;

  if (tipoEvento === "finalizacao") {
    const finalizacaoId = gerarIdFinalizacao(endereco);
  
    const itensDaFinalizacao =
  Array.isArray(itensEvento)
    ? itensEvento.map(
        (item, indice) => ({
          contagemId:
            item.contagemId ||
            `SEM-BASE-${Date.now()}-${indice}`,

          codigoBarras:
            item.codigoBarras ||
            item.ean ||
            "",

          codigo:
            item.codigo || "",

          quantidade:
            Number(
              item.quantidade
            ) || 0,

          usuario:
            item.usuario ||
            usuario,

          data:
            item.data ||
            agoraIso,
        })
      )
    : contagens
        .filter((c) => {
          return (
            c.ativo !== false &&
            Number(
              c.enderecoNumero
            ) ===
              Number(
                enderecoNumero
              ) &&
            !c.finalizacaoId
          );
        })
        .map((c) => ({
          contagemId: c.id,
          codigoBarras:
            c.codigoBarras,
          codigo: c.codigo || "",
          quantidade:
            Number(
              c.quantidade
            ) || 0,
          usuario: c.usuario,
          data: c.data,
        }));
  
        if (
          !Array.isArray(itensEvento)
        ) {
          contagens = contagens.map(
            (c) => {
              if (
                c.ativo !== false &&
                Number(
                  c.enderecoNumero
                ) ===
                  Number(
                    enderecoNumero
                  ) &&
                !c.finalizacaoId
              ) {
                return {
                  ...c,
                  finalizacaoId,
                };
              }
        
              return c;
            }
          );
        }
  
        endereco.finalizacoes.push({
          id: finalizacaoId,
        
          modoOperacao:
            normalizarModoOperacao(
              modoOperacao
            ),
      enderecoId: Number(endereco.id),
      enderecoNumero: Number(enderecoNumero),
      usuario,
      data: agoraIso,
      excluida: false,
      itens: itensDaFinalizacao,
    });
  
    endereco.finalizadoViaColetor = true;
    endereco.finalizadoEm = agoraIso;
  
    await salvarContagens();
  }
  
  const resumoFaixa = recalcularStatusFaixa(endereco);

endereco.status = resumoFaixa.status;
endereco.totalPosicoes = resumoFaixa.totalPosicoes;
endereco.posicoesConcluidas = resumoFaixa.concluidos;
endereco.posicoesPendentes = resumoFaixa.pendentes;
endereco.posicoesEmContagem = resumoFaixa.emContagem;
endereco.posicoesDuplicadas = resumoFaixa.duplicados;
  endereco.atualizadoEm = agoraIso;
  await salvarEnderecamentos();

  return endereco;
}

function gerarAuditoriaDuplicidadeEnderecos() {
  return enderecamentos.flatMap((item) => {
    const eventos = Array.isArray(item.transmissoes) ? item.transmissoes : [];
    const finalizacoes = Array.isArray(item.finalizacoes)
      ? item.finalizacoes.filter((f) => !f.excluida)
      : [];

    const auditoriaPorNumero = [];

    for (let numero = Number(item.inicio) || 0; numero <= (Number(item.fim) || 0); numero += 1) {
      const eventosDoNumero = eventos.filter(
        (e) => Number(e.enderecoNumero) === Number(numero)
      );

      const transmissoesDoNumero = eventosDoNumero.filter(
        (e) => e.tipo === "transmissao"
      );

      const finalizacoesDoNumero = finalizacoes.filter(
        (f) => Number(f.enderecoNumero) === Number(numero)
      );
      const transmissoesPendentesDoNumero = transmissoesDoNumero.filter(
        (t) => String(t.statusConsolidacao || "").toLowerCase() === "pendente"
      );
      
      const jaFoiConsolidado = !!obterConsolidacaoPorNumero(item, numero)?.consolidado;
      
      const duplicadoOperacional =
        finalizacoesDoNumero.length > 1 ||
        transmissoesPendentesDoNumero.length > 1 ||
        (jaFoiConsolidado && transmissoesPendentesDoNumero.length > 0);

        function montarItensDetalhados(itens, usuarioPadrao, dataPadrao) {
          return (Array.isArray(itens) ? itens : []).map((it, index) => {
            const codigoBarras = String(it.codigoBarras || "").trim();
        
            const produto = inventario.find(
              (p) => String(p.codigoBarras || "").trim() === codigoBarras
            );
        
            return {
              contagemId: it.contagemId || it.id || `ITEM-${Date.now()}-${index}`,
              codigoBarras,
              codigo: produto?.codigo || produto?.codigoInterno || it.codigo || "",
              descricao: produto?.descricao || it.descricao || "Item não encontrado",
              quantidade: Number(it.quantidade) || 0,
              usuario: it.usuario || usuarioPadrao || "--",
              data: it.data || dataPadrao || null,
            };
          });
        }
        
        const finalizacoesReaisDetalhadas = finalizacoesDoNumero.map((finalizacao) => {
          const itensDetalhados = montarItensDetalhados(
            finalizacao.itens,
            finalizacao.usuario,
            finalizacao.data
          );
        
          return {
            id: finalizacao.id,
            origem: "finalizacao",
            usuario: finalizacao.usuario || "--",
            data: finalizacao.data || null,
            enderecoNumero: Number(finalizacao.enderecoNumero) || Number(numero),
            itens: itensDetalhados,
            totalItens: itensDetalhados.length,
          };
        });
        
        const finalizacoesPorTransmissao = transmissoesDoNumero
  .map((transmissao) => ({
    transmissao,
    indiceOriginal: eventos.indexOf(transmissao),
  }))
  .filter(({ transmissao }) => !transmissao.excluida)
  .map(({ transmissao, indiceOriginal }) => {
            const itensDetalhados = montarItensDetalhados(
              transmissao.itens,
              transmissao.usuario,
              transmissao.data
            );
        
            return {
              id: transmissao.id || `TRANS-${item.id}-${numero}-${indiceOriginal}`,
              origem: "transmissao",
              usuario: transmissao.usuario || "--",
              data: transmissao.data || null,
              enderecoNumero: Number(transmissao.enderecoNumero) || Number(numero),
              itens: itensDetalhados,
              totalItens: itensDetalhados.length,
              statusConsolidacao: transmissao.statusConsolidacao || "pendente",
            };
          });
        
        const finalizacoesDetalhadas = [
          ...finalizacoesReaisDetalhadas,
          ...finalizacoesPorTransmissao,
        ];

      const consolidacaoNumero = obterConsolidacaoPorNumero(item, numero);

      const contagensAtivasDoNumero = contagens.filter((c) => {
        return (
          c &&
          c.ativo !== false &&
          Number(c.enderecoNumero) === Number(numero)
        );
      });

      auditoriaPorNumero.push({
        id: item.id,
        enderecoId: item.id,
        chaveAuditoria: `${item.id}-${numero}`,
        tipo: item.tipo,
        nome: item.nome,
        inicio: Number(numero),
        fim: Number(numero),
        enderecoNumero: Number(numero),
        status:
  duplicadoOperacional
    ? "recontagem"
    : finalizacoesDetalhadas.length === 1 || jaFoiConsolidado
    ? "concluido"
    : transmissoesDoNumero.length > 0
    ? "em-contagem"
    : "pendente",
        totalPosicoes: 1,
        contagensRecebidas: Math.max(
          contagensAtivasDoNumero.length,
          transmissoesDoNumero.length
        ),
        totalEventos: eventosDoNumero.length,
        totalTransmissoes: transmissoesDoNumero.length,
        totalFinalizacoes: finalizacoesDetalhadas.length,
        duplicado: duplicadoOperacional,
duplicidadeResolvida: !duplicadoOperacional,
        duplicidadeResolvidaEm: finalizacoesDetalhadas.length <= 1 ? new Date().toISOString() : null,
        duplicidadeResolvidaPor: finalizacoesDetalhadas.length <= 1 ? "sistema" : null,
        finalizacoes: finalizacoesDetalhadas,
        eventos: eventosDoNumero,
        ultimaContagemEm:
          eventosDoNumero.length > 0
            ? [...eventosDoNumero].sort((a, b) => new Date(b.data) - new Date(a.data))[0]?.data || null
            : null,
        finalizadoEm:
          finalizacoesDetalhadas.length > 0
            ? [...finalizacoesDetalhadas].sort((a, b) => new Date(b.data) - new Date(a.data))[0]?.data || null
            : null,
            finalizadoViaColetor:
            finalizacoesDetalhadas.length > 0 ||
            (
              !!item.finalizadoViaColetor &&
              transmissoesDoNumero.some((t) =>
                t.tipo === "transmissao" &&
                t.statusConsolidacao === "pendente"
              )
            ),
        consolidadoNoSistema: !!consolidacaoNumero?.consolidado,
        consolidadoEm: consolidacaoNumero?.consolidadoEm || null,
        consolidadoPor: consolidacaoNumero?.consolidadoPor || null,
      });
    }

    return auditoriaPorNumero;
  });
}
const FUSO_HORARIO_REALSTOCK =
  process.env.APP_TIMEZONE ||
  "America/Manaus";
  function formatarDataHoraTransmissao(valor) {
    if (!valor) return "--";
  
    const d = new Date(valor);
  
    if (isNaN(d.getTime())) {
      return "--";
    }
  
    return d.toLocaleString("pt-BR", {
      timeZone: FUSO_HORARIO_REALSTOCK,
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    });
  }

function contarItensDoEndereco(
  endereco,
  enderecoNumero = null
) {
  const numero = String(
    enderecoNumero || ""
  ).trim();

  if (
    modoOperacao === "sem-base"
  ) {
    const codigos = new Set();

    (
      Array.isArray(
        contagemSemBase
      )
        ? contagemSemBase
        : []
    ).forEach((item) => {
      const possuiEndereco =
        (
          Array.isArray(
            item.enderecos
          )
            ? item.enderecos
            : []
        ).some(
          (registro) =>
            String(
              registro.enderecoNumero ||
              ""
            ).trim() === numero &&
            Number(
              registro.quantidade
            ) > 0
        );

      if (!possuiEndereco) {
        return;
      }

      const chave =
        String(
          item.ean ||
          item.codigo ||
          ""
        ).trim();

      if (chave) {
        codigos.add(chave);
      }
    });

    return codigos.size;
  }

  const finalizacoes =
    Array.isArray(
      endereco?.finalizacoes
    )
      ? endereco.finalizacoes.filter(
          (item) =>
            !item.excluida &&
            (
              !numero ||
              Number(
                item.enderecoNumero
              ) === Number(numero)
            )
        )
      : [];

  const itens =
    finalizacoes.flatMap(
      (item) =>
        Array.isArray(item.itens)
          ? item.itens
          : []
    );

  return new Set(
    itens
      .map(
        (item) =>
          String(
            item.codigoBarras ||
            item.ean ||
            item.codigo ||
            ""
          ).trim()
      )
      .filter(Boolean)
  ).size;
}

function obterOrigemTransmissaoEndereco(endereco) {
  const eventos = Array.isArray(endereco?.transmissoes) ? endereco.transmissoes : [];
  if (!eventos.length) return "Coletor/App";

  const ultimo = [...eventos].sort((a, b) => new Date(b.data) - new Date(a.data))[0];
  return ultimo?.usuario || "Coletor/App";
}

function gerarPainelTransmissoesConsolidacao() {
  const auditoriaDuplicados = gerarAuditoriaDuplicidadeEnderecos();

  const recebidas = [];
  const fila = [];
  const consolidadas = [];
  const pendencias = [];

  auditoriaDuplicados.forEach((endereco) => {
    const temContagens = Number(endereco.contagensRecebidas || 0) > 0;
    if (!temContagens) return;

    const ultimoEvento = Array.isArray(endereco.eventos) && endereco.eventos.length
  ? [...endereco.eventos].sort((a, b) => new Date(b.data) - new Date(a.data))[0]
  : null;

const ultimoEnderecoContado =
  ultimoEvento?.enderecoNumero ??
  (Array.isArray(endereco.finalizacoes) && endereco.finalizacoes.length
    ? endereco.finalizacoes[endereco.finalizacoes.length - 1]?.enderecoNumero
    : null);

    const itemBase = {
      id: `END-${endereco.id}-${endereco.enderecoNumero}`,
      enderecoId: Number(endereco.id),
      nome: endereco.nome || "--",
      tipo: endereco.tipo || "--",
      faixa: `${endereco.inicio || "--"} até ${endereco.fim || "--"}`,
      enderecoExato: ultimoEnderecoContado || "",
      data: formatarDataHoraTransmissao(endereco.ultimaContagemEm),
      itens: contarItensDoEndereco(
        endereco,
        endereco.enderecoNumero ||
          ultimoEnderecoContado
      ),
      origem: obterOrigemTransmissaoEndereco(endereco),
      contagensRecebidas: Number(endereco.contagensRecebidas || 0),
      duplicado: !!endereco.duplicado,
      finalizadoViaColetor: !!endereco.finalizadoViaColetor ||                    // Original
                             (Array.isArray(endereco.transmissoes) &&             // ✅ Mobile
                              endereco.transmissoes.filter(t => t.tipo === "transmissao").length > 0),
      consolidadoNoSistema: !!endereco.consolidadoNoSistema
    };  // ← VÍRGULA AQUI!
    
    const duplicidadeResolvida = !!endereco.duplicidadeResolvida;

    if (itemBase.duplicado && !duplicidadeResolvida) {
      pendencias.push({
        ...itemBase,
        status: "Pendência",
        motivo: `Endereço duplicado (${endereco.nome || "--"} • ${endereco.inicio || "--"} até ${endereco.fim || "--"}). Revise no módulo de endereçamento antes de consolidar.`,
      });
      return;
    }

    if (itemBase.consolidadoNoSistema) {
      consolidadas.push({
        ...itemBase,
        status: "Consolidada",
      });
      return;
    }

    if (itemBase.finalizadoViaColetor && !itemBase.consolidadoNoSistema) {
      fila.push({
        ...itemBase,
        status: "Em fila",
      });
      return;
    }

    recebidas.push({
      ...itemBase,
      status: "Recebida",
    });
  });

  return {
    recebidas,
    fila,
    consolidadas,
    pendencias,
    totais: {
      recebidas: recebidas.length,
      fila: fila.length,
      consolidadas: consolidadas.length,
      erros: pendencias.length,
    },
    atualizadoEm: new Date().toISOString(),
  };
}

const possuiDatabaseUrl =
  Boolean(
    String(process.env.DATABASE_URL || "").trim()
  );

const valorUsePostgres = String(
  process.env.USE_POSTGRES || ""
)
  .trim()
  .toLowerCase();

const usarPostgres =
  possuiDatabaseUrl &&
  (
    process.env.NODE_ENV === "production" ||
    ["true", "1", "sim", "yes"].includes(
      valorUsePostgres
    )
  );

console.log("Configuração PostgreSQL:", {
  nodeEnv: process.env.NODE_ENV || "development",
  possuiDatabaseUrl,
  usePostgresEnv:
    process.env.USE_POSTGRES || "não informado",
  usarPostgres,
});
async function carregarUsuarios() {
  try {
    garantirPastaData();

    const usuariosDoBanco = usarPostgres
  ? await carregarUsuariosPostgres()
  : [];

if (usarPostgres && Array.isArray(usuariosDoBanco) && usuariosDoBanco.length > 0) {
      usuarios = usuariosDoBanco;

      fs.writeFileSync(
        usuariosPath,
        JSON.stringify(usuarios, null, 2),
        "utf8"
      );

      console.log(`✅ Usuários carregados do PostgreSQL: ${usuarios.length}`);
      return;
    }

    if (!fs.existsSync(usuariosPath)) {
      const usuariosIniciais = [
        {
          id: 1,
          nome: "Administrador",
          usuario: "admin",
          senha: "realstockj14",
          matricula: "RS-0001",
          funcao: "Administrador",
          telefone: "",
          status: "ativo",
          meta: 0,
          criadoEm: new Date().toISOString(),
        },
        {
          id: 2,
          nome: "Juliana Santos ",
          usuario: "Juliana",
          senha: "realstockj14",
          matricula: "RS-0002",
          funcao: "Operador",
          telefone: "",
          status: "ativo",
          meta: 0,
          criadoEm: new Date().toISOString(),
        },
      ];

      usuarios = usuariosIniciais;

      fs.writeFileSync(
        usuariosPath,
        JSON.stringify(usuarios, null, 2),
        "utf8"
      );

      if (usarPostgres) {
        await salvarUsuariosPostgres(usuarios);
      }

      console.log("✅ Usuários iniciais criados no JSON e PostgreSQL.");
      return;
    }

    const conteudo = fs.readFileSync(usuariosPath, "utf8");
    const lidos = JSON.parse(conteudo || "[]");

    usuarios = Array.isArray(lidos)
      ? lidos.map((u) => ({
          ...u,
          status: u.status || "ativo",
        }))
      : [];

      if (usarPostgres) {
        await salvarUsuariosPostgres(usuarios);
      }

    console.log(`✅ Usuários migrados do JSON para PostgreSQL: ${usuarios.length}`);
  } catch (erro) {
    console.error("Erro ao carregar usuários:", erro.message);

    try {
      const conteudo = fs.readFileSync(usuariosPath, "utf8");
      usuarios = JSON.parse(conteudo || "[]");
      console.log("⚠️ Usuários carregados do JSON como backup.");
    } catch {
      usuarios = [];
    }
  }
}

async function salvarUsuarios() {
  try {
    garantirPastaData();

    fs.writeFileSync(
      usuariosPath,
      JSON.stringify(usuarios, null, 2),
      "utf8"
    );

    if (usarPostgres) {
      await salvarUsuariosPostgres(usuarios);
    }

    console.log("✅ Usuários salvos no JSON e PostgreSQL.");
  } catch (erro) {
    console.error("Erro ao salvar usuários:", erro.message);
  }
}

function carregarLayoutTxt() {
  try {
    garantirPastaData();

    if (!fs.existsSync(layoutTxtPath)) {
      fs.writeFileSync(layoutTxtPath, JSON.stringify(layoutTxtPadrao, null, 2), "utf8");
      layoutTxt = JSON.parse(JSON.stringify(layoutTxtPadrao));
      return;
    }

    const conteudo = fs.readFileSync(layoutTxtPath, "utf8");
    const lido = JSON.parse(conteudo || "{}");

    layoutTxt = {
      ...JSON.parse(JSON.stringify(layoutTxtPadrao)),
      ...lido,
    };
  } catch (erro) {
    console.error("Erro ao carregar layout TXT:", erro);
    layoutTxt = JSON.parse(JSON.stringify(layoutTxtPadrao));
  }
}

function salvarLayoutTxt() {
  try {
    garantirPastaData();
    fs.writeFileSync(layoutTxtPath, JSON.stringify(layoutTxt, null, 2), "utf8");
  } catch (erro) {
    console.error("Erro ao salvar layout TXT:", erro);
  }
}

function carregarLayoutsSalvos() {
  try {
    garantirPastaData();

    if (!fs.existsSync(layoutsTxtPath)) {
      fs.writeFileSync(layoutsTxtPath, JSON.stringify([], null, 2), "utf8");
      layoutsSalvos = [];
      return;
    }

    const conteudo = fs.readFileSync(layoutsTxtPath, "utf8");
    layoutsSalvos = JSON.parse(conteudo || "[]");
  } catch (erro) {
    console.error("Erro ao carregar layouts salvos:", erro);
    layoutsSalvos = [];
  }
}

function salvarLayoutsSalvos() {
  try {
    garantirPastaData();
    fs.writeFileSync(layoutsTxtPath, JSON.stringify(layoutsSalvos, null, 2), "utf8");
  } catch (erro) {
    console.error("Erro ao salvar layouts salvos:", erro);
  }
}
function carregarLayoutsUniversais() {
  try {
    garantirPastaData();

    if (!fs.existsSync(layoutsUniversaisPath)) {
      fs.writeFileSync(
        layoutsUniversaisPath,
        JSON.stringify([], null, 2),
        "utf8"
      );

      layoutsUniversais = [];
      return;
    }

    const conteudo = fs.readFileSync(
      layoutsUniversaisPath,
      "utf8"
    );

    const dados = JSON.parse(conteudo || "[]");

    layoutsUniversais = Array.isArray(dados)
      ? dados
      : [];
  } catch (erro) {
    console.error(
      "Erro ao carregar layouts universais:",
      erro
    );

    layoutsUniversais = [];
  }
}

function salvarLayoutsUniversais() {
  try {
    garantirPastaData();

    fs.writeFileSync(
      layoutsUniversaisPath,
      JSON.stringify(
        layoutsUniversais,
        null,
        2
      ),
      "utf8"
    );
  } catch (erro) {
    console.error(
      "Erro ao salvar layouts universais:",
      erro
    );
  }
}

function gerarIdLayoutUniversal() {
  return [
    "LU",
    Date.now(),
    Math.random()
      .toString(36)
      .slice(2, 8),
  ].join("-");
}

let layoutsExportacao = [];

function carregarLayoutsExportacao() {
  try {
    garantirPastaData();

    if (!fs.existsSync(layoutsExportacaoPath)) {
      fs.writeFileSync(
        layoutsExportacaoPath,
        JSON.stringify([], null, 2),
        "utf8"
      );
      layoutsExportacao = [];
      return;
    }

    const conteudo = fs.readFileSync(layoutsExportacaoPath, "utf8");
    const dados = JSON.parse(conteudo || "[]");
    layoutsExportacao = Array.isArray(dados) ? dados : [];
  } catch (erro) {
    console.error("Erro ao carregar layouts de exportação:", erro);
    layoutsExportacao = [];
  }
}

function salvarLayoutsExportacao() {
  try {
    garantirPastaData();
    fs.writeFileSync(
      layoutsExportacaoPath,
      JSON.stringify(layoutsExportacao, null, 2),
      "utf8"
    );
  } catch (erro) {
    console.error("Erro ao salvar layouts de exportação:", erro);
  }
}

function gerarIdLayoutExportacao() {
  return ["LE", Date.now(), Math.random().toString(36).slice(2, 8)].join("-");
}

const COLUNAS_EXPORTACAO_DISPONIVEIS = [
  { chave: "codigoBarras", rotulo: "Código EAN" },
  { chave: "codigo", rotulo: "Código interno" },
  { chave: "codigoInterno", rotulo: "Código interno alternativo" },
  { chave: "descricao", rotulo: "Descrição" },
  { chave: "categoria", rotulo: "Categoria" },
  { chave: "custoUnitario", rotulo: "Custo unitário" },
  { chave: "qtdeCongelada", rotulo: "Qtde congelada" },
  { chave: "qtdeContada", rotulo: "Qtde contada" },
  { chave: "divergencia", rotulo: "Divergência" },
  { chave: "situacao", rotulo: "Situação" },
  { chave: "ajuste", rotulo: "Ajuste" },
  { chave: "endereco", rotulo: "Endereço" },
  { chave: "enderecoNumero", rotulo: "Número do endereço" },
  { chave: "coletaEndereco", rotulo: "Coleta endereço" },
  { chave: "valorCongelado", rotulo: "Valor congelado" },
  { chave: "valorContado", rotulo: "Valor contado" },
  { chave: "valorDivergencia", rotulo: "Valor divergência" },
];

const FORMATOS_EXPORTACAO_PERMITIDOS = new Set(["txt", "txt-retorno", "csv", "json", "xlsx"]);

function normalizarLayoutExportacao(dados = {}, layoutAtual = null) {
  const formato = String(dados.formato || "").trim().toLowerCase();
  const cliente = String(dados.cliente || layoutAtual?.cliente || "").trim();
  const nome = String(dados.nome || layoutAtual?.nome || "").trim();

  const colunasRecebidas = Array.isArray(dados.colunas) ? dados.colunas : [];
  const chavesValidas = new Set(
    COLUNAS_EXPORTACAO_DISPONIVEIS.map((c) => c.chave)
  );

  const colunas = colunasRecebidas
    .filter((col) => col && chavesValidas.has(col.chave))
    .map((col, indice) => ({
      chave: col.chave,
      rotulo: String(col.rotulo || col.chave || "").trim() || col.chave,
      incluida: col.incluida !== false,
      ordem: Number.isFinite(Number(col.ordem)) ? Number(col.ordem) : indice,
    }));

  return {
    id: layoutAtual?.id || gerarIdLayoutExportacao(),
    nome,
    cliente,
    observacao: String(dados.observacao || layoutAtual?.observacao || "").trim(),
    formato: FORMATOS_EXPORTACAO_PERMITIDOS.has(formato) ? formato : "csv",
    delimitador: String(dados.delimitador || layoutAtual?.delimitador || ";") || ";",
    possuiCabecalho: dados.possuiCabecalho !== false,
    txtRetorno: {
      prefixo: String((dados.txtRetorno && dados.txtRetorno.prefixo) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.prefixo) || "0001010"),
      sufixo: String((dados.txtRetorno && dados.txtRetorno.sufixo) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.sufixo) || "000000"),
      tamanhoEan: Number((dados.txtRetorno && dados.txtRetorno.tamanhoEan) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.tamanhoEan) || 13),
      tamanhoQtd: Number((dados.txtRetorno && dados.txtRetorno.tamanhoQtd) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.tamanhoQtd) || 8),
      tipoQuantidade: String((dados.txtRetorno && dados.txtRetorno.tipoQuantidade) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.tipoQuantidade) || "auto"),
      somenteContados: !!((dados.txtRetorno && dados.txtRetorno.somenteContados) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.somenteContados)),
      somenteDivergencia: !!((dados.txtRetorno && dados.txtRetorno.somenteDivergencia) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.somenteDivergencia)),
      somenteMaiorZero: !!((dados.txtRetorno && dados.txtRetorno.somenteMaiorZero) || (layoutAtual?.txtRetorno && layoutAtual.txtRetorno.somenteMaiorZero)),
    },
    colunas,
    criadoEm: layoutAtual?.criadoEm || new Date().toISOString(),
    atualizadoEm: new Date().toISOString(),
    ultimoUsoEm: layoutAtual?.ultimoUsoEm || null,
  };
}

function obterSituacaoItem(item) {
  const qtdeContada = Number(item.qtdeContada) || 0;
  const qtdeCongelada = Number(item.qtdeCongelada) || 0;
  const divergencia = qtdeContada - qtdeCongelada;

  if (qtdeContada === 0) return "Sem contagem";
  if (Math.abs(divergencia) < 0.000001) return "Neutro";
  return divergencia > 0 ? "Sobrando" : "Faltando";
}

function obterAjusteItem(item) {
  const qtdeContada = Number(item.qtdeContada) || 0;
  const qtdeCongelada = Number(item.qtdeCongelada) || 0;
  const divergencia = qtdeContada - qtdeCongelada;

  if (qtdeContada === 0) return qtdeCongelada;
  if (Math.abs(divergencia) < 0.000001) return 0;
  return divergencia;
}

function prepararLinhasExportacao(itens, colunas) {
  const colunasAtivas = colunas
    .filter((c) => c.incluida !== false)
    .sort((a, b) => (a.ordem ?? 0) - (b.ordem ?? 0));

  const linhas = itens.map((item) => {
    const linha = {};
    colunasAtivas.forEach((col) => {
      let valor = item[col.chave];
      if (col.chave === "situacao") valor = obterSituacaoItem(item);
      if (col.chave === "ajuste") valor = obterAjusteItem(item);
      if (col.chave === "coletaEndereco") {
        valor = item.enderecosContagem && item.enderecosContagem.length > 0
          ? item.enderecosContagem.map((e) => e.enderecoNumero || e).join(", ")
          : "";
      }
      if (valor === null || valor === undefined) valor = "";
      linha[col.rotulo || col.chave] = valor;
    });
    return linha;
  });

  return { colunasAtivas, linhas };
}

function gerarCsvExportacao(linhas, colunasAtivas, delimitador, possuiCabecalho) {
  const separador = delimitador || ";";
  const cabecalhos = colunasAtivas.map((c) => c.rotulo || c.chave);

  const escapar = (valor) => {
    const texto = String(valor ?? "");
    if (texto.includes(separador) || texto.includes("\n") || texto.includes('"')) {
      return `"${texto.replace(/"/g, '""')}"`;
    }
    return texto;
  };

  const conteudo = [];
  if (possuiCabecalho) {
    conteudo.push(cabecalhos.map(escapar).join(separador));
  }
  linhas.forEach((linha) => {
    conteudo.push(
      colunasAtivas
        .map((c) => escapar(linha[c.rotulo || c.chave]))
        .join(separador)
    );
  });

  return conteudo.join("\n");
}

function gerarTxtExportacao(linhas, colunasAtivas, delimitador, possuiCabecalho) {
  return gerarCsvExportacao(linhas, colunasAtivas, delimitador || "\t", possuiCabecalho);
}

function gerarTxtRetornoExportacao(itens, opts = {}) {
  const prefixo = opts.prefixo || "0001010";
  const sufixo = opts.sufixo || "000000";
  const tamanhoEan = Number(opts.tamanhoEan) || 13;
  const tamanhoQtd = Number(opts.tamanhoQtd) || 8;
  const tipoQuantidade = opts.tipoQuantidade || "auto";

  let lista = Array.isArray(itens) ? itens : [];

  if (opts.somenteContados) {
    lista = lista.filter((item) => (Number(item.qtdeContada) || 0) > 0);
  }
  if (opts.somenteDivergencia) {
    lista = lista.filter((item) => {
      const qtdeContada = Number(item.qtdeContada) || 0;
      const qtdeCongelada = Number(item.qtdeCongelada) || 0;
      return qtdeContada - qtdeCongelada !== 0;
    });
  }
  if (opts.somenteMaiorZero) {
    lista = lista.filter((item) => (Number(item.qtdeCongelada) || 0) > 0);
  }

  let conteudo = "";
  lista.forEach((item) => {
    const codigoBase = String(item.codigoBarras || item.codigo || "").replace(/\D/g, "");
    const codigoEan = codigoBase.padStart(tamanhoEan, "0").slice(-tamanhoEan);
    const qtde = parseFloat(item.qtdeContada) || 0;
    const textoCategoria = `${item.categoria || ""} ${item.descricao || ""}`.toLowerCase();

    let quantidadeFormatada = "";
    if (tipoQuantidade === "inteiro") {
      quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
    } else if (tipoQuantidade === "milhar") {
      quantidadeFormatada = Math.round(qtde * 1000).toString().padStart(tamanhoQtd, "0");
    } else {
      if (textoCategoria.includes("kg")) {
        quantidadeFormatada =
          qtde % 1 === 0
            ? qtde.toString().padStart(tamanhoQtd, "0")
            : Math.round(qtde * 1000).toString().padStart(tamanhoQtd, "0");
      } else {
        quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
      }
    }

    conteudo += `${prefixo}${codigoEan}${quantidadeFormatada}${sufixo}\n`;
  });

  return conteudo;
}

function gerarJsonExportacao(linhas, colunasAtivas) {
  const cabecalhos = colunasAtivas.map((c) => c.rotulo || c.chave);
  const saida = linhas.map((linha) => {
    const obj = {};
    cabecalhos.forEach((cab) => {
      obj[cab] = linha[cab];
    });
    return obj;
  });
  return JSON.stringify(saida, null, 2);
}

async function gerarExcelExportacao(linhas, colunasAtivas) {
  const ExcelJS = require("exceljs");
  const workbook = new ExcelJS.Workbook();
  const planilha = workbook.addWorksheet("Inventário");

  const cabecalhos = colunasAtivas.map((c) => c.rotulo || c.chave);
  planilha.addRow(cabecalhos);

  linhas.forEach((linha) => {
    planilha.addRow(cabecalhos.map((cab) => linha[cab]));
  });

  planilha.getRow(1).font = { bold: true };
  planilha.columns.forEach((col) => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const len = String(cell.value ?? "").length;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 2, 60);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function normalizarLayoutUniversal(
  dados = {},
  layoutAtual = null
) {
  const formatosPermitidos = new Set([
    "txt",
    "txt-retorno",
    "csv",
    "json",
    "xlsx",
  ]);

  const leiturasPermitidas = new Set([
    "automatico",
    "posicao-fixa",
    "delimitado",
    "json",
    "excel",
  ]);

  const destinosPermitidos = new Set([
    "base-principal",
    "saldo-atual",
    "complementar",
    "base-wms",
  ]);

  const formato = String(
    dados.formato || ""
  )
    .trim()
    .toLowerCase();

  const tipoLeitura = String(
    dados.tipoLeitura || "automatico"
  )
    .trim()
    .toLowerCase();

  const destino = String(
    dados.destino || ""
  )
    .trim()
    .toLowerCase();

  const mapeamento =
    dados.mapeamento &&
    typeof dados.mapeamento === "object" &&
    !Array.isArray(dados.mapeamento)
      ? dados.mapeamento
      : {};
      const transformacoes =
      normalizarTransformacoesUniversais(
        dados.transformacoes ||
        layoutAtual?.transformacoes ||
        {}
      );
  return {
    id:
      layoutAtual?.id ||
      gerarIdLayoutUniversal(),

    nome: String(
      dados.nome || layoutAtual?.nome || ""
    ).trim(),

    cliente: String(
      dados.cliente ||
      layoutAtual?.cliente ||
      ""
    ).trim(),

    observacao: String(
      dados.observacao ||
      layoutAtual?.observacao ||
      ""
    ).trim(),

    formato:
      formatosPermitidos.has(formato)
        ? formato
        : "",

    tipoLeitura:
      leiturasPermitidas.has(tipoLeitura)
        ? tipoLeitura
        : "automatico",

    delimitador: String(
      dados.delimitador || ""
    ),

    possuiCabecalho:
      dados.possuiCabecalho !== false,

    destino:
      destinosPermitidos.has(destino)
        ? destino
        : "",

    mapeamento,
    transformacoes,
    criadoEm:
      layoutAtual?.criadoEm ||
      new Date().toISOString(),

    atualizadoEm:
      new Date().toISOString(),

    ultimoUsoEm:
      layoutAtual?.ultimoUsoEm || null,
  };
}
function gerarNovaMatricula() {
  const numeros = usuarios
    .map((u) => {
      const match = String(u.matricula || "").match(/RS-(\d+)/);
      return match ? Number(match[1]) : 0;
    })
    .filter((n) => !isNaN(n));

  const maior = numeros.length ? Math.max(...numeros) : 0;
  const proximo = maior + 1;

  return `RS-${String(proximo).padStart(4, "0")}`;
}

function gerarNovoIdUsuario() {
  const ids = usuarios.map((u) => Number(u.id) || 0);
  return ids.length ? Math.max(...ids) + 1 : 1;
}

function gerarResumoUsuariosAtivos() {
  return usuarios
    .filter((u) => String(u.status || "").toLowerCase() === "ativo")
    .map((u) => {
      const registrosUsuario = contagens.filter((c) => c.usuario === u.usuario);
      const ultimaAtividade =
        registrosUsuario.length > 0
          ? registrosUsuario[registrosUsuario.length - 1].data
          : null;

      const totalContado = registrosUsuario.reduce(
        (acc, item) => acc + (Number(item.quantidade) || 0),
        0
      );

      return {
        id: u.id,
        nome: u.nome || u.usuario,
        usuario: u.usuario,
        senha: u.senha || "",
        matricula: u.matricula,
        funcao: u.funcao || "Operador",
        telefone: u.telefone || "",
        status: u.status || "ativo",
        meta: Number(u.meta) || 0,
        totalContado,
        ultimaAtividade,
      };
    });
}

function gerarIdLayout() {
  return `layout-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}

function registrarContagem(
  usuario,
  codigoBarras,
  quantidade,
  enderecoNumero = null,
  finalizacaoId = null,
  statusConsolidacao = "consolidado"
) {
  const qtd = Number(quantidade) || 0;
  if (qtd <= 0) return;

  const usuarioObj = usuarios.find((u) => u.usuario === usuario);
  const endereco = enderecoNumero !== null ? buscarEnderecoPorNumero(enderecoNumero) : null;

  const consolidado = statusConsolidacao === "consolidado";

  contagens.push({
    id: `CONT-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    usuario,
    matricula: usuarioObj?.matricula || "SEM-MATRICULA",
    codigoBarras,
    quantidade: qtd,
    enderecoId: endereco?.id || null,
    enderecoNumero: enderecoNumero !== null ? Number(enderecoNumero) : null,
    finalizacaoId: finalizacaoId || null,

    ativo: consolidado,
    statusConsolidacao: consolidado ? "consolidado" : "pendente",

    data: new Date().toISOString(),
  });

  salvarContagens();
}
function gerarRankingUsuariosSemBase() {
  const mapa = {};

  (Array.isArray(finalizacoesSemBase) ? finalizacoesSemBase : []).forEach((finalizacao) => {
    const usuario = String(finalizacao.usuario || "sem-usuario").trim() || "sem-usuario";
    const itens = Array.isArray(finalizacao.itens) ? finalizacao.itens : [];

    if (!mapa[usuario]) {
      mapa[usuario] = {
        usuario,
        matricula: "SEM-MATRICULA",
        movimentacoes: 0,
        itensUnicos: new Set(),
        totalContado: 0,
        ultimaAtualizacao: null,
      };
    }

    mapa[usuario].movimentacoes += 1;
    mapa[usuario].ultimaAtualizacao = finalizacao.data || mapa[usuario].ultimaAtualizacao;

    itens.forEach((item) => {
      const chave = String(item.ean || item.codigo || "").trim();
      if (chave) {
        mapa[usuario].itensUnicos.add(chave);
      }

      mapa[usuario].totalContado += Number(item.quantidade || 0);
    });
  });

  const ranking = Object.values(mapa).map((item) => ({
    usuario: item.usuario,
    matricula: item.matricula,
    totalContado: item.totalContado,
    movimentacoes: item.movimentacoes,
    itensUnicos: item.itensUnicos.size,
    ultimaAtualizacao: item.ultimaAtualizacao,
  }));

  ranking.sort((a, b) => b.movimentacoes - a.movimentacoes || b.itensUnicos - a.itensUnicos);

  const totalGeral = ranking.reduce((acc, item) => acc + item.movimentacoes, 0);

  return ranking.map((item) => ({
    ...item,
    percentual: totalGeral > 0 ? (item.movimentacoes / totalGeral) * 100 : 0,
  }));
}
function gerarResumoDashboardSemBase() {
  const lista = Array.isArray(contagemSemBase) ? contagemSemBase : [];

  const totalProdutos = 0;
  const totalItensUnicos = lista.length;
  const totalQuantidade = lista.reduce(
    (acc, item) => acc + (Number(item.quantidade) || 0),
    0
  );

  const totalEnderecosFinalizados = Array.isArray(finalizacoesSemBase)
    ? finalizacoesSemBase.length
    : 0;

  return {
    totalProdutos,
    totalItensUnicos,
    totalQuantidade,
    totalEnderecosFinalizados,
  };
}
function gerarRankingUsuarios() {
  if (!inventario || inventario.length === 0) {
    return [];
  }

  const mapa = {};

  /*
    Mapa rápido para descobrir se cada produto
    deve ser tratado como KG ou UN.
  */
  const mapaProdutos = new Map();

  inventario.forEach((produto) => {
    const codigoBarras = String(
      produto?.codigoBarras || ''
    ).trim();

    if (codigoBarras) {
      mapaProdutos.set(
        codigoBarras,
        produto
      );
    }
  });

  function produtoEhKg(codigoBarras) {
    const produto =
      mapaProdutos.get(
        String(codigoBarras || '').trim()
      );

    if (!produto) {
      return false;
    }

    const texto = `
      ${produto.descricao || ''}
      ${produto.categoria || ''}
      ${produto.tipo || ''}
    `.toLowerCase();

    return (
      texto.includes('kg') ||
      texto.includes('quilo')
    );
  }

  contagens
    .filter(
      (registro) =>
        registro &&
        registro.ativo !== false
    )
    .forEach((registro) => {
      const usuario =
        registro.usuario ||
        'desconhecido';

      if (!mapa[usuario]) {
        mapa[usuario] = {
          usuario,
          matricula:
            registro.matricula ||
            'SEM-MATRICULA',

          /*
            Mantemos separado:
            unidades físicas e peso.
          */
          totalContadoUn: 0,
          totalContadoKg: 0,

          movimentacoes: 0,
          itensUnicos: new Set(),
          ultimaAtualizacao:
            registro.data || null,
        };
      }

      const quantidade =
        Number(registro.quantidade) || 0;

      if (
        produtoEhKg(
          registro.codigoBarras
        )
      ) {
        mapa[usuario].totalContadoKg +=
          quantidade;
      } else {
        mapa[usuario].totalContadoUn +=
          quantidade;
      }

      mapa[usuario].movimentacoes += 1;

      if (registro.codigoBarras) {
        mapa[usuario].itensUnicos.add(
          String(
            registro.codigoBarras
          ).trim()
        );
      }

      if (registro.data) {
        mapa[usuario].ultimaAtualizacao =
          registro.data;
      }
    });

  const ranking =
    Object.values(mapa).map(
      (item) => ({
        usuario: item.usuario,
        matricula: item.matricula,

        totalContadoUn:
          item.totalContadoUn,

        totalContadoKg:
          item.totalContadoKg,

        /*
          Mantido por compatibilidade com
          partes antigas do sistema.

          Não deve ser usado para exibição
          de UN + KG.
        */
        totalContado:
          item.totalContadoUn,

        movimentacoes:
          item.movimentacoes,

        itensUnicos:
          item.itensUnicos.size,

        ultimaAtualizacao:
          item.ultimaAtualizacao,
      })
    );

  /*
    Como participação atualmente é calculada
    por endereços concluídos na rota
    /ranking-usuarios, não precisamos misturar
    KG e UN para determinar participação.
  */

  ranking.sort(
    (a, b) =>
      b.movimentacoes -
        a.movimentacoes ||
      b.itensUnicos -
        a.itensUnicos
  );

  return ranking;
}

function parseMoeda(valor) {
  if (
    valor === null ||
    valor === undefined ||
    String(valor).trim() === ""
  ) {
    return 0;
  }

  /*
    Usa a mesma conversão do Configurador Universal.

    Formatos aceitos:

    25.90       -> 25.90
    25,90       -> 25.90
    1.234,56    -> 1234.56
    1,234.56    -> 1234.56
    -25,90      -> -25.90
    -25.90      -> -25.90
  */
  const numero =
    converterNumeroUniversal(
      valor
    );

  return Number.isFinite(numero)
    ? numero
    : 0;
}
function parseQuantidade(valor) {
  if (!valor) return 0;

  const texto = String(valor).trim();

  if (texto.includes(".")) {
    return parseFloat(texto) || 0;
  }

  if (texto.includes(",")) {
    return parseFloat(texto.replace(",", ".")) || 0;
  }

  return parseFloat(texto) || 0;
}

function extrairCampoBruto(linha, campo) {
  const cfg = layoutTxt[campo];
  if (!cfg) return "";
  return String(linha || "").substring(cfg.inicio, cfg.fim);
}

function extrairCampoLinha(linha, campo) {
  const bruto = extrairCampoBruto(linha, campo);

  if (campo === "custoUnitario") {
    return parseMoeda(bruto);
  }

  if (campo === "qtdeCongelada") {
    return parseQuantidade(bruto);
  }

  return bruto.trim();
}
async function lerLinhasTxtPorStream(caminhoArquivo) {
  return new Promise((resolve, reject) => {
    const linhas = [];

    const stream = fs.createReadStream(caminhoArquivo, { encoding: "utf8" });

    const rl = readline.createInterface({
      input: stream,
      crlfDelay: Infinity,
    });

    rl.on("line", (linha) => {
      const limpa = String(linha || "").replace(/\r/g, "");
      if (limpa.trim() !== "") {
        linhas.push(limpa);
      }
    });

    rl.on("close", () => resolve(linhas));
    rl.on("error", reject);
    stream.on("error", reject);
  });
}

async function salvarProdutosNoBanco(listaProdutos) {
  const lista = Array.isArray(listaProdutos) ? listaProdutos : [];

  if (usarPostgres) {
    await salvarProdutosPostgres(lista);
    return;
  }

  return new Promise((resolve) => {
    db.serialize(() => {
      db.run("DELETE FROM produtos", (erroDelete) => {
        if (erroDelete) {
          console.error("Erro ao limpar tabela produtos:", erroDelete.message);
          return resolve();
        }

        const stmt = db.prepare(`
          INSERT INTO produtos (
            codigoBarras,
            codigo,
            descricao,
            categoria,
            custoUnitario,
            qtdeCongelada,
            qtdeContada
          ) VALUES (?, ?, ?, ?, ?, ?, ?)
        `);

        for (const item of lista) {
          stmt.run(
            item.codigoBarras || "",
            item.codigo || item.codigoInterno || "",
            item.descricao || "",
            item.categoria || "",
            Number(item.custoUnitario) || 0,
            Number(item.qtdeCongelada) || 0,
            Number(item.qtdeContada) || 0
          );
        }

        stmt.finalize((erroFinalize) => {
          if (erroFinalize) {
            console.error("Erro ao salvar produtos no SQLite:", erroFinalize.message);
          }
          resolve();
        });
      });
    });
  });
}
async function carregarProdutosDoBanco(callback = null) {
  if (usarPostgres) {
    try {
      inventario = await carregarProdutosPostgres();

      console.log(`✅ Inventário carregado do PostgreSQL: ${inventario.length} itens`);

      if (callback) callback();
      return;
    } catch (erro) {
      console.error("Erro ao carregar produtos do PostgreSQL:", erro.message);
      inventario = [];
      if (callback) callback();
      return;
    }
  }

  db.all("SELECT * FROM produtos", [], (err, rows) => {
    if (err) {
      console.error("Erro ao carregar produtos do banco:", err.message);
      inventario = [];
      if (callback) callback();
      return;
    }

    inventario = (rows || []).map((item) => ({
      codigoBarras: item.codigoBarras || "",
      codigo: item.codigo || "",
      codigoInterno: item.codigo || "",
      descricao: item.descricao || "",
      categoria: item.categoria || "",
      custoUnitario: Number(item.custoUnitario) || 0,
      qtdeCongelada: Number(item.qtdeCongelada) || 0,
      qtdeContada: Number(item.qtdeContada) || 0,
    }));

    console.log(`Inventário carregado do SQLite: ${inventario.length} itens`);

    if (callback) callback();
  });
}
app.get("/logo-realstock.png", (req, res) => {
  res.sendFile(path.join(__dirname, "logo-realstock.png"));
});

app.get("/login", (req, res) =>
res.sendFile(caminhoPublico("login.html"))
);

app.post("/login", (req, res) => {
  const { usuario, senha, redirect } = req.body;

  const encontrado = usuarios.find(
    (u) => u.usuario === usuario && u.senha === senha
  );

  if (!encontrado) {
    const destinoErro =
      redirect && String(redirect).trim() !== ""
        ? `/login?erro=Usuário ou senha inválidos&redirect=${encodeURIComponent(
            String(redirect).trim()
          )}`
        : "/login?erro=Usuário ou senha inválidos";

    return res.redirect(destinoErro);
  }
  if (
    String(encontrado.status || "ativo").toLowerCase() === "inativo"
  ) {
    const destinoErro =
      redirect && String(redirect).trim() !== ""
        ? `/login?erro=Usuário inativo. Procure o administrador.&redirect=${encodeURIComponent(
            String(redirect).trim()
          )}`
        : "/login?erro=Usuário inativo. Procure o administrador.";
  
    return res.redirect(destinoErro);
  }

  req.session.logado = true;
  req.session.usuario = {
    id: encontrado.id,
    nome: encontrado.nome,
    usuario: encontrado.usuario,
    funcao: encontrado.funcao,
    status: encontrado.status
  };

  const funcao = String(encontrado.funcao || "").toLowerCase();
let destino = "/";

if (funcao === "operador") {
  destino = "/coleta-mobile";
} else if (
  funcao === "líder" ||
  funcao === "lider" ||
  funcao === "administrador"
) {
  destino = "/";
} else {
  destino = "/login";
}

return res.redirect(destino);
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});
app.get("/quem-sou-eu", autenticar, (req, res) => {
  return res.json({
    usuario: req.session.usuario || null
  });
});
const mobileStatusUsuarios = {};

app.post("/mobile/status", autenticar, (req, res) => {
  const usuarioSessao = req.session.usuario || {};
  const usuario = usuarioSessao.usuario || usuarioSessao.nome || "desconhecido";

  const nome = usuarioSessao.nome || usuario;

const dadosStatus = {
  usuario,
  nome,
  ultimaAtividade: new Date().toISOString()
};

mobileStatusUsuarios[usuario] = dadosStatus;
mobileStatusUsuarios[nome] = dadosStatus;

  res.json({ ok: true });
});
app.get("/", autenticar, permitirSomenteLiderOuAdmin, (req, res) =>
res.sendFile(caminhoPublico("index.html"))
);
app.get("/coleta-mobile", autenticar, permitirSomenteOperador, (req, res) =>
  res.sendFile(caminhoPublico("contagem-mobile.html"))
);

app.get("/modo-operacao", autenticar, (req, res) => {
  return res.json({
    modoOperacao: normalizarModoOperacao(modoOperacao),
  });
});

app.post("/modo-operacao", autenticar, permitirSomenteLiderOuAdmin, async (req, res) => {
  try {
    const novoModo = String(
      req.body?.modoOperacao || req.body?.modo || ""
    ).trim();

    if (!["com-base", "sem-base", "wms"].includes(novoModo)) {
      return res.status(400).json({
        erro: "Modo de operação inválido.",
      });
    }

    modoOperacao = normalizarModoOperacao(novoModo);
    await salvarModoOperacao();

    return res.json({
      sucesso: true,
      modoOperacao,
      mensagem:
  modoOperacao === "sem-base"
    ? "Modo sem base ativado com sucesso."
    : modoOperacao === "wms"
    ? "Modo WMS ativado com sucesso."
    : "Modo com base ativado com sucesso.",
    });
  } catch (erro) {
    console.error("Erro ao salvar modo de operação:", erro);
    return res.status(500).json({
      erro: "Falha ao atualizar modo de operação.",
    });
  }
});

app.get(
  '/wms/base-esperada/produto',
  autenticar,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            'O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.',
        });
      }

      const enderecoWms = String(
        req.query?.enderecoWms || ''
      ).trim();

      const codigoBarras = String(
        req.query?.codigoBarras || ''
      ).trim();

      const codigo = String(
        req.query?.codigo || ''
      ).trim();

      if (!enderecoWms) {
        return res.status(400).json({
          erro:
            'Informe o endereço WMS.',
        });
      }

      if (!codigoBarras && !codigo) {
        return res.status(400).json({
          erro:
            'Informe o código de barras ou o código interno.',
        });
      }

      const item =
        await buscarBaseEsperadaWmsPostgres({
          enderecoWms,
          codigoBarras,
          codigo,
        });

      return res.json({
        sucesso: true,
        encontrado: Boolean(item),
        item,
      });
    } catch (erro) {
      console.error(
        'Erro ao localizar produto na base WMS:',
        erro
      );

      return res.status(500).json({
        erro:
          'Falha ao consultar a base esperada WMS.',
      });
    }
  }
);
app.get(
  "/wms/base-esperada/:id",
  autenticar,
  permitirSomenteLiderOuAdmin,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const item =
        await buscarBaseEsperadaWmsPorIdPostgres(
          req.params.id
        );

      if (!item) {
        return res.status(404).json({
          erro:
            "Registro da base esperada WMS não encontrado.",
        });
      }

      return res.json({
        sucesso: true,
        item,
      });
    } catch (erro) {
      console.error(
        "Erro ao consultar base esperada WMS:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao consultar a base esperada WMS.",
      });
    }
  }
);
app.get(
  "/wms/base-esperada",
  autenticar,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(503).json({
          sucesso: false,
          erro:
            "PostgreSQL não está ativo. Confirme DATABASE_URL e USE_POSTGRES=true.",
          diagnostico: {
            possuiDatabaseUrl,
            usePostgres:
              process.env.USE_POSTGRES ||
              "não informado",
            nodeEnv:
              process.env.NODE_ENV ||
              "development",
          },
        });
      }

      const enderecoWms = String(
        req.query?.enderecoWms || ""
      ).trim();

      const busca = String(
        req.query?.busca || ""
      ).trim();

      const itens =
        await carregarBaseEsperadaWmsPostgres({
          enderecoWms,
          busca,
        });

      return res.json({
        sucesso: true,
        total: Array.isArray(itens)
          ? itens.length
          : 0,
        itens: Array.isArray(itens)
          ? itens
          : [],
      });
    } catch (erro) {
      console.error(
        "Erro ao carregar Base Esperada WMS:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          erro?.message ||
          "Falha ao carregar a Base Esperada WMS.",
      });
    }
  }
);
app.put(
  "/wms/base-esperada/:id",
  autenticar,
  permitirSomenteLiderOuAdmin,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const item =
        await atualizarBaseEsperadaWmsPostgres(
          req.params.id,
          {
            enderecoWms:
              req.body?.enderecoWms,
            codigoBarras:
              req.body?.codigoBarras,
            codigo:
              req.body?.codigo,
            descricao:
              req.body?.descricao,
            quantidadeEsperada:
              req.body?.quantidadeEsperada,
          }
        );

      if (!item) {
        return res.status(404).json({
          erro:
            "Registro da base esperada WMS não encontrado.",
        });
      }

      return res.json({
        sucesso: true,
        item,
        mensagem:
          "Registro da base esperada WMS atualizado.",
      });
    } catch (erro) {
      console.error(
        "Erro ao atualizar base esperada WMS:",
        erro
      );

      const conflito =
        erro?.code === "23505" ||
        String(erro?.message || "").includes(
          "Já existe"
        );

      if (conflito) {
        return res.status(409).json({
          erro:
            "Já existe este produto no endereço WMS informado.",
        });
      }

      return res.status(400).json({
        erro:
          erro?.message ||
          "Falha ao atualizar a base esperada WMS.",
      });
    }
  }
);
app.delete(
  "/wms/base-esperada/:id",
  autenticar,
  permitirSomenteLiderOuAdmin,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const excluido =
        await excluirBaseEsperadaWmsPostgres(
          req.params.id
        );

      if (!excluido) {
        return res.status(404).json({
          erro:
            "Registro da base esperada WMS não encontrado.",
        });
      }

      return res.json({
        sucesso: true,
        mensagem:
          "Registro removido da base esperada WMS.",
      });
    } catch (erro) {
      console.error(
        "Erro ao excluir base esperada WMS:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao excluir o registro da base esperada WMS.",
      });
    }
  }
);

app.get("/wms/contagens", autenticar, async (req, res) => {
  try {
    if (!usarPostgres) {
      return res.status(400).json({
        erro: "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
      });
    }

    const dados = await carregarContagensWmsPostgres();

    return res.json({
      sucesso: true,
      itens: dados,
    });
  } catch (erro) {
    console.error("Erro ao carregar contagens WMS:", erro);
    return res.status(500).json({
      erro: "Falha ao carregar contagens WMS.",
    });
  }
});

app.post("/wms/contagens", autenticar, async (req, res) => {
  try {
    if (!usarPostgres) {
      return res.status(400).json({
        erro: "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
      });
    }

    const usuario =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

      const item = {
        id:
          req.body?.id ||
          `WMS-${Date.now()}-${Math.floor(
            Math.random() * 1000
          )}`,
      
        enderecoInventario: String(
          req.body?.enderecoInventario || ''
        ).trim(),
      
        enderecoWms: String(
          req.body?.enderecoWms || ''
        ).trim(),
      
        codigoBarras: String(
          req.body?.codigoBarras ||
          req.body?.ean ||
          ''
        ).trim(),
      
        codigo: String(
          req.body?.codigo || ''
        ).trim(),
      
        descricao: String(
          req.body?.descricao || ''
        ).trim(),
      
        quantidadeCliente: 0,
      
        quantidadeContada: Number(
          req.body?.quantidadeContada ||
          req.body?.quantidade ||
          0
        ),
      
        usuario,
        data: new Date().toISOString(),
      };

    if (!item.enderecoInventario || !item.enderecoWms) {
      return res.status(400).json({
        erro: "Informe endereço do inventário e endereço WMS.",
      });
    }

    if (!item.codigoBarras && !item.codigo) {
      return res.status(400).json({
        erro: "Informe código de barras ou código interno.",
      });
    }
    const baseEsperada =
  await buscarBaseEsperadaWmsPostgres({
    enderecoWms: item.enderecoWms,
    codigoBarras: item.codigoBarras,
    codigo: item.codigo,
  });

item.quantidadeCliente =
  Number(
    baseEsperada?.quantidadeEsperada
  ) || 0;

    const salvo = await salvarContagemWmsPostgres(item);

    return res.json({
      sucesso: true,
      item: salvo,
      baseEsperadaEncontrada:
        Boolean(baseEsperada),
    
      mensagem: baseEsperada
        ? 'Contagem WMS registrada com sucesso.'
        : 'Contagem registrada, mas o produto não foi encontrado na base esperada deste endereço.',
    });
  } catch (erro) {
    console.error("Erro ao registrar contagem WMS:", erro);
    return res.status(500).json({
      erro: "Falha ao registrar contagem WMS.",
    });
  }
});
app.delete(
  "/wms/contagens/:id",
  autenticar,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const id = String(
        req.params?.id || ""
      ).trim();

      if (!id) {
        return res.status(400).json({
          erro:
            "Identificador da leitura WMS não informado.",
        });
      }

      const usuario =
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "sistema";

      const excluida =
        await excluirContagemWmsPostgres(
          id,
          usuario
        );

      if (!excluida) {
        return res.status(404).json({
          erro:
            "Leitura WMS não encontrada ou já excluída.",
        });
      }

      return res.json({
        sucesso: true,
        mensagem:
          "Leitura WMS excluída com sucesso.",
      });
    } catch (erro) {
      console.error(
        "Erro ao excluir leitura WMS:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao excluir a leitura WMS.",
      });
    }
  }
);
app.post(
  "/wms/finalizar-endereco",
  autenticar,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const enderecoInventario = String(
        req.body?.enderecoInventario || ""
      ).trim();

      const enderecoWms = String(
        req.body?.enderecoWms || ""
      ).trim();

      const itens = Array.isArray(
        req.body?.itens
      )
        ? req.body.itens
        : [];

      if (
        !enderecoInventario ||
        !enderecoWms
      ) {
        return res.status(400).json({
          erro:
            "Informe o endereço do inventário e o endereço WMS.",
        });
      }

      if (!itens.length) {
        return res.status(400).json({
          erro:
            "Não existem itens para finalizar.",
        });
      }

      const usuario =
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "sistema";

      const idsUnicos = new Set(
        itens
          .map((item) =>
            String(
              item.idWms ||
                item.id ||
                item.codigoBarras ||
                item.codigo ||
                ""
            ).trim()
          )
          .filter(Boolean)
      );

      const totalVolume = itens.reduce(
        (total, item) =>
          total +
          (Number(item.quantidade) || 0),
        0
      );

      const finalizacao = {
        id:
          req.body?.id ||
          `FINAL-WMS-${Date.now()}-${Math.floor(
            Math.random() * 1000
          )}`,

        enderecoInventario,
        enderecoWms,
        usuario,
        data: new Date().toISOString(),
        totalItens: idsUnicos.size,
        totalVolume,
      };

      const salva =
      await salvarFinalizacaoWmsPostgres(
        finalizacao
      );
    
    /*
      O endereço WMS pertence à operação WMS,
      mas ranking, progresso e mapa continuam
      sendo controlados pelo endereço do inventário.
    */
    const enderecoInventarioAtualizado =
      await registrarEventoEndereco(
        enderecoInventario,
        "finalizacao",
        usuario
      );
    
    if (!enderecoInventarioAtualizado) {
      console.warn(
        `Finalização WMS salva, porém o endereço de inventário ${enderecoInventario} não foi localizado.`
      );
    }
    
    return res.json({
      sucesso: true,
      finalizacao: salva,
      enderecoInventario:
        enderecoInventarioAtualizado || null,
      mensagem:
        "Endereço WMS e endereço do inventário finalizados com sucesso.",
    });
    } catch (erro) {
      console.error(
        "Erro ao finalizar endereço WMS:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao finalizar o endereço WMS.",
      });
    }
  }
);
app.get(
  "/wms/finalizacoes",
  autenticar,
  async (req, res) => {
    try {
      if (!usarPostgres) {
        return res.status(400).json({
          erro:
            "O módulo WMS usa PostgreSQL. Ative USE_POSTGRES=true.",
        });
      }

      const itens =
        await carregarFinalizacoesWmsPostgres();

      return res.json({
        sucesso: true,
        itens,
      });
    } catch (erro) {
      console.error(
        "Erro ao carregar finalizações WMS:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao carregar finalizações WMS.",
      });
    }
  }
);
app.get("/sem-base/dados", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  return res.json({
    itens: Array.isArray(contagemSemBase) ? contagemSemBase : [],
  });
});
app.post(
  "/sem-base/abrir-endereco",
  autenticar,
  async (req, res) => {
    try {
      if (modoOperacao !== "sem-base") {
        return res.status(400).json({
          erro:
            "Esta operação está disponível somente no modo sem base.",
        });
      }

      const enderecoNumero = String(
        req.body?.enderecoNumero || ""
      ).trim();

      if (!enderecoNumero) {
        return res.status(400).json({
          erro: "Informe o endereço.",
        });
      }

      const usuario =
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "sistema";

      const endereco =
        buscarEnderecoPorNumero(
          enderecoNumero
        );

      if (!endereco) {
        return res.status(404).json({
          erro:
            "Endereço não encontrado no cadastro.",
        });
      }

      if (
        !Array.isArray(
          endereco.transmissoes
        )
      ) {
        endereco.transmissoes = [];
      }

      const eventosDoEndereco =
        endereco.transmissoes
          .filter(
            (evento) =>
              !evento.excluida &&
              Number(
                evento.enderecoNumero
              ) ===
                Number(enderecoNumero)
          )
          .sort(
            (a, b) =>
              new Date(b.data) -
              new Date(a.data)
          );

      const ultimoEvento =
        eventosDoEndereco[0];

      /*
        Evita gerar várias transmissões
        ao clicar repetidamente em abrir.
      */
      if (
        ultimoEvento?.tipo !==
        "transmissao"
      ) {
        await registrarEventoEndereco(
          Number(enderecoNumero),
          "transmissao",
          usuario,
          []
        );
      }

      return res.json({
        sucesso: true,
        mensagem:
          "Endereço aberto para contagem.",
      });
    } catch (erro) {
      console.error(
        "Erro ao abrir endereço sem base:",
        erro
      );

      return res.status(500).json({
        erro:
          "Não foi possível abrir o endereço.",
      });
    }
  }
);
app.post(
  "/sem-base/leitura",
  autenticar,
  async (req, res) => {
    try {
      if (modoOperacao !== "sem-base") {
        return res.status(400).json({
          erro:
            "A leitura sem base somente pode ser usada no modo sem base.",
        });
      }

      const valorInformado = String(
        req.body?.ean ||
          req.body?.codigo ||
          ""
      ).trim();

      if (!valorInformado) {
        return res.status(400).json({
          erro:
            "Informe um EAN ou código interno válido.",
        });
      }

      const enderecoNumero = String(
        req.body?.enderecoNumero || ""
      ).trim();

      if (!enderecoNumero) {
        return res.status(400).json({
          erro:
            "Abra um endereço antes de registrar a leitura.",
        });
      }

      const usuario =
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "sistema";

      const agora =
        new Date().toISOString();

      const pareceEan =
        /^\d{8,14}$/.test(
          valorInformado
        );

      const ean =
        pareceEan
          ? valorInformado
          : "";

      const codigo =
        pareceEan
          ? ""
          : valorInformado;

      const index =
        contagemSemBase.findIndex(
          (item) => {
            const itemEan =
              String(
                item.ean || ""
              ).trim();

            const itemCodigo =
              String(
                item.codigo || ""
              ).trim();

            return (
              itemEan === ean &&
              itemCodigo === codigo
            );
          }
        );

      let itemAtualizado;

      if (index >= 0) {
        const item =
          contagemSemBase[index];

        if (
          !Array.isArray(
            item.enderecos
          )
        ) {
          item.enderecos = [];
        }

        const indiceEndereco =
          item.enderecos.findIndex(
            (registro) =>
              String(
                registro.enderecoNumero ||
                  ""
              ).trim() === enderecoNumero
          );

        if (indiceEndereco >= 0) {
          const registro =
            item.enderecos[
              indiceEndereco
            ];

          registro.quantidade =
            (
              Number(
                registro.quantidade
              ) || 0
            ) + 1;

          registro.usuario =
            usuario;

          registro.data =
            agora;
        } else {
          item.enderecos.push({
            enderecoNumero,
            quantidade: 1,
            usuario,
            data: agora,
          });
        }

        /*
          O total do item sempre será
          a soma das quantidades de todos
          os endereços.
        */
        item.quantidade =
          item.enderecos.reduce(
            (total, registro) =>
              total +
              (
                Number(
                  registro.quantidade
                ) || 0
              ),
            0
          );

        item.ultimoUsuario =
          usuario;

        item.data =
          agora;

        itemAtualizado =
          item;
      } else {
        itemAtualizado = {
          ean,
          codigo,
          quantidade: 1,
          ultimoUsuario:
            usuario,
          data: agora,

          enderecos: [
            {
              enderecoNumero,
              quantidade: 1,
              usuario,
              data: agora,
            },
          ],
        };

        contagemSemBase.push(
          itemAtualizado
        );
      }

      await salvarContagemSemBase();

      return res.json({
        sucesso: true,
        item: itemAtualizado,
        mensagem:
          "Leitura registrada com sucesso.",
      });
    } catch (erro) {
      console.error(
        "Erro em /sem-base/leitura:",
        erro
      );

      return res.status(500).json({
        erro:
          "Falha ao registrar leitura do modo sem base.",
      });
    }
  }
);
app.post("/sem-base/reset", autenticar, permitirSomenteLiderOuAdmin, async (req, res) => {
  try {
    if (modoOperacao !== "sem-base") {
      return res.status(400).json({
        erro:"O reset sem base só pode ser usado no modo sem base."
      });
     }
    contagemSemBase = [];
    await salvarContagemSemBase();

    return res.json({
      sucesso: true,
      mensagem: "Contagem sem base limpa com sucesso.",
    });
  } catch (erro) {
    console.error("Erro em /sem-base/reset:", erro);
    return res.status(500).json({
      erro: "Falha ao limpar dados do modo sem base.",
    });
  }
});
app.get("/sem-base/exportar-txt", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  try {
    const separadorRaw = String(req.query?.separador || ";");
    const cabecalho = String(req.query?.cabecalho || "1") === "1";

    const separador =
      separadorRaw === "tab" ? "\t" : separadorRaw;

    const linhas = [];

    if (cabecalho) {
      linhas.push(["EAN", "Quantidade", "Última leitura"].join(separador));
    }

    contagemSemBase.forEach((item) => {
      linhas.push([
        String(item.ean || ""),
        String(Number(item.quantidade || 0)),
        String(item.ultimaLeituraEm || ""),
      ].join(separador));
    });

    const conteudo = linhas.join("\n");

    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="contagem-sem-base.txt"'
    );

    return res.send(conteudo);
  } catch (erro) {
    console.error("Erro ao exportar TXT sem base:", erro);
    return res.status(500).send("Falha ao exportar TXT.");
  }
});
app.post("/sem-base/finalizar-endereco", autenticar, async (req, res) => {
  try {
    const enderecoNumero = String(req.body?.enderecoNumero || "").trim();

    if (!enderecoNumero) {
      return res.status(400).json({
        erro: "Informe um endereço válido para finalizar.",
      });
    }

    const usuario =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

    const itensDoEndereco = (Array.isArray(contagemSemBase) ? contagemSemBase : [])
      .map((item) => {
        const enderecos = Array.isArray(item.enderecos) ? item.enderecos : [];
        const enderecoAtual = enderecos.find(
          (e) => String(e.enderecoNumero || "") === enderecoNumero
        );

        if (!enderecoAtual) return null;

        return {
          ean: item.ean || "",
          codigo: item.codigo || "",
          quantidade: Number(enderecoAtual.quantidade || 0),
        };
      })
      .filter(Boolean);

    if (!itensDoEndereco.length) {
      return res.status(400).json({
        erro: "Este endereço ainda não possui leituras para finalizar.",
      });
    }

    const finalizacao = {
      id: `SBF-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      enderecoNumero,
      usuario,
      data: new Date().toISOString(),
      itens: itensDoEndereco,
      totalItensUnicos: itensDoEndereco.length,
      totalVolume: itensDoEndereco.reduce(
        (acc, item) => acc + (Number(item.quantidade) || 0),
        0
      ),
    };

    finalizacoesSemBase.push(finalizacao);
    await salvarFinalizacoesSemBase();

    const enderecoAtualizado =
  await registrarEventoEndereco(
    Number(enderecoNumero),
    "finalizacao",
    usuario,
    itensDoEndereco
  );

    return res.json({
      sucesso: true,
      mensagem: "Endereço finalizado e integrado automaticamente ao sistema.",
      finalizacao,
      enderecoAtualizado,
    });
  } catch (erro) {
    console.error("Erro ao finalizar endereço sem base:", erro);
    return res.status(500).json({
      erro: "Falha ao finalizar endereço no modo sem base.",
    });
  }
});

app.get("/sem-base/exportar-pdf", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  try {
    const incluirUltimaLeitura =
      String(req.query?.ultimaLeitura || "1") === "1";

    const printer = new PdfPrinter(fonts);

    const cabecalho = incluirUltimaLeitura
      ? ["EAN", "Quantidade contada", "Última leitura"]
      : ["EAN", "Quantidade contada"];

    const body = [cabecalho];

    contagemSemBase.forEach((item) => {
      const linha = incluirUltimaLeitura
        ? [
            String(item.ean || ""),
            String(Number(item.quantidade || 0)),
            String(item.ultimaLeituraEm || ""),
          ]
        : [
            String(item.ean || ""),
            String(Number(item.quantidade || 0)),
          ];

      body.push(linha);
    });

    const docDefinition = {
      pageSize: "A4",
      pageMargins: [24, 24, 24, 24],
      content: [
        { text: "Relatório - Modo sem base", fontSize: 16, bold: true, margin: [0, 0, 0, 12] },
        {
          table: {
            headerRows: 1,
            widths: incluirUltimaLeitura ? ["*", 100, 160] : ["*", 120],
            body,
          },
          layout: {
            fillColor: function (rowIndex) {
              return rowIndex === 0 ? "#1e293b" : null;
            },
            hLineColor: function () {
              return "#cbd5e1";
            },
            vLineColor: function () {
              return "#cbd5e1";
            },
            hLineWidth: function () {
              return 0.6;
            },
            vLineWidth: function () {
              return 0.6;
            },
          },
        },
      ],
      defaultStyle: {
        font: "Helvetica",
        fontSize: 9,
      },
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="contagem-sem-base.pdf"'
    );

    pdfDoc.pipe(res);
    pdfDoc.end();
  } catch (erro) {
    console.error("Erro ao exportar PDF sem base:", erro);
    return res.status(500).send("Falha ao exportar PDF.");
  }
});
function garantirArquivoContagemSemBase() {
  garantirPastaData();

  if (!fs.existsSync(contagemSemBasePath)) {
    fs.writeFileSync(contagemSemBasePath, "[]", "utf8");
  }
}
function garantirArquivoFinalizacoesSemBase() {
  garantirPastaData();

  if (!fs.existsSync(finalizacoesSemBasePath)) {
    fs.writeFileSync(finalizacoesSemBasePath, "[]", "utf8");
  }
}

async function carregarFinalizacoesSemBase() {
  try {
    garantirArquivoFinalizacoesSemBase();

    if (usarPostgres) {
      const dadosBanco = await carregarFinalizacoesSemBasePostgres();

      if (Array.isArray(dadosBanco)) {
        finalizacoesSemBase = dadosBanco;
        return;
      }
    }

    const bruto = fs.readFileSync(finalizacoesSemBasePath, "utf8") || "[]";
    finalizacoesSemBase = JSON.parse(bruto);

    if (!Array.isArray(finalizacoesSemBase)) {
      finalizacoesSemBase = [];
    }
  } catch (erro) {
    console.error("Erro ao carregar finalizações sem base:", erro);
    finalizacoesSemBase = [];
  }
}

async function salvarFinalizacoesSemBase() {
  try {
    garantirArquivoFinalizacoesSemBase();

    fs.writeFileSync(
      finalizacoesSemBasePath,
      JSON.stringify(finalizacoesSemBase, null, 2),
      "utf8"
    );

    if (usarPostgres) {
      await salvarFinalizacoesSemBasePostgres(finalizacoesSemBase);
    }
  } catch (erro) {
    console.error("Erro ao salvar finalizações sem base:", erro);
  }
}
async function carregarContagemSemBase() {
  try {
    garantirArquivoContagemSemBase();

    if (usarPostgres) {
      const dadosBanco = await carregarContagemSemBasePostgres();

      if (Array.isArray(dadosBanco) && dadosBanco.length > 0) {
        contagemSemBase = dadosBanco;
        return;
      }
    }

    const bruto = fs.readFileSync(contagemSemBasePath, "utf8") || "[]";
    contagemSemBase = JSON.parse(bruto);

    if (!Array.isArray(contagemSemBase)) {
      contagemSemBase = [];
    }

    if (usarPostgres) {
      await salvarContagemSemBasePostgres(contagemSemBase);
    }
  } catch (erro) {
    console.error("Erro ao carregar contagem sem base:", erro.message);
    contagemSemBase = [];
  }
}

async function salvarContagemSemBase() {
  try {
    garantirArquivoContagemSemBase();

    fs.writeFileSync(
      contagemSemBasePath,
      JSON.stringify(contagemSemBase, null, 2),
      "utf8"
    );

    if (usarPostgres) {
      await salvarContagemSemBasePostgres(contagemSemBase);
    }
  } catch (erro) {
    console.error("Erro ao salvar contagem sem base:", erro.message);
  }
}
async function carregarModoOperacao() {
  try {
    garantirPastaData();

    // ===== PRIORIDADE: PostgreSQL =====
    if (usarPostgres) {
      try {
        const config = await carregarConfiguracaoPostgres(
          "modo_operacao",
          null
        );

        if (config) {
          const modoSalvo = String(
            config.modoOperacao || config.modo || "com-base"
          ).trim();

          
          modoOperacao = normalizarModoOperacao(modoSalvo);

          console.log(
            "Modo operacional carregado do PostgreSQL:",
            modoOperacao
          );

          return;
        }
      } catch (erroPg) {
        console.error(
          "Erro ao carregar modo do PostgreSQL:",
          erroPg.message
        );
      }
    }

    // ===== FALLBACK: JSON =====
    if (!fs.existsSync(configModoPath)) {
      modoOperacao = "com-base";
      await salvarModoOperacao();
      return;
    }

    const bruto = fs.readFileSync(configModoPath, "utf8") || "{}";
    const config = JSON.parse(bruto);

    const modoSalvo = String(
      config?.modoOperacao || config?.modo || "com-base"
    ).trim();

    
    modoOperacao = normalizarModoOperacao(modoSalvo);

    console.log("Modo operacional carregado do JSON:", modoOperacao);
  } catch (erro) {
    console.error("Erro ao carregar modo de operação:", erro);
    modoOperacao = "com-base";
  }
}
async function salvarModoOperacao() {
  try {
    garantirPastaData();

    fs.writeFileSync(
      configModoPath,
      JSON.stringify(
        {
          modoOperacao,
          modo: modoOperacao,
          atualizadoEm: new Date().toISOString(),
        },
        null,
        2
      ),
      "utf8"
    );

    if (usarPostgres) {
      await salvarConfiguracaoPostgres("modo_operacao", {
        modoOperacao,
        modo: modoOperacao,
        atualizadoEm: new Date().toISOString(),
      });
    }

    console.log("Modo operacional salvo:", modoOperacao);
  } catch (erro) {
    console.error("Erro ao salvar modo de operação:", erro);
  }
}
function removerDuplicadosPorCodigoInterno(lista) {
  const mapa = new Map();

  lista.forEach((item) => {
    const chave = String(item.codigo || item.codigoInterno || "").trim();

    if (!chave) return;

    if (!mapa.has(chave)) {
      mapa.set(chave, item);
    }
  });

  return Array.from(mapa.values());
}
function normalizarBuscaSemBase(valor) {
  return String(valor || "").trim().toLowerCase();
}

function montarLinhasSemBaseParaTabela({ busca }) {
  let lista = Array.isArray(contagemSemBase) ? [...contagemSemBase] : [];

  if (busca) {
    const termo = normalizarBuscaSemBase(busca);

    lista = lista.filter((item) => {
      const ean = normalizarBuscaSemBase(item.ean || "");
      const codigo = normalizarBuscaSemBase(item.codigo || "");
      return ean.includes(termo) || codigo.includes(termo);
    });
  }

  return lista.map((item) => {
    const enderecos = Array.isArray(item.enderecos) ? item.enderecos : [];

    const enderecoTexto = enderecos
  .map((e) => {
    const enderecoObj = buscarEnderecoPorNumero(Number(e.enderecoNumero));
    const nomeEndereco = enderecoObj?.nome || "ENDEREÇO";
    return `${nomeEndereco} • ${e.enderecoNumero} (${Number(e.quantidade || 0)})`;
  })
  .join(" | ");

    return {
      codigoBarras: item.ean || "",
      codigo: item.codigo || "",
      codigoInterno: item.codigo || "",
      descricao: "",
      categoria: "",
      custoUnitario: 0,
      qtdeCongelada: 0,
      qtdeContada: Number(item.quantidade || 0),
      divergencia: Number(item.quantidade || 0),
      enderecoNumero: enderecos.map((e) => e.enderecoNumero).join(", "),
      endereco: enderecoTexto,
      enderecosContagem: enderecos.map((e) => {
        const enderecoObj = buscarEnderecoPorNumero(Number(e.enderecoNumero));
        const nomeEndereco = enderecoObj?.nome || "ENDEREÇO";
      
        return {
          enderecoNumero: e.enderecoNumero,
          quantidade: Number(e.quantidade || 0),
          texto: `${nomeEndereco} • ${e.enderecoNumero} (${Number(e.quantidade || 0)})`,
        };
      }),
      valorCongelado: 0,
      valorContado: 0,
      valorDivergencia: 0,
      origemRegistro: "sem-base",
    };
  });
}
function filtrarInventario({ categoria, ordem, busca }) {
  let resultado = [...inventario];

  resultado = resultado.filter((item) => parseQuantidade(item.qtdeCongelada) > 0);
  resultado = removerDuplicadosPorCodigoInterno(resultado);

  if (categoria) {
    resultado = resultado.filter((item) =>
      item.categoria?.toLowerCase().includes(categoria.toLowerCase())
    );
  }

  if (busca) {
    const termo = busca.toLowerCase();
    resultado = resultado.filter(
      (item) =>
        item.codigoBarras?.toLowerCase().includes(termo) ||
        item.codigo?.toLowerCase().includes(termo) ||
        item.codigoInterno?.toLowerCase().includes(termo) ||
        item.descricao?.toLowerCase().includes(termo) ||
        item.categoria?.toLowerCase().includes(termo)
    );
  }

  resultado = resultado.map((item) => {
    const resumoContagens = obterResumoContagensDoProduto(item.codigoBarras);
    const qtdeContada = Number(resumoContagens.qtdeContada) || 0;
    const qtdeCongelada = parseQuantidade(item.qtdeCongelada);
    const custo = Number(item.custoUnitario) || 0;
    const divergencia = qtdeContada - qtdeCongelada;
  
    return {
      ...item,
      qtdeContada,
      qtdeCongelada,
      custoUnitario: custo,
      divergencia,
      enderecoNumero: resumoContagens.enderecoNumero || "",
      endereco: resumoContagens.endereco || "",
      enderecosContagem: resumoContagens.enderecos || [],
      valorCongelado: custo * qtdeCongelada,
      valorContado: custo * qtdeContada,
      valorDivergencia: custo * divergencia,
    };
  });
  if (ordem === "divergencia_asc") {
    resultado.sort((a, b) => a.divergencia - b.divergencia);
  } else if (ordem === "divergencia_desc") {
    resultado.sort((a, b) => b.divergencia - a.divergencia);
  } else if (ordem === "maior-divergencia") {
    resultado = resultado.filter((i) => i.divergencia > 0);
    resultado.sort((a, b) => b.divergencia - a.divergencia);
  } else if (ordem === "menor-divergencia") {
    resultado = resultado.filter((i) => i.divergencia < 0);
    resultado.sort((a, b) => a.divergencia - b.divergencia);
  } else if (ordem === "neutra") {
    resultado = resultado.filter((i) => {
      const qtdeContada = Number(i.qtdeContada) || 0;
      const qtdeCongelada = parseQuantidade(i.qtdeCongelada);
      return qtdeContada > 0 && Math.abs(qtdeContada - qtdeCongelada) < 0.000001;
    });
  } else if (ordem === "sem-contagem") {
    resultado = resultado.filter((i) => {
      const qtdeContada = Number(i.qtdeContada) || 0;
      return qtdeContada === 0;
    });
  }

  return resultado;
}
function localizarProdutoResumoOperacional(codigoBarras) {
  const codigoNormalizado = String(
    codigoBarras || ""
  ).trim();

  return inventario.find((produto) => {
    return (
      String(produto.codigoBarras || "").trim() ===
      codigoNormalizado
    );
  });
}
function produtoResumoEhKg(
  codigoBarras = '',
  itemReferencia = null
) {
  const produto =
    localizarProdutoResumoOperacional(
      codigoBarras
    );

  const texto = [
    produto?.descricao,
    produto?.categoria,
    produto?.tipo,

    itemReferencia?.descricao,
    itemReferencia?.categoria,
    itemReferencia?.tipo,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  return (
    texto.includes('kg') ||
    texto.includes('quilo')
  );
}
function calcularVolumesFinalizacao(
  itens = []
) {
  let totalUn = 0;
  let totalKg = 0;

  (
    Array.isArray(itens)
      ? itens
      : []
  ).forEach((item) => {
    const codigoBarras =
      String(
        item?.codigoBarras ||
        item?.ean ||
        ''
      ).trim();

    const quantidade =
      Number(
        item?.quantidade
      ) || 0;

    const ehKg =
      produtoResumoEhKg(
        codigoBarras,
        item
      );

    if (ehKg) {
      totalKg += quantidade;
    } else {
      totalUn += quantidade;
    }
  });

  return {
    totalUn,
    totalKg,
  };
}
function montarUltimasLeiturasDashboard(
  limite = 10
) {
  if (
    modoOperacao === "sem-base"
  ) {
    const leituras = [];

    (
      Array.isArray(
        contagemSemBase
      )
        ? contagemSemBase
        : []
    ).forEach((item) => {
      const enderecos =
        Array.isArray(
          item.enderecos
        )
          ? item.enderecos
          : [];

      enderecos.forEach(
        (endereco) => {
          leituras.push({
            id:
              `${item.ean || item.codigo}-` +
              `${endereco.enderecoNumero}`,

            codigoBarras:
              item.ean || "",

            codigo:
              item.codigo || "",

            descricao:
              item.ean
                ? "Produto sem base"
                : "Código interno sem base",

            quantidade:
              Number(
                endereco.quantidade
              ) || 0,

            usuario:
              endereco.usuario ||
              item.ultimoUsuario ||
              "Não informado",

            enderecoNumero:
              endereco.enderecoNumero ||
              null,

              data:
              endereco.data ||
              item.data ||
              item.ultimaLeituraEm ||
              null,
          });
        }
      );
    });

    return leituras
      .filter(
        (item) => item.data
      )
      .sort(
        (a, b) =>
          new Date(b.data) -
          new Date(a.data)
      )
      .slice(0, limite);
  }

  return [
    ...(
      Array.isArray(contagens)
        ? contagens
        : []
    ),
  ]
    .filter(
      (item) =>
        item &&
        item.ativo !== false &&
        item.codigoBarras &&
        item.data
    )
    .sort(
      (a, b) =>
        new Date(b.data) -
        new Date(a.data)
    )
    .slice(0, limite)
    .map((item) => {
      const produto =
        localizarProdutoResumoOperacional(
          item.codigoBarras
        );

        const ehKg =
        produtoResumoEhKg(
          item.codigoBarras,
          produto
        );
      
      return {
        id:
          item.id || '',
      
        codigoBarras:
          item.codigoBarras || '',
      
        codigo:
          produto?.codigo ||
          produto?.codigoInterno ||
          '',
      
        descricao:
          produto?.descricao ||
          'Sem descrição',
      
        categoria:
          produto?.categoria || '',
      
        tipo:
          produto?.tipo || '',
      
        quantidade:
          Number(
            item.quantidade
          ) || 0,
      
        unidade:
          ehKg
            ? 'kg'
            : 'un',
      
        ehKg,
      
        usuario:
          item.usuario ||
          'Não informado',
      
        enderecoNumero:
          item.enderecoNumero ??
          null,
      
        data:
          item.data,
      };
    });
}

function montarUltimasFinalizacoesDashboard(
  limite = 10
) {
  const modoAtual =
    normalizarModoOperacao(
      modoOperacao
    );

  /*
    ========================================================
    MODO SEM BASE
    ========================================================
  */

  if (modoAtual === "sem-base") {
    return [
      ...(
        Array.isArray(
          finalizacoesSemBase
        )
          ? finalizacoesSemBase
          : []
      ),
    ]
      .filter(
        (item) =>
          item &&
          !item.excluida &&
          item.data
      )
      .sort(
        (a, b) =>
          new Date(b.data) -
          new Date(a.data)
      )
      .slice(0, limite)
      .map((finalizacao) => {
        const itens =
          Array.isArray(
            finalizacao.itens
          )
            ? finalizacao.itens
            : [];

        const codigosUnicos =
          new Set(
            itens
              .map(
                (item) =>
                  String(
                    item.ean ||
                    item.codigoBarras ||
                    item.codigo ||
                    ""
                  ).trim()
              )
              .filter(Boolean)
          );

          const volumes =
  calcularVolumesFinalizacao(
    itens
  );

        return {
          id:
            finalizacao.id || "",

          enderecoId: 0,

          enderecoNumero:
            String(
              finalizacao.enderecoNumero ||
              ""
            ),

          enderecoNome:
            "Endereço",

          usuario:
            finalizacao.usuario ||
            "Não informado",

          totalItensUnicos:
            codigosUnicos.size,

            totalVolumeUn:
            volumes.totalUn,
          
          totalVolumeKg:
            volumes.totalKg,
          data:
            finalizacao.data,
        };
      });
  }


  /*
    ========================================================
    COM BASE / WMS

    Aqui filtramos PRIMEIRO os endereços do modo atual
    e DEPOIS as finalizações do mesmo modo.
    ========================================================
  */

  const lista = [];

  const enderecosDoModoAtual = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).filter(
    (endereco) =>
      normalizarModoOperacao(
        endereco?.modoOperacao
      ) === modoAtual
  );


  enderecosDoModoAtual.forEach(
    (endereco) => {

      const finalizacoes =
        Array.isArray(
          endereco?.finalizacoes
        )
          ? endereco.finalizacoes.filter(
              (finalizacao) =>
                finalizacao &&
                !finalizacao.excluida &&
                finalizacao.data &&
                eventoPertenceAoModo(
                  finalizacao,
                  modoAtual
                )
            )
          : [];


      finalizacoes.forEach(
        (finalizacao) => {

          const itens =
            Array.isArray(
              finalizacao.itens
            )
              ? finalizacao.itens
              : [];

              const volumes =
              calcularVolumesFinalizacao(
                itens
              );


          lista.push({
            id:
              finalizacao.id || "",

            enderecoId:
              Number(
                finalizacao.enderecoId
              ) ||
              Number(
                endereco.id
              ) ||
              0,

            enderecoNumero:
              Number(
                finalizacao.enderecoNumero
              ) || 0,

            enderecoNome:
              endereco.nome ||
              endereco.tipo ||
              "Endereço",

            usuario:
              finalizacao.usuario ||
              "Não informado",

            totalItensUnicos:
              new Set(
                itens
                  .map(
                    (item) =>
                      String(
                        item.codigoBarras ||
                        item.ean ||
                        item.codigo ||
                        ""
                      ).trim()
                  )
                  .filter(Boolean)
              ).size,

              totalVolumeUn:
              volumes.totalUn,
            
            totalVolumeKg:
              volumes.totalKg,
            data:
              finalizacao.data,
          });
        }
      );
    }
  );


  return lista
    .sort(
      (a, b) =>
        new Date(b.data) -
        new Date(a.data)
    )
    .slice(0, limite);
}

function montarMapaEnderecosDashboard() {
  const resultado = [];

  const modoAtual = normalizarModoOperacao(
    modoOperacao
  );

  const pertenceAoModoAtual = (evento) =>
    eventoPertenceAoModo(
      evento,
      modoAtual
    );

  const enderecosDoModoAtual = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).filter(
    (endereco) =>
      normalizarModoOperacao(
        endereco?.modoOperacao
      ) === modoAtual
  );

  enderecosDoModoAtual.forEach((endereco) => {

    const inicio = Number(endereco.inicio) || 0;
    const fim = Number(endereco.fim) || 0;

    const finalizacoes =
    Array.isArray(
      endereco.finalizacoes
    )
      ? endereco.finalizacoes.filter(
          (item) =>
            !item.excluida &&
            pertenceAoModoAtual(
              item
            )
        )
      : [];
      const transmissoes =
      Array.isArray(
        endereco.transmissoes
      )
        ? endereco.transmissoes.filter(
            (item) =>
              !item.excluida &&
              pertenceAoModoAtual(
                item
              )
          )
        : [];
    for (
      let numero = inicio;
      numero <= fim;
      numero += 1
    ) {
      const finalizacoesNumero =
        finalizacoes.filter(
          (item) =>
            Number(item.enderecoNumero) === numero
        );

      const transmissoesNumero =
        transmissoes.filter(
          (item) =>
            Number(item.enderecoNumero) === numero &&
            item.tipo === "transmissao"
        );

        let status = "pendente";

        if (finalizacoesNumero.length > 1) {
          status = "duplicado";
        } else if (finalizacoesNumero.length === 1) {
          status = "concluido";
        } else if (transmissoesNumero.length > 0) {
          status = "em-contagem";
        }
        
        
        /*
          =====================================================
          OPERADOR RESPONSÁVEL PELO ENDEREÇO
          =====================================================
        
          CONCLUÍDO:
          deve mostrar quem FINALIZOU a contagem.
        
          EM CONTAGEM:
          mostra quem realizou a transmissão mais recente.
        
          DUPLICADO:
          usa a finalização mais recente.
        
          Nunca usar uma consolidação posterior feita
          pelo admin para substituir o conferente real.
        */
        
        const ultimaFinalizacao =
          [...finalizacoesNumero]
            .filter((item) => item?.data)
            .sort(
              (a, b) =>
                new Date(b.data).getTime() -
                new Date(a.data).getTime()
            )[0] || null;
        
        
        const ultimaTransmissao =
          [...transmissoesNumero]
            .filter((item) => item?.data)
            .sort(
              (a, b) =>
                new Date(b.data).getTime() -
                new Date(a.data).getTime()
            )[0] || null;
        
        
        let eventoResponsavel = null;
        
        if (
          status === "concluido" ||
          status === "duplicado"
        ) {
          eventoResponsavel =
            ultimaFinalizacao;
        }
        else if (
          status === "em-contagem"
        ) {
          eventoResponsavel =
            ultimaTransmissao;
        }

        resultado.push({
          enderecoId:
            Number(endereco.id) || 0,
        
          enderecoNumero:
            numero,
        
          setor:
            endereco.nome ||
            endereco.tipo ||
            "Endereço",
        
          status,
        
          usuario:
            eventoResponsavel?.usuario ||
            "Não informado",
        
          ultimaAtividade:
            eventoResponsavel?.data ||
            null,
        });
    }
  });

  return resultado;
}

app.get(
  "/api/dashboard-operacional",
  autenticar,
  permitirSomenteLiderOuAdmin,
  (req, res) => {
    try {
      res.set({
        "Cache-Control":
          "no-store, no-cache, must-revalidate, private",
        Pragma: "no-cache",
        Expires: "0",
      });

      return res.json({
        sucesso: true,
        ultimasLeituras:
          montarUltimasLeiturasDashboard(10),
        ultimasFinalizacoes:
          montarUltimasFinalizacoesDashboard(10),
        mapaEnderecos:
          montarMapaEnderecosDashboard(),
        atualizadoEm: new Date().toISOString(),
      });
    } catch (erro) {
      console.error(
        "Erro ao montar dashboard operacional:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          "Falha ao carregar o resumo operacional.",
      });
    }
  }
);
app.get("/inventario", autenticar, (req, res) => {
  if (modoOperacao === "sem-base") {
    const resultado = montarLinhasSemBaseParaTabela({
      busca: req.query.busca || "",
    });

    return res.json(resultado);
  }

  const resultado = filtrarInventario({
    categoria: req.query.categoria || "",
    ordem: req.query.ordem || "",
    busca: req.query.busca || "",
  });

  res.json(resultado);
});

function normalizarChaveUsuarioRanking(valor) {
  return String(valor || "")
    .trim()
    .toLowerCase();
}
function gerarMapaEnderecosConcluidosPorUsuario() {
  const mapa = new Map();

 
  function adicionarEndereco(usuario, chaveEndereco) {
    const chaveUsuario = normalizarChaveUsuarioRanking(usuario);
    const chave = String(chaveEndereco || "").trim();

    if (!chaveUsuario || !chave) {
      return;
    }

    if (!mapa.has(chaveUsuario)) {
      mapa.set(chaveUsuario, new Set());
    }

    mapa.get(chaveUsuario).add(chave);
  }

  if (modoOperacao === "sem-base") {
    const finalizacoes = Array.isArray(finalizacoesSemBase)
      ? finalizacoesSemBase
      : [];

    finalizacoes.forEach((finalizacao) => {
      if (!finalizacao || finalizacao.excluida) {
        return;
      }

      const enderecoNumero = String(
        finalizacao.enderecoNumero || ""
      ).trim();

      adicionarEndereco(
        finalizacao.usuario,
        `SEM-BASE:${enderecoNumero}`
      );
    });

    return mapa;
  }

  const listaEnderecamentos = Array.isArray(enderecamentos)
    ? enderecamentos
    : [];

    listaEnderecamentos.forEach((endereco) => {
      const finalizacoes =
        Array.isArray(
          endereco?.finalizacoes
        )
          ? endereco.finalizacoes.filter(
              (finalizacao) =>
                !finalizacao?.excluida &&
                eventoPertenceAoModo(
                  finalizacao
                )
            )
          : [];

    finalizacoes.forEach((finalizacao) => {
      if (!finalizacao || finalizacao.excluida) {
        return;
      }

      const enderecoId =
        Number(finalizacao.enderecoId) ||
        Number(endereco.id) ||
        0;

      const enderecoNumero =
        Number(finalizacao.enderecoNumero) || 0;

      if (!enderecoNumero) {
        return;
      }

      adicionarEndereco(
        finalizacao.usuario,
        `${enderecoId}:${enderecoNumero}`
      );
    });
  });

  return mapa;
}

function montarDetalhesProgresso() {
  const agora = Date.now();

  const modoAtual =
    normalizarModoOperacao(
      modoOperacao
    );

  const listaEnderecamentos = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  ).filter(
    (endereco) =>
      normalizarModoOperacao(
        endereco?.modoOperacao
      ) === modoAtual
  );
  const finalizacoesValidas = [];


/*
  ========================================================
  SEM BASE

  Usa finalizacoesSemBase porque nela estão:
  - operador real
  - produtos
  - quantidades
  - volume
  ========================================================
*/
if (modoAtual === "sem-base") {

  (
    Array.isArray(
      finalizacoesSemBase
    )
      ? finalizacoesSemBase
      : []
  )
    .filter(
      (finalizacao) =>
        finalizacao &&
        !finalizacao.excluida
    )
    .forEach(
      (finalizacao) => {

        const enderecoNumero =
          Number(
            finalizacao.enderecoNumero
          ) || 0;


        if (!enderecoNumero) {
          return;
        }


        const endereco =
          buscarEnderecoPorNumero(
            enderecoNumero,
            modoAtual
          );


        if (!endereco) {
          return;
        }


        const itens =
          Array.isArray(
            finalizacao.itens
          )
            ? finalizacao.itens.map(
                (item) => ({
                  ...item,

                  codigoBarras:
                    item.codigoBarras ||
                    item.ean ||
                    "",

                  codigo:
                    item.codigo ||
                    "",

                  quantidade:
                    Number(
                      item.quantidade
                    ) || 0,
                })
              )
            : [];


        finalizacoesValidas.push({
          id:
            finalizacao.id || "",

          enderecoId:
            Number(
              endereco.id
            ) || 0,

          enderecoNumero,

          enderecoNome:
            endereco.nome ||
            endereco.tipo ||
            "Endereço",

          usuario:
            finalizacao.usuario ||
            "Não informado",

          data:
            finalizacao.data ||
            null,

          itens,
        });
      }
    );


/*
  ========================================================
  COM BASE / WMS

  Mantém a fonte normal de endereçamentos.
  ========================================================
*/
} else {

  listaEnderecamentos.forEach(
    (endereco) => {

      const finalizacoes =
        Array.isArray(
          endereco?.finalizacoes
        )
          ? endereco.finalizacoes.filter(
              (finalizacao) =>
                !finalizacao?.excluida &&
                eventoPertenceAoModo(
                  finalizacao,
                  modoAtual
                )
            )
          : [];


      finalizacoes.forEach(
        (finalizacao) => {

          if (
            !finalizacao ||
            finalizacao.excluida
          ) {
            return;
          }


          finalizacoesValidas.push({
            id:
              finalizacao.id || "",

            enderecoId:
              Number(
                finalizacao.enderecoId
              ) ||
              Number(
                endereco.id
              ) ||
              0,

            enderecoNumero:
              Number(
                finalizacao.enderecoNumero
              ) || 0,

            enderecoNome:
              endereco.nome ||
              endereco.tipo ||
              "Endereço",

            usuario:
              finalizacao.usuario ||
              "Não informado",

            data:
              finalizacao.data ||
              null,

            itens:
              Array.isArray(
                finalizacao.itens
              )
                ? finalizacao.itens
                : [],
          });
        }
      );
    }
  );
}
  const chavesConcluidas = new Set(
    finalizacoesValidas
      .filter((item) => item.enderecoNumero)
      .map(
        (item) =>
          `${item.enderecoId}:${item.enderecoNumero}`
      )
  );

  const totalPosicoes = listaEnderecamentos.reduce(
    (total, endereco) => {
      const inicio = Number(endereco?.inicio) || 0;
      const fim = Number(endereco?.fim) || 0;

      if (!inicio || fim < inicio) {
        return total;
      }

      return total + (fim - inicio + 1);
    },
    0
  );

  const concluidos = chavesConcluidas.size;
  const pendentes = Math.max(0, totalPosicoes - concluidos);

  const operadoresAtivos = Object.entries(
    mobileStatusUsuarios || {}
  )
    .filter(([, status]) => {
      if (!status?.ultimaAtividade) return false;

      const ultimaAtividade = new Date(
        status.ultimaAtividade
      ).getTime();

      return (
        Number.isFinite(ultimaAtividade) &&
        agora - ultimaAtividade <= 30000
      );
    })
    .map(([usuario, status]) => ({
      usuario,
      ultimaAtividade:
        status.ultimaAtividade || null,
    }));

  const temposPorEndereco = [];

  finalizacoesValidas.forEach((finalizacao) => {
    const endereco = listaEnderecamentos.find(
      (item) =>
        Number(item.id) ===
        Number(finalizacao.enderecoId)
    );

    if (!endereco || !finalizacao.data) return;

    const eventos = Array.isArray(
      endereco.transmissoes
    )
      ? endereco.transmissoes.filter(
          (evento) =>
            !evento?.excluida &&
            eventoPertenceAoModo(
              evento
            )
        )
      : [];

    const transmissoesDoNumero = eventos
      .filter(
        (evento) =>
          evento?.tipo === "transmissao" &&
          Number(evento.enderecoNumero) ===
            Number(finalizacao.enderecoNumero) &&
          evento.data
      )
      .sort(
        (a, b) =>
          new Date(a.data).getTime() -
          new Date(b.data).getTime()
      );

    const primeiraTransmissao =
      transmissoesDoNumero[0];

    if (!primeiraTransmissao) return;

    const inicio = new Date(
      primeiraTransmissao.data
    ).getTime();

    const fim = new Date(
      finalizacao.data
    ).getTime();

    const duracao = fim - inicio;

    if (
      Number.isFinite(duracao) &&
      duracao > 0
    ) {
      temposPorEndereco.push(duracao);
    }
  });

  const tempoMedioMs = temposPorEndereco.length
    ? temposPorEndereco.reduce(
        (total, valor) => total + valor,
        0
      ) / temposPorEndereco.length
    : 0;

  const umaHoraAtras = agora - 60 * 60 * 1000;

  const finalizadosUltimaHora =
    finalizacoesValidas.filter((item) => {
      const horario = new Date(item.data).getTime();

      return (
        Number.isFinite(horario) &&
        horario >= umaHoraAtras &&
        horario <= agora
      );
    }).length;

  const ritmoPorHora = finalizadosUltimaHora;

  let previsaoConclusao = null;

  if (pendentes === 0 && concluidos > 0) {
    previsaoConclusao = new Date().toISOString();
  } else if (ritmoPorHora > 0 && pendentes > 0) {
    const horasRestantes =
      pendentes / ritmoPorHora;

    previsaoConclusao = new Date(
      agora +
        horasRestantes *
          60 *
          60 *
          1000
    ).toISOString();
  }

  const ultimasFinalizacoes = [
    ...finalizacoesValidas,
  ]
    .filter((item) => item.data)
    .sort(
      (a, b) =>
        new Date(b.data).getTime() -
        new Date(a.data).getTime()
    )
    .slice(0, 10)
    .map((item) => ({
      id: item.id,
      enderecoId: item.enderecoId,
      enderecoNumero: item.enderecoNumero,
      enderecoNome: item.enderecoNome,
      usuario: item.usuario,
      data: item.data,
      totalItensUnicos: new Set(
        item.itens
          .map(
            (produto) =>
              produto.codigoBarras ||
              produto.codigo ||
              produto.contagemId ||
              ""
          )
          .filter(Boolean)
      ).size,
      totalVolume: item.itens.reduce(
        (total, produto) =>
          total +
          (Number(produto.quantidade) || 0),
        0
      ),
    }));

  const percentual = totalPosicoes
    ? (concluidos / totalPosicoes) * 100
    : 0;

  return {
    totalPosicoes,
    concluidos,
    pendentes,
    percentual,
    tempoMedioMs,
    ritmoPorHora,
    previsaoConclusao,
    operadoresAtivos,
    totalOperadoresAtivos:
      operadoresAtivos.length,
    ultimasFinalizacoes,
    atualizadoEm: new Date().toISOString(),
  };
}

app.get(
  "/api/progresso-detalhes",
  autenticar,
  permitirSomenteLiderOuAdmin,
  (req, res) => {
    try {
      res.set({
        "Cache-Control":
          "no-store, no-cache, must-revalidate, private",
        Pragma: "no-cache",
        Expires: "0",
      });

      const detalhes =
        montarDetalhesProgresso();

      return res.status(200).json({
        sucesso: true,
        detalhes,
      });
    } catch (erro) {
      console.error(
        "Erro ao montar detalhes do progresso:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          "Falha ao carregar os detalhes do progresso.",
      });
    }
  }
);

app.get("/ranking-usuarios", autenticar, (req, res) => {
  const ranking =
    modoOperacao === "sem-base"
      ? gerarRankingUsuariosSemBase()
      : gerarRankingUsuarios();

  const agora = Date.now();

  const mapaEnderecosConcluidos =
    gerarMapaEnderecosConcluidosPorUsuario();


    const modoAtualRanking =
  normalizarModoOperacao(
    modoOperacao
  );

const totalEnderecosCadastrados = (
  Array.isArray(enderecamentos)
    ? enderecamentos
    : []
)
  .filter(
    (endereco) =>
      normalizarModoOperacao(
        endereco?.modoOperacao
      ) === modoAtualRanking
  )
  .reduce(
    (total, endereco) => {
      const inicio =
        Number(endereco?.inicio) || 0;

      const fim =
        Number(endereco?.fim) || 0;

      if (
        !inicio ||
        !fim ||
        fim < inicio
      ) {
        return total;
      }

      return (
        total +
        (fim - inicio + 1)
      );
    },
    0
  );

  const rankingComStatus = ranking.map((item) => {
    const possiveisChaves = [
      item.usuario,
      item.nome,
      item.usuarioNome,
      item.nomeUsuario,
    ].filter(Boolean);

    const status = possiveisChaves
      .map((chave) => mobileStatusUsuarios[chave])
      .find(Boolean);

    const online =
      status &&
      agora -
        new Date(status.ultimaAtividade).getTime() <=
        30000;

    const chaveUsuario =
      possiveisChaves
        .map(normalizarChaveUsuarioRanking)
        .find((chave) =>
          mapaEnderecosConcluidos.has(chave)
        ) ||
      normalizarChaveUsuarioRanking(item.usuario);

    const enderecosConcluidos =
      mapaEnderecosConcluidos.get(chaveUsuario)?.size || 0;

      return {
        ...item,
      
        enderecosConcluidos,
      
        /*
          Participação = endereços concluídos
          pelo operador / total cadastrado
          no modo atual.
        */
        percentual:
          totalEnderecosCadastrados > 0
            ? (
                enderecosConcluidos /
                totalEnderecosCadastrados
              ) * 100
            : 0,
      
        mobileOnline:
          !!online,
      
          ultimaAtividadeMobile:
          status?.ultimaAtividade || null,
      };
  });

  rankingComStatus.sort((a, b) => {
    return (
      Number(b.enderecosConcluidos || 0) -
        Number(a.enderecosConcluidos || 0) ||
      Number(b.totalContado || 0) -
        Number(a.totalContado || 0)
    );
  });

  res.json(rankingComStatus);
});
app.get("/sem-base/resumo-dashboard", autenticar, (req, res) => {
  return res.json(gerarResumoDashboardSemBase());
});
app.get("/layout-txt", autenticar, (req, res) => {
  res.json(layoutTxt);
});

app.post("/layout-txt", autenticar, (req, res) => {
  try {
    const recebido = req.body || {};

    layoutTxt = {
      codigoBarras: recebido.codigoBarras || layoutTxt.codigoBarras,
      codigo: recebido.codigo || layoutTxt.codigo,
      descricao: recebido.descricao || layoutTxt.descricao,
      custoUnitario: recebido.custoUnitario || layoutTxt.custoUnitario,
      qtdeCongelada: recebido.qtdeCongelada || layoutTxt.qtdeCongelada,
      categoria: recebido.categoria || layoutTxt.categoria,
      tipo: recebido.tipo || layoutTxt.tipo,
    };

    salvarLayoutTxt();
    res.json({ sucesso: true, layoutTxt });
  } catch (erro) {
    console.error("Erro ao salvar layout TXT:", erro);
    res.status(500).json({ erro: "Falha ao salvar layout TXT." });
  }
});
/* ==========================================================
   LAYOUTS UNIVERSAIS
========================================================== */

app.get(
  "/configurador-universal/layouts",
  autenticar,
  (req, res) => {
    const ordenados = [
      ...layoutsUniversais,
    ].sort((a, b) => {
      return String(a.nome || "")
        .localeCompare(
          String(b.nome || ""),
          "pt-BR",
          {
            sensitivity: "base",
          }
        );
    });

    return res.json({
      sucesso: true,
      total: ordenados.length,
      layouts: ordenados,
    });
  }
);

app.post(
  "/configurador-universal/layouts",
  autenticar,
  async (req, res) => {
    try {
      const nome = String(
        req.body?.nome || ""
      ).trim();

      if (!nome) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Informe o nome do layout.",
        });
      }

      const layout =
        normalizarLayoutUniversal(
          req.body
        );

      layoutsUniversais.push(layout);
      salvarLayoutsUniversais();

      return res.status(201).json({
        sucesso: true,
        mensagem:
          "Layout universal salvo com sucesso.",
        layout,
      });
    } catch (erro) {
      console.error(
        "Erro ao salvar layout universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          "Não foi possível salvar o layout universal.",
      });
    }
  }
);

app.put(
  "/configurador-universal/layouts/:id",
  autenticar,
  async (req, res) => {
    try {
      const id = String(
        req.params.id || ""
      ).trim();

      const indice =
        layoutsUniversais.findIndex(
          (item) => item.id === id
        );

      if (indice === -1) {
        return res.status(404).json({
          sucesso: false,
          erro:
            "Layout universal não encontrado.",
        });
      }

      const atual =
        layoutsUniversais[indice];

      const atualizado =
        normalizarLayoutUniversal(
          req.body,
          atual
        );

      if (!atualizado.nome) {
        return res.status(400).json({
          sucesso: false,
          erro:
            "Informe o nome do layout.",
        });
      }

      layoutsUniversais[indice] =
        atualizado;

      salvarLayoutsUniversais();

      return res.json({
        sucesso: true,
        mensagem:
          "Layout universal atualizado.",
        layout: atualizado,
      });
    } catch (erro) {
      console.error(
        "Erro ao atualizar layout universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          "Não foi possível atualizar o layout universal.",
      });
    }
  }
);

app.delete(
  "/configurador-universal/layouts/:id",
  autenticar,
  (req, res) => {
    try {
      const id = String(
        req.params.id || ""
      ).trim();

      const indice =
        layoutsUniversais.findIndex(
          (item) => item.id === id
        );

      if (indice === -1) {
        return res.status(404).json({
          sucesso: false,
          erro:
            "Layout universal não encontrado.",
        });
      }

      layoutsUniversais.splice(
        indice,
        1
      );

      salvarLayoutsUniversais();

      return res.json({
        sucesso: true,
        mensagem:
          "Layout universal excluído.",
      });
    } catch (erro) {
      console.error(
        "Erro ao excluir layout universal:",
        erro
      );

      return res.status(500).json({
        sucesso: false,
        erro:
          "Não foi possível excluir o layout universal.",
      });
    }
  }
);
app.get("/layouts-txt", autenticar, (req, res) => {
  res.json(layoutsSalvos);
});

app.post("/layouts-txt", autenticar, (req, res) => {
  try {
    const { nome, cliente, observacao } = req.body || {};

    if (!nome || !String(nome).trim()) {
      return res.status(400).json({ erro: "Nome do layout é obrigatório." });
    }

    const novoLayout = {
      id: gerarIdLayout(),
      nome: String(nome).trim(),
      cliente: String(cliente || "").trim(),
      observacao: String(observacao || "").trim(),
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      campos: JSON.parse(JSON.stringify(layoutTxt)),
    };

    layoutsSalvos.push(novoLayout);
    salvarLayoutsSalvos();

    res.json({ sucesso: true, layout: novoLayout });
  } catch (erro) {
    console.error("Erro ao criar layout salvo:", erro);
    res.status(500).json({ erro: "Falha ao criar layout salvo." });
  }
});

app.post("/layouts-txt/:id/aplicar", autenticar, (req, res) => {
  try {
    const { id } = req.params;
    const layout = layoutsSalvos.find((item) => item.id === id);

    if (!layout) {
      return res.status(404).json({ erro: "Layout não encontrado." });
    }

    layoutTxt = JSON.parse(JSON.stringify(layout.campos || layoutTxtPadrao));
    salvarLayoutTxt();

    layout.ultimoUsoEm = new Date().toISOString();
    layout.atualizadoEm = new Date().toISOString();
    salvarLayoutsSalvos();

    res.json({ sucesso: true, layoutTxt });
  } catch (erro) {
    console.error("Erro ao aplicar layout salvo:", erro);
    res.status(500).json({ erro: "Falha ao aplicar layout salvo." });
  }
});

app.delete("/layouts-txt/:id", autenticar, (req, res) => {
  try {
    const { id } = req.params;
    const index = layoutsSalvos.findIndex((item) => item.id === id);

    if (index === -1) {
      return res.status(404).json({ erro: "Layout não encontrado." });
    }

    layoutsSalvos.splice(index, 1);
    salvarLayoutsSalvos();

    res.json({ sucesso: true });
  } catch (erro) {
    console.error("Erro ao excluir layout salvo:", erro);
    res.status(500).json({ erro: "Falha ao excluir layout salvo." });
  }
});

app.get("/exportacao-universal/colunas", autenticar, (req, res) => {
  res.json({
    colunas: COLUNAS_EXPORTACAO_DISPONIVEIS,
    formatos: Array.from(FORMATOS_EXPORTACAO_PERMITIDOS),
  });
});

app.get("/exportacao-universal/layouts", autenticar, (req, res) => {
  const cliente = String(req.query.cliente || "").trim().toLowerCase();
  const lista = cliente
    ? layoutsExportacao.filter(
        (l) => String(l.cliente || "").trim().toLowerCase() === cliente
      )
    : layoutsExportacao;
  res.json(lista);
});

app.post("/exportacao-universal/layouts", autenticar, (req, res) => {
  try {
    const layout = normalizarLayoutExportacao(req.body || {});

    if (!layout.nome) {
      return res.status(400).json({
        sucesso: false,
        erro: "Informe o nome do layout de exportação.",
      });
    }

    if (!layout.colunas.length) {
      return res.status(400).json({
        sucesso: false,
        erro: "Selecione ao menos uma coluna.",
      });
    }

    layoutsExportacao.push(layout);
    salvarLayoutsExportacao();

    return res.status(201).json({
      sucesso: true,
      mensagem: "Layout de exportação salvo com sucesso.",
      layout,
    });
  } catch (erro) {
    console.error("Erro ao salvar layout de exportação:", erro);
    return res.status(500).json({
      sucesso: false,
      erro: "Não foi possível salvar o layout de exportação.",
    });
  }
});

app.put("/exportacao-universal/layouts/:id", autenticar, (req, res) => {
  try {
    const id = String(req.params.id || "").trim();
    const indice = layoutsExportacao.findIndex((item) => item.id === id);

    if (indice === -1) {
      return res.status(404).json({
        sucesso: false,
        erro: "Layout de exportação não encontrado.",
      });
    }

    const atualizado = normalizarLayoutExportacao(
      req.body || {},
      layoutsExportacao[indice]
    );

    if (!atualizado.nome) {
      return res.status(400).json({
        sucesso: false,
        erro: "Informe o nome do layout.",
      });
    }

    if (!atualizado.colunas.length) {
      return res.status(400).json({
        sucesso: false,
        erro: "Selecione ao menos uma coluna.",
      });
    }

    layoutsExportacao[indice] = atualizado;
    salvarLayoutsExportacao();

    return res.json({
      sucesso: true,
      mensagem: "Layout de exportação atualizado.",
      layout: atualizado,
    });
  } catch (erro) {
    console.error("Erro ao atualizar layout de exportação:", erro);
    return res.status(500).json({
      sucesso: false,
      erro: "Não foi possível atualizar o layout de exportação.",
    });
  }
});

app.delete("/exportacao-universal/layouts/:id", autenticar, (req, res) => {
  try {
    const id = String(req.params.id || "").trim();
    const indice = layoutsExportacao.findIndex((item) => item.id === id);

    if (indice === -1) {
      return res.status(404).json({
        sucesso: false,
        erro: "Layout de exportação não encontrado.",
      });
    }

    layoutsExportacao.splice(indice, 1);
    salvarLayoutsExportacao();

    return res.json({
      sucesso: true,
      mensagem: "Layout de exportação excluído.",
    });
  } catch (erro) {
    console.error("Erro ao excluir layout de exportação:", erro);
    return res.status(500).json({
      sucesso: false,
      erro: "Não foi possível excluir o layout de exportação.",
    });
  }
});

app.get("/exportacao-universal/gerar", autenticar, async (req, res) => {
  try {
    let layout;
    let layoutId = String(req.query.layoutId || "").trim();

    if (layoutId) {
      layout = layoutsExportacao.find((l) => l.id === layoutId);
    }

    if (!layout) {
      const colunasQuery = String(req.query.colunas || "").trim();
      const formato = String(req.query.formato || "csv").trim().toLowerCase();

      const chaves = colunasQuery
        ? colunasQuery.split(",").map((c) => c.trim()).filter(Boolean)
        : COLUNAS_EXPORTACAO_DISPONIVEIS.map((c) => c.chave);

      const colunas = chaves.map((chave, indice) => {
        const encontrada = COLUNAS_EXPORTACAO_DISPONIVEIS.find(
          (c) => c.chave === chave
        );
        return {
          chave,
          rotulo: encontrada ? encontrada.rotulo : chave,
          incluida: true,
          ordem: indice,
        };
      });

      layout = normalizarLayoutExportacao({
        nome:
          String(
            req.query.nome ||
            "Exportação temporária"
          ).trim(),
      
        cliente:
          String(
            req.query.cliente || ""
          ).trim(),
      
        formato,
      
        delimitador:
          String(
            req.query.delimitador || ";"
          ),
      
        possuiCabecalho:
          req.query.possuiCabecalho !==
          "false",
      
        colunas,
      });
    } else {
      layout.ultimoUsoEm = new Date().toISOString();
      salvarLayoutsExportacao();
    }

    const dadosInventario = usarPostgres
      ? await carregarProdutosPostgres()
      : inventario;

    const itensComCalculos = dadosInventario.map((item) => ({
      ...item,
      divergencia:
        (Number(item.qtdeContada) || 0) - (Number(item.qtdeCongelada) || 0),
      valorCongelado:
        (Number(item.custoUnitario) || 0) * (Number(item.qtdeCongelada) || 0),
      valorContado:
        (Number(item.custoUnitario) || 0) * (Number(item.qtdeContada) || 0),
      valorDivergencia:
        (Number(item.custoUnitario) || 0) *
        ((Number(item.qtdeContada) || 0) - (Number(item.qtdeCongelada) || 0)),
    }));

    const { colunasAtivas, linhas } = prepararLinhasExportacao(
      itensComCalculos,
      layout.colunas
    );

    if (!colunasAtivas.length) {
      return res.status(400).json({
        sucesso: false,
        erro: "Nenhuma coluna ativa no layout.",
      });
    }

    const hoje = new Date();

const dataArquivo =
  hoje.getFullYear() +
  "-" +
  String(
    hoje.getMonth() + 1
  ).padStart(2, "0") +
  "-" +
  String(
    hoje.getDate()
  ).padStart(2, "0");

function limparParteNomeArquivo(
  valor,
  padrao
) {
  const texto = String(
    valor || padrao
  )
    .normalize("NFD")
    .replace(
      /[\u0300-\u036f]/g,
      ""
    )
    .trim()
    .replace(
      /[^a-zA-Z0-9\s_-]/g,
      ""
    )
    .replace(
      /\s+/g,
      "_"
    )
    .replace(
      /_+/g,
      "_"
    )
    .replace(
      /^_+|_+$/g,
      ""
    )
    .toUpperCase();

  return texto || padrao;
}

const nomeLayout =
  limparParteNomeArquivo(
    layout.nome,
    "LAYOUT"
  );

const nomeCliente =
  limparParteNomeArquivo(
    layout.cliente,
    "CLIENTE"
  );

const extensaoArquivo =
  layout.formato === "xlsx"
    ? "xlsx"
    : layout.formato === "txt-retorno"
      ? "txt"
      : layout.formato;

const nomeArquivo =
  `${nomeLayout}_${nomeCliente}_${dataArquivo}.${extensaoArquivo}`;

    if (layout.formato === "csv") {
      const conteudo = gerarCsvExportacao(
        linhas,
        colunasAtivas,
        layout.delimitador,
        layout.possuiCabecalho
      );
      res.setHeader("Content-Type", "text/csv; charset=utf-8");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${nomeArquivo}"`
      );
      return res.send(conteudo);
    }

    if (layout.formato === "txt") {
      const conteudo = gerarTxtExportacao(
        linhas,
        colunasAtivas,
        layout.delimitador,
        layout.possuiCabecalho
      );
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${nomeArquivo}"`
      );
      return res.send(conteudo);
    }

    if (layout.formato === "txt-retorno") {
      const t = layout.txtRetorno || {};
      const conteudo = gerarTxtRetornoExportacao(itensComCalculos, {
        prefixo: req.query.prefixo || t.prefixo || "0001010",
        sufixo: req.query.sufixo || t.sufixo || "000000",
        tamanhoEan: Number(req.query.tamanhoEan) || t.tamanhoEan || 13,
        tamanhoQtd: Number(req.query.tamanhoQtd) || t.tamanhoQtd || 8,
        tipoQuantidade: req.query.tipoQuantidade || t.tipoQuantidade || "auto",
        somenteContados: req.query.somenteContados === "1" || !!t.somenteContados,
        somenteDivergencia: req.query.somenteDivergencia === "1" || !!t.somenteDivergencia,
        somenteMaiorZero: req.query.somenteMaiorZero === "1" || !!t.somenteMaiorZero,
      });
      
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${nomeArquivo}"`
      );
      return res.send(conteudo);
    }

    if (layout.formato === "json") {
      const conteudo = gerarJsonExportacao(linhas, colunasAtivas);
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${nomeArquivo}"`
      );
      return res.send(conteudo);
    }

    if (layout.formato === "xlsx") {
      const buffer = await gerarExcelExportacao(linhas, colunasAtivas);
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${nomeArquivo}"`
      );
      return res.send(buffer);
    }

    return res.status(400).json({
      sucesso: false,
      erro: "Formato de exportação não suportado.",
    });
  } catch (erro) {
    console.error("Erro ao gerar exportação universal:", erro);
    return res.status(500).json({
      sucesso: false,
      erro: "Não foi possível gerar a exportação.",
    });
  }
});

app.post("/preview-txt", autenticar, (req, res) => {
  try {
    if (!req.files || !req.files.arquivo) {
      return res.status(400).json({ erro: "Nenhum arquivo enviado." });
    }

    const conteudo = req.files.arquivo.data.toString("utf8");
    const linhas = conteudo
      .split("\n")
      .map((l) => l.replace(/\r/g, ""))
      .filter((l) => l.trim() !== "")
      .slice(0, 15);

    const preview = linhas.map((linha, index) => ({
      linha: index + 1,
      codigoBarras: extrairCampoLinha(linha, "codigoBarras"),
      codigo: extrairCampoLinha(linha, "codigo"),
      descricao: extrairCampoLinha(linha, "descricao"),
      custoUnitario: extrairCampoLinha(linha, "custoUnitario"),
      qtdeCongelada: extrairCampoLinha(linha, "qtdeCongelada"),
      categoria: extrairCampoLinha(linha, "categoria"),
      tipo: extrairCampoLinha(linha, "tipo"),
      bruto: linha,
    }));

    res.json(preview);
  } catch (erro) {
    console.error("Erro ao gerar preview TXT:", erro);
    res.status(500).json({ erro: "Falha ao gerar preview do TXT." });
  }
});

app.post("/importar-txt", autenticar, async (req, res) => {
  try {
    if (!req.files || !req.files.arquivo) {
      return res.status(400).send("Nenhum arquivo enviado.");
    }

    garantirPastaData();

    const arquivo = req.files.arquivo;
    const caminhoTemporario = path.join(
      dataDir,
      `tmp-import-base-${Date.now()}.txt`
    );

    await arquivo.mv(caminhoTemporario);

    const linhas = await lerLinhasTxtPorStream(caminhoTemporario);

    const itensImportados = linhas.map((linha) => ({
      codigoBarras: extrairCampoLinha(linha, "codigoBarras"),
      codigo: extrairCampoLinha(linha, "codigo"),
      codigoInterno: extrairCampoLinha(linha, "codigo"),
      descricao: extrairCampoLinha(linha, "descricao"),
      custoUnitario: extrairCampoLinha(linha, "custoUnitario"),
      qtdeCongelada: extrairCampoLinha(linha, "qtdeCongelada"),
      qtdeContada: 0,
      categoria: extrairCampoLinha(linha, "categoria"),
      tipo: extrairCampoLinha(linha, "tipo"),
    }));

    const totalImportadoBruto = itensImportados.length;
    const itensUnicosImportados = removerDuplicadosPorCodigoInterno(itensImportados);
    const totalUnicosBruto = itensUnicosImportados.length;
    const duplicatasRemovidas = totalImportadoBruto - totalUnicosBruto;
    const itensZeradosIgnorados = itensUnicosImportados.filter(
      (item) => parseQuantidade(item.qtdeCongelada) <= 0
    ).length;

    auditoriaImportacao = {
      totalImportadoBruto,
      totalUnicosBruto,
      duplicatasRemovidas,
      itensZeradosIgnorados,
    };

    // reinicia o ciclo do inventário
    limparContagensPersistidas();
    historicoAlteracoes = [];
    historicoAuditoriaItens = [];
    enderecamentos = [];
    itemAuditoriaAtual = null;

    inventario = itensUnicosImportados;
    await salvarProdutosNoBanco(inventario);

modoOperacao = "com-base";
salvarModoOperacao();

tipoUltimaImportacao = "Importação base";
await salvarEnderecamentos();
broadcastInventario();
    fs.unlinkSync(caminhoTemporario);

    if (req.headers.accept?.includes("application/json")) {
      return res.json({
        sucesso: true,
        mensagem: "Arquivo base importado com sucesso.",
        total: inventario.length,
        modoOperacao
      });
    }
    
    res.redirect("/");
  } catch (erro) {
    console.error("Erro ao importar TXT:", erro);
    res.status(500).send(`Erro ao importar TXT: ${erro.message}`);
  }
});

app.get("/auditoria-importacao", autenticar, (req, res) => {
  const totalImportado =
    auditoriaImportacao.totalImportadoBruto ?? inventario.length;

  const totalUnicosPorCodigoInterno =
    auditoriaImportacao.totalUnicosBruto ?? inventario.length;

  const duplicatasRemovidas =
    auditoriaImportacao.duplicatasRemovidas ?? 0;

  const semZero = inventario.filter(
    (item) => parseQuantidade(item.qtdeCongelada) > 0
  );

  const totalComQtdeMaiorQueZero = semZero.length;

  const itensZeradosIgnorados =
    auditoriaImportacao.itensZeradosIgnorados ??
    inventario.filter((item) => parseQuantidade(item.qtdeCongelada) <= 0).length;

  const valorCongeladoTotal = semZero.reduce((acc, item) => {
    const qtde = parseQuantidade(item.qtdeCongelada);
    const custo = Number(item.custoUnitario) || 0;
    return acc + qtde * custo;
  }, 0);

  res.json({
    totalImportado,
    totalComQtdeMaiorQueZero,
    totalUnicosPorCodigoInterno,
    duplicatasRemovidas,
    itensZeradosIgnorados,
    valorCongeladoTotal: Number(valorCongeladoTotal.toFixed(2)),
    amostra: semZero.slice(0, 30).map((item) => ({
      codigoBarras: item.codigoBarras,
      codigo: item.codigo || item.codigoInterno,
      descricao: item.descricao,
      custoUnitario: Number(item.custoUnitario) || 0,
      qtdeCongelada: parseQuantidade(item.qtdeCongelada),
      valorCongelado: Number(
        (
          ((Number(item.custoUnitario) || 0) *
            parseQuantidade(item.qtdeCongelada))
        ).toFixed(2)
      ),
    })),
  });
});
app.post("/importar-saldo-atual-txt", autenticar, async (req, res) => {
  let caminhoTemporario = null;

  try {
    if (!req.files || !req.files.arquivo) {
      return res.status(400).json({
        sucesso: false,
        erro: "Nenhum arquivo enviado.",
      });
    }

    garantirPastaData();

    const arquivo = req.files.arquivo;

    caminhoTemporario = path.join(
      dataDir,
      `tmp-import-saldo-${Date.now()}.txt`
    );

    await arquivo.mv(caminhoTemporario);

    const linhas = await lerLinhasTxtPorStream(caminhoTemporario);

    const itensImportados = linhas.map((linha) => ({
      codigoBarras: extrairCampoLinha(linha, "codigoBarras"),
      codigo: extrairCampoLinha(linha, "codigo"),
      codigoInterno: extrairCampoLinha(linha, "codigo"),
      descricao: extrairCampoLinha(linha, "descricao"),
      custoUnitario: extrairCampoLinha(linha, "custoUnitario"),
      qtdeCongelada: extrairCampoLinha(linha, "qtdeCongelada"),
      categoria: extrairCampoLinha(linha, "categoria"),
      tipo: extrairCampoLinha(linha, "tipo"),
    }));

    const totalImportadoBruto = itensImportados.length;
    const itensUnicosImportados = removerDuplicadosPorCodigoInterno(itensImportados);
    const totalUnicosBruto = itensUnicosImportados.length;
    const duplicatasRemovidas = totalImportadoBruto - totalUnicosBruto;

    const itensZeradosIgnorados = itensUnicosImportados.filter(
      (item) => parseQuantidade(item.qtdeCongelada) <= 0
    ).length;

    auditoriaImportacao = {
      totalImportadoBruto,
      totalUnicosBruto,
      duplicatasRemovidas,
      itensZeradosIgnorados,
    };

    const mapaInventarioAtual = new Map(
      inventario.map((item) => [
        String(item.codigo || item.codigoInterno || "").trim(),
        item,
      ])
    );

    inventario = itensUnicosImportados.map((novoItem) => {
      const codigoBarrasNovo = String(novoItem.codigoBarras || "").trim();
    
      const totalContado = contagens
        .filter((c) => {
          return (
            c &&
            c.ativo !== false &&
            String(c.codigoBarras || "").trim() === codigoBarrasNovo
          );
        })
        .reduce((acc, c) => acc + (Number(c.quantidade) || 0), 0);
    
      return {
        ...novoItem,
        qtdeContada: totalContado,
      };
    });

    await salvarProdutosNoBanco(inventario);

    tipoUltimaImportacao = "Atualização de saldo";

    ultimaImportacao = {
      arquivo: arquivo.name,
      tipo: "Atualização de saldo",
      horario: new Date().toISOString(),
      status: "Sucesso na importação",
      observacao: `Saldo atualizado com sucesso. ${inventario.length} itens carregados. Contagens, usuários e endereços foram preservados.`,
    };

    broadcastInventario();

    if (caminhoTemporario && fs.existsSync(caminhoTemporario)) {
      fs.unlinkSync(caminhoTemporario);
    }

    return res.json({
      sucesso: true,
      mensagem: "Saldo atual importado com sucesso.",
      ultimaImportacao,
      total: inventario.length,
    });
  } catch (erro) {
    console.error("Erro ao importar saldo atual TXT:", erro);

    ultimaImportacao = {
      arquivo: req.files?.arquivo?.name || "--",
      tipo: "Atualização de saldo",
      horario: new Date().toISOString(),
      status: "Erro na importação",
      observacao: erro.message,
    };

    if (caminhoTemporario && fs.existsSync(caminhoTemporario)) {
      fs.unlinkSync(caminhoTemporario);
    }

    return res.status(500).json({
      sucesso: false,
      erro: erro.message,
      ultimaImportacao,
    });
  }
});

app.post("/importar-csv", autenticar, (req, res) => {
  if (!req.files || !req.files.arquivo) {
    return res.status(400).send("Nenhum arquivo enviado.");
  }

  const conteudo = req.files.arquivo.data.toString("utf8");
  const linhas = conteudo.split("\n").filter((l) => l.trim() !== "");

  const dados = linhas.slice(1).map((linha) => {
    const colunas = linha.split("\t");
    const [
      codigo,
      codigoBarras,
      categoria,
      descricao,
      _embalagem,
      qtdeEstoque,
      custoUnitario,
      qtdeContada,
    ] = colunas;

    return {
      codigo: codigo?.trim(),
      codigoInterno: codigo?.trim(),
      codigoBarras: codigoBarras?.trim(),
      categoria: categoria?.trim(),
      descricao: descricao?.trim(),
      custoUnitario:
        parseFloat(custoUnitario?.replace("R$", "").replace(",", ".").trim()) || 0,
      qtdeCongelada: parseFloat(qtdeEstoque?.replace(",", ".").trim()) || 0,
      qtdeContada: parseFloat(qtdeContada?.replace(",", ".").trim()) || 0,
    };
  });

  inventario = [...inventario, ...dados];
  broadcastInventario();
  res.redirect("/");
});
app.get("/filtro-app", autenticar, (req, res) => {
  try {
    if (modoOperacao === "sem-base") {
      return res.json({
        atualizadoEm: new Date().toLocaleString("pt-BR"),
        total: 0,
        totalEnderecos: 0,
        itens: [],
        enderecos: []
      });
    }

    carregarEnderecamentos();

    const itens = inventario
      .map((item) => ({
        codigoBarras: String(item.codigoBarras || "").trim(),
        codigo: String(item.codigo || item.codigoInterno || "").trim(),
        descricao: String(item.descricao || "").trim(),
      }))
      .filter((item) => item.codigoBarras || item.codigo || item.descricao)
      .sort((a, b) => {
        const descA = a.descricao.toLowerCase();
        const descB = b.descricao.toLowerCase();
        return descA.localeCompare(descB, "pt-BR");
      });

    const enderecos = [];

    enderecamentos.forEach((item) => {
      const inicio = Number(item.inicio) || 0;
      const fim = Number(item.fim) || 0;

      if (inicio > 0 && fim >= inicio) {
        for (let numero = inicio; numero <= fim; numero++) {
          enderecos.push({
            enderecoNumero: String(numero),
            id: item.id,
            nome: item.nome || "",
            tipo: item.tipo || "",
            inicio,
            fim
          });
        }
      }
    });

    return res.json({
      atualizadoEm: new Date().toLocaleString("pt-BR"),
      total: itens.length,
      totalEnderecos: enderecos.length,
      itens,
      enderecos
    });
  } catch (erro) {
    console.error("Erro ao gerar filtro do app:", erro);
    res.status(500).json({ erro: "Falha ao gerar filtro do app." });
  }
});
app.get("/transmissoes-consolidacao", autenticar, (req, res) => {
  try {
    const painel = gerarPainelTransmissoesConsolidacao();
    res.json(painel);
  } catch (erro) {
    console.error("Erro ao gerar painel de transmissões:", erro);
    res.status(500).json({ erro: "Falha ao carregar transmissões e consolidação." });
  }
});


app.post("/transmissoes-consolidacao/:id/consolidar", autenticar, async (req, res) => {
  try {
    const idParam = String(req.params.id || "");
    const match = idParam.match(/^END-(\d+)-(\d+)$/);

    if (!match) {
      return res.status(400).json({ erro: "Identificador de transmissão inválido." });
    }

    const enderecoId = Number(match[1]);
    const enderecoNumero = Number(match[2]);

    const endereco = enderecamentos.find((item) => Number(item.id) === enderecoId);

    if (!endereco) {
      return res.status(404).json({ erro: "Transmissão/endereço não encontrado." });
    }

    const transmissaoPendente = [...(endereco.transmissoes || [])]
      .reverse()
      .find((t) =>
        t.tipo === "transmissao" &&
        t.statusConsolidacao === "pendente" &&
        Number(t.enderecoNumero) === Number(enderecoNumero)
      );

    if (!transmissaoPendente) {
      return res.status(400).json({
        erro: "Nenhuma transmissão pendente encontrada para consolidar."
      });
    }

    if (!Array.isArray(endereco.consolidacoesPorNumero)) {
      endereco.consolidacoesPorNumero = [];
    }

    const jaConsolidado = endereco.consolidacoesPorNumero.some(
      (item) => Number(item.enderecoNumero) === enderecoNumero && item.consolidado
    );

    if (jaConsolidado) {
      return res.status(400).json({ erro: "Esta transmissão já foi consolidada." });
    }

    const agoraIso = new Date().toISOString();
    const usuarioConsolidacao =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

    const novosRegistros = (transmissaoPendente.itens || [])
      .map((item, index) => ({
        id: `CONT-${Date.now()}-${index}-${Math.floor(Math.random() * 1000)}`,
        usuario: transmissaoPendente.usuario || usuarioConsolidacao,
        matricula: "SEM-MATRICULA",
        codigoBarras: String(item.codigoBarras || "").trim(),
        quantidade: Number(item.quantidade) || 0,
        enderecoId: Number(endereco.id),
        enderecoNumero: Number(enderecoNumero),
        ativo: true,
        statusConsolidacao: "consolidado",
        consolidadoEm: agoraIso,
        consolidadoPor: usuarioConsolidacao,
        data: transmissaoPendente.data || agoraIso,
      }))
      .filter((item) => item.codigoBarras && item.quantidade > 0);

    contagens.push(...novosRegistros);
    salvarContagens();

    transmissaoPendente.statusConsolidacao = "consolidado";
    transmissaoPendente.consolidadoEm = agoraIso;
    transmissaoPendente.consolidadoPor = usuarioConsolidacao;

    endereco.consolidacoesPorNumero.push({
      enderecoNumero,
      consolidado: true,
      consolidadoEm: agoraIso,
      consolidadoPor: usuarioConsolidacao,
    });

    registrarEventoEndereco(Number(enderecoNumero), "finalizacao", usuarioConsolidacao);

    endereco.atualizadoEm = agoraIso;
    salvarEnderecamentos();

    recalcularInventarioComBaseNasContagens();
    await salvarProdutosNoBanco(inventario);
    broadcastInventario();

    return res.json({
      sucesso: true,
      mensagem: `Transmissão ${idParam} consolidada com sucesso.`,
      painel: gerarPainelTransmissoesConsolidacao(),
    });
  } catch (erro) {
    console.error("Erro ao consolidar transmissão:", erro);
    return res.status(500).json({ erro: "Falha ao consolidar transmissão." });
  }
});

app.post("/transmissoes-consolidacao/consolidar", autenticar, async (req, res) => {
  try {
    const painel = gerarPainelTransmissoesConsolidacao();
    const fila = Array.isArray(painel.fila) ? painel.fila : [];

    let totalConsolidadas = 0;
    let totalErros = 0;

    fila.forEach((itemFila) => {
      const idParam = String(itemFila.id || "");
      const match = idParam.match(/^END-(\d+)-(\d+)$/);

      if (!match) {
        totalErros += 1;
        return;
      }

      const enderecoId = Number(match[1]);
      const enderecoNumero = Number(match[2]);

      const endereco = enderecamentos.find((item) => Number(item.id) === enderecoId);
      if (!endereco) {
        totalErros += 1;
        return;
      }

      const transmissaoPendente = [...(endereco.transmissoes || [])]
        .reverse()
        .find((t) =>
          t.tipo === "transmissao" &&
          t.statusConsolidacao === "pendente" &&
          Number(t.enderecoNumero) === Number(enderecoNumero)
        );

      if (!transmissaoPendente) {
        totalErros += 1;
        return;
      }

      if (!Array.isArray(endereco.consolidacoesPorNumero)) {
        endereco.consolidacoesPorNumero = [];
      }

      const jaConsolidado = endereco.consolidacoesPorNumero.some(
        (c) => Number(c.enderecoNumero) === enderecoNumero && c.consolidado
      );

      if (jaConsolidado) return;

      const agoraIso = new Date().toISOString();
      const usuarioConsolidacao =
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "sistema";

        const novosRegistros = (transmissaoPendente.itens || [])
        .map((item, index) => {
          const codigoBarras = String(
            item.codigoBarras || item.eanOuCodigo || item.ean || ""
          ).trim();
      
          const codigo = String(item.codigo || "").trim();
      
          return {
            id: `CONT-${Date.now()}-${index}-${Math.floor(Math.random() * 1000)}`,
            usuario: transmissaoPendente.usuario || usuarioConsolidacao,
            matricula: "SEM-MATRICULA",
            codigoBarras,
            codigo,
            quantidade: Number(item.quantidade) || 0,
            enderecoId: Number(endereco.id),
            enderecoNumero: Number(enderecoNumero),
            ativo: true,
            statusConsolidacao: "consolidado",
            consolidadoEm: agoraIso,
            consolidadoPor: usuarioConsolidacao,
            data: transmissaoPendente.data || agoraIso,
          };
        })
        .filter((item) => (item.codigoBarras || item.codigo) && item.quantidade > 0);
      
      if (modoOperacao === "sem-base") {
        novosRegistros.forEach((registro) => {
          const chaveNova = String(registro.codigoBarras || registro.codigo || "").trim();
      
          const existente = contagemSemBase.find((item) => {
            const chaveAtual = String(item.ean || item.codigo || "").trim();
            return chaveAtual === chaveNova;
          });
      
          if (existente) {
            existente.quantidade =
              (Number(existente.quantidade) || 0) + Number(registro.quantidade || 0);
      
            if (!Array.isArray(existente.enderecos)) {
              existente.enderecos = [];
            }
      
            existente.enderecos.push({
              enderecoNumero,
            
              quantidade:
                Number(
                  registro.quantidade
                ) || 0,
            
              usuario:
                registro.usuario ||
                transmissaoPendente.usuario ||
                usuarioConsolidacao,
            
              data:
                registro.data ||
                agoraIso,
            });
      
            existente.ultimoUsuario = registro.usuario;
            existente.ultimaLeituraEm = registro.data;
          } else {
            contagemSemBase.push({
              ean:
                registro.codigoBarras,
            
              codigo:
                registro.codigo,
            
              quantidade:
                Number(
                  registro.quantidade
                ) || 0,
            
              ultimoUsuario:
                registro.usuario ||
                transmissaoPendente.usuario ||
                usuarioConsolidacao,
            
              ultimaLeituraEm:
                registro.data ||
                agoraIso,
            
              enderecos: [
                {
                  enderecoNumero,
            
                  quantidade:
                    Number(
                      registro.quantidade
                    ) || 0,
            
                  usuario:
                    registro.usuario ||
                    transmissaoPendente.usuario ||
                    usuarioConsolidacao,
            
                  data:
                    registro.data ||
                    agoraIso,
                },
              ],
            });
          }
        });
      
        finalizacoesSemBase.push({
          id: `SBF-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          enderecoNumero,
          usuario: transmissaoPendente.usuario || usuarioConsolidacao,
          data: agoraIso,
          itens: novosRegistros.map((item) => ({
            ean: item.codigoBarras,
            codigo: item.codigo,
            quantidade: item.quantidade,
          })),
          totalItensUnicos: novosRegistros.length,
          totalVolume: novosRegistros.reduce(
            (acc, item) => acc + Number(item.quantidade || 0),
            0
          ),
        });
      } else {
        contagens.push(...novosRegistros);
      }

      transmissaoPendente.statusConsolidacao = "consolidado";
      transmissaoPendente.consolidadoEm = agoraIso;
      transmissaoPendente.consolidadoPor = usuarioConsolidacao;

      endereco.consolidacoesPorNumero.push({
        enderecoNumero,
        consolidado: true,
        consolidadoEm: agoraIso,
        consolidadoPor: usuarioConsolidacao,
      });

      registrarEventoEndereco(Number(enderecoNumero), "finalizacao", usuarioConsolidacao);

      endereco.atualizadoEm = agoraIso;
      totalConsolidadas += 1;
    });

    salvarEnderecamentos();

    if (modoOperacao === "sem-base") {
      await salvarContagemSemBase();
      await salvarFinalizacoesSemBase();
    } else {
      salvarContagens();
      recalcularInventarioComBaseNasContagens();
      await salvarProdutosNoBanco(inventario);
    }
    
    broadcastInventario();
    return res.json({
      sucesso: true,
      mensagem: `${totalConsolidadas} transmissão(ões) consolidada(s). ${totalErros} erro(s).`,
      painel: gerarPainelTransmissoesConsolidacao(),
    });
  } catch (erro) {
    console.error("Erro ao consolidar transmissões em fila:", erro);
    return res.status(500).json({
      erro: "Falha ao consolidar transmissões em fila.",
    });
  }
});
app.post("/enderecamentos/:id/resolver-duplicidade", autenticar, (req, res) => {
  try {
    const id = Number(req.params.id);
    const endereco = enderecamentos.find((item) => Number(item.id) === id);

    if (!endereco) {
      return res.status(404).json({ erro: "Endereço não encontrado." });
    }

    endereco.duplicidadeResolvida = true;
    endereco.duplicidadeResolvidaEm = new Date().toISOString();
    endereco.duplicidadeResolvidaPor =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

    if (Number(endereco.contagensRecebidas || 0) > 0) {
      if (endereco.finalizadoViaColetor) {
        endereco.status = "em-fila";
      } else {
        endereco.status = "em-contagem";
      }
    }

    endereco.atualizadoEm = new Date().toISOString();

    salvarEnderecamentos();

    return res.json({
      sucesso: true,
      mensagem: "Duplicidade resolvida com sucesso."
    });
  } catch (erro) {
    console.error("Erro ao resolver duplicidade:", erro);
    return res.status(500).json({ erro: "Falha ao resolver duplicidade." });
  }
});
app.post("/atualizar-campo", autenticar, (req, res) => {
  const { codigoBarras, campo, valor } = req.body;
  const item = inventario.find((i) => i.codigoBarras === codigoBarras);

  if (!item) {
    return res.status(404).json({ erro: "Produto não encontrado" });
  }

  const valorAntigo = item[campo];

  item[campo] = ["qtdeContada", "qtdeCongelada", "custoUnitario"].includes(campo)
    ? parseFloat(String(valor).replace(",", ".")) || 0
    : valor;

  registrarAlteracao(
    req.session.usuario,
    codigoBarras,
    campo,
    valorAntigo,
    item[campo]
  );

  if (campo === "qtdeContada") {
    const antigo = Number(valorAntigo) || 0;
    const novo = Number(item[campo]) || 0;
    const delta = novo - antigo;

    if (delta > 0) {
      registrarContagem(req.session.usuario, codigoBarras, delta);
    }
  }

  broadcastInventario();
  res.json({ sucesso: true });
});

app.get("/historico-dados", autenticar, (req, res) =>
  res.json(historicoAlteracoes)
);

app.get("/historico", autenticar, (req, res) =>
res.sendFile(caminhoPublico("historico.html"))
);

app.get("/exportar-txt-alternativo", autenticar, (req, res) => {
  let conteudo = "";

  const cliente = req.query.cliente || "cliente";
  const loja = req.query.loja || "loja";
  const data = req.query.data || new Date().toISOString().slice(0, 10);

  const prefixo = req.query.prefixo || "0001010";
  const sufixo = req.query.sufixo || "000000";
  const tamanhoEan = Number(req.query.tamanhoEan) || 13;
  const tamanhoQtd = Number(req.query.tamanhoQtd) || 8;
  const tipoQuantidade = req.query.tipoQuantidade || "auto";

  const somenteContados = req.query.somenteContados === "1";
  const somenteDivergencia = req.query.somenteDivergencia === "1";
  const somenteMaiorZero = req.query.somenteMaiorZero === "1";

  let lista = [];

  if (modoOperacao === "sem-base") {
    lista = (Array.isArray(contagemSemBase) ? contagemSemBase : []).map((item) => ({
      codigoBarras: item.ean || "",
      codigo: item.codigo || "",
      categoria: "",
      descricao: "",
      qtdeCongelada: 0,
      qtdeContada: Number(item.quantidade || 0),
    }));

    if (somenteContados) {
      lista = lista.filter((item) => (Number(item.qtdeContada) || 0) > 0);
    }

    if (somenteDivergencia) {
      lista = lista.filter((item) => (Number(item.qtdeContada) || 0) !== 0);
    }

    lista.forEach((item) => {
      const codigoBase = String(item.codigoBarras || item.codigo || "")
        .replace(/\D/g, "");

      const codigoEan = codigoBase
        .padStart(tamanhoEan, "0")
        .slice(-tamanhoEan);

      const qtde = parseFloat(item.qtdeContada) || 0;

      let quantidadeFormatada = "";

      if (tipoQuantidade === "inteiro") {
        quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
      } else if (tipoQuantidade === "milhar") {
        quantidadeFormatada = Math.round(qtde * 1000).toString().padStart(tamanhoQtd, "0");
      } else {
        quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
      }

      conteudo += `${prefixo}${codigoEan}${quantidadeFormatada}${sufixo}\n`;
    });
  } else {
    recalcularInventarioComBaseNasContagens();
  
    lista = filtrarInventario({
      categoria: req.query.categoria || "",
      ordem: req.query.ordem || "",
      busca: req.query.busca || "",
    });
    if (somenteContados) {
      lista = lista.filter((item) => (Number(item.qtdeContada) || 0) > 0);
    }

    if (somenteDivergencia) {
      lista = lista.filter((item) => {
        const qtdeContada = Number(item.qtdeContada) || 0;
        const qtdeCongelada = Number(item.qtdeCongelada) || 0;
        return qtdeContada - qtdeCongelada !== 0;
      });
    }

    if (somenteMaiorZero) {
      lista = lista.filter((item) => (Number(item.qtdeCongelada) || 0) > 0);
    }

    lista.forEach((item) => {
      const codigoEan = String(item.codigoBarras || "")
        .padStart(tamanhoEan, "0")
        .slice(-tamanhoEan);

      const qtde = parseFloat(item.qtdeContada) || 0;
      const textoCategoria = `${item.categoria || ""} ${item.descricao || ""}`.toLowerCase();

      let quantidadeFormatada = "";

      if (tipoQuantidade === "inteiro") {
        quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
      } else if (tipoQuantidade === "milhar") {
        quantidadeFormatada = Math.round(qtde * 1000).toString().padStart(tamanhoQtd, "0");
      } else {
        if (textoCategoria.includes("kg")) {
          quantidadeFormatada =
            qtde % 1 === 0
              ? qtde.toString().padStart(tamanhoQtd, "0")
              : Math.round(qtde * 1000).toString().padStart(tamanhoQtd, "0");
        } else {
          quantidadeFormatada = Math.round(qtde).toString().padStart(tamanhoQtd, "0");
        }
      }

      conteudo += `${prefixo}${codigoEan}${quantidadeFormatada}${sufixo}\n`;
    });
  }

  const nomeArquivo = `${cliente}_${loja}_${data}_retorno.txt`
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .replace(/[^\w.-]+/g, "_")
  .replace(/_+/g, "_");

  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${nomeArquivo}"`
  );
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  return res.send(conteudo);
  });
  
  app.get("/exportar-csv", autenticar, (req, res) => {
  const cabecalho = [
    "Código de Barras",
    "Código Interno",
    "Descrição",
    "Custo Unitário",
    "Categoria",
    "Qtd. Congelada",
    "Qtd. Contada",
    "Divergência",
    "Valor Congelado",
    "Valor Contado",
    "Valor Divergente",
  ];

  const linhas = inventario.map((item) => {
    const qtdeContada = Number(item.qtdeContada) || 0;
    const qtdeCongelada = Number(item.qtdeCongelada) || 0;
    const custo = Number(item.custoUnitario) || 0;
    const divergencia = qtdeContada - qtdeCongelada;

    return [
      item.codigoBarras,
      item.codigo || item.codigoInterno || "",
      item.descricao,
      custo.toFixed(2),
      item.categoria,
      qtdeCongelada,
      qtdeContada,
      divergencia,
      (custo * qtdeCongelada).toFixed(2),
      (custo * qtdeContada).toFixed(2),
      (custo * divergencia).toFixed(2),
    ].join(";");
  });

  const conteudo = [cabecalho.join(";"), ...linhas].join("\n");
  const cliente = req.query.cliente || "cliente";
  const loja = req.query.loja || "loja";
  const data = req.query.data || new Date().toISOString().slice(0, 10);

  const nomeArquivo = `${cliente}_${loja}_${data}_inventario.csv`.replace(/\s+/g, "_");

  res.setHeader("Content-disposition", `attachment; filename=${nomeArquivo}`);
  res.setHeader("Content-Type", "text/csv");
  res.send(conteudo);
});

app.get("/exportar-excel", autenticar, async (req, res) => {
  try {
    const resultado = filtrarInventario({
      categoria: req.query.categoria || "",
      ordem: req.query.ordem || "",
      busca: req.query.busca || "",
    });

    const cliente = req.query.cliente || "Cliente não informado";
    const loja = req.query.loja || "Loja não informada";
    const dataInventario =
      req.query.data || new Date().toLocaleDateString("pt-BR");

    const geradoEm = new Date().toLocaleString("pt-BR");

    const totalProdutos = resultado.length;
    const totalItensComDivergencia = resultado.filter(
      (item) => Number(item.divergencia) !== 0
    ).length;
    const totalValorCongelado =
  modoOperacao === "sem-base"
    ? 0
    : resultado.reduce((sum, item) => sum + (Number(item.valorCongelado) || 0), 0);

const totalValorContado =
  modoOperacao === "sem-base"
    ? 0
    : resultado.reduce((sum, item) => sum + (Number(item.valorContado) || 0), 0);

const totalValorDivergencia =
  modoOperacao === "sem-base"
    ? 0
    : resultado.reduce((sum, item) => sum + (Number(item.valorDivergencia) || 0), 0);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "RealStock";
    workbook.lastModifiedBy = "RealStock";
    workbook.created = new Date();
    workbook.modified = new Date();

    const worksheet = workbook.addWorksheet("Inventário", {
      views: [{ state: "frozen", ySplit: 11 }],
      properties: { defaultRowHeight: 20 },
    });

    const caminhoLogoJpg = path.join(__dirname, "public", "logo-realstock.png");
    const caminhoLogoPng = path.join(__dirname, "public", "logo-realstock.png");

    if (fs.existsSync(caminhoLogoJpg)) {
      const logoId = workbook.addImage({
        filename: caminhoLogoJpg,
        extension: "jpeg",
      });

      worksheet.addImage(logoId, {
        tl: { col: 0, row: 0 },
        ext: { width: 210, height: 75 },
      });
    } else if (fs.existsSync(caminhoLogoPng)) {
      const logoId = workbook.addImage({
        filename: caminhoLogoPng,
        extension: "png",
      });

      worksheet.addImage(logoId, {
        tl: { col: 0, row: 0 },
        ext: { width: 210, height: 75 },
      });
    }

    worksheet.columns = [
      { key: "codigoBarras", width: 18 },
      { key: "codigo", width: 14 },
      { key: "descricao", width: 40 },
      { key: "categoria", width: 22 },
      { key: "custoUnitario", width: 14 },
      { key: "qtdeCongelada", width: 14 },
      { key: "qtdeContada", width: 14 },
      { key: "divergencia", width: 14 },
      { key: "ajuste", width: 12 },
      { key: "situacao", width: 18 },
      { key: "valorCongelado", width: 16 },
      { key: "valorContado", width: 16 },
      { key: "valorDivergencia", width: 17 },
    ];

    worksheet.mergeCells("A1:M1");
    worksheet.mergeCells("A2:M2");
    worksheet.mergeCells("A3:M3");
    worksheet.mergeCells("A4:M4");

    worksheet.getCell("A1").value = "RELATÓRIO DE INVENTÁRIO";
    worksheet.getCell("A2").value = `Cliente: ${cliente}`;
    worksheet.getCell("A3").value = `Loja: ${loja}`;
    worksheet.getCell("A4").value = `Data do Inventário: ${dataInventario}   |   Gerado em: ${geradoEm}`;

    worksheet.getCell("A1").font = {
      bold: true,
      size: 18,
      color: { argb: "FF0F172A" },
    };
    worksheet.getCell("A1").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    ["A2", "A3", "A4"].forEach((cellRef) => {
      worksheet.getCell(cellRef).font = {
        bold: true,
        size: 11,
        color: { argb: "FF334155" },
      };
      worksheet.getCell(cellRef).alignment = {
        horizontal: "left",
        vertical: "middle",
      };
    });

    worksheet.getRow(1).height = 28;
    worksheet.getRow(2).height = 22;
    worksheet.getRow(3).height = 22;
    worksheet.getRow(4).height = 22;

    worksheet.mergeCells("A6:C6");
    worksheet.mergeCells("D6:F6");
    worksheet.mergeCells("G6:J6");
    worksheet.mergeCells("K6:M6");

    worksheet.getCell("A6").value = `Produtos\n${totalProdutos.toLocaleString("pt-BR")}`;
    worksheet.getCell("D6").value = `Itens com Divergência\n${totalItensComDivergencia.toLocaleString("pt-BR")}`;
    worksheet.getCell("G6").value = `Valor Congelado\nR$ ${formatarValorPdf(totalValorCongelado)}`;
    worksheet.getCell("K6").value = `Valor Divergente\nR$ ${formatarValorPdf(totalValorDivergencia)}`;

    ["A6", "D6", "G6", "K6"].forEach((cellRef) => {
      const cell = worksheet.getCell(cellRef);
      cell.font = {
        bold: true,
        size: 11,
        color: { argb: "FF0F172A" },
      };
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEFF6FF" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } },
      };
    });

    worksheet.getRow(6).height = 36;
    worksheet.getRow(7).height = 8;

    const headerRowNumber = 8;
    const headerRow = worksheet.getRow(headerRowNumber);

    headerRow.values = [
      "Código de Barras",
      "Código Interno",
      "Descrição",
      "Categoria",
      "Custo Unitário",
      "Qtd. Congelada",
      "Qtd. Contada",
      "Divergência",
      "Ajuste",
      "Situação",
      "Valor Congelado",
      "Valor Contado",
      "Valor Divergente",
    ];

    headerRow.height = 24;

    headerRow.eachCell((cell) => {
      cell.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
        size: 10,
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1E293B" },
      };
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } },
      };
    });

    resultado.forEach((item) => {
      const qtdeContada = Number(item.qtdeContada) || 0;
      const qtdeCongelada = Number(item.qtdeCongelada) || 0;

      let situacao = "SEM DIVERGÊNCIA";
      if (qtdeContada === 0) situacao = "ZERADOS";
      else if (qtdeContada > qtdeCongelada) situacao = "SOBRA";
      else if (qtdeContada < qtdeCongelada) situacao = "FALTA";

      worksheet.addRow({
        codigoBarras: item.codigoBarras || "",
        codigo: item.codigo || item.codigoInterno || "",
        descricao: item.descricao || "",
        categoria: item.categoria || "",
        custoUnitario: Number(item.custoUnitario) || 0,
        qtdeCongelada: Number(item.qtdeCongelada) || 0,
        qtdeContada: Number(item.qtdeContada) || 0,
        divergencia: Number(item.divergencia) || 0,
        ajuste: "",
        situacao,
        valorCongelado: Number(item.valorCongelado) || 0,
        valorContado: Number(item.valorContado) || 0,
        valorDivergencia: Number(item.valorDivergencia) || 0,
      });
    });

    const firstDataRow = 9;
    const lastDataRow = worksheet.rowCount;

    for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      row.eachCell((cell) => {
        cell.font = {
          size: 10,
          color: { argb: "FF0F172A" },
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
      });

      row.getCell(3).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      row.getCell(4).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      row.getCell(5).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(6).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(7).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(8).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(9).alignment = { vertical: "middle", horizontal: "center" };
      row.getCell(10).alignment = { vertical: "middle", horizontal: "center" };
      row.getCell(11).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(12).alignment = { vertical: "middle", horizontal: "right" };
      row.getCell(13).alignment = { vertical: "middle", horizontal: "right" };

      row.getCell(5).numFmt = "R$ #,##0.00";
      row.getCell(6).numFmt = "#,##0.000";
      row.getCell(7).numFmt = "#,##0.000";
      row.getCell(8).numFmt = "#,##0.000";
      row.getCell(11).numFmt = "R$ #,##0.00";
      row.getCell(12).numFmt = "R$ #,##0.00";
      row.getCell(13).numFmt = "R$ #,##0.00";
    }

    const totalRow = worksheet.addRow({
      descricao: "TOTAIS",
      valorCongelado: totalValorCongelado,
      valorContado: totalValorContado,
      valorDivergencia: totalValorDivergencia,
    });

    totalRow.font = {
      bold: true,
      size: 10,
      color: { argb: "FF0F172A" },
    };

    totalRow.getCell(11).numFmt = "R$ #,##0.00";
    totalRow.getCell(12).numFmt = "R$ #,##0.00";
    totalRow.getCell(13).numFmt = "R$ #,##0.00";

    totalRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } },
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE2E8F0" },
      };
      cell.alignment = {
        vertical: "middle",
        horizontal: "right",
      };
    });

    totalRow.getCell(3).alignment = {
      vertical: "middle",
      horizontal: "left",
    };

    worksheet.autoFilter = {
      from: "A8",
      to: "M8",
    };

    const nomeArquivo = `${cliente}_${loja}_${dataInventario}_inventario.xlsx`
      .replace(/\s+/g, "_")
      .replace(/[\/\\:]/g, "-");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename=${nomeArquivo}`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (erro) {
    console.error("Erro ao exportar Excel:", erro);
    res.status(500).send("Erro ao exportar Excel.");
  }
});

function formatarValorPdf(valor) {
  return (Number(valor) || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatarNumeroPdf(valor) {
  const n = Number(valor) || 0;

  if (Number.isInteger(n)) {
    return n.toLocaleString("pt-BR");
  }

  return n.toLocaleString("pt-BR", {
    minimumFractionDigits: 3,
    maximumFractionDigits: 3,
  });
}

function obterSituacaoItem(item) {
  const qtdeContada = Number(item.qtdeContada) || 0;
  const qtdeCongelada = Number(item.qtdeCongelada) || 0;

  if (qtdeContada === 0) return "SEM CONTAGEM";
  if (qtdeContada > qtdeCongelada) return "SOBRA";
  if (qtdeContada < qtdeCongelada) return "FALTA";
  return "SEM DIVERGÊNCIA";
}

function calcularCurvaABC(itens) {
  const base = [...itens]
    .map((item) => {
      const valorBase = Math.abs(Number(item.valorCongelado) || 0);

      return {
        ...item,
        valorBase,
      };
    })
    .sort((a, b) => b.valorBase - a.valorBase);

  const totalValorBase = base.reduce((acc, item) => acc + item.valorBase, 0);

  let acumulado = 0;

  const classificados = base.map((item) => {
    const participacao = totalValorBase > 0 ? (item.valorBase / totalValorBase) * 100 : 0;
    acumulado += participacao;

    let classe = "C";
    if (acumulado <= 80) classe = "A";
    else if (acumulado <= 95) classe = "B";

    return {
      ...item,
      participacao,
      acumulado,
      classe,
    };
  });

  return {
    totalValorBase,
    itens: classificados,
  };
}

function resumirCurvaPorClasse(itensABC) {
  const resumo = {
    A: { classe: "A", itens: 0, valor: 0, percentual: 0 },
    B: { classe: "B", itens: 0, valor: 0, percentual: 0 },
    C: { classe: "C", itens: 0, valor: 0, percentual: 0 },
  };

  const total = itensABC.reduce((acc, item) => acc + (Number(item.valorBase) || 0), 0);

  itensABC.forEach((item) => {
    const classe = item.classe || "C";
    if (!resumo[classe]) return;

    resumo[classe].itens += 1;
    resumo[classe].valor += Number(item.valorBase) || 0;
  });

  Object.values(resumo).forEach((grupo) => {
    grupo.percentual = total > 0 ? (grupo.valor / total) * 100 : 0;
  });

  return Object.values(resumo);
}

function resumirDivergencias(itens) {
  let sobraQtd = 0;
  let faltaQtd = 0;
  let sobraValor = 0;
  let faltaValor = 0;
  let semDivergencia = 0;
  let semContagem = 0;

  itens.forEach((item) => {
    const divergencia = Number(item.divergencia) || 0;
    const valorDivergencia = Number(item.valorDivergencia) || 0;
    const qtdeContada = Number(item.qtdeContada) || 0;

    if (qtdeContada === 0) {
      semContagem += 1;
    }

    if (divergencia > 0) {
      sobraQtd += divergencia;
      sobraValor += Math.abs(valorDivergencia);
    } else if (divergencia < 0) {
      faltaQtd += Math.abs(divergencia);
      faltaValor += Math.abs(valorDivergencia);
    } else {
      semDivergencia += 1;
    }
  });

  return {
    sobraQtd,
    faltaQtd,
    sobraValor,
    faltaValor,
    semDivergencia,
    semContagem,
  };
}

function resumirCategoriasDivergencia(itens) {
  const mapa = {};

  itens.forEach((item) => {
    const categoria = String(item.categoria || "SEM CATEGORIA").trim() || "SEM CATEGORIA";
    const divergencia = Number(item.divergencia) || 0;
    const valorDivergencia = Number(item.valorDivergencia) || 0;

    if (!mapa[categoria]) {
      mapa[categoria] = {
        categoria,
        sobraQtd: 0,
        faltaQtd: 0,
        sobraValor: 0,
        faltaValor: 0,
        resultadoValor: 0,
      };
    }

    if (divergencia > 0) {
      mapa[categoria].sobraQtd += divergencia;
      mapa[categoria].sobraValor += Math.abs(valorDivergencia);
    } else if (divergencia < 0) {
      mapa[categoria].faltaQtd += Math.abs(divergencia);
      mapa[categoria].faltaValor += Math.abs(valorDivergencia);
    }

    mapa[categoria].resultadoValor =
      (mapa[categoria].sobraValor || 0) - (mapa[categoria].faltaValor || 0);
  });

  return Object.values(mapa).sort((a, b) => {
    const totalA = Math.abs(a.resultadoValor || 0);
    const totalB = Math.abs(b.resultadoValor || 0);
    return totalB - totalA;
  });
}

app.get(
  "/exportar-pdf",
  autenticar,
  async (req, res) => {
    let resultado = [];

    if (modoOperacao === "sem-base") {
      resultado =
        montarLinhasSemBaseParaTabela({
          busca:
            req.query.busca || "",
        });
    } else if (
      modoOperacao === "wms"
    ) {
      if (!usarPostgres) {
        return res.status(503).json({
          sucesso: false,
          erro:
            "O PostgreSQL precisa estar ativo para gerar o relatório WMS.",
        });
      }
    
      const baseEsperadaWms =
        await carregarBaseEsperadaWmsPostgres(
          {}
        );
    
      const contagensWms =
        await carregarContagensWmsPostgres();
    
      const normalizarCodigoWms = (
        valor
      ) =>
        String(valor || "")
          .trim()
          .replace(/^0+/, "");
    
      resultado = (
        Array.isArray(baseEsperadaWms)
          ? baseEsperadaWms
          : []
      ).map((itemBase) => {
        const enderecoWms =
          String(
            itemBase.enderecoWms || ""
          ).trim();
    
        const codigoBarrasBase =
          normalizarCodigoWms(
            itemBase.codigoBarras
          );
    
        const codigoBase =
          normalizarCodigoWms(
            itemBase.codigo
          );
    
        const produtoCadastro =
          inventario.find((produto) => {
            const eanProduto =
              normalizarCodigoWms(
                produto.codigoBarras
              );
    
            const codigoProduto =
              normalizarCodigoWms(
                produto.codigo ||
                produto.codigoInterno
              );
    
            return (
              (
                codigoBarrasBase &&
                eanProduto ===
                  codigoBarrasBase
              ) ||
              (
                codigoBase &&
                codigoProduto ===
                  codigoBase
              )
            );
          });
    
        const quantidadeEsperada =
          Number(
            itemBase.quantidadeEsperada
          ) || 0;
    
        const quantidadeContada =
          (
            Array.isArray(contagensWms)
              ? contagensWms
              : []
          )
            .filter((contagem) => {
              const mesmoEndereco =
                String(
                  contagem.enderecoWms ||
                  ""
                ).trim() === enderecoWms;
    
              if (!mesmoEndereco) {
                return false;
              }
    
              const eanContagem =
                normalizarCodigoWms(
                  contagem.codigoBarras
                );
    
              const codigoContagem =
                normalizarCodigoWms(
                  contagem.codigo
                );
    
              return (
                (
                  codigoBarrasBase &&
                  eanContagem ===
                    codigoBarrasBase
                ) ||
                (
                  codigoBase &&
                  codigoContagem ===
                    codigoBase
                )
              );
            })
            .reduce(
              (
                total,
                contagem
              ) =>
                total +
                (
                  Number(
                    contagem.quantidadeContada
                  ) || 0
                ),
              0
            );
    
        const custoUnitario =
          Number(
            produtoCadastro
              ?.custoUnitario
          ) || 0;
    
        const divergencia =
          quantidadeContada -
          quantidadeEsperada;
    
        return {
          enderecoWms,
    
          codigoBarras:
            itemBase.codigoBarras ||
            produtoCadastro
              ?.codigoBarras ||
            "",
    
          codigo:
            itemBase.codigo ||
            produtoCadastro?.codigo ||
            produtoCadastro
              ?.codigoInterno ||
            "",
    
          descricao:
            itemBase.descricao ||
            produtoCadastro
              ?.descricao ||
            "Sem descrição",
    
          categoria:
            produtoCadastro
              ?.categoria || "",
    
          custoUnitario,
    
          qtdeCongelada:
            quantidadeEsperada,
    
          qtdeContada:
            quantidadeContada,
    
          divergencia,
    
          valorCongelado:
            quantidadeEsperada *
            custoUnitario,
    
          valorContado:
            quantidadeContada *
            custoUnitario,
    
          valorDivergencia:
            divergencia *
            custoUnitario,
    
          enderecosContagem: [],
        };
      });
    
      const busca =
        String(
          req.query.busca || ""
        )
          .trim()
          .toLowerCase();
    
      if (busca) {
        resultado =
          resultado.filter(
            (item) =>
              [
                item.enderecoWms,
                item.codigoBarras,
                item.codigo,
                item.descricao,
                item.categoria,
              ].some((valor) =>
                String(valor || "")
                  .toLowerCase()
                  .includes(busca)
              )
          );
      }
    
      const categoria =
        String(
          req.query.categoria || ""
        )
          .trim()
          .toLowerCase();
    
      if (categoria) {
        resultado =
          resultado.filter(
            (item) =>
              String(
                item.categoria || ""
              )
                .toLowerCase()
                .includes(categoria)
          );
      }
    } else {
      resultado =
        filtrarInventario({
          categoria:
            req.query.categoria || "",
    
          ordem:
            req.query.ordem || "",
    
          busca:
            req.query.busca || "",
        });
    
      resultado =
        resultado.filter(
          (item) =>
            parseQuantidade(
              item.qtdeCongelada
            ) > 0
        );
    }

  const colunasSelecionadas = (req.query.colunas || "")
    .split(",")
    .map((c) => c.trim())
    .filter(Boolean);
    let cabecalhosPersonalizados = {};

try {
  cabecalhosPersonalizados =
    JSON.parse(
      String(
        req.query.cabecalhos || "{}"
      )
    );

  if (
    !cabecalhosPersonalizados ||
    typeof cabecalhosPersonalizados !==
      "object" ||
    Array.isArray(
      cabecalhosPersonalizados
    )
  ) {
    cabecalhosPersonalizados = {};
  }
} catch (erro) {
  cabecalhosPersonalizados = {};
}
    function obterEnderecosPdf(item) {
      const lista = Array.isArray(item.enderecosContagem)
        ? item.enderecosContagem
        : [];
    
      return lista.map((e) => {
        const numero = e.enderecoNumero || "";
        const enderecoObj = buscarEnderecoPorNumero(Number(numero));
        const nome = e.nome || enderecoObj?.nome || "ENDEREÇO";
    
        return {
          nome,
          numero,
          quantidade: Number(e.quantidade || 0),
        };
      });
    }
    
    function formatarEnderecosPdf(item) {
      return obterEnderecosPdf(item)
        .map((e) => `${e.nome} • ${e.numero}`)
        .join("\n\n");
    }
    
    function formatarColetaEnderecoPdf(item) {
      return obterEnderecosPdf(item)
        .map((e) => formatarNumeroPdf(e.quantidade))
        .join("\n\n");
    }
  const mapaColunas = {
    codigoBarras: {
      titulo: "Código EAN",
      valor: (item) => item.codigoBarras || "",
      width: 95,
      align: "left",
    },
    codigo: {
      titulo: "Cód. Interno",
      valor: (item) => item.codigo || item.codigoInterno || "",
      width: 75,
      align: "left",
    },
    descricao: {
      titulo: "Descrição",
      valor: (item) => item.descricao || "",
      width: "*",
      align: "left",
    },
    categoria: {
      titulo: "Categoria",
      valor: (item) => item.categoria || "",
      width: 105,
      align: "left",
    },
    custoUnitario: {
      titulo: "Custo Unit.",
      valor: (item) => `R$ ${formatarValorPdf(item.custoUnitario)}`,
      width: 78,
      align: "right",
    },
    qtdeCongelada: {
      titulo: "Qtd. Cong.",
      valor: (item) => formatarNumeroPdf(item.qtdeCongelada),
      width: 68,
      align: "right",
    },
    qtdeContada: {
      titulo: "Qtd. Cont.",
      valor: (item) => formatarNumeroPdf(item.qtdeContada),
      width: 68,
      align: "right",
    },
    divergencia: {
      titulo: "Diverg.",
      valor: (item) => formatarNumeroPdf(item.divergencia),
      width: 70,
      align: "right",
    },
    situacao: {
      titulo: "Situação",
      valor: (item) => {
        if (modoOperacao === "sem-base") {
          return Number(item.qtdeContada || 0) > 0 ? "CONTADO" : "ZERADOS";
        }
    
        const qtdeContada = Number(item.qtdeContada) || 0;
        const qtdeCongelada = Number(item.qtdeCongelada) || 0;
    
        if (qtdeContada === 0) return "ZERADOS";
        if (qtdeContada > qtdeCongelada) return "SOBRA";
        if (qtdeContada < qtdeCongelada) return "FALTA";
        return "SEM DIVERGÊNCIA";
      },
      width: 90,
      align: "center",
    },
    endereco: {
      titulo: "Endereço",
      valor: (item) => formatarEnderecosPdf(item),
      width: 95,
      align: "left",
    },
    enderecoWms: {
      titulo: "Endereço WMS",
    
      valor: (item) =>
        String(
          item.enderecoWms || ""
        ),
    
      width: 82,
      align: "left",
    },
    coletaEndereco: {
      titulo: "Coleta End",
      valor: (item) => formatarColetaEnderecoPdf(item),
      width: 60,
      align: "right",
    },
    ajuste: {
      titulo: "Ajuste",
      valor: () => "",
      width: 65,
      align: "center",
    },
    valorCongelado: {
      titulo: "Valor Cong.",
      valor: (item) => `R$ ${formatarValorPdf(item.valorCongelado)}`,
      width: 88,
      align: "right",
    },
    valorContado: {
      titulo: "Valor Cont.",
      valor: (item) => `R$ ${formatarValorPdf(item.valorContado)}`,
      width: 88,
      align: "right",
    },
    valorDivergencia: {
      titulo: "Valor Diverg.",
      valor: (item) => `R$ ${formatarValorPdf(item.valorDivergencia)}`,
      width: 92,
      align: "right",
    },
  };

  const colunasPadrao = [
    "codigoBarras",
    "codigo",
    "descricao",
    "categoria",
    "endereco",
    "enderecoWms",
"coletaEndereco",
"custoUnitario",
    "qtdeCongelada",
    "qtdeContada",
    "divergencia",
    "situacao",
    "ajuste",
    "valorCongelado",
    "valorContado",
    "valorDivergencia",
  ];
  const colunasAtivas = (colunasSelecionadas.length ? colunasSelecionadas : colunasPadrao)
    .filter((coluna) => mapaColunas[coluna]);

    const margemEsquerdaPdf = 12;
    const margemDireitaPdf = 22;

const larguraPaginaA4Paisagem = 841.89;

// Área útil real da A4 paisagem.
// Desconto extra de segurança para bordas e paddings internos do pdfmake.
const larguraDisponivelPdf =
  larguraPaginaA4Paisagem - margemEsquerdaPdf - margemDireitaPdf - 28;

const largurasBasePdf = {
  codigoBarras: 88,
  codigo: 55,
  descricao: 240,
  categoria: 70,
  custoUnitario: 62,
  qtdeCongelada: 68,
  qtdeContada: 68,
  divergencia: 72,
  situacao: 70,
  ajuste: 70,
  endereco: 68,
  enderecoWms: 82,
  coletaEndereco: 68,
  valorCongelado: 82,
  valorContado: 82,
  valorDivergencia: 92,
};

const larguraTotalSelecionadaPdf = colunasAtivas.reduce(
  (total, coluna) => total + (largurasBasePdf[coluna] || 65),
  0
);

const fatorAjustePdf = larguraDisponivelPdf / larguraTotalSelecionadaPdf;

let largurasTabelaPdf = colunasAtivas.map((coluna) => {
  const larguraBase = largurasBasePdf[coluna] || 65;
  return Math.floor(larguraBase * fatorAjustePdf);
});
const somaLargurasPdf = largurasTabelaPdf.reduce((acc, largura) => acc + largura, 0);
const diferencaPdf = Math.floor(larguraDisponivelPdf - somaLargurasPdf);

const indiceDescricaoPdf = colunasAtivas.indexOf("descricao");
if (indiceDescricaoPdf >= 0 && diferencaPdf > 0) {
  largurasTabelaPdf[indiceDescricaoPdf] += diferencaPdf;
}

    const fonteTabelaPdf = 9;
  const printer = new PdfPrinter(fonts);

  const tableBody = [
    colunasAtivas.map((coluna) => ({
      text:
  cabecalhosPersonalizados[coluna] ||
  mapaColunas[coluna].titulo,
      style: "tableHeader",
      alignment: mapaColunas[coluna].align || "left",
    })),
  ];

  if (resultado.length === 0) {
    const linhaVazia = new Array(colunasAtivas.length).fill("");
    linhaVazia[0] = {
      text: "Nenhum produto encontrado",
      colSpan: colunasAtivas.length,
      alignment: "center",
      margin: [0, 8, 0, 8],
    };
    tableBody.push(linhaVazia);
  } else {
    resultado.forEach((item) => {
      tableBody.push(
        colunasAtivas.map((coluna) => ({
          text: String(mapaColunas[coluna].valor(item)),
style: "tableCell",
alignment: mapaColunas[coluna].align || "left",
fontSize: fonteTabelaPdf,
noWrap: false,
        }))
      );
    });
  }

  const totalProdutos = resultado.length;
  const totalValorCongelado = resultado.reduce(
    (acc, item) => acc + (Number(item.valorCongelado) || 0),
    0
  );
  const totalValorContado = resultado.reduce(
    (acc, item) => acc + (Number(item.valorContado) || 0),
    0
  );
  const totalValorDivergencia = resultado.reduce(
    (acc, item) => acc + (Number(item.valorDivergencia) || 0),
    0
  );
  const totalItensComDivergencia = resultado.filter(
    (item) => Number(item.divergencia) !== 0
  ).length;

  const cliente = req.query.cliente || "Cliente não informado";
  const loja = req.query.loja || "Loja não informada";
  const agoraPdf = new Date();

const dataInventario =
  req.query.data ||
  agoraPdf.toLocaleDateString("pt-BR", {
    timeZone: "America/Manaus",
  });

const dataHoraGeradoPdf = agoraPdf.toLocaleString("pt-BR", {
  timeZone: "America/Manaus",
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
  hour: "2-digit",
  minute: "2-digit",
  second: "2-digit",
});

const docDefinition = {
  pageSize: "A4",
pageOrientation: "landscape",
pageMargins: [margemEsquerdaPdf, 24, margemDireitaPdf, 24],
    footer: function (currentPage, pageCount) {
      return {
        margin: [16, 4, 16, 0],
        columns: [
          {
            text: "RealStock • Relatório de Inventário",
            alignment: "left",
            fontSize: 8,
            color: "#64748b",
          },
          {
            text: `Página ${currentPage} de ${pageCount}`,
            alignment: "right",
            fontSize: 8,
            color: "#64748b",
          },
        ],
      };
    },

    content: [
      {
        columns: [
          [
            {
              text: modoOperacao === "sem-base"
  ? "RELATÓRIO DE INVENTÁRIO - MODO SEM BASE"
  : "RELATÓRIO DE INVENTÁRIO",
              style: "header",
            },
            {
              text: "Resumo operacional dos itens filtrados no sistema",
              style: "subheader",
            },
          ],
          [
            {
              text: `Gerado em: ${dataHoraGeradoPdf}`,
              alignment: "right",
              fontSize: 8,
              color: "#475569",
              margin: [0, 6, 0, 0],
            },
          ],
        ],
        margin: [0, 0, 0, 8],
      },

      {
        table: {
          widths: ["*", "*", "*"],
          body: [
            [
              { text: `Cliente\n${cliente}`, style: "infoBox" },
              { text: `Loja\n${loja}`, style: "infoBox" },
              { text: `Data do Inventário\n${dataInventario}`, style: "infoBox" },
            ],
          ],
        },
        layout: "noBorders",
        margin: [0, 0, 0, 8],
      },

      ...(modoOperacao === "sem-base"
  ? [
      {
        table: {
          widths: ["*", "*"],
          body: [
            [
              { text: `Produtos\n${totalProdutos}`, style: "metricBox" },
              {
                text: `Itens Contados\n${resultado.filter((item) => Number(item.qtdeContada || 0) > 0).length}`,
                style: "metricBox",
              },
            ],
          ],
        },
        layout: "noBorders",
        margin: [0, 0, 0, 10],
      },
    ]
  : [
      {
        table: {
          widths: ["*", "*", "*", "*"],
          body: [
            [
              { text: `Produtos\n${totalProdutos}`, style: "metricBox" },
              {
                text: `Itens com Divergência\n${totalItensComDivergencia}`,
                style: "metricBox",
              },
              {
                text: `Valor Congelado\nR$ ${formatarValorPdf(totalValorCongelado)}`,
                style: "metricBox",
              },
              {
                text: `Valor Divergência\nR$ ${formatarValorPdf(totalValorDivergencia)}`,
                style: "metricBox",
              },
            ],
          ],
        },
        layout: "noBorders",
        margin: [0, 0, 0, 10],
      },
    ]),

      {
        table: {
          headerRows: 1,
          widths: largurasTabelaPdf,
          body: tableBody,
        },
        layout: {
          paddingLeft: function () {
            return 2;
          },
          paddingRight: function () {
            return 2;
          },
          paddingTop: function () {
            return 4;
          },
          paddingBottom: function () {
            return 4;
          },
          fillColor: function (rowIndex) {
            return rowIndex === 0 ? "#1e293b" : rowIndex % 2 === 0 ? "#f8fafc" : null;
          },
          hLineColor: function () {
            return "#cbd5e1";
          },
          vLineColor: function () {
            return "#cbd5e1";
          },
          hLineWidth: function () {
            return 0.5;
          },
          vLineWidth: function () {
            return 0.5;
          },
        },
      },
      ...(modoOperacao === "sem-base"
  ? []
  : [
      {
        margin: [0, 10, 0, 0],
        table: {
          widths: ["*", 120],
          body: [
            [
              { text: "RESUMO FINANCEIRO", style: "summaryTitle", colSpan: 2 },
              {},
            ],
            [
              { text: "Valor congelado total", style: "summaryLabel" },
              {
                text: `R$ ${formatarValorPdf(totalValorCongelado)}`,
                style: "summaryValue",
                alignment: "right",
              },
            ],
            [
              { text: "Valor contado total", style: "summaryLabel" },
              {
                text: `R$ ${formatarValorPdf(totalValorContado)}`,
                style: "summaryValue",
                alignment: "right",
              },
            ],
            [
              { text: "Valor divergente total", style: "summaryLabel" },
              {
                text: `R$ ${formatarValorPdf(totalValorDivergencia)}`,
                style: "summaryValue",
                alignment: "right",
              },
            ],
          ],
        },
        layout: {
          fillColor: function (rowIndex) {
            return rowIndex === 0 ? "#e2e8f0" : "#ffffff";
          },
          hLineColor: function () {
            return "#cbd5e1";
          },
          vLineColor: function () {
            return "#cbd5e1";
          },
          hLineWidth: function () {
            return 0.5;
          },
          vLineWidth: function () {
            return 0.5;
          },
          paddingLeft: function () {
            return 6;
          },
          paddingRight: function () {
            return 6;
          },
          paddingTop: function () {
            return 3;
          },
          paddingBottom: function () {
            return 3;
          },
        },
      },
    ]),
  ],
    styles: {
      header: {
        fontSize: 16,
        bold: true,
        color: "#0f172a",
      },
      subheader: {
        fontSize: 8,
        color: "#475569",
        margin: [0, 2, 0, 0],
      },
      infoBox: {
        fontSize: 8.5,
        bold: true,
        color: "#0f172a",
        fillColor: "#eff6ff",
        margin: [6, 6, 6, 6],
      },
      metricBox: {
        fontSize: 9,
        bold: true,
        color: "#0f172a",
        fillColor: "#f8fafc",
        margin: [6, 6, 6, 6],
        alignment: "center",
      },
      tableHeader: {
        bold: true,
        fontSize: fonteTabelaPdf + 1.5,
        color: "#ffffff",
        fillColor: "#1e293b",
        margin: [2, 6, 2, 6],
      },
      tableCell: {
        fontSize: fonteTabelaPdf,
        color: "#0f172a",
        margin: [1, 2, 1, 2],
      },
      summaryTitle: {
        bold: true,
        fontSize: 9.5,
        color: "#0f172a",
      },
      summaryLabel: {
        fontSize: 9,
        bold: true,
        color: "#334155",
      },
      summaryValue: {
        fontSize: 9,
        bold: true,
        color: "#0f172a",
      },
    },

    defaultStyle: {
      font: "Helvetica",
    },
  };

  const pdfDoc = printer.createPdfKitDocument(docDefinition);
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    'inline; filename="Relatorio-inventario.pdf"'
  );
  pdfDoc.pipe(res);
  pdfDoc.end();
});

app.post("/adicionar-produto", autenticar, (req, res) => {
  const { codigoBarras, codigoInterno } = req.body;

  if (!codigoBarras) {
    return res.status(400).json({ mensagem: "Código de barras obrigatório" });
  }

  const index = inventario.findIndex((p) => p.codigoBarras === codigoBarras);

  if (index >= 0) {
    inventario[index].codigoInterno =
      codigoInterno || inventario[index].codigoInterno || inventario[index].codigo;

    const valorAntigo = Number(inventario[index].qtdeContada) || 0;
    inventario[index].qtdeContada = valorAntigo + 1;

    registrarAlteracao(
      req.session.usuario,
      codigoBarras,
      "qtdeContada",
      valorAntigo,
      inventario[index].qtdeContada
    );

    registrarContagem(req.session.usuario, codigoBarras, 1);
  } else {
    inventario.push({
      codigoBarras,
      codigo: codigoInterno || "",
      codigoInterno: codigoInterno || "",
      qtdeContada: 1,
      qtdeCongelada: 0,
      custoUnitario: 0,
      descricao: "",
      categoria: "",
      tipo: "",
    });

    registrarAlteracao(
      req.session.usuario,
      codigoBarras,
      "qtdeContada",
      0,
      1
    );

    registrarContagem(req.session.usuario, codigoBarras, 1);
  }

  broadcastInventario();
  res.json({ mensagem: "Produto adicionado/atualizado com sucesso" });
});

app.get("/exportar-curva-abc-pdf", autenticar, (req, res) => {
  try {
    let resultado = filtrarInventario({
      categoria: req.query.categoria || "",
      ordem: req.query.ordem || "",
      busca: req.query.busca || "",
    });

    resultado = resultado.filter((item) => parseQuantidade(item.qtdeCongelada) > 0);

    const cliente = req.query.cliente || "Cliente não informado";
    const loja = req.query.loja || "Loja não informada";
    const codigoLoja = req.query.codigoLoja || "";
    const dataInventario = req.query.data || new Date().toLocaleDateString("pt-BR");
    const cidade = req.query.cidade || "";
    const uf = req.query.uf || "";
    const responsavelCliente = req.query.responsavelCliente || "Não informado";
    const liderOperacao = req.query.liderOperacao || "Não informado";
    const observacoes = req.query.observacoes || "";

    const printer = new PdfPrinter(fonts);
    const curva = calcularCurvaABC(resultado);
    const itensABC = curva.itens;
    const resumoClasses = resumirCurvaPorClasse(itensABC);
    const resumoDivergencias = resumirDivergencias(resultado);
    const categoriasResumo = resumirCategoriasDivergencia(resultado);

    const totalProdutos = resultado.length;

    const totalValorCongelado = resultado.reduce(
      (acc, item) => acc + (Number(item.valorCongelado) || 0),
      0
    );

    const totalValorContado = resultado.reduce(
      (acc, item) => acc + (Number(item.valorContado) || 0),
      0
    );

    const totalValorDivergenciaLiquido = resultado.reduce(
      (acc, item) => acc + (Number(item.valorDivergencia) || 0),
      0
    );

    const totalValorDivergenciaAbsoluto = resultado.reduce(
      (acc, item) => acc + Math.abs(Number(item.valorDivergencia) || 0),
      0
    );

    const itensComDivergencia = resultado.filter(
      (item) => Number(item.divergencia) !== 0
    ).length;

    const valorSobras = Number(resumoDivergencias.sobraValor) || 0;
    const valorFaltas = Number(resumoDivergencias.faltaValor) || 0;
    const resultadoLiquido = valorSobras - valorFaltas;

    const qtdSobras = Number(resumoDivergencias.sobraQtd) || 0;
    const qtdFaltas = Number(resumoDivergencias.faltaQtd) || 0;
    const resultadoQtd = qtdSobras - qtdFaltas;

    const classeA = resumoClasses.find((i) => i.classe === "A") || { itens: 0, valor: 0, percentual: 0 };
    const classeB = resumoClasses.find((i) => i.classe === "B") || { itens: 0, valor: 0, percentual: 0 };
    const classeC = resumoClasses.find((i) => i.classe === "C") || { itens: 0, valor: 0, percentual: 0 };

    const topCategorias = categoriasResumo.slice(0, 8);

    const bodyCategorias = [
      [
        { text: "Categoria", style: "tableHeader" },
        { text: "Sobra (Qtd.)", style: "tableHeader", alignment: "right" },
        { text: "Falta (Qtd.)", style: "tableHeader", alignment: "right" },
        { text: "Sobra (R$)", style: "tableHeader", alignment: "right" },
        { text: "Falta (R$)", style: "tableHeader", alignment: "right" },
        { text: "Resultado (R$)", style: "tableHeader", alignment: "right" },
      ],
    ];

    if (topCategorias.length === 0) {
      bodyCategorias.push([
        {
          text: "Nenhuma divergência por categoria encontrada.",
          colSpan: 6,
          alignment: "center",
          margin: [0, 10, 0, 10],
        },
        {},
        {},
        {},
        {},
        {},
      ]);
    } else {
      topCategorias.forEach((item) => {
        bodyCategorias.push([
          { text: item.categoria || "SEM CATEGORIA", style: "tableCell" },
          { text: formatarNumeroPdf(item.sobraQtd), style: "tableCell", alignment: "right" },
          { text: formatarNumeroPdf(item.faltaQtd), style: "tableCell", alignment: "right" },
          { text: `R$ ${formatarValorPdf(item.sobraValor)}`, style: "tableCell", alignment: "right" },
          { text: `R$ ${formatarValorPdf(item.faltaValor)}`, style: "tableCell", alignment: "right" },
          {
            text: `R$ ${formatarValorPdf(item.resultadoValor)}`,
            style: "tableCell",
            alignment: "right",
          },
        ]);
      });
    }

    const docDefinition = {
      pageSize: "A4",
      pageOrientation: "portrait",
      pageMargins: [28, 24, 28, 30],

      footer: function (currentPage, pageCount) {
        return {
          margin: [28, 6, 28, 0],
          columns: [
            {
              text: "RealStock • Curva ABC Financeira",
              alignment: "left",
              fontSize: 8,
              color: "#64748b",
            },
            {
              text: `Página ${currentPage} de ${pageCount}`,
              alignment: "right",
              fontSize: 8,
              color: "#64748b",
            },
          ],
        };
      },

      content: [
        {
          stack: [
            {
              canvas: [
                {
                  type: "rect",
                  x: 0,
                  y: 0,
                  w: 539,
                  h: 92,
                  r: 12,
                  color: "#0f172a",
                },
              ],
            },
            {
              absolutePosition: { x: 42, y: 42 },
              text: "REALSTOCK",
              fontSize: 23,
              bold: true,
              color: "#ffffff",
            },
            {
              absolutePosition: { x: 42, y: 70 },
              text: "RELATÓRIO DE CURVA ABC",
              fontSize: 15,
              bold: true,
              color: "#93c5fd",
            },
            {
              absolutePosition: { x: 42, y: 92 },
              text: "Resumo executivo financeiro do inventário",
              fontSize: 9,
              color: "#cbd5e1",
            },
          ],
          margin: [0, 0, 0, 86],
        },

        {
          table: {
            widths: ["*", "*"],
            body: [
              [
                { text: `Cliente\n${cliente}`, style: "infoBox" },
                { text: `Loja\n${loja}${codigoLoja ? ` (${codigoLoja})` : ""}`, style: "infoBox" },
              ],
              [
                { text: `Data do Inventário\n${dataInventario}`, style: "infoBox" },
                { text: `Local\n${cidade || "--"}${uf ? ` - ${uf}` : ""}`, style: "infoBox" },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 12],
        },

        {
          text: "RESUMO EXECUTIVO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            widths: ["*", "*"],
            body: [
              [
                { text: `Produtos Base\n${formatarNumeroPdf(totalProdutos)}`, style: "metricBoxDark" },
                { text: `Itens com Divergência\n${formatarNumeroPdf(itensComDivergencia)}`, style: "metricBoxDark" },
              ],
              [
                { text: `Valor Congelado\nR$ ${formatarValorPdf(totalValorCongelado)}`, style: "metricBoxLight" },
                { text: `Valor Contado\nR$ ${formatarValorPdf(totalValorContado)}`, style: "metricBoxLight" },
              ],
              [
                { text: `Divergência Líquida\nR$ ${formatarValorPdf(totalValorDivergenciaLiquido)}`, style: "metricBoxLight" },
                { text: `Divergência Absoluta\nR$ ${formatarValorPdf(totalValorDivergenciaAbsoluto)}`, style: "metricBoxLight" },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 14],
        },

        {
          text: "COMPARATIVO FINANCEIRO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            widths: ["*", 28, "*", 28, "*"],
            body: [
              [
                {
                  stack: [
                    { text: "SOBRAS", style: "compareTitlePositive" },
                    { text: `R$ ${formatarValorPdf(valorSobras)}`, style: "compareValuePositive" },
                    { text: `Qtd.: ${formatarNumeroPdf(qtdSobras)}`, style: "compareSub" },
                  ],
                  style: "compareBoxPositive",
                },
                {
                  text: "−",
                  alignment: "center",
                  bold: true,
                  fontSize: 20,
                  color: "#475569",
                  margin: [0, 22, 0, 0],
                },
                {
                  stack: [
                    { text: "FALTAS", style: "compareTitleNegative" },
                    { text: `R$ ${formatarValorPdf(valorFaltas)}`, style: "compareValueNegative" },
                    { text: `Qtd.: ${formatarNumeroPdf(qtdFaltas)}`, style: "compareSub" },
                  ],
                  style: "compareBoxNegative",
                },
                {
                  text: "=",
                  alignment: "center",
                  bold: true,
                  fontSize: 20,
                  color: "#475569",
                  margin: [0, 22, 0, 0],
                },
                {
                  stack: [
                    { text: "RESULTADO", style: "compareTitleResult" },
                    { text: `R$ ${formatarValorPdf(resultadoLiquido)}`, style: "compareValueResult" },
                    { text: `Qtd.: ${formatarNumeroPdf(resultadoQtd)}`, style: "compareSub" },
                  ],
                  style: "compareBoxResult",
                },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 10],
        },

        {
          text: "RESUMO DA CURVA ABC",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            widths: ["*", "*", "*"],
            body: [
              [
                {
                  text:
                    `CLASSE A\n\n` +
                    `Itens: ${formatarNumeroPdf(classeA.itens)}\n` +
                    `Valor: R$ ${formatarValorPdf(classeA.valor)}\n` +
                    `Participação: ${formatarValorPdf(classeA.percentual)}%`,
                  style: "abcBoxA",
                },
                {
                  text:
                    `CLASSE B\n\n` +
                    `Itens: ${formatarNumeroPdf(classeB.itens)}\n` +
                    `Valor: R$ ${formatarValorPdf(classeB.valor)}\n` +
                    `Participação: ${formatarValorPdf(classeB.percentual)}%`,
                  style: "abcBoxB",
                },
                {
                  text:
                    `CLASSE C\n\n` +
                    `Itens: ${formatarNumeroPdf(classeC.itens)}\n` +
                    `Valor: R$ ${formatarValorPdf(classeC.valor)}\n` +
                    `Participação: ${formatarValorPdf(classeC.percentual)}%`,
                  style: "abcBoxC",
                },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 10],
        },

        {
          text: "DIVERGÊNCIAS POR CATEGORIA",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            headerRows: 1,
            widths: ["*", 58, 58, 82, 82, 82],
            body: bodyCategorias,
          },
          layout: {
            fillColor: function (rowIndex) {
              if (rowIndex === 0) return "#0f172a";
              return rowIndex % 2 === 0 ? "#f8fafc" : "#ffffff";
            },
            hLineColor: function () {
              return "#cbd5e1";
            },
            vLineColor: function () {
              return "#cbd5e1";
            },
            hLineWidth: function () {
              return 0.5;
            },
            vLineWidth: function () {
              return 0.5;
            },
            paddingLeft: function () {
              return 6;
            },
            paddingRight: function () {
              return 6;
            },
            paddingTop: function () {
              return 5;
            },
            paddingBottom: function () {
              return 5;
            },
          },
          margin: [0, 0, 0, 10],
        },

        {
          text: "ASSINATURAS",
          style: "sectionTitle",
          margin: [0, 0, 0, 10],
        },

        {
          table: {
            widths: ["*", 24, "*"],
            body: [
              [
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: responsavelCliente, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável da loja", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
                {
                  text: "",
                  border: [false, false, false, false],
                },
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: liderOperacao, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável RealStock", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
              ],
            ],
          },
          layout: "noBorders",
        },
      ],

      styles: {
        header: {
          fontSize: 18,
          bold: true,
          color: "#0f172a",
        },
        subheader: {
          fontSize: 8.5,
          color: "#475569",
        },
        sectionTitle: {
          fontSize: 11,
          bold: true,
          color: "#0f172a",
        },
        infoBox: {
          fontSize: 9,
          bold: true,
          color: "#0f172a",
          fillColor: "#eff6ff",
          margin: [8, 8, 8, 8],
        },
        metricBoxDark: {
          fontSize: 9,
          bold: true,
          color: "#ffffff",
          fillColor: "#0f172a",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        metricBoxLight: {
          fontSize: 9,
          bold: true,
          color: "#0f172a",
          fillColor: "#f8fafc",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        compareBoxPositive: {
          fillColor: "#ecfdf5",
          margin: [6, 6, 6, 6],
        },
        compareBoxNegative: {
          fillColor: "#fef2f2",
          margin: [8, 10, 8, 10],
        },
        compareBoxResult: {
          fillColor: "#eff6ff",
          margin: [8, 10, 8, 10],
        },
        compareTitlePositive: {
          fontSize: 10,
          bold: true,
          color: "#065f46",
          alignment: "center",
        },
        compareTitleNegative: {
          fontSize: 10,
          bold: true,
          color: "#991b1b",
          alignment: "center",
        },
        compareTitleResult: {
          fontSize: 10,
          bold: true,
          color: "#1d4ed8",
          alignment: "center",
        },
        compareValuePositive: {
          fontSize: 13,
          bold: true,
          color: "#065f46",
          alignment: "center",
          margin: [0, 6, 0, 4],
        },
        compareValueNegative: {
          fontSize: 13,
          bold: true,
          color: "#991b1b",
          alignment: "center",
          margin: [0, 6, 0, 4],
        },
        compareValueResult: {
          fontSize: 13,
          bold: true,
          color: "#1d4ed8",
          alignment: "center",
          margin: [0, 6, 0, 4],
        },
        compareSub: {
          fontSize: 9,
          color: "#475569",
          alignment: "center",
        },
        abcBoxA: {
          fontSize: 10,
          bold: true,
          color: "#0f172a",
          fillColor: "#dbeafe",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        abcBoxB: {
          fontSize: 10,
          bold: true,
          color: "#0f172a",
          fillColor: "#fef3c7",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        abcBoxC: {
          fontSize: 10,
          bold: true,
          color: "#0f172a",
          fillColor: "#f1f5f9",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        tableHeader: {
          bold: true,
          fontSize: 9,
          color: "#ffffff",
        },
        tableCell: {
          fontSize: 8.7,
          color: "#0f172a",
        },
        approvalLabel: {
          fontSize: 9.2,
          bold: true,
          color: "#334155",
        },
        approvalValue: {
          fontSize: 10,
          bold: true,
          color: "#0f172a",
        },
        approvalText: {
          fontSize: 9,
          color: "#0f172a",
        },
      },

      defaultStyle: {
        font: "Helvetica",
      },
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="Curva-ABC-RealStock.pdf"'
    );
    pdfDoc.pipe(res);
    pdfDoc.end();
  } catch (erro) {
    console.error("Erro ao exportar Curva ABC PDF:", erro);
    res.status(500).send("Erro ao exportar Curva ABC PDF.");
  }
});

app.post("/editar", autenticar, (req, res) => {
  const { codigo, novaQuantidade } = req.body;
  const item = inventario.find((i) => i.codigoBarras === codigo);

  if (item) {
    const valorAntigo = Number(item.qtdeContada) || 0;
    const valorNovo = parseFloat(novaQuantidade) || 0;

    item.qtdeContada = valorNovo;

    registrarAlteracao(
      req.session.usuario,
      codigo,
      "qtdeContada",
      valorAntigo,
      valorNovo
    );

    const delta = valorNovo - valorAntigo;
    if (delta > 0) {
      registrarContagem(req.session.usuario, codigo, delta);
    }

    broadcastInventario();
    return res.status(200).json({ sucesso: true });
  }

  res.status(404).json({ erro: "Produto não encontrado" });
});
app.put('/editar-item-finalizacao', autenticar, (req, res) => {
  try {
    const { finalizacaoId, contagemId, novaQuantidade } = req.body;

    const quantidade = Number(novaQuantidade);

    if (!finalizacaoId || !contagemId) {
      return res.status(400).json({ erro: 'finalizacaoId e contagemId são obrigatórios.' });
    }

    if (Number.isNaN(quantidade) || quantidade < 0) {
      return res.status(400).json({ erro: 'Nova quantidade inválida.' });
    }

    let alterado = false;

    enderecamentos = enderecamentos.map((end) => {
      const finalizacoes = Array.isArray(
        endereco?.finalizacoes
      )
        ? endereco.finalizacoes.filter(
            (finalizacao) =>
              !finalizacao?.excluida &&
              eventoPertenceAoModo(
                finalizacao
              )
          )
        : [];
      const novasFinalizacoes = finalizacoes.map((fin) => {
        if (String(fin.id) !== String(finalizacaoId)) return fin;

        const itens = Array.isArray(fin.itens) ? fin.itens : [];

        const novosItens = itens.map((item) => {
          if (String(item.contagemId) !== String(contagemId)) return item;

          alterado = true;

          return {
            ...item,
            quantidade
          };
        });

        return {
          ...fin,
          itens: novosItens
        };
      });

      return {
        ...end,
        finalizacoes: novasFinalizacoes,
        atualizadoEm: new Date().toISOString()
      };
    });

    if (!alterado) {
      return res.status(404).json({ erro: 'Item não encontrado para edição.' });
    }

    salvarEnderecamentos();
    return res.json({
      sucesso: true,
      mensagem: 'Quantidade atualizada com sucesso.'
    });
  } catch (erro) {
    console.error('Erro ao editar item da finalização:', erro);
    return res.status(500).json({ erro: 'Falha ao editar item da finalização.' });
  }
});
app.post('/auditoria/corrigir-item', autenticar, (req, res) => {
  try {
    const { codigo, novaQuantidade, motivo, observacao, usuario } = req.body;

    if (!codigo) {
      return res.status(400).json({ erro: 'Código não informado' });
    }

    const quantidadeNumero = Number(novaQuantidade);
    if (isNaN(quantidadeNumero)) {
      return res.status(400).json({ erro: 'Quantidade inválida' });
    }

    const agora = new Date().toISOString();
    if (
      modoOperacao ===
      "sem-base"
    ) {
      const indice =
        contagemSemBase.findIndex(
          (item) =>
            String(
              item.ean ||
              item.codigo ||
              ""
            ).trim() ===
            String(codigo).trim()
        );
    
      if (indice === -1) {
        return res.status(404).json({
          erro:
            "Item não encontrado na contagem sem base.",
        });
      }
    
      const item =
        contagemSemBase[
          indice
        ];
    
      if (
        !Array.isArray(
          item.enderecos
        )
      ) {
        item.enderecos = [];
      }
    
      const quantidadeAnterior =
        Number(
          item.quantidade
        ) || 0;
    
      const diferenca =
        quantidadeNumero -
        quantidadeAnterior;
    
      /*
        A correção será aplicada ao endereço
        da leitura mais recente. Isso mantém
        o total e o detalhamento sincronizados.
      */
      const enderecosOrdenados =
        [...item.enderecos]
          .filter(
            (registro) =>
              registro &&
              registro.enderecoNumero
          )
          .sort(
            (a, b) =>
              new Date(
                b.data || 0
              ) -
              new Date(
                a.data || 0
              )
          );
    
      const enderecoMaisRecente =
        enderecosOrdenados[0];
    
      if (!enderecoMaisRecente) {
        return res.status(400).json({
          erro:
            "O item não possui endereço associado para receber a correção.",
        });
      }
    
      enderecoMaisRecente.quantidade =
        Math.max(
          0,
          (
            Number(
              enderecoMaisRecente.quantidade
            ) || 0
          ) + diferenca
        );
    
      enderecoMaisRecente.usuario =
        usuario ||
        req.session?.usuario?.usuario ||
        req.session?.usuario?.nome ||
        "auditoria";
    
      enderecoMaisRecente.data =
        agora;
    
      item.quantidade =
        item.enderecos.reduce(
          (total, registro) =>
            total +
            (
              Number(
                registro.quantidade
              ) || 0
            ),
          0
        );
    
      item.ultimoUsuario =
        enderecoMaisRecente.usuario;
    
      item.data =
        agora;
    
      historicoAuditoriaItens.push({
        tipo:
          "correcao-sem-base",
    
        codigoBarras:
          item.ean || "",
    
        codigo:
          item.codigo || "",
    
        quantidadeAnterior,
    
        novaQuantidade:
          item.quantidade,
    
        enderecoNumero:
          enderecoMaisRecente
            .enderecoNumero,
    
        motivo:
          motivo || "",
    
        observacao:
          observacao || "",
    
        usuario:
          enderecoMaisRecente
            .usuario,
    
        data:
          agora,
      });
    
      salvarContagemSemBase()
        .then(() => {
          broadcastInventario();
        })
        .catch((erroSalvar) => {
          console.error(
            "Erro ao salvar correção sem base:",
            erroSalvar
          );
        });
    
      return res.json({
        ok: true,
        sucesso: true,
        quantidade:
          item.quantidade,
      });
    }

    // 🔴 GARANTE ARRAY
    if (!Array.isArray(contagens)) contagens = [];
    if (!Array.isArray(historicoAuditoriaItens)) historicoAuditoriaItens = [];

    // 🔴 DESATIVA CONTAGENS ANTIGAS
    for (let i = 0; i < contagens.length; i++) {
      const c = contagens[i];

      if (
        c &&
        c.ativo !== false &&
        String(c.codigoBarras) === String(codigo)
      ) {
        contagens[i] = {
          ...c,
          ativo: false,
          substituidaPorAuditoria: true,
          dataSubstituicao: agora
        };
      }
    }

    // 🔴 NOVA CONTAGEM OFICIAL
    const novaContagem = {
      id: `CORR-${Date.now()}`,
      codigoBarras: codigo,
      quantidade: quantidadeNumero,
      usuario: usuario || 'auditoria',
      data: agora,
      origem: 'auditoria',
      ativo: true,
      enderecoNumero: null,
      motivoCorrecao: motivo || '',
      observacaoCorrecao: observacao || ''
    };

    contagens.push(novaContagem);

    salvarContagens();

    // 🔴 ATUALIZA INVENTÁRIO CORRETAMENTE
    for (let i = 0; i < inventario.length; i++) {
      if (String(inventario[i].codigoBarras) === String(codigo)) {
        inventario[i].qtdeContada = quantidadeNumero;
      }
    }

    // 🔴 LOG DE AUDITORIA
    historicoAuditoriaItens.push({
      codigoBarras: codigo,
      novaQuantidade: quantidadeNumero,
      motivo: motivo || '',
      observacao: observacao || '',
      usuario: usuario || 'auditoria',
      data: agora
    });

    return res.json({ ok: true });

  } catch (erro) {
    console.error('🔥 ERRO REAL NA ROTA:', erro);
    return res.status(500).json({ erro: erro.message });
  }
});
app.delete('/excluir-finalizacao/:finalizacaoId', autenticar, async (req, res) => {
  try {
    const finalizacaoId = String(req.params.finalizacaoId || '').trim();

    if (!finalizacaoId) {
      return res.status(400).json({ erro: 'finalizacaoId é obrigatório.' });
    }

    let encontrou = false;
    const agoraIso = new Date().toISOString();
    const usuarioExclusao =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      'sistema';

    enderecamentos = enderecamentos.map((endereco) => {
      const finalizacoes = Array.isArray(endereco.finalizacoes) ? endereco.finalizacoes : [];
      const transmissoes = Array.isArray(endereco.transmissoes) ? endereco.transmissoes : [];

      const novasFinalizacoes = finalizacoes.map((fin) => {
        if (String(fin.id) !== finalizacaoId || fin.excluida) return fin;

        encontrou = true;

        return {
          ...fin,
          excluida: true,
          excluidaEm: agoraIso,
          excluidaPor: usuarioExclusao,
        };
      });

      const novasTransmissoes = transmissoes.map((trans, index) => {

        const idReal = String(trans.id || '').trim();
const idVirtual = `TRANS-${Number(endereco.id)}-${Number(trans.enderecoNumero)}-${index}`;

if (
  trans.excluida ||
  (
    idReal !== finalizacaoId &&
    idVirtual !== finalizacaoId
  )
) {
  return trans;
}
        encontrou = true;

        return {
          ...trans,
          excluida: true,
          statusConsolidacao: 'excluida',
          excluidaEm: agoraIso,
          excluidaPor: usuarioExclusao,
        };
      });

      const enderecoAtualizado = {
        ...endereco,
        finalizacoes: novasFinalizacoes,
        transmissoes: novasTransmissoes,
        atualizadoEm: agoraIso,
      };

      const resumoFaixa = recalcularStatusFaixa(enderecoAtualizado);

      return {
        ...enderecoAtualizado,
        status: resumoFaixa.status,
        totalPosicoes: resumoFaixa.totalPosicoes,
        posicoesConcluidas: resumoFaixa.concluidos,
        posicoesPendentes: resumoFaixa.pendentes,
        posicoesEmContagem: resumoFaixa.emContagem,
        posicoesDuplicadas: resumoFaixa.duplicados,
        contagensRecebidas: novasTransmissoes.filter(
          (t) => t.tipo === 'transmissao' && !t.excluida
        ).length,
      };
    });

    if (!encontrou) {
      return res.status(404).json({ erro: 'Finalização/transmissão não encontrada.' });
    }

    await salvarEnderecamentos();

    if (
      modoOperacao !== "sem-base"
    ) {
      recalcularInventarioComBaseNasContagens();
    
      await salvarProdutosNoBanco(
        inventario
      );
    
      broadcastInventario();
    }
const painelAtualizado = gerarPainelTransmissoesConsolidacao();

    return res.json({
      sucesso: true,
      mensagem: 'Finalização excluída com sucesso.',
      painel: painelAtualizado,
    });
  } catch (erro) {
    console.error('Erro ao excluir finalização:', erro);
    return res.status(500).json({ erro: 'Erro ao excluir finalização.' });
  }
});
app.post("/encerrar-inventario", autenticar, async (req, res) => {
  try {
    garantirPastaEncerramentos();

    const modoEncerrado =
      normalizarModoOperacao(
        modoOperacao
      );

    if (
      !haAtividadeNoModo(
        modoEncerrado
      )
    ) {
      return res.status(400).json({
        erro:
          `Não há inventário ativo no modo ${modoEncerrado} para encerrar.`
      });
    }

    const timestamp = gerarTimestampEncerramento();
    const nomePasta = `encerramento-${timestamp}`;
    const pastaEncerramento = path.join(encerramentosDir, nomePasta);

    fs.mkdirSync(pastaEncerramento, { recursive: true });

    const snapshot = montarSnapshotEncerramento(req.session.usuario);

    fs.writeFileSync(
      path.join(pastaEncerramento, "snapshot.json"),
      JSON.stringify(snapshot, null, 2),
      "utf8"
    );

    const linhasEncerramento =
  modoEncerrado === "sem-base"
    ? montarLinhasSemBaseParaTabela({
        busca: ""
      })
    : inventario;

fs.writeFileSync(
  path.join(
    pastaEncerramento,
    "inventario-final.csv"
  ),
  gerarCsvInventario(
    linhasEncerramento
  ),
  "utf8"
);

    const resumoTxt = [
      `Encerrado em: ${snapshot.encerradoEm}`,
      `Encerrado por: ${snapshot.encerradoPor}`,
      `Total de itens: ${snapshot.resumo.totalItens}`,
      `Total de alterações: ${snapshot.resumo.totalAlteracoes}`,
      `Total de auditorias: ${snapshot.resumo.totalAuditorias}`,
      `Total de contagens: ${snapshot.resumo.totalContagens}`,
      `Total de endereçamentos: ${snapshot.resumo.totalEnderecamentos}`
    ].join("\n");

    fs.writeFileSync(
      path.join(pastaEncerramento, "resumo.txt"),
      resumoTxt,
      "utf8"
    );

    fs.writeFileSync(
      ultimoEncerramentoPath,
      JSON.stringify({
        nomePasta,
        snapshotPath: path.join(pastaEncerramento, "snapshot.json"),
        encerradoEm: snapshot.encerradoEm
      }, null, 2),
      "utf8"
    );

    await resetarSistemaAposEncerramento(
      modoEncerrado
    );
    
    broadcastInventario();
    
    return res.json({
      sucesso: true,
      nomePasta,
      mensagem:
        `Inventário do modo ${modoEncerrado} encerrado com sucesso.`
    });

  } catch (erro) {
    console.error("Erro ao encerrar inventário:", erro);
    return res.status(500).json({
      erro: "Erro ao encerrar inventário."
    });
  }
});
app.post("/restaurar-ultimo-encerramento", autenticar, async (req, res) => {
  try {
    garantirPastaEncerramentos();

    if (!fs.existsSync(ultimoEncerramentoPath)) {
      return res.status(404).json({
        erro: "Nenhum encerramento anterior foi encontrado."
      });
    }

    const ultimo = JSON.parse(fs.readFileSync(ultimoEncerramentoPath, "utf8") || "{}");
    const snapshotPath = ultimo.snapshotPath;

    if (!snapshotPath || !fs.existsSync(snapshotPath)) {
      return res.status(404).json({
        erro: "Snapshot do último encerramento não encontrado."
      });
    }

    const snapshot = JSON.parse(fs.readFileSync(snapshotPath, "utf8") || "{}");

    inventario = Array.isArray(snapshot.inventario) ? snapshot.inventario : [];
    historicoAlteracoes = Array.isArray(snapshot.historicoAlteracoes) ? snapshot.historicoAlteracoes : [];
    historicoAuditoriaItens = Array.isArray(snapshot.historicoAuditoriaItens) ? snapshot.historicoAuditoriaItens : [];
    contagens = Array.isArray(snapshot.contagens) ? snapshot.contagens : [];
    enderecamentos = Array.isArray(snapshot.enderecamentos)
      ? snapshot.enderecamentos.map(normalizarEnderecoSalvo)
      : [];

    auditoriaImportacao = snapshot.auditoriaImportacao || {
      totalImportadoBruto: 0,
      totalUnicosBruto: 0,
      duplicatasRemovidas: 0,
      itensZeradosIgnorados: 0,
    };

    itemAuditoriaAtual = null;

    await salvarProdutosNoBanco(inventario);
    await salvarContagens();
    await salvarEnderecamentos();

    broadcastInventario();

    return res.json({
      sucesso: true,
      nomePasta: ultimo.nomePasta || null,
      mensagem: "Último encerramento restaurado com sucesso."
    });
  } catch (erro) {
    console.error("Erro ao restaurar último encerramento:", erro);
    return res.status(500).json({
      erro: "Erro ao restaurar último encerramento."
    });
  }
});
app.get(
  "/exportar-checklist-encerramento-pdf-auto",
  autenticar,
  (req, res) => {
    try {
      /*
        Tamanho da fonte utilizado exclusivamente
        nas tabelas do checklist de encerramento.
      */
      const fonteTabelaPdf = 9;

      const printer =
        new PdfPrinter(fonts);

    const cliente = req.query.cliente || "Cliente não informado";
    const loja = req.query.loja || "Loja não informada";
    const codigoLoja = req.query.codigoLoja || "";
    const dataInventario = req.query.dataInventario || "";
    const cidade = req.query.cidade || "";
    const uf = req.query.uf || "";
    const responsavelCliente = req.query.responsavelCliente || "Não informado";
    const responsavelRealstock = req.query.responsavelRealstock || "Não informado";
    const dataEncerramento = req.query.dataEncerramento || "";
    const horaEncerramento = req.query.horaEncerramento || "";

    const checklist = [
      { label: "Contagem finalizada pela equipe", marcado: true },
      { label: "Itens sem contagem revisados", marcado: true },
      { label: "Divergências revisadas", marcado: true },
      { label: "Categorias críticas analisadas", marcado: true },
      { label: "Itens classe A revisados", marcado: true },
      { label: "Cliente acompanhou o fechamento", marcado: true },
      { label: "Números apresentados ao cliente", marcado: true },
      { label: "Responsável RealStock validou a operação", marcado: true },
      { label: "Operação autorizada para encerramento", marcado: true },
    ];

    const totalMarcados = checklist.filter((item) => item.marcado).length;

    const bodyChecklist = [
      [
        { text: "Status", style: "tableHeader", alignment: "center" },
        { text: "Item de verificação", style: "tableHeader" },
      ],
      ...checklist.map((item) => [
        {
          text: item.marcado ? "CONCLUÍDO" : "PENDENTE",
          style: item.marcado ? "statusOk" : "statusPendente",
          alignment: "center",
        },
        {
          text: item.label,
          style: "tableCell",
        },
      ]),
    ];

    const docDefinition = {
      pageSize: "A4",
      pageOrientation: "portrait",
      pageMargins: [28, 24, 28, 30],

      footer: function (currentPage, pageCount) {
        return {
          margin: [28, 6, 28, 0],
          columns: [
            {
              text: "RealStock • Checklist de Encerramento",
              alignment: "left",
              fontSize: 8,
              color: "#64748b",
            },
            {
              text: `Página ${currentPage} de ${pageCount}`,
              alignment: "right",
              fontSize: 8,
              color: "#64748b",
            },
          ],
        };
      },

      content: [
        {
          stack: [
            {
              canvas: [
                {
                  type: "rect",
                  x: 0,
                  y: 0,
                  w: 539,
                  h: 92,
                  r: 12,
                  color: "#0f172a",
                },
              ],
            },
            {
              absolutePosition: { x: 42, y: 42 },
              text: "REALSTOCK",
              fontSize: 23,
              bold: true,
              color: "#ffffff",
            },
            {
              absolutePosition: { x: 42, y: 70 },
              text: "CHECKLIST DE ENCERRAMENTO",
              fontSize: 15,
              bold: true,
              color: "#93c5fd",
            },
            {
              absolutePosition: { x: 42, y: 92 },
              text: "Fechamento formal da operação de inventário",
              fontSize: 9,
              color: "#cbd5e1",
            },
          ],
          margin: [0, 0, 0, 86],
        },

        {
          table: {
            widths: ["*", "*"],
            body: [
              [
                { text: `Cliente\n${cliente}`, style: "infoBox" },
                { text: `Loja\n${loja}${codigoLoja ? ` (${codigoLoja})` : ""}`, style: "infoBox" },
              ],
              [
                { text: `Data do Inventário\n${dataInventario || "--"}`, style: "infoBox" },
                { text: `Local\n${cidade || "--"}${uf ? ` - ${uf}` : ""}`, style: "infoBox" },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 12],
        },

        {
          text: "DADOS DO ENCERRAMENTO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            widths: ["*", "*"],
            body: [
              [
                { text: `Status\nENCERRADO`, style: "metricBoxDark" },
                { text: `Checklist concluído\n${totalMarcados}/9 itens marcados`, style: "metricBoxDark" },
              ],
              [
                { text: `Data do Encerramento\n${dataEncerramento || "--"}`, style: "metricBoxLight" },
                { text: `Hora do Encerramento\n${horaEncerramento || "--:--"}`, style: "metricBoxLight" },
              ],
              [
                { text: `Responsável RealStock\n${responsavelRealstock}`, style: "metricBoxLight" },
                { text: `Responsável da loja\n${responsavelCliente}`, style: "metricBoxLight" },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 14],
        },

        {
          text: "CHECKLIST DA OPERAÇÃO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            headerRows: 1,
            widths: [92, "*"],
            body: bodyChecklist,
          },
          layout: {
            fillColor: function (rowIndex) {
              if (rowIndex === 0) return "#0f172a";
              return rowIndex % 2 === 0 ? "#f8fafc" : "#ffffff";
            },
            hLineColor: function () {
              return "#cbd5e1";
            },
            vLineColor: function () {
              return "#cbd5e1";
            },
            hLineWidth: function () {
              return 0.5;
            },
            vLineWidth: function () {
              return 0.5;
            },
            paddingLeft: function () {
              return 6;
            },
            paddingRight: function () {
              return 6;
            },
            paddingTop: function () {
              return 6;
            },
            paddingBottom: function () {
              return 6;
            },
          },
          margin: [0, 0, 0, 18],
        },

        {
          text: "ASSINATURAS",
          style: "sectionTitle",
          margin: [0, 0, 0, 10],
        },

        {
          table: {
            widths: ["*", 24, "*"],
            body: [
              [
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: responsavelCliente, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável da loja", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
                {
                  text: "",
                  border: [false, false, false, false],
                },
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: responsavelRealstock, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável RealStock", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
              ],
            ],
          },
          layout: "noBorders",
        },
      ],

      styles: {
        sectionTitle: {
          fontSize: 11,
          bold: true,
          color: "#0f172a",
        },
        infoBox: {
          fontSize: 9,
          bold: true,
          color: "#0f172a",
          fillColor: "#eff6ff",
          margin: [8, 8, 8, 8],
        },
        metricBoxDark: {
          fontSize: 9,
          bold: true,
          color: "#ffffff",
          fillColor: "#0f172a",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        metricBoxLight: {
          fontSize: 9,
          bold: true,
          color: "#0f172a",
          fillColor: "#f8fafc",
          alignment: "center",
          margin: [6, 6, 6, 6],
        },
        tableHeader: {
          bold: true,
          fontSize: fonteTabelaPdf,
          color: "#ffffff",
          margin: [2, 4, 2, 4],
        },
        tableCell: {
          fontSize: fonteTabelaPdf,
          color: "#0f172a",
          margin: [1, 2, 1, 2],
        },
        statusOk: {
          fontSize: 8.5,
          bold: true,
          color: "#065f46",
        },
        statusPendente: {
          fontSize: 8.5,
          bold: true,
          color: "#991b1b",
        },
      },

      defaultStyle: {
        font: "Helvetica",
      },
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="Checklist-Encerramento-RealStock.pdf"'
    );
    pdfDoc.pipe(res);
    pdfDoc.end();
  } catch (erro) {
    console.error("Erro ao exportar checklist automático PDF:", erro);
    res.status(500).send("Erro ao exportar checklist automático PDF.");
  }
});

app.get("/exportar-avaliacao-inventario-pdf", autenticar, (req, res) => {
  try {
    const printer = new PdfPrinter(fonts);

    const cliente = req.query.cliente || "Cliente não informado";
    const loja = req.query.loja || "Loja não informada";
    const codigoLoja = req.query.codigoLoja || "";
    const dataInventario = req.query.dataInventario || "";
    const cidade = req.query.cidade || "";
    const uf = req.query.uf || "";
    const responsavelCliente = req.query.responsavelCliente || "Não informado";
    const responsavelRealstock = req.query.responsavelRealstock || "Não informado";

    const perguntas = [
      "Organização geral da operação de inventário",
      "Postura e profissionalismo da equipe durante o trabalho",
      "Clareza na comunicação entre equipe RealStock e cliente",
      "Confiança transmitida durante a execução do inventário",
      "Qualidade da condução da operação pelo líder RealStock",
      "Agilidade e eficiência da equipe durante a contagem",
      "Compromisso e pontualidade da equipe durante a operação",
      "Satisfação geral com o serviço prestado pela RealStock",
    ];

    const bodyAvaliacao = [
      [
        { text: "Pergunta", style: "tableHeader" },
        { text: "Ótimo", style: "tableHeader", alignment: "center" },
        { text: "Bom", style: "tableHeader", alignment: "center" },
        { text: "Regular", style: "tableHeader", alignment: "center" },
        { text: "Ruim", style: "tableHeader", alignment: "center" },
      ],
      ...perguntas.map((pergunta) => [
        { text: pergunta, style: "tableCell" },
        { text: "[   ]", style: "markCell", alignment: "center" },
        { text: "[   ]", style: "markCell", alignment: "center" },
        { text: "[   ]", style: "markCell", alignment: "center" },
        { text: "[   ]", style: "markCell", alignment: "center" },
      ]),
    ];

    const docDefinition = {
      pageSize: "A4",
      pageOrientation: "portrait",
      pageMargins: [28, 24, 28, 30],

      footer: function (currentPage, pageCount) {
        return {
          margin: [28, 6, 28, 0],
          columns: [
            {
              text: "RealStock • Avaliação do Inventário",
              alignment: "left",
              fontSize: 8,
              color: "#64748b",
            },
            {
              text: `Página ${currentPage} de ${pageCount}`,
              alignment: "right",
              fontSize: 8,
              color: "#64748b",
            },
          ],
        };
      },

      content: [
        {
          stack: [
            {
              canvas: [
                {
                  type: "rect",
                  x: 0,
                  y: 0,
                  w: 539,
                  h: 92,
                  r: 12,
                  color: "#0f172a",
                },
              ],
            },
            {
              absolutePosition: { x: 42, y: 42 },
              text: "REALSTOCK",
              fontSize: 23,
              bold: true,
              color: "#ffffff",
            },
            {
              absolutePosition: { x: 42, y: 70 },
              text: "AVALIAÇÃO DO INVENTÁRIO",
              fontSize: 15,
              bold: true,
              color: "#93c5fd",
            },
            {
              absolutePosition: { x: 42, y: 92 },
              text: "Relatório final de percepção e satisfação da operação",
              fontSize: 9,
              color: "#cbd5e1",
            },
          ],
          margin: [0, 0, 0, 86],
        },

        {
          table: {
            widths: ["*", "*"],
            body: [
              [
                { text: `Cliente\n${cliente}`, style: "infoBox" },
                { text: `Loja\n${loja}${codigoLoja ? ` (${codigoLoja})` : ""}`, style: "infoBox" },
              ],
              [
                { text: `Data do Inventário\n${dataInventario || "--"}`, style: "infoBox" },
                { text: `Local\n${cidade || "--"}${uf ? ` - ${uf}` : ""}`, style: "infoBox" },
              ],
            ],
          },
          layout: "noBorders",
          margin: [0, 0, 0, 12],
        },

        {
          text: "ORIENTAÇÃO DE PREENCHIMENTO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          text: "Solicitamos ao cliente que marque um X em apenas uma opção para cada pergunta, conforme sua percepção sobre a operação realizada.",
          style: "instructionBox",
          margin: [0, 0, 0, 14],
        },

        {
          text: "QUESTIONÁRIO DE AVALIAÇÃO",
          style: "sectionTitle",
          margin: [0, 0, 0, 8],
        },

        {
          table: {
            headerRows: 1,
            widths: ["*", 52, 52, 58, 52],
            body: bodyAvaliacao,
          },
          layout: {
            fillColor: function (rowIndex) {
              if (rowIndex === 0) return "#0f172a";
              return rowIndex % 2 === 0 ? "#f8fafc" : "#ffffff";
            },
            hLineColor: function () {
              return "#cbd5e1";
            },
            vLineColor: function () {
              return "#cbd5e1";
            },
            hLineWidth: function () {
              return 0.5;
            },
            vLineWidth: function () {
              return 0.5;
            },
            paddingLeft: function () {
              return 6;
            },
            paddingRight: function () {
              return 6;
            },
            paddingTop: function () {
              return 6;
            },
            paddingBottom: function () {
              return 6;
            },
          },
          margin: [0, 0, 0, 14],
        },
        {
          text: "OBSERVAÇÕES DO CLIENTE",
          style: "sectionTitle",
          margin: [0, 0, 0, 6],
        },

        {
          text:
            "________________________________________________________________________________________\n\n" +
            "________________________________________________________________________________________\n\n" +
            "________________________________________________________________________________________",
          style: "observationBox",
          margin: [0, 0, 0, 16],
        },
        {
          table: {
            widths: ["*", 90],
            body: [
              [
                { text: "Nota geral da operação", style: "summaryLabel" },
                { text: "_______ / 10", style: "summaryValue", alignment: "center" },
              ],
            ],
          },
          layout: {
            fillColor: function (rowIndex, node, columnIndex) {
              return columnIndex === 0 ? "#e2e8f0" : "#ffffff";
            },
            hLineColor: function () {
              return "#cbd5e1";
            },
            vLineColor: function () {
              return "#cbd5e1";
            },
            hLineWidth: function () {
              return 0.5;
            },
            vLineWidth: function () {
              return 0.5;
            },
            paddingLeft: function () {
              return 6;
            },
            paddingRight: function () {
              return 6;
            },
            paddingTop: function () {
              return 8;
            },
            paddingBottom: function () {
              return 8;
            },
          },
          margin: [0, 0, 0, 16],
        },

        {
          text: "ASSINATURAS",
          style: "sectionTitle",
          margin: [0, 0, 0, 10],
        },

        {
          table: {
            widths: ["*", 24, "*"],
            body: [
              [
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: responsavelCliente, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável da loja", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
                {
                  text: "",
                  border: [false, false, false, false],
                },
                {
                  stack: [
                    {
                      canvas: [
                        {
                          type: "line",
                          x1: 20,
                          y1: 28,
                          x2: 220,
                          y2: 28,
                          lineWidth: 1,
                          lineColor: "#334155",
                        },
                      ],
                      margin: [0, 0, 0, 10],
                    },
                    { text: responsavelRealstock, alignment: "center", bold: true, fontSize: 10 },
                    { text: "Responsável RealStock", alignment: "center", fontSize: 9, color: "#64748b" },
                  ],
                  margin: [8, 10, 8, 10],
                },
              ],
            ],
          },
          layout: "noBorders",
        },
      ],

      styles: {
        sectionTitle: {
          fontSize: 11,
          bold: true,
          color: "#0f172a",
        },
        infoBox: {
          fontSize: 9,
          bold: true,
          color: "#0f172a",
          fillColor: "#eff6ff",
          margin: [8, 8, 8, 8],
        },
        instructionBox: {
          fontSize: 9,
          color: "#0f172a",
          fillColor: "#f8fafc",
          margin: [8, 8, 8, 8],
        },
        tableHeader: {
          bold: true,
          fontSize: 9,
          color: "#ffffff",
        },
        tableCell: {
          fontSize: 8.7,
          color: "#0f172a",
        },
        markCell: {
          fontSize: 11,
          bold: true,
          color: "#0f172a",
        },
        summaryLabel: {
          fontSize: 9,
          bold: true,
          color: "#334155",
        },
        observationBox: {
          fontSize: 9,
          color: "#0f172a",
          fillColor: "#f8fafc",
          margin: [8, 8, 8, 8],
        },
        summaryValue: {
          fontSize: 10,
          bold: true,
          color: "#0f172a",
        },
      },

      defaultStyle: {
        font: "Helvetica",
      },
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="Avaliacao-Inventario-RealStock.pdf"'
    );
    pdfDoc.pipe(res);
    pdfDoc.end();
  } catch (erro) {
    console.error("Erro ao exportar avaliação do inventário PDF:", erro);
    res.status(500).send("Erro ao exportar avaliação do inventário PDF.");
  }
});

app.post("/transmitir-endereco", autenticar, (req, res) => {
  try {
    const { enderecoNumero, itens } = req.body || {};

    if (!enderecoNumero) {
      return res.status(400).json({ erro: "enderecoNumero obrigatório" });
    }

    if (!Array.isArray(itens) || itens.length === 0) {
      return res.status(400).json({ erro: "Nenhum item para transmitir." });
    }

    const usuario =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

    const endereco = buscarEnderecoPorNumero(Number(enderecoNumero));

    if (!endereco) {
      return res.status(404).json({
        erro: "Endereço não encontrado no endereçamento cadastrado."
      });
    }

    if (!Array.isArray(endereco.transmissoes)) {
      endereco.transmissoes = [];
    }

    const agoraIso = new Date().toISOString();

    const itensValidos = itens
      .map((item) => ({
        codigoBarras: String(item.codigoBarras || item.codigo || "").trim(),
        quantidade: Number(item.quantidade) || 0,
      }))
      .filter((i) => i.codigoBarras && i.quantidade > 0);

    if (!itensValidos.length) {
      return res.status(400).json({ erro: "Nenhum item válido." });
    }

    endereco.transmissoes.push({
      tipo: "transmissao",
    
      modoOperacao: normalizarModoOperacao(
        modoOperacao
      ),
    
      statusConsolidacao: "pendente",
      enderecoNumero: Number(enderecoNumero),
      usuario,
      data: agoraIso,
      itens: itensValidos,
    });

    endereco.finalizadoViaColetor = true;
endereco.finalizadoEm = agoraIso;

    endereco.contagensRecebidas =
      Number(endereco.contagensRecebidas || 0) + 1;

    endereco.ultimaContagemEm = agoraIso;
    endereco.atualizadoEm = agoraIso;

    salvarEnderecamentos();

    const painel = gerarPainelTransmissoesConsolidacao();

    return res.json({
      sucesso: true,
      mensagem: "Transmissão recebida.",
      endereco,
      painel,
    });
  } catch (erro) {
    console.error("Erro em /transmitir-endereco:", erro);
    return res.status(500).json({ erro: "Falha transmissão" });
  }
});
app.post("/finalizar-endereco", autenticar, (req, res) => {
  try {
    const { enderecoNumero } = req.body || {};

    if (enderecoNumero === undefined || enderecoNumero === null || enderecoNumero === "") {
      return res.status(400).json({ erro: "enderecoNumero é obrigatório." });
    }

    const endereco = registrarEventoEndereco(
      enderecoNumero,
      "finalizacao",
      req.session?.usuario || "sistema"
    );

    if (!endereco) {
      return res.status(404).json({ erro: "Nenhum endereçamento encontrado para esse número." });
    }

    return res.json({
      ok: true,
      mensagem: "Endereço finalizado com sucesso.",
      endereco,
    });
  } catch (erro) {
    console.error("Erro ao finalizar endereço:", erro);
    return res.status(500).json({ erro: "Falha ao finalizar endereço." });
  }
});
app.get(
  "/validar-endereco-mobile/:numero",
  autenticar,
  async (req, res) => {
    try {

      /*
        Aguarda os endereçamentos serem carregados.

        No Render, quando usar PostgreSQL,
        essa leitura é assíncrona.
      */
      await carregarEnderecamentos();

      const numero =
        Number(req.params.numero);

      if (
        !Number.isFinite(numero) ||
        numero <= 0
      ) {
        return res.status(400).json({
          valido: false,
          erro: "Endereço inválido.",
        });
      }

      const modoAtual =
        normalizarModoOperacao(
          modoOperacao
        );

      const endereco =
        buscarEnderecoPorNumero(
          numero,
          modoAtual
        );

      if (!endereco) {
        return res.status(404).json({
          valido: false,

          erro:
            `Endereço ${numero} não cadastrado no modo ${modoAtual}.`,
        });
      }

      return res.json({
        valido: true,

        endereco: {
          id:
            endereco.id,

          nome:
            endereco.nome,

          tipo:
            endereco.tipo,

          inicio:
            endereco.inicio,

          fim:
            endereco.fim,

          enderecoNumero:
            numero,

          modoOperacao:
            endereco.modoOperacao,
        },
      });

    } catch (erro) {
      console.error(
        "Erro ao validar endereço mobile:",
        erro
      );

      return res.status(500).json({
        valido: false,
        erro:
          "Falha ao validar endereço.",
      });
    }
  }
);
app.get("/enderecamentos", autenticar, (req, res) => {
  try {
    
    const modoAtual =
    normalizarModoOperacao(
      modoOperacao
    );
  
  const listaAtualizada = (
    Array.isArray(enderecamentos)
      ? enderecamentos
      : []
  )
    .filter(
      (endereco) =>
        normalizarModoOperacao(
          endereco?.modoOperacao
        ) === modoAtual
    )
    .map((endereco) => {

      const normalizado =
        normalizarEnderecoSalvo(endereco);

      const resumo =
        recalcularStatusFaixa(normalizado);

      return {
        ...normalizado,

        status: resumo.status,

        totalPosicoes:
          resumo.totalPosicoes,

        posicoesConcluidas:
          resumo.concluidos,

        posicoesPendentes:
          resumo.pendentes,

        posicoesEmContagem:
          resumo.emContagem,

        posicoesDuplicadas:
          resumo.duplicados,
      };
    });

    
    return res.json(listaAtualizada);
  } catch (erro) {
    console.error(
      "Erro ao recalcular endereçamentos:",
      erro
    );

    return res.status(500).json({
      erro:
        "Falha ao carregar os endereçamentos.",
    });
  }
});

app.get("/enderecamentos/posicoes-pendentes", autenticar, (req, res) => {
  try {
    const pendentes = [];

    enderecamentos.forEach((endereco) => {
      const inicio = Number(endereco.inicio) || 0;
      const fim = Number(endereco.fim) || 0;

      const finalizacoesAtivas = Array.isArray(
        endereco.finalizacoes
      )
        ? endereco.finalizacoes.filter(
            (finalizacao) =>
              !finalizacao?.excluida &&
              eventoPertenceAoModo(
                finalizacao
              )
          )
        : [];
      
      const transmissoesAtivas = Array.isArray(
        endereco.transmissoes
      )
        ? endereco.transmissoes.filter(
            (transmissao) =>
              !transmissao?.excluida &&
              eventoPertenceAoModo(
                transmissao
              )
          )
        : [];

      for (let numero = inicio; numero <= fim; numero++) {
        const qtdFinalizacoes = finalizacoesAtivas.filter(
          (f) => Number(f.enderecoNumero) === Number(numero)
        ).length;

        if (qtdFinalizacoes >= 1) {
          continue;
        }

        const qtdTransmissoes = transmissoesAtivas.filter(
          (t) =>
            t.tipo === "transmissao" &&
            Number(t.enderecoNumero) === Number(numero)
        ).length;

        pendentes.push({
          enderecoId: endereco.id,
          nome: endereco.nome || "--",
          tipo: endereco.tipo || "--",
          sequencia: endereco.sequencia || "",
          posicao: numero,
          faixa: `${inicio} até ${fim}`,
          status: qtdTransmissoes > 0
            ? "Aguardando finalização"
            : "Sem contagem/finalização",
          transmissoes: qtdTransmissoes,
          finalizacoes: qtdFinalizacoes,
        });
      }
    });

    return res.json({
      total: pendentes.length,
      pendentes,
    });
  } catch (erro) {
    console.error("Erro ao listar posições pendentes:", erro);
    return res.status(500).json({
      erro: "Falha ao listar posições pendentes.",
    });
  }
});
app.post("/enderecamentos", autenticar, (req, res) => {
  try {
    const {
      tipo,
      nome,
      inicio,
      fim,
      sequencia,
      status,
      observacoes,
    } = req.body || {};

    const tipoLimpo = normalizarTextoEndereco(tipo);
    const nomeLimpo = normalizarTextoEndereco(nome);
    const nInicio = Number(inicio);
    const nFim = Number(fim);
    const nSequencia = Number(sequencia) || 0;
    const statusLimpo = normalizarTextoEndereco(status || "pendente").toLowerCase();
    const observacoesLimpas = normalizarTextoEndereco(observacoes);

    if (!tipoLimpo || !nomeLimpo) {
      return res.status(400).json({ erro: "Tipo e nome da área são obrigatórios." });
    }

    if (!Number.isFinite(nInicio) || !Number.isFinite(nFim)) {
      return res.status(400).json({ erro: "Número inicial e final devem ser válidos." });
    }

    if (nFim < nInicio) {
      return res.status(400).json({ erro: "O número final não pode ser menor que o inicial." });
    }

    const conflito = existeFaixaDuplicadaOuSobreposta({
      tipo: tipoLimpo,
      inicio: nInicio,
      fim: nFim,
    });

    if (conflito) {
      return res.status(400).json({
        erro: `Já existe uma faixa igual ou sobreposta em ${conflito.tipo}: ${conflito.inicio} até ${conflito.fim}.`,
      });
    }

    const novoEndereco = {
      id: gerarNovoIdEnderecamento(),
      tipo: tipoLimpo,

      modoOperacao:
      normalizarModoOperacao(
        modoOperacao
      ),

      nome: nomeLimpo,
      inicio: nInicio,
      fim: nFim,
      sequencia: nSequencia,
      status: statusLimpo,
      observacoes: observacoesLimpas,
      totalPosicoes: calcularTotalPosicoesEndereco(nInicio, nFim),
      transmissoes: [],
      contagensRecebidas: 0,
      finalizadoViaColetor: false,
      ultimaContagemEm: null,
      finalizadoEm: null,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
    };

    enderecamentos.push(novoEndereco);
    salvarEnderecamentos();

    res.json({
      sucesso: true,
      endereco: novoEndereco,
    });
  } catch (erro) {
    console.error("Erro ao criar endereçamento:", erro);
    res.status(500).json({ erro: "Falha ao criar endereçamento." });
  }
});

app.put("/enderecamentos/:id", autenticar, (req, res) => {
  try {
    const id = Number(req.params.id);
    const {
      tipo,
      nome,
      inicio,
      fim,
      sequencia,
      status,
      observacoes,
    } = req.body || {};

    const index = enderecamentos.findIndex((e) => Number(e.id) === id);
    if (index === -1) {
      return res.status(404).json({ erro: "Endereço não encontrado." });
    }

    const tipoLimpo = normalizarTextoEndereco(tipo);
    const nomeLimpo = normalizarTextoEndereco(nome);
    const nInicio = Number(inicio);
    const nFim = Number(fim);
    const nSequencia = Number(sequencia) || 0;
    const statusLimpo = normalizarTextoEndereco(status || "pendente").toLowerCase();
    const observacoesLimpas = normalizarTextoEndereco(observacoes);

    if (!tipoLimpo || !nomeLimpo) {
      return res.status(400).json({ erro: "Tipo e nome da área são obrigatórios." });
    }

    if (!Number.isFinite(nInicio) || !Number.isFinite(nFim)) {
      return res.status(400).json({ erro: "Número inicial e final devem ser válidos." });
    }

    if (nFim < nInicio) {
      return res.status(400).json({ erro: "O número final não pode ser menor que o inicial." });
    }

    const conflito = existeFaixaDuplicadaOuSobreposta({
      idIgnorar: id,
      tipo: tipoLimpo,
      inicio: nInicio,
      fim: nFim,
    });

    if (conflito) {
      return res.status(400).json({
        erro: `Já existe uma faixa igual ou sobreposta em ${conflito.tipo}: ${conflito.inicio} até ${conflito.fim}.`,
      });
    }

    enderecamentos[index] = {
      ...enderecamentos[index],
      tipo: tipoLimpo,
      nome: nomeLimpo,
      inicio: nInicio,
      fim: nFim,
      sequencia: nSequencia,
      status: statusLimpo,
      observacoes: observacoesLimpas,
      totalPosicoes: calcularTotalPosicoesEndereco(nInicio, nFim),
      transmissoes: Array.isArray(enderecamentos[index].transmissoes)
        ? enderecamentos[index].transmissoes
        : [],
      contagensRecebidas: Number(enderecamentos[index].contagensRecebidas) || 0,
      finalizadoViaColetor: !!enderecamentos[index].finalizadoViaColetor,
      ultimaContagemEm: enderecamentos[index].ultimaContagemEm || null,
      finalizadoEm: enderecamentos[index].finalizadoEm || null,
      atualizadoEm: new Date().toISOString(),
    };

    salvarEnderecamentos();

    res.json({
      sucesso: true,
      endereco: enderecamentos[index],
    });
  } catch (erro) {
    console.error("Erro ao atualizar endereçamento:", erro);
    res.status(500).json({ erro: "Falha ao atualizar endereçamento." });
  }
});
app.get('/enderecos-itens-contados', autenticar, (req, res) => {
  try {
    const {
      endereco = '',
      termo = '',
      somenteAtivos = 'true'
    } = req.query;

    const termoBusca = String(termo || '').trim().toLowerCase();
    const enderecoBusca = String(endereco || '').trim();
    if (
      modoOperacao ===
      "sem-base"
    ) {
      const linhas = [];
    
      (
        Array.isArray(
          contagemSemBase
        )
          ? contagemSemBase
          : []
      ).forEach((item) => {
        const chaveItem =
          String(
            item.ean ||
            item.codigo ||
            ""
          ).trim();
    
        const enderecos =
          Array.isArray(
            item.enderecos
          )
            ? item.enderecos
            : [];
    
        /*
          Agrupa novamente por endereço
          para proteger dados antigos que
          possam ter sido duplicados.
        */
        const mapaEnderecos =
          new Map();
    
        enderecos.forEach(
          (registro) => {
            const numero =
              String(
                registro.enderecoNumero ||
                  ""
              ).trim();
    
            if (!numero) {
              return;
            }
    
            const atual =
              mapaEnderecos.get(
                numero
              ) || {
                quantidade: 0,
                usuario:
                  registro.usuario ||
                  item.ultimoUsuario ||
                  "--",
                data:
                  registro.data ||
                  item.data ||
                  null,
              };
    
            atual.quantidade +=
              Number(
                registro.quantidade
              ) || 0;
    
            /*
              Mantém o registro mais recente.
            */
            if (
              registro.data &&
              (
                !atual.data ||
                new Date(
                  registro.data
                ) >
                  new Date(
                    atual.data
                  )
              )
            ) {
              atual.data =
                registro.data;
    
              atual.usuario =
                registro.usuario ||
                atual.usuario;
            }
    
            mapaEnderecos.set(
              numero,
              atual
            );
          }
        );
    
        mapaEnderecos.forEach(
          (
            registro,
            numero
          ) => {
            if (
              enderecoBusca &&
              numero !==
                enderecoBusca
            ) {
              return;
            }
    
            const codigoInterno =
              String(
                item.codigo || ""
              );
    
            const ean =
              String(
                item.ean || ""
              );
    
            const textoBusca =
              `${ean} ${codigoInterno} ${numero}`
                .toLowerCase();
    
            if (
              termoBusca &&
              !textoBusca.includes(
                termoBusca
              )
            ) {
              return;
            }
    
            const enderecoObj =
              buscarEnderecoPorNumero(
                Number(numero)
              );
    
            linhas.push({
              id:
                `SEMBASE|${encodeURIComponent(
                  chaveItem
                )}|${encodeURIComponent(
                  numero
                )}`,
    
              codigoBarras:
                ean,
    
              codigoInterno,
    
              descricao:
                "Item sem base",
    
              enderecoNumero:
                numero,
    
              enderecoNome:
                enderecoObj?.nome ||
                "ENDEREÇO",
    
              quantidadeContadaNoEndereco:
                registro.quantidade,
    
              origem:
                "sem-base",
    
              ativo: true,
    
              usuario:
                registro.usuario,
    
              data:
                registro.data,
            });
          }
        );
      });
    
      linhas.sort(
        (a, b) => {
          const enderecoA =
            Number(
              a.enderecoNumero
            ) || 0;
    
          const enderecoB =
            Number(
              b.enderecoNumero
            ) || 0;
    
          if (
            enderecoA !==
            enderecoB
          ) {
            return (
              enderecoA -
              enderecoB
            );
          }
    
          return String(
            a.codigoBarras ||
            a.codigoInterno ||
            ""
          ).localeCompare(
            String(
              b.codigoBarras ||
              b.codigoInterno ||
              ""
            ),
            "pt-BR"
          );
        }
      );
    
      return res.json(linhas);
    }

    const linhas = (Array.isArray(contagens) ? contagens : [])
      .filter((item) => {
        if (!item) return false;

        if (somenteAtivos === 'true' && item.ativo === false) {
          return false;
        }

        if (enderecoBusca && String(item.enderecoNumero || '') !== enderecoBusca) {
          return false;
        }

        const produto = inventario.find(
          (p) => String(p.codigoBarras || '') === String(item.codigoBarras || '')
        );

        const codigoInterno = String(produto?.codigo || produto?.codigoInterno || '').toLowerCase();
        const ean = String(item.codigoBarras || '').toLowerCase();
        const descricao = String(produto?.descricao || '').toLowerCase();
        const enderecoTexto = String(item.enderecoNumero || '').toLowerCase();

        if (termoBusca) {
          const combinado = `${codigoInterno} ${ean} ${descricao} ${enderecoTexto}`;
          if (!combinado.includes(termoBusca)) {
            return false;
          }
        }

        return true;
      })
      .map((item) => {
        const produto = inventario.find(
          (p) => String(p.codigoBarras || '') === String(item.codigoBarras || '')
        );

        const enderecoObj = item.enderecoId
          ? enderecamentos.find((e) => Number(e.id) === Number(item.enderecoId))
          : buscarEnderecoPorNumero(item.enderecoNumero);

        return {
          id: item.id,
          codigoBarras: item.codigoBarras || '',
          codigoInterno: produto?.codigo || produto?.codigoInterno || '',
          descricao: produto?.descricao || 'Item não encontrado',
          enderecoNumero: Number(item.enderecoNumero) || 0,
          enderecoNome: enderecoObj?.nome || 'ENDEREÇO',
          quantidadeContadaNoEndereco: Number(item.quantidade) || 0,
          origem: item.origem || 'coleta',
          ativo: item.ativo !== false,
          usuario: item.usuario || '--',
          data: item.data || null
        };
      })
      .sort((a, b) => {
        if (a.enderecoNumero !== b.enderecoNumero) {
          return a.enderecoNumero - b.enderecoNumero;
        }
        return String(a.descricao || '').localeCompare(String(b.descricao || ''), 'pt-BR');
      });

    res.json(linhas);
  } catch (erro) {
    console.error('Erro ao listar itens contados por endereço:', erro);
    res.status(500).json({ erro: 'Falha ao listar itens contados por endereço.' });
  }
});
app.post('/enderecos-itens-contados/editar', autenticar, (req, res) => {
  try {
    const { id, novaQuantidade } = req.body;

    if (!id) {
      return res.status(400).json({ erro: 'ID da contagem não informado.' });
    }

    const quantidadeNumero = Number(novaQuantidade);
    if (isNaN(quantidadeNumero)) {
      return res.status(400).json({ erro: 'Quantidade inválida.' });
    }

    const indice = contagens.findIndex((item) => String(item.id) === String(id));

    if (indice === -1) {
      return res.status(404).json({ erro: 'Contagem não encontrada.' });
    }

    const anterior = Number(contagens[indice].quantidade) || 0;

    contagens[indice] = {
      ...contagens[indice],
      quantidade: quantidadeNumero,
      dataEdicao: new Date().toISOString(),
      editadoManualPorEndereco: true,
      usuarioEdicao:
        req.session?.usuario?.usuario || req.session?.usuario?.nome || 'sistema'
    };

    salvarContagens();

    historicoAuditoriaItens.push({
      tipo: 'edicao-endereco',
      contagemId: contagens[indice].id,
      codigoBarras: contagens[indice].codigoBarras,
      enderecoNumero: contagens[indice].enderecoNumero || null,
      quantidadeAnterior: anterior,
      novaQuantidade: quantidadeNumero,
      usuario:
        req.session?.usuario?.usuario || req.session?.usuario?.nome || 'sistema',
      data: new Date().toISOString()
    });

    return res.json({
      ok: true,
      mensagem: 'Quantidade do item no endereço atualizada com sucesso.'
    });
  } catch (erro) {
    console.error('Erro ao editar item contado por endereço:', erro);
    return res.status(500).json({ erro: 'Falha ao editar item contado por endereço.' });
  }
});

function formatarDataHoraManaus(
  valor = new Date()
) {
  const data =
    valor instanceof Date
      ? valor
      : new Date(valor);

  if (
    Number.isNaN(
      data.getTime()
    )
  ) {
    return "--";
  }

  return new Intl.DateTimeFormat(
    "pt-BR",
    {
      timeZone:
        "America/Manaus",

      day: "2-digit",
      month: "2-digit",
      year: "numeric",

      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",

      hour12: false,
    }
  ).format(data);
}


function montarLinhasPdfItensPorEndereco({
  endereco = "",
  termo = "",
  somenteAtivos = "true",
} = {}) {

  const termoBusca =
    String(
      termo || ""
    )
      .trim()
      .toLowerCase();

  const enderecoBusca =
    String(
      endereco || ""
    ).trim();


  /*
    ========================================================
    SEM BASE
    ========================================================
  */

  if (
    normalizarModoOperacao(
      modoOperacao
    ) === "sem-base"
  ) {

    const linhas = [];


    (
      Array.isArray(
        contagemSemBase
      )
        ? contagemSemBase
        : []
    ).forEach((item) => {

      const chaveItem =
        String(
          item.ean ||
          item.codigo ||
          ""
        ).trim();


      const mapaEnderecos =
        new Map();


      (
        Array.isArray(
          item.enderecos
        )
          ? item.enderecos
          : []
      ).forEach(
        (registro) => {

          const numero =
            String(
              registro.enderecoNumero ||
              ""
            ).trim();


          if (!numero) {
            return;
          }


          const atual =
            mapaEnderecos.get(
              numero
            ) || {
              quantidade: 0,

              usuario:
                registro.usuario ||
                item.ultimoUsuario ||
                "--",

              data:
                registro.data ||
                item.data ||
                item.ultimaLeituraEm ||
                null,
            };


          atual.quantidade +=
            Number(
              registro.quantidade
            ) || 0;


          if (
            registro.data &&
            (
              !atual.data ||
              new Date(
                registro.data
              ) >
                new Date(
                  atual.data
                )
            )
          ) {
            atual.data =
              registro.data;

            atual.usuario =
              registro.usuario ||
              atual.usuario;
          }


          mapaEnderecos.set(
            numero,
            atual
          );
        }
      );


      mapaEnderecos.forEach(
        (
          registro,
          numero
        ) => {

          if (
            enderecoBusca &&
            numero !==
              enderecoBusca
          ) {
            return;
          }


          const ean =
            String(
              item.ean || ""
            );

          const codigoInterno =
            String(
              item.codigo || ""
            );


          const descricao =
            "Item sem base";


          const combinado =
            (
              `${ean} ` +
              `${codigoInterno} ` +
              `${descricao} ` +
              `${numero}`
            ).toLowerCase();


          if (
            termoBusca &&
            !combinado.includes(
              termoBusca
            )
          ) {
            return;
          }


          const enderecoObj =
            buscarEnderecoPorNumero(
              Number(numero),
              "sem-base"
            );


          linhas.push({
            codigoBarras:
              ean,

            codigoInterno,

            descricao,

            endereco:
              `${
                enderecoObj?.nome ||
                "ENDEREÇO"
              } • ${numero}`,

            enderecoNumero:
              Number(numero) || 0,

            quantidade:
              Number(
                registro.quantidade
              ) || 0,

            origem:
              "sem-base",

            usuario:
              registro.usuario ||
              "--",

            data:
              registro.data
                ? formatarDataHoraManaus(
                    registro.data
                  )
                : "--",

            chaveItem,
          });
        }
      );
    });


    return linhas.sort(
      (a, b) => {

        if (
          a.enderecoNumero !==
          b.enderecoNumero
        ) {
          return (
            a.enderecoNumero -
            b.enderecoNumero
          );
        }

        return String(
          a.codigoBarras ||
          a.codigoInterno ||
          ""
        ).localeCompare(
          String(
            b.codigoBarras ||
            b.codigoInterno ||
            ""
          ),
          "pt-BR"
        );
      }
    );
  }


  /*
    ========================================================
    COM BASE / WMS
    ========================================================
  */

  return (
    Array.isArray(contagens)
      ? contagens
      : []
  )
    .filter((item) => {

      if (!item) {
        return false;
      }

      if (
        somenteAtivos === "true" &&
        item.ativo === false
      ) {
        return false;
      }

      if (
        enderecoBusca &&
        String(
          item.enderecoNumero || ""
        ) !== enderecoBusca
      ) {
        return false;
      }


      const produto =
        inventario.find(
          (p) =>
            String(
              p.codigoBarras || ""
            ) ===
            String(
              item.codigoBarras || ""
            )
        );


      const combinado =
        (
          `${
            produto?.codigo ||
            produto?.codigoInterno ||
            ""
          } ` +
          `${item.codigoBarras || ""} ` +
          `${produto?.descricao || ""} ` +
          `${item.enderecoNumero || ""}`
        ).toLowerCase();


      if (
        termoBusca &&
        !combinado.includes(
          termoBusca
        )
      ) {
        return false;
      }


      return true;
    })
    .map((item) => {

      const produto =
        inventario.find(
          (p) =>
            String(
              p.codigoBarras || ""
            ) ===
            String(
              item.codigoBarras || ""
            )
        );


      const enderecoObj =
        item.enderecoId
          ? enderecamentos.find(
              (e) =>
                Number(e.id) ===
                Number(
                  item.enderecoId
                )
            )
          : buscarEnderecoPorNumero(
              item.enderecoNumero
            );


      return {
        codigoBarras:
          item.codigoBarras ||
          "",

        codigoInterno:
          produto?.codigo ||
          produto?.codigoInterno ||
          "",

        descricao:
          produto?.descricao ||
          "Item não encontrado",

        endereco:
          `${
            enderecoObj?.nome ||
            "ENDEREÇO"
          } • ${
            Number(
              item.enderecoNumero
            ) || 0
          }`,

        enderecoNumero:
          Number(
            item.enderecoNumero
          ) || 0,

        quantidade:
          Number(
            item.quantidade
          ) || 0,

        origem:
          item.origem ||
          "coleta",

        usuario:
          item.usuario ||
          "--",

        data:
          item.data
            ? formatarDataHoraManaus(
                item.data
              )
            : "--",
      };
    })
    .sort(
      (a, b) => {

        if (
          a.enderecoNumero !==
          b.enderecoNumero
        ) {
          return (
            a.enderecoNumero -
            b.enderecoNumero
          );
        }

        return String(
          a.descricao || ""
        ).localeCompare(
          String(
            b.descricao || ""
          ),
          "pt-BR"
        );
      }
    );
}

app.get('/enderecos-itens-contados/pdf', autenticar, (req, res) => {
  try {
    const {
      endereco = '',
      termo = '',
      somenteAtivos = 'true'
    } = req.query;

    const linhas =
    montarLinhasPdfItensPorEndereco({
      endereco,
      termo,
      somenteAtivos,
    });

    const printer = new PdfPrinter(fonts);

    const body = [
      [
        {
          text: "EAN",
          style: "tableHeader",
        },
        {
          text: "Cód. interno",
          style: "tableHeader",
        },
        {
          text: "Descrição",
          style: "tableHeader",
        },
        {
          text: "Endereço",
          style: "tableHeader",
        },
        {
          text: "Quantidade",
          style: "tableHeader",
          alignment: "right",
        },
        {
          text: "Usuário",
          style: "tableHeader",
        },
      ],
    
      ...linhas.map((item, indice) => {
        const fundo =
          indice % 2 === 0
            ? "#f8fafc"
            : "#ffffff";
    
        return [
          {
            text: item.codigoBarras,
            style: "tableCell",
            fillColor: fundo,
          },
          {
            text: item.codigoInterno,
            style: "tableCell",
            fillColor: fundo,
          },
          {
            text: item.descricao,
            style: "tableCell",
            fillColor: fundo,
          },
          {
            text: item.endereco,
            style: "tableCell",
            fillColor: fundo,
          },
          {
            text: Number(
              item.quantidade || 0
            ).toLocaleString("pt-BR"),
            style: "tableCell",
            alignment: "right",
            bold: true,
            color: "#0f766e",
            fillColor: fundo,
          },
          {
            text: item.usuario,
            style: "tableCell",
            fillColor: fundo,
          },
        ];
      }),
    ];
    
    const docDefinition = {
      pageOrientation: "landscape",
      pageMargins: [28, 72, 28, 48],
    
      header: {
        margin: [28, 18, 28, 0],
    
        columns: [
          {
            width: "*",
            stack: [
              {
                text: "REALSTOCK",
                color: "#14b8a6",
                fontSize: 17,
                bold: true,
              },
              {
                text:
                  "Inventário Inteligente · Relatório operacional",
                color: "#64748b",
                fontSize: 8,
                margin: [0, 2, 0, 0],
              },
            ],
          },
    
          {
            width: "auto",
            text:
  formatarDataHoraManaus(
    new Date()
  ),
            color: "#64748b",
            fontSize: 8,
            alignment: "right",
            margin: [0, 7, 0, 0],
          },
        ],
      },
    
      footer: function (
        paginaAtual,
        totalPaginas
      ) {
        return {
          margin: [28, 10, 28, 0],
    
          columns: [
            {
              text: "RealStock",
              color: "#14b8a6",
              bold: true,
              fontSize: 9,
            },
    
            {
              text:
                `Página ${paginaAtual} de ${totalPaginas}`,
              alignment: "right",
              color: "#64748b",
              fontSize: 8,
            },
          ],
        };
      },
    
      content: [
        {
          canvas: [
            {
              type: "rect",
              x: 0,
              y: 0,
              w: 786,
              h: 4,
              color: "#14b8a6",
            },
          ],
          margin: [0, 0, 0, 14],
        },
    
        {
          text: "Itens contados por endereço",
          style: "title",
        },
    
        {
          text:
            `${linhas.length.toLocaleString(
              "pt-BR"
            )} registro(s) listado(s)`,
          style: "subTitle",
          margin: [0, 3, 0, 14],
        },
    
        {
          table: {
            headerRows: 1,
    
            widths: [
              105,
              82,
              "*",
              115,
              82,
              90,
            ],
    
            body,
          },
    
          layout: {
            fillColor: function (rowIndex) {
              return rowIndex === 0
                ? "#123554"
                : null;
            },
    
            hLineColor: function () {
              return "#dbe5ef";
            },
    
            vLineColor: function () {
              return "#dbe5ef";
            },
    
            hLineWidth: function (
              linhaIndex
            ) {
              return linhaIndex === 0 ? 0 : 0.6;
            },
    
            vLineWidth: function () {
              return 0.6;
            },
    
            paddingLeft: function () {
              return 8;
            },
    
            paddingRight: function () {
              return 8;
            },
    
            paddingTop: function () {
              return 7;
            },
    
            paddingBottom: function () {
              return 7;
            },
          },
        },
      ],
    
      styles: {
        title: {
          fontSize: 18,
          bold: true,
          color: "#10243a",
        },
    
        subTitle: {
          fontSize: 9,
          color: "#64748b",
        },
    
        tableHeader: {
          bold: true,
          fontSize: 8,
          color: "#ffffff",
        },
    
        tableCell: {
          fontSize: 8,
          color: "#263649",
        },
      },
    
      defaultStyle: {
        font: "Helvetica",
      },
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);

    res.setHeader(
      "Content-Type",
      "application/pdf"
    );
    
    res.setHeader(
      "Content-Disposition",
      'inline; filename="itens-contados-por-endereco.pdf"'
    );
    
    res.setHeader(
      "Cache-Control",
      "no-store, no-cache, must-revalidate, private"
    );
    
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");

    pdfDoc.pipe(res);
    pdfDoc.end();
  } catch (erro) {
    console.error('Erro ao gerar PDF de itens por endereço:', erro);
    res.status(500).send('Erro ao gerar PDF.');
  }
});
app.delete("/enderecamentos/:id", autenticar, (req, res) => {
  try {
    const id = Number(req.params.id);

    const index = enderecamentos.findIndex((e) => Number(e.id) === id);
    if (index === -1) {
      return res.status(404).json({ erro: "Endereço não encontrado." });
    }

    enderecamentos.splice(index, 1);
    salvarEnderecamentos();

    res.json({ sucesso: true });
  } catch (erro) {
    console.error("Erro ao excluir endereçamento:", erro);
    res.status(500).json({ erro: "Falha ao excluir endereçamento." });
  }
});

app.get("/auditoria-enderecos-duplicados", autenticar, (req, res) => {
  try {
    const auditoria = gerarAuditoriaDuplicidadeEnderecos();
    res.json(auditoria);
  } catch (erro) {
    console.error("Erro ao gerar auditoria de endereços duplicados:", erro);
    res.status(500).json({ erro: "Falha ao gerar auditoria de endereços duplicados." });
  }
});

app.get("/usuarios", autenticar, (req, res) => {
  res.json(
    gerarResumoUsuariosAtivos().concat(
      usuarios
        .filter((u) => String(u.status || "").toLowerCase() !== "ativo")
        .map((u) => {
          const registrosUsuario = contagens.filter((c) => c.usuario === u.usuario);
          const ultimaAtividade =
            registrosUsuario.length > 0
              ? registrosUsuario[registrosUsuario.length - 1].data
              : null;

          const totalContado = registrosUsuario.reduce(
            (acc, item) => acc + (Number(item.quantidade) || 0),
            0
          );

          return {
            id: u.id,
            nome: u.nome || u.usuario,
            usuario: u.usuario,
            senha: u.senha || "",
            matricula: u.matricula,
            funcao: u.funcao || "Operador",
            telefone: u.telefone || "",
            status: u.status || "inativo",
            meta: Number(u.meta) || 0,
            totalContado,
            ultimaAtividade,
          };
        })
    )
  );
});

app.post("/usuarios", autenticar, async (req, res) => {
  try {
    const {
      nome,
      usuario,
      senha,
      funcao,
      telefone,
      status,
      meta,
    } = req.body || {};

    if (!nome || !usuario || !senha) {
      return res.status(400).json({ erro: "Nome, usuário e senha são obrigatórios." });
    }

    const usuarioExistente = usuarios.find(
      (u) => String(u.usuario).toLowerCase() === String(usuario).toLowerCase()
    );

    if (usuarioExistente) {
      return res.status(400).json({ erro: "Já existe um usuário com esse login." });
    }

    const novoUsuario = {
      id: gerarNovoIdUsuario(),
      nome: String(nome).trim(),
      usuario: String(usuario).trim(),
      senha: String(senha).trim(),
      matricula: gerarNovaMatricula(),
      funcao: String(funcao || "Operador").trim(),
      telefone: String(telefone || "").trim(),
      status: String(status || "ativo").trim(),
      meta: Number(meta) || 0,
      criadoEm: new Date().toISOString(),
    };

    usuarios.push(novoUsuario);
    await salvarUsuarios();
    return res.json({
      sucesso: true,
      usuario: novoUsuario,
    });
  } catch (erro) {
    console.error("Erro ao criar usuário:", erro);
    return res.status(500).json({ erro: "Falha ao criar usuário." });
  }
});

app.put("/usuarios/:id/meta", autenticar, async (req, res) => {
  try {
    const { id } = req.params;
    const { meta } = req.body || {};

    const usuario = usuarios.find((u) => String(u.id) === String(id));

    if (!usuario) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    usuario.meta = Number(meta) || 0;
    await salvarUsuarios();

    res.json({ sucesso: true, usuario });
  } catch (erro) {
    console.error("Erro ao atualizar meta:", erro);
    res.status(500).json({ erro: "Falha ao atualizar meta." });
  }
});

app.put("/usuarios/:id", autenticar, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const { nome, usuario, senha, funcao, telefone, meta, status } = req.body || {};

    const usuarioIndex = usuarios.findIndex((u) => Number(u.id) === id);
    if (usuarioIndex === -1) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    const loginDuplicado = usuarios.some(
      (u) =>
        Number(u.id) !== id &&
        String(u.usuario || "").toLowerCase() === String(usuario || "").toLowerCase()
    );

    if (loginDuplicado) {
      return res.status(400).json({ erro: "Já existe outro usuário com esse login." });
    }

    usuarios[usuarioIndex] = {
      ...usuarios[usuarioIndex],
      nome: nome ?? usuarios[usuarioIndex].nome,
      usuario: usuario ?? usuarios[usuarioIndex].usuario,
      senha: senha ?? usuarios[usuarioIndex].senha,
      funcao: funcao ?? usuarios[usuarioIndex].funcao,
      telefone: telefone ?? usuarios[usuarioIndex].telefone,
      meta: Number(meta) || 0,
      status: status ?? usuarios[usuarioIndex].status,
    };

    await salvarUsuarios();

    return res.json({ ok: true, usuario: usuarios[usuarioIndex] });
  } catch (erro) {
    console.error("Erro ao editar usuário:", erro);
    return res.status(500).json({ erro: "Falha ao editar usuário." });
  }
});

app.patch("/usuarios/:id/status", autenticar, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const { status } = req.body;

    const usuario = usuarios.find((u) => Number(u.id) === id);
    if (!usuario) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    if (!["ativo", "inativo"].includes(status)) {
      return res.status(400).json({ erro: "Status inválido." });
    }

    usuario.status = status;
    await salvarUsuarios();

    return res.json({ ok: true, usuario });
  } catch (erro) {
    console.error("Erro ao alterar status do usuário:", erro);
    return res.status(500).json({ erro: "Falha ao alterar status." });
  }
});

app.delete("/usuarios/:id", autenticar, async (req, res) => {
  try {
    const id = Number(req.params.id);

    const index = usuarios.findIndex((u) => Number(u.id) === id);
    if (index === -1) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    const usuario = usuarios[index];

    if ((usuario.usuario || "").toLowerCase() === "admin") {
      return res.status(400).json({ erro: "O administrador principal não pode ser excluído." });
    }

    usuarios.splice(index, 1);
    await salvarUsuarios();

    return res.json({ ok: true });
  } catch (erro) {
    console.error("Erro ao excluir usuário:", erro);
    return res.status(500).json({ erro: "Falha ao excluir usuário." });
  }
});
app.get("/tipo-ultima-importacao", autenticar, (req, res) => {
  res.json({
    tipoUltimaImportacao,
    ultimaImportacao,
  });
});
app.get("/usuarios-atividade", autenticar, (req, res) => {
  const atividade = usuarios.map((u) => {
    const registros = contagens
      .filter((c) => c.usuario === u.usuario)
      .sort((a, b) => new Date(b.data) - new Date(a.data))
      .slice(0, 5);

    return {
      id: u.id,
      nome: u.nome || u.usuario,
      usuario: u.usuario,
      matricula: u.matricula,
      funcao: u.funcao || "Operador",
      ultimaAtividade: registros[0]?.data || null,
      registros,
    };
  });

  res.json(atividade);
});

const dadosInventarioPadrao = {
  cliente: "",
  loja: "",
  codigoLoja: "",
  dataInventario: "",
  tipoInventario: "geral",
  escopoInventario: "loja",
  modoContagem: "cego",
  cidade: "",
  uf: "",
  responsavelCliente: "",
  liderOperacao: "",
  horaInicio: "",
  horaMeta: "",
  observacoes: "",
};

app.get("/dados-inventario", autenticar, async (req, res) => {
  try {
    const dados = usarPostgres
      ? await carregarConfiguracaoPostgres("dados_inventario", dadosInventarioPadrao)
      : dadosInventarioPadrao;

    return res.json({
      ...dadosInventarioPadrao,
      ...dados,
    });
  } catch (erro) {
    console.error("Erro ao carregar dados do inventário:", erro.message);
    return res.status(500).json({
      erro: "Erro ao carregar dados do inventário.",
    });
  }
});

app.post("/dados-inventario", autenticar, async (req, res) => {
  try {
    const dados = {
      cliente: String(req.body.cliente || "").trim(),
      loja: String(req.body.loja || "").trim(),
      codigoLoja: String(req.body.codigoLoja || "").trim(),
      dataInventario: String(req.body.dataInventario || ""),
      tipoInventario: String(req.body.tipoInventario || "geral"),
      escopoInventario: String(req.body.escopoInventario || "loja"),
      modoContagem: String(req.body.modoContagem || "cego"),
      cidade: String(req.body.cidade || "").trim(),
      uf: String(req.body.uf || "").trim().toUpperCase(),
      responsavelCliente: String(req.body.responsavelCliente || "").trim(),
      liderOperacao: String(req.body.liderOperacao || "").trim(),
      horaInicio: String(req.body.horaInicio || ""),
      horaMeta: String(req.body.horaMeta || ""),
      observacoes: String(req.body.observacoes || "").trim(),
    };

    if (usarPostgres) {
      await salvarConfiguracaoPostgres("dados_inventario", dados);
    }

    return res.json({
      sucesso: true,
      mensagem: "Dados do inventário salvos.",
      dados,
    });
  } catch (erro) {
    console.error("Erro ao salvar dados do inventário:", erro.message);
    return res.status(500).json({
      erro: "Erro ao salvar dados do inventário.",
    });
  }
});

async function iniciarServidor() {
  try {
    if (usarPostgres) {
      await testarConexao();
      await criarTabelas();
    } else {
      console.log("PostgreSQL ignorado no StackBlitz/local.");
    }
    
    carregarLayoutTxt();
carregarLayoutsSalvos();
carregarLayoutsUniversais();
carregarLayoutsExportacao();

    await carregarContagemSemBase();
    await carregarFinalizacoesSemBase();
    await carregarModoOperacao();
    
    await carregarProdutosDoBanco();
    await carregarUsuarios();
    await carregarEnderecamentos();
    await carregarContagens();

    server.listen(port, () => {
      console.log(`Servidor rodando em http://localhost:${port}`);
      console.log(`✅ Sistema iniciado com ${inventario.length} produtos, ${enderecamentos.length} endereçamento(s) e ${contagens.length} contagem(ns).`);
    });
  } catch (erro) {
    console.error("❌ Erro ao iniciar servidor:", erro);
    process.exit(1);
  }
}

iniciarServidor();