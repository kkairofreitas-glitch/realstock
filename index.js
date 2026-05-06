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

  return caminhoPublic;
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
let contagemSemBase = [];
let modoOperacao = "com-base";
let finalizacoesSemBase = [];

const dataDir = path.join(__dirname, "data");
const contagensPath = path.join(dataDir, "contagens.json");
const layoutTxtPath = path.join(dataDir, "layout-txt.json");
const layoutsTxtPath = path.join(dataDir, "layouts-txt.json");
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
  if (req.session && req.session.logado) {
    return next();
  }

  const aceitaJson =
    req.xhr ||
    req.headers.accept?.includes("application/json") ||
    req.headers["content-type"]?.includes("application/json");

  if (aceitaJson) {
    return res.status(401).json({ erro: "Sessão expirada" });
  }

  return res.redirect(`/login?redirect=${encodeURIComponent(req.originalUrl || "/")}`);
}
function permitirSomenteLiderOuAdmin(req, res, next) {
  if (!req.session?.logado || !req.session?.usuario) {
    return res.redirect('/login');
  }

  const funcao = String(req.session.usuario.funcao || '').toLowerCase();

  if (funcao === 'líder' || funcao === 'lider' || funcao === 'administrador') {
    return next();
  }

  return res.redirect('/coleta-mobile');
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

function resetarSistemaAposEncerramento() {
  inventario = [];
  historicoAlteracoes = [];
  historicoAuditoriaItens = [];
  contagens = [];
  enderecamentos = [];
  contagemSemBase = [];
  finalizacoesSemBase = [];
  itemAuditoriaAtual = null;
  modoOperacao = "com-base";

  auditoriaImportacao = {
    totalImportadoBruto: 0,
    totalUnicosBruto: 0,
    duplicatasRemovidas: 0,
    itensZeradosIgnorados: 0,
  };

  salvarContagens();
  salvarEnderecamentos();
  salvarContagemSemBase();
  salvarFinalizacoesSemBase();
  salvarModoOperacao();
}
function garantirPastaData() {
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }
}

function carregarContagens() {
  try {
    garantirPastaData();

    if (!fs.existsSync(contagensPath)) {
      fs.writeFileSync(contagensPath, JSON.stringify([], null, 2), "utf8");
      contagens = [];
      return;
    }

    const conteudo = fs.readFileSync(contagensPath, "utf8");
    contagens = JSON.parse(conteudo || "[]");
  } catch (erro) {
    console.error("Erro ao carregar contagens:", erro);
    contagens = [];
  }
}

function salvarContagens() {
  try {
    garantirPastaData();
    fs.writeFileSync(contagensPath, JSON.stringify(Array.isArray(contagens) ? contagens : [], null, 2), "utf8");
  } catch (erro) {
    console.error("Erro ao salvar contagens:", erro);
  }
}
function limparContagensPersistidas() {
  contagens = [];
  salvarContagens();
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
function carregarEnderecamentos() {
  try {
    garantirPastaData();

    if (!fs.existsSync(enderecamentosPath)) {
      fs.writeFileSync(enderecamentosPath, JSON.stringify([], null, 2), "utf8");
      enderecamentos = [];
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
  } catch (erro) {
    console.error("Erro ao carregar endereçamentos:", erro);
    enderecamentos = [];
  }
}

function salvarEnderecamentos() {
  try {
    garantirPastaData();
    fs.writeFileSync(enderecamentosPath, JSON.stringify(enderecamentos, null, 2), "utf8");
  } catch (erro) {
    console.error("Erro ao salvar endereçamentos:", erro);
  }
}

function gerarNovoIdEnderecamento() {
  const ids = enderecamentos.map((e) => Number(e.id) || 0);
  return ids.length ? Math.max(...ids) + 1 : 1;
}

function existeFaixaDuplicadaOuSobreposta({ idIgnorar = null, tipo, inicio, fim }) {
  const nInicio = Number(inicio) || 0;
  const nFim = Number(fim) || 0;
  const tipoNormalizado = normalizarTextoEndereco(tipo).toLowerCase();

  return enderecamentos.find((item) => {
    if (idIgnorar !== null && Number(item.id) === Number(idIgnorar)) {
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

function buscarEnderecoPorNumero(enderecoNumero) {
  const numero = Number(enderecoNumero);

  return enderecamentos.find((item) => {
    const inicio = Number(item.inicio) || 0;
    const fim = Number(item.fim) || 0;
    return numero >= inicio && numero <= fim;
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

  const eventos = Array.isArray(endereco?.transmissoes) ? endereco.transmissoes : [];
  const finalizacoesAtivas = Array.isArray(endereco?.finalizacoes)
    ? endereco.finalizacoes.filter((f) => !f.excluida)
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
      pendentes += 1;
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
function registrarEventoEndereco(enderecoNumero, tipoEvento, usuario = "sistema") {
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
    enderecoNumero: Number(enderecoNumero),
    usuario,
    data: agoraIso,
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
  
    const itensDaFinalizacao = contagens
      .filter((c) => {
        return (
          c.ativo !== false &&
          Number(c.enderecoNumero) === Number(enderecoNumero) &&
          !c.finalizacaoId
        );
      })
      .map((c) => ({
        contagemId: c.id,
        codigoBarras: c.codigoBarras,
        quantidade: Number(c.quantidade) || 0,
        usuario: c.usuario,
        data: c.data,
      }));
  
    contagens = contagens.map((c) => {
      if (
        c.ativo !== false &&
        Number(c.enderecoNumero) === Number(enderecoNumero) &&
        !c.finalizacaoId
      ) {
        return {
          ...c,
          finalizacaoId,
        };
      }
  
      return c;
    });
  
    endereco.finalizacoes.push({
      id: finalizacaoId,
      enderecoId: Number(endereco.id),
      enderecoNumero: Number(enderecoNumero),
      usuario,
      data: agoraIso,
      excluida: false,
      itens: itensDaFinalizacao,
    });
  
    endereco.finalizadoViaColetor = true;
    endereco.finalizadoEm = agoraIso;
  
    salvarContagens();
  }
  
  const resumoFaixa = recalcularStatusFaixa(endereco);

endereco.status = resumoFaixa.status;
endereco.totalPosicoes = resumoFaixa.totalPosicoes;
endereco.posicoesConcluidas = resumoFaixa.concluidos;
endereco.posicoesPendentes = resumoFaixa.pendentes;
endereco.posicoesEmContagem = resumoFaixa.emContagem;
endereco.posicoesDuplicadas = resumoFaixa.duplicados;
  endereco.atualizadoEm = agoraIso;
  salvarEnderecamentos();

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
function formatarDataHoraTransmissao(valor) {
  if (!valor) return "--";

  const d = new Date(valor);
  if (isNaN(d.getTime())) return "--";

  return d.toLocaleString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function contarItensDoEndereco(endereco) {
  if (!endereco) return 0;

  const inicio = Number(endereco.inicio) || 0;
  const fim = Number(endereco.fim) || 0;

  if (fim < inicio) return 0;

  const itensNoIntervalo = inventario.filter((item) => {
    const codigo = Number(item.codigo || item.codigoInterno || 0);
    return codigo >= inicio && codigo <= fim;
  });

  const contados = itensNoIntervalo.filter((item) => Number(item.qtdeContada || 0) > 0);
  return contados.length;
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
      itens: contarItensDoEndereco(endereco),
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
function carregarUsuarios() {
  try {
    garantirPastaData();

    if (!fs.existsSync(usuariosPath)) {
      const usuariosIniciais = [
        {
          id: 1,
          nome: "Administrador",
          usuario: "admin",
          senha: "1234",
          matricula: "RS-0001",
          funcao: "Administrador",
          telefone: "",
          status: "ativo",
          meta: 0,
          criadoEm: new Date().toISOString(),
        },
        {
          id: 2,
          nome: "João",
          usuario: "joao",
          senha: "abcd",
          matricula: "RS-0002",
          funcao: "Operador",
          telefone: "",
          status: "ativo",
          meta: 0,
          criadoEm: new Date().toISOString(),
        },
      ];

      fs.writeFileSync(usuariosPath, JSON.stringify(usuariosIniciais, null, 2), "utf8");
      usuarios = usuariosIniciais;
      return;
    }

    const conteudo = fs.readFileSync(usuariosPath, "utf8");
    const lidos = JSON.parse(conteudo || "[]");
    usuarios = Array.isArray(lidos) ? lidos : [];
  } catch (erro) {
    console.error("Erro ao carregar usuários:", erro);
    usuarios = [];
  }
}

function salvarUsuarios() {
  try {
    garantirPastaData();
    fs.writeFileSync(usuariosPath, JSON.stringify(usuarios, null, 2), "utf8");
  } catch (erro) {
    console.error("Erro ao salvar usuários:", erro);
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

  contagens
    .filter((registro) => registro && registro.ativo !== false)
    .forEach((registro) => {
      const usuario = registro.usuario || "desconhecido";

      if (!mapa[usuario]) {
        mapa[usuario] = {
          usuario,
          matricula: registro.matricula || "SEM-MATRICULA",
          totalContado: 0,
          movimentacoes: 0,
          itensUnicos: new Set(),
          ultimaAtualizacao: registro.data || null,
        };
      }

      mapa[usuario].totalContado += Number(registro.quantidade) || 0;
      mapa[usuario].movimentacoes += 1;

      if (registro.codigoBarras) {
        mapa[usuario].itensUnicos.add(String(registro.codigoBarras).trim());
      }

      if (registro.data) {
        mapa[usuario].ultimaAtualizacao = registro.data;
      }
    });

  const ranking = Object.values(mapa).map((item) => ({
    usuario: item.usuario,
    matricula: item.matricula,
    totalContado: item.totalContado,
    movimentacoes: item.movimentacoes,
    itensUnicos: item.itensUnicos.size,
    ultimaAtualizacao: item.ultimaAtualizacao,
  }));

  ranking.sort((a, b) => b.totalContado - a.totalContado);

  const totalGeral = ranking.reduce((acc, item) => acc + item.totalContado, 0);

  return ranking.map((item) => ({
    ...item,
    percentual: totalGeral > 0 ? (item.totalContado / totalGeral) * 100 : 0,
  }));
}

function parseMoeda(valor) {
  if (!valor) return 0;

  return (
    parseFloat(
      String(valor)
        .replace(/\./g, "")
        .replace(",", ".")
        .trim()
    ) || 0
  );
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

function salvarProdutosNoBanco(listaProdutos) {
  db.serialize(() => {
    db.run("DELETE FROM produtos");

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

    listaProdutos.forEach((item) => {
      stmt.run(
        item.codigoBarras || "",
        item.codigo || item.codigoInterno || "",
        item.descricao || "",
        item.categoria || "",
        Number(item.custoUnitario) || 0,
        Number(item.qtdeCongelada) || 0,
        Number(item.qtdeContada) || 0
      );
    });

    stmt.finalize();
  });
}
function carregarProdutosDoBanco(callback = null) {
  db.all("SELECT * FROM produtos", [], (err, rows) => {
    if (err) {
      console.error("Erro ao carregar produtos do banco:", err);
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
app.get("/", autenticar, permitirSomenteLiderOuAdmin, (req, res) =>
res.sendFile(caminhoPublico("index.html"))
);
app.get("/coleta-mobile", autenticar, permitirSomenteOperador, (req, res) =>
  res.sendFile(caminhoPublico("contagem-mobile.html"))
);

app.get("/modo-operacao", autenticar, (req, res) => {
  return res.json({
    modoOperacao: modoOperacao === "sem-base" ? "sem-base" : "com-base",
  });
});

app.post("/modo-operacao", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  try {
    const novoModo = String(req.body?.modoOperacao || "").trim();

    if (novoModo !== "com-base" && novoModo !== "sem-base") {
      return res.status(400).json({
        erro: "Modo de operação inválido.",
      });
    }

    modoOperacao = novoModo;
    salvarModoOperacao();

    return res.json({
      sucesso: true,
      modoOperacao,
      mensagem:
        modoOperacao === "sem-base"
          ? "Modo sem base ativado com sucesso."
          : "Modo com base ativado com sucesso.",
    });
  } catch (erro) {
    console.error("Erro ao salvar modo de operação:", erro);
    return res.status(500).json({
      erro: "Falha ao atualizar modo de operação.",
    });
  }
});
app.get("/sem-base/dados", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  return res.json({
    itens: Array.isArray(contagemSemBase) ? contagemSemBase : [],
  });
});

app.post("/sem-base/leitura", autenticar, (req, res) => {
  try {
    const valorInformado = String(
      req.body?.ean || req.body?.codigo || ""
    ).trim();

    if (!valorInformado) {
      return res.status(400).json({
        erro: "Informe um EAN ou código interno válido.",
      });
    }

    const usuario =
      req.session?.usuario?.usuario ||
      req.session?.usuario?.nome ||
      "sistema";

    const enderecoNumero = String(req.body?.enderecoNumero || "").trim();
    const pareceEan = /^\d{8,14}$/.test(valorInformado);

    const ean = pareceEan ? valorInformado : "";
    const codigo = pareceEan ? "" : valorInformado;

    const index = contagemSemBase.findIndex((item) => {
      const itemEan = String(item.ean || "").trim();
      const itemCodigo = String(item.codigo || "").trim();
      return itemEan === ean && itemCodigo === codigo;
    });

    let itemAtualizado = null;

    if (index >= 0) {
      const item = contagemSemBase[index];
      item.quantidade = (Number(item.quantidade) || 0) + 1;
      item.ultimoUsuario = usuario;

      if (!Array.isArray(item.enderecos)) {
        item.enderecos = [];
      }

      if (enderecoNumero) {
        const idxEndereco = item.enderecos.findIndex(
          (e) => String(e.enderecoNumero || "") === enderecoNumero
        );

        if (idxEndereco >= 0) {
          item.enderecos[idxEndereco].quantidade =
            (Number(item.enderecos[idxEndereco].quantidade) || 0) + 1;
        } else {
          item.enderecos.push({
            enderecoNumero,
            quantidade: 1,
          });
        }
      }

      itemAtualizado = item;
    } else {
      const novoItem = {
        ean,
        codigo,
        quantidade: 1,
        ultimoUsuario: usuario,
        enderecos: enderecoNumero
          ? [
              {
                enderecoNumero,
                quantidade: 1,
              },
            ]
          : [],
      };

      contagemSemBase.push(novoItem);
      itemAtualizado = novoItem;
    }

    salvarContagemSemBase();

    return res.json({
      sucesso: true,
      item: itemAtualizado,
      mensagem: "Leitura registrada com sucesso.",
    });
  } catch (erro) {
    console.error("Erro em /sem-base/leitura:", erro);
    return res.status(500).json({
      erro: "Falha ao registrar leitura do modo sem base.",
    });
  }
});

app.post("/sem-base/reset", autenticar, permitirSomenteLiderOuAdmin, (req, res) => {
  try {
    if (modoOperacao !== "sem-base") {
      return res.status(400).json({
        erro:"O reset sem base só pode ser usado no modo sem base."
      });
     }
    contagemSemBase = [];
    salvarContagemSemBase();

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
app.post("/sem-base/finalizar-endereco", autenticar, (req, res) => {
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
    salvarFinalizacoesSemBase();

    const enderecoAtualizado = registrarEventoEndereco(
      Number(enderecoNumero),
      "finalizacao",
      usuario
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
          layout: "lightHorizontalLines",
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

function carregarFinalizacoesSemBase() {
  try {
    garantirArquivoFinalizacoesSemBase();
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

function salvarFinalizacoesSemBase() {
  try {
    garantirArquivoFinalizacoesSemBase();
    fs.writeFileSync(
      finalizacoesSemBasePath,
      JSON.stringify(finalizacoesSemBase, null, 2),
      "utf8"
    );
  } catch (erro) {
    console.error("Erro ao salvar finalizações sem base:", erro);
  }
}
function carregarContagemSemBase() {
  try {
    garantirArquivoContagemSemBase();
    const bruto = fs.readFileSync(contagemSemBasePath, "utf8") || "[]";
    contagemSemBase = JSON.parse(bruto);

    if (!Array.isArray(contagemSemBase)) {
      contagemSemBase = [];
    }
  } catch (erro) {
    console.error("Erro ao carregar contagem sem base:", erro);
    contagemSemBase = [];
  }
}

function salvarContagemSemBase() {
  try {
    garantirArquivoContagemSemBase();
    fs.writeFileSync(
      contagemSemBasePath,
      JSON.stringify(contagemSemBase, null, 2),
      "utf8"
    );
  } catch (erro) {
    console.error("Erro ao salvar contagem sem base:", erro);
  }
}
function carregarModoOperacao() {
  try {
    garantirPastaData();

    if (!fs.existsSync(configModoPath)) {
      fs.writeFileSync(
        configModoPath,
        JSON.stringify({ modoOperacao: "com-base" }, null, 2),
        "utf8"
      );
      modoOperacao = "com-base";
      return;
    }

    const bruto = fs.readFileSync(configModoPath, "utf8") || "{}";
    const config = JSON.parse(bruto);

    modoOperacao =
      config?.modoOperacao === "sem-base" ? "sem-base" : "com-base";
  } catch (erro) {
    console.error("Erro ao carregar modo de operação:", erro);
    modoOperacao = "com-base";
  }
}

function salvarModoOperacao() {
  try {
    garantirPastaData();

    fs.writeFileSync(
      configModoPath,
      JSON.stringify({ modoOperacao }, null, 2),
      "utf8"
    );
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

app.get("/ranking-usuarios", autenticar, (req, res) => {
  const ranking =
    modoOperacao === "sem-base"
      ? gerarRankingUsuariosSemBase()
      : gerarRankingUsuarios();

  res.json(ranking);
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
    salvarProdutosNoBanco(inventario);

    tipoUltimaImportacao = "Importação base";
    salvarEnderecamentos();
    broadcastInventario();

    fs.unlinkSync(caminhoTemporario);

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
  try {
    if (!req.files || !req.files.arquivo) {
      return res.status(400).send("Nenhum arquivo enviado.");
    }

    garantirPastaData();

    const arquivo = req.files.arquivo;
    const caminhoTemporario = path.join(
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

    const mapaInventarioAtual = new Map(
      inventario.map((item) => [
        String(item.codigo || item.codigoInterno || "").trim(),
        item,
      ])
    );

    const inventarioAtualizado = itensUnicosImportados.map((novoItem) => {
      const chave = String(novoItem.codigo || novoItem.codigoInterno || "").trim();
      const itemAtual = mapaInventarioAtual.get(chave);

      if (!itemAtual) {
        return {
          ...novoItem,
          qtdeContada: 0,
        };
      }

      return {
        ...novoItem,
        qtdeContada: Number(itemAtual.qtdeContada) || 0,
      };
    });

    inventario = inventarioAtualizado;
    salvarProdutosNoBanco(inventario);
    
    tipoUltimaImportacao = "Atualização de saldo";
    broadcastInventario();

    fs.unlinkSync(caminhoTemporario);

    res.redirect("/");
  } catch (erro) {
    console.error("Erro ao importar saldo atual TXT:", erro);
    res.status(500).send(`Erro ao importar saldo atual TXT: ${erro.message}`);
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
        atualizadoEm: new Date().toISOString(),
        total: 0,
        itens: [],
      });
    }
    const filtro = inventario
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

    res.json({
      atualizadoEm: new Date().toISOString(),
      total: filtro.length,
      itens: filtro,
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


app.post("/transmissoes-consolidacao/:id/consolidar", autenticar, (req, res) => {
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
    salvarProdutosNoBanco(inventario);
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

app.post("/transmissoes-consolidacao/consolidar", autenticar, (req, res) => {
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

    salvarContagens();
    salvarEnderecamentos();

    recalcularInventarioComBaseNasContagens();
    salvarProdutosNoBanco(inventario);
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

    const caminhoLogoJpg = path.join(__dirname, "public", "logo-realstock.jpg");
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

app.get("/exportar-pdf", autenticar, (req, res) => {
  let resultado = [];

  if (modoOperacao === "sem-base") {
    resultado = montarLinhasSemBaseParaTabela({
      busca: req.query.busca || "",
    });
  } else {
    resultado = filtrarInventario({
      categoria: req.query.categoria || "",
      ordem: req.query.ordem || "",
      busca: req.query.busca || "",
    });

    resultado = resultado.filter(
      (item) => parseQuantidade(item.qtdeCongelada) > 0
    );
  }

  const colunasSelecionadas = (req.query.colunas || "")
    .split(",")
    .map((c) => c.trim())
    .filter(Boolean);

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
      titulo: "Divergência",
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
      valor: (item) => {
        const texto = String(item.endereco || item.enderecoNumero || "");
        return texto.replace(/\s*\|\s*/g, "\n\n");
      },
      width: 95,
      align: "left",
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

  const printer = new PdfPrinter(fonts);

  const tableBody = [
    colunasAtivas.map((coluna) => ({
      text: mapaColunas[coluna].titulo,
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
noWrap: coluna !== "descricao" && coluna !== "endereco",
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
  const dataInventario = req.query.data || new Date().toLocaleDateString("pt-BR");

  const docDefinition = {
    pageOrientation: "landscape",
    pageMargins: [16, 18, 16, 24],

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
              text: `Gerado em: ${new Date().toLocaleString("pt-BR")}`,
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
          widths: colunasAtivas.map((coluna) => mapaColunas[coluna].width),
          body: tableBody,
        },
        layout: {
          fillColor: function (rowIndex) {
            if (rowIndex === 0) return "#1e293b";
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
            return 5;
          },
          paddingRight: function () {
            return 5;
          },
          paddingTop: function () {
            return 4;
          },
          paddingBottom: function () {
            return 4;
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
        fontSize: 9,
        color: "#ffffff",
      },
      tableCell: {
        fontSize: 9,
        color: "#0f172a",
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
      const finalizacoes = Array.isArray(end.finalizacoes) ? end.finalizacoes : [];

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
app.delete('/excluir-finalizacao/:finalizacaoId', autenticar, (req, res) => {
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

    salvarEnderecamentos();
    recalcularInventarioComBaseNasContagens();
salvarProdutosNoBanco(inventario);
broadcastInventario();

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
app.post("/encerrar-inventario", autenticar, (req, res) => {
  try {
    garantirPastaEncerramentos();

    if (!inventario.length) {
      return res.status(400).json({
        erro: "Não há inventário ativo para encerrar."
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

    fs.writeFileSync(
      path.join(pastaEncerramento, "inventario-final.csv"),
      gerarCsvInventario(inventario),
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

    resetarSistemaAposEncerramento();
    broadcastInventario();

    return res.json({
      sucesso: true,
      nomePasta,
      mensagem: "Inventário encerrado com sucesso."
    });
  } catch (erro) {
    console.error("Erro ao encerrar inventário:", erro);
    return res.status(500).json({
      erro: "Erro ao encerrar inventário."
    });
  }
});
app.post("/restaurar-ultimo-encerramento", autenticar, (req, res) => {
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
    enderecamentos = Array.isArray(snapshot.enderecamentos) ? snapshot.enderecamentos.map(normalizarEnderecoSalvo) : [];
    auditoriaImportacao = snapshot.auditoriaImportacao || {
      totalImportadoBruto: 0,
      totalUnicosBruto: 0,
      duplicatasRemovidas: 0,
      itensZeradosIgnorados: 0,
    };
    itemAuditoriaAtual = null;

    salvarContagens();
    salvarEnderecamentos();
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

app.get("/exportar-checklist-encerramento-pdf-auto", autenticar, (req, res) => {
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
          fontSize: 9,
          color: "#ffffff",
        },
        tableCell: {
          fontSize: 9,
          color: "#0f172a",
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

app.get("/enderecamentos", autenticar, (req, res) => {
  res.json(enderecamentos);
});

app.get("/enderecamentos/posicoes-pendentes", autenticar, (req, res) => {
  try {
    const pendentes = [];

    enderecamentos.forEach((endereco) => {
      const inicio = Number(endereco.inicio) || 0;
      const fim = Number(endereco.fim) || 0;

      const finalizacoesAtivas = Array.isArray(endereco.finalizacoes)
        ? endereco.finalizacoes.filter((f) => !f.excluida)
        : [];

      const transmissoesAtivas = Array.isArray(endereco.transmissoes)
        ? endereco.transmissoes.filter((t) => !t.excluida)
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
app.get('/enderecos-itens-contados/pdf', autenticar, (req, res) => {
  try {
    const {
      endereco = '',
      termo = '',
      somenteAtivos = 'true'
    } = req.query;

    const termoBusca = String(termo || '').trim().toLowerCase();
    const enderecoBusca = String(endereco || '').trim();

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
          codigoBarras: item.codigoBarras || '',
          codigoInterno: produto?.codigo || produto?.codigoInterno || '',
          descricao: produto?.descricao || 'Item não encontrado',
          endereco: `${enderecoObj?.nome || 'ENDEREÇO'} • ${Number(item.enderecoNumero) || 0}`,
          quantidade: Number(item.quantidade) || 0,
          origem: item.origem || 'coleta',
          usuario: item.usuario || '--',
          data: item.data
            ? new Date(item.data).toLocaleString('pt-BR')
            : '--'
        };
      })
      .sort((a, b) => {
        const enderecoA = Number(String(a.endereco).split('•')[1] || 0);
        const enderecoB = Number(String(b.endereco).split('•')[1] || 0);

        if (enderecoA !== enderecoB) {
          return enderecoA - enderecoB;
        }

        return String(a.descricao || '').localeCompare(String(b.descricao || ''), 'pt-BR');
      });

    const printer = new PdfPrinter(fonts);

    const body = [
      [
        { text: 'EAN', style: 'tableHeader' },
        { text: 'Cód. Interno', style: 'tableHeader' },
        { text: 'Descrição', style: 'tableHeader' },
        { text: 'Endereço', style: 'tableHeader' },
        { text: 'Quantidade contada no endereço', style: 'tableHeader' },
        { text: 'Origem', style: 'tableHeader' },
        { text: 'Usuário', style: 'tableHeader' },
        { text: 'Data', style: 'tableHeader' },
      ],
      ...linhas.map((item) => ([
        { text: item.codigoBarras, style: 'tableCell' },
        { text: item.codigoInterno, style: 'tableCell' },
        { text: item.descricao, style: 'tableCell' },
        { text: item.endereco, style: 'tableCell' },
        { text: String(item.quantidade).replace('.', ','), style: 'tableCell', alignment: 'right' },
        { text: item.origem, style: 'tableCell' },
        { text: item.usuario, style: 'tableCell' },
        { text: item.data, style: 'tableCell' },
      ]))
    ];

    const docDefinition = {
      pageOrientation: 'landscape',
      pageMargins: [20, 20, 20, 20],
      content: [
        {
          text: 'Itens contados por endereço',
          style: 'title'
        },
        {
          text: `Gerado em: ${new Date().toLocaleString('pt-BR')}`,
          style: 'subTitle',
          margin: [0, 0, 0, 10]
        },
        {
          table: {
            headerRows: 1,
            widths: [90, 75, '*', 90, 90, 55, 70, 85],
            body
          },
          layout: {
            fillColor: function (rowIndex) {
              return rowIndex === 0 ? '#e5e7eb' : null;
            }
          }
        }
      ],
      styles: {
        title: {
          fontSize: 16,
          bold: true,
          color: '#162f4a'
        },
        subTitle: {
          fontSize: 9,
          color: '#475569'
        },
        tableHeader: {
          bold: true,
          fontSize: 9,
          color: '#0f172a',
          margin: [0, 4, 0, 4]
        },
        tableCell: {
          fontSize: 8,
          margin: [0, 4, 0, 4]
        }
      },
      defaultStyle: {
        font: 'Helvetica'
      }
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      'inline; filename="itens-contados-por-endereco.pdf"'
    );

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

app.post("/usuarios", autenticar, (req, res) => {
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
    salvarUsuarios();

    return res.json({
      sucesso: true,
      usuario: novoUsuario,
    });
  } catch (erro) {
    console.error("Erro ao criar usuário:", erro);
    return res.status(500).json({ erro: "Falha ao criar usuário." });
  }
});

app.put("/usuarios/:id/meta", autenticar, (req, res) => {
  try {
    const { id } = req.params;
    const { meta } = req.body || {};

    const usuario = usuarios.find((u) => String(u.id) === String(id));

    if (!usuario) {
      return res.status(404).json({ erro: "Usuário não encontrado." });
    }

    usuario.meta = Number(meta) || 0;
    salvarUsuarios();

    res.json({ sucesso: true, usuario });
  } catch (erro) {
    console.error("Erro ao atualizar meta:", erro);
    res.status(500).json({ erro: "Falha ao atualizar meta." });
  }
});

app.put("/usuarios/:id", autenticar, (req, res) => {
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

    salvarUsuarios();

    return res.json({ ok: true, usuario: usuarios[usuarioIndex] });
  } catch (erro) {
    console.error("Erro ao editar usuário:", erro);
    return res.status(500).json({ erro: "Falha ao editar usuário." });
  }
});

app.patch("/usuarios/:id/status", autenticar, (req, res) => {
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
    salvarUsuarios();

    return res.json({ ok: true, usuario });
  } catch (erro) {
    console.error("Erro ao alterar status do usuário:", erro);
    return res.status(500).json({ erro: "Falha ao alterar status." });
  }
});

app.delete("/usuarios/:id", autenticar, (req, res) => {
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
    salvarUsuarios();

    return res.json({ ok: true });
  } catch (erro) {
    console.error("Erro ao excluir usuário:", erro);
    return res.status(500).json({ erro: "Falha ao excluir usuário." });
  }
});
app.get("/tipo-ultima-importacao", autenticar, (req, res) => {
  res.json({ tipoUltimaImportacao });
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

carregarEnderecamentos();
carregarUsuarios();
carregarContagens();
carregarLayoutTxt();
carregarLayoutsSalvos();
carregarContagemSemBase();
carregarModoOperacao();
carregarProdutosDoBanco(() => {
  server.listen(port, () =>
    console.log(`Servidor rodando em http://localhost:${port}`)
  );
});