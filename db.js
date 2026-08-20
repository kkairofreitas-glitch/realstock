const { Pool } = require("pg");

console.log("DATABASE_URL carregada?", process.env.DATABASE_URL ? "SIM" : "NÃO");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false,
  },
  connectionTimeoutMillis: 5000,
  idleTimeoutMillis: 5000,
});

async function testarConexao() {
  console.log("Testando conexão PostgreSQL...");

  try {
    const resultado = await pool.query("SELECT NOW()");
    console.log("✅ PostgreSQL conectado:", resultado.rows[0]);
  } catch (erro) {
    console.error("❌ Erro PostgreSQL:", erro.message);
  }
}

async function criarTabelas() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        nome TEXT,
        usuario TEXT UNIQUE NOT NULL,
        senha TEXT NOT NULL,
        matricula TEXT,
        funcao TEXT,
        telefone TEXT,
        status TEXT DEFAULT 'ativo',
        meta NUMERIC DEFAULT 0,
        criado_em TIMESTAMP DEFAULT NOW(),
        atualizado_em TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS produtos (
        id SERIAL PRIMARY KEY,
        codigo_barras TEXT,
        codigo TEXT,
        codigo_interno TEXT,
        descricao TEXT,
        categoria TEXT,
        custo_unitario NUMERIC DEFAULT 0,
        qtde_congelada NUMERIC DEFAULT 0,
        qtde_contada NUMERIC DEFAULT 0,
        criado_em TIMESTAMP DEFAULT NOW(),
        atualizado_em TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS contagens (
        id TEXT PRIMARY KEY,
        usuario TEXT,
        matricula TEXT,
        codigo_barras TEXT,
        codigo TEXT,
        quantidade NUMERIC DEFAULT 0,
        endereco_id INTEGER,
        endereco_numero INTEGER,
        finalizacao_id TEXT,
        ativo BOOLEAN DEFAULT true,
        status_consolidacao TEXT DEFAULT 'consolidado',
        consolidado_em TIMESTAMP,
        consolidado_por TEXT,
        data TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS enderecamentos (
        id INTEGER PRIMARY KEY,
        tipo TEXT,
        nome TEXT,
        inicio INTEGER,
        fim INTEGER,
        sequencia INTEGER DEFAULT 0,
        status TEXT DEFAULT 'pendente',
        observacoes TEXT,
        dados JSONB DEFAULT '{}'::jsonb,
        transmissoes JSONB DEFAULT '[]'::jsonb,
        finalizacoes JSONB DEFAULT '[]'::jsonb,
        consolidacoes_por_numero JSONB DEFAULT '[]'::jsonb,
        criado_em TIMESTAMP DEFAULT NOW(),
        atualizado_em TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS contagem_sem_base (
        id SERIAL PRIMARY KEY,
        ean TEXT,
        codigo TEXT,
        quantidade NUMERIC DEFAULT 0,
        ultimo_usuario TEXT,
        ultima_leitura_em TIMESTAMP DEFAULT NOW(),
        enderecos JSONB DEFAULT '[]'::jsonb
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS finalizacoes_sem_base (
        id TEXT PRIMARY KEY,
        endereco_numero TEXT,
        usuario TEXT,
        data TIMESTAMP DEFAULT NOW(),
        itens JSONB DEFAULT '[]'::jsonb,
        total_itens_unicos INTEGER DEFAULT 0,
        total_volume NUMERIC DEFAULT 0
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS configuracoes (
        chave TEXT PRIMARY KEY,
        valor JSONB NOT NULL,
        atualizado_em TIMESTAMP DEFAULT NOW()
      );
    `);
    await pool.query(`
  CREATE TABLE IF NOT EXISTS contagens_wms (
    id TEXT PRIMARY KEY,
    endereco_inventario TEXT NOT NULL,
    endereco_wms TEXT NOT NULL,
    codigo_barras TEXT,
    codigo TEXT,
    descricao TEXT,
    quantidade_cliente NUMERIC DEFAULT 0,
    quantidade_contada NUMERIC DEFAULT 0,
    divergencia NUMERIC DEFAULT 0,
    status TEXT DEFAULT 'pendente',
    usuario TEXT,
    data TIMESTAMP DEFAULT NOW(),
    ativo BOOLEAN DEFAULT TRUE,
    excluida_em TIMESTAMP,
    excluida_por TEXT
  )
`);

await pool.query(`
  ALTER TABLE contagens_wms
  ADD COLUMN IF NOT EXISTS ativo BOOLEAN DEFAULT TRUE
`);

await pool.query(`
  ALTER TABLE contagens_wms
  ADD COLUMN IF NOT EXISTS excluida_em TIMESTAMP
`);

await pool.query(`
  ALTER TABLE contagens_wms
  ADD COLUMN IF NOT EXISTS excluida_por TEXT
`);

await pool.query(`
  CREATE TABLE IF NOT EXISTS finalizacoes_wms (
    id TEXT PRIMARY KEY,
    endereco_inventario TEXT NOT NULL,
    endereco_wms TEXT NOT NULL,
    usuario TEXT,
    data TIMESTAMP DEFAULT NOW(),
    total_itens INTEGER DEFAULT 0,
    total_volume NUMERIC DEFAULT 0
  )
`);
await pool.query(`
  CREATE TABLE IF NOT EXISTS base_esperada_wms (
    id SERIAL PRIMARY KEY,
    endereco_wms TEXT NOT NULL,
    codigo_barras TEXT DEFAULT '',
    codigo TEXT DEFAULT '',
    descricao TEXT DEFAULT '',
    quantidade_esperada NUMERIC DEFAULT 0,
    importado_em TIMESTAMP DEFAULT NOW(),
    atualizado_em TIMESTAMP DEFAULT NOW()
  )
`);
await pool.query(`
  CREATE INDEX IF NOT EXISTS idx_base_esperada_wms_endereco
  ON base_esperada_wms (endereco_wms)
`);

await pool.query(`
  CREATE INDEX IF NOT EXISTS idx_base_esperada_wms_codigo_barras
  ON base_esperada_wms (codigo_barras)
`);

await pool.query(`
  CREATE INDEX IF NOT EXISTS idx_base_esperada_wms_codigo
  ON base_esperada_wms (codigo)
`);
await pool.query(`
  CREATE UNIQUE INDEX IF NOT EXISTS idx_base_esperada_wms_unica
  ON base_esperada_wms (
    endereco_wms,
    COALESCE(NULLIF(codigo_barras, ''), codigo)
  )
`);
    console.log("✅ Tabelas PostgreSQL verificadas/criadas.");
  } catch (erro) {
    console.error("❌ Erro ao criar tabelas PostgreSQL:", erro.message);
  }
}

async function carregarUsuariosPostgres() {
  const resultado = await pool.query(`
    SELECT
      id,
      nome,
      usuario,
      senha,
      matricula,
      funcao,
      telefone,
      status,
      meta,
      criado_em
    FROM usuarios
    ORDER BY id ASC
  `);

  return resultado.rows.map((u) => ({
    id: u.id,
    nome: u.nome || "",
    usuario: u.usuario || "",
    senha: u.senha || "",
    matricula: u.matricula || "",
    funcao: u.funcao || "Operador",
    telefone: u.telefone || "",
    status: u.status || "ativo",
    meta: Number(u.meta) || 0,
    criadoEm: u.criado_em || null,
  }));
}

async function salvarUsuarioPostgres(usuario) {
  const resultado = await pool.query(
    `
    INSERT INTO usuarios (
      nome,
      usuario,
      senha,
      matricula,
      funcao,
      telefone,
      status,
      meta
    )
    VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
    ON CONFLICT (usuario)
    DO UPDATE SET
      nome = EXCLUDED.nome,
      senha = EXCLUDED.senha,
      matricula = EXCLUDED.matricula,
      funcao = EXCLUDED.funcao,
      telefone = EXCLUDED.telefone,
      status = EXCLUDED.status,
      meta = EXCLUDED.meta,
      atualizado_em = NOW()
    RETURNING *
    `,
    [
      usuario.nome || "",
      usuario.usuario,
      usuario.senha,
      usuario.matricula || "",
      usuario.funcao || "Operador",
      usuario.telefone || "",
      usuario.status || "ativo",
      Number(usuario.meta) || 0,
    ]
  );

  return resultado.rows[0];
}

async function salvarUsuariosPostgres(listaUsuarios) {
  const lista = Array.isArray(listaUsuarios) ? listaUsuarios : [];

  for (const usuario of lista) {
    if (!usuario?.usuario || !usuario?.senha) continue;
    await salvarUsuarioPostgres(usuario);
  }

  console.log(`✅ Usuários salvos no PostgreSQL: ${lista.length}`);
}
async function salvarProdutosPostgres(listaProdutos = []) {
  await pool.query("DELETE FROM produtos");

  for (const item of listaProdutos) {
    await pool.query(
      `
      INSERT INTO produtos (
        codigo_barras,
        codigo,
        codigo_interno,
        descricao,
        categoria,
        custo_unitario,
        qtde_congelada,
        qtde_contada
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      `,
      [
        item.codigoBarras || "",
        item.codigo || item.codigoInterno || "",
        item.codigoInterno || item.codigo || "",
        item.descricao || "",
        item.categoria || "",
        Number(item.custoUnitario) || 0,
        Number(item.qtdeCongelada) || 0,
        Number(item.qtdeContada) || 0,
      ]
    );
  }

  console.log(`✅ Produtos salvos no PostgreSQL: ${listaProdutos.length}`);
}

async function carregarProdutosPostgres() {
  const resultado = await pool.query(`
    SELECT
      codigo_barras,
      codigo,
      codigo_interno,
      descricao,
      categoria,
      custo_unitario,
      qtde_congelada,
      qtde_contada
    FROM produtos
    ORDER BY id ASC
  `);

  return resultado.rows.map((item) => ({
    codigoBarras: item.codigo_barras || "",
    codigo: item.codigo || "",
    codigoInterno: item.codigo_interno || item.codigo || "",
    descricao: item.descricao || "",
    categoria: item.categoria || "",
    custoUnitario: Number(item.custo_unitario) || 0,
    qtdeCongelada: Number(item.qtde_congelada) || 0,
    qtdeContada: Number(item.qtde_contada) || 0,
  }));
}

async function salvarEnderecamentosPostgres(
  listaEnderecamentos = [],
  opcoes = {}
) {
  const lista =
    Array.isArray(
      listaEnderecamentos
    )
      ? listaEnderecamentos
      : [];

  const permitirVazio =
    opcoes?.permitirVazio === true;


  /*
    ========================================================
    PROTEÇÃO CONTRA APAGAMENTO ACIDENTAL
    ========================================================

    Lista vazia normalmente NÃO apaga o PostgreSQL.

    Somente operações explícitas, como ENCERRAMENTO,
    poderão enviar:

    { permitirVazio: true }
  */
  if (
    lista.length === 0 &&
    !permitirVazio
  ) {
    console.warn(
      "⚠️ salvarEnderecamentosPostgres recebeu lista vazia. " +
      "A tabela enderecamentos não será apagada."
    );

   

    return;
  }


  const client =
    await pool.connect();


  try {
    await client.query(
      "BEGIN"
    );


    /*
      Mantemos a sincronização completa,
      mas somente quando existe conteúdo
      válido para persistir.
    */
    await client.query(
      "DELETE FROM enderecamentos"
    );


    for (const item of lista) {

      await client.query(
        `
        INSERT INTO enderecamentos (
          id,
          tipo,
          nome,
          inicio,
          fim,
          sequencia,
          status,
          observacoes,
          dados,
          transmissoes,
          finalizacoes,
          consolidacoes_por_numero
        )
        VALUES (
          $1,$2,$3,$4,$5,$6,
          $7,$8,$9,$10,$11,$12
        )
        ON CONFLICT (id)
        DO UPDATE SET
          tipo =
            EXCLUDED.tipo,

          nome =
            EXCLUDED.nome,

          inicio =
            EXCLUDED.inicio,

          fim =
            EXCLUDED.fim,

          sequencia =
            EXCLUDED.sequencia,

          status =
            EXCLUDED.status,

          observacoes =
            EXCLUDED.observacoes,

          dados =
            EXCLUDED.dados,

          transmissoes =
            EXCLUDED.transmissoes,

          finalizacoes =
            EXCLUDED.finalizacoes,

          consolidacoes_por_numero =
            EXCLUDED.consolidacoes_por_numero,

          atualizado_em =
            NOW()
        `,
        [
          Number(item.id) || 0,

          item.tipo || "",

          item.nome || "",

          Number(item.inicio) || 0,

          Number(item.fim) || 0,

          Number(item.sequencia) || 0,

          item.status || "pendente",

          item.observacoes || "",

          JSON.stringify(
            item
          ),

          JSON.stringify(
            item.transmissoes || []
          ),

          JSON.stringify(
            item.finalizacoes || []
          ),

          JSON.stringify(
            item.consolidacoesPorNumero ||
            []
          ),
        ]
      );
    }


    await client.query(
      "COMMIT"
    );


    console.log(
      `✅ Endereçamentos salvos no PostgreSQL: ${lista.length}`
    );

  } catch (erro) {

    await client.query(
      "ROLLBACK"
    );


    console.error(
      "Erro ao salvar endereçamentos no PostgreSQL:",
      erro
    );


    throw erro;

  } finally {

    client.release();

  }
}

async function carregarEnderecamentosPostgres() {
  const resultado = await pool.query(`
    SELECT dados
    FROM enderecamentos
    ORDER BY id ASC
  `);

  return resultado.rows.map((row) => row.dados || {});
}


async function salvarContagensPostgres(listaContagens = []) {
  await pool.query("DELETE FROM contagens");

  for (const item of listaContagens) {
    await pool.query(
      `
      INSERT INTO contagens (
        id,
        usuario,
        matricula,
        codigo_barras,
        codigo,
        quantidade,
        endereco_id,
        endereco_numero,
        finalizacao_id,
        ativo,
        status_consolidacao,
        consolidado_em,
        consolidado_por,
        data
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)
      ON CONFLICT (id)
      DO UPDATE SET
        usuario = EXCLUDED.usuario,
        matricula = EXCLUDED.matricula,
        codigo_barras = EXCLUDED.codigo_barras,
        codigo = EXCLUDED.codigo,
        quantidade = EXCLUDED.quantidade,
        endereco_id = EXCLUDED.endereco_id,
        endereco_numero = EXCLUDED.endereco_numero,
        finalizacao_id = EXCLUDED.finalizacao_id,
        ativo = EXCLUDED.ativo,
        status_consolidacao = EXCLUDED.status_consolidacao,
        consolidado_em = EXCLUDED.consolidado_em,
        consolidado_por = EXCLUDED.consolidado_por,
        data = EXCLUDED.data
      `,
      [
        item.id || `CONT-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        item.usuario || "",
        item.matricula || "",
        item.codigoBarras || "",
        item.codigo || "",
        Number(item.quantidade) || 0,
        item.enderecoId ? Number(item.enderecoId) : null,
        item.enderecoNumero ? Number(item.enderecoNumero) : null,
        item.finalizacaoId || null,
        item.ativo !== false,
        item.statusConsolidacao || "consolidado",
        item.consolidadoEm || null,
        item.consolidadoPor || null,
        item.data || new Date().toISOString(),
      ]
    );
  }

  console.log(`✅ Contagens salvas no PostgreSQL: ${listaContagens.length}`);
}

async function carregarContagensPostgres() {
  const resultado = await pool.query(`
    SELECT
      id,
      usuario,
      matricula,
      codigo_barras,
      codigo,
      quantidade,
      endereco_id,
      endereco_numero,
      finalizacao_id,
      ativo,
      status_consolidacao,
      consolidado_em,
      consolidado_por,
      data
    FROM contagens
    ORDER BY data ASC
  `);

  return resultado.rows.map((item) => ({
    id: item.id,
    usuario: item.usuario || "",
    matricula: item.matricula || "",
    codigoBarras: item.codigo_barras || "",
    codigo: item.codigo || "",
    quantidade: Number(item.quantidade) || 0,
    enderecoId: item.endereco_id,
    enderecoNumero: item.endereco_numero,
    finalizacaoId: item.finalizacao_id,
    ativo: item.ativo !== false,
    statusConsolidacao: item.status_consolidacao || "consolidado",
    consolidadoEm: item.consolidado_em || null,
    consolidadoPor: item.consolidado_por || null,
    data: item.data || null,
  }));
}

async function salvarConfiguracaoPostgres(chave, valor) {
  await pool.query(
    `
    INSERT INTO configuracoes (chave, valor, atualizado_em)
    VALUES ($1, $2, NOW())
    ON CONFLICT (chave)
    DO UPDATE SET
      valor = EXCLUDED.valor,
      atualizado_em = NOW()
    `,
    [chave, JSON.stringify(valor || {})]
  );

  console.log(`✅ Configuração salva no PostgreSQL: ${chave}`);
}

async function carregarConfiguracaoPostgres(chave, valorPadrao = {}) {
  const resultado = await pool.query(
    `
    SELECT valor
    FROM configuracoes
    WHERE chave = $1
    LIMIT 1
    `,
    [chave]
  );

  if (!resultado.rows.length) return valorPadrao;

  return resultado.rows[0].valor || valorPadrao;
}

async function limparInventarioPostgres() {
  await pool.query("DELETE FROM produtos");
  await pool.query("DELETE FROM contagens");
  await pool.query("DELETE FROM enderecamentos");
  await pool.query("DELETE FROM contagem_sem_base");
  await pool.query("DELETE FROM finalizacoes_sem_base");
  await pool.query("DELETE FROM configuracoes WHERE chave IN ('dados_inventario')");

  console.log("✅ Inventário limpo no PostgreSQL.");
}
function normalizarTextoWms(valor) {
  return String(valor || '').trim();
}

function normalizarEnderecoWms(valor) {
  return normalizarTextoWms(valor).toUpperCase();
}
async function salvarContagemSemBasePostgres(lista = []) {
  await pool.query("DELETE FROM contagem_sem_base");

  for (const item of Array.isArray(lista) ? lista : []) {
    await pool.query(
      `
      INSERT INTO contagem_sem_base (
        ean,
        codigo,
        quantidade,
        ultimo_usuario,
        ultima_leitura_em,
        enderecos
      )
      VALUES ($1,$2,$3,$4,$5,$6)
      `,
      [
        item.ean || "",
        item.codigo || "",
        Number(item.quantidade) || 0,
        item.ultimoUsuario || "",
        item.ultimaLeituraEm || new Date().toISOString(),
        JSON.stringify(item.enderecos || []),
      ]
    );
  }
}

async function carregarContagemSemBasePostgres() {
  const resultado = await pool.query(`
    SELECT ean, codigo, quantidade, ultimo_usuario, ultima_leitura_em, enderecos
    FROM contagem_sem_base
    ORDER BY id ASC
  `);

  return resultado.rows.map((item) => ({
    ean: item.ean || "",
    codigo: item.codigo || "",
    quantidade: Number(item.quantidade) || 0,
    ultimoUsuario: item.ultimo_usuario || "",
    ultimaLeituraEm: item.ultima_leitura_em || null,
    enderecos: item.enderecos || [],
  }));
}

async function salvarFinalizacoesSemBasePostgres(lista = []) {
  await pool.query("DELETE FROM finalizacoes_sem_base");

  for (const item of Array.isArray(lista) ? lista : []) {
    await pool.query(
      `
      INSERT INTO finalizacoes_sem_base (
        id,
        endereco_numero,
        usuario,
        data,
        itens,
        total_itens_unicos,
        total_volume
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7)
      ON CONFLICT (id)
      DO UPDATE SET
        endereco_numero = EXCLUDED.endereco_numero,
        usuario = EXCLUDED.usuario,
        data = EXCLUDED.data,
        itens = EXCLUDED.itens,
        total_itens_unicos = EXCLUDED.total_itens_unicos,
        total_volume = EXCLUDED.total_volume
      `,
      [
        item.id,
        item.enderecoNumero || "",
        item.usuario || "",
        item.data || new Date().toISOString(),
        JSON.stringify(item.itens || []),
        Number(item.totalItensUnicos) || 0,
        Number(item.totalVolume) || 0,
      ]
    );
  }
}

async function carregarFinalizacoesSemBasePostgres() {
  const resultado = await pool.query(`
    SELECT id, endereco_numero, usuario, data, itens, total_itens_unicos, total_volume
    FROM finalizacoes_sem_base
    ORDER BY data ASC
  `);

  return resultado.rows.map((item) => ({
    id: item.id,
    enderecoNumero: item.endereco_numero || "",
    usuario: item.usuario || "",
    data: item.data || null,
    itens: item.itens || [],
    totalItensUnicos: Number(item.total_itens_unicos) || 0,
    totalVolume: Number(item.total_volume) || 0,
  }));
}
async function salvarBaseEsperadaWmsPostgres(item) {
  const enderecoWms = normalizarEnderecoWms(
    item.enderecoWms
  );

  const codigoBarras = normalizarTextoWms(
    item.codigoBarras
  );

  const codigo = normalizarTextoWms(
    item.codigo
  );

  const descricao = normalizarTextoWms(
    item.descricao
  );

  const quantidadeEsperada =
    Number(item.quantidadeEsperada) || 0;

  if (!enderecoWms) {
    throw new Error(
      'O endereço WMS é obrigatório.'
    );
  }

  if (!codigoBarras && !codigo) {
    throw new Error(
      'Informe o código de barras ou o código interno.'
    );
  }

  const resultado = await pool.query(
    `
    INSERT INTO base_esperada_wms (
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_esperada,
      importado_em,
      atualizado_em
    )
    VALUES ($1, $2, $3, $4, $5, NOW(), NOW())

    ON CONFLICT (
      endereco_wms,
      (COALESCE(NULLIF(codigo_barras, ''), codigo))
    )
    DO UPDATE SET
      codigo_barras = EXCLUDED.codigo_barras,
      codigo = EXCLUDED.codigo,
      descricao = EXCLUDED.descricao,
      quantidade_esperada =
        EXCLUDED.quantidade_esperada,
      atualizado_em = NOW()

    RETURNING
      id,
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_esperada,
      importado_em,
      atualizado_em
    `,
    [
      enderecoWms,
      codigoBarras,
      codigo,
      descricao,
      quantidadeEsperada,
    ]
  );

  const registro = resultado.rows[0];

  return {
    id: registro.id,
    enderecoWms:
      registro.endereco_wms || '',
    codigoBarras:
      registro.codigo_barras || '',
    codigo:
      registro.codigo || '',
    descricao:
      registro.descricao || '',
    quantidadeEsperada:
      Number(registro.quantidade_esperada) || 0,
    importadoEm:
      registro.importado_em || null,
    atualizadoEm:
      registro.atualizado_em || null,
  };
}
async function salvarBaseEsperadaWmsLotePostgres(
  itens = [],
  substituirBase = false
) {
  const cliente = await pool.connect();

  try {
    await cliente.query('BEGIN');

    if (substituirBase) {
      await cliente.query(`
        DELETE FROM base_esperada_wms
      `);
    }

    let registrosSalvos = 0;
    let registrosIgnorados = 0;

    for (const item of itens) {
      const enderecoWms =
        normalizarEnderecoWms(
          item.enderecoWms
        );

      const codigoBarras =
        normalizarTextoWms(
          item.codigoBarras
        );

      const codigo =
        normalizarTextoWms(
          item.codigo
        );

      const descricao =
        normalizarTextoWms(
          item.descricao
        );

      const quantidadeEsperada =
        Number(item.quantidadeEsperada) || 0;

      if (
        !enderecoWms ||
        (!codigoBarras && !codigo)
      ) {
        registrosIgnorados++;
        continue;
      }

      await cliente.query(
        `
        INSERT INTO base_esperada_wms (
          endereco_wms,
          codigo_barras,
          codigo,
          descricao,
          quantidade_esperada,
          importado_em,
          atualizado_em
        )
        VALUES (
          $1, $2, $3, $4, $5, NOW(), NOW()
        )

        ON CONFLICT (
          endereco_wms,
          (COALESCE(NULLIF(codigo_barras, ''), codigo))
        )
        DO UPDATE SET
          codigo_barras =
            EXCLUDED.codigo_barras,
          codigo =
            EXCLUDED.codigo,
          descricao =
            EXCLUDED.descricao,
          quantidade_esperada =
            EXCLUDED.quantidade_esperada,
          atualizado_em = NOW()
        `,
        [
          enderecoWms,
          codigoBarras,
          codigo,
          descricao,
          quantidadeEsperada,
        ]
      );

      registrosSalvos++;
    }

    await cliente.query('COMMIT');

    return {
      registrosRecebidos: itens.length,
      registrosSalvos,
      registrosIgnorados,
      substituirBase,
    };
  } catch (erro) {
    await cliente.query('ROLLBACK');
    throw erro;
  } finally {
    cliente.release();
  }
}
async function buscarBaseEsperadaWmsPostgres({
  enderecoWms,
  codigoBarras,
  codigo,
}) {
  const enderecoNormalizado =
    normalizarEnderecoWms(enderecoWms);

  const eanNormalizado =
    normalizarTextoWms(codigoBarras);

  const codigoNormalizado =
    normalizarTextoWms(codigo);

  if (!enderecoNormalizado) {
    return null;
  }

  if (!eanNormalizado && !codigoNormalizado) {
    return null;
  }

  const resultado = await pool.query(
    `
    SELECT
      id,
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_esperada,
      importado_em,
      atualizado_em
    FROM base_esperada_wms
    WHERE endereco_wms = $1
      AND (
        (
          $2 <> ''
          AND codigo_barras = $2
        )
        OR
        (
          $3 <> ''
          AND codigo = $3
        )
      )
    ORDER BY
      CASE
        WHEN $2 <> ''
          AND codigo_barras = $2
        THEN 0
        ELSE 1
      END
    LIMIT 1
    `,
    [
      enderecoNormalizado,
      eanNormalizado,
      codigoNormalizado,
    ]
  );

  if (!resultado.rows.length) {
    return null;
  }

  const item = resultado.rows[0];

  return {
    id: item.id,
    enderecoWms:
      item.endereco_wms || '',
    codigoBarras:
      item.codigo_barras || '',
    codigo:
      item.codigo || '',
    descricao:
      item.descricao || '',
    quantidadeEsperada:
      Number(item.quantidade_esperada) || 0,
    importadoEm:
      item.importado_em || null,
    atualizadoEm:
      item.atualizado_em || null,
  };
}
async function buscarBaseEsperadaWmsPorIdPostgres(id) {
  const idNumerico = Number(id);

  if (!Number.isInteger(idNumerico) || idNumerico <= 0) {
    return null;
  }

  const resultado = await pool.query(
    `
    SELECT
      id,
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_esperada,
      importado_em,
      atualizado_em
    FROM base_esperada_wms
    WHERE id = $1
    LIMIT 1
    `,
    [idNumerico]
  );

  if (!resultado.rows.length) {
    return null;
  }

  const item = resultado.rows[0];

  return {
    id: item.id,
    enderecoWms: item.endereco_wms || "",
    codigoBarras: item.codigo_barras || "",
    codigo: item.codigo || "",
    descricao: item.descricao || "",
    quantidadeEsperada:
      Number(item.quantidade_esperada) || 0,
    importadoEm: item.importado_em || null,
    atualizadoEm: item.atualizado_em || null,
  };
}
async function atualizarBaseEsperadaWmsPostgres(
  id,
  item
) {
  const idNumerico = Number(id);

  if (!Number.isInteger(idNumerico) || idNumerico <= 0) {
    throw new Error(
      "Identificador inválido da base esperada WMS."
    );
  }

  const enderecoWms = normalizarEnderecoWms(
    item.enderecoWms
  );

  const codigoBarras = normalizarTextoWms(
    item.codigoBarras
  );

  const codigo = normalizarTextoWms(
    item.codigo
  );

  const descricao = normalizarTextoWms(
    item.descricao
  );

  const quantidadeEsperada = Number(
    item.quantidadeEsperada
  );

  if (!enderecoWms) {
    throw new Error(
      "O endereço WMS é obrigatório."
    );
  }

  if (!codigoBarras && !codigo) {
    throw new Error(
      "Informe o código de barras ou o código interno."
    );
  }

  if (
    !Number.isFinite(quantidadeEsperada) ||
    quantidadeEsperada < 0
  ) {
    throw new Error(
      "A quantidade esperada deve ser um número maior ou igual a zero."
    );
  }

  try {
    const resultado = await pool.query(
      `
      UPDATE base_esperada_wms
      SET
        endereco_wms = $2,
        codigo_barras = $3,
        codigo = $4,
        descricao = $5,
        quantidade_esperada = $6,
        atualizado_em = NOW()
      WHERE id = $1
      RETURNING
        id,
        endereco_wms,
        codigo_barras,
        codigo,
        descricao,
        quantidade_esperada,
        importado_em,
        atualizado_em
      `,
      [
        idNumerico,
        enderecoWms,
        codigoBarras,
        codigo,
        descricao,
        quantidadeEsperada,
      ]
    );

    if (!resultado.rows.length) {
      return null;
    }

    const registro = resultado.rows[0];

    return {
      id: registro.id,
      enderecoWms:
        registro.endereco_wms || "",
      codigoBarras:
        registro.codigo_barras || "",
      codigo:
        registro.codigo || "",
      descricao:
        registro.descricao || "",
      quantidadeEsperada:
        Number(registro.quantidade_esperada) || 0,
      importadoEm:
        registro.importado_em || null,
      atualizadoEm:
        registro.atualizado_em || null,
    };
  } catch (erro) {
    if (erro?.code === "23505") {
      throw new Error(
        "Já existe este produto no endereço WMS informado."
      );
    }

    throw erro;
  }
}
async function excluirBaseEsperadaWmsPostgres(id) {
  const idNumerico = Number(id);

  if (!Number.isInteger(idNumerico) || idNumerico <= 0) {
    return false;
  }

  const resultado = await pool.query(
    `
    DELETE FROM base_esperada_wms
    WHERE id = $1
    RETURNING id
    `,
    [idNumerico]
  );

  return resultado.rowCount > 0;
}

async function carregarBaseEsperadaWmsPostgres({
  enderecoWms = '',
  busca = '',
} = {}) {
  const filtros = [];
  const valores = [];

  const enderecoNormalizado =
    normalizarEnderecoWms(enderecoWms);

  const buscaNormalizada =
    normalizarTextoWms(busca);

  if (enderecoNormalizado) {
    valores.push(enderecoNormalizado);

    filtros.push(
      `endereco_wms = $${valores.length}`
    );
  }

  if (buscaNormalizada) {
    valores.push(`%${buscaNormalizada}%`);

    filtros.push(`
      (
        codigo_barras ILIKE $${valores.length}
        OR codigo ILIKE $${valores.length}
        OR descricao ILIKE $${valores.length}
      )
    `);
  }

  const where =
    filtros.length
      ? `WHERE ${filtros.join(' AND ')}`
      : '';

  const resultado = await pool.query(
    `
    SELECT
      id,
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_esperada,
      importado_em,
      atualizado_em
    FROM base_esperada_wms
    ${where}
    ORDER BY
      endereco_wms ASC,
      descricao ASC,
      codigo_barras ASC,
      codigo ASC
    `,
    valores
  );

  return resultado.rows.map((item) => ({
    id: item.id,
    enderecoWms:
      item.endereco_wms || '',
    codigoBarras:
      item.codigo_barras || '',
    codigo:
      item.codigo || '',
    descricao:
      item.descricao || '',
    quantidadeEsperada:
      Number(item.quantidade_esperada) || 0,
    importadoEm:
      item.importado_em || null,
    atualizadoEm:
      item.atualizado_em || null,
  }));
}
async function carregarContagensWmsPostgres() {
  const resultado = await pool.query(`
    SELECT
      id,
      endereco_inventario,
      endereco_wms,
      codigo_barras,
      codigo,
      descricao,
      quantidade_cliente,
      quantidade_contada,
      divergencia,
      status,
      usuario,
      data,
      ativo
    FROM contagens_wms
    WHERE ativo IS DISTINCT FROM FALSE
    ORDER BY data DESC
  `);

  return resultado.rows.map((item) => ({
    id: item.id,
    enderecoInventario: item.endereco_inventario || "",
    enderecoWms: item.endereco_wms || "",
    codigoBarras: item.codigo_barras || "",
    codigo: item.codigo || "",
    descricao: item.descricao || "",
    quantidadeCliente: Number(item.quantidade_cliente) || 0,
    quantidadeContada: Number(item.quantidade_contada) || 0,
    divergencia: Number(item.divergencia) || 0,
    status: item.status || "pendente",
    usuario: item.usuario || "",
    data: item.data || null,
    ativo: item.ativo !== false,
  }));
}

async function salvarContagemWmsPostgres(item) {
  const quantidadeCliente = Number(item.quantidadeCliente) || 0;
  const quantidadeContada = Number(item.quantidadeContada) || 0;
  const divergencia = quantidadeContada - quantidadeCliente;

  const status =
    divergencia === 0
      ? "ok"
      : divergencia > 0
      ? "sobrando"
      : "faltando";

      await pool.query(
        `
        INSERT INTO contagens_wms (
          id,
          endereco_inventario,
          endereco_wms,
          codigo_barras,
          codigo,
          descricao,
          quantidade_cliente,
          quantidade_contada,
          divergencia,
          status,
          usuario,
          data,
          ativo,
          excluida_em,
          excluida_por
        )
        VALUES (
          $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,
          TRUE,NULL,NULL
        )
        ON CONFLICT (id)
        DO UPDATE SET
          endereco_inventario = EXCLUDED.endereco_inventario,
          endereco_wms = EXCLUDED.endereco_wms,
          codigo_barras = EXCLUDED.codigo_barras,
          codigo = EXCLUDED.codigo,
          descricao = EXCLUDED.descricao,
          quantidade_cliente = EXCLUDED.quantidade_cliente,
          quantidade_contada = EXCLUDED.quantidade_contada,
          divergencia = EXCLUDED.divergencia,
          status = EXCLUDED.status,
          usuario = EXCLUDED.usuario,
          data = EXCLUDED.data,
          ativo = TRUE,
          excluida_em = NULL,
          excluida_por = NULL
        `,
        [
          item.id,
          item.enderecoInventario || "",
          item.enderecoWms || "",
          item.codigoBarras || "",
          item.codigo || "",
          item.descricao || "",
          quantidadeCliente,
          quantidadeContada,
          divergencia,
          status,
          item.usuario || "",
          item.data || new Date().toISOString(),
        ]
      );

  return {
    ...item,
    quantidadeCliente,
    quantidadeContada,
    divergencia,
    status,
  };
}
async function excluirContagemWmsPostgres(id, usuario = "") {
  const resultado = await pool.query(
    `
    UPDATE contagens_wms
    SET
      ativo = FALSE,
      excluida_em = NOW(),
      excluida_por = $2
    WHERE id = $1
      AND ativo IS DISTINCT FROM FALSE
    RETURNING id
    `,
    [id, usuario]
  );

  return resultado.rowCount > 0;
}

async function salvarFinalizacaoWmsPostgres(finalizacao) {
  await pool.query(
    `
    INSERT INTO finalizacoes_wms (
      id,
      endereco_inventario,
      endereco_wms,
      usuario,
      data,
      total_itens,
      total_volume
    )
    VALUES ($1,$2,$3,$4,$5,$6,$7)
    ON CONFLICT (id)
    DO UPDATE SET
      endereco_inventario = EXCLUDED.endereco_inventario,
      endereco_wms = EXCLUDED.endereco_wms,
      usuario = EXCLUDED.usuario,
      data = EXCLUDED.data,
      total_itens = EXCLUDED.total_itens,
      total_volume = EXCLUDED.total_volume
    `,
    [
      finalizacao.id,
      finalizacao.enderecoInventario || "",
      finalizacao.enderecoWms || "",
      finalizacao.usuario || "",
      finalizacao.data || new Date().toISOString(),
      Number(finalizacao.totalItens) || 0,
      Number(finalizacao.totalVolume) || 0,
    ]
  );

  return finalizacao;
}

async function carregarFinalizacoesWmsPostgres() {
  const resultado = await pool.query(`
    SELECT
      id,
      endereco_inventario,
      endereco_wms,
      usuario,
      data,
      total_itens,
      total_volume
    FROM finalizacoes_wms
    ORDER BY data DESC
  `);

  return resultado.rows.map((item) => ({
    id: item.id,
    enderecoInventario:
      item.endereco_inventario || "",
    enderecoWms: item.endereco_wms || "",
    usuario: item.usuario || "",
    data: item.data || null,
    totalItens:
      Number(item.total_itens) || 0,
    totalVolume:
      Number(item.total_volume) || 0,
  }));
}
module.exports = {
  pool,
  testarConexao,
  criarTabelas,
  carregarUsuariosPostgres,
  salvarUsuarioPostgres,
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
};