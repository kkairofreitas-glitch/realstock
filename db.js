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

async function salvarEnderecamentosPostgres(listaEnderecamentos = []) {
  await pool.query("DELETE FROM enderecamentos");

  for (const item of listaEnderecamentos) {
    await pool.query(
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
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
      ON CONFLICT (id)
      DO UPDATE SET
        tipo = EXCLUDED.tipo,
        nome = EXCLUDED.nome,
        inicio = EXCLUDED.inicio,
        fim = EXCLUDED.fim,
        sequencia = EXCLUDED.sequencia,
        status = EXCLUDED.status,
        observacoes = EXCLUDED.observacoes,
        dados = EXCLUDED.dados,
        transmissoes = EXCLUDED.transmissoes,
        finalizacoes = EXCLUDED.finalizacoes,
        consolidacoes_por_numero = EXCLUDED.consolidacoes_por_numero,
        atualizado_em = NOW()
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
        JSON.stringify(item),
        JSON.stringify(item.transmissoes || []),
        JSON.stringify(item.finalizacoes || []),
        JSON.stringify(item.consolidacoesPorNumero || []),
      ]
    );
  }

  console.log(`✅ Endereçamentos salvos no PostgreSQL: ${listaEnderecamentos.length}`);
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
};