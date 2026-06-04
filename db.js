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

module.exports = {
  pool,
  testarConexao,
  criarTabelas,
  carregarUsuariosPostgres,
  salvarUsuarioPostgres,
  salvarUsuariosPostgres,
};