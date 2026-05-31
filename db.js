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

module.exports = {
  pool,
  testarConexao,
};