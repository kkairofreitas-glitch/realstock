<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8" />
  <title>Histórico de Alterações</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
      background: #f5f5f5;
      transition: background 0.3s, color 0.3s;
    }
    body.dark {
      background: #1e1e1e;
      color: #ddd;
    }
    h1 {
      color: #1a4731;
      text-align: center;
      margin-bottom: 20px;
    }
    body.dark h1 { color: #9ccc9c; }
    table {
      width: 100%;
      border-collapse: collapse;
      background: white;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    body.dark table { background: #2b2b2b; }
    th, td {
      padding: 10px;
      text-align: left;
      border-bottom: 1px solid #ddd;
    }
    body.dark th, body.dark td { border-bottom: 1px solid #444; }
    th {
      background-color: #1a4731;
      color: white;
      font-weight: bold;
      text-transform: uppercase;
      font-size: 14px;
    }
    tr:nth-child(even) { background-color: #f9f9f9; }
    body.dark tr:nth-child(even) { background-color: #333; }
    tr:hover { background-color: #e2f0d9; }
    body.dark tr:hover { background-color: #3a523a; }
    .positivo { color: green; font-weight: bold; }
    .negativo { color: red; font-weight: bold; }
    /* Botões fixos */
    .top-buttons {
      position: fixed;
      top: 10px;
      left: 10px;
      display: flex;
      gap: 10px;
      z-index: 10;
    }
    .btn {
      padding: 8px 12px;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      background-color: #1a4731;
      color: white;
      font-size: 14px;
      position: relative;
    }
    .btn:hover { background-color: #145524; }
    /* Tooltip */
    .btn[data-tooltip]:hover::after {
      content: attr(data-tooltip);
      position: absolute;
      bottom: -28px;
      left: 50%;
      transform: translateX(-50%);
      background: black;
      color: white;
      padding: 5px 8px;
      border-radius: 4px;
      white-space: nowrap;
      font-size: 12px;
      opacity: 0.85;
    }
    /* Confirmação visual */
    #confirmacao {
      position: fixed;
      top: 20px;
      right: 20px;
      background: #4caf50;
      color: white;
      padding: 10px 15px;
      border-radius: 6px;
      display: none;
      font-weight: bold;
    }
  </style>
</head>
<body>
  <!-- Botões -->
  <div class="top-buttons">
    <button class="btn" onclick="window.location.href='/'" data-tooltip="Voltar ao inventário">⬅ Voltar</button>
    <button class="btn" onclick="alternarDarkMode()" data-tooltip="Alternar tema">🌙 Dark Mode</button>
  </div>

  <h1>Histórico de Alterações</h1>

  <table>
    <thead>
      <tr>
        <th>Usuário</th>
        <th>Código de Barras</th>
        <th>Campo</th>
        <th>Valor Antigo</th>
        <th>Valor Novo</th>
        <th>Data</th>
      </tr>
    </thead>
    <tbody id="historico-corpo"></tbody>
  </table>

  <!-- Confirmação -->
  <div id="confirmacao">✔ Dados carregados</div>

  <script>
    async function carregarHistorico() {
      const res = await fetch('/historico-dados');
      const dados = await res.json();
      const tbody = document.getElementById('historico-corpo');
      tbody.innerHTML = '';

      dados.forEach(item => {
        const tr = document.createElement('tr');
        const positivo = Number(item.valorNovo) > Number(item.valorAntigo);
        const negativo = Number(item.valorNovo) < Number(item.valorAntigo);

        tr.innerHTML = `
          <td>${item.usuario}</td>
          <td>${item.codigoBarras}</td>
          <td>${item.campo}</td>
          <td class="${negativo ? 'negativo' : ''}">${item.valorAntigo}</td>
          <td class="${positivo ? 'positivo' : negativo ? 'negativo' : ''}">${item.valorNovo}</td>
          <td>${new Date(item.data).toLocaleString()}</td>
        `;
        tbody.appendChild(tr);
      });

      mostrarConfirmacao();
    }

    function mostrarConfirmacao() {
      const conf = document.getElementById('confirmacao');
      conf.style.display = 'block';
      setTimeout(() => conf.style.display = 'none', 2000);
    }

    function alternarDarkMode() {
      document.body.classList.toggle('dark');
    }

    carregarHistorico();
  </script>
</body>
</html>
