// Nomes dos arquivos gerados na exportação
const NOME_ARQUIVO_ZERADOS = "produtos_zerados.xlsx";
const NOME_ARQUIVO_ENCONTRADOS = "produtos_incluir.xlsx";

function criarElemento(tipo, atributos, texto) {
  const elemento = document.createElement(tipo);
  if (atributos) {
    Object.keys(atributos).forEach((chave) => {
      elemento.setAttribute(chave, atributos[chave]);
    });
  }
  if (texto) elemento.textContent = texto;
  return elemento;
}

// Detecta a linha do cabeçalho procurando "Referência"/"Código"/"Disponível"
function encontrarCabecalho(rawData) {
  for (let i = 0; i < rawData.length; i++) {
    const linha = rawData[i] ?? [];
    const temColunas = linha.some((c) =>
      ["Referência", "Código", "Disponível"].includes(String(c).trim())
    );
    if (temColunas) return i;
  }
  return 1; // fallback pro comportamento antigo
}

function lerPlanilha(buffer) {
  const workBook = XLSX.read(buffer, { type: "array" });
  const sheet = workBook.Sheets[workBook.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const headerIndex = encontrarCabecalho(rawData);
  const header = rawData[headerIndex];
  const rows = rawData.slice(headerIndex + 1);

  return rows.map((row) => {
    const item = {};
    header.forEach((col, index) => {
      item[String(col).trim()] = row[index];
    });
    return {
      referencia: String(item["Referência"] ?? "").trim(),
      descricao: String(item["Descrição"] ?? "").trim(),
      disponivel: Number(String(item["Disponível"] ?? "0").replace(",", ".")),
      codigoEAN: String(item["Código"] ?? "").trim(),
      subgrupo: String(item["Sub-Grupo"] ?? "").trim(),
      grupo: String(item["Grupo"] ?? "").trim(),
    };
  });
}

// Regra única: estoque negativo com Código e Descrição preenchidos
// (Referência e Sub-Grupo são opcionais — compatível com planilhas enxutas
// que não têm essas colunas)
function filtrarProdutos(data) {
  return data.filter(
    (item) => item.disponivel < 0 && item.codigoEAN && item.descricao
  );
}

// Varre as linhas da tabela (inclusive as escondidas pela busca) e separa
// em dois grupos, preservando a ordem original das linhas.
function lerLinhasTabela() {
  const tbody = document.getElementById("tbodyOutput");
  const zerados = [];
  const encontrados = [];
  if (!tbody) return { zerados, encontrados };

  tbody.querySelectorAll(".linha-output").forEach((linha) => {
    const codigo = linha.dataset.codigo || "";
    const input = linha.querySelector(".qtd-input");
    const qtd = Number(input?.value) || 0;
    const item = { codigo, qtd };
    if (qtd > 0) encontrados.push(item);
    else zerados.push(item);
  });

  return { zerados, encontrados };
}

// Formata um grupo pro texto copiável, renumerando a partir de 1
function formatarLinhasCopia(grupo) {
  return grupo
    .map(
      (item, index) =>
        `${index + 1}. Código: ${item.codigo}  |  Quantidade: ${item.qtd}`
    )
    .join("\n");
}

function copiarGrupo(grupo, mensagemVazio) {
  if (grupo.length === 0) {
    Toastify({
      text: mensagemVazio,
      duration: 3000,
      gravity: "top",
      position: "right",
      backgroundColor: "#f59e0b",
    }).showToast();
    return;
  }

  navigator.clipboard.writeText(formatarLinhasCopia(grupo)).then(() => {
    Toastify({
      text: "Texto copiado com sucesso!",
      duration: 3000,
      gravity: "top",
      position: "right",
      backgroundColor: "#4BB543",
    }).showToast();
  });
}

// Molda um grupo já separado pro formato de planilha ({Código, Quantidade})
function montarDadosExportar(grupo) {
  return grupo.map((item) => ({
    Código: item.codigo,
    Quantidade: item.qtd,
  }));
}

function exportarParaExcel(dados, nomeArquivo) {
  const novaPlanilha = XLSX.utils.json_to_sheet(dados);
  novaPlanilha["!cols"] = [{ wpx: 150 }, { wpx: 120 }];
  const novoWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, "Resultado");
  XLSX.writeFile(novoWorkbook, nomeArquivo);
}

function exportarGrupo(grupo, nomeArquivo, mensagemVazio) {
  if (grupo.length === 0) {
    Toastify({
      text: mensagemVazio,
      duration: 3000,
      gravity: "top",
      position: "right",
      backgroundColor: "#f59e0b",
    }).showToast();
    return;
  }

  exportarParaExcel(montarDadosExportar(grupo), nomeArquivo);
}

// Limpa o resultado anterior antes de renderizar o novo
function limparResultado() {
  const secao = document.getElementById("secao");
  if (secao) secao.innerHTML = "";
}

// Filtra as linhas já renderizadas pelo código (esconde/mostra <tr>).
// Não recria a tabela — os valores digitados nas quantidades sobrevivem.
function filtrarLinhasPorBusca(termo, tbody) {
  const termoNormalizado = termo.trim().toLowerCase().replace(/\s+/g, "");
  const linhas = tbody.querySelectorAll(".linha-output");
  linhas.forEach((linha) => {
    const codigo = (linha.dataset.codigo || "").toLowerCase().replace(/\s+/g, "");
    linha.style.display = codigo.includes(termoNormalizado) ? "" : "none";
  });
}

function renderizarResultado(dados) {
  const secao = document.getElementById("secao");
  if (!secao) return;

  limparResultado();

  const produtos = filtrarProdutos(dados);
  const total = produtos.length;

  // 1) Resumo
  const resumo = criarElemento("div", { class: "resumo" });
  const resumoIncluir = criarElemento(
    "span",
    { class: "resumo-incluir" },
    "Encontrados (incluir): 0"
  );
  const resumoZerar = criarElemento(
    "span",
    { class: "resumo-zerar" },
    `Zerados: ${total}`
  );
  resumo.appendChild(resumoIncluir);
  resumo.appendChild(resumoZerar);
  secao.appendChild(resumo);

  // Recalcula os contadores a partir do estado atual dos inputs
  function atualizarResumo() {
    const linhas = tbody.querySelectorAll(".linha-output");
    let encontrados = 0;
    linhas.forEach((linha) => {
      const input = linha.querySelector(".qtd-input");
      if (Number(input.value) > 0) encontrados++;
    });
    resumoIncluir.textContent = `Encontrados (incluir): ${encontrados}`;
    resumoZerar.textContent = `Zerados: ${linhas.length - encontrados}`;
  }

  // 2) Busca
  const buscaInput = criarElemento("input", {
    id: "buscaCodigo",
    class: "busca-input",
    type: "text",
    placeholder: "Buscar código...",
  });
  secao.appendChild(buscaInput);

  // 3) Tabela
  const tabela = criarElemento("table", { class: "tabela-output" });

  const thead = criarElemento("thead");
  const trHead = criarElemento("tr");
  trHead.appendChild(criarElemento("th", null, "Código"));
  trHead.appendChild(criarElemento("th", null, "Quantidade"));
  thead.appendChild(trHead);
  tabela.appendChild(thead);

  const tbody = criarElemento("tbody", { id: "tbodyOutput" });

  produtos.forEach((item) => {
    const tr = criarElemento("tr", {
      class: "linha-output",
      "data-codigo": item.codigoEAN,
    });

    tr.appendChild(criarElemento("td", null, item.codigoEAN));

    const tdQtd = criarElemento("td");
    const inputQtd = criarElemento("input", {
      type: "number",
      class: "qtd-input",
      value: "0",
      min: "0",
      step: "1",
    });

    // 4) Realce + contador: reage a cada alteração de quantidade
    inputQtd.addEventListener("input", () => {
      const encontrado = Number(inputQtd.value) > 0;
      inputQtd.classList.toggle("editado", encontrado);
      tr.classList.toggle("editado", encontrado);
      atualizarResumo();
    });

    tdQtd.appendChild(inputQtd);
    tr.appendChild(tdQtd);
    tbody.appendChild(tr);
  });

  tabela.appendChild(tbody);
  secao.appendChild(tabela);

  // Estado inicial (tudo 0 / zerado) — garante resumo coerente na carga
  atualizarResumo();

  buscaInput.addEventListener("input", () =>
    filtrarLinhasPorBusca(buscaInput.value, tbody)
  );

  // 5) Botões de saída
  const containerBtn = criarElemento("div", { class: "container-btn" });

  const btnCopiarZerados = criarElemento(
    "button",
    { id: "btnCopiarZerados", class: "btn-copy" },
    "Copiar zerados"
  );
  btnCopiarZerados.onclick = () => {
    const { zerados } = lerLinhasTabela();
    copiarGrupo(zerados, "Nenhum produto zerado");
  };
  containerBtn.appendChild(btnCopiarZerados);

  const btnCopiarEncontrados = criarElemento(
    "button",
    { id: "btnCopiarEncontrados", class: "btn-copy" },
    "Copiar encontrados"
  );
  btnCopiarEncontrados.onclick = () => {
    const { encontrados } = lerLinhasTabela();
    copiarGrupo(encontrados, "Nenhum produto encontrado");
  };
  containerBtn.appendChild(btnCopiarEncontrados);

  const btnExportarZerados = criarElemento(
    "button",
    { id: "btnExportarZerados", class: "btn-export" },
    "Exportar zerados"
  );
  btnExportarZerados.onclick = () => {
    const { zerados } = lerLinhasTabela();
    exportarGrupo(zerados, NOME_ARQUIVO_ZERADOS, "Nenhum produto zerado");
  };
  containerBtn.appendChild(btnExportarZerados);

  const btnExportarEncontrados = criarElemento(
    "button",
    { id: "btnExportarEncontrados", class: "btn-export" },
    "Exportar encontrados"
  );
  btnExportarEncontrados.onclick = () => {
    const { encontrados } = lerLinhasTabela();
    exportarGrupo(
      encontrados,
      NOME_ARQUIVO_ENCONTRADOS,
      "Nenhum produto encontrado"
    );
  };
  containerBtn.appendChild(btnExportarEncontrados);

  secao.appendChild(containerBtn);
}

function processarArquivo(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  file
    .arrayBuffer()
    .then((buffer) => {
      const dados = lerPlanilha(buffer);
      renderizarResultado(dados);
    })
    .catch((err) => {
      console.error(err);
      Toastify({
        text: "Erro ao ler o arquivo.",
        duration: 3000,
        gravity: "top",
        position: "right",
        backgroundColor: "#e74c3c",
      }).showToast();
    });
}

function solicitarArquivo() {
  const novoInput = document.createElement("input");
  novoInput.type = "file";
  novoInput.accept = ".xlsx, .xls";
  novoInput.style.display = "none";
  novoInput.addEventListener("change", (e) => {
    processarArquivo(e);
    novoInput.remove(); // limpa o input do DOM depois do uso
  });
  document.body.appendChild(novoInput);
  novoInput.click();
}

// Guarda: só liga o evento se o botão existir
const btnProcessar = document.getElementById("btnProcessar");
if (btnProcessar)
  btnProcessar.addEventListener("click", () => solicitarArquivo());
