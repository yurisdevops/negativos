// Nome do arquivo gerado na exportação
const NOME_ARQUIVO_EXPORT = "produtos_negativos.xlsx";

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

function formatarProdutos(data) {
  return filtrarProdutos(data)
    .map(
      (item, index) =>
        `${index + 1}. Código: ${item.codigoEAN}  |  Quantidade: 0`
    )
    .join("\n");
}

function montarDadosExportar(data) {
  return filtrarProdutos(data).map((item) => ({
    Código: item.codigoEAN,
    Quantidade: 0,
  }));
}

function exportarParaExcel(dados, nomeArquivo) {
  const novaPlanilha = XLSX.utils.json_to_sheet(dados);
  novaPlanilha["!cols"] = [{ wpx: 150 }, { wpx: 120 }];
  const novoWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, "Resultado");
  XLSX.writeFile(novoWorkbook, nomeArquivo);
}

// Limpa o resultado anterior antes de renderizar o novo
function limparResultado() {
  const secao = document.getElementById("secao");
  if (secao) secao.innerHTML = "";
}

function renderizarResultado(dados) {
  const secao = document.getElementById("secao");
  if (!secao) return;

  limparResultado();

  const resultado = formatarProdutos(dados);

  const title = criarElemento(
    "h2",
    { class: "title-Output" },
    "Produtos Negativos"
  );
  secao.appendChild(title);

  const textArea = criarElemento("textarea", {
    class: "output",
    id: "output",
    rows: "35",
    cols: "100",
  });
  textArea.value = resultado;
  secao.appendChild(textArea);

  const div = criarElemento("div", { class: "container-btn" });
  secao.appendChild(div);

  const btnCopiar = criarElemento(
    "button",
    { class: "btn-copy" },
    "📋 Copiar"
  );
  btnCopiar.onclick = () => {
    navigator.clipboard.writeText(textArea.value).then(() => {
      Toastify({
        text: "Texto copiado com sucesso!",
        duration: 3000,
        gravity: "top",
        position: "right",
        backgroundColor: "#4BB543",
      }).showToast();
    });
  };
  div.appendChild(btnCopiar);

  const btnExportar = criarElemento(
    "button",
    { class: "btn-export" },
    "⬇️ Exportar Excel"
  );
  const dadosExportar = montarDadosExportar(dados);
  btnExportar.onclick = () =>
    exportarParaExcel(dadosExportar, NOME_ARQUIVO_EXPORT);
  div.appendChild(btnExportar);
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
