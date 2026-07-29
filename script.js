// Config por tipo — elimina a enxurrada de ternários
const CONFIG = {
  zerar: {
    idTextarea: "outputZerar",
    idBotaoCopiar: "btn-copyZerar",
    idExportar: "btn-exportZerar",
    idTitle: "Zerar",
    titulo: "Zerar",
    nomeArquivo: "produtos_zerar.xlsx",
  },
  incluir: {
    idTextarea: "outputIncluir",
    idBotaoCopiar: "btn-copyIncluir",
    idExportar: "btn-exportIncluir",
    idTitle: "Incluir",
    titulo: "Incluir",
    nomeArquivo: "produtos_incluir.xlsx",
  },
};

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

// Detecta a linha do cabeçalho procurando "Referência"/"Código"
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
      incluir: Number(String(item["Incluir"] ?? "0").replace(",", ".")),
    };
  });
}

function filtrarProdutos(data, tipo) {
  const isZerar = tipo === "zerar";
  return data.filter((item) => {
    const base =
      item.disponivel < 0 &&
      item.referencia &&
      item.descricao &&
      item.subgrupo;
    return isZerar ? base && item.incluir <= 0 : base && item.incluir > 0;
  });
}

function formatarProdutos(data, tipo) {
  const isZerar = tipo === "zerar";
  return filtrarProdutos(data, tipo)
    .map(
      (item, index) =>
        `${index + 1}. Código: ${item.codigoEAN}  |  Quantidade: ${
          isZerar ? "0" : item.incluir
        }  |  Ação: ${isZerar ? "Zerar" : "Incluir"}`
    )
    .join("\n");
}

function montarDadosExportar(data, tipo) {
  const isZerar = tipo === "zerar";
  return filtrarProdutos(data, tipo).map((item) => ({
    Código: item.codigoEAN,
    Quantidade: isZerar ? 0 : item.incluir,
    Ação: isZerar ? "Zerar" : "Incluir",
  }));
}

function exportarParaExcel(dados, nomeArquivo) {
  const novaPlanilha = XLSX.utils.json_to_sheet(dados);
  novaPlanilha["!cols"] = [{ wpx: 150 }, { wpx: 120 }, { wpx: 120 }];
  const novoWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, "Resultado");
  XLSX.writeFile(novoWorkbook, nomeArquivo);
}

// Limpa o resultado anterior antes de renderizar o novo
function limparResultado() {
  const secao = document.getElementById("secao");
  if (secao) secao.innerHTML = "";
}

function renderizarResultado(dados, tipo) {
  const cfg = CONFIG[tipo];
  const secao = document.getElementById("secao");
  if (!secao) return;

  limparResultado();

  const resultado = formatarProdutos(dados, tipo);

  const title = criarElemento(
    "h2",
    { class: "title-Output", id: cfg.idTitle },
    cfg.titulo
  );
  secao.appendChild(title);

  const textArea = criarElemento("textarea", {
    class: "output",
    id: cfg.idTextarea,
    rows: "35",
    cols: "100",
  });
  textArea.value = resultado;
  secao.appendChild(textArea);

  const div = criarElemento("div", {
    class: "container-btn",
    id: "btn-actions",
  });
  secao.appendChild(div);

  const btnCopiar = criarElemento(
    "button",
    { class: "btn-copy", id: cfg.idBotaoCopiar },
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
    { class: "btn-export", id: cfg.idExportar },
    "⬇️ Exportar Excel"
  );
  const dadosExportar = montarDadosExportar(dados, tipo);
  btnExportar.onclick = () => exportarParaExcel(dadosExportar, cfg.nomeArquivo);
  div.appendChild(btnExportar);
}

function processarArquivo(event, tipo) {
  const file = event.target.files?.[0];
  if (!file) return;

  file
    .arrayBuffer()
    .then((buffer) => {
      const dados = lerPlanilha(buffer);
      renderizarResultado(dados, tipo);
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

function solicitarArquivo(tipo) {
  const novoInput = document.createElement("input");
  novoInput.type = "file";
  novoInput.accept = ".xlsx, .xls";
  novoInput.style.display = "none";
  novoInput.addEventListener("change", (e) => {
    processarArquivo(e, tipo);
    novoInput.remove(); // limpa o input do DOM depois do uso
  });
  document.body.appendChild(novoInput);
  novoInput.click();
}

// Guarda: só liga os eventos se os botões existirem
const btnZerar = document.getElementById("btnZerar");
const btnIncluir = document.getElementById("btnIncluir");
if (btnZerar) btnZerar.addEventListener("click", () => solicitarArquivo("zerar"));
if (btnIncluir)
  btnIncluir.addEventListener("click", () => solicitarArquivo("incluir"));