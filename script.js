function criarElemento(tipo, atributos, texto) {
  const elemento = document.createElement(tipo);
  if (atributos) {
    Object.keys(atributos).forEach((chave) => {
      elemento.setAttribute(chave, atributos[chave]);
    });
  }
  if (texto) {
    elemento.textContent = texto;
  }
  return elemento;
}

function processarArquivo(event, tipo) {
  const file = event.target.files?.[0];
  if (!file) return;

  file.arrayBuffer().then((buffer) => {
    const workBook = XLSX.read(buffer, { type: "buffer" });
    const sheet = workBook.Sheets[workBook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const header = rawData[1];
    const rows = rawData.slice(2);

    const dadosFormatados = rows.map((row) => {
      const item = {};
      header.forEach((col, index) => {
        item[col] = row[index];
      });
      return {
        referencia: String(item["Referência"] ?? "").trim(),
        descricao: String(item["Descrição"] ?? "").trim(),
        preco: parseFloat(String(item["Venda"] ?? "0").replace(",", ".")),
        disponivel: Number(String(item["Disponível"] ?? "0").replace(",", ".")),
        codigo: String(item["Código"] ?? "").trim(),
        subgrupo: String(item["Sub-Grupo"] ?? "").trim(),
        grupo: String(item["Grupo"] ?? "").trim(),
        incluir: Number(String(item["Incluir"] ?? "0").replace(",", ".")),
      };
    });

    const resultado =
      tipo === "zerar"
        ? formatarProdutos(dadosFormatados, "zerar")
        : formatarProdutos(dadosFormatados, "incluir");

    const idTextarea = tipo === "zerar" ? "outputZerar" : "outputIncluir";
    const idBotao = tipo === "zerar" ? "btn-copyZerar" : "btn-copyIncluir";
    const idTitle = tipo === "zerar" ? "Zerar" : "Incluir";
    const idOcultar = tipo === "zerar" ? "outputIncluir" : "outputZerar";
    const botaoOcultar = tipo === "zerar" ? "btn-copyIncluir" : "btn-copyZerar";
    const tituloOcultar = tipo === "zerar" ? "Incluir" : "Zerar";
    const exportOcultar =
      tipo === "zerar" ? "btn-exportIncluir" : "btn-exportZerar";

    // Remover elementos que não são necessários
    document.getElementById(idOcultar)?.remove();
    document.getElementById(botaoOcultar)?.remove();
    document.getElementById(tituloOcultar)?.remove();
    document.getElementById(exportOcultar)?.remove();

    const section = document.getElementById("secao");

    let textArea = document.getElementById(idTextarea);
    if (!textArea) {
      textArea = criarElemento("textarea", {
        class: "output",
        id: idTextarea,
        rows: "35",
        cols: "100",
      });
      section.after(textArea);
    }
    textArea.value = resultado;

    let div = criarElemento("div", {
      class: "container-btn",
      id: "btn-actions",
    });

    textArea.after(div);

    let btnCopiar = document.getElementById(idBotao);
    if (!btnCopiar) {
      btnCopiar = criarElemento(
        "button",
        {
          class: "btn-copy",
          id: idBotao,
        },
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
    }

    let title = document.createElement("h2");
    if (title) {
      title.setAttribute("class", "title-Output");
      title.setAttribute("id", idTitle);
      title.textContent = idTitle;
      textArea.before(title);
    }

    const nomeArquivo =
      tipo === "zerar" ? "produtos_zerar.xlsx" : "produtos_incluir.xlsx";

    const dadosExportar =
      tipo === "zerar"
        ? dadosFormatados
            .filter(
              (item) =>
                item.incluir <= 0 &&
                item.disponivel < 0 &&
                item.referencia &&
                item.descricao &&
                item.subgrupo
            )
            .map(({ incluir, grupo, ...item }) => ({ ...item, Zerar: "sim" }))
        : dadosFormatados
            .filter(
              (item) =>
                item.incluir > 0 &&
                item.disponivel < 0 &&
                item.referencia &&
                item.descricao &&
                item.subgrupo &&
                item.incluir
            )
            .map(({ grupo, ...item }) => ({ ...item }));

    const exportarParaExcel = (dados, nomeArquivo) => {
      const novaPlanilha = XLSX.utils.json_to_sheet(dados);
      const estiloAlinhamentoDireita = {
        alignment: { horizontal: "right" },
        font: { bold: true },
      };

      // Define a largura das colunas. Ajuste os valores conforme a necessidade.
      novaPlanilha["!cols"] = [
        { wpx: 120 },
        { wpx: 250 },
        { wpx: 150 },
        { wpx: 150 },
        { wpx: 150 },
        { wpx: 150 },
        { wpx: 150 },
      ];
      if (novaPlanilha["!ref"]) {
        const range = XLSX.utils.decode_range(novaPlanilha["!ref"]);
        for (let R = range.s.r; R <= range.e.r; ++R) {
          // Loop para todas as linhas
          for (let C = range.s.c; C <= range.e.c; ++C) {
            // Loop para todas as colunas
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (novaPlanilha[cellAddress]) {
              // Aplica o alinhamento, mantendo outros estilos se existirem
              if (!novaPlanilha[cellAddress].s) {
                novaPlanilha[cellAddress].s = {};
              }
              Object.assign(
                novaPlanilha[cellAddress].s,
                estiloAlinhamentoDireita
              );
            }
          }
        }
      }

      const novoWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, "Resultado");
      XLSX.writeFile(novoWorkbook, nomeArquivo);
    };

    let btnExportar = document.createElement("button");
    const idExportar =
      tipo === "zerar" ? "btn-exportZerar" : "btn-exportIncluir";
    btnExportar.setAttribute("id", idExportar);
    btnExportar.setAttribute("class", "btn-export");
    btnExportar.textContent = "⬇️ Exportar Excel";
    btnExportar.onclick = () => exportarParaExcel(dadosExportar, nomeArquivo);
    div.appendChild(btnExportar);

    const input = event.target;
    const clone = input.cloneNode(true);
    input.replaceWith(clone);

    clone.addEventListener("change", (e) => processarArquivo(e, tipo));
  });
}

function formatarProdutos(data, tipo) {
  const isZerar = tipo === "zerar";
  return data
    .filter((item) =>
      isZerar
        ? item.incluir <= 0 &&
          item.disponivel < 0 &&
          item.referencia &&
          item.descricao &&
          item.subgrupo
        : item.incluir > 0 &&
          item.disponivel < 0 &&
          item.referencia &&
          item.descricao &&
          item.subgrupo &&
          item.incluir
    )
    .map(
      (item, index) =>
        `${index + 1}.Produto: ${item.descricao}\n  Referência: ${
          item.referencia
        }  |  Grupo: ${item.subgrupo}\n  Código: ${
          item.codigo
        }\n  Quantidade: ${isZerar ? "" : item.incluir}\n  Ação: ${
          isZerar ? "Zerar Estoque" : "Incluir ao Estoque"
        }\n`
    )
    .join("\n");
}

document
  .getElementById("btnZerar")
  .addEventListener("click", () => solicitarArquivo("zerar"));
document
  .getElementById("btnIncluir")
  .addEventListener("click", () => solicitarArquivo("incluir"));

function solicitarArquivo(tipo) {
  const novoInput = document.createElement("input");
  novoInput.type = "file";
  novoInput.accept = ".xlsx, .xls";
  novoInput.style.display = "none";

  novoInput.addEventListener("change", (e) => processarArquivo(e, tipo));

  document.body.appendChild(novoInput);
  novoInput.click();
}
