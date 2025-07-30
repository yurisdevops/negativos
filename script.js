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
        ? formatarProdutosZerar(dadosFormatados)
        : formatarProdutosIncluir(dadosFormatados);

    const idTextarea = tipo === "zerar" ? "outputZerar" : "outputIncluir";
    const idBotao = tipo === "zerar" ? "btn-copyZerar" : "btn-copyIncluir";
    const idTitle = tipo === "zerar" ? "Zerar" : "Incluir";

    const section = document.getElementById("secao");

    let textArea = document.getElementById(idTextarea);
    if (!textArea) {
      textArea = document.createElement("textarea");
      textArea.setAttribute("class", "output");
      textArea.setAttribute("id", idTextarea);
      textArea.rows = 35;
      textArea.cols = 100;
      section.after(textArea);
    }
    textArea.value = resultado;

    let btnCopiar = document.getElementById(idBotao);
    if (!btnCopiar) {
      btnCopiar = document.createElement("button");
      btnCopiar.setAttribute("class", "btn-copy");
      btnCopiar.setAttribute("id", idBotao);
      btnCopiar.textContent = "📋 Copiar";
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
      textArea.after(btnCopiar);
    }

    let title = document.createElement("h2");
    if (title) {
      title.setAttribute("class", "title-Output");
      title.setAttribute("id", idTitle);
      title.textContent = idTitle;
      textArea.before(title);
    }

    e.target.value = "";
  });
}

document
  .getElementById("fileZerar")
  .addEventListener("change", (e) => processarArquivo(e, "zerar"));

document
  .getElementById("fileIncluir")
  .addEventListener("change", (e) => processarArquivo(e, "incluir"));

function formatarProdutosZerar(data) {
  console.log("Dados recebidos pela função:", data);
  return data
    .filter(
      (item) =>
        item.incluir <= 0 &&
        item.disponivel < 0 &&
        item.referencia &&
        item.descricao &&
        item.subgrupo
    )
    .map(
      (item, index) =>
        `${index + 1}.Produto: ${item.descricao}\n  Referência: ${
          item.referencia
        }  |  Grupo: ${item.subgrupo}\n  Código: ${
          item.codigo
        }\n  Ação: Zerar Estoque\n`
    )
    .join("\n");
}

function formatarProdutosIncluir(data) {
  console.log("Dados recebidos pela função:", data);
  return data
    .filter(
      (item) =>
        item.incluir > 0 &&
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
        }\n  Quantidade: ${item.incluir}\n  Ação: Incluir ao Estoque\n`
    )
    .join("\n");
}

function copiarResultado() {
  const texto = document.getElementsByClassId("output").value;
  navigator.clipboard.writeText(texto).then(() => {
    alert("Texto copiado com sucesso!");
  });
}
