function gerarDoc() {
  const { Document, Packer, Paragraph, TextRun } = docx;

  function campo(id) {
    return document.getElementById(id).value.trim();
  }

  const conteudo = [];

  // Primeira linha (obrigatória)
  conteudo.push(new Paragraph({
    children: [new TextRun({ text: campo("tipificacao"), bold: true })]
  }));

  // Linha extra (se existir)
  if (campo("linhaExtra")) {
    conteudo.push(new Paragraph({
      children: [new TextRun({ text: campo("linhaExtra"), bold: true })]
    }));
  }

  // Cabeçalho
  conteudo.push(new Paragraph({
    children: [new TextRun({ text: campo("cabecalho"), bold: true })]
  }));

  conteudo.push(new Paragraph(`Data/Hora: ${campo("dataHora")}`));
  conteudo.push(new Paragraph(`Local: ${campo("local")}`));

  if (campo("preso")) {
    conteudo.push(new Paragraph(`Preso: ${campo("preso")}`));
    conteudo.push(new Paragraph(`Antecedentes: ${campo("antecedentesPreso") || "Não possui"}`));
    conteudo.push(new Paragraph(`OrCrim: ${campo("orcrimPreso") || "Desconhecido"}`));
  }

  if (campo("material")) {
    conteudo.push(new Paragraph("Material apreendido:"));
    campo("material").split("\n").forEach(item => {
      conteudo.push(new Paragraph(item));
    });
  }

  conteudo.push(new Paragraph({
    text: "Resumo do fato:",
    bold: true
  }));

  conteudo.push(new Paragraph(campo("resumo")));

  const doc = new Document({
    sections: [{ children: conteudo }]
  });

  Packer.toBlob(doc).then(blob => {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "release.docx";
    link.click();
  });
}
