const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

function doGet(e) {
  if (e.parameter.api) {
    const dados = getDadosDashboard();
    return ContentService.createTextOutput(JSON.stringify(dados)).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('SAP Quality Analytics')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  const dados = sheet.getDataRange().getValues();
  dados.shift(); 
  
  return dados.map(function(linha, indice) {
    let dataFormatada = "-";
    if (linha[5]) {
      try {
        const d = (linha[5] instanceof Date) ? linha[5] : new Date(linha[5]);
        dataFormatada = Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy");
      } catch (e) {}
    }

    return {
      cd: String(linha[0] || ""),
      os: String(linha[1] || ""),
      pn: String(linha[2] || ""),
      oc: String(linha[3] || ""),
      aplic: String(linha[4] || ""),
      data: dataFormatada,
      qtd: Number(linha[6]) || 0,
      parecer: String(linha[7] || "Sem Parecer"),
      anexo1: String(linha[8] || ""), // Coluna I
      anexo2: String(linha[9] || "")  // Coluna J
    };
  });
}

// FUNÇÃO PARA O FORMULÁRIO SALVAR NA PLANILHA
function salvarDados(obj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("ControleCds");
    sheet.appendRow([
      obj.cd, obj.os, obj.pn, obj.oc, obj.aplic, 
      new Date(obj.data + "T12:00:00"), // Ajuste de fuso
      obj.qtd, obj.parecer, obj.anexo1, obj.anexo2
    ]);
    return "Sucesso!";
  } catch (e) {
    return "Erro: " + e.toString();
  }
}
