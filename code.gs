const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

function doGet(e) {
  // Se o GitHub pedir dados (?api=true), envia JSON. Se abrir no Google, abre o HTML.
  if (e.parameter.api) {
    return ContentService.createTextOutput(JSON.stringify(getDadosDashboard()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SAP Quality Analytics')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  const dados = sheet.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho
  
  return dados.map((linha, indice) => {
    let dataFormatada = "-";
    let dataISO = "";
    if (linha[5]) {
      const dataObj = (linha[5] instanceof Date) ? linha[5] : new Date(linha[5]);
      if (!isNaN(dataObj.getTime())) {
        dataFormatada = Utilities.formatDate(dataObj, "GMT-3", "dd/MM/yyyy");
        dataISO = dataObj.toISOString().split('T')[0];
      }
    }
    return {
      idLinha: indice + 2,
      cd: String(linha[0] || ""),
      os: String(linha[1] || ""),
      pn: String(linha[2] || ""),
      oc: String(linha[3] || ""),
      aplic: String(linha[4] || ""),
      dataParaExibir: dataFormatada,
      dataISO: dataISO,
      qtd: Number(linha[6]) || 0,
      parecer: linha[7] ? String(linha[7]).trim() : "Sem Parecer"
    };
  });
}
