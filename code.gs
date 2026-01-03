const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

/**
 * Lida com requisições GET. 
 * Se houver o parâmetro ?api=true, retorna JSON. Caso contrário, renderiza o HTML.
 */
function doGet(e) {
  if (e.parameter.api) {
    const dados = getDadosDashboard();
    return ContentService.createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SAP Quality Analytics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Busca e formata os dados da planilha "ControleCds".
 */
function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  if (!sheet) return [];

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];
  dados.shift(); // Remove cabeçalho
  
  return dados.map(function(linha, indice) {
    let dataFormatada = "-";
    if (linha[5]) { 
      try {
        const dataObj = (linha[5] instanceof Date) ? linha[5] : new Date(linha[5]);
        if (!isNaN(dataObj.getTime())) {
          dataFormatada = Utilities.formatDate(dataObj, "GMT-3", "dd/MM/yyyy");
        }
      } catch (e) { dataFormatada = String(linha[5]); }
    }

    return {
      idLinha: indice + 2,
      cd: String(linha[0] || ""),
      os: String(linha[1] || ""),
      pn: String(linha[2] || ""),
      oc: String(linha[3] || ""),
      aplic: String(linha[4] || ""),
      dataParaExibir: dataFormatada,
      qtd: Number(linha[6]) || 0,
      parecer: linha[7] ? String(linha[7]).trim() : "Sem Parecer"
    };
  });
}
