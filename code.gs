const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SAP Quality Analytics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  if (!sheet) return [];

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];
  dados.shift(); // Remove cabeçalho
  
  return dados.map(function(linha, indice) {
    let dataFormatada = "-";
    let dataISO = ""; 
    
    // Tratamento da Data (Coluna F / Índice 5)
    if (linha[5]) { 
      try {
        const dataObj = (linha[5] instanceof Date) ? linha[5] : new Date(linha[5]);
        if (!isNaN(dataObj.getTime())) {
          dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          dataISO = dataObj.toISOString().split('T')[0]; 
        }
      } catch (e) { dataFormatada = linha[5].toString(); }
    }

    const linkImg = linha[8] ? linha[8].toString() : "";
    const linkPpt = linha[9] ? linha[9].toString() : "";
    let htmlLinks = "";
    if (linkImg.includes("http")) htmlLinks += '<a href="' + linkImg + '" target="_blank">📷</a> ';
    if (linkPpt.includes("http")) htmlLinks += '<a href="' + linkPpt + '" target="_blank">📄</a>';

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
      parecer: linha[7] ? String(linha[7]).trim() : "Sem Parecer",
      anexosHtml: htmlLinks || "-"
    };
  });
}
