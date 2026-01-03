const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

/**
 * Função principal para lidar com requisições GET.
 * Serve tanto para abrir o site no Google quanto para fornecer dados ao GitHub.
 */
function doGet(e) {
  // Se a requisição vier do GitHub (contendo ?api=true), envia os dados como JSON
  if (e.parameter.api) {
    const dados = getDadosDashboard();
    return ContentService.createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Se abrir pelo link do Google Script, renderiza a página HTML
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SAP Quality Analytics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Extrai e formata os dados da planilha.
 */
function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  if (!sheet) return [];

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];
  dados.shift(); // Remove o cabeçalho
  
  return dados.map(function(linha, indice) {
    let dataFormatada = "-";
    let dataISO = ""; 
    
    // CORREÇÃO DA DATA: Formata no servidor para evitar 'undefined' no navegador
    if (linha[5]) { 
      try {
        const dataObj = (linha[5] instanceof Date) ? linha[5] : new Date(linha[5]);
        if (!isNaN(dataObj.getTime())) {
          dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          dataISO = dataObj.toISOString().split('T')[0]; 
        }
      } catch (e) { 
        dataFormatada = String(linha[5]); 
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
      parecer: linha[7] ? String(linha[7]).trim() : "Sem Parecer",
      anexosHtml: (linha[8] || linha[9]) ? "Sim" : "-"
    };
  });
}
