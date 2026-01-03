// Configurações Globais
const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";
const ID_PASTA_DRIVE = "1uCQrm_OyCz_O7QT6AhWdGlFD-HmOeYn_";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAP Quality Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Ponto de entrada corrigido para lidar com requisições externas via POST
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.data;
    let resultado;

    if (action === 'getDados') {
      resultado = getDadosDashboard();
    } else if (action === 'salvar') {
      resultado = { message: salvarDados(payload) };
    }
    
    // Retorno em JSON formatado corretamente para evitar erros de conexão
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  if (!sheet) return [];

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];

  dados.shift(); // Remove cabeçalho
  
  return dados.map(function(linha, indice) {
    let dataFormatada = "";
    if (linha[5] instanceof Date) {
      dataFormatada = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    
    const linkImg = linha[8] ? linha[8].toString() : "";
    const linkPpt = linha[9] ? linha[9].toString() : "";
    let htmlLinks = "";
    if (linkImg.includes("http")) htmlLinks += '<a href="' + linkImg + '" target="_blank">📷</a>';
    if (linkPpt.includes("http")) htmlLinks += '<a href="' + linkPpt + '" target="_blank">📄</a>';

    return {
      idLinha: indice + 2,
      cd: linha[0], os: linha[1], pn: linha[2], oc: linha[3],
      aplic: linha[4], dataInput: dataFormatada, 
      qtd: linha[6] || 0, parecer: linha[7] || "",
      anexosHtml: htmlLinks || "-"
    };
  });
}

function salvarDados(form) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  const linhaDados = [form.cd, form.os || "", form.pn || "", form.oc || "", form.aplic, form.data, form.qtd, form.parecer, form.linkImgAntigo || "", form.linkPptAntigo || ""];

  if (form.idLinha) {
    sheet.getRange(parseInt(form.idLinha), 1, 1, 10).setValues([linhaDados]);
    return "Registro atualizado!";
  } else {
    sheet.appendRow(linhaDados);
    return "Novo registro criado!";
  }
}
