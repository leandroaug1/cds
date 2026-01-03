// Configurações do Projeto
const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";
const ID_PASTA_DRIVE = "1uCQrm_OyCz_O7QT6AhWdGlFD-HmOeYn_";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAP Quality Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Ponto de entrada da API para o GitHub
function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var action = request.action;
    var payload = request.data;
    var resultado;

    if (action === 'getDados') {
      resultado = getDadosDashboard();
    } else if (action === 'salvar') {
      resultado = { message: salvarDados(payload) };
    }
    
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDadosDashboard() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = ss.getSheetByName("ControleCds");
  if (!sheet) return [];

  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];

  dados.shift(); // Remove cabeçalho
  
  return dados.map(function(linha, indice) {
    var dataVisual = "";
    if (linha[5]) {
      try {
        // Formatação rigorosa da data para evitar "undefined"
        dataVisual = Utilities.formatDate(new Date(linha[5]), Session.getScriptTimeZone(), "dd/MM/yyyy");
      } catch (e) { dataVisual = linha[5].toString(); }
    }

    var linkImg = linha[8] ? linha[8].toString() : "";
    var linkPpt = linha[9] ? linha[9].toString() : "";
    var htmlLinks = "";
    if (linkImg.includes("http")) htmlLinks += '<a href="' + linkImg + '" target="_blank">📷</a>';
    if (linkPpt.includes("http")) htmlLinks += '<a href="' + linkPpt + '" target="_blank">📄</a>';

    return {
      idLinha: indice + 2,
      cd: linha[0] || "",
      os: linha[1] || "",
      pn: linha[2] || "",
      oc: linha[3] || "",
      aplic: linha[4] || "",
      dataInput: linha[5], 
      dataVisual: dataVisual, // Campo usado na tabela
      qtd: Number(linha[6]) || 0,
      parecer: linha[7] ? linha[7].toString().trim() : "",
      anexosHtml: htmlLinks || "-"
    };
  });
}

function salvarDados(form) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = ss.getSheetByName("ControleCds");
  var linhaDados = [form.cd, form.os, form.pn, form.oc, form.aplic, form.data, form.qtd, form.parecer, form.linkImgAntigo || "", form.linkPptAntigo || ""];

  if (form.idLinha) {
    sheet.getRange(parseInt(form.idLinha), 1, 1, 10).setValues([linhaDados]);
    return "Registro atualizado!";
  } else {
    sheet.appendRow(linhaDados);
    return "Novo registro criado!";
  }
}
