// Configurações extraídas do seu projeto original
const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";
const ID_PASTA_DRIVE = "1uCQrm_OyCz_O7QT6AhWdGlFD-HmOeYn_";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAP Quality Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Ponto de entrada para requisições externas (API)
 * Essencial para funcionamento fora do ambiente Google (ex: GitHub Pages)
 */
function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var action = request.action;
    var payload = request.data;

    if (action === 'getDados') {
      var dados = getDadosDashboard();
      return ContentService.createTextOutput(JSON.stringify(dados))
        .setMimeType(ContentService.MimeType.JSON);
    } 
    
    if (action === 'salvar') {
      var msg = salvarDados(payload);
      return ContentService.createTextOutput(JSON.stringify({ message: msg }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDadosDashboard() {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILHA);
    var sheet = ss.getSheetByName("ControleCds");
    if (!sheet) return [];

    var dados = sheet.getDataRange().getValues();
    if (dados.length <= 1) return [];

    dados.shift(); // Remove cabeçalho
    
    return dados.map(function(linha, indice) {
      var linhaReal = indice + 2; 
      var dataFormatada = "";
      
      if (linha[5]) {
        try {
          dataFormatada = (linha[5] instanceof Date) 
            ? Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "yyyy-MM-dd")
            : linha[5].toString();
        } catch (e) { dataFormatada = ""; }
      }
      
      var dataVisual = dataFormatada;
      if(dataFormatada.includes("-")) {
         var partes = dataFormatada.split("-");
         if(partes.length === 3) dataVisual = partes[2] + "/" + partes[1] + "/" + partes[0];
      }

      var linkImg = linha[8] ? linha[8].toString() : "";
      var linkPpt = linha[9] ? linha[9].toString() : "";
      var htmlLinks = "";
      if (linkImg.toLowerCase().includes("http")) htmlLinks += '<a href="' + linkImg + '" target="_blank" style="text-decoration:none; margin-right:8px; font-size:16px">📷</a>';
      if (linkPpt.toLowerCase().includes("http")) htmlLinks += '<a href="' + linkPpt + '" target="_blank" style="text-decoration:none; font-size:16px">📄</a>';

      return {
        idLinha: linhaReal,
        cd: linha[0] ? linha[0].toString() : "",
        os: linha[1] ? linha[1].toString() : "",
        pn: linha[2] ? linha[2].toString() : "",
        oc: linha[3] ? linha[3].toString() : "",
        aplic: linha[4] ? linha[4].toString() : "",
        dataInput: dataFormatada,
        dataVisual: dataVisual,
        qtd: (typeof linha[6] === 'number' ? linha[6] : 0),
        parecer: linha[7] ? linha[7].toString().trim() : "",
        linkImg: linkImg,
        linkPpt: linkPpt,
        anexosHtml: htmlLinks || "-"
      };
    });
  } catch (erro) { throw new Error(erro.message); }
}

function salvarDados(form) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = ss.getSheetByName("ControleCds");
  var folder = DriveApp.getFolderById(ID_PASTA_DRIVE);

  var urlImgFinal = form.linkImgAntigo;
  if (form.arquivoImg && form.arquivoImg.data) {
    var blobImg = Utilities.newBlob(Utilities.base64Decode(form.arquivoImg.data), form.arquivoImg.mimeType, form.arquivoImg.name);
    var arqImg = folder.createFile(blobImg);
    arqImg.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urlImgFinal = arqImg.getUrl();
  }

  var urlPptFinal = form.linkPptAntigo;
  if (form.arquivoPpt && form.arquivoPpt.data) {
    var blobPpt = Utilities.newBlob(Utilities.base64Decode(form.arquivoPpt.data), form.arquivoPpt.mimeType, form.arquivoPpt.name);
    var arqPpt = folder.createFile(blobPpt);
    arqPpt.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urlPptFinal = arqPpt.getUrl();
  }
  
  var linhaDados = [form.cd, form.os, form.pn, form.oc, form.aplic, form.data, form.qtd, form.parecer, urlImgFinal, urlPptFinal];

  if (form.idLinha) {
    sheet.getRange(parseInt(form.idLinha), 1, 1, 10).setValues([linhaDados]);
    return "Registro atualizado!";
  } else {
    sheet.appendRow(linhaDados);
    return "Registro criado!";
  }
}
