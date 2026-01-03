const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

function doGet(e) {
  try {
    const dados = getDadosDashboard();
    return ContentService.createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "erro", "msg": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("ControleCds");
    
    let url1 = data.anexo1Base64 ? uploadParaDrive(data.anexo1Base64, data.anexo1Nome) : "";
    let url2 = data.anexo2Base64 ? uploadParaDrive(data.anexo2Base64, data.anexo2Nome) : "";

    sheet.appendRow([
      data.cd, data.os, data.pn, data.oc, data.aplic, 
      new Date(data.data + "T12:00:00"), 
      data.qtd, data.parecer, url1, url2
    ]);

    return ContentService.createTextOutput(JSON.stringify({ "result": "Sucesso!" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "Erro", "error": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  const dados = sheet.getDataRange().getValues();
  dados.shift(); 

  return dados.map(linha => ({
    cd: String(linha[0] || ""),
    os: String(linha[1] || ""),
    pn: String(linha[2] || ""),
    oc: String(linha[3] || ""),
    aplic: String(linha[4] || ""),
    data: linha[5] instanceof Date ? Utilities.formatDate(linha[5], "GMT-3", "dd/MM/yyyy") : String(linha[5] || "-"),
    qtd: Number(linha[6]) || 0,
    parecer: String(linha[7] || "Sem Parecer"),
    anexo1: String(linha[8] || ""),
    anexo2: String(linha[9] || "")
  }));
}

function uploadParaDrive(base64Data, fileName) {
  const contentType = base64Data.substring(5, base64Data.indexOf(';'));
  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = DriveApp.getRootFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
