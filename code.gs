const ID_PLANILHA = "1rU7ETLF7vxQY3mQNFjVSpVmWts6lcZltzb22GQWy9sQ";

/**
 * Rota GET: Retorna os dados para o Dashboard ou processa salvamento simples
 */
function doGet(e) {
  // Se houver parâmetro 'dados', estamos tentando salvar
  if (e.parameter.action === 'salvar') {
    return salvarDadosExterno(e.parameter);
  }

  // Caso contrário, retorna os dados para o Dashboard
  const dados = getDadosDashboard();
  return ContentService.createTextOutput(JSON.stringify(dados))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Rota POST: Para receber arquivos (Base64) e objetos grandes
 */
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const resultado = salvarDados(params);
  return ContentService.createTextOutput(JSON.stringify({result: resultado}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName("ControleCds");
  const dados = sheet.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho

  return dados.map(linha => ({
    cd: String(linha[0] || ""),
    os: String(linha[1] || ""),
    pn: String(linha[2] || ""),
    oc: String(linha[3] || ""),
    aplic: String(linha[4] || ""),
    data: linha[5] ? Utilities.formatDate(new Date(linha[5]), "GMT-3", "dd/MM/yyyy") : "-",
    qtd: Number(linha[6]) || 0,
    parecer: String(linha[7] || "Sem Parecer"),
    anexo1: String(linha[8] || ""),
    anexo2: String(linha[9] || "")
  }));
}

function uploadParaDrive(base64Data, fileName) {
  try {
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = DriveApp.getRootFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "";
  }
}

function salvarDados(obj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("ControleCds");
    sheet.appendRow([
      obj.cd, obj.os, obj.pn, obj.oc, obj.aplic, 
      new Date(obj.data + "T12:00:00"), 
      obj.qtd, obj.parecer, obj.anexo1, obj.anexo2
    ]);
    return "Sucesso!";
  } catch (e) {
    return "Erro: " + e.toString();
  }
}
