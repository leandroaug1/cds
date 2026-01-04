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
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("ControleCds");
    const data = JSON.parse(e.postData.contents);
    
    SpreadsheetApp.flush();
    const rows = sheet.getDataRange().getValues();

    if (data.action === "ADD") {
      sheet.appendRow([
        data.cd, data.os, data.pn, data.oc, data.aplic, 
        new Date(data.data + "T12:00:00"), 
        data.qtd, data.parecer, 
        data.anexo1Base64 ? uploadParaDrive(data.anexo1Base64, data.anexo1Nome) : "",
        data.anexo2Base64 ? uploadParaDrive(data.anexo2Base64, data.anexo2Nome) : ""
      ]);
    } 
    else if (data.action === "EDIT") {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0].toString().trim() === data.idOriginal.toString().trim()) {
          let url1 = data.anexo1Base64 ? uploadParaDrive(data.anexo1Base64, data.anexo1Nome) : (data.anexo1Existente || "");
          let url2 = data.anexo2Base64 ? uploadParaDrive(data.anexo2Base64, data.anexo2Nome) : (data.anexo2Existente || "");
          sheet.getRange(i + 1, 1, 1, 10).setValues([[
            data.cd, data.os, data.pn, data.oc, data.aplic, 
            new Date(data.data + "T12:00:00"), data.qtd, data.parecer, url1, url2
          ]]);
          break;
        }
      }
    } 
    else if (data.action === "DELETE") {
      // Percorre de baixo para cima para garantir a integridade dos índices ao remover
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]).trim() === String(data.id).trim()) {
          sheet.deleteRow(i + 1); // Remove a linha física
          SpreadsheetApp.flush(); // Força a atualização no banco de dados
          break;
        }
      }
    }

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
  const values = sheet.getDataRange().getValues();
  const dados = values.filter((linha, index) => {
    return index > 0 && linha[0] !== "" && linha[0] !== null;
  });
  return dados.map(linha => {
    let d = linha[5];
    return {
      id: linha[0],
      cd: String(linha[0] || ""),
      os: String(linha[1] || ""),
      pn: String(linha[2] || ""),
      oc: String(linha[3] || ""),
      aplic: String(linha[4] || ""),
      dataRaw: d instanceof Date ? d.toISOString().split('T')[0] : "", 
      dataExibicao: d instanceof Date ? Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy") : String(d || "-"),
      qtd: Number(linha[6]) || 0,
      parecer: String(linha[7] || "Sem Parecer"),
      anexo1: String(linha[8] || ""),
      anexo2: String(linha[9] || "")
    };
  });
}

function uploadParaDrive(base64Data, fileName) {
  const contentType = base64Data.substring(5, base64Data.indexOf(';'));
  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = DriveApp.getRootFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
