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
    const rows = sheet.getDataRange().getValues();

    if (data.action === "ADD") {
      sheet.appendRow([
        data.cd, data.os, data.pn, data.oc, data.aplic, 
        new Date(data.data + "T12:00:00"), 
        data.qtd, data.parecer, "", ""
      ]);
    } 
    else if (data.action === "EDIT") {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.idOriginal) { // Localiza pela chave original (CD)
          sheet.getRange(i + 1, 1, 1, 8).setValues([[
            data.cd, data.os, data.pn, data.oc, data.aplic, 
            new Date(data.data + "T12:00:00"), data.qtd, data.parecer
          ]]);
          break;
        }
      }
    } 
    else if (data.action === "DELETE") {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.id) {
          sheet.deleteRow(i + 1);
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
  const dados = sheet.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho

  return dados.map(linha => {
    let d = linha[5];
    return {
      id: linha[0], // Usamos o CD como ID único
      cd: String(linha[0] || ""),
      os: String(linha[1] || ""),
      pn: String(linha[2] || ""),
      oc: String(linha[3] || ""),
      aplic: String(linha[4] || ""),
      dataRaw: d instanceof Date ? d.toISOString().split('T')[0] : "", 
      dataExibicao: d instanceof Date ? Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy") : String(d || "-"),
      qtd: Number(linha[6]) || 0,
      parecer: String(linha[7] || "Sem Parecer")
    };
  });
}
