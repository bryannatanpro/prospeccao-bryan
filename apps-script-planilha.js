// Atualize o código no projeto Apps Script JÁ vinculado à planilha.
// Execute configurarColunaOrigem uma vez para criar a coluna imediatamente.
// Depois atualize a implantação existente para uma nova versão, mantendo a URL.
// Republique também os dois HTMLs para que enviem a origem.

function prepararColunaOrigem(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Data', 'Hora', 'Seção', 'ID da Mensagem', 'Texto da Mensagem']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  const cabecalhos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let coluna = cabecalhos.indexOf('Origem da cópia') + 1;
  if (!coluna) {
    coluna = Math.max(6, sheet.getLastColumn() + 1);
    if (coluna > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(1, coluna).setValue('Origem da cópia').setFontWeight('bold');
  }
  return coluna;
}

function configurarColunaOrigem() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    prepararColunaOrigem(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse(e.postData.contents);
    lock.waitLock(30000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const coluna = prepararColunaOrigem(sheet);
    // Versões antigas não identificam o perfil; não atribuir por suposição.
    const origem = ['Vitória', 'Bryan'].includes(data.origem) ? data.origem : 'Não informada';
    const linha = Array(coluna).fill('');
    [data.data, data.hora, data.secao, data.id, data.texto].forEach((valor, i) => {
      linha[i] = valor == null ? '' : valor;
    });
    linha[coluna - 1] = origem;
    sheet.appendRow(linha);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
