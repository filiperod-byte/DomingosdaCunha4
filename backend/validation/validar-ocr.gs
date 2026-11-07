/**
 * TESTE ISOLADO — não altera qualquer linha do Google Sheets ou fatura original.
 * Executar apenas uma vez no editor do Apps Script e eliminar este ficheiro depois.
 * Usa o PDF da manutenção do portão (2026/52) como referência.
 */
function validarOCRFaturaDC4() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/script.external_request'
  ]);
  const id='19CVakGEdHjxnx69HbzsLAMlDTduS65gK';
  const file=DriveApp.getFileById(id);
  const pdf=file.getBlob(),media=pdf.getBytes();
  if(pdf.getContentType()!=='application/pdf')throw Error('O documento de teste não é um PDF');
  const boundary='DC4_'+Utilities.getUuid().replace(/-/g,'');
  const metadata=JSON.stringify({name:'DC4-TESTE-OCR-TEMP',
    mimeType:'application/vnd.google-apps.document'});
  const prefix='--'+boundary+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+
    metadata+'\r\n--'+boundary+'\r\nContent-Type: application/pdf\r\n\r\n';
  const suffix='\r\n--'+boundary+'--\r\n';
  const bytes=Utilities.newBlob(prefix).getBytes()
    .concat(media,Utilities.newBlob(suffix).getBytes());
  const auth={Authorization:'Bearer '+ScriptApp.getOAuthToken()};
  let tempId='';
  try {
    const upload=UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=pt&fields=id',
      {method:'post',headers:auth,contentType:'multipart/related; boundary='+boundary,
       payload:bytes,muteHttpExceptions:true});
    if(upload.getResponseCode()!==200 && upload.getResponseCode()!==201)
      throw Error('Conversão falhou. HTTP '+upload.getResponseCode()+
        ' — '+upload.getContentText().slice(0,350));
    tempId=JSON.parse(upload.getContentText()).id;
    if(!tempId)throw Error('A conversão não devolveu o ID do Google Docs');
    const exportacao=UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(tempId)+
      '/export?mimeType=text%2Fplain',
      {headers:auth,muteHttpExceptions:true});
    if(exportacao.getResponseCode()!==200)
      throw Error('Exportação falhou. HTTP '+exportacao.getResponseCode());
    const texto=exportacao.getContentText('UTF-8');
    const resultado={
      teste:'OCR Google Drive via Apps Script',
      conversao:'OK',exportacao:'OK',
      carateres: texto.length,
      numeroReconhecido:/2026\s*\/\s*52/.test(texto),
      totalReconhecido:/295[,.]20/.test(texto)
    };
    Logger.log(JSON.stringify(resultado));
    if(!resultado.numeroReconhecido||!resultado.totalReconhecido)
      Logger.log('OCR disponível, mas alguns campos precisam de revisão.');
  } finally {
    if(tempId) {
      try {DriveApp.getFileById(tempId).setTrashed(true);
        Logger.log('Documento temporário eliminado.');
      }catch(e){Logger.log('Eliminar manualmente DC4-TESTE-OCR-TEMP: '+e.message);}
    }
  }
}
