/**
 * Ponto de entrada do Google Apps Script Web App
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Finance Intelligence')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Função para incluir outros arquivos HTML dentro do HTML principal
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
