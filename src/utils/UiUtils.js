function includeFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}