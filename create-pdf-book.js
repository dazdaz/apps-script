/**
 * Google Apps Script to create PDFs from linked chapter Google Docs.
 * 
 * This script assumes it's run from a main Google Doc that contains hyperlinks
 * to separate chapter Docs. It extracts those links, assuming they are to Google Docs,
 * and generates a PDF for each, saving them in the same folder as the main Doc.
 * 
 * To use: 
 * 1. Open the main Doc in Google Docs.
 * 2. Go to Extensions > Apps Script.
 * 3. Paste this code into the script editor.
 * 4. Save and run the function `createPDFsFromChapterLinks`.
 * 
 * Note: Grant permissions when prompted. Handles basic error logging.
 */

function createPDFsFromChapterLinks() {
  var mainDoc = DocumentApp.getActiveDocument();
  var body = mainDoc.getBody();
  var chapterLinks = extractChapterLinks(body);
  
  var folder = DriveApp.getFileById(mainDoc.getId()).getParents().next();
  
  for (var i = 0; i < chapterLinks.length; i++) {
    var chapter = chapterLinks[i];
    var docId = chapter.docId;
    var name = chapter.name;
    
    try {
      var pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
      pdfBlob.setName(name + '.pdf');
      folder.createFile(pdfBlob);
      Logger.log('Created PDF: ' + name + '.pdf');
    } catch (e) {
      Logger.log('Error creating PDF for ' + name + ': ' + e.message);
    }
  }
}

/**
 * Extracts hyperlinks to Google Docs from the body, along with their display text.
 * @param {Body} body - The document body.
 * @return {Array} Array of objects {docId: string, name: string}.
 */
function extractChapterLinks(body) {
  var links = [];
  var paragraphs = body.getParagraphs();
  
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var textElement = p.editAsText();
    var text = textElement.getText();
    var offset = 0;
    
    while (offset < text.length) {
      var url = textElement.getLinkUrl(offset);
      if (url) {
        var start = offset;
        while (offset < text.length && textElement.getLinkUrl(offset) === url) {
          offset++;
        }
        var linkText = text.substring(start, offset).trim();
        var match = url.match(/https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)\/?.*/);
        if (match) {
          var docId = match[1];
          var chapterName = linkText || 'Chapter_' + (links.length + 1);
          links.push({docId: docId, name: chapterName});
        }
      } else {
        offset++;
      }
    }
  }
  
  // Remove duplicates by docId (last occurrence wins)
  var uniqueLinks = {};
  for (var j = 0; j < links.length; j++) {
    uniqueLinks[links[j].docId] = links[j];
  }
  
  return Object.values(uniqueLinks);
}
