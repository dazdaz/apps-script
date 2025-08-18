/**
 * Google Apps Script to create PDFs from linked chapter Google Docs and concatenate them into one PDF.
 *
 * This script assumes it's run from a main Google Doc that contains hyperlinks
 * to separate chapter Docs. It extracts those links, assuming they are to Google Docs,
 * generates a PDF for each, saves them in a dedicated folder, and then creates a merged PDF
 * by concatenating the chapter contents in sequence.
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
  var mainName = mainDoc.getName();
  var body = mainDoc.getBody();
  var chapterLinks = extractChapterLinks(body);
  
  var parentFolder = DriveApp.getFileById(mainDoc.getId()).getParents().next();
  var pdfFolderName = mainName + "_PDFs";
  var pdfFolder;
  
  // Get existing folder or create a new one
  var folders = parentFolder.getFoldersByName(pdfFolderName);
  if (folders.hasNext()) {
    pdfFolder = folders.next();
  } else {
    pdfFolder = parentFolder.createFolder(pdfFolderName);
  }
  
  // Create individual PDFs
  for (var i = 0; i < chapterLinks.length; i++) {
    var chapter = chapterLinks[i];
    var docId = chapter.docId;
    var name = chapter.name;
    
    try {
      var pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
      pdfBlob.setName(name + '.pdf');
      pdfFolder.createFile(pdfBlob);
      Logger.log('Created PDF: ' + name + '.pdf');
    } catch (e) {
      Logger.log('Error creating PDF for ' + name + ': ' + e.message);
    }
  }
  
  // Create merged PDF by appending chapter contents to a temporary Doc
  if (chapterLinks.length > 0) {
    var mergedDoc = DocumentApp.create(mainName + "_Merged_Temp");
    var mergedBody = mergedDoc.getBody();
    
    for (var i = 0; i < chapterLinks.length; i++) {
      var chapter = chapterLinks[i];
      var docId = chapter.docId;
      
      try {
        appendDocTo(mergedDoc, docId);
        if (i < chapterLinks.length - 1) {
          mergedBody.appendPageBreak();
        }
      } catch (e) {
        Logger.log('Error appending chapter ' + chapter.name + ': ' + e.message);
      }
    }
    
    try {
      var mergedPdfBlob = mergedDoc.getAs('application/pdf');
      mergedPdfBlob.setName(mainName + "_Merged.pdf");
      pdfFolder.createFile(mergedPdfBlob);
      Logger.log('Created merged PDF: ' + mainName + '_Merged.pdf');
      
      // Trash the temporary merged Doc
      DriveApp.getFileById(mergedDoc.getId()).setTrashed(true);
    } catch (e) {
      Logger.log('Error creating merged PDF: ' + e.message);
    }
  }
}

/**
 * Appends the body content of a source Doc to a target Doc.
 * @param {Document} targetDoc - The target Document.
 * @param {string} sourceDocId - The ID of the source Document.
 */
function appendDocTo(targetDoc, sourceDocId) {
  var sourceDoc = DocumentApp.openById(sourceDocId);
  var sourceBody = sourceDoc.getBody();
  var targetBody = targetDoc.getBody();
  
  var totalElements = sourceBody.getNumChildren();
  for (var j = 0; j < totalElements; ++j) {
    var element = sourceBody.getChild(j).copy();
    var type = element.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(element);
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(element);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(element);
    } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      targetBody.appendHorizontalRule(element);
    } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
      targetBody.appendPageBreak(element);
    } // Add more element types if needed
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
