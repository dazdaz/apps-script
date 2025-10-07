function onOpen() {
  DocumentApp.getUi()
    .createMenu('Markdown Converter')
    .addItem('Open Converter', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Markdown to Docs');
  DocumentApp.getUi().showSidebar(html);
}

function convertMarkdownToDoc(markdownText) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    // Add a line break before new content if document isn't empty
    if (body.getNumChildren() > 0) {
      body.appendParagraph('');
    }
    
    const lines = markdownText.split('\n');
    let inCodeBlock = false;
    let codeBlockContent = [];
    
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i];
      
      // Handle code blocks
      if (line.trim().startsWith('```')) {
        if (!inCodeBlock) {
          inCodeBlock = true;
          codeBlockContent = [];
        } else {
          // End code block
          const codeText = codeBlockContent.join('\n');
          const para = body.appendParagraph(codeText);
          para.editAsText()
            .setFontFamily('Courier New')
            .setBackgroundColor('#f4f4f4');
          inCodeBlock = false;
          codeBlockContent = [];
        }
        continue;
      }
      
      if (inCodeBlock) {
        codeBlockContent.push(line);
        continue;
      }
      
      // Skip empty lines
      if (line.trim() === '') {
        body.appendParagraph('');
        continue;
      }
      
      // Headers
      if (line.startsWith('# ')) {
        const para = body.appendParagraph(line.substring(2));
        para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      } else if (line.startsWith('## ')) {
        const para = body.appendParagraph(line.substring(3));
        para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } else if (line.startsWith('### ')) {
        const para = body.appendParagraph(line.substring(4));
        para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      }
      // Bulleted lists (with various markers)
      else if (line.match(/^\s*\*\s+/)) {
        const text = line.replace(/^\s*\*\s+/, '');
        const para = body.appendListItem(text);
        para.setGlyphType(DocumentApp.GlyphType.BULLET);
        applyInlineFormatting(para);
      }
      else if (line.match(/^[\-]\s+/)) {
        const para = body.appendListItem(line.substring(2));
        para.setGlyphType(DocumentApp.GlyphType.BULLET);
        applyInlineFormatting(para);
      }
      // Ordered lists
      else if (line.match(/^\d+\.\s+/)) {
        const text = line.replace(/^\d+\.\s+/, '');
        const para = body.appendListItem(text);
        para.setGlyphType(DocumentApp.GlyphType.NUMBER);
        applyInlineFormatting(para);
      }
      // Regular paragraphs
      else {
        const para = body.appendParagraph(line);
        applyInlineFormatting(para);
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error in convertMarkdownToDoc:', error);
    throw new Error('Conversion failed: ' + error.message);
  }
}

function applyInlineFormatting(paragraph) {
  const text = paragraph.getText();
  const editText = paragraph.editAsText();
  
  // Process bold text marked with **text**
  let regex = /\*\*([^*]+)\*\*/g;
  let match;
  let offset = 0;
  
  while ((match = regex.exec(text)) !== null) {
    const startIndex = match.index - offset;
    const content = match[1];
    
    // Replace **text** with text and apply bold
    editText.deleteText(startIndex, startIndex + 1); // Remove first **
    editText.deleteText(startIndex + content.length, startIndex + content.length + 1); // Remove second **
    editText.setBold(startIndex, startIndex + content.length - 1, true);
    
    offset += 4; // Removed 4 characters (two **)
  }
  
  // Update text for further processing
  const updatedText = paragraph.getText();
  
  // Process italic text marked with *text*
  regex = /\*([^*]+)\*/g;
  offset = 0;
  
  while ((match = regex.exec(updatedText)) !== null) {
    const startIndex = match.index - offset;
    const content = match[1];
    
    // Replace *text* with text and apply italic
    editText.deleteText(startIndex, startIndex); // Remove first *
    editText.deleteText(startIndex + content.length, startIndex + content.length); // Remove second *
    editText.setItalic(startIndex, startIndex + content.length - 1, true);
    
    offset += 2; // Removed 2 characters (two *)
  }
  
  // Process code text marked with `text`
  const finalText = paragraph.getText();
  regex = /`([^`]+)`/g;
  offset = 0;
  
  while ((match = regex.exec(finalText)) !== null) {
    const startIndex = match.index - offset;
    const content = match[1];
    
    // Replace `text` with text and apply code formatting
    editText.deleteText(startIndex, startIndex); // Remove first `
    editText.deleteText(startIndex + content.length, startIndex + content.length); // Remove second `
    editText.setFontFamily(startIndex, startIndex + content.length - 1, 'Courier New');
    editText.setBackgroundColor(startIndex, startIndex + content.length - 1, '#f4f4f4');
    
    offset += 2; // Removed 2 characters (two `)
  }
}
