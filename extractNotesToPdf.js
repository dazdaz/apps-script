/**
 * This script extracts speaker notes from a Google Slides presentation,
 * creates a new Google Doc with the formatted notes, and then converts
 * and saves that document as a PDF in your Google Drive. Each note entry
 * includes a link back to the original slide.
 * Include the slide title (up to 55 chars) on the same line as the page number.
 * All text formatting (bold, italic, underline) is removed - only plain text is copied.
 */
function extractNotesToPdf() {
  try {
    // Get the active Google Slides presentation.
    const presentation = SlidesApp.getActivePresentation();
    const presentationTitle = presentation.getName();
    const slides = presentation.getSlides();

    // Create a new Google Doc to hold the notes.
    const doc = DocumentApp.create(`${presentationTitle} - Extracted Notes`);
    const body = doc.getBody();

    // Iterate through each slide and extract the notes.
    slides.forEach((slide, index) => {
      
      let slideTitle = 'No Title';
      
      // 1. Try to get the text from the Title placeholder
      const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
      if (titlePlaceholder && titlePlaceholder.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        slideTitle = titlePlaceholder.asShape().getText().asString();
      } 
      
      // 2. If no title found, look for text with font size > 21
      if (slideTitle === 'No Title' || slideTitle.trim() === '') {
        const largestText = findLargestTextOnSlide(slide);
        if (largestText) {
          slideTitle = largestText;
        }
      }
      
      // 3. Fallback to the Slide Layout name if the above is empty
      if (slideTitle === 'No Title' || slideTitle.trim() === '') {
        try {
          slideTitle = slide.getLayout().getName();
        } catch (e) {
          slideTitle = 'Untitled Slide'; // Final safe fallback
        }
      }

      // Truncate the title to 55 characters and replace newlines with a space
      const truncatedTitle = slideTitle.replace(/[\n\r]/g, ' ').substring(0, 55).trim();
      
      // Get the speaker notes content.
      const speakerNotesShape = slide.getNotesPage().getSpeakerNotesShape();
      let slideNotesText = '';

      // Check if the shape exists before trying to get its text.
      if (speakerNotesShape) {
        // Get only the plain text string
        slideNotesText = speakerNotesShape.getText().asString();
      }

      const slideNumber = index + 1;
      const slideUrl = `https://docs.google.com/presentation/d/${presentation.getId()}/edit#slide=id.${slide.getObjectId()}`;

      // If the slide has notes, add them to the new document.
      if (slideNotesText && slideNotesText.trim() !== "") {
        
        // --- Slide {Number} - {Title} --- (Header)
        const headerParagraph = body.appendParagraph(`--- Slide ${slideNumber} - ${truncatedTitle} ---`);
        headerParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);

        // Add a link to the slide.
        const linkParagraph = body.appendParagraph('Go to slide');
        const linkStyle = {};
        linkStyle[DocumentApp.Attribute.LINK_URL] = slideUrl;
        linkParagraph.setAttributes(linkStyle);

        // Add the notes content as plain text
        body.appendParagraph('').setSpacingAfter(0).setSpacingBefore(0);

        // Split the notes by line breaks to handle paragraphs
        const notesLines = slideNotesText.split('\n');
        
        notesLines.forEach((line) => {
          const newParagraph = body.appendParagraph('');
          
          // Handle bullet points (simple detection)
          if (line.trim().startsWith('•') || line.trim().startsWith('-')) {
            const glyphType = DocumentApp.GlyphType.BULLET;
            newParagraph.setGlyphType(glyphType);
          }
          
          // Add the line as plain text without any formatting
          if (line.length > 0) {
            newParagraph.appendText(line);
            // No formatting applied - just plain text
          }
        });
        
        // Add a blank line/spacing after the notes for separation.
        body.appendParagraph('').setSpacingAfter(20);
      }
    });

    // Close the document to save changes and get its ID.
    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());

    // Save the new Google Doc as a PDF file in Google Drive.
    const pdfBlob = docFile.getAs(MimeType.PDF);
    pdfBlob.setName(`${presentationTitle} - Speaker Notes.pdf`);

    // Get the root folder of Drive and save the PDF there.
    DriveApp.createFile(pdfBlob);

    // Log success message to the console.
    Logger.log('PDF created successfully in your Google Drive.');

  } catch (error) {
    Logger.log(`An error occurred: ${error.message}`);
  }
}

/**
 * Helper function to find text with font size > 21 on a slide
 * Returns the text with the largest font size (prioritizing the first found)
 */
function findLargestTextOnSlide(slide) {
  let largestText = null;
  let largestFontSize = 21; // Minimum threshold
  
  try {
    // Get all shapes on the slide
    const shapes = slide.getShapes();
    
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      
      // Check if the shape has text
      if (shape.getText()) {
        const textRange = shape.getText();
        const text = textRange.asString().trim();
        
        // Skip empty text
        if (text === '') continue;
        
        // Get the runs to check font sizes
        const runs = textRange.getRuns();
        
        // Check the font size of the first run (assuming title would be consistent)
        if (runs.length > 0) {
          const fontSize = runs[0].getTextStyle().getFontSize();
          
          // If fontSize is specified and larger than our threshold
          if (fontSize !== null && fontSize > largestFontSize) {
            largestFontSize = fontSize;
            largestText = text;
          }
        }
      }
    }
    
    // Also check tables for large text
    const tables = slide.getTables();
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          const cell = table.getCell(row, col);
          const textRange = cell.getText();
          const text = textRange.asString().trim();
          
          if (text === '') continue;
          
          const runs = textRange.getRuns();
          if (runs.length > 0) {
            const fontSize = runs[0].getTextStyle().getFontSize();
            
            if (fontSize !== null && fontSize > largestFontSize) {
              largestFontSize = fontSize;
              largestText = text;
            }
          }
        }
      }
    }
    
  } catch (e) {
    Logger.log(`Error finding large text on slide: ${e.message}`);
  }
  
  return largestText;
}
