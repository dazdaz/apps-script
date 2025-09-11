/**
 * This script extracts speaker notes from a Google Slides presentation,
 * creates a new Google Doc with the formatted notes, and then converts
 * and saves that document as a PDF in your Google Drive. Each note entry
 * includes a link back to the original slide.
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
      // Corrected line to properly access the text from the speaker notes shape.
      const speakerNotesShape = slide.getNotesPage().getSpeakerNotesShape();
      let slideNotes = '';

      // Check if the shape exists before trying to get its text.
      if (speakerNotesShape) {
        slideNotes = speakerNotesShape.getText().asString();
      }

      const slideNumber = index + 1;
      const slideUrl = `https://docs.google.com/presentation/d/${presentation.getId()}/edit#slide=id.${slide.getObjectId()}`;

      // If the slide has notes, add them to the new document.
      if (slideNotes && slideNotes.trim() !== "") {
        body.appendParagraph(`--- Slide ${slideNumber} ---`);
        
        // Add a link to the slide.
        const linkParagraph = body.appendParagraph('Go to slide');
        const linkStyle = {};
        linkStyle[DocumentApp.Attribute.LINK_URL] = slideUrl;
        linkParagraph.setAttributes(linkStyle);
        
        // Add the notes content.
        body.appendParagraph(slideNotes).setSpacingAfter(10);
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
