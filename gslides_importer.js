/**
 * Google Slides Content Importer
 * Parses formatted content and imports slides with speaker notes
 */

// Add menu when the presentation opens
function onOpen() {
  SlidesApp.getUi()
    .createMenu('Content Importer')
    .addItem('Import Slides Content', 'showImportDialog')
    .addToUi();
}

// Show the import dialog
function showImportDialog() {
  const html = HtmlService.createHtmlOutput(getDialogHtml())
    .setWidth(700)
    .setHeight(600)
    .setTitle('Import Slides Content');
  
  SlidesApp.getUi().showModalDialog(html, 'Import Slides Content');
}

// HTML for the dialog
function getDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * {
      box-sizing: border-box;
      font-family: 'Google Sans', Arial, sans-serif;
    }
    body {
      margin: 0;
      padding: 20px;
      background: #f8f9fa;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
    }
    .container {
      background: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.12);
    }
    label {
      font-weight: 500;
      color: #3c4043;
      display: block;
      margin-bottom: 8px;
    }
    textarea {
      width: 100%;
      height: 350px;
      padding: 12px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-family: 'Roboto Mono', monospace;
      font-size: 12px;
      resize: vertical;
    }
    textarea:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px rgba(26,115,232,0.2);
    }
    .button-row {
      margin-top: 16px;
      display: flex;
      gap: 12px;
      justify-content: flex-end;
    }
    button {
      padding: 10px 24px;
      border: none;
      border-radius: 4px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.2s;
    }
    .primary {
      background: #1a73e8;
      color: white;
    }
    .primary:hover {
      background: #1557b0;
    }
    .secondary {
      background: #f1f3f4;
      color: #3c4043;
    }
    .secondary:hover {
      background: #e8eaed;
    }
    .status {
      margin-top: 12px;
      padding: 10px;
      border-radius: 4px;
      display: none;
    }
    .status.success {
      background: #e6f4ea;
      color: #137333;
      display: block;
    }
    .status.error {
      background: #fce8e6;
      color: #c5221f;
      display: block;
    }
    .status.loading {
      background: #e8f0fe;
      color: #1a73e8;
      display: block;
    }
    .help {
      font-size: 12px;
      color: #5f6368;
      margin-top: 8px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📊 Import Slides Content</h2>
    <label for="content">Paste your formatted content below:</label>
    <textarea id="content" placeholder="# SECTION 1: Introduction

---
**Slide 1**
**Title:** Your Slide Title
**Content:**
* Bullet point 1
* Bullet point 2

**Speaker Notes:**
Your speaker notes here...

---
**Slide 2**
..."></textarea>
    <p class="help">
      Format: Use <code>---</code> to separate slides. Include <code>**Title:**</code>, 
      <code>**Subtitle:**</code>, <code>**Content:**</code>, and <code>**Speaker Notes:**</code> sections.
    </p>
    <div id="status" class="status"></div>
    <div class="button-row">
      <button class="secondary" onclick="google.script.host.close()">Cancel</button>
      <button class="secondary" onclick="clearContent()">Clear</button>
      <button class="primary" onclick="importContent()">Import Slides</button>
    </div>
  </div>

  <script>
    function importContent() {
      const content = document.getElementById('content').value.trim();
      const status = document.getElementById('status');
      
      if (!content) {
        status.className = 'status error';
        status.textContent = '⚠️ Please paste content to import.';
        return;
      }
      
      status.className = 'status loading';
      status.textContent = '⏳ Importing slides... Please wait.';
      
      google.script.run
        .withSuccessHandler(function(result) {
          status.className = 'status success';
          status.textContent = '✅ ' + result;
          setTimeout(function() {
            google.script.host.close();
          }, 2000);
        })
        .withFailureHandler(function(error) {
          status.className = 'status error';
          status.textContent = '❌ Error: ' + error.message;
        })
        .processContent(content);
    }
    
    function clearContent() {
      document.getElementById('content').value = '';
      document.getElementById('status').className = 'status';
    }
  </script>
</body>
</html>
  `;
}

// Process the content and create slides
function processContent(content) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = parseContent(content);
  
  if (slides.length === 0) {
    throw new Error('No valid slides found in the content. Check your formatting.');
  }
  
  let slidesCreated = 0;
  
  slides.forEach(function(slideData) {
    createSlide(presentation, slideData);
    slidesCreated++;
  });
  
  return 'Successfully imported ' + slidesCreated + ' slide(s)!';
}

// Parse the content into slide objects
function parseContent(content) {
  const slides = [];
  
  // Split by slide separator (---)
  const sections = content.split(/\n---\n|\n---|\n-{3,}\n/);
  
  sections.forEach(function(section) {
    section = section.trim();
    if (!section) return;
    
    // Check if this section contains slide content
    if (section.includes('**Slide') || section.includes('**Title:**')) {
      const slideData = parseSlideSection(section);
      if (slideData.title || slideData.content) {
        slides.push(slideData);
      }
    }
  });
  
  return slides;
}

// Parse individual slide section
function parseSlideSection(section) {
  const slideData = {
    title: '',
    subtitle: '',
    content: '',
    speakerNotes: ''
  };
  
  // Extract Title
  const titleMatch = section.match(/\*\*Title:\*\*\s*(.+?)(?=\n\*\*|\n\n|$)/s);
  if (titleMatch) {
    slideData.title = cleanText(titleMatch[1]);
  }
  
  // Extract Subtitle
  const subtitleMatch = section.match(/\*\*Subtitle:\*\*\s*(.+?)(?=\n\*\*|\n\n|$)/s);
  if (subtitleMatch) {
    slideData.subtitle = cleanText(subtitleMatch[1]);
  }
  
  // Extract Content
  const contentMatch = section.match(/\*\*Content:\*\*\s*([\s\S]+?)(?=\n\*\*Speaker Notes:\*\*|$)/);
  if (contentMatch) {
    slideData.content = cleanContent(contentMatch[1]);
  }
  
  // Extract Speaker Notes
  const notesMatch = section.match(/\*\*Speaker Notes:\*\*\s*([\s\S]+?)$/);
  if (notesMatch) {
    slideData.speakerNotes = cleanText(notesMatch[1]);
  }
  
  return slideData;
}

// Clean text by removing markdown formatting
function cleanText(text) {
  return text
    .replace(/\*\*/g, '')
    .replace(/\*/g, '')
    .replace(/\[|\]/g, '')
    .replace(/^\s+|\s+$/g, '')
    .trim();
}

// Clean content text and preserve structure
function cleanContent(text) {
  return text
    .replace(/\*\*/g, '')
    .replace(/^\*\s+/gm, '• ')
    .replace(/^\d+\.\s+/gm, function(match) { return match; })
    .replace(/\[|\]/g, '')
    .trim();
}

// Create a slide with the parsed data
function createSlide(presentation, slideData) {
  // Choose layout based on content
  let layout;
  if (slideData.subtitle) {
    layout = SlidesApp.PredefinedLayout.TITLE;
  } else if (slideData.content) {
    layout = SlidesApp.PredefinedLayout.TITLE_AND_BODY;
  } else {
    layout = SlidesApp.PredefinedLayout.TITLE_ONLY;
  }
  
  const slide = presentation.appendSlide(layout);
  const shapes = slide.getShapes();
  
  // Find and populate placeholders
  shapes.forEach(function(shape) {
    const placeholderType = shape.getPlaceholderType();
    
    if (placeholderType === SlidesApp.PlaceholderType.TITLE || 
        placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
      if (slideData.title) {
        shape.getText().setText(slideData.title);
      }
    }
    else if (placeholderType === SlidesApp.PlaceholderType.SUBTITLE) {
      if (slideData.subtitle) {
        shape.getText().setText(slideData.subtitle);
      } else if (slideData.content && !hasBodyPlaceholder(shapes)) {
        shape.getText().setText(slideData.content);
      }
    }
    else if (placeholderType === SlidesApp.PlaceholderType.BODY) {
      if (slideData.content) {
        shape.getText().setText(slideData.content);
      }
    }
  });
  
  // If we have content but no body placeholder was found, create a text box
  if (slideData.content && !hasBodyPlaceholder(shapes)) {
    const contentBox = slide.insertTextBox(slideData.content, 50, 150, 600, 300);
    contentBox.getText().getTextStyle().setFontSize(14);
  }
  
  // Add speaker notes
  if (slideData.speakerNotes) {
    slide.getNotesPage().getSpeakerNotesShape().getText().setText(slideData.speakerNotes);
  }
  
  return slide;
}

// Check if shapes contain a body placeholder
function hasBodyPlaceholder(shapes) {
  for (let i = 0; i < shapes.length; i++) {
    if (shapes[i].getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
      return true;
    }
  }
  return false;
}
