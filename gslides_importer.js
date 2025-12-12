/**
 * Google Slides Content Importer
 * FIXED: Uses 'Split' logic to prevent Speaker Notes from leaking into Body.
 */

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Content Importer')
    .addItem('Import Slides Content', 'showImportDialog')
    .addToUi();
}

function showImportDialog() {
  const html = HtmlService.createHtmlOutput(getDialogHtml())
    .setWidth(700)
    .setHeight(600)
    .setTitle('Import Slides Content');
  
  SlidesApp.getUi().showModalDialog(html, 'Import Slides Content');
}

function getDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; font-family: 'Google Sans', Arial, sans-serif; }
    body { margin: 0; padding: 20px; background: #f8f9fa; }
    h2 { color: #1a73e8; margin-top: 0; display: flex; align-items: center; gap: 10px; }
    .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
    label { font-weight: 500; color: #3c4043; display: block; margin-bottom: 8px; }
    textarea { 
      width: 100%; height: 350px; padding: 12px; 
      border: 1px solid #dadce0; border-radius: 4px; 
      font-family: 'Roboto Mono', monospace; font-size: 12px; 
      resize: vertical; line-height: 1.5;
    }
    textarea:focus { outline: none; border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
    .button-row { margin-top: 16px; display: flex; gap: 12px; justify-content: flex-end; }
    button { padding: 10px 24px; border: none; border-radius: 4px; font-size: 14px; font-weight: 500; cursor: pointer; transition: background 0.2s; }
    .primary { background: #1a73e8; color: white; }
    .primary:hover { background: #1557b0; }
    .secondary { background: #f1f3f4; color: #3c4043; }
    .secondary:hover { background: #e8eaed; }
    .status { margin-top: 12px; padding: 10px; border-radius: 4px; display: none; }
    .status.success { background: #e6f4ea; color: #137333; display: block; }
    .status.error { background: #fce8e6; color: #c5221f; display: block; }
    .status.loading { background: #e8f0fe; color: #1a73e8; display: block; }
    .help { font-size: 12px; color: #5f6368; margin-top: 8px; line-height: 1.4; }
    code { background: #f1f3f4; padding: 2px 4px; border-radius: 3px; font-family: monospace; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📊 Import Slides Content</h2>
    <label for="content">Paste your content:</label>
    <textarea id="content" placeholder="--- [SLIDE 1] ---
Title: Workshop Objectives
Body:
1. Understand GKE Fleets
2. Define Platform Tenants

Speaker Notes:
By the end of this session..."></textarea>
    
    <p class="help">
      <strong>Supported Keys:</strong> <code>Title:</code>, <code>Subtitle:</code>, <code>Body:</code> (or <code>Content:</code>), and <code>Notes:</code>.<br>
      Use <code>---</code> to separate slides.
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
          setTimeout(function() { google.script.host.close(); }, 2000);
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

function processContent(content) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = parseContent(content);
  
  if (slides.length === 0) {
    throw new Error('No valid slides found. Check separators (---).');
  }
  
  let slidesCreated = 0;
  slides.forEach(function(slideData) {
    createSlide(presentation, slideData);
    slidesCreated++;
  });
  
  return 'Successfully imported ' + slidesCreated + ' slide(s)!';
}

function parseContent(content) {
  const slides = [];
  // Splits by ---, --- [SLIDE X] ---, or --------
  const sections = content.split(/^(?:\s*-{3,}.*)$/gm);
  
  sections.forEach(function(section) {
    section = section.trim();
    if (!section) return;
    
    // Check if section contains valid keywords
    if (/Title:|Body:|Content:|Notes:/i.test(section)) {
      const slideData = parseSlideSection(section);
      if (slideData.title || slideData.content) {
        slides.push(slideData);
      }
    }
  });
  
  return slides;
}

// ROBUST SPLIT PARSER
function parseSlideSection(section) {
  const slideData = {
    title: '',
    subtitle: '',
    content: '',
    speakerNotes: ''
  };

  // 1. Identify where Speaker Notes begin
  // We look for "Speaker Notes:" or "Notes:" (case insensitive, optional bold/indent)
  const notesHeaderRegex = /\n\s*(?:\*\*)?(?:Speaker\s*)?Notes(?:\*\*|:)/i;
  const notesMatch = section.match(notesHeaderRegex);
  
  let bodySection = section;
  let notesSection = '';

  // If we found a notes section, split the text into two parts
  if (notesMatch) {
    const splitIndex = notesMatch.index;
    bodySection = section.substring(0, splitIndex); // Everything before the notes
    notesSection = section.substring(splitIndex + notesMatch[0].length); // Everything after
    
    slideData.speakerNotes = cleanText(notesSection);
  }

  // 2. Extract Title from the "Body Section"
  // Match Title: at start of string or new line
  const titleMatch = bodySection.match(/(?:^|[\r\n]+)\s*(?:\*\*)?Title(?:\*\*|:)\s*(.+?)(?=\n|$)/i);
  if (titleMatch) {
    slideData.title = cleanText(titleMatch[1]);
  }

  // 3. Extract Subtitle from the "Body Section"
  const subtitleMatch = bodySection.match(/(?:^|[\r\n]+)\s*(?:\*\*)?Subtitle(?:\*\*|:)\s*(.+?)(?=\n|$)/i);
  if (subtitleMatch) {
    slideData.subtitle = cleanText(subtitleMatch[1]);
  }

  // 4. Extract Content/Body from the "Body Section"
  // It captures everything after "Body:" until the end of bodySection
  const contentMatch = bodySection.match(/(?:^|[\r\n]+)\s*(?:\*\*)?(?:Body|Content)(?:\*\*|:)\s*([\s\S]+)/i);
  if (contentMatch) {
    slideData.content = cleanContent(contentMatch[1]);
  }

  return slideData;
}

function cleanText(text) {
  return text
    .replace(/\*\*/g, '')
    .replace(/\*/g, '')
    .replace(/\[|\]/g, '')
    .replace(/^\s+|\s+$/g, '')
    .trim();
}

function cleanContent(text) {
  return text
    .replace(/\*\*/g, '')
    .replace(/^\s*[\*\-]\s+/gm, '• ')
    .replace(/^\d+\.\s+/gm, function(match) { return match; })
    .replace(/\[|\]/g, '')
    .trim();
}

function createSlide(presentation, slideData) {
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
  
  shapes.forEach(function(shape) {
    const placeholderType = shape.getPlaceholderType();
    
    if (placeholderType === SlidesApp.PlaceholderType.TITLE || 
        placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
      if (slideData.title) shape.getText().setText(slideData.title);
    }
    else if (placeholderType === SlidesApp.PlaceholderType.SUBTITLE) {
      if (slideData.subtitle) {
        shape.getText().setText(slideData.subtitle);
      } else if (slideData.content && !hasBodyPlaceholder(shapes)) {
        shape.getText().setText(slideData.content);
      }
    }
    else if (placeholderType === SlidesApp.PlaceholderType.BODY) {
      if (slideData.content) shape.getText().setText(slideData.content);
    }
  });
  
  if (slideData.content && !hasBodyPlaceholder(shapes)) {
    const contentBox = slide.insertTextBox(slideData.content, 50, 150, 600, 300);
    contentBox.getText().getTextStyle().setFontSize(14);
  }
  
  if (slideData.speakerNotes) {
    slide.getNotesPage().getSpeakerNotesShape().getText().setText(slideData.speakerNotes);
  }
  
  return slide;
}

function hasBodyPlaceholder(shapes) {
  for (let i = 0; i < shapes.length; i++) {
    if (shapes[i].getPlaceholderType() === SlidesApp.PlaceholderType.BODY) return true;
  }
  return false;
}
