### **Installation Instructions**

1. **Open Google Docs and access Apps Script:**
   - Open your Google Doc
   - Click on `Extensions` → `Apps Script`

2. **Create the Code.gs file:**
   - Delete any existing code in the default file
   - Copy and paste the entire Code.gs content from above
   - The file should already be named `Code.gs`

3. **Create the Sidebar.html file:**
   - Click the `+` button next to "Files" in the left sidebar
   - Select `HTML`
   - Name it `Sidebar` (without the .html extension)
   - Replace all content with the HTML code from above

4. **Save and authorize:**
   - Click the save icon or press `Ctrl+S` (or `Cmd+S` on Mac)
   - Close the Apps Script editor tab
   - Return to your Google Doc and refresh the page
   - You should see a new menu called "Markdown Converter"
   - Click `Markdown Converter` → `Open Converter`
   - You may need to authorize the script on first use

### **Supported Markdown Features**

The converter supports the following Markdown syntax:

- **Headers:** `#` for H1, `##` for H2, `###` for H3
- **Bold text:** `**text**`
- **Italic text:** `*text*`
- **Bulleted lists:** `*` or `-` at the start of a line
- **Numbered lists:** `1.`, `2.`, etc.
- **Code blocks:** Text between triple backticks
- **Inline code:** Text between single backticks

### **Usage Example**

Once installed, you can paste Markdown like this into the sidebar:

```markdown
# My Document Title

This is a paragraph with **bold text** and *italic text*.

## Section 1

Here's a list:
* First item
* Second item
* Third item

### Subsection

1. Numbered item one
2. Numbered item two

```
function example() {
  return "Hello World";
}
```

Some inline `code` here.
```

The converter will transform it into properly formatted Google Docs content with appropriate headings, formatting, and styles.
