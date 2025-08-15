// This Google Apps Script function copies comments (including replies) from one Google Doc to another.
// Note: This assumes the target document has identical or similar content to the source, as comments use anchors that reference specific positions in the document.
// If the content differs significantly, the anchors may not align properly.

// Prerequisites:
// 1. In the Apps Script editor, go to Resources > Advanced Google services, and enable "Drive API" (Identifier: Drive, Version: v2).
// 2. In the Google Cloud Console[](https://console.cloud.google.com/apis/library/drive.googleapis.com), enable the Drive API for your project.

// Function to copy comments from source Doc to target Doc
function copyComments(sourceDocId, targetDocId) {
  // Retrieve the list of comments from the source document
  var commentList = Drive.Comments.list(sourceDocId);
  
  // Iterate through each comment
  commentList.items.forEach(function(item) {
    // Temporarily store replies and remove them from the item (as insert doesn't accept replies directly)
    var replies = item.replies || [];
    delete item.replies;
    
    // Insert the comment into the target document and get its new ID
    var newCommentId = Drive.Comments.insert(item, targetDocId).commentId;
    
    // Insert each reply into the new comment
    replies.forEach(function(reply) {
      Drive.Replies.insert(reply, targetDocId, newCommentId);
    });
  });
  
  Logger.log('Comments copied successfully from ' + sourceDocId + ' to ' + targetDocId);
}

// Example usage: Copy comments to a new duplicate of the source document
function copyCommentsToNewDoc() {
  var sourceDocId = 'YOUR_SOURCE_DOCUMENT_ID_HERE'; // Replace with your source Google Doc ID
  
  // Make a copy of the source document
  var sourceFile = DriveApp.getFileById(sourceDocId);
  var targetFile = sourceFile.makeCopy('Copy of ' + sourceFile.getName());
  var targetDocId = targetFile.getId();
  
  // Copy the comments
  copyComments(sourceDocId, targetDocId);
  
  Logger.log('New document created with ID: ' + targetDocId);
}

// To use for two existing documents:
// - Call copyComments('SOURCE_ID', 'TARGET_ID');
// Note: This copies only open comments by default. To include deleted/resolved comments, add {includeDeleted: true} to Drive.Comments.list(sourceDocId, {includeDeleted: true}).
// Dates of comments will reflect the copy time, as create/modified times cannot be preserved.
