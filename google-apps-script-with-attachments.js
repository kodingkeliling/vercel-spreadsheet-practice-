// Variabel wbook untuk mengintegrasikan Apps Script dengan halamanan Spredsheet
const wbook = SpreadsheetApp.getActive();

// menginisialisasi nama di wbook spreadsheet google
const sheet = wbook.getSheetByName('note');

// Function Get buat dengan Konsep doGet
function doGet() {
  let data = [];
  const rlen = sheet.getLastRow();
  const clen = sheet.getLastColumn();
  const rows = sheet.getRange(1, 1, rlen, clen).getValues();

  for(let i = 0; i < rows.length; i++) {
    const dataRow = rows[i];
    let record = {};
    
    for(let j = 0; j < clen; j++) {
      record[rows[0][j]] = dataRow[j];
    }
    
    if(i > 0) {
      data.push(record);
    }
  }

  const response = {
    data: data
  };

  console.log(response);
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// Function Post untuk CRUD operations
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch (action) {
      case 'create':
        return handleCreate(data.data);
      case 'update':
        return handleUpdate(data.id, data.data);
      case 'delete':
        return handleDelete(data.id);
      default:
        throw new Error('Invalid action: ' + action);
    }
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle Create Note
function handleCreate(noteData) {
  // Convert tags to string properly
  let tagsString = '';
  if (noteData.tags) {
    if (typeof noteData.tags === 'string') {
      tagsString = noteData.tags;
    } else if (Array.isArray(noteData.tags)) {
      tagsString = noteData.tags.join(';');
    } else {
      tagsString = String(noteData.tags);
    }
  }
  
  // Convert attachments to string properly
  let attachmentsString = '';
  if (noteData.attachments) {
    if (typeof noteData.attachments === 'string') {
      attachmentsString = noteData.attachments;
    } else if (Array.isArray(noteData.attachments)) {
      attachmentsString = noteData.attachments.join(';');
    } else {
      attachmentsString = String(noteData.attachments);
    }
  }
  
  const newRow = [
    noteData.id || `note_${Date.now()}`,
    noteData.title || '',
    noteData.content || '',
    noteData.createdAt || new Date().toISOString(),
    noteData.updatedAt || new Date().toISOString(),
    tagsString,
    noteData.isArchived || false,
    attachmentsString // Add attachments column
  ];
  
  sheet.appendRow(newRow);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      data: noteData 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle Update Note
function handleUpdate(noteId, noteData) {
  const rlen = sheet.getLastRow();
  const clen = sheet.getLastColumn();
  const rows = sheet.getRange(1, 1, rlen, clen).getValues();
  const headers = rows[0];
  const idColumnIndex = headers.indexOf('id');
  
  if (idColumnIndex === -1) {
    throw new Error('ID column not found');
  }
  
  // Find the row with the matching ID
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idColumnIndex] === noteId) {
      rowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Note not found with ID: ' + noteId);
  }
  
  // Convert tags to string properly
  let tagsString = rows[rowIndex - 1][headers.indexOf('tags')] || '';
  if (noteData.tags) {
    if (typeof noteData.tags === 'string') {
      tagsString = noteData.tags;
    } else if (Array.isArray(noteData.tags)) {
      tagsString = noteData.tags.join(';');
    } else {
      tagsString = String(noteData.tags);
    }
  }
  
  // Convert attachments to string properly
  let attachmentsString = rows[rowIndex - 1][headers.indexOf('attachments')] || '';
  if (noteData.attachments) {
    if (typeof noteData.attachments === 'string') {
      attachmentsString = noteData.attachments;
    } else if (Array.isArray(noteData.attachments)) {
      attachmentsString = noteData.attachments.join(';');
    } else {
      attachmentsString = String(noteData.attachments);
    }
  }
  
  // Update the row
  const updatedRow = [
    noteData.id || noteId,
    noteData.title || rows[rowIndex - 1][headers.indexOf('title')] || '',
    noteData.content || rows[rowIndex - 1][headers.indexOf('content')] || '',
    rows[rowIndex - 1][headers.indexOf('createdAt')] || new Date().toISOString(),
    noteData.updatedAt || new Date().toISOString(),
    tagsString,
    noteData.isArchived !== undefined ? noteData.isArchived : rows[rowIndex - 1][headers.indexOf('isArchived')] || false,
    attachmentsString // Add attachments column
  ];
  
  // Update each cell in the row
  for (let i = 0; i < updatedRow.length; i++) {
    sheet.getRange(rowIndex, i + 1).setValue(updatedRow[i]);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      data: noteData 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle Delete Note
function handleDelete(noteId) {
  const rlen = sheet.getLastRow();
  const clen = sheet.getLastColumn();
  const rows = sheet.getRange(1, 1, rlen, clen).getValues();
  const headers = rows[0];
  const idColumnIndex = headers.indexOf('id');
  
  if (idColumnIndex === -1) {
    throw new Error('ID column not found');
  }
  
  // Find the row with the matching ID
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idColumnIndex] === noteId) {
      rowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Note not found with ID: ' + noteId);
  }
  
  // Delete the row
  sheet.deleteRow(rowIndex);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: 'Note deleted successfully' 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

