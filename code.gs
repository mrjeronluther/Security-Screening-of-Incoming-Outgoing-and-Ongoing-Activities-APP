function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
  .setTitle("MCD Security Screening of Incoming, Outgoing, and Ongoing Activities");
}

function getSheetTabs() {
  try {
    // Ignore the cache and always fetch the sheet names from the spreadsheet
    const spreadsheetId = '1618bKjBlFS7Vf3wtMHmMcLxJlXWba4XE0NTuWTliYLE';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets()
                              .map(sheet => sheet.getName())
                              .filter(name => name !== 'Sheet1' && name !== 'conso' && name !== 'MirrorParticularTransaction'); // Exclude 'Sheet1' and 'Conso Sheet' MirrorParticularTransaction

    // Optionally, cache the new sheet names
    const cache = CacheService.getScriptCache();
    cache.put('sheets', JSON.stringify(sheets), 3600); // Cache for 1 hour

    return sheets;
  } catch (error) {
    Logger.log('Error fetching sheets: ' + error.message);
    return [];
  }
}

function submitData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(formData.selectedSheet);
  const sheetconso = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("conso");
  const folder = DriveApp.getFolderById('18-Y0m6vCCpszsYk5p59XJ95BkaCbsMSR'); // Replace with your folder ID

  // Save the uploaded files
  const file1 = formData.file1 ? saveFile(formData.file1, folder) : null;
  const file2 = formData.file2 ? saveFile(formData.file2, folder) : null;
  const file3 = formData.file3 ? saveFile(formData.file3, folder) : null;
  const file4 = formData.file4 ? saveFile(formData.file4, folder) : null;
  const file5 = formData.file5 ? saveFile(formData.file5, folder) : null;
  const file6 = formData.file6 ? saveFile(formData.file6, folder) : null;
  const file7 = formData.file7 ? saveFile(formData.file7, folder) : null;
  const file8 = formData.file8 ? saveFile(formData.file8, folder) : null;
  const file9 = formData.file9 ? saveFile(formData.file9, folder) : null;
  const file10 = formData.file10 ? saveFile(formData.file10, folder) : null;
  const file11 = formData.file11 ? saveFile(formData.file11, folder) : null;
  const file12 = formData.file12 ? saveFile(formData.file12, folder) : null;
  

  // Get the current timestamp
  const timestamp = new Date();

  // Prepare the row data
  let rowData = [];

  if (formData.transaction === 'Incoming STL Client') {
    rowData = [
      formData.selectedSheet,  // Sheet name
      formData.name,           // Name of Operator
      formData.dateofreporting, // Date of Reporting
      formData.transaction,    // Transaction type
      formData.tnbn,           // Trade Name/Business Name
      formData.moanum,         // MOA Number
      formData.doi,            // Date of Ingress
      formData.ds,             // Date Start
      formData.de,             // Date End
      formData.certificatePresented,  // Certificate Presented
      file1 ? file1.getUrl() : '',     // File 1 URL
      formData.methodofApproval,      // Method of Approval
      file2 ? file2.getUrl() : '',     // File 2 URL
      formData.whoApprove,          // Who Approved
      timestamp                    // Timestamp
    ];
  } else if (formData.transaction === 'Outgoing STL Client') {
    rowData = [
      formData.selectedSheet,  // Sheet name
      formData.name,           // Name of Operator
      formData.dateofreporting, // Date of Reporting
      formData.transaction,    // Transaction type
      formData.osctnbn,        // Trade Name/Business Name
      formData.oscmoanum,      // MOA Number
      formData.oscdoi,         // Date of Ingress
      formData.oscds,          // Date Start
      formData.oscde,          // Date End
      formData.osccertificatePresented,  // Certificate Presented
      file3 ? file3.getUrl() : '',     // File 3 URL
      formData.oscmethodofApproval,    // Method of Approval
      file4 ? file4.getUrl() : '',     // File 4 URL
      formData.oscwhoApprove,         // Who Approved
      timestamp                    // Timestamp
    ];
  } else if (formData.transaction === 'Incoming Inline Client') {
    rowData = [
      formData.selectedSheet,  // Sheet name
      formData.name,           // Name of Operator
      formData.dateofreporting, // Date of Reporting
      formData.transaction,    // Transaction type
      formData.iictnbn,        // Trade Name/Business Name
      formData.iicdoi,         // Date of Ingress
      formData.iicds,          // Date Start
      formData.iicde,          // Date End
      formData.iiccertificatePresented,  // Certificate Presented
      file5 ? file5.getUrl() : '',     // File 5 URL
      formData.iicmethodofApproval,    // Method of Approval
      file6 ? file6.getUrl() : '',     // File 6 URL
      formData.iicwhoApprove,         // Who Approved
      timestamp                    // Timestamp
    ];
  } else if (formData.transaction === 'Outgoing Inline Client') {
    rowData = [
      formData.selectedSheet,  
      formData.name,          
      formData.dateofreporting, 
      formData.transaction,    
      formData.oictnbn,       
      formData.oicdoi,         
      formData.oicds,          
      formData.oicde,          
      formData.oiccertificatePresented,  
      file7 ? file7.getUrl() : '',    
      formData.oicmethodofApproval,    
      file8 ? file8.getUrl() : '',     
      formData.oicwhoApprove,         
      timestamp                    
    ];
  } else if (formData.transaction === 'Incoming Marketing Event Supplier') {
    rowData = [
      formData.selectedSheet,  
      formData.name,          
      formData.dateofreporting, 
      formData.transaction,    
      formData.imecn,       
      formData.imenp,        
      formData.imede,        
      formData.imesed,
      formData.imeep,
      formData.imeli,
      formData.imeqi,
      formData.imecertificatePresented,                
      file9 ? file9.getUrl(): '',    
      formData.imewhoApprove,    
      formData.imemalltenant,       
      timestamp                  
    ];
  } else if (formData.transaction === 'Outgoing Marketing Event Supplier') {
    rowData = [
      formData.selectedSheet,  
      formData.name,          
      formData.dateofreporting, 
      formData.transaction,    
      formData.omecn,       
      formData.omenp,        
      formData.omesed,        
      formData.omeep,
      formData.omeli,
      formData.omecertificatePresented,                
      file10 ? file10.getUrl() : '',    
      formData.omewhoApprove,    
      formData.omemalltenant,       
      timestamp                  
    ];
  } else if (formData.transaction === 'Delivery') {
    rowData = [
      formData.selectedSheet,  
      formData.name,          
      formData.dateofreporting, 
      formData.transaction,    

      formData.ddtsn,       
      formData.dddd,        
      formData.dddt,        
      formData.dddep,
      formData.ddcertificatePresented, 
      formData.coiothers,
      formData.ddqoei,
       
      formData.ddworkpermitapproval,  
      file12 ? file12.getUrl() : '',             
      formData.ddwhoApprove,    
       
      timestamp                  
    ];
  } else if (formData.transaction === 'Pullout') {
    rowData = [
      formData.selectedSheet,  
      formData.name,          
      formData.dateofreporting, 
      formData.transaction,    

      formData.plldtsn,       
      formData.plldd,        
      formData.pllt,        
      formData.pllext,
      formData.pllloi, 
      formData.pllqty,
       
      formData.pllcertificatePresented,               
      formData.pllwhoApprove,    
       
      timestamp                  
    ];
  } else if (formData.transaction === 'Incoming Contractor') {
  // Adjust catOptions logic
  const catOptionsValue = formData.cat === 'Options' ? formData.catothers : formData.cat;

  // Adjust seotsset logic
  const seotssetValue = formData.seotsset === 'Options' ? formData.cattoa : formData.seotsset;

  // Build the rowData array with the adjusted values
  rowData = [
    formData.selectedSheet,  
    formData.name,          
    formData.dateofreporting, 
    formData.transaction,    

    formData.icdtow,       
    catOptionsValue,           // Using the adjusted catOptions value
    formData.catothers, 

    formData.catcn,
    formData.catwl, 
    formData.catesdw, 
    formData.cateedw,
    formData.catmbi,
    formData.seot,

    seotssetValue,             // Using the adjusted seotsset value
    formData.cattoa,
    formData.mwd,
    file11 ? file11.getUrl() : '', 
    formData.catwap,               
    timestamp                  
  ];
}
 else {
    rowData = [
      formData.selectedSheet, 
      formData.name,
      formData.dateofreporting,
      formData.transaction,
      timestamp       
    ];
  }

  // Append the row to both sheets
  sheet.appendRow(rowData);
  sheetconso.appendRow(rowData);
}


function saveFile(fileData, folder) {
  if (!fileData || !fileData.data) return null;

  // Decode the base64 file data and create the file
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, fileData.name);
  const file = folder.createFile(blob);

  if (file.getSize() <= 0) {
    file.setTrashed(true);
    throw new Error('File upload failed: ' + fileData.name);
  }

  return file;
}

