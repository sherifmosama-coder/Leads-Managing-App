/**
 * Utility to fetch all application settings
 * Uses getActiveSpreadsheet() and our header map logic
 */
function getAppConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APP_CONFIG");
  const data = sheet.getDataRange().getValues();
  const map = getHeaderMap("APP_CONFIG");
  const config = {};
  
  // Skip the header row (index 0)
  for (let i = 1; i < data.length; i++) {
    let key = data[i][map["Setting_Key"] - 1];
    let value = data[i][map["Value"] - 1];
    if (key) {
      // Trim keys to avoid mismatch issues
      config[key.toString().trim()] = value; 
    }
  }
  return config;
}

/**
 * 1. Web App Routing
 */
function doGet(e) {
  const config = getAppConfig();
  const appName = config["App_Name"] || "System Login"; // Fallback if blank
  
  // We use createTemplateFromFile so we can inject variables
  var template = HtmlService.createTemplateFromFile('Login');
  template.appName = appName; 
  
  return template.evaluate()
      .setTitle(appName + ' - Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 5. Dashboard Routing (SPA)
 * Returns the evaluated HTML of the Dashboard to inject into the page seamlessly.
 */
function getDashboardView() {
  const config = getAppConfig();
  var template = HtmlService.createTemplateFromFile('Dashboard');
  template.appName = config["App_Name"] || "CRM Dashboard";
  return template.evaluate().getContent();
}

/**
 * 6. Data Engine: Fetch Kanban Data
 * Retrieves pipeline stages and active deals, mapping contacts automatically.
 * Includes Row-Level Security based on user email and role.
 */
function getKanbanData(userEmail, userRole) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get Pipeline Stages from vertical SYS_DROPDOWNS
    const ddSheet = ss.getSheetByName("SYS_DROPDOWNS");
    const ddData = ddSheet.getDataRange().getValues();
    const ddMap = getHeaderMap("SYS_DROPDOWNS");
    let stages = [];
    
    // Check if Pipeline_Stage column exists
    const stageColIndex = ddMap["Pipeline_Stage"] ? ddMap["Pipeline_Stage"] - 1 : -1;
    
    if (stageColIndex !== -1) {
      // Loop from row 2 downwards (index 1)
      for (let i = 1; i < ddData.length; i++) {
        let stageName = ddData[i][stageColIndex];
        if (stageName) {
          stages.push({
            name: stageName,
            order: i, // Uses row number for sorting
            color: 'var(--primary)' // Default color for vertical lists
          });
        }
      }
    }

    // 2. Get Contacts for cross-referencing names
    const contactSheet = ss.getSheetByName("DB_CONTACTS");
    const contactData = contactSheet.getDataRange().getValues();
    const contactMap = getHeaderMap("DB_CONTACTS");
    let contactDict = {};
    
    for (let i = 1; i < contactData.length; i++) {
      let cid = contactData[i][contactMap["Contact_ID"] - 1];
      let fName = contactData[i][contactMap["First_Name"] - 1];
      let lName = contactData[i][contactMap["Last_Name"] - 1];
      if (cid) contactDict[cid] = (fName + " " + lName).trim();
    }

    // 3. Get Deals
    const dealSheet = ss.getSheetByName("DB_DEALS");
    const dealData = dealSheet.getDataRange().getValues();
    const dealMap = getHeaderMap("DB_DEALS");
    let deals = [];
    
    for (let i = 1; i < dealData.length; i++) {
      // Safely check if mapping exists to avoid undefined math
      let dealIdIndex = dealMap["Deal_ID"] ? dealMap["Deal_ID"] - 1 : -1;
      if (dealIdIndex === -1 || !dealData[i][dealIdIndex]) continue; // Skip if invalid or empty
      
      let contactId = dealMap["Contact_ID"] ? dealData[i][dealMap["Contact_ID"] - 1] : "";
      let assignedTo = dealMap["Assigned_To"] ? dealData[i][dealMap["Assigned_To"] - 1] : "";
      
      // ROW-LEVEL SECURITY: Filter out deals if user is not an Admin AND not the assigned agent
      if (userRole !== "Admin" && String(assignedTo).trim().toLowerCase() !== String(userEmail).trim().toLowerCase()) {
         continue; 
      }
      
      deals.push({
        id: dealData[i][dealIdIndex],
        dealName: dealMap["Deal_Name"] ? dealData[i][dealMap["Deal_Name"] - 1] : "Unnamed",
        contactName: contactDict[contactId] || "Unknown Contact",
        stage: dealMap["Pipeline_Stage"] ? dealData[i][dealMap["Pipeline_Stage"] - 1] : "",
        value: dealMap["Estimated_Value"] ? dealData[i][dealMap["Estimated_Value"] - 1] || 0 : 0,
        probability: dealMap["Probability_Percent"] ? dealData[i][dealMap["Probability_Percent"] - 1] || 0 : 0,
        assignedTo: dealMap["Assigned_To"] ? dealData[i][dealMap["Assigned_To"] - 1] : "",
        expectedClose: dealMap["Expected_Close_Date"] ? dealData[i][dealMap["Expected_Close_Date"] - 1] : ""
      });
    }

    // Bulletproof Serialization: Strips undefined values and sanitizes dates natively
    const safePayload = JSON.parse(JSON.stringify({ success: true, stages: stages, deals: deals }));
    return safePayload;
    
  } catch (error) {
    return { success: false, message: "Error fetching pipeline data: " + error.message };
  }
}

/**
 * 8. File Management Engine
 * Ensures a centralized folder exists for the CRM and returns the ID.
 */
function getOrCreateRootFolder() {
  const folderName = "CRM_System_Attachments";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

/**
 * 7. Data Engine: Process Full Deal Move
 * Handles: Audit Logs, Note Entry, Admin-only Routing, and File Attachment URLs.
 */
function processDealMove(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();
    const userEmail = payload.userEmail;
    
    // --- 1. UPDATE DB_DEALS ---
    const dealSheet = ss.getSheetByName("DB_DEALS");
    const dealData = dealSheet.getDataRange().getValues();
    const dealMap = getHeaderMap("DB_DEALS");
    
    let oldStage = "Unknown";
    let currentAgent = "";
    let dealRow = -1;
    
    for (let i = 1; i < dealData.length; i++) {
      if (String(dealData[i][dealMap["Deal_ID"] - 1]) === String(payload.dealId)) {
        dealRow = i + 1;
        oldStage = dealData[i][dealMap["Pipeline_Stage"] - 1];
        currentAgent = dealData[i][dealMap["Assigned_To"] - 1];
        break;
      }
    }
    
    if (dealRow === -1) return { success: false, message: "Deal record not found." };

    // SECURITY: Only Admin can change Assigned_To.
    const finalAgent = (payload.userRole === "Admin") ? payload.assignedTo : currentAgent;

    // Apply updates to the sheet
    if(dealMap["Pipeline_Stage"]) dealSheet.getRange(dealRow, dealMap["Pipeline_Stage"]).setValue(payload.newStage);
    if(dealMap["Assigned_To"]) dealSheet.getRange(dealRow, dealMap["Assigned_To"]).setValue(finalAgent);
    if(dealMap["Estimated_Value"]) dealSheet.getRange(dealRow, dealMap["Estimated_Value"]).setValue(payload.value);
    
    // Format: Store 50 as 0.5 for Sheet Number Formatting
    if(dealMap["Probability_Percent"]) dealSheet.getRange(dealRow, dealMap["Probability_Percent"]).setValue(Number(payload.probability) / 100);
    
    if(dealMap["Expected_Close_Date"]) dealSheet.getRange(dealRow, dealMap["Expected_Close_Date"]).setValue(payload.closeDate);
    if(dealMap["Last_Updated_By"]) dealSheet.getRange(dealRow, dealMap["Last_Updated_By"]).setValue(userEmail);
    if(dealMap["Last_Updated_Date"]) dealSheet.getRange(dealRow, dealMap["Last_Updated_Date"]).setValue(timestamp);
    
    // NEW: Save the Attachment URL if provided
    if(payload.attachmentUrl && dealMap["Attachment_URL"]) {
      dealSheet.getRange(dealRow, dealMap["Attachment_URL"]).setValue(payload.attachmentUrl);
    }

    // --- 2. WRITE TO DB_LOGS ---
    const logSheet = ss.getSheetByName("DB_LOGS");
    logSheet.appendRow([
      "LOG-" + Utilities.getUuid().substring(0,8).toUpperCase(), 
      timestamp, 
      userEmail, 
      "Stage Change", 
      payload.dealId, 
      oldStage, 
      payload.newStage
    ]);

    // --- 3. WRITE TO DB_NOTES ---
    if (payload.noteText && payload.noteText.trim() !== "") {
      const noteSheet = ss.getSheetByName("DB_NOTES");
      noteSheet.appendRow([
        "NOT-" + Utilities.getUuid().substring(0,8).toUpperCase(), 
        payload.dealId, 
        payload.noteText.trim(), 
        userEmail, 
        timestamp
      ]);
    }
    
    return { success: true, message: "Deal updated successfully." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Fetch all agents and currency config for the modal
 */
function getModalRequirements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAppConfig();
  const ddSheet = ss.getSheetByName("SYS_DROPDOWNS");
  const ddData = ddSheet.getDataRange().getValues();
  const ddMap = getHeaderMap("SYS_DROPDOWNS");
  
  let agents = [];
  const agentCol = ddMap["Agents"] - 1;
  if(agentCol >= 0) {
    for(let i=1; i<ddData.length; i++) {
       if(ddData[i][agentCol]) agents.push(ddData[i][agentCol]);
    }
  }
  
  return {
    agents: agents,
    currency: config["Currency_Symbol"] || "$"
  };
}

/**
 * 7. Data Engine: Process Full Deal Move
 * Updates DB_DEALS, writes an audit trail to DB_LOGS, and saves comments to DB_NOTES.
 */
function processDealMove(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();
    
    // --- 1. UPDATE DB_DEALS ---
    const dealSheet = ss.getSheetByName("DB_DEALS");
    const dealData = dealSheet.getDataRange().getValues();
    const dealMap = getHeaderMap("DB_DEALS");
    
    let oldStage = "Unknown";
    let dealFound = false;
    
    for (let i = 1; i < dealData.length; i++) {
      if (String(dealData[i][dealMap["Deal_ID"] - 1]) === String(payload.dealId)) {
        oldStage = dealData[i][dealMap["Pipeline_Stage"] - 1];
        const row = i + 1;
        
        // Update core fields
        if(dealMap["Pipeline_Stage"]) dealSheet.getRange(row, dealMap["Pipeline_Stage"]).setValue(payload.newStage);
        if(dealMap["Assigned_To"]) dealSheet.getRange(row, dealMap["Assigned_To"]).setValue(payload.assignedTo);
        if(dealMap["Estimated_Value"]) dealSheet.getRange(row, dealMap["Estimated_Value"]).setValue(payload.value);
        if(dealMap["Probability_Percent"]) dealSheet.getRange(row, dealMap["Probability_Percent"]).setValue(payload.probability);
        if(dealMap["Expected_Close_Date"]) dealSheet.getRange(row, dealMap["Expected_Close_Date"]).setValue(payload.closeDate);
        
        // Stamp tracking fields
        if(dealMap["Last_Updated_By"]) dealSheet.getRange(row, dealMap["Last_Updated_By"]).setValue(payload.userEmail);
        if(dealMap["Last_Updated_Date"]) dealSheet.getRange(row, dealMap["Last_Updated_Date"]).setValue(timestamp);
        
        dealFound = true;
        break;
      }
    }
    
    if (!dealFound) return { success: false, message: "Deal record could not be found." };

    // --- 2. WRITE TO DB_LOGS ---
    const logSheet = ss.getSheetByName("DB_LOGS");
    const logId = "LOG-" + Utilities.getUuid().substring(0,8).toUpperCase();
    logSheet.appendRow([
      logId, 
      timestamp, 
      payload.userEmail, 
      "Stage Change", 
      payload.dealId, 
      oldStage, 
      payload.newStage
    ]);

    // --- 3. WRITE TO DB_NOTES (If a note was added) ---
    if (payload.noteText && payload.noteText.trim() !== "") {
      const noteSheet = ss.getSheetByName("DB_NOTES");
      const noteId = "NOT-" + Utilities.getUuid().substring(0,8).toUpperCase();
      noteSheet.appendRow([
        noteId, 
        payload.dealId, 
        payload.noteText.trim(), 
        payload.userEmail, 
        timestamp
      ]);
    }
    
    return { success: true, message: "Deal successfully updated." };
  } catch (error) {
    return { success: false, message: "System error: " + error.message };
  }
}

/**
 * 2. HTML Templating Utility
 * Allows us to separate CSS and JS into their own files later.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 3. Header Mapping Utility
 * Dynamically finds column indexes based on header names.
 * Uses getActiveSpreadsheet() as required.
 */
function getHeaderMap(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error("Sheet not found: " + sheetName);
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  headers.forEach((header, index) => {
    if (header) {
      // Trim spaces just in case of manual entry errors
      map[header.toString().trim()] = index + 1; 
    }
  });
  return map;
}

/**
 * 4. Authentication Logic
 * Checks the DB_USERS sheet for matching credentials using header mapping.
 */
function verifyLogin(email, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DB_USERS");
    const data = sheet.getDataRange().getValues();
    const map = getHeaderMap("DB_USERS");
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      let userEmail = data[i][map["Email"] - 1];
      // Update headers to match relational architecture and force string conversion
      let userPass = String(data[i][map["Password_Hash"] - 1]); 
      let userStatus = data[i][map["Account_Status"] - 1];
      
      // Force input password to string as well to ensure a strict type match
      if (userEmail === email && userPass === String(password)) {
        if (userStatus !== "Active") {
          return { success: false, message: "Account is inactive. Please contact the administrator." };
        }
        
        // Return safe user data (never return the password to the frontend)
        return {
          success: true,
          userData: {
            id: data[i][map["User_ID"] - 1],
            name: data[i][map["Full_Name"] - 1],
            role: data[i][map["Role"] - 1],
            theme: data[i][map["Theme_Preference"] - 1]
          }
        };
      }
    }
    return { success: false, message: "Invalid email or password." };
  } catch (error) {
    return { success: false, message: "System error: " + error.message };
  }
}

/**
 * 9. File Upload Engine
 * Uploads a file to a deal-specific subfolder and returns the public URL.
 */
function uploadFileToDeal(fileData, fileName, dealId) {
  try {
    const rootFolder = getOrCreateRootFolder();
    let dealFolder;
    
    // البحث عن فولدر خاص بالديل أو إنشاؤه
    const subFolders = rootFolder.getFoldersByName(dealId);
    if (subFolders.hasNext()) {
      dealFolder = subFolders.next();
    } else {
      dealFolder = rootFolder.createFolder(dealId);
    }
    
    // تحويل البيانات لملف وحفظه
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), fileData.split(',')[0].split(':')[1].split(';')[0], fileName);
    const file = dealFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    throw new Error("File Upload Failed: " + e.message);
  }
}
