/**
 * Serves the web app and passes the app's URL to the HTML template.
 * This allows for safe redirection on login/logout.
 */
function doGet(e) {
  const userCache = CacheService.getUserCache();
  const isLoggedIn = userCache.get('isLoggedIn'); // Check the cache
  const scriptUrl = ScriptApp.getService().getUrl();

  // Clear any stale login state if needed
  Logger.log('Login status: ' + isLoggedIn);
  Logger.log('Script URL: ' + scriptUrl);

  if (isLoggedIn === 'true') {
    const template = HtmlService.createTemplateFromFile('Index');
    template.url = scriptUrl;
    return template.evaluate()
      .setTitle('PMS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    const template = HtmlService.createTemplateFromFile('Login');
    template.url = scriptUrl;
    return template.evaluate()
      .setTitle('PMS Login');
  }
}

/**
 * Authenticates a user against the 'Doer list' sheet.
 */
function authenticateUser(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Doer list');
    if (!sheet) return { success: false, message: 'Configuration error: "Doer list" sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const nameIndex = headers.indexOf('Name');
    const passwordIndex = headers.indexOf('Password');
    const roleIndex = headers.indexOf('Role');
    // --- UPDATED: Look for ISO column ---
    const isoIndex = headers.indexOf('ISO No.'); 

    const userRow = data.find(row => row[nameIndex].toString().trim() === username.trim());

    if (userRow && userRow[passwordIndex].toString() === password) {
      const userCache = CacheService.getUserCache();
      userCache.put('isLoggedIn', 'true', 14400); 

      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('username', userRow[nameIndex].toString());
      
      if (roleIndex !== -1) {
        userProperties.setProperty('userRole', userRow[roleIndex].toString());
      }
      
      // --- UPDATED: Store the ISO Number ---
      if (isoIndex !== -1 && userRow[isoIndex]) {
        // userProperties.setProperty('assignedIso', userRow[isoIndex].toString());
        
        // Clean the ISO before storing it in the user session
        userProperties.setProperty('assignedIso', cleanIso(userRow[isoIndex]));
      } else {
        userProperties.deleteProperty('assignedIso');
      }

      return { success: true, message: 'Login successful.' };
    } else {
      return { success: false, message: 'Invalid username or password.' };
    }
  } catch (e) {
    Logger.log(`Authentication Error: ${e.toString()}`);
    return { success: false, message: 'An error occurred during authentication.' };
  }
}

/**
 * Retrieves the current user's information.
 */
function getUserInfo() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    return { 
      username: userProperties.getProperty('username'), 
      userRole: userProperties.getProperty('userRole') || 'N/A',
      assignedIso: userProperties.getProperty('assignedIso') || null // Return the ISO
    };
  } catch (e) {
    return null;
  }
}

/**
 * Fetches the assigned project for a specific user based on the 'Doer list'.
 * 1. Finds the ISO No. for the username in 'Doer list'.
 * 2. Finds the Project Name for that ISO No. in 'Project Details Form'.
 * @param {string} username The logged-in username.
 * @returns {object} { projectName: "...", isoNo: "..." } or null if not found.
 */
function getCustomerAssignedProject(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get ISO No. from 'Doer list'
    const doerSheet = ss.getSheetByName('Doer list');
    if (!doerSheet) return null;
    
    const doerData = doerSheet.getDataRange().getValues();
    const doerHeaders = doerData.shift(); // Headers: Name, Password, Role, ISO No.
    
    const nameIdx = doerHeaders.indexOf('Name');
    const isoIdx = doerHeaders.indexOf('ISO No.');
    
    if (nameIdx === -1 || isoIdx === -1) return null;
    
    const doerRow = doerData.find(row => row[nameIdx].toString().trim() === username);
    if (!doerRow) return null;
    
    const assignedIso = doerRow[isoIdx];
    if (!assignedIso) return null;

    // 2. Get Project Name from 'Project Details Form' using that ISO No.
    const detailsSheet = ss.getSheetByName('Project Details Form');
    if (!detailsSheet) return null;
    
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projNameIdx = detailsHeaders.indexOf('PROJECT NAME');
    const detailsIsoIdx = detailsHeaders.indexOf('ISO No.');
    
    if (projNameIdx === -1 || detailsIsoIdx === -1) return null;
    
    const projectRow = detailsData.find(row => row[detailsIsoIdx] === assignedIso);
    
    if (projectRow) {
      return {
        projectName: projectRow[projNameIdx],
        isoNo: assignedIso
      };
    }
    
    return null;

  } catch (e) {
    Logger.log('Error in getCustomerAssignedProject: ' + e.toString());
    return null;
  }
}

/**
 * NEW HELPER: Finds the Project Name based on the ISO No.
 * This is needed because your Dashboard functions require Project Name to load data.
 */
function getProjectNameFromIso(isoNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project Details Form');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const isoColIdx = headers.indexOf('ISO No.');
  const projectColIdx = headers.indexOf('PROJECT NAME');

  if (isoColIdx === -1 || projectColIdx === -1) return null;

  const projectRow = data.find(row => row[isoColIdx] === isoNo);
  
  // Return the Project Name if found
  return projectRow ? projectRow[projectColIdx] : null; 
}


/**
 * Logs the user out by deleting their session properties.
 */
function logout() {
  try {
    const userCache = CacheService.getUserCache();
    userCache.remove('isLoggedIn'); // Remove the session from the cache

    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteAllProperties(); // Also clear username/role
    
    Logger.log('User logged out successfully');
    return { success: true };


  } catch (e) {
    Logger.log('Logout error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Includes the content of another HTML file into a template.
 * This allows for component-based HTML structure (e.g., sidebars, navbars).
 * @param {string} filename The name of the HTML file to include (without .html).
 * @return {string} The HTML content of the file.
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/**
 * Modified: Fetches projects that are pending OR marked as 'NR'
 */
function getPendingProjects() {
  try {
    const externalSpreadsheetId = '1RnTKOiplVFLiZouq86fBdOm3PmaGpG0izXDwCZl0ow8';
    const sheetName = 'Main';
    const ss = SpreadsheetApp.openById(externalSpreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [['Error'], ['External sheet not found.']];

    const headerRow = 7;
    if (sheet.getLastRow() < headerRow) return [['Error'], ['No data in external sheet.']];

    const range = sheet.getRange(headerRow, 1, sheet.getLastRow() - headerRow + 1, sheet.getLastColumn());
    const values = range.getDisplayValues();
    const headers = values.shift().map(h => h.trim());
    const statusIndex = headers.indexOf('Status');
    const isoNoIndex = headers.indexOf('ISO No.');
    const bucketIndex = headers.indexOf('Bucket');

    if (statusIndex === -1 || isoNoIndex === -1 || bucketIndex === -1) {
      return [['Error'], ['Required headers "Status", "ISO No.", or "Bucket" not found.']];
    }

    const filteredRows = values.filter(row => {
      const statusValue = row[statusIndex].trim();
      const isoNoValue = row[isoNoIndex].trim();
      const bucketValue = row[bucketIndex].trim();
      
      // MODIFIED: Include both blank buckets and 'NR' buckets
      return statusValue === '' && isoNoValue !== '' && (bucketValue === '' || bucketValue === 'NR');
    });

    return [headers, ...filteredRows];

  } catch (e) {
    return [['Error'], [e.message]];
  }
}

/**
 * Finds a project by its ISO No. in the external sheet and updates its Bucket column.
 * @param {string} isoNo The unique ISO Number to identify the row.
 * @param {string} bucket The bucket value to set.
 */
function updateProjectBucket(isoNo, bucket) {
  try {
    const externalSpreadsheetId = '1RnTKOiplVFLiZouq86fBdOm3PmaGpG0izXDwCZl0ow8';
    const sheetName = 'Main';

    const ss = SpreadsheetApp.openById(externalSpreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('External sheet not found.');

    const headerRow = 7;
    const range = sheet.getRange(headerRow, 1, sheet.getLastRow() - headerRow + 1, sheet.getLastColumn());
    const values = range.getValues();
    const headers = values.shift().map(h => h.trim());

    const isoNoIndex = headers.indexOf('ISO No.');
    const bucketColIndex = headers.indexOf('Bucket');

    if (isoNoIndex === -1 || bucketColIndex === -1) {
      throw new Error('Could not find required columns (ISO No. or Bucket).');
    }

    const rowIndex = values.findIndex(row => row[isoNoIndex] === isoNo);

    if (rowIndex === -1) {
      throw new Error('Project with ISO No. ' + isoNo + ' not found.');
    }

    const sheetRow = rowIndex + headerRow + 1;
    sheet.getRange(sheetRow, bucketColIndex + 1).setValue(bucket);

    return "Success";

  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Gets a list of project names filtered by a specific bucket.
 * @param {string} bucketName The name of the bucket to filter by (e.g., 'STP CIVIL').
 * @returns {string[]} An array of project names that match the bucket.
 */
function getProjectNamesByBucket(bucketName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Project Details Form');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const projectNameIndex = headers.indexOf('PROJECT NAME');
    const bucketFieldIndex = headers.indexOf('BUCKET FIELD');

    if (projectNameIndex === -1 || bucketFieldIndex === -1) {
      return []; // Return empty if columns are missing
    }

    const filteredProjectNames = data
      .filter(row => row[bucketFieldIndex] === bucketName)
      .map(row => row[projectNameIndex]);
      
    // Return a unique list of names
    return [...new Set(filteredProjectNames)];

  } catch (e) {
    Logger.log('Error in getProjectNamesByBucket: ' + e.toString());
    return []; // Return empty array on error
  }
}

/**
 * MASTER FUNCTION (OPTIMIZED): Fetches ALL sheets (including MOM) in one go.
 */
function getAllDashboardData(projectName) {
  const result = {
    details: null,
    kpis: null,
    timelineChart: [],
    milestoneChart: [],
    momData: { headers: [], rows: [] } // Added container for MOM data
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. BATCH FETCH for KPIs (Values only is fine for math)
    const sheetNames = [
      'Project Details Form', 'PROJECT PLANNING', 'Manpower Cost', 
      'INVENTORY', 'Auditor Sheet Manpower', 'Auditor Sheet Item', 
      'Auditor Sheet Drawing', 'FINANCE/ACCOUNTS', 'Billing', 'WBS Sheet', 'Expenses'
    ];

    const allData = {};
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      allData[name] = sheet ? sheet.getDataRange().getValues() : [];
    });

    // 2. Process Project Details
    const detailsData = allData['Project Details Form'];
    if (detailsData.length > 0) {
      const headers = detailsData[0].map(h => String(h).trim());
      const projectIndex = headers.indexOf('PROJECT NAME');
      const isoIndex = headers.indexOf('ISO No.');
      
      const projectRow = detailsData.slice(1).find(row => row[projectIndex] === projectName);
      
      if (projectRow) {
        // const isoNo = projectRow[isoIndex];
        const isoNo = cleanIso(projectRow[isoIndex]);

        result.details = {
          isoNo: isoNo,
          projectAdd: projectRow[headers.indexOf('PROJECT ADD')],
          contactPerson: projectRow[headers.indexOf('CONTACT PERSON')],
          item: projectRow[headers.indexOf('item')]
        };

        if (isoNo) {
          // 3. Process KPIs and Charts
          result.kpis = calculateKpisInMemory(isoNo, allData, result.details.orderAmount);
          result.timelineChart = calculateTimelineChartInMemory(isoNo, allData['WBS Sheet']);
          result.milestoneChart = calculateMilestoneChartInMemory(isoNo, allData['PROJECT PLANNING'], allData['Billing']);

          // 4. Process MOM Data (SPECIFICALLY FOR RICH TEXT/LINKS)
          const momSheet = ss.getSheetByName('MOM');
          if (momSheet) {
             // We fetch this separately here because we need getRichTextValues for links
             const momRange = momSheet.getDataRange();
             const displayValues = momRange.getDisplayValues();
             const richTextValues = momRange.getRichTextValues();
             
             const momHeaders = displayValues.shift().map(h => h.trim());
             richTextValues.shift();
             
             const momIsoIndex = momHeaders.indexOf('ISO No.');
             
             // Columns that contain links
             const linkColumns = ['Upload Image 1', 'Upload Image 2', 'Upload Image 3', 'Upload Image 4', 'Upload Image 5', 'PDF Link', 'Client Signature (Client Attendee 1)', 'Mentor Signature (Mentor Attendee 1)'];
             const linkColumnIndices = linkColumns.map(h => momHeaders.indexOf(h));

             if (momIsoIndex > -1) {
               const filteredMomRows = [];
               displayValues.forEach((row, rowIndex) => {
                 if (row[momIsoIndex] === isoNo) {
                   // Extract links
                   linkColumnIndices.forEach(colIndex => {
                     if (colIndex !== -1) {
                       const link = richTextValues[rowIndex][colIndex].getLinkUrl();
                       if (link) row[colIndex] = link;
                     }
                   });
                   filteredMomRows.push(row);
                 }
               });
               result.momData = { headers: momHeaders, rows: filteredMomRows };
             }
          }
        }
      } else {
        return { error: "Project not found." };
      }
    }
    return result;

  } catch (e) {
    Logger.log("Error in getAllDashboardData: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * OPTIMIZED KPI CALCULATOR
 * Calculates all KPIs using data already loaded in memory. No spreadsheet calls.
 */
function calculateKpisInMemory(isoNo, allData, orderAmountFromDetails) {
  const kpis = {};
  
  // Helper to sum in memory
  const sumInMemory = (sheetName, headerRowIdx, colName, isoColName) => {
    const data = allData[sheetName];
    if (!data || data.length <= headerRowIdx) return 0;
    
    const headers = data[headerRowIdx - 1].map(h => String(h).trim());
    const norm = s => String(s).replace(/\s+/g, ' ').trim().toUpperCase();
    const valIdx = headers.findIndex(h => norm(h) === norm(colName));
    const isoIdx = headers.findIndex(h => norm(h) === norm(isoColName));
    
    if (valIdx === -1 || isoIdx === -1) return 0;

    return data.slice(headerRowIdx)
      .filter(row => cleanIso(row[isoIdx]) === cleanIso(isoNo))
      .reduce((sum, row) => sum + (parseFloat(String(row[valIdx]).replace(/,/g, '')) || 0), 0);
  };

  // --- 1. Project Info (Sum Order Amount + Handle Commas) ---
  const detailsData = allData['Project Details Form'];
  const dHeaders = detailsData[0].map(h => String(h).trim());
  const dIsoIdx = dHeaders.indexOf('ISO No.');
  const dRateIdx = dHeaders.indexOf('BOQ RATE(Rs)');
  const dManagerIdx = dHeaders.indexOf('PROJECT MANAGER');

  const dSiteInchargeIdx = dHeaders.indexOf('SITE INCHARGE'); 
  
  // Find ALL rows for this project to sum the amount
  const projectRows = detailsData.slice(1).filter(r => r[dIsoIdx] === isoNo);
  
  // Calculate Total Order Amount
  let totalOrderAmount = 0;
  if (dRateIdx > -1) {
    totalOrderAmount = projectRows.reduce((sum, row) => {
      // 1. Convert to string to safely use replace
      // 2. Remove all commas (e.g., "1,50,000" -> "150000")
      // 3. Parse to float
      let rawValue = String(row[dRateIdx] || 0);
      let cleanValue = parseFloat(rawValue.replace(/,/g, ''));
      
      return sum + (isNaN(cleanValue) ? 0 : cleanValue);
    }, 0);
  }
  kpis.orderAmount = totalOrderAmount;

  const firstRow = projectRows.length > 0 ? projectRows[0] : null;
  
  kpis.projectManager = (firstRow && dManagerIdx > -1) ? firstRow[dManagerIdx] : 'N/A';

  kpis.siteIncharge = (firstRow && dSiteInchargeIdx > -1) ? firstRow[dSiteInchargeIdx] : 'N/A';
  
  // Payment Terms
  if (firstRow) {
      kpis.paymentTerms = {
        advance: firstRow[dHeaders.indexOf('ADVANCE (%)')] || 0,
        againstPi: firstRow[dHeaders.indexOf('AGAINST PI (%)')] || 0,
        installTest: firstRow[dHeaders.indexOf('INSTALLATION & TESTING (%)')] || 0,
        commission: firstRow[dHeaders.indexOf('COMMISSIONING & HANDOVER (%)')] || 0,
        security: firstRow[dHeaders.indexOf('SECURITY AMOUNT (%)')] || 0,
      };
      
      const poLinkCell = firstRow[dHeaders.indexOf('CUSTOMER PO COPY')];
      kpis.poCopy = (typeof poLinkCell === 'string' && poLinkCell.startsWith('http')) ? poLinkCell : '#';
  }

  // --- 2. Expenses ---
  kpis.manpowerUsedRs = sumInMemory("Manpower Cost", 1, "Manpower Cost", "ISO No.");
  kpis.productUsedRs = sumInMemory("INVENTORY", 2, "CONSU.2", "ISO No.");

  // --- 2. Expenses (UPDATED for Debit Amount) ---
  const expenseSheetData = allData['Expenses'];
  if(expenseSheetData && expenseSheetData.length > 0) {
      const eHeaders = expenseSheetData[0].map(h => String(h).trim());
      
      // Robust search for ISO column (handles "ISO No." or "ISO No")
      let eIsoIdx = eHeaders.indexOf('ISO No.');
      if (eIsoIdx === -1) eIsoIdx = eHeaders.indexOf('ISO No');

      // Search for Debit Amount column
      const eDebitIdx = eHeaders.indexOf('Debit Amount');
      
      if (eIsoIdx > -1 && eDebitIdx > -1) {
          const targetIso = String(isoNo).trim();
          
          kpis.expensesRs = expenseSheetData.slice(1)
            .filter(r => {
                // Extract clean ISO (removes " - Client Name" if present)
                const rowIso = String(r[eIsoIdx]).split(' - ')[0].trim();
                return rowIso === targetIso;
            })
            .reduce((sum, r) => {
                // Remove commas and sum the Debit Amount
                let rawVal = String(r[eDebitIdx] || 0);
                let val = parseFloat(rawVal.replace(/,/g, ''));
                return sum + (isNaN(val) ? 0 : val);
            }, 0);
      } else { 
        Logger.log("Expenses Error: 'ISO No.' or 'Debit Amount' column not found.");
        kpis.expensesRs = 0; 
      }
  } else { kpis.expensesRs = 0; }

  // --- 3. Product Status ---
  kpis.indentPlanned = sumInMemory("INVENTORY", 2, "Planned Indent Qty", "ISO No.");
  kpis.actualIndent = sumInMemory("INVENTORY", 2, "INDENT1", "ISO No.");
  kpis.received = sumInMemory("INVENTORY", 2, "REC.1", "ISO No.");
  kpis.used = sumInMemory("INVENTORY", 2, "CONSU.1", "ISO No.");

  // --- 4. Manpower & Engineer ---
  const pData = allData['PROJECT PLANNING'];
  if (pData.length > 0) {
      const pHeaders = pData[0].map(h => String(h).trim());
      const pIsoIdx = pHeaders.indexOf("ISO No.");
      const doerIdx = pHeaders.indexOf("Engineer Name (Doer)");
      const reqHrsIdx = pHeaders.indexOf("Required Hours");
      const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");
      
      if (pIsoIdx > -1) {
          const projRows = pData.slice(1).filter(r => r[pIsoIdx] === isoNo);
          
          if (doerIdx > -1) {
             const doers = new Set(projRows.map(r => r[doerIdx]).filter(String));
             kpis.engWorked = doers.size;
          }
          if (reqHrsIdx > -1) kpis.plannedManHrs = projRows.reduce((sum, r) => sum + (parseFloat(r[reqHrsIdx])||0), 0);
          if (workedHrsIdx > -1) kpis.actualManHrs = projRows.reduce((sum, r) => sum + (parseFloat(r[workedHrsIdx])||0), 0);
      }
  }

  // --- 5. Audit ---
  kpis.auditManpower = sumInMemory("Auditor Sheet Manpower", 2, "Deviation", "ISO No.");
  kpis.auditProduct = sumInMemory("Auditor Sheet Item", 2, "Deviation", "ISO No.");
  kpis.auditDrawings = sumInMemory("Auditor Sheet Drawing", 2, "Deviation", "ISO NO.");

  
  // --- 6. Milestone Summary ---
  const wbsData = allData['WBS Sheet'];
  if (wbsData && wbsData.length > 0) {
    const wHeaders = wbsData[0].map(h => String(h).trim());
    const wIsoIdx = wHeaders.indexOf("ISO No.");
    const wMileIdx = wHeaders.findIndex(h => h.replace(/\s+/g, ' ').toUpperCase() === "MILESTONE/BOQ ITEM");
    const wActualEndIdx = wHeaders.findIndex(h => h.replace(/\s+/g, ' ').toUpperCase() === "ACTUAL END DATE");

    if (wIsoIdx === -1 || wMileIdx === -1) Logger.log('WBS headers not found: ' + JSON.stringify(wHeaders));
    if (wIsoIdx > -1 && wMileIdx > -1) {
      const targetIso = String(isoNo).trim();
      const projWbsRows = wbsData.slice(1).filter(r => cleanIso(r[wIsoIdx]) === targetIso);

      // Group by Unique Milestone Name to get the count of 12
      const milestoneGroups = {}; 
      projWbsRows.forEach(row => {
        const mName = String(row[wMileIdx]).trim();
        if (!mName || mName === "N/A" || mName === "") return;
        if (!milestoneGroups[mName]) milestoneGroups[mName] = [];
        milestoneGroups[mName].push(row);
      });

      const uniqueMilestoneNames = Object.keys(milestoneGroups);
      kpis.totalMilestones = uniqueMilestoneNames.length; // This will now show 12

      let mCompleted = 0;
      uniqueMilestoneNames.forEach(name => {
        const tasksInThisMilestone = milestoneGroups[name];
        // A milestone is "Done" if every row has an Actual End Date
        const isAllDone = tasksInThisMilestone.every(t => {
          const dateVal = t[wActualEndIdx];
          return dateVal && String(dateVal).trim() !== "" && String(dateVal).toUpperCase() !== "N/A";
        });
        if (isAllDone) mCompleted++;
      });

      kpis.completedMilestones = mCompleted;
      kpis.pendingMilestones = kpis.totalMilestones - mCompleted;
      kpis.milestoneDonePercent = kpis.totalMilestones > 0 ? Math.round((mCompleted / kpis.totalMilestones) * 100) : 0;
    }
  }

  // --- 7. Task Status (Keep this pulling from PROJECT PLANNING) ---
  if (pData && pData.length > 0) {
      const pHeaders = pData[0].map(h => String(h).trim());
      const pIsoIdx = pHeaders.indexOf("ISO No.");
      const statusIdx = pHeaders.indexOf("Status");

      if (pIsoIdx > -1 && statusIdx > -1) {
          const targetIsoStr = String(isoNo).trim();
          
          // 1. Filter all rows that match this ISO
          const projTasks = pData.slice(1).filter(row => {
            const rowIso = String(row[pIsoIdx] || "").trim();
            return rowIso === targetIsoStr;
          });

          let done = 0, hold = 0, inprog = 0;
          
          // 2. Process every task found for this ISO
          projTasks.forEach(r => {
              const s = String(r[statusIdx] || "").trim().toUpperCase();

              if (s === 'DONE') {
                  done++;
              } else if (s === 'IN-PROGRESS' || s === 'IN PROGRESS') {
                  inprog++;
              } else {
                  // UPDATED: This counts 'HOLD' as well as any EMPTY/BLANK cells as HOLD
                  hold++;
              }
          });

          // 3. Update the KPIs
          kpis.totalTasks = done + hold + inprog; // This will now show the actual total
          kpis.tasksDone = done;
          kpis.tasksInProgress = inprog;
          kpis.tasksOnHold = hold;
      } else {
          Logger.log("Error: Could not find 'ISO No.' or 'Status' column in PROJECT PLANNING");
      }
  }

  // --- 8. Billing FMS (FIXED & ROBUST) ---
  const bData = allData['Billing'];
  if (bData && bData.length > 1) {
      const bHeaders = bData[0].map(h => String(h).trim());
      
      // 1. Find Headers (Robust Search)
      // Look for "ISO NO" OR "ISO No."
      let bIsoIdx = bHeaders.indexOf("ISO NO");
      if (bIsoIdx === -1) bIsoIdx = bHeaders.indexOf("ISO No.");
      
      const bMilestoneIdx = bHeaders.indexOf("Milestones");
      
      // Look for "Status5" OR "Status"
      let bStatusIdx = bHeaders.indexOf("Status5");
      if (bStatusIdx === -1) bStatusIdx = bHeaders.indexOf("Status");

      if (bIsoIdx > -1 && bMilestoneIdx > -1 && bStatusIdx > -1) {
          const targetIso = String(isoNo).trim(); // This is already cleaned by your previous update
          
          // 2. Filter Rows using cleanIso on the sheet data
          const bRows = bData.slice(1).filter(r => {
             // We clean the row data on the fly (e.g. remove "- COMPANY NAME")
             return cleanIso(r[bIsoIdx]) === targetIso;
          });

          const totalB = bRows.length;
          
          if (totalB > 0) {
              // Count Unique Milestones
              const uniqueM = new Set(bRows.map(r => String(r[bMilestoneIdx]).trim()).filter(Boolean));
              kpis.totalBillingMilestones = uniqueM.size;
              
              // Count Completed Milestones
              let completedMCount = 0;
              uniqueM.forEach(m => {
                  const mRows = bRows.filter(r => String(r[bMilestoneIdx]).trim() === m);
                  // Check if ALL rows for this milestone are DONE
                  const allDone = mRows.every(r => String(r[bStatusIdx]).trim().toUpperCase() === 'DONE');
                  if(allDone && mRows.length > 0) completedMCount++;
              });
              kpis.completedBillingMilestones = completedMCount;

              // Calculate Percentages
              const doneB = bRows.filter(r => String(r[bStatusIdx]).trim().toUpperCase() === 'DONE').length;
              
              // Avoid division by zero
              if (totalB > 0) {
                  kpis.overallBillingDonePercent = Math.round((doneB / totalB) * 100);
                  kpis.overallBillingPendingPercent = 100 - kpis.overallBillingDonePercent;
              }
          }
      } else {
        Logger.log("Billing Error: Could not find ISO, Milestones, or Status5/Status columns.");
      }
  }

  // --- 9. Finance ---
  const fData = allData['FINANCE/ACCOUNTS'];
  if (fData && fData.length > 1) {
      const fHeaders = fData[0].map(h => String(h).trim());
      const fIsoIdx = fHeaders.indexOf("ISO No.");
      const payIdx = fHeaders.indexOf("% of  Pymt recd");
      
      if (fIsoIdx > -1 && payIdx > -1) {
          // Use slice(1) to skip header, find first match
          const fRow = fData.slice(1).find(r => r[fIsoIdx] === isoNo);
          kpis.paymentReceivedPercent = fRow ? (fRow[payIdx] || 0) : 0;
      }
  }

  // --- 10. Profit Loss ---
  const orderAmt = parseFloat(kpis.orderAmount) || 0;
  const costs = (kpis.manpowerUsedRs || 0) + (kpis.productUsedRs || 0) + (kpis.expensesRs || 0);
  kpis.profitLoss = orderAmt - costs;

  return kpis;
}

/**
 * OPTIMIZED CHART CALCULATOR (Timeline)
 */
function calculateTimelineChartInMemory(isoNo, wbsData) {
  if (!wbsData || wbsData.length < 2) return [];
  const headers = wbsData[0].map(h => String(h).trim());
  const isoIdx = headers.indexOf('ISO No.');
  const mileIdx = headers.indexOf('MILESTONE/BOQ ITEM');
  const planIdx = headers.indexOf('PLANNED START DATE');
  const actIdx = headers.indexOf('ACTUAL END DATE');

  if (isoIdx === -1) return [];

  return wbsData.slice(1)
    .filter(row => row[isoIdx] === isoNo)
    .map(row => {
      const pDate = row[planIdx];
      const aDate = row[actIdx];
      
      // Basic date check
      const fPDate = (pDate instanceof Date) ? Utilities.formatDate(pDate, Session.getScriptTimeZone(), "dd/MM/yyyy") : null;
      const fADate = (aDate instanceof Date && aDate.getTime() > 0) ? Utilities.formatDate(aDate, Session.getScriptTimeZone(), "dd/MM/yyyy") : null;

      return {
        milestone: row[mileIdx],
        planned: fPDate,
        actual: fADate
      };
    })
    .filter(d => d.milestone && d.planned);
}

/**
 * CORRECTED: Fetches and combines data for the Milestone Progress Combo Chart.
 * Uses cleanIso to ensure data is found even if suffixes exist in the Billing sheet.
 */
function calculateMilestoneChartInMemory(isoNo, planningData, billingData) {
  if (!planningData || !billingData || !isoNo) return [];
  
  const targetIso = cleanIso(isoNo);

  // --- 1. Process Planning Data ---
  const pHeaders = planningData[0].map(h => String(h).trim());
  const pIsoIdx = pHeaders.indexOf("ISO No.");
  const pMileIdx = pHeaders.indexOf("Milestone");
  const pStatIdx = pHeaders.indexOf("Status %");
  
  // Filter planning tasks using cleanIso
  const pTasks = pIsoIdx > -1 ? planningData.slice(1).filter(r => cleanIso(r[pIsoIdx]) === targetIso) : [];

  // --- 2. Process Billing Data ---
  const bHeaders = billingData[0].map(h => String(h).trim());
  
  // Robust Billing Header Search (Same as KPI logic)
  let bIsoIdx = bHeaders.indexOf("ISO NO");
  if (bIsoIdx === -1) bIsoIdx = bHeaders.indexOf("ISO No.");
  
  const bMileIdx = bHeaders.indexOf("Milestones");
  
  let bStatusIdx = bHeaders.indexOf("Status5");
  if (bStatusIdx === -1) bStatusIdx = bHeaders.indexOf("Status");

  // Filter billing tasks using cleanIso
  const bTasks = bIsoIdx > -1 ? billingData.slice(1).filter(r => cleanIso(r[bIsoIdx]) === targetIso) : [];

  // --- 3. Combine Milestone Names for the X-Axis ---
  const allMiles = new Set([
    ...pTasks.map(r => String(r[pMileIdx] || "").trim()),
    ...bTasks.map(r => String(r[bMileIdx] || "").trim())
  ].filter(Boolean));

  const chartData = [];
  allMiles.forEach(m => {
      // Calculate Task Progress (Brown Bar)
      const myTasks = pTasks.filter(r => String(r[pMileIdx] || "").trim() === m);
      let tProg = 0;
      if (myTasks.length > 0) {
          const sumPct = myTasks.reduce((s, r) => s + (parseFloat(r[pStatIdx]) || 0), 0);
          // Check if data is decimal (0.8) or whole (80)
          tProg = sumPct / myTasks.length;
          if (tProg <= 1 && tProg > 0) tProg = tProg * 100; 
      }

      // Calculate Billing Status (Yellow Bar)
      const myBills = bTasks.filter(r => String(r[bMileIdx] || "").trim() === m);
      let bStat = 0;
      if (myBills.length > 0) {
          const done = myBills.filter(r => String(r[bStatusIdx] || "").trim().toUpperCase() === 'DONE').length;
          bStat = (done / myBills.length) * 100;
      }

      chartData.push({
          milestone: m,
          taskProgress: tProg,
          billingStatus: bStat
      });
  });

  // Sort by Task Progress descending so the chart looks organized
  return chartData.sort((a, b) => b.taskProgress - a.taskProgress);
}

// /**
//  * OPTIMIZED CHART CALCULATOR (Milestone Combo)
//  */
// function calculateMilestoneChartInMemory(isoNo, planningData, billingData) {
//   if (!planningData || !billingData) return [];
  
//   // Process Planning
//   const pHeaders = planningData[0].map(h => String(h).trim());
//   const pIsoIdx = pHeaders.indexOf("ISO No.");
//   const pMileIdx = pHeaders.indexOf("Milestone");
//   const pStatIdx = pHeaders.indexOf("Status %");
//   const pTasks = pIsoIdx > -1 ? planningData.slice(1).filter(r => r[pIsoIdx] === isoNo) : [];

//   // Process Billing
//   const bHeaders = billingData[0].map(h => String(h).trim());
//   const bIsoIdx = bHeaders.indexOf("ISO NO");
//   const bMileIdx = bHeaders.indexOf("Milestones");
//   const bStatIdx = bHeaders.indexOf("Status5");
//   const bTasks = bIsoIdx > -1 ? billingData.slice(1).filter(r => r[bIsoIdx] === isoNo) : [];

//   // Combine Milestones
//   const allMiles = new Set([
//     ...pTasks.map(r => r[pMileIdx]),
//     ...bTasks.map(r => r[bMileIdx])
//   ].filter(Boolean));

//   const chartData = [];
//   allMiles.forEach(m => {
//       // Task Progress
//       const myTasks = pTasks.filter(r => r[pMileIdx] === m);
//       let tProg = 0;
//       if (myTasks.length > 0) {
//           const sumPct = myTasks.reduce((s, r) => s + (parseFloat(r[pStatIdx]) || 0), 0);
//           tProg = (sumPct / myTasks.length) * 100; // Assuming raw data is 0.5 for 50%, if 50, remove *100
//           // Wait, earlier we found your data is likely 0-1 or 0-100?
//           // If your data is 0.8 for 80%, keep *100. If it's 80, remove *100.
//           // Based on previous code "parseFloat(row[statusPercentIdx]) || 0", let's assume it's already 0-100 based on your chart fix.
//           // ADJUSTMENT: If your sheet has "80%", parseFloat is 80. If "0.8", it is 0.8.
//           // Let's assume standard number 80.
//           tProg = sumPct / myTasks.length; 
//       }

//       // Billing Status
//       const myBills = bTasks.filter(r => r[bMileIdx] === m);
//       let bStat = 0;
//       if (myBills.length > 0) {
//           const done = myBills.filter(r => String(r[bStatIdx]).trim().toUpperCase() === 'DONE').length;
//           bStat = (done / myBills.length) * 100;
//       }

//       chartData.push({
//           milestone: m,
//           taskProgress: tProg,
//           billingStatus: bStat
//       });
//   });

//   return chartData;
// }

/**
 * Gets specific details for a selected project name.
 */
function getProjectDetails(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project Details Form');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const projectIndex = headers.indexOf('PROJECT NAME');
  const isoIndex = headers.indexOf('ISO No.');
  const addIndex = headers.indexOf('PROJECT ADD');
  const personIndex = headers.indexOf('CONTACT PERSON');
  const itemIndex = headers.indexOf('item');

  const projectRow = data.find(row => row[projectIndex] === projectName);

  if (projectRow) {
    return {
      isoNo: projectRow[isoIndex],
      projectAdd: projectRow[addIndex],
      contactPerson: projectRow[personIndex],
      item: projectRow[itemIndex]
    };
  }
  return null;
}

/**
 * Fetches and prepares data for the D3.js timeline chart, using ISO No.
 * @param {string} isoNo The ISO Number of the project.
 */
function getChartData(isoNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wbsSheet = ss.getSheetByName('WBS Sheet');

  if (!wbsSheet || !isoNo) return [];

  const wbsData = wbsSheet.getDataRange().getValues();
  const wbsHeaders = wbsData.shift();

  const wbsIsoIndex = wbsHeaders.indexOf('ISO No.');
  const milestoneIndex = wbsHeaders.indexOf('MILESTONE/BOQ ITEM');
  const plannedStartIndex = wbsHeaders.indexOf('PLANNED START DATE');
  const actualEndIndex = wbsHeaders.indexOf('ACTUAL END DATE');
  if (wbsIsoIndex === -1) return [];

  const chartData = wbsData
    .filter(row => row[wbsIsoIndex] === isoNo)
    .map(row => {
      const plannedDateValue = row[plannedStartIndex];
      const actualEndDateValue = row[actualEndIndex];

      const plannedDate = (plannedDateValue instanceof Date) ?
      Utilities.formatDate(plannedDateValue, Session.getScriptTimeZone(), "dd/MM/yyyy") :
      null;

      const actualEndDate = (actualEndDateValue instanceof Date && actualEndDateValue.getTime() > 0) ?
      Utilities.formatDate(actualEndDateValue, Session.getScriptTimeZone(), "dd/MM/yyyy") :
      null;

      return {
        milestone: row[milestoneIndex],
        planned: plannedDate,
        actual: actualEndDate
      };
    })
    .filter(d => d.milestone && d.planned);
  return chartData;
}

/**
 * Gathers a comprehensive set of Key Performance Indicators (KPIs) for a given project.
 * @param {string} projectName The name of the project to fetch KPIs for.
 * @returns {object} An object containing all the calculated KPIs.
 */
function getProjectKpis(projectName = 'Hyatt regency') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpis = {}; // Object to hold all our results

  // 1. Get the ISO No. for the project (our primary key)
  const detailsSheet = ss.getSheetByName('Project Details Form');
  if (!detailsSheet) return { error: "Project Details Form not found." };
  const detailsData = detailsSheet.getDataRange().getValues();
  const detailsHeaders = detailsData.shift();
  const projectIndex = detailsHeaders.indexOf('PROJECT NAME');
  const isoIndex = detailsHeaders.indexOf('ISO No.');

  const projectRow = detailsData.find(row => row[projectIndex] === projectName);
  if (!projectRow) return { error: `Project '${projectName}' not found.` };
  const isoNo = projectRow[isoIndex];

  // 2. Fetch data from "Project Details Form"
  kpis.orderAmount = projectRow[detailsHeaders.indexOf('BOQ RATE(Rs)')] || 0;
  kpis.projectManager = projectRow[detailsHeaders.indexOf('PROJECT MANAGER')] || 'N/A';

  const poCopyIndex = detailsHeaders.indexOf('CUSTOMER PO COPY');
  let poLinkUrl = '#';
  const projectRowIndexInData = detailsData.findIndex(row => row[projectIndex] === projectName);
  if (poCopyIndex !== -1 && projectRowIndexInData !== -1) {
    const sheetRowNumber = projectRowIndexInData + 2;
    const sheetColumnNumber = poCopyIndex + 1;
    const poCopyCell = detailsSheet.getRange(sheetRowNumber, sheetColumnNumber);
    const poLink = poCopyCell.getRichTextValue();
    if (poLink) {
        poLinkUrl = poLink.getLinkUrl() || '#';
    }
  }
  kpis.poCopy = poLinkUrl;

  kpis.paymentTerms = {
    advance: projectRow[detailsHeaders.indexOf('ADVANCE (%)')] || 0,
    againstPi: projectRow[detailsHeaders.indexOf('AGAINST PI (%)')] || 0,
    installTest: projectRow[detailsHeaders.indexOf('INSTALLATION & TESTING (%)')] || 0,
    commission: projectRow[detailsHeaders.indexOf('COMMISSIONING & HANDOVER (%)')] || 0,
    security: projectRow[detailsHeaders.indexOf('SECURITY AMOUNT (%)')] || 0,
  };

  // Helper function to sum a column in a sheet based on ISO No.
  const sumColumnByIso = (sheetName, headerRowIndex, columnToSum, isoColumn) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    if (data.length < headerRowIndex) return 0;
    const headers = data[headerRowIndex - 1];
    const colIdx = headers.indexOf(columnToSum);
    const isoColIdx = headers.indexOf(isoColumn);
    if (colIdx === -1 || isoColIdx === -1) return 0;

    return data
      .slice(headerRowIndex)
      .filter(row => row[isoColIdx] === isoNo)
      .reduce((sum, row) => sum + (parseFloat(row[colIdx]) || 0), 0);
  };

  // 3. Calculate Expense Data
  kpis.manpowerUsedRs = sumColumnByIso("Manpower Cost", 1, "Manpower Cost", "ISO No.");
  kpis.productUsedRs = sumColumnByIso("INVENTORY", 2, "CONSU.2", "ISO No.");
  kpis.expensesRs = sumColumnByIso("Expenses", 1, "Amount", "ISO No.");

  // 4. Product Status Data
  kpis.indentPlanned = sumColumnByIso("INVENTORY", 2, "Qty", "ISO No.");      // NEW: From 'Qty' column
  kpis.actualIndent = sumColumnByIso("INVENTORY", 2, "INDENT1", "ISO No.");   // RENAMED: from 'indent'
  kpis.received = sumColumnByIso("INVENTORY", 2, "REC.1", "ISO No.");
  kpis.used = sumColumnByIso("INVENTORY", 2, "CONSU.1", "ISO No.");

  // 5. Manpower & Engineer Data
  const planningSheet = ss.getSheetByName("PROJECT PLANNING");
  let engWorked = 0, plannedManHrs = 0, actualManHrs = 0;
  if (planningSheet) {
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    // --- THIS IS THE CORRECTED LINE ---
    const doerNameIdx = pHeaders.indexOf("Engineer Name (Doer)"); // Targeting the correct column
    const reqHrsIdx = pHeaders.indexOf("Required Hours");
    const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");

    if (pIsoIdx > -1 && doerNameIdx > -1) {
        const projectPlanningData = pData.filter(r => r[pIsoIdx] === isoNo);
        
        // --- THIS IS THE CORRECTED LOGIC ---
        // It now maps the correct column, filters out any blank names, and then creates the unique Set.
        const uniqueEngineers = new Set(projectPlanningData.map(r => r[doerNameIdx]).filter(String));
        
        engWorked = uniqueEngineers.size;
        plannedManHrs = projectPlanningData.reduce((sum, r) => sum + (parseFloat(r[reqHrsIdx]) || 0), 0);
        actualManHrs = projectPlanningData.reduce((sum, r) => sum + (parseFloat(r[workedHrsIdx]) || 0), 0);
    }
  }
  kpis.engWorked = engWorked;
  kpis.plannedManHrs = plannedManHrs;
  kpis.actualManHrs = actualManHrs;

  // 6. Audit Data
  // Headers are on row 2 for all audit sheets. Data starts on row 3.
  // The slice() method in the helper correctly handles this.
  kpis.auditManpower = sumColumnByIso("Auditor Sheet Manpower", 2, "Deviation", "ISO No.");
  kpis.auditProduct = sumColumnByIso("Auditor Sheet Item", 2, "Deviation", "ISO No."); 
  kpis.auditDrawings = sumColumnByIso("Auditor Sheet Drawing", 2, "Deviation", "ISO No.");

  // 7. Task Status Data (4-COLUMN LOGIC)
  let totalTasks = 0, doneCount = 0, holdCount = 0, inProgressCount = 0;
  if (planningSheet) {
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const statusIdx = pHeaders.indexOf("Status");

    if (pIsoIdx > -1 && statusIdx > -1) {
        const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);
        projectTasks.forEach(task => {
            const status = (task[statusIdx] || '').toString().toUpperCase().trim();
            switch (status) {
                case 'DONE': doneCount++; break;
                case 'HOLD': holdCount++; break;
                case 'IN-PROGRESS': inProgressCount++; break;
            }
        });
        totalTasks = doneCount + holdCount + inProgressCount;
    }
  }
  kpis.totalTasks = totalTasks;
  kpis.tasksDone = doneCount;
  kpis.tasksInProgress = inProgressCount; // Now a separate value
  kpis.tasksOnHold = holdCount;         // Now a separate value


  // ====================================================================
  // ===== 8. Finance Data  =====
  // ====================================================================
  // --- Part 1: Payment Received % ---
  const financeSheet = ss.getSheetByName("FINANCE/ACCOUNTS");
  kpis.paymentReceivedPercent = 0;
  if (financeSheet) {
    const fData = financeSheet.getDataRange().getValues();
    if (fData.length > 1) {
        const fHeaders = fData[0];
        const fIsoIdx = fHeaders.indexOf("ISO No.");
        const paymentRecdIdx = fHeaders.indexOf("% of  Pymt recd");
        if (fIsoIdx > -1) {
            const dataRows = fData.slice(1);
            const financeRow = dataRows.find(r => r[fIsoIdx] === isoNo);
            if (financeRow) {
              kpis.paymentReceivedPercent = financeRow[paymentRecdIdx] || 0;
            }
        }
    }
  }

  // ====================================================================
  // ============= 9. MILESTONE SUMMARY STATUS DATA =====================
  // ====================================================================
  kpis.totalMilestones = 0;
  kpis.completedMilestones = 0;
  kpis.pendingMilestones = 0;

  const wbsData = allData['WBS Sheet'];
  if (wbsData && wbsData.length > 0) {
    const wHeaders = wbsData[0].map(h => String(h).trim());
    const wIsoIdx = wHeaders.indexOf("ISO No.");
    const wMileIdx = wHeaders.indexOf("MILESTONE/BOQ ITEM");
    const wActualEndIdx = wHeaders.indexOf("ACTUAL END DATE");

    if (wIsoIdx > -1 && wMileIdx > -1 && wActualEndIdx > -1) {
      // 1. Filter rows for this project (Using TRIM for robust matching)
      const projectIsoStr = String(isoNo).trim();
      const projWbsRows = wbsData.slice(1).filter(r => String(r[wIsoIdx]).trim() === projectIsoStr);

      
      // 2. Group all 236+ rows by their UNIQUE Milestone Name
      const milestoneGroups = {}; 
      projWbsRows.forEach(row => {
        const mName = String(row[wMileIdx]).trim();
        if (!mName) return; // Skip empty rows
        if (!milestoneGroups[mName]) milestoneGroups[mName] = [];
        milestoneGroups[mName].push(row);
      });

      const uniqueMilestoneNames = Object.keys(milestoneGroups);
      let completedCount = 0;

      // 3. Evaluate each Unique Milestone group
      uniqueMilestoneNames.forEach(name => {
        const tasksInThisMilestone = milestoneGroups[name];
        
        // A milestone is considered "DONE" only if EVERY row inside it 
        // has an Actual End Date filled in.
        const isAllDone = tasksInThisMilestone.every(t => {
          const dateVal = t[wActualEndIdx];
          return dateVal && String(dateVal).trim() !== "" && String(dateVal).toUpperCase() !== "N/A";
        });

        if (isAllDone) completedCount++;
      });

      // 4. Update the KPIs with unique counts
      kpis.totalMilestones = uniqueMilestoneNames.length;
      kpis.completedMilestones = completedCount;
      kpis.pendingMilestones = kpis.totalMilestones - completedCount;
    }
  }
  
  // ====================================================================
  // ===== END OF WBS SUMMARY LOGIC =====
  // ====================================================================

  // ====================================================================
  // ===== 10. NEW: BILLING FMS SUMMARY DATA =====
  // ====================================================================
  kpis.totalBillingMilestones = 0;
  kpis.completedBillingMilestones = 0; // NEW value to return
  kpis.overallBillingDonePercent = 0;
  kpis.overallBillingPendingPercent = 0;
  const billingSheet = ss.getSheetByName("Billing");
  if (billingSheet) {
    const bData = billingSheet.getDataRange().getValues();
    if (bData.length > 1) {
      const bHeaders = bData[0].map(h => h.trim());
      const bIsoIdx = bHeaders.indexOf("ISO NO");
      const bMilestoneIdx = bHeaders.indexOf("Milestones");
      const bStatusIdx = bHeaders.indexOf("Status5");

      if (bIsoIdx > -1 && bMilestoneIdx > -1 && bStatusIdx > -1) {
        const projectBillings = bData.slice(1).filter(r => r[bIsoIdx] === isoNo);
        const totalBillings = projectBillings.length;

        if (totalBillings > 0) {
          const milestoneGroups = {};
          projectBillings.forEach(row => {
              const milestoneName = row[bMilestoneIdx];
              if (!milestoneName) return;
              if (!milestoneGroups[milestoneName]) {
                  milestoneGroups[milestoneName] = [];
              }
              milestoneGroups[milestoneName].push(row);
          });

          kpis.totalBillingMilestones = Object.keys(milestoneGroups).length;
          
          let completedMilestoneCount = 0;
          for (const milestoneName in milestoneGroups) {
              const billingsForMilestone = milestoneGroups[milestoneName];
              const allDone = billingsForMilestone.every(b => (b[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE');
              if (allDone) {
                  completedMilestoneCount++;
              }
          }
          kpis.completedBillingMilestones = completedMilestoneCount;

          // Overall percentage calculation remains the same
          const doneBillings = projectBillings.filter(r => (r[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE').length;
          const donePercent = Math.round((doneBillings / totalBillings) * 100);
          
          kpis.overallBillingDonePercent = donePercent;
          kpis.overallBillingPendingPercent = 100 - donePercent;
        }
      }
    }
  }

  // ====================================================================
  // ===== 11. NEW: PROFIT / LOSS CALCULATION =====
  // ====================================================================
  const orderAmount = parseFloat(kpis.orderAmount) || 0;
  const totalCosts = (kpis.manpowerUsedRs || 0) + (kpis.productUsedRs || 0) + (kpis.expensesRs || 0);
  kpis.profitLoss = orderAmount - totalCosts;

  
  // ====================================================================
  // ===== END OF NEW SECTION =====
  // ====================================================================
  
  return kpis;
}

/**
 * Fetches all WBS sheet data for a specific project based on its ISO No.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows for the project.
 */
function getWbsDataForProject(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const wbsSheet = ss.getSheetByName('WBS Sheet');

    if (!detailsSheet || !wbsSheet) {
      return { error: "Required sheet ('Project Details Form' or 'WBS Sheet') not found." };
    }

    // Step 1: Find the ISO No. from the project name
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');

    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) {
      return { error: `Project '${projectName}' not found.` };
    }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) {
      return { error: "ISO No. not found for this project."};
    }

    // Step 2: Get and filter WBS Data using getDisplayValues to preserve formatting
    const wbsData = wbsSheet.getDataRange().getDisplayValues();
    const wbsHeaders = wbsData.shift();
    const wbsIsoIndex = wbsHeaders.indexOf('ISO No.');

    if (wbsIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'WBS Sheet'."};
    }

    const filteredRows = wbsData.filter(row => row[wbsIsoIndex] === isoNo);

    // Step 3: Return the headers and the filtered data
    return { headers: wbsHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getWbsDataForProject: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches all PROJECT PLANNING sheet data for a specific project based on its ISO No.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows for the project.
 */
function getProjectPlanningData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');

    if (!detailsSheet || !planningSheet) {
      return { error: "Required sheet ('Project Details Form' or 'PROJECT PLANNING') not found." };
    }

    // Step 1: Find the ISO No. from the project name
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');

    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) {
      return { error: `Project '${projectName}' not found.` };
    }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) {
      return { error: "ISO No. not found for this project." };
    }

    // Step 2: Get and filter PROJECT PLANNING Data using getDisplayValues
    const planningData = planningSheet.getDataRange().getDisplayValues();
    const planningHeaders = planningData.shift();
    const planningIsoIndex = planningHeaders.indexOf('ISO No.');

    if (planningIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'PROJECT PLANNING' sheet." };
    }

    const filteredRows = planningData.filter(row => row[planningIsoIndex] === isoNo);

    // Step 3: Return the headers and the filtered data
    return { headers: planningHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getProjectPlanningData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches all INVENTORY sheet data for a specific project.
 * It also extracts the URL from the 'PO NO' column if it exists.
 * Note: This function assumes headers are on row 2 of the 'INVENTORY' sheet.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getInventoryData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const inventorySheet = ss.getSheetByName('INVENTORY');

    if (!detailsSheet || !inventorySheet) {
      return { error: "Required sheet ('Project Details Form' or 'INVENTORY') not found." };
    }

    // Step 1: Find the ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');

    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Step 2: Get and filter INVENTORY Data, handling headers on row 2
    if (inventorySheet.getLastRow() < 2) {
      return { error: "Not enough data in 'INVENTORY' sheet." };
    }
    
    const inventoryRange = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, inventorySheet.getLastColumn());
    const displayValues = inventoryRange.getDisplayValues();
    const richTextValues = inventoryRange.getRichTextValues(); // Get rich text to find links
    
    const inventoryHeaders = displayValues.shift();
    richTextValues.shift(); // Also remove header row from rich text values
    
    const inventoryIsoIndex = inventoryHeaders.indexOf('ISO No.');
    const poNoIndex = inventoryHeaders.indexOf('PO NO'); // Find the PO NO column index

    if (inventoryIsoIndex === -1) {
      return { error: "'ISO No.' column not found on row 2 of 'INVENTORY' sheet." };
    }

    const filteredRows = [];
    displayValues.forEach((row, rowIndex) => {
        // Filter by ISO No.
        if (row[inventoryIsoIndex] === isoNo) {
            // If a PO NO column exists, check for a link
            if (poNoIndex !== -1) {
                const poLink = richTextValues[rowIndex][poNoIndex].getLinkUrl();
                if (poLink) {
                    // If a link is found, replace the cell's text with the URL
                    row[poNoIndex] = poLink;
                }
            }
            filteredRows.push(row);
        }
    });

    // Step 3: Return the headers and the modified, filtered data
    return { headers: inventoryHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getInventoryData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches data from the 'Manpower Cost' sheet for a specific project.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getManpowerCostData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const costSheet = ss.getSheetByName('Manpower Cost');

    if (!detailsSheet || !costSheet) {
      return { error: "Required sheet ('Project Details Form' or 'Manpower Cost') not found." };
    }

    // Step 1: Find the ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');

    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Step 2: Get and filter Manpower Cost Data
    const costData = costSheet.getDataRange().getDisplayValues();
    const costHeaders = costData.shift();
    const costIsoIndex = costHeaders.indexOf('ISO No.');

    if (costIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'Manpower Cost' sheet." };
    }

    const filteredRows = costData.filter(row => row[costIsoIndex] === isoNo);

    // Step 3: Return the headers and the filtered data
    return { headers: costHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getManpowerCostData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches data from the 'FINANCE/ACCOUNTS' sheet for a specific project.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getFinanceData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const financeSheet = ss.getSheetByName('FINANCE/ACCOUNTS');

    if (!detailsSheet || !financeSheet) {
      return { error: "Required sheet ('Project Details Form' or 'FINANCE/ACCOUNTS') not found." };
    }

    // Step 1: Find the ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');

    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Step 2: Get and filter FINANCE/ACCOUNTS Data
    const financeData = financeSheet.getDataRange().getDisplayValues();
    const financeHeaders = financeData.shift(); // Assuming headers are on row 1
    const financeIsoIndex = financeHeaders.indexOf('ISO No.');

    if (financeIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'FINANCE/ACCOUNTS' sheet." };
    }

    // Filter rows. It might be one or more rows per project.
    const filteredRows = financeData.filter(row => row[financeIsoIndex] === isoNo);

    // Step 3: Return the headers and the filtered data
    return { headers: financeHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getFinanceData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches the specific row of data from 'Project Details Form' for a given project.
 * It also extracts URLs from cells with hyperlinks.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and the single data row.
 */
function getProjectDetailsFormData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Project Details Form');
    if (!sheet) {
      return { error: "'Project Details Form' sheet not found." };
    }

    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    const richTextValues = range.getRichTextValues(); // Get rich text to find links

    const headers = values.shift();
    richTextValues.shift(); // Also remove header row from rich text values

    const projectColIdx = headers.indexOf('PROJECT NAME');
    if (projectColIdx === -1) {
      return { error: "'PROJECT NAME' column not found." };
    }

    // Find the index of the row for the specified project
    const projectRowIndex = values.findIndex(row => row[projectColIdx] === projectName);

    if (projectRowIndex === -1) {
      return { error: `Project '${projectName}' not found in 'Project Details Form'.` };
    }
    
    const projectRowData = values[projectRowIndex];
    const projectRowRichText = richTextValues[projectRowIndex];

    // List of headers that might contain a hyperlink
    const linkHeaders = ['SITE PICS', 'INFO-SHEET', 'PROJECT DRAWING', 'CUSTOMER PO COPY', 'CLIENT SITE VISITS/MOM'];

    // Iterate through the headers to check for links
    headers.forEach((header, index) => {
      if (linkHeaders.includes(header)) {
        const link = projectRowRichText[index].getLinkUrl();
        if (link) {
          projectRowData[index] = link; // Replace the cell text with the actual URL
        }
      }
    });

    return { headers: headers, row: projectRowData };

  } catch (e) {
    Logger.log('Error in getProjectDetailsFormData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches audit data from 'Auditor Sheet Manpower' for a specific project.
 * Headers are on row 2.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getAuditManpowerData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const auditSheet = ss.getSheetByName('Auditor Sheet Manpower');

    if (!detailsSheet || !auditSheet) {
      return { error: "Required sheet ('Project Details Form' or 'Auditor Sheet Manpower') not found." };
    }

    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Get data from audit sheet (headers on row 2)
    if (auditSheet.getLastRow() < 2) return { headers: [], rows: [] };
    const auditRange = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, auditSheet.getLastColumn());
    const auditData = auditRange.getDisplayValues();
    const auditHeaders = auditData.shift();
    const auditIsoIndex = auditHeaders.indexOf('ISO No.');

    if (auditIsoIndex === -1) {
      return { error: "'ISO No.' column not found on row 2 of 'Auditor Sheet Manpower'." };
    }

    const filteredRows = auditData.filter(row => row[auditIsoIndex] === isoNo);
    return { headers: auditHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getAuditManpowerData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Fetches audit data from 'Auditor Sheet Item' for a specific project.
 * Headers are on row 2.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getAuditItemData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const auditSheet = ss.getSheetByName('Auditor Sheet Item');

    if (!detailsSheet || !auditSheet) {
      return { error: "Required sheet ('Project Details Form' or 'Auditor Sheet Item') not found." };
    }

    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Get data from audit sheet (headers on row 2)
    if (auditSheet.getLastRow() < 2) return { headers: [], rows: [] };
    const auditRange = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, auditSheet.getLastColumn());
    const auditData = auditRange.getDisplayValues();
    const auditHeaders = auditData.shift();
    const auditIsoIndex = auditHeaders.indexOf('ISO No.'); // Using 'ISO No 2' as specified

    if (auditIsoIndex === -1) {
      return { error: "'ISO No.' column not found on row 2 of 'Auditor Sheet Item'." };
    }

    const filteredRows = auditData.filter(row => row[auditIsoIndex] === isoNo);
    return { headers: auditHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getAuditItemData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Fetches audit data from 'Auditor Sheet Drawing' for a specific project.
 * Headers are on row 2. Extracts hyperlinks from specified columns.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getAuditDrawingData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const auditSheet = ss.getSheetByName('Auditor Sheet Drawing');

    if (!detailsSheet || !auditSheet) {
      return { error: "Required sheet ('Project Details Form' or 'Auditor Sheet Drawing') not found." };
    }
    
    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Get data from audit sheet (headers on row 2)
    if (auditSheet.getLastRow() < 2) return { headers: [], rows: [] };
    const auditRange = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, auditSheet.getLastColumn());
    const auditData = auditRange.getDisplayValues();
    const richTextValues = auditRange.getRichTextValues(); // For getting links
    
    const auditHeaders = auditData.shift();
    richTextValues.shift(); // Sync rich text with data
    const auditIsoIndex = auditHeaders.indexOf('ISO No.');
    
    if (auditIsoIndex === -1) {
      return { error: "'ISO No.' column not found on row 2 of 'Auditor Sheet Drawing'." };
    }

    const filteredRows = [];
    const linkHeaders = ['Drawing Attachment', 'Final Drawing'];
    const headerIndicesToGetLinks = linkHeaders.map(h => auditHeaders.indexOf(h));

    auditData.forEach((row, rowIndex) => {
      if (row[auditIsoIndex] === isoNo) {
        // Create a copy of the row to modify
        const newRow = [...row];
        const richTextRow = richTextValues[rowIndex];

        headerIndicesToGetLinks.forEach(headerIndex => {
          if (headerIndex !== -1) {
            const link = richTextRow[headerIndex].getLinkUrl();
            if (link) {
              newRow[headerIndex] = link; // Replace cell text with the actual URL
            }
          }
        });
        filteredRows.push(newRow);
      }
    });

    return { headers: auditHeaders, rows: filteredRows };
  } catch (e) {
    Logger.log('Error in getAuditDrawingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * REPLACED: Fetches a summary of work done by each engineer, summing their total hours.
 * @param {string} projectName The name of the project.
 * @returns {object} An array of objects, each containing an engineer's name and their total worked hours.
 */
function getEngineerWorkSummary(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    if (!detailsSheet || !planningSheet) return { error: "Required sheets not found." };

    // Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get Planning Data
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const doerNameIdx = pHeaders.indexOf("Engineer Name (Doer)");
    const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");

    if (pIsoIdx === -1 || doerNameIdx === -1 || workedHrsIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    const projectPlanningData = pData.filter(r => r[pIsoIdx] === isoNo);
    
    // Use an object to sum hours for each engineer
    const engineerHours = {};
    projectPlanningData.forEach(row => {
        const name = row[doerNameIdx];
        const hours = parseFloat(row[workedHrsIdx]) || 0;
        if (name && hours > 0) {
            if (engineerHours[name]) {
                engineerHours[name] += hours;
            } else {
                engineerHours[name] = hours;
            }
        }
    });
    
    // Convert the object into an array for the frontend
    const summaryData = Object.keys(engineerHours).map(name => {
        return { name: name, hours: engineerHours[name] };
    });
    
    return summaryData.sort((a, b) => b.hours - a.hours); // Sort by hours descending

  } catch (e) {
    Logger.log('Error in getEngineerWorkSummary: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * REPLACED: Fetches data from the 'MOM' sheet for a specific project, using ISO No.
 * @param {string} isoNo The ISO Number of the project.
 * @returns {object} An object containing the headers and filtered rows with URLs.
 */
function getMomData(isoNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const momSheet = ss.getSheetByName('MOM');

    if (!momSheet) {
      return { error: "Required sheet ('MOM') not found." };
    }
    if (!isoNo) {
      return { error: "ISO No. was not provided to the server." };
    }

    // Get MOM Data with both display values and rich text for links
    const momRange = momSheet.getDataRange();
    const displayValues = momRange.getDisplayValues();
    const richTextValues = momRange.getRichTextValues();

    const momHeaders = displayValues.shift().map(h => h.trim());
    richTextValues.shift(); // Keep rich text array in sync

    const momIsoIndex = momHeaders.indexOf('ISO No.');
    if (momIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'MOM' sheet." };
    }

    // Define columns that might contain hyperlinks
    const linkColumns = [
        'Upload Image 1', 'Upload Image 2', 'Upload Image 3', 
        'Upload Image 4', 'Upload Image 5', 'PDF Link', 'Client Signature (Client Attendee 1)',
        'Mentor Signature (Mentor Attendee 1)'
    ];
    const linkColumnIndices = linkColumns.map(h => momHeaders.indexOf(h));

    const filteredRows = [];
    displayValues.forEach((row, rowIndex) => {
        // Filter by the project's ISO No.
        if (row[momIsoIndex] === isoNo) {
            // Check for links in the designated columns for this row
            linkColumnIndices.forEach(colIndex => {
                if (colIndex !== -1) {
                    const link = richTextValues[rowIndex][colIndex].getLinkUrl();
                    if (link) {
                        row[colIndex] = link;
                    }
                }
            });
            filteredRows.push(row);
        }
    });

    return { headers: momHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getMomData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches the detailed task counts (with combined pending) for each milestone of a project.
 * Used to populate the milestone details modal.
 * @param {string} projectName The name of the project.
 * @returns {object} An array of objects, each with milestone name and task counts.
 */
function getMilestoneDetailsData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');

    if (!detailsSheet || !planningSheet) {
      return { error: "Required sheets not found." };
    }

    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) return { error: "ISO No. not found for this project." };
    
    
    // Get planning data and calculate detailed stats for each milestone
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const milestoneIdx = pHeaders.indexOf("Milestone");
    const statusIdx = pHeaders.indexOf("Status");

    if (pIsoIdx === -1 || milestoneIdx === -1 || statusIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);
    const milestoneStats = {};
    const milestoneStatusCounts = [];

    projectTasks.forEach(task => {
      const milestoneName = task[milestoneIdx];
      if (!milestoneName) return;

      if (!milestoneStats[milestoneName]) {
        milestoneStats[milestoneName] = { done: 0, inProgress: 0, hold: 0 };
      }
      
      const status = (task[statusIdx] || '').toString().toUpperCase().trim();
      switch (status) {
        case 'DONE': milestoneStats[milestoneName].done++; break;
        case 'IN-PROGRESS': milestoneStats[milestoneName].inProgress++; break;
        case 'HOLD': milestoneStats[milestoneName].hold++; break;
      }
    });

    for (const milestoneName in milestoneStats) {
      const stats = milestoneStats[milestoneName];
      const total = stats.done + stats.inProgress + stats.hold;
      const pending = stats.inProgress + stats.hold; // Combine them here
      milestoneStatusCounts.push({
        milestone: milestoneName,
        total: total,
        done: stats.done,
        pending: pending // Return the combined value
      });
    }
    
    return milestoneStatusCounts;

  } catch (e) {
    Logger.log('Error in getMilestoneDetailsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches product consumption data from the 'INVENTORY' sheet for a project.
 * This is specifically for the 'Product Used (Rs.)' modal. Headers are on row 2.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getProductUsedData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const inventorySheet = ss.getSheetByName('INVENTORY');

    if (!detailsSheet || !inventorySheet) {
      return { error: "Required sheets ('Project Details Form' or 'INVENTORY') not found." };
    }

    // Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) { return { error: `Project '${projectName}' not found.` }; }
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) { return { error: "ISO No. not found for this project." }; }

    // Get Inventory Data (headers on row 2)
    if (inventorySheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }
    const inventoryRange = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, inventorySheet.getLastColumn());
    const inventoryData = inventoryRange.getDisplayValues();
    const inventoryHeaders = inventoryData.shift();
    const inventoryIsoIndex = inventoryHeaders.indexOf('ISO No.');

    if (inventoryIsoIndex === -1) {
      return { error: "'ISO No.' column not found on row 2 of 'INVENTORY' sheet." };
    }

    // Filter rows for the project AND rows where the consumption cost ('CONSU.2') has a value > 0
    const consu2Index = inventoryHeaders.indexOf('CONSU.2');
    const filteredRows = inventoryData.filter(row => {
        const isProjectMatch = row[inventoryIsoIndex] === isoNo;
        const hasConsumption = consu2Index !== -1 && parseFloat(row[consu2Index]) > 0;
        return isProjectMatch && hasConsumption;
    });

    return { headers: inventoryHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getProductUsedData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * REPLACED FUNCTION: Fetches all data from the 'Expenses' sheet for a project.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getExpenseDetailsData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const expenseSheet = ss.getSheetByName('Expenses');

    if (!detailsSheet || !expenseSheet) {
      return { error: "Required sheet ('Project Details Form' or 'Expenses') not found." };
    }

    // Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get and filter data from the 'Expenses' sheet
    const expenseData = expenseSheet.getDataRange().getDisplayValues();
    const expenseHeaders = expenseData.shift(); // Headers on row 1
    const expenseIsoIndex = expenseHeaders.indexOf('ISO No.');

    if (expenseIsoIndex === -1) {
      return { error: "'ISO No.' column not found in 'Expenses' sheet." };
    }

    const filteredRows = expenseData.filter(row => row[expenseIsoIndex] === isoNo);
    
    return { headers: expenseHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getExpenseDetailsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

// /**
//  * REPLACED: Fetches tasks from 'PROJECT PLANNING' filtered by status.
//  * @param {string} projectName The name of the project.
//  * @param {string} filterType The status to filter by ('DONE', 'IN-PROGRESS', or 'HOLD').
//  * @returns {object} An object containing the headers and filtered rows.
//  */
// function getFilteredTasks(projectName, filterType) {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const detailsSheet = ss.getSheetByName('Project Details Form');
//     const planningSheet = ss.getSheetByName('PROJECT PLANNING');
//     if (!detailsSheet || !planningSheet) return { error: "Required sheets not found." };

//     // --- THIS IS THE CORRECTED BLOCK ---
//     // Using the standard, reliable method to get the ISO No.
//     const detailsData = detailsSheet.getDataRange().getValues();
//     const detailsHeaders = detailsData.shift();
//     const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
//     const isoColIdx = detailsHeaders.indexOf('ISO No.');
//     const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
//     if (!projectRow) return { error: `Project '${projectName}' not found.` };
//     // const isoNo = projectRow[isoColIdx];
//     const isoNo = cleanIso(projectRow[isoColIdx]);
//     if (!isoNo) return { error: "ISO No. not found for this project." };
//     // --- END OF CORRECTION ---

//     // Get and filter PROJECT PLANNING Data
//     const pData = planningSheet.getDataRange().getDisplayValues();
//     const pHeaders = pData.shift();
//     const pIsoIdx = pHeaders.indexOf("ISO No.");
//     const statusIdx = pHeaders.indexOf("Status");

//     if (pIsoIdx === -1 || statusIdx === -1) {
//       return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
//     }

//     const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);
//     let filteredRows = [];

//     // Use the filterType directly (this part was correct)
//     if (['DONE', 'IN-PROGRESS', 'HOLD'].includes(filterType)) {
//       filteredRows = projectTasks.filter(task => (task[statusIdx] || '').toString().toUpperCase().trim() === filterType);
//     }

//     return { headers: pHeaders, rows: filteredRows };

//   } catch (e) {
//     Logger.log('Error in getFilteredTasks: ' + e.toString());
//     return { error: 'An error occurred on the server: ' + e.message };
//   }
// }

/**
 * REPLACED: Fetches tasks from 'PROJECT PLANNING' filtered by status.
 * @param {string} projectName The name of the project.
 * @param {string} filterType The status to filter by ('DONE', 'IN-PROGRESS', or 'HOLD').
 * @returns {object} An object containing the headers and filtered rows.
 */
function getFilteredTasks(projectName, filterType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    if (!detailsSheet || !planningSheet) return { error: "Required sheets not found." };

    // Using the standard, reliable method to get the ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get and filter PROJECT PLANNING Data
    const pData = planningSheet.getDataRange().getDisplayValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const statusIdx = pHeaders.indexOf("Status");

    if (pIsoIdx === -1 || statusIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);
    let filteredRows =[];

    // --- UPDATED LOGIC TO MATCH KPI DASHBOARD MATH ---
    if (['DONE', 'IN-PROGRESS', 'HOLD'].includes(filterType)) {
      filteredRows = projectTasks.filter(task => {
        const statusVal = (task[statusIdx] || '').toString().toUpperCase().trim();
        
        if (filterType === 'HOLD') {
          // KPI Logic: Anything NOT Done and NOT In-Progress is considered HOLD (including blanks)
          return statusVal !== 'DONE' && statusVal !== 'IN-PROGRESS' && statusVal !== 'IN PROGRESS';
        } else if (filterType === 'IN-PROGRESS') {
          // Handle both variations of In-Progress
          return statusVal === 'IN-PROGRESS' || statusVal === 'IN PROGRESS';
        } else {
          // Exact match for DONE
          return statusVal === filterType;
        }
      });
    }

    return { headers: pHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getFilteredTasks: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches tasks from 'PROJECT PLANNING' that have planned hours assigned.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getPlannedHrsData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    if (!detailsSheet || !planningSheet) return { error: "Required sheets not found." };

    // Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get and filter PROJECT PLANNING Data
    const pData = planningSheet.getDataRange().getDisplayValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const reqHrsIdx = pHeaders.indexOf("Required Hours");

    if (pIsoIdx === -1 || reqHrsIdx === -1) return { error: "Required columns not found." };

    const filteredRows = pData.filter(row => {
        const isProjectMatch = row[pIsoIdx] === isoNo;
        const hasPlannedHours = parseFloat(row[reqHrsIdx]) > 0;
        return isProjectMatch && hasPlannedHours;
    });

    return { headers: pHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getPlannedHrsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches tasks from 'PROJECT PLANNING' that have actual hours logged.
 * @param {string} projectName The name of the project.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getActualHrsData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    if (!detailsSheet || !planningSheet) return { error: "Required sheets not found." };
    
    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = projectRow[isoColIdx];
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get and filter PROJECT PLANNING Data
    const pData = planningSheet.getDataRange().getDisplayValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");

    if (pIsoIdx === -1 || workedHrsIdx === -1) return { error: "Required columns not found." };

    const filteredRows = pData.filter(row => {
        const isProjectMatch = row[pIsoIdx] === isoNo;
        const hasActualHours = parseFloat(row[workedHrsIdx]) > 0;
        return isProjectMatch && hasActualHours;
    });

    return { headers: pHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getActualHrsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * REPLACED: Fetches milestone names from the WBS Sheet filtered by ACTUAL END DATE.
 * @param {string} projectName The name of the project.
 * @param {string} filterType The status to filter by ('COMPLETED' or 'PENDING').
 */
function getFilteredMilestones(projectName, filterType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const wbsSheet = ss.getSheetByName('WBS Sheet');
    if (!detailsSheet || !wbsSheet) return { error: "Required sheets not found." };

    // 1. Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };

    // const isoNo = String(projectRow[isoColIdx]).trim();
    // --- UPDATED LINE ---
    const isoNo = cleanIso(projectRow[isoColIdx]);

    // 2. Get WBS Data
    const wbsData = wbsSheet.getDataRange().getValues();
    const wHeaders = wbsData[0].map(h => String(h).trim());
    const wIsoIdx = wHeaders.indexOf("ISO No.");
    const wMileIdx = wHeaders.indexOf("MILESTONE/BOQ ITEM");
    const wActualEndIdx = wHeaders.indexOf("ACTUAL END DATE");

    if (wIsoIdx === -1 || wMileIdx === -1 || wActualEndIdx === -1) {
      return { error: "Required columns not found in 'WBS Sheet'." };
    }

    // 3. Filter rows for this project and Group by Milestone Name
    const projWbsRows = wbsData.slice(1).filter(r => String(r[wIsoIdx]).trim() === isoNo);
    const milestoneGroups = {};
    
    projWbsRows.forEach(row => {
      const mName = String(row[wMileIdx]).trim();
      if (!mName || mName === "N/A" || mName === "") return;
      if (!milestoneGroups[mName]) milestoneGroups[mName] = [];
      milestoneGroups[mName].push(row);
    });

    const resultMilestones = [];

    // 4. Evaluate each group based on ACTUAL END DATE
    for (const milestoneName in milestoneGroups) {
      const tasks = milestoneGroups[milestoneName];
      
      // A milestone is COMPLETED only if EVERY row in it has an Actual End Date
      const isComplete = tasks.every(t => {
        const dateVal = t[wActualEndIdx];
        return dateVal && String(dateVal).trim() !== "" && String(dateVal).toUpperCase() !== "N/A";
      });
      
      if (filterType === 'COMPLETED' && isComplete) {
        resultMilestones.push(milestoneName);
      } else if (filterType === 'PENDING' && !isComplete) {
        resultMilestones.push(milestoneName);
      }
    }

    return resultMilestones.sort();

  } catch (e) {
    Logger.log('Error in getFilteredMilestones: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Fetches data from the 'INVENTORY' sheet, filtered to rows where a specific column has a value > 0.
 * @param {string} projectName The name of the project.
 * @param {string} filterColumn The name of the column to filter on (e.g., 'Qty', 'INDENT1').
 * @returns {object} An object containing the headers and filtered rows.
 */
function getFilteredInventoryData(projectName, filterColumn) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const inventorySheet = ss.getSheetByName('INVENTORY');
    if (!detailsSheet || !inventorySheet) return { error: "Required sheets not found." };

    // Get ISO No. for the project
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get Inventory Data (headers on row 2)
    if (inventorySheet.getLastRow() < 2) return { headers: [], rows: [] };
    const inventoryRange = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, inventorySheet.getLastColumn());
    const displayValues = inventoryRange.getDisplayValues();
    const inventoryHeaders = displayValues.shift();
    
    const inventoryIsoIndex = inventoryHeaders.indexOf('ISO No.');
    const filterColumnIndex = inventoryHeaders.indexOf(filterColumn);

    if (inventoryIsoIndex === -1 || filterColumnIndex === -1) {
      return { error: `Required column ('ISO No.' or '${filterColumn}') not found in 'INVENTORY' sheet.` };
    }

    // Filter rows for the project AND where the specified filter column has a value > 0
    const filteredRows = displayValues.filter(row => {
        const isProjectMatch = cleanIso(row[inventoryIsoIndex]) === isoNo;
        const hasValue = parseFloat(String(row[filterColumnIndex]).replace(/,/g, '')) > 0;
        return isProjectMatch && hasValue;
    });

    return { headers: inventoryHeaders, rows: filteredRows };

  } catch(e) {
    Logger.log('Error in getFilteredInventoryData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

function getBillingData(projectName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const billingSheet = ss.getSheetByName('Billing');
    if (!detailsSheet || !billingSheet) return { error: "Required sheets not found." };

    // 1. Get Clean ISO
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    
    // Uses the helper function you added earlier
    const isoNo = cleanIso(projectRow[isoColIdx]); 

    // 2. Get Billing Headers
    const bRange = billingSheet.getDataRange();
    const displayValues = bRange.getDisplayValues();
    const richTextValues = bRange.getRichTextValues();
    
    const bHeaders = displayValues.shift().map(h => h.trim());
    richTextValues.shift();

    // Robust Header Search
    let bIsoIdx = bHeaders.indexOf("ISO NO");
    if (bIsoIdx === -1) bIsoIdx = bHeaders.indexOf("ISO No.");
    const suppDocIdx = bHeaders.indexOf("Supp Doc");

    if (bIsoIdx === -1) return { error: "'ISO NO' column not found in 'Billing' sheet." };

    // 3. Filter Rows using cleanIso
    const filteredRows = [];
    displayValues.forEach((row, rowIndex) => {
      // Clean the sheet data before comparing
      if (cleanIso(row[bIsoIdx]) === isoNo) {
        if (suppDocIdx !== -1) {
          const link = richTextValues[rowIndex][suppDocIdx].getLinkUrl();
          if (link) {
            row[suppDocIdx] = link;
          }
        }
        filteredRows.push(row);
      }
    });

    return { headers: bHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getBillingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * REPLACED: Fetches and prepares data for the task progress timeline chart, using ISO No.
 * This version includes the most robust date parsing to prevent errors.
 */
function getTaskStatusChartData( isoNo = 'MWEPL/2024-25/01250') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    if (!planningSheet || !isoNo) {
      Logger.log("getTaskStatusChartData: Sheet not found or no ISO No provided.");
      return []; // Always return an array
    }

    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift().map(h => h.trim());
    
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const taskIdx = pHeaders.indexOf("Task");
    const statusPercentIdx = pHeaders.indexOf("Status %");
    const plannedDateIdx = pHeaders.indexOf("Planned Date");

    if (pIsoIdx === -1 || taskIdx === -1 || statusPercentIdx === -1) {
      Logger.log("getTaskStatusChartData: A required column was not found.");
      return []; // Always return an array
    }

    const chartData = pData
      .filter(row => row[pIsoIdx] === isoNo && row[taskIdx])
      .map(row => {
        const status = parseFloat(row[statusPercentIdx]) || 0;
        let plannedDateValue = (plannedDateIdx > -1) ? row[plannedDateIdx] : null;
        
        let jsDate = null;
        if (plannedDateValue instanceof Date && !isNaN(plannedDateValue)) {
            jsDate = plannedDateValue;
        } else if (typeof plannedDateValue === 'string') {
            const parts = plannedDateValue.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
            if (parts) {
                jsDate = new Date(parts[3], parts[2] - 1, parts[1]);
            }
        }
        
        const formattedDate = (jsDate && !isNaN(jsDate)) 
            ? Utilities.formatDate(jsDate, Session.getScriptTimeZone(), "dd-MM-yyyy") 
            : "N/A";

        return {
          milestone: row[taskIdx],
          planned: status,
          actual: formattedDate,
          rawDate: jsDate // Will be a valid Date object or null
        };
      });

    chartData.sort((a, b) => {
        if (!a.rawDate) return 1;
        if (!b.rawDate) return -1;
        return a.rawDate - b.rawDate;
    });

    Logger.log(chartData)
    
    return chartData; // Ensure it returns the final array

  } catch (e) {
    Logger.log('FATAL Error in getTaskStatusChartData: ' + e.toString());
    return []; // ALWAYS return an array, even on fatal error
  }
}

/**
 * Fetches data from the 'Billing' sheet, filtered by status ('DONE' or 'PENDING').
 * @param {string} projectName The name of the project.
 * @param {string} filterType The status to filter by.
 * @returns {object} An object containing the headers and filtered rows.
 */
function getFilteredBillingData(projectName, filterType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName('Project Details Form');
    const billingSheet = ss.getSheetByName('Billing');
    if (!detailsSheet || !billingSheet) return { error: "Required sheets not found." };

    // Get ISO No.
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projectColIdx = detailsHeaders.indexOf('PROJECT NAME');
    const isoColIdx = detailsHeaders.indexOf('ISO No.');
    const projectRow = detailsData.find(row => row[projectColIdx] === projectName);
    if (!projectRow) return { error: `Project '${projectName}' not found.` };
    // const isoNo = projectRow[isoColIdx];
    const isoNo = cleanIso(projectRow[isoColIdx]);
    if (!isoNo) return { error: "ISO No. not found for this project." };

    // Get Billing Data
    const bData = billingSheet.getDataRange().getDisplayValues();
    const bHeaders = bData.shift().map(h => h.trim());
    const bIsoIdx = bHeaders.indexOf("ISO NO");
    const bStatusIdx = bHeaders.indexOf("Status5");
    if (bIsoIdx === -1 || bStatusIdx === -1) return { error: "Required columns not found in 'Billing' sheet." };
    
    const projectBillings = bData.filter(r => r[bIsoIdx] === isoNo);
    let filteredRows = [];

    if (filterType === 'DONE') {
      filteredRows = projectBillings.filter(r => (r[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE');
    } else if (filterType === 'PENDING') {
      filteredRows = projectBillings.filter(r => (r[bStatusIdx] || '').toString().trim().toUpperCase() !== 'DONE');
    }

    return { headers: bHeaders, rows: filteredRows };

  } catch (e) {
    Logger.log('Error in getFilteredBillingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * REPLACED: Fetches and combines data for the Milestone Progress Combo Chart.
 * This version correctly filters out undefined or null milestone names.
 * @param {string} isoNo The ISO Number of the project.
 * @returns {object[]} An array of objects formatted for the D3 combo chart.
 */
function getMilestoneComboChartData(isoNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName('PROJECT PLANNING');
    const billingSheet = ss.getSheetByName('Billing');
    if (!planningSheet || !billingSheet || !isoNo) return [];

    // --- 1. Get all data for the project ---
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift().map(h => h.trim());
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const pMilestoneIdx = pHeaders.indexOf("Milestone");
    const pStatusPercentIdx = pHeaders.indexOf("Status %");
    const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);

    const bData = billingSheet.getDataRange().getValues();
    const bHeaders = bData.shift().map(h => h.trim());
    const bIsoIdx = bHeaders.indexOf("ISO NO");
    const bMilestoneIdx = bHeaders.indexOf("Milestones");
    const bStatusIdx = bHeaders.indexOf("Status5");
    const projectBillings = bData.filter(r => r[bIsoIdx] === isoNo);

    // --- 2. Create a unique, combined list of all milestones (CORRECTED) ---
    const taskMilestones = projectTasks.map(r => r[pMilestoneIdx]);
    const billingMilestones = projectBillings.map(r => r[bMilestoneIdx]);
    // The .filter(Boolean) will remove any null, undefined, or empty strings.
    const allMilestones = [...new Set([...taskMilestones, ...billingMilestones])].filter(Boolean);

    // --- 3. Calculate stats for each milestone ---
    const chartData = allMilestones.map(milestoneName => {
      // a) Calculate Task Progress %
      const tasksForMilestone = projectTasks.filter(r => r[pMilestoneIdx] === milestoneName);
      let taskProgress = 0;
      if (tasksForMilestone.length > 0) {
        const totalPercent = tasksForMilestone.reduce((sum, task) => sum + (parseFloat(task[pStatusPercentIdx]) || 0), 0);
        taskProgress = totalPercent / tasksForMilestone.length;
      }

      // b) Calculate Billing Status %
      const billingsForMilestone = projectBillings.filter(r => r[bMilestoneIdx] === milestoneName);
      let billingStatus = 0;
      if (billingsForMilestone.length > 0) {
        const doneCount = billingsForMilestone.filter(b => (b[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE').length;
        billingStatus = (doneCount / billingsForMilestone.length) * 100;
      }
      
      return {
        milestone: milestoneName,
        taskProgress: taskProgress,
        billingStatus: billingStatus
      };
    });

    return chartData;

  } catch (e) {
    Logger.log('Error in getMilestoneComboChartData: ' + e.toString());
    return [];
  }
}

/**
 * Helper to clean the ISO string.
 * Converts "MWEPL/2025-26/02641 - NATUROVILLE HOSPITALITY" -> "MWEPL/2025-26/02641"
 */
function cleanIso(rawIso) {
  if (!rawIso) return "";
  // Convert to string, split by " - " (space hyphen space), take the first part, and trim whitespace
  return String(rawIso).split(' - ')[0].trim();
}


function debugHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['INVENTORY',2],['Auditor Sheet Item',2],['Auditor Sheet Drawing',2],['WBS Sheet',1]].forEach(([n,r]) => {
    const s = ss.getSheetByName(n);
    if (!s) return Logger.log(n + ': SHEET NOT FOUND');
    Logger.log(n + ' headers: ' + JSON.stringify(s.getRange(r,1,1,s.getLastColumn()).getValues()[0]));
    Logger.log(n + ' sample row: ' + JSON.stringify(s.getRange(r+1,1,1,s.getLastColumn()).getValues()[0]));
  });
}