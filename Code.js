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
    // 1. BATCH FETCH for KPIs (Values only is fine for math)
    const sheetNames = [
      'Project Details Form', 'PROJECT PLANNING', 'Manpower Cost', 
      'INVENTORY', 'Auditor Sheet Manpower', 'Auditor Sheet Item', 
      'Auditor Sheet Drawing', 'FINANCE/ACCOUNTS', 'Billing', 'WBS Sheet', 'Expenses'
    ];

    // ONE Sheets API call for all tabs (instead of 11 SpreadsheetApp reads)
    const allData = fetchSheetsBatch_(getSsId_(), sheetNames);

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
          const momGrid = fetchMomWithLinks_(getSsId_());
          if (momGrid && momGrid.displayValues.length) {
             const displayValues = momGrid.displayValues;
             const linkValues = momGrid.links;

             const momHeaders = displayValues.shift().map(h => String(h).trim());
             linkValues.shift();

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
                       const link = linkValues[rowIndex][colIndex];
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
 * Reads many tabs in ONE Sheets API (v4) call.
 * Returns { sheetName: 2D array } shaped like getDataRange().getValues()
 * (rows padded to equal width; missing tabs -> []). Dates come back as serial numbers.
 */
function fetchSheetsBatch_(ssId, sheetNames) {
  // readTabs_ already pads rows and retries tab-by-tab if one tab is missing,
  // so there is no need to open the spreadsheet (slow) just to list its tabs.
  const t = readTabs_(ssId, sheetNames.map(n => ({ name: n })), { unformatted: true });
  const out = {};
  sheetNames.forEach(n => { out[n] = t[n] ? t[n].values : []; });
  return out;
}

/**
 * Reads the MOM tab (display text + hyperlinks) in one Sheets API call.
 * Returns { displayValues: string[][], links: (string|null)[][] } or null.
 */
function fetchMomWithLinks_(ssId) {
  try {
    const t = readTabs_(ssId, [{ name: 'MOM' }], { links: true });
    if (!t['MOM']) return null;
    return { displayValues: t['MOM'].values, links: t['MOM'].links };
  } catch (e) {
    Logger.log('fetchMomWithLinks_ error: ' + e);
    return null; // sheet missing -> no MOM data
  }
}

/** Sheets serial date -> "dd/MM/yyyy" (null for blank/0/non-numeric). */
function serialToDdMmYyyy_(v) {
  if (v instanceof Date) {
    return v.getTime() > 0 ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : null;
  }
  if (typeof v !== 'number' || v <= 0) return null;
  const ms = Math.round((v - 25569) * 86400 * 1000); // days since 1899-12-30 -> epoch
  return Utilities.formatDate(new Date(ms), 'UTC', 'dd/MM/yyyy');
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
      // Sheets API returns dates as serial numbers (UNFORMATTED_VALUE)
      const fPDate = serialToDdMmYyyy_(pDate);
      const fADate = serialToDdMmYyyy_(aDate);

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
 * ====================================================================
 * SHEETS API HELPERS (shared by every dashboard data function)
 * ====================================================================
 */

/**
 * Spreadsheet ID without opening the (large) spreadsheet on every call.
 * Cached per script for 6 hours; falls back to the active spreadsheet.
 */
function getSsId_() {
  const cache = CacheService.getScriptCache();
  let id = cache.get('SS_ID');
  if (!id) {
    id = SpreadsheetApp.getActiveSpreadsheet().getId();
    cache.put('SS_ID', id, 21600);
  }
  return id;
}

/** 1 -> "A", 27 -> "AA" */
function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Per-execution memo of prefetched tabs (null = off). Reset on every Apps Script run. */
var TAB_MEMO_ = null;
function memoKey_(s, opt) {
  return s.name + '|' + (s.startRow || 1) + '|' + (opt && opt.unformatted ? 'u' : 'f');
}

/**
 * Prefetches the given tabs (3 Sheets API calls max) into TAB_MEMO_ so the
 * get...Data functions of one bundle part run without touching the spreadsheet again.
 * Each list item is {name, startRow?}.
 */
function prefetchTabs_(plain, withLinks, raw) {
  const ssId = getSsId_();
  const memo = {};
  if (plain.length) {
    const f = readTabs_(ssId, plain, {});
    plain.forEach(s => { if (f[s.name]) memo[memoKey_(s, {})] = f[s.name]; });
  }
  if (withLinks.length) {
    const l = readTabs_(ssId, withLinks, { links: true });
    withLinks.forEach(s => { if (l[s.name]) memo[memoKey_(s, {})] = l[s.name]; });
  }
  if (raw.length) {
    const u = readTabs_(ssId, raw, { unformatted: true });
    raw.forEach(s => { if (u[s.name]) memo[memoKey_(s, { unformatted: true })] = u[s.name]; });
  }
  TAB_MEMO_ = memo;
}

/**
 * Loads ONE part of the dashboard as JSON. The UI fires all parts in parallel, shows each
 * section as soon as its part returns, and drives a progress bar from the finished parts.
 *   part: 'dashboard' | 'planning' | 'wbs' | 'inventory' | 'billing' | 'other'
 * Returns { part, data } where data = { "<functionName>|<JSON args>": result }.
 * (For 'dashboard', data = the getAllDashboardData result.)
 * Only the selected project's rows are returned.
 */
function getBundlePart(projectName, part) {
  const t0 = Date.now();
  const res = getBundlePartInner_(projectName, part);
  res.ms = Date.now() - t0; // server time for this part (logged in the browser console)
  Logger.log('getBundlePart ' + part + ' took ' + res.ms + ' ms');
  return res;
}

function getBundlePartInner_(projectName, part) {
  const p = projectName;
  const D = { name: 'Project Details Form' };
  try {
    let calls = [];
    if (part === 'dashboard') {
      return { part: part, data: getAllDashboardData(projectName) };
    } else if (part === 'planning') {
      prefetchTabs_([D, { name: 'PROJECT PLANNING' }], [], [D, { name: 'PROJECT PLANNING' }]);
      calls = [
        [getProjectPlanningData, 'getProjectPlanningData', [p]],
        [getFilteredTasks, 'getFilteredTasks', [p, 'DONE']],
        [getFilteredTasks, 'getFilteredTasks', [p, 'IN-PROGRESS']],
        [getFilteredTasks, 'getFilteredTasks', [p, 'HOLD']],
        [getEngineerWorkSummary, 'getEngineerWorkSummary', [p]],
        [getPlannedHrsData, 'getPlannedHrsData', [p]],
        [getActualHrsData, 'getActualHrsData', [p]]
      ];
    } else if (part === 'wbs') {
      prefetchTabs_([D, { name: 'WBS Sheet' }], [], [D, { name: 'WBS Sheet' }]);
      calls = [
        [getWbsDataForProject, 'getWbsDataForProject', [p]],
        [getFilteredMilestones, 'getFilteredMilestones', [p, 'COMPLETED']],
        [getFilteredMilestones, 'getFilteredMilestones', [p, 'PENDING']]
      ];
    } else if (part === 'inventory') {
      prefetchTabs_([], [D, { name: 'INVENTORY', startRow: 2 }], []);
      calls = [
        [getInventoryData, 'getInventoryData', [p]],
        [getProductUsedData, 'getProductUsedData', [p]],
        [getFilteredInventoryData, 'getFilteredInventoryData', [p, 'Planned Indent Qty']],
        [getFilteredInventoryData, 'getFilteredInventoryData', [p, 'INDENT1']],
        [getFilteredInventoryData, 'getFilteredInventoryData', [p, 'REC.1']],
        [getFilteredInventoryData, 'getFilteredInventoryData', [p, 'CONSU.1']]
      ];
    } else if (part === 'billing') {
      prefetchTabs_([], [D, { name: 'Billing' }], []);
      calls = [
        [getBillingData, 'getBillingData', [p]],
        [getFilteredBillingData, 'getFilteredBillingData', [p, 'DONE']],
        [getFilteredBillingData, 'getFilteredBillingData', [p, 'PENDING']]
      ];
    } else if (part === 'other') {
      prefetchTabs_(
        [{ name: 'Manpower Cost' }, { name: 'FINANCE/ACCOUNTS' }, { name: 'Expenses' },
         { name: 'Auditor Sheet Manpower', startRow: 2 }, { name: 'Auditor Sheet Item', startRow: 2 }],
        [D, { name: 'Auditor Sheet Drawing', startRow: 2 }],
        []
      );
      // Manpower/Finance/Expenses/Audit tabs also need the details tab in plain form (links version covers it)
      calls = [
        [getManpowerCostData, 'getManpowerCostData', [p]],
        [getFinanceData, 'getFinanceData', [p]],
        [getProjectDetailsFormData, 'getProjectDetailsFormData', [p]],
        [getAuditManpowerData, 'getAuditManpowerData', [p]],
        [getAuditItemData, 'getAuditItemData', [p]],
        [getAuditDrawingData, 'getAuditDrawingData', [p]],
        [getExpenseDetailsData, 'getExpenseDetailsData', [p]]
      ];
    } else {
      return { part: part, error: 'Unknown part: ' + part };
    }

    const data = {};
    calls.forEach(c => { data[c[1] + '|' + JSON.stringify(c[2])] = c[0].apply(null, c[2]); });
    return { part: part, data: data };
  } catch (e) {
    Logger.log('Error in getBundlePart(' + part + '): ' + e.toString());
    return { part: part, error: e.toString() };
  } finally {
    TAB_MEMO_ = null;
  }
}

/**
 * Reads several tabs in ONE Sheets API call.
 * @param {string} ssId Spreadsheet ID.
 * @param {{name:string,startRow?:number}[]} specs Tabs to read (startRow = header row, default 1).
 * @param {{unformatted?:boolean, links?:boolean}} opt
 *   unformatted: raw values (numbers/date serials) like getValues(); default = displayed text like getDisplayValues().
 *   links: also return hyperlinks per cell (uses Spreadsheets.get; values are always displayed text).
 * @returns {Object} { tabName: { values: any[][], links?: (string|null)[][] } } (missing tabs are absent)
 */
function readTabs_(ssId, specs, opt) {
  opt = opt || {};

  // Serve from the per-request memo when getBundlePart has prefetched the tabs.
  if (TAB_MEMO_) {
    const hit = {};
    const all = specs.every(s => {
      const e = TAB_MEMO_[memoKey_(s, opt)];
      if (e && (!opt.links || e.links)) { hit[s.name] = e; return true; }
      return false;
    });
    if (all) return hit;
  }

  const out = {};
  const rangeOf = s => "'" + s.name.replace(/'/g, "''") + "'" + ((s.startRow || 1) > 1 ? '!A' + s.startRow + ':ZZ' : '');
  const width = rows => rows.reduce((m, r) => Math.max(m, r.length), 0);

  try {
    if (opt.links) {
      // Step 1: cheap values read (empty grid cells are trimmed) to learn each tab's real data extent.
      const vres = Sheets.Spreadsheets.Values.batchGet(ssId, {
        ranges: specs.map(rangeOf),
        valueRenderOption: 'FORMATTED_VALUE',
        majorDimension: 'ROWS'
      });
      const vals = {};
      const tight = [];
      (vres.valueRanges || []).forEach((vr, i) => {
        const s = specs[i];
        const rows = vr.values || [];
        const w = width(rows);
        vals[s.name] = rows.map(r => { const c = r.slice(); while (c.length < w) c.push(''); return c; });
        if (rows.length && w) {
          const start = s.startRow || 1;
          tight.push("'" + s.name.replace(/'/g, "''") + "'!A" + start + ':' + colLetter_(w) + (start + rows.length - 1));
        }
      });

      // Step 2: hyperlinks only for the cells that actually hold data.
      const linkGrid = {};
      if (tight.length) {
        const res = Sheets.Spreadsheets.get(ssId, {
          ranges: tight,
          fields: 'sheets(properties.title,data.rowData.values.hyperlink)'
        });
        (res.sheets || []).forEach(sh => {
          linkGrid[sh.properties.title] = ((sh.data || [])[0] || {}).rowData || [];
        });
      }

      specs.forEach(s => {
        if (!vals[s.name]) return;
        const values = vals[s.name];
        const rd = linkGrid[s.name] || [];
        const links = values.map((row, r) => {
          const cells = (rd[r] && rd[r].values) || [];
          return row.map((_, c) => (cells[c] && cells[c].hyperlink) || null);
        });
        out[s.name] = { values: values, links: links };
      });
    } else {
      const res = Sheets.Spreadsheets.Values.batchGet(ssId, {
        ranges: specs.map(rangeOf),
        valueRenderOption: opt.unformatted ? 'UNFORMATTED_VALUE' : 'FORMATTED_VALUE',
        majorDimension: 'ROWS'
      });
      (res.valueRanges || []).forEach((vr, i) => {
        const rows = vr.values || [];
        const w = width(rows);
        out[specs[i].name] = {
          values: rows.map(r => { const c = r.slice(); while (c.length < w) c.push(''); return c; })
        };
      });
    }
  } catch (e) {
    // A missing tab makes the whole batch fail -> retry tab by tab so the others still load.
    Logger.log('readTabs_ batch failed (' + e + ')');
    if (specs.length > 1) {
      specs.forEach(s => Object.assign(out, readTabs_(ssId, [s], opt)));
    }
  }
  return out;
}

/**
 * Looks up a project's ISO No. in 'Project Details Form' and loads one data tab,
 * both in a single Sheets API call.
 * @param {string} projectName
 * @param {string} tab Data tab name.
 * @param {{headerRow?:number, unformatted?:boolean, links?:boolean, rawIso?:boolean, allowNoIso?:boolean}} o
 *   headerRow: row holding the headers in `tab` (default 1). rawIso: don't run cleanIso on the ISO.
 * @returns {{error:string}|{isoNo:string, headers:any[], rows:any[][], links:(string|null)[][]|null}}
 */
function loadProjectTab_(projectName, tab, o) {
  o = o || {};
  const DET = 'Project Details Form';
  const ssId = getSsId_();
  const specs = [{ name: DET, startRow: 1 }];
  if (tab !== DET) specs.push({ name: tab, startRow: o.headerRow || 1 });

  const t = readTabs_(ssId, specs, { unformatted: o.unformatted, links: o.links });
  if (!t[DET] || !t[tab]) {
    return { error: "Required sheet ('" + DET + "' or '" + tab + "') not found." };
  }

  const dv = t[DET].values;
  const dh = dv[0] || [];
  const pIdx = dh.indexOf('PROJECT NAME');
  const iIdx = dh.indexOf('ISO No.');
  const projectRow = dv.slice(1).find(r => r[pIdx] === projectName);
  if (!projectRow) return { error: `Project '${projectName}' not found.` };

  const isoNo = o.rawIso ? projectRow[iIdx] : cleanIso(projectRow[iIdx]);
  if (!isoNo && !o.allowNoIso) return { error: "ISO No. not found for this project." };

  const v = t[tab].values;
  return {
    isoNo: isoNo,
    headers: v[0] || [],
    rows: v.slice(1),
    links: t[tab].links ? t[tab].links.slice(1) : null
  };
}

/**
 * Gets specific details for a selected project name.
 */
function getProjectDetails(projectName) {
  const ssId = getSsId_();
  const t = readTabs_(ssId, [{ name: 'Project Details Form' }], { unformatted: true });
  if (!t['Project Details Form']) return null;

  const data = t['Project Details Form'].values;
  const headers = data[0] || [];

  const projectIndex = headers.indexOf('PROJECT NAME');
  const isoIndex = headers.indexOf('ISO No.');
  const addIndex = headers.indexOf('PROJECT ADD');
  const personIndex = headers.indexOf('CONTACT PERSON');
  const itemIndex = headers.indexOf('item');

  const projectRow = data.slice(1).find(row => row[projectIndex] === projectName);

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
  if (!isoNo) return [];
  const ssId = getSsId_();
  const t = readTabs_(ssId, [{ name: 'WBS Sheet' }], { unformatted: true });
  if (!t['WBS Sheet']) return [];

  const wbsData = t['WBS Sheet'].values;
  const wbsHeaders = wbsData[0] || [];

  const wbsIsoIndex = wbsHeaders.indexOf('ISO No.');
  const milestoneIndex = wbsHeaders.indexOf('MILESTONE/BOQ ITEM');
  const plannedStartIndex = wbsHeaders.indexOf('PLANNED START DATE');
  const actualEndIndex = wbsHeaders.indexOf('ACTUAL END DATE');
  if (wbsIsoIndex === -1) return [];

  return wbsData.slice(1)
    .filter(row => row[wbsIsoIndex] === isoNo)
    .map(row => ({
      milestone: row[milestoneIndex],
      planned: serialToDdMmYyyy_(row[plannedStartIndex]),
      actual: serialToDdMmYyyy_(row[actualEndIndex])
    }))
    .filter(d => d.milestone && d.planned);
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
    const d = loadProjectTab_(projectName, 'WBS Sheet');
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'WBS Sheet'." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getWbsDataForProject: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches all PROJECT PLANNING sheet data for a specific project based on its ISO No.
 */
function getProjectPlanningData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING');
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'PROJECT PLANNING' sheet." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getProjectPlanningData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches all INVENTORY sheet data for a specific project (headers on row 2).
 * Replaces the 'PO NO' cell text with its hyperlink when one exists.
 */
function getInventoryData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'INVENTORY', { headerRow: 2, links: true });
    if (d.error) return { error: d.error };

    if (!d.headers.length) return { error: "Not enough data in 'INVENTORY' sheet." };

    const isoIdx = d.headers.indexOf('ISO No.');
    const poIdx = d.headers.indexOf('PO NO');
    if (isoIdx === -1) return { error: "'ISO No.' column not found on row 2 of 'INVENTORY' sheet." };

    const rows = [];
    d.rows.forEach((row, i) => {
      if (row[isoIdx] === d.isoNo) {
        if (poIdx !== -1) {
          const link = d.links[i][poIdx];
          if (link) row[poIdx] = link;
        }
        rows.push(row);
      }
    });
    return { headers: d.headers, rows: rows };
  } catch (e) {
    Logger.log('Error in getInventoryData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches data from the 'Manpower Cost' sheet for a specific project.
 */
function getManpowerCostData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'Manpower Cost');
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'Manpower Cost' sheet." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getManpowerCostData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches data from the 'FINANCE/ACCOUNTS' sheet for a specific project.
 */
function getFinanceData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'FINANCE/ACCOUNTS');
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'FINANCE/ACCOUNTS' sheet." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getFinanceData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Fetches the row of 'Project Details Form' for a given project, replacing
 * hyperlinked cells with their URLs.
 */
function getProjectDetailsFormData(projectName) {
  try {
    const DET = 'Project Details Form';
    const ssId = getSsId_();
    const t = readTabs_(ssId, [{ name: DET }], { links: true });
    if (!t[DET]) return { error: "'Project Details Form' sheet not found." };

    const values = t[DET].values;
    const links = t[DET].links;
    const headers = values[0] || [];

    const projectColIdx = headers.indexOf('PROJECT NAME');
    if (projectColIdx === -1) return { error: "'PROJECT NAME' column not found." };

    const rowIdx = values.slice(1).findIndex(row => row[projectColIdx] === projectName);
    if (rowIdx === -1) return { error: `Project '${projectName}' not found in 'Project Details Form'.` };

    const rowData = values[rowIdx + 1];
    const rowLinks = links[rowIdx + 1];
    const linkHeaders = ['SITE PICS', 'INFO-SHEET', 'PROJECT DRAWING', 'CUSTOMER PO COPY', 'CLIENT SITE VISITS/MOM'];

    headers.forEach((header, index) => {
      if (linkHeaders.includes(header) && rowLinks[index]) rowData[index] = rowLinks[index];
    });

    return { headers: headers, row: rowData };
  } catch (e) {
    Logger.log('Error in getProjectDetailsFormData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/** Shared body for the three audit tabs without links (headers on row 2). */
function getAuditTabData_(projectName, tab) {
  try {
    const d = loadProjectTab_(projectName, tab, { headerRow: 2 });
    if (d.error) return { error: d.error };
    if (!d.headers.length) return { headers: [], rows: [] };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found on row 2 of '" + tab + "'." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getAuditTabData_ (' + tab + '): ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/** Audit data from 'Auditor Sheet Manpower' (headers on row 2). */
function getAuditManpowerData(projectName) {
  return getAuditTabData_(projectName, 'Auditor Sheet Manpower');
}

/** Audit data from 'Auditor Sheet Item' (headers on row 2). */
function getAuditItemData(projectName) {
  return getAuditTabData_(projectName, 'Auditor Sheet Item');
}

/**
 * Audit data from 'Auditor Sheet Drawing' (headers on row 2).
 * Replaces 'Drawing Attachment' / 'Final Drawing' text with hyperlinks.
 */
function getAuditDrawingData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'Auditor Sheet Drawing', { headerRow: 2, links: true });
    if (d.error) return { error: d.error };
    if (!d.headers.length) return { headers: [], rows: [] };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found on row 2 of 'Auditor Sheet Drawing'." };

    const linkIdx = ['Drawing Attachment', 'Final Drawing'].map(h => d.headers.indexOf(h));
    const rows = [];
    d.rows.forEach((row, i) => {
      if (row[isoIdx] === d.isoNo) {
        const newRow = row.slice();
        linkIdx.forEach(ci => {
          if (ci !== -1 && d.links[i][ci]) newRow[ci] = d.links[i][ci];
        });
        rows.push(newRow);
      }
    });
    return { headers: d.headers, rows: rows };
  } catch (e) {
    Logger.log('Error in getAuditDrawingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Summary of worked hours per engineer for a project.
 */
function getEngineerWorkSummary(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING', { unformatted: true, rawIso: true });
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    const doerIdx = d.headers.indexOf('Engineer Name (Doer)');
    const hrsIdx = d.headers.indexOf('Worked (Hrs)');
    if (isoIdx === -1 || doerIdx === -1 || hrsIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    const engineerHours = {};
    d.rows.filter(r => r[isoIdx] === d.isoNo).forEach(row => {
      const name = row[doerIdx];
      const hours = parseFloat(row[hrsIdx]) || 0;
      if (name && hours > 0) engineerHours[name] = (engineerHours[name] || 0) + hours;
    });

    return Object.keys(engineerHours)
      .map(name => ({ name: name, hours: engineerHours[name] }))
      .sort((a, b) => b.hours - a.hours);
  } catch (e) {
    Logger.log('Error in getEngineerWorkSummary: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Fetches data from the 'MOM' sheet for a project ISO No., with hyperlinks.
 * @param {string} isoNo The ISO Number of the project.
 */
function getMomData(isoNo) {
  try {
    if (!isoNo) return { error: "ISO No. was not provided to the server." };

    const ssId = getSsId_();
    const t = readTabs_(ssId, [{ name: 'MOM' }], { links: true });
    if (!t['MOM']) return { error: "Required sheet ('MOM') not found." };

    const values = t['MOM'].values;
    const links = t['MOM'].links;
    const momHeaders = (values[0] || []).map(h => String(h).trim());

    const isoIdx = momHeaders.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'MOM' sheet." };

    const linkColumns = [
      'Upload Image 1', 'Upload Image 2', 'Upload Image 3',
      'Upload Image 4', 'Upload Image 5', 'PDF Link', 'Client Signature (Client Attendee 1)',
      'Mentor Signature (Mentor Attendee 1)'
    ];
    const linkIdx = linkColumns.map(h => momHeaders.indexOf(h));

    const rows = [];
    values.slice(1).forEach((row, i) => {
      if (row[isoIdx] === isoNo) {
        linkIdx.forEach(ci => {
          if (ci !== -1 && links[i + 1][ci]) row[ci] = links[i + 1][ci];
        });
        rows.push(row);
      }
    });
    return { headers: momHeaders, rows: rows };
  } catch (e) {
    Logger.log('Error in getMomData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Detailed task counts per milestone (used by the milestone details modal).
 */
function getMilestoneDetailsData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING', { unformatted: true, rawIso: true });
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    const mileIdx = d.headers.indexOf('Milestone');
    const statusIdx = d.headers.indexOf('Status');
    if (isoIdx === -1 || mileIdx === -1 || statusIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    const stats = {};
    d.rows.filter(r => r[isoIdx] === d.isoNo).forEach(task => {
      const name = task[mileIdx];
      if (!name) return;
      if (!stats[name]) stats[name] = { done: 0, inProgress: 0, hold: 0 };
      const status = (task[statusIdx] || '').toString().toUpperCase().trim();
      if (status === 'DONE') stats[name].done++;
      else if (status === 'IN-PROGRESS') stats[name].inProgress++;
      else if (status === 'HOLD') stats[name].hold++;
    });

    return Object.keys(stats).map(name => {
      const s = stats[name];
      return {
        milestone: name,
        total: s.done + s.inProgress + s.hold,
        done: s.done,
        pending: s.inProgress + s.hold
      };
    });
  } catch (e) {
    Logger.log('Error in getMilestoneDetailsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Product consumption rows from 'INVENTORY' (headers on row 2) where CONSU.2 > 0.
 */
function getProductUsedData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'INVENTORY', { headerRow: 2, rawIso: true });
    if (d.error) return { error: d.error };
    if (!d.headers.length) return { headers: [], rows: [] };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found on row 2 of 'INVENTORY' sheet." };

    const consu2Idx = d.headers.indexOf('CONSU.2');
    const rows = d.rows.filter(row =>
      row[isoIdx] === d.isoNo && consu2Idx !== -1 && parseFloat(row[consu2Idx]) > 0
    );
    return { headers: d.headers, rows: rows };
  } catch (e) {
    Logger.log('Error in getProductUsedData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * All rows of 'Expenses' for a project.
 */
function getExpenseDetailsData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'Expenses', { rawIso: true });
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    if (isoIdx === -1) return { error: "'ISO No.' column not found in 'Expenses' sheet." };

    return { headers: d.headers, rows: d.rows.filter(row => row[isoIdx] === d.isoNo) };
  } catch (e) {
    Logger.log('Error in getExpenseDetailsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Tasks from 'PROJECT PLANNING' filtered by status ('DONE', 'IN-PROGRESS', or 'HOLD').
 */
function getFilteredTasks(projectName, filterType) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING');
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    const statusIdx = d.headers.indexOf('Status');
    if (isoIdx === -1 || statusIdx === -1) {
      return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };
    }

    let rows = [];
    if (['DONE', 'IN-PROGRESS', 'HOLD'].includes(filterType)) {
      rows = d.rows.filter(r => r[isoIdx] === d.isoNo).filter(task => {
        const s = (task[statusIdx] || '').toString().toUpperCase().trim();
        if (filterType === 'HOLD') {
          // Matches KPI logic: anything not Done / In-Progress (incl. blank) is HOLD
          return s !== 'DONE' && s !== 'IN-PROGRESS' && s !== 'IN PROGRESS';
        } else if (filterType === 'IN-PROGRESS') {
          return s === 'IN-PROGRESS' || s === 'IN PROGRESS';
        }
        return s === filterType;
      });
    }
    return { headers: d.headers, rows: rows };
  } catch (e) {
    Logger.log('Error in getFilteredTasks: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Tasks from 'PROJECT PLANNING' that have planned hours assigned.
 */
function getPlannedHrsData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING', { rawIso: true });
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    const hrsIdx = d.headers.indexOf('Required Hours');
    if (isoIdx === -1 || hrsIdx === -1) return { error: "Required columns not found." };

    return {
      headers: d.headers,
      rows: d.rows.filter(row => row[isoIdx] === d.isoNo && parseFloat(row[hrsIdx]) > 0)
    };
  } catch (e) {
    Logger.log('Error in getPlannedHrsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Tasks from 'PROJECT PLANNING' that have actual hours logged.
 */
function getActualHrsData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'PROJECT PLANNING', { rawIso: true });
    if (d.error) return { error: d.error };

    const isoIdx = d.headers.indexOf('ISO No.');
    const hrsIdx = d.headers.indexOf('Worked (Hrs)');
    if (isoIdx === -1 || hrsIdx === -1) return { error: "Required columns not found." };

    return {
      headers: d.headers,
      rows: d.rows.filter(row => row[isoIdx] === d.isoNo && parseFloat(row[hrsIdx]) > 0)
    };
  } catch (e) {
    Logger.log('Error in getActualHrsData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * Milestone names from the WBS Sheet filtered by ACTUAL END DATE ('COMPLETED' or 'PENDING').
 */
function getFilteredMilestones(projectName, filterType) {
  try {
    const d = loadProjectTab_(projectName, 'WBS Sheet', { unformatted: true, allowNoIso: true });
    if (d.error) return { error: d.error };

    const wHeaders = d.headers.map(h => String(h).trim());
    const isoIdx = wHeaders.indexOf('ISO No.');
    const mileIdx = wHeaders.indexOf('MILESTONE/BOQ ITEM');
    const endIdx = wHeaders.indexOf('ACTUAL END DATE');
    if (isoIdx === -1 || mileIdx === -1 || endIdx === -1) {
      return { error: "Required columns not found in 'WBS Sheet'." };
    }

    const groups = {};
    d.rows.filter(r => String(r[isoIdx]).trim() === d.isoNo).forEach(row => {
      const name = String(row[mileIdx]).trim();
      if (!name || name === 'N/A') return;
      (groups[name] = groups[name] || []).push(row);
    });

    const result = [];
    for (const name in groups) {
      // COMPLETED only if EVERY row has an Actual End Date
      const isComplete = groups[name].every(t => {
        const v = t[endIdx];
        return v && String(v).trim() !== '' && String(v).toUpperCase() !== 'N/A';
      });
      if ((filterType === 'COMPLETED' && isComplete) || (filterType === 'PENDING' && !isComplete)) {
        result.push(name);
      }
    }
    return result.sort();
  } catch (e) {
    Logger.log('Error in getFilteredMilestones: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * 'INVENTORY' rows (headers on row 2) for a project where `filterColumn` > 0.
 */
function getFilteredInventoryData(projectName, filterColumn) {
  try {
    const d = loadProjectTab_(projectName, 'INVENTORY', { headerRow: 2 });
    if (d.error) return { error: d.error };
    if (!d.headers.length) return { headers: [], rows: [] };

    const isoIdx = d.headers.indexOf('ISO No.');
    const colIdx = d.headers.indexOf(filterColumn);
    if (isoIdx === -1 || colIdx === -1) {
      return { error: `Required column ('ISO No.' or '${filterColumn}') not found in 'INVENTORY' sheet.` };
    }

    const rows = d.rows.filter(row =>
      cleanIso(row[isoIdx]) === d.isoNo && parseFloat(String(row[colIdx]).replace(/,/g, '')) > 0
    );
    return { headers: d.headers, rows: rows };
  } catch (e) {
    Logger.log('Error in getFilteredInventoryData: ' + e.toString());
    return { error: 'An error occurred on the server: ' + e.message };
  }
}

/**
 * 'Billing' rows for a project; 'Supp Doc' text is replaced with its hyperlink.
 */
function getBillingData(projectName) {
  try {
    const d = loadProjectTab_(projectName, 'Billing', { links: true, allowNoIso: true });
    if (d.error) return { error: d.error };

    const bHeaders = d.headers.map(h => String(h).trim());
    let isoIdx = bHeaders.indexOf('ISO NO');
    if (isoIdx === -1) isoIdx = bHeaders.indexOf('ISO No.');
    const suppIdx = bHeaders.indexOf('Supp Doc');
    if (isoIdx === -1) return { error: "'ISO NO' column not found in 'Billing' sheet." };

    const rows = [];
    d.rows.forEach((row, i) => {
      if (cleanIso(row[isoIdx]) === d.isoNo) {
        if (suppIdx !== -1 && d.links[i][suppIdx]) row[suppIdx] = d.links[i][suppIdx];
        rows.push(row);
      }
    });
    return { headers: bHeaders, rows: rows };
  } catch (e) {
    Logger.log('Error in getBillingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Data for the task progress chart, using ISO No. (dates arrive as serial numbers).
 */
function getTaskStatusChartData(isoNo = 'MWEPL/2024-25/01250') {
  try {
    if (!isoNo) return [];
    const ssId = getSsId_();
    const t = readTabs_(ssId, [{ name: 'PROJECT PLANNING' }], { unformatted: true });
    if (!t['PROJECT PLANNING']) {
      Logger.log('getTaskStatusChartData: Sheet not found.');
      return [];
    }

    const pData = t['PROJECT PLANNING'].values;
    const pHeaders = (pData[0] || []).map(h => String(h).trim());
    const isoIdx = pHeaders.indexOf('ISO No.');
    const taskIdx = pHeaders.indexOf('Task');
    const pctIdx = pHeaders.indexOf('Status %');
    const dateIdx = pHeaders.indexOf('Planned Date');

    if (isoIdx === -1 || taskIdx === -1 || pctIdx === -1) {
      Logger.log('getTaskStatusChartData: A required column was not found.');
      return [];
    }

    const chartData = pData.slice(1)
      .filter(row => row[isoIdx] === isoNo && row[taskIdx])
      .map(row => {
        const raw = dateIdx > -1 ? row[dateIdx] : null;
        let ms = null;
        if (typeof raw === 'number' && raw > 0) {
          ms = Math.round((raw - 25569) * 86400 * 1000); // serial -> epoch ms (UTC)
        } else if (typeof raw === 'string') {
          const m = raw.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
          if (m) ms = Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        }
        return {
          milestone: row[taskIdx],
          planned: parseFloat(row[pctIdx]) || 0,
          actual: ms !== null ? Utilities.formatDate(new Date(ms), 'UTC', 'dd-MM-yyyy') : 'N/A',
          rawDate: ms // epoch ms or null (Date objects can't be returned to the client)
        };
      });

    chartData.sort((a, b) => {
      if (a.rawDate === null) return 1;
      if (b.rawDate === null) return -1;
      return a.rawDate - b.rawDate;
    });
    return chartData;
  } catch (e) {
    Logger.log('FATAL Error in getTaskStatusChartData: ' + e.toString());
    return [];
  }
}

/**
 * 'Billing' rows filtered by status ('DONE' or 'PENDING').
 */
function getFilteredBillingData(projectName, filterType) {
  try {
    const d = loadProjectTab_(projectName, 'Billing');
    if (d.error) return { error: d.error };

    const bHeaders = d.headers.map(h => String(h).trim());
    const isoIdx = bHeaders.indexOf('ISO NO');
    const statusIdx = bHeaders.indexOf('Status5');
    if (isoIdx === -1 || statusIdx === -1) return { error: "Required columns not found in 'Billing' sheet." };

    const isDone = r => (r[statusIdx] || '').toString().trim().toUpperCase() === 'DONE';
    const projectBillings = d.rows.filter(r => r[isoIdx] === d.isoNo);
    let rows = [];
    if (filterType === 'DONE') rows = projectBillings.filter(isDone);
    else if (filterType === 'PENDING') rows = projectBillings.filter(r => !isDone(r));

    return { headers: bHeaders, rows: rows };
  } catch (e) {
    Logger.log('Error in getFilteredBillingData: ' + e.toString());
    return { error: 'An error occurred: ' + e.message };
  }
}

/**
 * Data for the Milestone Progress combo chart (planning + billing in one API call).
 * @param {string} isoNo The ISO Number of the project.
 */
function getMilestoneComboChartData(isoNo) {
  try {
    if (!isoNo) return [];
    const ssId = getSsId_();
    const t = readTabs_(ssId, [{ name: 'PROJECT PLANNING' }, { name: 'Billing' }], { unformatted: true });
    if (!t['PROJECT PLANNING'] || !t['Billing']) return [];

    const pData = t['PROJECT PLANNING'].values;
    const pHeaders = (pData[0] || []).map(h => String(h).trim());
    const pIsoIdx = pHeaders.indexOf('ISO No.');
    const pMileIdx = pHeaders.indexOf('Milestone');
    const pPctIdx = pHeaders.indexOf('Status %');
    const projectTasks = pData.slice(1).filter(r => r[pIsoIdx] === isoNo);

    const bData = t['Billing'].values;
    const bHeaders = (bData[0] || []).map(h => String(h).trim());
    const bIsoIdx = bHeaders.indexOf('ISO NO');
    const bMileIdx = bHeaders.indexOf('Milestones');
    const bStatusIdx = bHeaders.indexOf('Status5');
    const projectBillings = bData.slice(1).filter(r => r[bIsoIdx] === isoNo);

    const allMilestones = [...new Set([
      ...projectTasks.map(r => r[pMileIdx]),
      ...projectBillings.map(r => r[bMileIdx])
    ])].filter(Boolean);

    return allMilestones.map(name => {
      const tasks = projectTasks.filter(r => r[pMileIdx] === name);
      let taskProgress = 0;
      if (tasks.length > 0) {
        taskProgress = tasks.reduce((s, r) => s + (parseFloat(r[pPctIdx]) || 0), 0) / tasks.length;
      }

      const bills = projectBillings.filter(r => r[bMileIdx] === name);
      let billingStatus = 0;
      if (bills.length > 0) {
        const done = bills.filter(b => (b[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE').length;
        billingStatus = (done / bills.length) * 100;
      }
      return { milestone: name, taskProgress: taskProgress, billingStatus: billingStatus };
    });
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