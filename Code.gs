function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FinTrack Core // Elite Expense Matrix')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Hydrates core initialization state based on visitor verification profiling.
 */
function getApplicationInitialState() {
  const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
  const authProfile = SecurityEngine.validateUserAccess(activeEmail);
  
  if (!authProfile.authorized) {
    return { authorized: false, reason: "Unauthorized email signature access profile." };
  }
  
  const accessibleCategories = DatabaseController.getCategoriesForUser(activeEmail, authProfile.rights);
  const usersList = DatabaseController.getAllUsersProfileList();
  
  return {
    authorized: true,
    userProfile: authProfile.profile,
    rights: authProfile.rights,
    categories: accessibleCategories,
    systemUsers: usersList
  };
}

// =========================================================================
// SYSTEM CONTROLLER INTERACTION BRIDGE
// =========================================================================

function getExpensesData(allowedCategories) {
  return DatabaseController.getExpensesData(allowedCategories);
}

function getSettlementsData(allowedCategories) {
  return DatabaseController.getSettlementsData(allowedCategories);
}

function createNewCategoryWorkspace(categoryPayload) {
  return DatabaseController.createNewCategoryWorkspace(categoryPayload);
}

function commitExpenseRow(payload) {
  return DatabaseController.commitExpenseRow(payload);
}

function recordDirectSettlement(settlementPayload) {
  return DatabaseController.recordDirectSettlement(settlementPayload);
}

// =========================================================================
// SECURITY LAYER ENGINE MODULE
// =========================================================================
const SecurityEngine = {
  validateUserAccess: function(email) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName("USERS");
    const accessSheet = ss.getSheetByName("ACCESS_RIGHTS");
    
    if(!usersSheet || !accessSheet) return { authorized: false };
    
    const usersData = usersSheet.getDataRange().getValues();
    let userRecord = null;
    
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][2].toLowerCase() === email) {
        userRecord = { id: usersData[i][0], name: usersData[i][1], email: usersData[i][2], photo: usersData[i][3], role: usersData[i][4], status: usersData[i][5] };
        break;
      }
    }
    
    if (!userRecord || userRecord.status !== "Active") return { authorized: false };
    
    const accessData = accessSheet.getDataRange().getValues();
    let rightsRecord = { categories: [], months: [] };
    
    for (let j = 1; j < accessData.length; j++) {
      if (accessData[j][0].toLowerCase() === email) {
        rightsRecord.categories = JSON.parse(accessData[j][1]);
        rightsRecord.months = JSON.parse(accessData[j][2]);
        break;
      }
    }
    
    return { authorized: true, profile: userRecord, rights: rightsRecord };
  }
};

// =========================================================================
// MASTER PERSISTENCE CONTROLLER
// =========================================================================
const DatabaseController = {
  getAllUsersProfileList: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("USERS");
    const data = sheet.getDataRange().getValues();
    const list = [];
    for(let i = 1; i < data.length; i++) {
      if(data[i][5] === "Active") {
        list.push({ name: data[i][1], email: data[i][2].toLowerCase() });
      }
    }
    return list;
  },

  getCategoriesForUser: function(email, rights) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CATEGORY_MASTER");
    const data = sheet.getDataRange().getValues();
    const output = [];
    
    for(let i = 1; i < data.length; i++) {
      const catId = data[i][0];
      if (rights.categories.includes(catId) || rights.categories.includes("*")) {
        if(data[i][7] === "Active") {
          output.push({ id: catId, name: data[i][1], description: data[i][2], members: JSON.parse(data[i][4]) });
        }
      }
    }
    return output;
  },

  getExpensesData: function(allowedCategories) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("EXPENSES_MASTER");
    if(!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const records = [];
    
    for(let i = 1; i < data.length; i++) {
      const catId = data[i][1];
      if(allowedCategories.includes(catId) || allowedCategories.includes("*")) {
        let cleanBreakdown = [];
        try {
          cleanBreakdown = typeof data[i][8] === 'string' ? JSON.parse(data[i][8]) : data[i][8];
        } catch(e) {
          cleanBreakdown = [];
        }

        records.push({
          id: data[i][0], catId: catId, title: data[i][2], amount: parseFloat(data[i][3]) || 0,
          type: data[i][4], date: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          yearMonth: data[i][6] ? String(data[i][6]).trim() : "", paidBy: data[i][7] ? String(data[i][7]).toLowerCase().trim() : "", 
          breakdown: cleanBreakdown, receipt: data[i][9] || ""
        });
      }
    }
    return records;
  },

  getSettlementsData: function(allowedCategories) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SETTLEMENTS_MASTER");
    if(!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const records = [];
    
    for(let i = 1; i < data.length; i++) {
      const catId = data[i][1];
      if(allowedCategories.includes(catId) || allowedCategories.includes("*")) {
        records.push({
          id: data[i][0], catId: catId, fromUser: data[i][2].toLowerCase(), toUser: data[i][3].toLowerCase(),
          amount: parseFloat(data[i][4]), date: Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd"), proof: data[i][6]
        });
      }
    }
    return records;
  },

  createNewCategoryWorkspace: function(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const catMaster = ss.getSheetByName("CATEGORY_MASTER");
      const data = catMaster.getDataRange().getValues();
      
      const newCatIndex = data.length;
      const newCatId = "CAT_" + String(newCatIndex).padStart(3, '0');
      
      const rowData = [
        newCatId, payload.name, payload.description, payload.createdBy,
        JSON.stringify(payload.members), payload.type, new Date(), "Active"
      ];
      catMaster.appendRow(rowData);
      
      const templateTab = ss.insertSheet(payload.name);
      const headers = [["Expense ID", "Category ID", "Title", "Amount", "Type", "Expense Date", "Year-Month", "Paid By (Abono)", "Participant Breakdown JSON", "Receipt URL"]];
      templateTab.getRange(1, 1, 1, 10).setValues(headers).setFontWeight("bold").setBackground("#0f172a").setFontColor("#f8fafc");
      
      const accessSheet = ss.getSheetByName("ACCESS_RIGHTS");
      const accessData = accessSheet.getDataRange().getValues();
      
      payload.members.forEach(memberEmail => {
        let matched = false;
        for(let r = 1; r < accessData.length; r++) {
          if(accessData[r][0].toLowerCase() === memberEmail.toLowerCase()) {
            let existingCats = JSON.parse(accessData[r][1]);
            if(!existingCats.includes(newCatId) && !existingCats.includes("*")) {
              existingCats.push(newCatId);
              accessSheet.getRange(r + 1, 2).setValue(JSON.stringify(existingCats));
            }
            matched = true; break;
          }
        }
        if(!matched) {
          accessSheet.appendRow([memberEmail.toLowerCase(), JSON.stringify([newCatId]), JSON.stringify(["*"])]);
        }
      });
      
      return { success: true, catId: newCatId };
    } finally {
      lock.releaseLock();
    }
  },

  commitExpenseRow: function(expenseObj) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const masterSheet = ss.getSheetByName("EXPENSES_MASTER");
      const categorySheet = ss.getSheetByName(expenseObj.categoryName);
      
      const newId = "EXP_" + new Date().getFullYear() + "_" + Math.floor(10000 + Math.random() * 90000);
      const rowData = [
        newId, expenseObj.catId, expenseObj.title, expenseObj.amount, expenseObj.type,
        expenseObj.date, expenseObj.date.substring(0,7), expenseObj.paidBy.toLowerCase(),
        JSON.stringify(expenseObj.breakdown), expenseObj.receipt || ""
      ];
      
      masterSheet.appendRow(rowData);
      if(categorySheet) categorySheet.appendRow(rowData);
      
      return { success: true, id: newId };
    } finally {
      lock.releaseLock();
    }
  },

  recordDirectSettlement: function(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const settlementSheet = ss.getSheetByName("SETTLEMENTS_MASTER");
      if (!settlementSheet) throw new Error("Missing sheet: SETTLEMENTS_MASTER");
      
      const newId = "SET_" + new Date().getFullYear() + "_" + Math.floor(10000 + Math.random() * 90000);
      
      settlementSheet.appendRow([
        newId, payload.catId, payload.fromUser.toLowerCase().trim(), payload.toUser.toLowerCase().trim(),
        payload.amount, payload.date, payload.proof || ""
      ]);
      
      const catMaster = ss.getSheetByName("CATEGORY_MASTER");
      let clearTabName = null;
      if (catMaster) {
        const catData = catMaster.getDataRange().getValues();
        for (let c = 1; c < catData.length; c++) {
          if (catData[c][0] === payload.catId) {
            clearTabName = catData[c][1];
            break;
          }
        }
      }
      
      const masterSheet = ss.getSheetByName("EXPENSES_MASTER");
      if (!masterSheet) return { success: true };
      
      const data = masterSheet.getDataRange().getValues();
      
      for(let i = 1; i < data.length; i++) {
        if(data[i][1] === payload.catId && data[i][7].toLowerCase().trim() === payload.toUser.toLowerCase().trim()) {
          let breakdown = [];
          try {
            breakdown = typeof data[i][8] === 'string' ? JSON.parse(data[i][8]) : data[i][8];
          } catch(e) {
            continue;
          }
          
          let matrixUpdated = false;
          breakdown = breakdown.map(p => {
            if(p.email.toLowerCase().trim() === payload.fromUser.toLowerCase().trim() && p.status === "Unpaid") {
              p.status = "Paid";
              matrixUpdated = true;
            }
            return p;
          });
          
          if(matrixUpdated) {
            masterSheet.getRange(i + 1, 9).setValue(JSON.stringify(breakdown));
            if (clearTabName) {
              const categorySheet = ss.getSheetByName(clearTabName); 
              if(categorySheet) {
                const catData = categorySheet.getDataRange().getValues();
                for(let k = 1; k < catData.length; k++) {
                  if(catData[k][0] === data[i][0]) {
                    categorySheet.getRange(k + 1, 9).setValue(JSON.stringify(breakdown));
                    break;
                  }
                }
              }
            }
          }
        }
      }
      return { success: true };
    } catch(err) {
      throw new Error("Backend Matrix Failure: " + err.message);
    } finally {
      lock.releaseLock();
    }
  }
};

// =========================================================================
// PIPELINE DATA-WRITING EXTRA BRIDGES FOR NEW HUB UPGRADES
// =========================================================================

function getQRDirectoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MEMBERS_PAYMENT_HUB");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const list = [];
  for(let i = 1; i < data.length; i++) {
    list.push({ email: data[i][0].toLowerCase().trim(), bank: data[i][1], number: data[i][2], url: data[i][3] });
  }
  return list;
}

function getBudgetEstimatesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BUDGET_ESTIMATOR");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const list = [];
  for(let i = 1; i < data.length; i++) {
    list.push({ id: data[i][0], catId: data[i][1], title: data[i][2], amount: parseFloat(data[i][3]) || 0 });
  }
  return list;
}

function getCombinedDashboardData(allowedCategories) {
  return {
    expenses: getExpensesData(allowedCategories),
    settlements: getSettlementsData(allowedCategories),
    qrDirectory: getQRDirectoryData(),
    estimates: getBudgetEstimatesData()
  };
}

function saveQRHubRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MEMBERS_PAYMENT_HUB");
    if(!sheet) throw new Error("Missing tab: MEMBERS_PAYMENT_HUB");
    
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0].toLowerCase().trim() === payload.email.toLowerCase().trim()) {
        rowIdx = i + 1; break;
      }
    }
    
    if(rowIdx !== -1) {
      sheet.getRange(rowIdx, 2).setValue(payload.bank);
      sheet.getRange(rowIdx, 3).setValue(payload.number);
      sheet.getRange(rowIdx, 4).setValue(payload.url);
    } else {
      sheet.appendRow([payload.email.toLowerCase().trim(), payload.bank, payload.number, payload.url]);
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function commitEstimatorProposedLine(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BUDGET_ESTIMATOR");
    if(!sheet) throw new Error("Missing tab: BUDGET_ESTIMATOR");
    
    const newId = "EST_" + Math.floor(10000 + Math.random() * 90000);
    sheet.appendRow([newId, payload.catId, payload.title, payload.amount]);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}
