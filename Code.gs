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

function updateCategoryWorkspace(payload) {
  return DatabaseController.updateCategoryWorkspace(payload);
}

function deleteCategoryWorkspace(catId) {
  return DatabaseController.deleteCategoryWorkspace(catId);
}

function updateExpenseRow(payload) {
  return DatabaseController.updateExpenseRow(payload);
}

function deleteExpenseRow(expenseId, categoryName) {
  return DatabaseController.deleteExpenseRow(expenseId, categoryName);
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
        try {
          rightsRecord.categories = accessData[j][1] ? JSON.parse(accessData[j][1]) : [];
          rightsRecord.months = accessData[j][2] ? JSON.parse(accessData[j][2]) : [];
        } catch(e) {
          rightsRecord.categories = [];
          rightsRecord.months = [];
        }
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
          let members = [];
          try {
            members = data[i][4] ? JSON.parse(data[i][4]) : [];
          } catch(e) {
            members = [];
          }
          output.push({
            id: catId,
            name: data[i][1],
            description: data[i][2],
            createdBy: data[i][3],
            members: members,
            type: data[i][5]
          });
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
          breakdown: cleanBreakdown, receipt: data[i][9] || "",
          recordedBy: data[i][10] ? data[i][10].toLowerCase() : ""
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
      
      this.syncCategoryAccess(newCatId, payload.members);
      
      return { success: true, catId: newCatId };
    } finally {
      lock.releaseLock();
    }
  },

  syncCategoryAccess: function(catId, members) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const accessSheet = ss.getSheetByName("ACCESS_RIGHTS");
    const accessData = accessSheet.getDataRange().getValues();
    const membersLower = members.map(m => m.toLowerCase().trim());
    const processedEmails = [];

    for (let r = 1; r < accessData.length; r++) {
      const email = accessData[r][0].toLowerCase().trim();
      processedEmails.push(email);
      let existingCats = [];
      try {
        existingCats = accessData[r][1] ? JSON.parse(accessData[r][1]) : [];
      } catch (e) {
        existingCats = [];
      }

      const hasId = existingCats.includes(catId);
      const shouldHaveId = membersLower.includes(email);
      const isSuperAdmin = existingCats.includes("*");

      if (shouldHaveId && !hasId && !isSuperAdmin) {
        existingCats.push(catId);
        accessSheet.getRange(r + 1, 2).setValue(JSON.stringify(existingCats));
      } else if (!shouldHaveId && hasId) {
        const index = existingCats.indexOf(catId);
        if (index > -1) {
          existingCats.splice(index, 1);
          accessSheet.getRange(r + 1, 2).setValue(JSON.stringify(existingCats));
        }
      }
    }

    // Add new members who don't have an entry yet
    membersLower.forEach(email => {
      if (!processedEmails.includes(email)) {
        accessSheet.appendRow([email, JSON.stringify([catId]), JSON.stringify(["*"])]);
      }
    });
  },

  updateCategoryWorkspace: function(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
      const auth = SecurityEngine.validateUserAccess(activeEmail);
      if (auth.profile.role !== 'Admin') throw new Error("Unauthorized: Only Admins can edit groups.");

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("CATEGORY_MASTER");
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          const oldName = data[i][1];
          sheet.getRange(i + 1, 2, 1, 2).setValues([[payload.name, payload.description]]);
          sheet.getRange(i + 1, 5, 1, 2).setValues([[JSON.stringify(payload.members), payload.type]]);

          if (oldName !== payload.name) {
            const oldSheet = ss.getSheetByName(oldName);
            if (oldSheet) oldSheet.setName(payload.name);
          }

          this.syncCategoryAccess(payload.id, payload.members);

          return { success: true };
        }
      }
      throw new Error("Category not found.");
    } finally {
      lock.releaseLock();
    }
  },

  deleteCategoryWorkspace: function(catId) {
     const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
      const auth = SecurityEngine.validateUserAccess(activeEmail);
      if (auth.profile.role !== 'Admin') throw new Error("Unauthorized: Only Admins can delete groups.");

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("CATEGORY_MASTER");
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === catId) {
          sheet.getRange(i + 1, 8).setValue("Inactive");
          return { success: true };
        }
      }
      throw new Error("Category not found.");
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
      const activeUser = Session.getActiveUser().getEmail().toLowerCase();
      
      const newId = "EXP_" + new Date().getFullYear() + "_" + Math.floor(10000 + Math.random() * 90000);
      const rowData = [
        newId, expenseObj.catId, expenseObj.title, expenseObj.amount, expenseObj.type,
        expenseObj.date, expenseObj.date.substring(0,7), expenseObj.paidBy.toLowerCase(),
        JSON.stringify(expenseObj.breakdown), expenseObj.receipt || "",
        activeUser
      ];
      
      masterSheet.appendRow(rowData);
      if(categorySheet) categorySheet.appendRow(rowData);
      
      return { success: true, id: newId };
    } finally {
      lock.releaseLock();
    }
  },

  updateExpenseRow: function(expenseObj) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
      const auth = SecurityEngine.validateUserAccess(activeEmail);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const masterSheet = ss.getSheetByName("EXPENSES_MASTER");
      const data = masterSheet.getDataRange().getValues();

      let rowIndex = -1;
      let recordedBy = "";
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === expenseObj.id) {
          rowIndex = i + 1;
          recordedBy = data[i][10] ? data[i][10].toLowerCase() : "";
          break;
        }
      }

      if (rowIndex === -1) throw new Error("Expense not found.");
      if (auth.profile.role !== 'Admin' && recordedBy !== activeEmail) {
        throw new Error("Unauthorized: You can only edit expenses you recorded.");
      }

      const rowData = [
        expenseObj.id, expenseObj.catId, expenseObj.title, expenseObj.amount, expenseObj.type,
        expenseObj.date, expenseObj.date.substring(0,7), expenseObj.paidBy.toLowerCase(),
        JSON.stringify(expenseObj.breakdown), expenseObj.receipt || ""
      ];

      masterSheet.getRange(rowIndex, 1, 1, 10).setValues([rowData]);

      const categorySheet = ss.getSheetByName(expenseObj.categoryName);
      if (categorySheet) {
        const catData = categorySheet.getDataRange().getValues();
        for (let j = 1; j < catData.length; j++) {
          if (catData[j][0] === expenseObj.id) {
            categorySheet.getRange(j + 1, 1, 1, 10).setValues([rowData]);
            break;
          }
        }
      }

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  },

  deleteExpenseRow: function(expenseId, categoryName) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
      const auth = SecurityEngine.validateUserAccess(activeEmail);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const masterSheet = ss.getSheetByName("EXPENSES_MASTER");
      const data = masterSheet.getDataRange().getValues();

      let rowIndex = -1;
      let recordedBy = "";
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === expenseId) {
          rowIndex = i + 1;
          recordedBy = data[i][10] ? data[i][10].toLowerCase() : "";
          break;
        }
      }

      if (rowIndex === -1) throw new Error("Expense not found.");
      if (auth.profile.role !== 'Admin' && recordedBy !== activeEmail) {
        throw new Error("Unauthorized: You can only delete expenses you recorded.");
      }

      masterSheet.deleteRow(rowIndex);

      const categorySheet = ss.getSheetByName(categoryName);
      if (categorySheet) {
        const catData = categorySheet.getDataRange().getValues();
        for (let j = 1; j < catData.length; j++) {
          if (catData[j][0] === expenseId) {
            categorySheet.deleteRow(j + 1);
            break;
          }
        }
      }

      return { success: true };
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
    list.push({
      id: data[i][0],
      catId: data[i][1],
      title: data[i][2],
      amount: parseFloat(data[i][3]) || 0,
      createdBy: data[i][4] ? data[i][4].toLowerCase() : ""
    });
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

function updateEstimatorLine(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
    const auth = SecurityEngine.validateUserAccess(activeEmail);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BUDGET_ESTIMATOR");
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let createdBy = "";
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        rowIndex = i + 1;
        createdBy = data[i][4] ? data[i][4].toLowerCase() : "";
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Line item not found.");
    if (auth.profile.role !== 'Admin' && createdBy !== activeEmail) {
      throw new Error("Unauthorized: You can only edit your own proposed goals.");
    }

    sheet.getRange(rowIndex, 2, 1, 2).setValues([[payload.catId, payload.title]]);
    sheet.getRange(rowIndex, 4).setValue(payload.amount);

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function deleteEstimatorLine(estId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
    const auth = SecurityEngine.validateUserAccess(activeEmail);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BUDGET_ESTIMATOR");
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let createdBy = "";
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === estId) {
        rowIndex = i + 1;
        createdBy = data[i][4] ? data[i][4].toLowerCase() : "";
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Line item not found.");
    if (auth.profile.role !== 'Admin' && createdBy !== activeEmail) {
      throw new Error("Unauthorized: You can only delete your own proposed goals.");
    }

    sheet.deleteRow(rowIndex);
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
    const activeUser = Session.getActiveUser().getEmail().toLowerCase();
    
    const newId = "EST_" + Math.floor(10000 + Math.random() * 90000);
    sheet.appendRow([newId, payload.catId, payload.title, payload.amount, activeUser]);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}
