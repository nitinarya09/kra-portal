/**
 * Authentication and User Management Service
 */

/**
 * Authenticates a user by username and password hash.
 * @param {string} username 
 * @param {string} passwordHash (SHA-256 hex string)
 * @returns {object} {success: true, token, user} or throws error
 */
function authenticateUser(username, passwordHash) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    initializeDatabase();
    usersSheet = ss.getSheetByName('Users');
  }
  var data = usersSheet.getDataRange().getValues();
  
  var user = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][6] === true) { // Check username and isActive
      if (data[i][1] === passwordHash) {
        user = {
          username: data[i][0],
          displayName: data[i][2],
          role: data[i][3],
          sections: data[i][4],
          designation: data[i][5]
        };
      }
      break;
    }
  }
  
  // Log attempt
  var attemptsSheet = ss.getSheetByName('Login_Attempts');
  attemptsSheet.appendRow([new Date(), username, "0.0.0.0", user !== null]); // Dummy IP as AS doesn't provide client IP directly
  
  if (!user) {
    throw new Error("Invalid username or password, or account is inactive.");
  }
  
  // Create session
  var token = Utilities.getUuid();
  var sessionsSheet = ss.getSheetByName('Sessions');
  var expiry = new Date();
  expiry.setHours(expiry.getHours() + 24); // 24 hours expiry
  
  sessionsSheet.appendRow([token, user.username, user.role, user.sections, new Date(), expiry]);
  
  return { success: true, token: token, user: user };
}

/**
 * Validates a session token.
 * @param {string} token 
 * @returns {object|null} user object or null
 */
function validateSession(token) {
  if (!token) return null;
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sessionsSheet = ss.getSheetByName('Sessions');
  var data = sessionsSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var expiryTime = data[i][5];
      if (expiryTime && expiryTime.getTime() > new Date().getTime()) {
        // Find user to get latest details
        var usersSheet = ss.getSheetByName('Users');
        var userData = usersSheet.getDataRange().getValues();
        for (var j = 1; j < userData.length; j++) {
          if (userData[j][0] === data[i][1] && userData[j][6] === true) {
            return {
              username: userData[j][0],
              displayName: userData[j][2],
              role: userData[j][3],
              sections: userData[j][4],
              designation: userData[j][5]
            };
          }
        }
      } else {
        // Token expired
        return null;
      }
    }
  }
  return null;
}

/**
 * Logs out a user by deleting their session.
 * @param {string} token 
 * @returns {object} {success: true}
 */
function logoutUser(token) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sessionsSheet = ss.getSheetByName('Sessions');
  var data = sessionsSheet.getDataRange().getValues();
  
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === token) {
      sessionsSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: true }; // Token not found, consider already logged out
}

/**
 * Changes a user's password.
 * @param {string} token 
 * @param {string} oldHash 
 * @param {string} newHash 
 * @returns {object} {success: true}
 */
function changePassword(token, oldHash, newHash) {
  var user = validateSession(token);
  if (!user) throw new Error("Invalid or expired session.");
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === user.username) {
      if (data[i][1] !== oldHash) {
        throw new Error("Old password incorrect.");
      }
      // Update password hash
      usersSheet.getRange(i + 1, 2).setValue(newHash);
      return { success: true };
    }
  }
  throw new Error("User not found.");
}

/**
 * Creates a new user (Admin only).
 * @param {string} token 
 * @param {object} userData {username, passwordHash, displayName, role, sections, designation, isActive}
 * @returns {object} {success: true}
 */
function createUser(token, userData) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  
  // Check if user already exists
  var data = usersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userData.username) {
      throw new Error("User already exists.");
    }
  }
  
  usersSheet.appendRow([
    userData.username,
    userData.passwordHash,
    userData.displayName,
    userData.role,
    userData.sections,
    userData.designation,
    userData.isActive !== undefined ? userData.isActive : true,
    new Date(),
    sessionUser.username
  ]);
  
  return { success: true };
}

/**
 * Updates a user's details (Admin only).
 * @param {string} token 
 * @param {string} username 
 * @param {object} updates 
 * @returns {object} {success: true}
 */
function updateUser(token, username, updates) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (updates.displayName !== undefined) usersSheet.getRange(i + 1, 3).setValue(updates.displayName);
      if (updates.role !== undefined) usersSheet.getRange(i + 1, 4).setValue(updates.role);
      if (updates.sections !== undefined) usersSheet.getRange(i + 1, 5).setValue(updates.sections);
      if (updates.designation !== undefined) usersSheet.getRange(i + 1, 6).setValue(updates.designation);
      if (updates.isActive !== undefined) usersSheet.getRange(i + 1, 7).setValue(updates.isActive);
      return { success: true };
    }
  }
  
  throw new Error("User not found.");
}

/**
 * Soft deletes a user (Admin only).
 * @param {string} token 
 * @param {string} username 
 * @returns {object} {success: true}
 */
function deleteUser(token, username) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      usersSheet.getRange(i + 1, 7).setValue(false); // Set isActive to FALSE
      return { success: true };
    }
  }
  
  throw new Error("User not found.");
}

/**
 * Lists all users without passwords (Admin only).
 * @param {string} token 
 * @returns {Array} List of users
 */
function listUsers(token) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      username: data[i][0],
      displayName: data[i][2],
      role: data[i][3],
      sections: data[i][4],
      designation: data[i][5],
      isActive: data[i][6],
      createdDate: data[i][7],
      createdBy: data[i][8]
    });
  }
  
  return users;
}

/**
 * Resets a user's password (Admin only).
 * @param {string} token 
 * @param {string} targetUsername 
 * @param {string} newHash 
 * @returns {object} {success: true}
 */
function resetPassword(token, targetUsername, newHash) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetUsername) {
      usersSheet.getRange(i + 1, 2).setValue(newHash);
      return { success: true };
    }
  }
  
  throw new Error("User not found.");
}

/**
 * Re-seeds all default accounts into the Users sheet if missing.
 */
function seedDefaultAccounts(token) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  initializeDatabase();
  return { status: "SUCCESS", message: "All 59 official accounts initialized successfully." };
}
